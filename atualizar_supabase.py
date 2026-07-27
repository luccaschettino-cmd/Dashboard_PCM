import os
import sys
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials
from supabase import create_client

# ── CONFIGURAÇÕES ──────────────────────────────────────────
SUPABASE_URL = 'https://vnpexblbdiwgwdjsqsuh.supabase.co'

SHEET_ID  = '1k0opt7fDCp-rECaGL0N5ZU-2Q-nKf18tvY8dO1xvndc'
ABA_NOME  = 'Acumulado'

MESES = {
    1: 'Janeiro',  2: 'Fevereiro', 3: 'Março',    4: 'Abril',
    5: 'Maio',     6: 'Junho',     7: 'Julho',     8: 'Agosto',
    9: 'Setembro', 10: 'Outubro',  11: 'Novembro', 12: 'Dezembro',
}


def _parse_date(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, str):
        valor = valor.strip()
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
            try:
                return datetime.strptime(valor, fmt).date()
            except:
                continue
    return None


def _executar(credenciais_json, supabase_key):

    # ── CONEXÕES ──
    print('🔌 Conectando ao Google Sheets...')
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds  = Credentials.from_service_account_file(credenciais_json, scopes=scopes)
    gc     = gspread.authorize(creds)
    sheet  = gc.open_by_key(SHEET_ID).worksheet(ABA_NOME)

    print('🔌 Conectando ao Supabase...')
    supabase = create_client(SUPABASE_URL, supabase_key)

    # ── LÊ PLANILHA ──
    print('📊 Lendo dados da planilha...')
    linhas = sheet.get_all_records()
    print(f'   {len(linhas)} linhas encontradas')

    # ── PROCESSA DADOS ──
    totais_mes = {}
    os_individuais = []

    for i, row in enumerate(linhas, start=2):
        try:
            codigo        = str(row.get('Código', '')).strip()
            setor         = str(row.get('Setor', '')).strip()
            categoria     = str(row.get('Categoria', '')).strip()
            situacao      = str(row.get('Situação', '')).strip()
            solicitante   = str(row.get('Solicitante', '')).strip()
            prioridade    = str(row.get('Prioridade', '')).strip()
            abertura_raw  = row.get('Abertura', '')
            conclusao_raw = row.get('Conclusão', '')
            entra         = str(row.get('Entra na contagem?', '')).strip()
            descricao     = str(row.get('Descrição', '')).strip()

            if not codigo:
                continue

            dt_abertura = _parse_date(abertura_raw)
            dt_conclusao_raw = _parse_date(conclusao_raw)
            mes_num = dt_abertura.month if dt_abertura else 0

            # ── OS INDIVIDUAL ──
            os_individuais.append({
                'codigo':         codigo,
                'setor':          setor,
                'categoria':      categoria,
                'situacao':       situacao,
                'solicitante':    solicitante,
                'prioridade':     prioridade,
                'abertura':       dt_abertura.isoformat() if dt_abertura else None,
                'conclusao':      dt_conclusao_raw.isoformat() if dt_conclusao_raw else None,
                'entra_contagem': entra,
                'descricao':      descricao,
                'mes_num':        mes_num,
            })

            # ── INICIALIZA MÊS ──
            if mes_num not in totais_mes:
                totais_mes[mes_num] = {
                    'abertas': 0, 'conc30': 0, 'nconc30': 0,
                    'urgente': 0, 'alta': 0, 'normal': 0,
                    'tempos': [], 'em_aberto': 0,
                }

            m = totais_mes[mes_num]

            # ── EM ABERTO (independente de entrar na contagem) ──
            situacao_lower = situacao.lower()
            if 'aguardando' in situacao_lower or 'andamento' in situacao_lower:
                m['em_aberto'] += 1

            # ── TOTAIS (só as que entram na contagem) ──
            if entra.lower() != 'sim' or not dt_abertura:
                continue

            m['abertas'] += 1

            prioridade_lower = prioridade.lower()
            if 'urgente' in prioridade_lower:
                m['urgente'] += 1
            elif 'alta' in prioridade_lower:
                m['alta'] += 1
            else:
                m['normal'] += 1

            if 'conclu' in situacao_lower and conclusao_raw:
                dt_conclusao = _parse_date(conclusao_raw)
                if dt_conclusao:
                    dias = (dt_conclusao - dt_abertura).days
                    m['tempos'].append(dias)
                    if dias <= 30:
                        m['conc30'] += 1
                    else:
                        m['nconc30'] += 1

        except Exception as e:
            print(f'   ⚠️ Erro na linha {i}: {e}')
            continue

    # ── MONTA REGISTROS MENSAIS ──
    registros_mensais = []
    for mes_num in sorted(totais_mes.keys()):
        m = totais_mes[mes_num]
        tempos_validos = [t for t in m['tempos'] if t >= 0]
        tempo_medio = (
            round(sum(tempos_validos) / len(tempos_validos), 2)
            if tempos_validos else 0
        )
        indicador = round(m['conc30'] / m['abertas'] * 100, 2) if m['abertas'] else 0

        registros_mensais.append({
            'mes_num':      mes_num,
            'mes':          MESES.get(mes_num, str(mes_num)),
            'abertas':      m['abertas'],
            'conc30':       m['conc30'],
            'nconc30':      m['nconc30'],
            'tempo_medio':  tempo_medio,
            'indicador':    indicador,
            'urgente':      m['urgente'],
            'alta':         m['alta'],
            'normal':       m['normal'],
            'total_dias':   round(sum(tempos_validos), 2),
            'os_com_tempo': len(tempos_validos),
            'em_aberto':    m['em_aberto'],
        })

    print(f'\n📋 Meses encontrados: {[r["mes"] for r in registros_mensais]}')
    for r in registros_mensais:
        print(f'   {r["mes"]}: {r["abertas"]} OS | {r["indicador"]}% | em aberto: {r["em_aberto"]}')
    print(f'\n📋 OS individuais: {len(os_individuais)}')

    # ── ATUALIZA SUPABASE — TOTAIS MENSAIS ──
    print('\n☁️  Atualizando tabela os_intranet...')
    if registros_mensais:
        supabase.table('os_intranet').upsert(registros_mensais, on_conflict='mes_num').execute()
        print(f'   ✅ {len(registros_mensais)} meses atualizados!')
    else:
        print('   ⚠️ Nenhum mês encontrado — abortando para não apagar dados existentes.')

    # ── ATUALIZA SUPABASE — OS INDIVIDUAIS (upsert primeiro, depois limpa obsoletos) ──
    print('\n☁️  Atualizando tabela os_individual...')

    lote = 500
    total = len(os_individuais)
    if total == 0:
        print('   ⚠️ Nenhuma OS individual encontrada na planilha — abortando para não apagar dados existentes.')
    else:
        for i in range(0, total, lote):
            fatia = os_individuais[i:i+lote]
            supabase.table('os_individual').upsert(fatia, on_conflict='codigo').execute()
            print(f'   Enviando... {min(i+lote, total)}/{total}')
        print(f'   ✅ {total} OS individuais atualizadas!')

        # remove códigos que não existem mais na planilha
        codigos_atuais = [o['codigo'] for o in os_individuais]
        try:
            existentes = supabase.table('os_individual').select('codigo').execute().data
            codigos_banco = [r['codigo'] for r in existentes]
            obsoletos = [c for c in codigos_banco if c not in codigos_atuais]
            if obsoletos:
                for i in range(0, len(obsoletos), lote):
                    fatia_obs = obsoletos[i:i+lote]
                    supabase.table('os_individual').delete().in_('codigo', fatia_obs).execute()
                print(f'   🗑️ {len(obsoletos)} OS obsoletas removidas.')
        except Exception as e:
            print(f'   ⚠️ Não foi possível limpar códigos obsoletos: {e}')

    print('\n🎉 Supabase atualizado!')


def main():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    credenciais_json = os.path.join(base, 'credenciais.json')
    supabase_key     = os.environ.get('SUPABASE_KEY', '')

    if not supabase_key:
        print('❌ SUPABASE_KEY não definida!')
        return

    _executar(credenciais_json, supabase_key)


def main_gui(credenciais_json, supabase_key):
    _executar(credenciais_json, supabase_key)


if __name__ == '__main__':
    main()
