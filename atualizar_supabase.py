import gspread
from google.oauth2.service_account import Credentials
from supabase import create_client
from datetime import datetime, date
import os

# ── CONFIGURAÇÕES ──────────────────────────────────────────
CREDENCIAIS_JSON = os.path.join(os.path.dirname(__file__), 'credenciais.json')

SUPABASE_URL = 'https://vnpexblbdiwgwdjsqsuh.supabase.co'
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')

SHEET_ID     = '1k0opt7fDCp-rECaGL0N5ZU-2Q-nKf18tvY8dO1xvndc'
ABA_NOME     = 'Acumulado'

MESES = {
    1:'Janeiro', 2:'Fevereiro', 3:'Março',    4:'Abril',
    5:'Maio',    6:'Junho',     7:'Julho',     8:'Agosto',
    9:'Setembro',10:'Outubro',  11:'Novembro', 12:'Dezembro'
}

# ── CONEXÕES ───────────────────────────────────────────────
print('🔌 Conectando ao Google Sheets...')
scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
creds  = Credentials.from_service_account_file(CREDENCIAIS_JSON, scopes=scopes)
gc     = gspread.authorize(creds)
sheet  = gc.open_by_key(SHEET_ID).worksheet(ABA_NOME)

print('🔌 Conectando ao Supabase...')
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# ── LÊ DADOS DA PLANILHA ───────────────────────────────────
print('📊 Lendo dados da planilha...')
linhas = sheet.get_all_records(head=2)  # cabeçalho na linha 2
print(f'   {len(linhas)} linhas encontradas')

# ── PROCESSA OS DADOS ──────────────────────────────────────
totais_mes = {}  # { mes_num: { abertas, conc30, nconc30, urgente, alta, normal, tempos[] } }

for i, row in enumerate(linhas, start=3):
    try:
        # Verifica se entra na contagem
        entra = str(row.get('Entra na contagem?', '')).strip().lower()
        if entra != 'sim':
            continue

        # Pega a data de abertura pra determinar o mês
        abertura_raw = row.get('Abertura', '')
        if not abertura_raw:
            continue

        # Converte data de abertura
        if isinstance(abertura_raw, str):
            abertura_raw = abertura_raw.strip()
            for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
                try:
                    dt_abertura = datetime.strptime(abertura_raw, fmt).date()
                    break
                except:
                    continue
            else:
                print(f'   ⚠️ Linha {i}: data de abertura inválida: {abertura_raw}')
                continue
        elif isinstance(abertura_raw, (datetime, date)):
            dt_abertura = abertura_raw if isinstance(abertura_raw, date) else abertura_raw.date()
        else:
            continue

        mes_num = dt_abertura.month

        # Inicializa mês se necessário
        if mes_num not in totais_mes:
            totais_mes[mes_num] = {
                'abertas': 0, 'conc30': 0, 'nconc30': 0,
                'urgente': 0, 'alta': 0, 'normal': 0, 'tempos': []
            }

        m = totais_mes[mes_num]
        m['abertas'] += 1

        # Prioridade
        prioridade = str(row.get('Prioridade', '')).strip().lower()
        if 'urgente' in prioridade:
            m['urgente'] += 1
        elif 'alta' in prioridade:
            m['alta'] += 1
        else:
            m['normal'] += 1

        # Calcula tempo de conclusão
        situacao = str(row.get('Situação', '')).strip().lower()
        conclusao_raw = row.get('Conclusão', '')

        if 'conclu' in situacao and conclusao_raw:
            if isinstance(conclusao_raw, str):
                conclusao_raw = conclusao_raw.strip()
                for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
                    try:
                        dt_conclusao = datetime.strptime(conclusao_raw, fmt).date()
                        break
                    except:
                        continue
                else:
                    dt_conclusao = None
            elif isinstance(conclusao_raw, (datetime, date)):
                dt_conclusao = conclusao_raw if isinstance(conclusao_raw, date) else conclusao_raw.date()
            else:
                dt_conclusao = None

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

# ── MONTA REGISTROS PRA INSERIR ────────────────────────────
registros = []
for mes_num in sorted(totais_mes.keys()):
    m = totais_mes[mes_num]
    tempos_validos = [t for t in m['tempos'] if t >= 0]
    tempo_medio = round(sum(tempos_validos) / len(tempos_validos), 2) if tempos_validos else 0
    indicador   = round(m['conc30'] / m['abertas'] * 100, 2) if m['abertas'] else 0

    registros.append({
        'mes_num':     mes_num,
        'mes':         MESES.get(mes_num, str(mes_num)),
        'abertas':     m['abertas'],
        'conc30':      m['conc30'],
        'nconc30':     m['nconc30'],
        'tempo_medio': tempo_medio,
        'indicador':   indicador,
        'urgente':     m['urgente'],
        'alta':        m['alta'],
        'normal':      m['normal'],
        'total_dias':   round(sum(tempos_validos), 2),
        'os_com_tempo': len(tempos_validos),
    })

print(f'\n📋 Meses encontrados: {[r["mes"] for r in registros]}')
for r in registros:
    print(f'   {r["mes"]}: {r["abertas"]} OS | indicador {r["indicador"]}% | tempo médio {r["tempo_medio"]}d')

# ── ATUALIZA SUPABASE ──────────────────────────────────────
print('\n☁️  Atualizando Supabase...')

# Apaga tudo e reinsere (upsert limpo)
supabase.table('os_intranet').delete().neq('id', 0).execute()
print('   Dados antigos removidos.')

resultado = supabase.table('os_intranet').insert(registros).execute()
print(f'   ✅ {len(registros)} meses inseridos com sucesso!')

print('\n🎉 Supabase atualizado! O dashboard vai refletir os novos dados em instantes.')
