# -*- coding: utf-8 -*-
"""
Exportador de Dados PCM -> JSON
--------------------------------
Lê a base de dados de OS (Excel) e gera um arquivo "data.json" limpo,
pronto para ser usado pela página web (index.html) com filtros.

Como usar:
  1. Ajuste ARQUIVO_EXCEL abaixo.
  2. Rode no terminal do VS Code: python exportar_dados_web.py
  3. Um arquivo "data.json" será criado. Suba esse arquivo pro GitHub
     junto com o index.html, toda vez que quiser atualizar a página.
"""

import pandas as pd
from datetime import time, timedelta, datetime
import json

# ============================================================
# CONFIGURAÇÃO — AJUSTE AQUI
# ============================================================
ARQUIVO_EXCEL = "Base_de_dados_PCM.xlsx"
LINHA_CABECALHO = 5
ARQUIVO_SAIDA = "data.json"


def normalizar_texto(valor):
    if pd.isna(valor):
        return "Não informado"
    texto = str(valor).strip().lower()
    if texto in ("", "-", "n", "sem dados"):
        return "Não informado"
    correcoes = {"melhroia": "melhoria", "inspeção/acompanhamento": "inspeção"}
    texto = correcoes.get(texto, texto)
    return texto.capitalize()


def converter_para_horas(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, timedelta):
        return round(valor.total_seconds() / 3600, 2)
    if isinstance(valor, time):
        return round(valor.hour + valor.minute / 60 + valor.second / 3600, 2)
    if isinstance(valor, datetime):
        return round(valor.hour + valor.minute / 60 + valor.second / 3600, 2)
    return 0.0


def main():
    print(f"Lendo dados de: {ARQUIVO_EXCEL}")
    df = pd.read_excel(ARQUIVO_EXCEL, header=LINHA_CABECALHO - 1)
    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]

    df["Setor"] = df["Setor"].apply(normalizar_texto)
    df["Tipo de Manteunção"] = df["Tipo de Manteunção"].apply(normalizar_texto)
    df["Responsavel"] = df["Responsavel"].apply(normalizar_texto)
    df["Horas_manutencao"] = df["Tempo de manutenção"].apply(converter_para_horas)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df.dropna(subset=["Data"])

    registros = []
    for _, row in df.iterrows():
        registros.append({
            "data": row["Data"].strftime("%Y-%m-%d"),
            "setor": row["Setor"],
            "responsavel": row["Responsavel"],
            "tipo": row["Tipo de Manteunção"],
            "horas": row["Horas_manutencao"],
        })

    with open(ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, indent=None)

    print(f"{len(registros)} registros exportados para {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()
