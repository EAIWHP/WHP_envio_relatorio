#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cruzar_cpfs_unicos_banco.py

Cruza os CPFs únicos da hierarquia com a base do banco PROD_WHP_Participantes_banco.
Gera um arquivo Excel com o cruzamento, CPFs só na hierarquia, CPFs em ambos
e resumo por revenda.
"""

import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
HIERARQUIA_DIR = BASE_DIR / "bases" / "bases_cadastro_hierarquia"


def formatar_cpf(cpf):
    """Formata CPF com pontos e traço para leitura humana."""
    if not cpf or len(cpf) != 11:
        return cpf
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


def main():
    # Localiza os arquivos mais recentes
    candidatos_validacao = sorted(HIERARQUIA_DIR.glob("validacao_hierarquia_banco_*.xlsx"))
    candidatos_unicos = sorted(HIERARQUIA_DIR.glob("cpfs_unicos_hierarquia_*.xlsx"))

    if not candidatos_validacao:
        raise FileNotFoundError(f"Arquivo de validação não encontrado em: {HIERARQUIA_DIR}")
    if not candidatos_unicos:
        raise FileNotFoundError(f"Arquivo de CPFs únicos não encontrado em: {HIERARQUIA_DIR}")

    arquivo_validacao = candidatos_validacao[-1]
    arquivo_unicos = candidatos_unicos[-1]

    logger.info(f"Lendo CPFs únicos: {arquivo_unicos.name}")
    df_unicos = pd.read_excel(arquivo_unicos, sheet_name="CPFs Unicos", dtype=str)

    logger.info(f"Lendo base do banco do arquivo de validação: {arquivo_validacao.name}")
    df_banco = pd.read_excel(arquivo_validacao, sheet_name="Banco", dtype=str)

    # Garante que estamos trabalhando com CPFs únicos do banco também
    df_banco_unicos = df_banco.drop_duplicates(subset=["cpf_limp"], keep="first").copy()

    total_unicos_hier = len(df_unicos)
    total_unicos_banco = df_banco_unicos["cpf_limp"].nunique()
    logger.info(f"CPFs únicos na hierarquia: {total_unicos_hier:,}")
    logger.info(f"CPFs únicos no banco: {total_unicos_banco:,}")

    # Cruzamento: left join dos CPFs únicos da hierarquia com o banco
    colunas_banco = [
        "cpf_limp", "nome", "ativo", "ativo_str", "data_inclusao",
        "data_aceite_regulamento", "data_bloqueio", "motivo_bloqueio",
        "origem_cadastro", "matricula", "email", "telefone", "celular",
        "revenda_banco",
    ]
    colunas_banco_presentes = [c for c in colunas_banco if c in df_banco_unicos.columns]

    df_banco_merge = df_banco_unicos[colunas_banco_presentes].rename(columns={"nome": "nome_banco"})

    df_cruzamento = df_unicos.merge(
        df_banco_merge,
        on="cpf_limp",
        how="left",
        suffixes=("", "_banco"),
    )

    # Flag de presença no banco
    df_cruzamento["encontrado_no_banco"] = df_cruzamento["nome_banco"].notna().map({True: "Sim", False: "Não"})

    # Separa grupos
    df_ambos = df_cruzamento[df_cruzamento["encontrado_no_banco"] == "Sim"].copy()
    df_somente_hier = df_cruzamento[df_cruzamento["encontrado_no_banco"] == "Não"].copy()

    # CPFs únicos só no banco (não estão na hierarquia)
    cpfs_hier = set(df_unicos["cpf_limp"].unique())
    df_somente_banco = df_banco_unicos[~df_banco_unicos["cpf_limp"].isin(cpfs_hier)].copy()

    logger.info(f"CPFs únicos em ambos: {len(df_ambos):,}")
    logger.info(f"CPFs únicos só na hierarquia: {len(df_somente_hier):,}")
    logger.info(f"CPFs únicos só no banco: {len(df_somente_banco):,}")

    # Resumo por revenda (hierarquia)
    df_resumo = df_cruzamento.groupby("revenda_norm").agg(
        total_cpfs_unicos=("cpf_limp", "nunique"),
        no_banco=("encontrado_no_banco", lambda x: (x == "Sim").sum()),
        nao_no_banco=("encontrado_no_banco", lambda x: (x == "Não").sum()),
        ativos_no_banco=("ativo_str", lambda x: (x == "Sim").sum()),
        inativos_no_banco=("ativo_str", lambda x: (x == "Não").sum()),
    ).reset_index()
    df_resumo["pct_no_banco"] = (df_resumo["no_banco"] / df_resumo["total_cpfs_unicos"] * 100).round(2)
    df_resumo = df_resumo.sort_values("total_cpfs_unicos", ascending=False)

    # Gera arquivo de saída
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = HIERARQUIA_DIR / f"cpfs_unicos_hierarquia_cruzamento_banco_{timestamp}.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df_cruzamento.to_excel(writer, sheet_name="Cruzamento Completo", index=False)
        df_ambos.to_excel(writer, sheet_name="Em Ambos", index=False)
        df_somente_hier.to_excel(writer, sheet_name="So na Hierarquia", index=False)
        df_somente_banco.to_excel(writer, sheet_name="So no Banco", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo por Revenda", index=False)

    logger.info(f"Arquivo de cruzamento gerado: {arquivo_saida}")
    print(f"\nArquivo gerado: {arquivo_saida}")
    print(f"CPFs únicos na hierarquia: {total_unicos_hier:,}")
    print(f"CPFs únicos no banco: {total_unicos_banco:,}")
    print(f"Em ambos: {len(df_ambos):,}")
    print(f"Só na hierarquia: {len(df_somente_hier):,}")
    print(f"Só no banco: {len(df_somente_banco):,}")


if __name__ == "__main__":
    main()
