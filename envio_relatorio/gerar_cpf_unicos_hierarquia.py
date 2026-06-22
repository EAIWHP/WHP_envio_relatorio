#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_cpf_unicos_hierarquia.py

Gera um arquivo Excel com CPFs únicos a partir da hierarquia consolidada.
Mantém o primeiro registro encontrado para cada CPF e adiciona metadados
sobre duplicidade (quantidade de ocorrências, arquivos e revendas).
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
    # Localiza o arquivo de validação mais recente
    candidatos = sorted(HIERARQUIA_DIR.glob("validacao_hierarquia_banco_*.xlsx"))
    if not candidatos:
        raise FileNotFoundError(f"Arquivo de validação não encontrado em: {HIERARQUIA_DIR}")

    arquivo_validacao = candidatos[-1]
    logger.info(f"Lendo arquivo de validação: {arquivo_validacao.name}")

    df_hier = pd.read_excel(arquivo_validacao, sheet_name="Hierarquia Consolidada", dtype=str)

    # Mantém apenas registros com CPF válido
    df_hier = df_hier[df_hier["cpf_limp"].notna()].copy()

    total_registros = len(df_hier)
    total_cpfs_unicos = df_hier["cpf_limp"].nunique()
    logger.info(f"{total_registros:,} registros -> {total_cpfs_unicos:,} CPFs únicos")

    # Conta ocorrências por CPF
    ocorrencias = df_hier.groupby("cpf_limp").agg(
        qtd_ocorrencias=("cpf_limp", "size"),
        qtd_arquivos_distintos=("arquivo_origem", "nunique"),
        qtd_revendas_distintos=("revenda_norm", "nunique"),
        qtd_lojas_distintos=("cod_loja", "nunique"),
        arquivos_origem=("arquivo_origem", lambda x: " | ".join(x.unique())),
        revendas_origem=("revenda_norm", lambda x: " | ".join(x.unique())),
        lojas_origem=("cod_loja", lambda x: " | ".join(x.dropna().astype(str).unique())),
    ).reset_index()

    # Remove duplicados mantendo o primeiro registro
    df_unicos = df_hier.drop_duplicates(subset=["cpf_limp"], keep="first").copy()

    # Adiciona metadados de duplicidade
    df_unicos = df_unicos.merge(ocorrencias, on="cpf_limp", how="left")

    # Adiciona CPF formatado
    df_unicos["cpf_formatado"] = df_unicos["cpf_limp"].apply(formatar_cpf)

    # Reordena colunas
    colunas_inicio = [
        "cpf_limp", "cpf_formatado", "nome", "revenda", "revenda_norm",
        "cod_loja", "cnpj", "cnpj_limp", "cargo", "vendedor", "gerente_loja",
        "gerente_regional", "diretor", "desligado",
    ]
    colunas_fim = [
        "qtd_ocorrencias", "qtd_arquivos_distintos", "qtd_revendas_distintos",
        "qtd_lojas_distintos", "arquivos_origem", "revendas_origem", "lojas_origem",
        "arquivo_origem",
    ]
    colunas_meio = [c for c in df_unicos.columns if c not in colunas_inicio + colunas_fim]
    df_unicos = df_unicos[colunas_inicio + colunas_meio + colunas_fim]

    # Aba 2: Resumo por revenda (baseado em CPFs únicos)
    df_resumo = df_unicos.groupby("revenda_norm").agg(
        total_cpfs_unicos=("cpf_limp", "nunique"),
        cpfs_duplicados=("qtd_ocorrencias", lambda x: (x > 1).sum()),
        max_ocorrencias=("qtd_ocorrencias", "max"),
    ).reset_index()
    df_resumo = df_resumo.sort_values("total_cpfs_unicos", ascending=False)

    # Aba 3: Apenas CPFs que tinham duplicados (resumo)
    df_dup_resumo = df_unicos[df_unicos["qtd_ocorrencias"] > 1][[
        "cpf_limp", "cpf_formatado", "nome", "revenda_norm", "qtd_ocorrencias",
        "qtd_arquivos_distintos", "qtd_revendas_distintos", "qtd_lojas_distintos",
        "arquivos_origem", "revendas_origem", "lojas_origem",
    ]].copy()
    df_dup_resumo = df_dup_resumo.sort_values("qtd_ocorrencias", ascending=False)

    # Gera arquivo de saída
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = HIERARQUIA_DIR / f"cpfs_unicos_hierarquia_{timestamp}.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df_unicos.to_excel(writer, sheet_name="CPFs Unicos", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo por Revenda", index=False)
        df_dup_resumo.to_excel(writer, sheet_name="CPFs Duplicados Resumo", index=False)

    logger.info(f"Arquivo de CPFs únicos gerado: {arquivo_saida}")
    print(f"\nArquivo gerado: {arquivo_saida}")
    print(f"Total de registros na hierarquia: {total_registros:,}")
    print(f"Total de CPFs únicos: {total_cpfs_unicos:,}")
    print(f"CPFs que apareciam duplicados: {len(df_dup_resumo):,}")


if __name__ == "__main__":
    main()
