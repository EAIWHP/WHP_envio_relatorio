#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_relatorio_duplicados_hierarquia.py

Gera um arquivo Excel separado com os CPFs duplicados nas bases de hierarquia,
organizado em abas para facilitar a verificação manual.
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
        raise FileNotFoundError("Arquivo de validação não encontrado em: {HIERARQUIA_DIR}")

    arquivo_validacao = candidatos[-1]
    logger.info(f"Lendo arquivo de validação: {arquivo_validacao.name}")

    df_hier = pd.read_excel(arquivo_validacao, sheet_name="Hierarquia Consolidada", dtype=str)

    # Mantém apenas registros com CPF válido
    df_hier = df_hier[df_hier["cpf_limp"].notna()].copy()

    # Identifica CPFs duplicados
    cpfs_duplicados = df_hier[df_hier.duplicated(subset=["cpf_limp"], keep=False)]["cpf_limp"].unique()
    logger.info(f"{len(cpfs_duplicados)} CPFs únicos duplicados encontrados")

    df_dup = df_hier[df_hier["cpf_limp"].isin(cpfs_duplicados)].copy()
    df_dup["cpf_formatado"] = df_dup["cpf_limp"].apply(formatar_cpf)

    # Ordena para facilitar leitura
    df_dup = df_dup.sort_values(["cpf_limp", "arquivo_origem", "cod_loja"])

    # Aba 1: Resumo por CPF
    resumo_rows = []
    for cpf in sorted(cpfs_duplicados):
        subset = df_dup[df_dup["cpf_limp"] == cpf]
        resumo_rows.append({
            "cpf_limp": cpf,
            "cpf_formatado": formatar_cpf(cpf),
            "nome": subset["nome"].iloc[0],
            "total_ocorrencias": len(subset),
            "qtd_arquivos_distintos": subset["arquivo_origem"].nunique(),
            "qtd_revendas_distintas": subset["revenda_norm"].nunique(),
            "qtd_lojas_distintas": subset["cod_loja"].nunique(),
            "revendas": " | ".join(subset["revenda_norm"].unique()),
            "arquivos": " | ".join(subset["arquivo_origem"].unique()),
            "lojas": " | ".join(subset["cod_loja"].dropna().astype(str).unique()),
        })

    df_resumo = pd.DataFrame(resumo_rows)
    df_resumo["tipo_duplicacao"] = df_resumo.apply(
        lambda r: "Arquivos diferentes" if r["qtd_arquivos_distintos"] > 1 else "Mesmo arquivo",
        axis=1,
    )

    # Aba 2: Detalhamento completo
    df_detalhe = df_dup[[
        "cpf_limp", "cpf_formatado", "revenda", "revenda_norm", "cod_loja",
        "cnpj", "cnpj_limp", "nome", "cargo", "vendedor", "gerente_loja",
        "gerente_regional", "diretor", "desligado", "arquivo_origem",
    ]].copy()

    # Aba 3: Duplicados em arquivos diferentes
    df_arquivos_diferentes = df_resumo[df_resumo["qtd_arquivos_distintos"] > 1].copy()

    # Aba 4: Duplicados no mesmo arquivo
    df_mesmo_arquivo = df_resumo[df_resumo["qtd_arquivos_distintos"] == 1].copy()

    # Gera arquivo de saída
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = HIERARQUIA_DIR / f"duplicados_hierarquia_{timestamp}.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo por CPF", index=False)
        df_detalhe.to_excel(writer, sheet_name="Detalhamento Completo", index=False)
        df_arquivos_diferentes.to_excel(writer, sheet_name="Arquivos Diferentes", index=False)
        df_mesmo_arquivo.to_excel(writer, sheet_name="Mesmo Arquivo", index=False)

    logger.info(f"Arquivo de duplicados gerado: {arquivo_saida}")
    print(f"\nArquivo gerado: {arquivo_saida}")
    print(f"Total de CPFs duplicados: {len(cpfs_duplicados):,}")
    print(f"Duplicados em arquivos diferentes: {len(df_arquivos_diferentes):,}")
    print(f"Duplicados no mesmo arquivo: {len(df_mesmo_arquivo):,}")


if __name__ == "__main__":
    main()
