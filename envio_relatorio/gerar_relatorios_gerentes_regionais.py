#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera relatórios individuais para cada Gerente Regional usando a hierarquia como fonte.

Regra de responsabilidade:
  • O gerente regional é identificado pela coluna GERENTE REGIONAL = SIM na hierarquia.
  • Cada gerente possui um COD LOJA / CNPJ associado na própria hierarquia.
  • A equipe do gerente = todos os CPFs da mesma revenda + COD LOJA + CNPJ.

Cada relatório contém:
  • Aba Resumo: nome do gerente regional, CPF, revenda, código da loja e CNPJ
  • Aba Detalhamento: equipe da loja com status de cadastro e treinamentos obrigatórios do mês.

Os arquivos são salvos em:
  envio_relatorio/relatorios_gerentes_regionais/

Uso:
  python3 gerar_relatorios_gerentes_regionais.py
"""

import logging
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import numpy as np

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from gerar_relatorio_semanal import (
    carregar_bases,
    detectar_cursos_obrigatorios,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "relatorios_gerentes_regionais"


def formatar_mes_referencia(mes_referencia):
    """Converte '2026-06' para 'Junho/2026'."""
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
    }
    p = pd.Period(mes_referencia, freq="M")
    return f"{meses_pt[p.month]}/{p.year}"


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    logger.info("Carregando bases...")

    bases = carregar_bases()
    df_hier = bases["hierarquia"]
    df_cad = bases["cadastro"]
    df_trein = bases["treinamentos"]
    mes_referencia = bases["mes_referencia"]

    # ------------------------------------------------------------------
    # Identifica gerentes regionais (GERENTE REGIONAL = SIM) na hierarquia
    # ------------------------------------------------------------------
    mask_gerente = df_hier["gerente_regional"].astype(str).str.strip().str.upper() == "SIM"
    df_gerentes = df_hier[mask_gerente].copy()

    if df_gerentes.empty:
        logger.warning("Nenhum gerente regional encontrado na hierarquia (GERENTE REGIONAL = SIM).")
        return

    logger.info(f"{len(df_gerentes)} registros de gerente regional encontrados na hierarquia")

    # Cruza com cadastro para obter nome do gerente
    df_gerentes = df_gerentes.merge(
        df_cad[["cpf_limp", "nome"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="left",
        suffixes=("", "_cad"),
    )
    df_gerentes["nome"] = df_gerentes["nome"].fillna(df_gerentes["nome_hier"])

    # Normaliza colunas para chave de loja
    def _norm_cod_loja(x):
        if pd.isna(x):
            return None
        s = str(x).strip().replace(".0", "")
        if s.lower() in ("", "nan"):
            return None
        return s

    def _norm_cnpj(x):
        if pd.isna(x):
            return None
        s = re.sub(r"[^0-9]", "", str(x).strip())
        if s == "":
            return None
        return s

    df_gerentes["cod_loja"] = df_gerentes["cod_loja"].apply(_norm_cod_loja)
    df_hier["cod_loja"] = df_hier["cod_loja"].apply(_norm_cod_loja)
    df_gerentes["cnpj"] = df_gerentes["cnpj"].apply(_norm_cnpj)
    df_hier["cnpj"] = df_hier["cnpj"].apply(_norm_cnpj)

    # Remove duplicados de gerente (mesmo CPF pode aparecer em mais de um arquivo)
    df_gerentes = df_gerentes.drop_duplicates(subset=["cpf_limp"], keep="first")
    logger.info(f"{len(df_gerentes)} gerentes regionais únicos")

    # ------------------------------------------------------------------
    # Prepara dados de treinamento
    # ------------------------------------------------------------------
    curso1, curso2, nome1, nome2, sku1, sku2 = detectar_cursos_obrigatorios(
        df_trein, mes_referencia
    )

    mes_dt = pd.Period(mes_referencia, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ]
    cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique()) if curso1 else set()
    cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique()) if curso2 else set()

    nome_col1 = f"Treinamento - Conteúdo Obrigatório | {mes_referencia} | {sku1}"
    nome_col2 = f"Treinamento - Conteúdo Obrigatório | {mes_referencia} | {sku2}"

    # ------------------------------------------------------------------
    # Gera um relatório por gerente regional
    # ------------------------------------------------------------------
    arquivos_gerados = []

    for _, gerente in df_gerentes.iterrows():
        revenda = gerente["revenda"]
        cod_loja = gerente["cod_loja"]
        cnpj = gerente["cnpj"]

        # Equipe = mesma revenda + cod_loja (ou cnpj se cod_loja ausente)
        mask_equipe = df_hier["revenda"] == revenda
        if pd.notna(cod_loja):
            mask_equipe = mask_equipe & (df_hier["cod_loja"] == cod_loja)
        elif pd.notna(cnpj):
            mask_equipe = mask_equipe & (df_hier["cnpj"] == cnpj)
        else:
            logger.warning(
                f"Gerente {gerente['nome']} ({revenda}) sem COD LOJA e sem CNPJ na hierarquia. Ignorado."
            )
            continue

        equipe_hier = df_hier[mask_equipe].copy()

        if equipe_hier.empty:
            logger.warning(
                f"Nenhum CPF encontrado para loja {cod_loja} ({revenda}) do gerente {gerente['nome']}"
            )
            continue

        # Cruza com cadastro
        equipe = equipe_hier.merge(
            df_cad[["cpf_limp", "nome", "cargo", "status"]].drop_duplicates("cpf_limp"),
            on="cpf_limp",
            how="left",
        )
        equipe["nome"] = equipe["nome"].fillna(equipe["nome_hier"])
        equipe["cargo"] = equipe["cargo"].fillna(equipe.get("cargo_hier"))

        # Flags de treinamento
        equipe[nome_col1] = np.where(equipe["cpf_limp"].isin(cpf_curso1), "Sim", "Não")
        equipe[nome_col2] = np.where(equipe["cpf_limp"].isin(cpf_curso2), "Sim", "Não")
        equipe["Fez os 2 Treinamentos"] = np.where(
            equipe["cpf_limp"].isin(cpf_curso1) & equipe["cpf_limp"].isin(cpf_curso2),
            "Sim", "Não"
        )

        # Mês de referência legível
        mes_legivel = formatar_mes_referencia(mes_referencia)
        equipe["Mês de Referência"] = mes_legivel

        # Renomeia colunas para o padrão do exemplo
        equipe = equipe.rename(columns={
            "revenda": "Revenda",
            "cod_loja": "Código da Loja",
            "cnpj": "CNPJ",
            "nome": "Nome",
            "cargo": "Cargo",
            "status": "Status Cadastro",
        })

        # Aba Resumo
        df_resumo = pd.DataFrame({
            "Indicador": [
                "Regional",
                "CPF do Regional",
                "Revenda",
                "Código da Loja",
                "CNPJ",
            ],
            "Valor": [
                gerente["nome"],
                gerente["cpf_limp"],
                revenda,
                cod_loja,
                cnpj,
            ],
        })

        # Aba Detalhamento (sem CPF dos participantes)
        colunas_det = [
            "Revenda", "Código da Loja", "CNPJ", "Nome", "Cargo",
            "Status Cadastro", nome_col1, nome_col2,
            "Fez os 2 Treinamentos", "Mês de Referência",
        ]
        df_det = equipe[colunas_det].copy()
        df_det = df_det.sort_values(["Nome"])

        # Nome do arquivo (usa cod_loja; se ausente, usa cnpj)
        nome_gerente = re.sub(r"[^\w\s-]", "", str(gerente["nome"])).strip().replace(" ", "_")
        id_loja = str(cod_loja) if pd.notna(cod_loja) else (str(cnpj) if pd.notna(cnpj) else "sem_loja")
        caminho = OUTPUT_DIR / (
            f"relatorio_gerente_regional_{nome_gerente.lower()}_"
            f"{revenda.lower().replace(' ', '_')}_{id_loja}_"
            f"{date.today():%Y%m%d}.xlsx"
        )

        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
            df_det.to_excel(writer, sheet_name="Detalhamento", index=False)

            for ws in writer.sheets.values():
                for col in ws.columns:
                    max_length = max(len(str(cell.value or "")) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

        arquivos_gerados.append(caminho.name)
        logger.info(
            f"Gerado: {caminho.name} ({len(df_det)} CPFs) - "
            f"{gerente['nome']} | {revenda} | Loja {cod_loja}"
        )

    logger.info(f"Total de relatórios gerados: {len(arquivos_gerados)}")
    logger.info(f"Pasta: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
