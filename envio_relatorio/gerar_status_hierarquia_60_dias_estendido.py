#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_status_hierarquia_60_dias_estendido.py

Lê o arquivo status_hierarquia_60_dias_YYYYMMDD.xlsx já gerado e cria uma cópia
estendida com informações de treinamentos obrigatórios e aceites no resumo.

Adiciona às abas Resumo_Regional e Resumo_Revenda:
- Realizaram ambos os cursos obrigatórios do mês
- Não realizaram
- % de realização
- Aceitaram o mês de referência
- Não aceitaram
- % de aceite
"""

import logging
import re
from datetime import date
from pathlib import Path

import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
ENTRADA_DIR = BASE_DIR / "relatorios_gerados"
SAIDA_DIR = ENTRADA_DIR
SAIDA_DIR.mkdir(exist_ok=True)

CADASTRO_FILE = DATA_DIR / "cadastro.xlsx"
TREIN_FILE = DATA_DIR / "Base_treinamentos.xlsx"
ACEITE_FILE = DATA_DIR / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"

ENTRADA_FILE = ENTRADA_DIR / f"status_hierarquia_60_dias_{date.today():%Y%m%d}.xlsx"
SAIDA_FILE = SAIDA_DIR / f"status_hierarquia_60_dias_{date.today():%Y%m%d}_estendido.xlsx"


def limpar_cpf(cpf):
    """Normaliza CPF em texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    cpf_str = str(cpf).strip().replace("'", "")
    cpf_limpo = re.sub(r"[^0-9]", "", cpf_str)
    if not cpf_limpo:
        return None
    return cpf_limpo.zfill(11)


def carregar_cadastro():
    """Carrega cadastro para mapeamento regional/revenda."""
    logger.info(f"Lendo cadastro: {CADASTRO_FILE.name}")
    df = pd.read_excel(CADASTRO_FILE, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    df["cpf_limp"] = df["cpf/cnpj"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()
    df["status"] = df["status"].fillna("").astype(str).str.strip().str.title()
    df["grupo"] = df["grupo"].fillna("").astype(str).str.strip()
    df["regional"] = df["regional"].fillna("").astype(str).str.strip()
    return df


def carregar_treinamentos():
    """Carrega base de treinamentos."""
    logger.info(f"Lendo treinamentos: {TREIN_FILE.name}")
    xl = pd.ExcelFile(TREIN_FILE)
    df = pd.read_excel(TREIN_FILE, sheet_name=xl.sheet_names[0])
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf)
    df["Conclusão"] = pd.to_datetime(df["Conclusão"], errors="coerce")
    df["Estado_norm"] = df["Estado"].astype(str).str.strip().str.lower()
    df["Trilha_norm"] = df["Trilha"].astype(str).str.strip().str.lower()
    return df


def carregar_aceites():
    """Carrega base de aceites mensais."""
    aceite_path = ACEITE_FILE
    if not aceite_path.exists():
        candidatos = sorted(DATA_DIR.glob("*Aceite*.xlsx"))
        if candidatos:
            aceite_path = candidatos[-1]
            logger.info(f"Arquivo padrão de aceites não encontrado. Usando: {aceite_path.name}")
        else:
            raise FileNotFoundError("Arquivo de aceites não encontrado em envio_relatorio/bases/")

    xl = pd.ExcelFile(aceite_path)
    # Usa a última aba por padrão
    aba = xl.sheet_names[-1]
    logger.info(f"Lendo aceites: {aceite_path.name} (aba {aba})")
    df = pd.read_excel(aceite_path, sheet_name=aba)
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf)
    df["DataAceite"] = pd.to_datetime(df["DataAceite"], errors="coerce")
    df["mes_aceite"] = df["DataAceite"].dt.to_period("M")
    return df


def detectar_cursos_obrigatorios(df_trein, ano_mes):
    """Detecta os 2 cursos obrigatórios do mês."""
    mes_dt = pd.Period(ano_mes, freq="M")
    obr = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado_norm"] == "concluido")
        & (df_trein["Trilha_norm"].str.contains("obrigat", case=False, na=False))
    ].copy()

    if obr.empty:
        return []

    top = obr["Curso"].value_counts().head(2)
    return top.index.tolist()


def calcular_treinamentos_e_aceites(df_status, df_trein, df_aceite, ano_mes):
    """
    Cruza os CPFs do status_hierarquia com treinamentos e aceites.
    Retorna DataFrames de resumo por regional e revenda enriquecidos.
    """
    cpfs_base = set(df_status["cpf_limp"].unique())

    # --- Treinamentos ---
    cursos = detectar_cursos_obrigatorios(df_trein, ano_mes)
    logger.info(f"Cursos obrigatórios detectados para {ano_mes}: {cursos}")

    if len(cursos) >= 2:
        mes_dt = pd.Period(ano_mes, freq="M")
        trein_mes = df_trein[
            (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
            & (df_trein["Estado_norm"] == "concluido")
        ]
        cpf_curso1 = set(trein_mes[trein_mes["Curso"] == cursos[0]]["cpf_limp"].unique())
        cpf_curso2 = set(trein_mes[trein_mes["Curso"] == cursos[1]]["cpf_limp"].unique())
        cpf_ambos = cpf_curso1 & cpf_curso2
    else:
        cpf_ambos = set()

    # --- Aceites ---
    mes_dt = pd.Period(ano_mes, freq="M")
    aceite_mes = df_aceite[df_aceite["mes_aceite"] == mes_dt]
    cpfs_aceitaram = set(aceite_mes["cpf_limp"].unique())
    logger.info(f"Aceites em {ano_mes}: {len(cpfs_aceitaram):,} CPFs únicos")

    # Marca no status
    df = df_status.copy()
    df["fez_2_cursos"] = df["cpf_limp"].isin(cpf_ambos)
    df["aceitou"] = df["cpf_limp"].isin(cpfs_aceitaram)

    return df


def main():
    logger.info(f"Lendo arquivo base: {ENTRADA_FILE.name}")
    if not ENTRADA_FILE.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {ENTRADA_FILE}")

    # Lê o consolidado do arquivo de status
    df_status = pd.read_excel(ENTRADA_FILE, sheet_name="Resumo_Revenda")
    df_status.columns = [c.strip() for c in df_status.columns]

    # Recupera os CPFs detalhados por revenda para cruzar com treinamentos/aceites
    xl = pd.ExcelFile(ENTRADA_FILE)
    dfs_revenda = []
    for aba in xl.sheet_names:
        if aba in ("Resumo_Regional", "Resumo_Revenda"):
            continue
        df = pd.read_excel(ENTRADA_FILE, sheet_name=aba, dtype=str)
        df["cpf_limp"] = df["cpf_limp"].astype(str).str.replace(r"[^0-9]", "", regex=True).str.zfill(11)
        df["revenda_arquivo"] = aba
        dfs_revenda.append(df)

    df_detalhe = pd.concat(dfs_revenda, ignore_index=True)
    df_detalhe["cpf_limp"] = df_detalhe["cpf_limp"].astype(str).str.zfill(11)

    # Carrega bases
    df_cad = carregar_cadastro()
    df_trein = carregar_treinamentos()
    df_aceite = carregar_aceites()

    # Determina mês de referência a partir dos treinamentos obrigatórios
    obr_concl = df_trein[
        (df_trein["Estado_norm"] == "concluido")
        & (df_trein["Trilha_norm"].str.contains("obrigat", case=False, na=False))
        & (df_trein["Conclusão"].notna())
    ]
    if not obr_concl.empty:
        ano_mes = str(obr_concl["Conclusão"].dt.to_period("M").max())
    else:
        ano_mes = date.today().strftime("%Y-%m")
    logger.info(f"Mês de referência: {ano_mes}")

    # Cruza treinamentos e aceites
    df_detalhe = calcular_treinamentos_e_aceites(df_detalhe, df_trein, df_aceite, ano_mes)

    # Carrega resumos originais
    resumo_reg = pd.read_excel(ENTRADA_FILE, sheet_name="Resumo_Regional")
    resumo_rev = pd.read_excel(ENTRADA_FILE, sheet_name="Resumo_Revenda")

    # Calcula métricas por revenda
    rev_metrics = (
        df_detalhe.groupby("revenda_arquivo")
        .agg(
            fez_2_cursos=("fez_2_cursos", "sum"),
            aceitou=("aceitou", "sum"),
        )
        .reset_index()
    )
    rev_metrics["fez_2_cursos"] = rev_metrics["fez_2_cursos"].astype(int)
    rev_metrics["aceitou"] = rev_metrics["aceitou"].astype(int)

    # Merge com resumo por revenda
    resumo_rev = resumo_rev.merge(rev_metrics, left_on="Revenda", right_on="revenda_arquivo", how="left")
    resumo_rev["Não fez 2 cursos"] = resumo_rev["Total participantes"] - resumo_rev["fez_2_cursos"]
    resumo_rev["% Fez 2 cursos"] = (resumo_rev["fez_2_cursos"] / resumo_rev["Total participantes"] * 100).round(1)
    resumo_rev["Não aceitou"] = resumo_rev["Total participantes"] - resumo_rev["aceitou"]
    resumo_rev["% Aceite"] = (resumo_rev["aceitou"] / resumo_rev["Total participantes"] * 100).round(1)

    # Reordena colunas
    cols_rev = [
        "Regional", "Revenda", "Total participantes", "Ativos no +TOP",
        "Inativos no +TOP", "% Ativos", "fez_2_cursos", "Não fez 2 cursos",
        "% Fez 2 cursos", "aceitou", "Não aceitou", "% Aceite"
    ]
    resumo_rev = resumo_rev[cols_rev].rename(columns={
        "fez_2_cursos": "Fez 2 cursos",
        "aceitou": "Aceitaram",
    })

    # Calcula métricas por regional
    reg_metrics = (
        resumo_rev.groupby("Regional")
        .agg(
            total=("Total participantes", "sum"),
            fez_2_cursos=("Fez 2 cursos", "sum"),
            aceitou=("Aceitaram", "sum"),
        )
        .reset_index()
    )

    resumo_reg = resumo_reg.merge(reg_metrics, on="Regional", how="left", suffixes=("", "_calc"))
    resumo_reg["Não fez 2 cursos"] = resumo_reg["Total participantes"] - resumo_reg["fez_2_cursos"]
    resumo_reg["% Fez 2 cursos"] = (resumo_reg["fez_2_cursos"] / resumo_reg["Total participantes"] * 100).round(1)
    resumo_reg["Não aceitou"] = resumo_reg["Total participantes"] - resumo_reg["aceitou"]
    resumo_reg["% Aceite"] = (resumo_reg["aceitou"] / resumo_reg["Total participantes"] * 100).round(1)

    cols_reg = [
        "Regional", "Total participantes", "Ativos no +TOP", "Inativos no +TOP",
        "% Ativos", "fez_2_cursos", "Não fez 2 cursos", "% Fez 2 cursos",
        "aceitou", "Não aceitou", "% Aceite"
    ]
    resumo_reg = resumo_reg[cols_reg].rename(columns={
        "fez_2_cursos": "Fez 2 cursos",
        "aceitou": "Aceitaram",
    })

    # Salva cópia estendida
    logger.info(f"Salvando arquivo estendido em: {SAIDA_FILE.name}")
    with pd.ExcelWriter(SAIDA_FILE, engine="openpyxl") as writer:
        resumo_reg.to_excel(writer, sheet_name="Resumo_Regional", index=False)
        resumo_rev.to_excel(writer, sheet_name="Resumo_Revenda", index=False)

        # Copia as demais abas do arquivo original
        for aba in xl.sheet_names:
            if aba in ("Resumo_Regional", "Resumo_Revenda"):
                continue
            df = pd.read_excel(ENTRADA_FILE, sheet_name=aba)
            df.to_excel(writer, sheet_name=aba, index=False)

    print("\n" + "=" * 80)
    print("ARQUIVO ESTENDIDO GERADO")
    print("=" * 80)
    print(f"Arquivo: {SAIDA_FILE}")
    print(f"Mês de referência: {ano_mes}")
    print("\nResumo por Regional:")
    print(resumo_reg.to_string(index=False))
    print("=" * 80)


if __name__ == "__main__":
    main()
