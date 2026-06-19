#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório Detalhado — MM Atacado, MM Varejo e Guaibim
----------------------------------------------------
Gera relatório separado com:
  • Total de participantes
  • Quantidade por status
  • Quantidade que deu aceite
  • CNPJ e nome das lojas
  • Quantidade que fez os 2 treinamentos obrigatórios do mês e quantos não

Período de referência: 01/05/2026 a 16/06/2026

Fontes:
  - envio_relatorio/bases/cadastro.xlsx
  - envio_relatorio/bases/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx
  - envio_relatorio/bases/WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx
  - envio_relatorio/bases/Base_treinamentos.xlsx
  - validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx

Uso:
  python3 gerar_relatorio_revendas_mm_guaibim.py
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
BANCO_PARTICIPANTES_PATH = Path("/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx")

OUTPUT_DIR.mkdir(exist_ok=True)

REVENDAS_ALVO = {"MM Atacado", "MM Varejo", "Guaibim"}
PRAZO_PRE_CADASTRO_INATIVO = 90
DATA_INICIO = pd.Timestamp("2026-05-01")
DATA_FIM = pd.Timestamp("2026-06-16")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("relatorio_mm_guaibim")


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ---------------------------------------------------------------------------
def limpar_cpf(cpf):
    """Normaliza CPF/CNPJ para texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    try:
        if isinstance(cpf, float):
            cpf = int(cpf)
    except Exception:
        pass
    s = str(cpf).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(11)


def limpar_cnpj(cnpj):
    """Normaliza CNPJ para texto com 14 dígitos."""
    if pd.isna(cnpj):
        return None
    try:
        if isinstance(cnpj, float):
            cnpj = int(cnpj)
    except Exception:
        pass
    s = str(cnpj).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(14)


def detectar_cursos_obrigatorios(df_trein, ano_mes):
    """Detecta os 2 cursos obrigatórios do mês."""
    mes_dt = pd.Period(ano_mes, freq="M")
    obr = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
        & (df_trein["Trilha"].str.contains("OBRIGAT", case=False, na=False))
    ].copy()

    if obr.empty:
        return None, None

    top = obr["Curso"].value_counts().head(2)
    return tuple(top.index.tolist()) if len(top) >= 1 else (None, None)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    logger.info("=" * 60)
    logger.info("Iniciando relatório detalhado MM Atacado / MM Varejo / Guaibim")
    logger.info(f"Período: {DATA_INICIO:%d/%m/%Y} a {DATA_FIM:%d/%m/%Y}")

    # --- Cadastro ---
    cadastro_path = DATA_DIR / "cadastro.xlsx"
    df_cad = pd.read_excel(cadastro_path)
    df_cad["cpf_limp"] = df_cad["cpf/cnpj"].apply(limpar_cpf)
    logger.info(f"Cadastro carregado: {len(df_cad):,} registros")

    # Aplica regra Pré-Cadastrado 90+ dias na base do banco
    if BANCO_PARTICIPANTES_PATH.exists():
        df_banco = pd.read_excel(BANCO_PARTICIPANTES_PATH)
        df_banco["cpf_limp"] = df_banco["Cpf"].apply(limpar_cpf)
        df_banco["DataInclusao_dt"] = pd.to_datetime(df_banco["DataInclusao"], errors="coerce")
        banco_clean = df_banco[["cpf_limp", "DataInclusao_dt"]].drop_duplicates(subset=["cpf_limp"], keep="first")
        df_cad = df_cad.merge(banco_clean, on="cpf_limp", how="left")
        dias_desde_inclusao = (pd.Timestamp("today").normalize() - df_cad["DataInclusao_dt"]).dt.days
        mask_pre_inativo = (
            (df_cad["status"] == "Pré-Cadastrado")
            & (df_cad["DataInclusao_dt"].notna())
            & (dias_desde_inclusao >= PRAZO_PRE_CADASTRO_INATIVO)
        )
        qtd = mask_pre_inativo.sum()
        if qtd > 0:
            df_cad.loc[mask_pre_inativo, "status"] = "Inativo"
            logger.info(f"{qtd:,} Pré-Cadastrados com 90+ dias na base do banco convertidos para Inativo")
        df_cad = df_cad.drop(columns=["DataInclusao_dt"])

    # Exclui Aguardando Aprovação
    df_cad = df_cad[df_cad["status"] != "Aguardando Aprovação"].copy()

    # Filtra revendas alvo
    df_alvo = df_cad[df_cad["grupo"].isin(REVENDAS_ALVO)].copy()
    logger.info(f"Revendas alvo: {len(df_alvo):,} registros")

    # --- Cadastro Revenda x Loja x Regional (CNPJ/nome loja) ---
    det_path = Path("/home/thamiresvieira/projetos/Programa_mais_top/bases/Tabelas/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx")
    df_det = pd.read_excel(det_path)
    df_det["cpf_limp"] = df_det["Cpf"].apply(limpar_cpf)
    df_det["CNPJ_Loja"] = df_det["CNPJ_Loja"].apply(limpar_cnpj)
    df_det["CNPJ_Revenda"] = df_det["CNPJ_Revenda"].apply(limpar_cnpj)
    df_det = df_det.rename(columns={
        "Revenda": "revenda_det",
        "Loja": "loja",
        "CNPJ_Loja": "cnpj_loja",
        "CNPJ_Revenda": "cnpj_revenda",
        "Regional_da_Loja": "regional_loja",
        "Regional_da_Revenda": "regional_revenda",
    })
    logger.info(f"Detalhamento revenda/loja carregado: {len(df_det):,} registros")

    # Cruzar cadastro com detalhamento
    df_alvo = df_alvo.merge(
        df_det[["cpf_limp", "revenda_det", "loja", "cnpj_loja", "cnpj_revenda", "regional_loja", "regional_revenda"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="left",
    )

    # --- Aceites ---
    aceite_path = DATA_DIR / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
    xl_aceite = pd.ExcelFile(aceite_path)
    aba_maior = None
    for aba in xl_aceite.sheet_names:
        if "MAI" in aba.upper() or "MAY" in aba.upper():
            aba_maior = aba
            break
    if aba_maior is None:
        aba_maior = xl_aceite.sheet_names[-1]
        logger.warning(f"Aba de maio não encontrada. Usando: {aba_maior}")

    df_aceite = pd.read_excel(aceite_path, sheet_name=aba_maior)
    df_aceite["cpf_limp"] = df_aceite["CPF"].apply(limpar_cpf)
    df_aceite["DataAceite"] = pd.to_datetime(df_aceite["DataAceite"], errors="coerce")
    aceitaram = set(df_aceite["cpf_limp"].unique())
    logger.info(f"Aceites ({aba_maior}): {len(aceitaram):,} CPFs únicos")

    df_alvo["deu_aceite"] = df_alvo["cpf_limp"].isin(aceitaram)

    # --- Treinamentos ---
    trein_path = DATA_DIR / "Base_treinamentos.xlsx"
    df_trein = pd.read_excel(trein_path)
    df_trein["cpf_limp"] = df_trein["CPF"].apply(limpar_cpf)
    df_trein["Conclusão"] = pd.to_datetime(df_trein["Conclusão"], errors="coerce")

    # Detecta cursos obrigatórios de junho/2026
    curso1, curso2 = detectar_cursos_obrigatorios(df_trein, "2026-06")
    if curso1 is None or curso2 is None:
        logger.warning("Não foi possível detectar 2 cursos obrigatórios de junho/2026")
        curso1, curso2 = None, None
    else:
        logger.info(f"Cursos obrigatórios detectados: '{curso1}' e '{curso2}'")

    if curso1 and curso2:
        trein_mes = df_trein[
            (df_trein["Conclusão"] >= DATA_INICIO)
            & (df_trein["Conclusão"] <= DATA_FIM)
            & (df_trein["Estado"].str.lower() == "concluido")
        ].copy()
        cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique())
        cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique())
        cpf_ambos = cpf_curso1 & cpf_curso2
        df_alvo["fez_curso1"] = df_alvo["cpf_limp"].isin(cpf_curso1)
        df_alvo["fez_curso2"] = df_alvo["cpf_limp"].isin(cpf_curso2)
        df_alvo["fez_ambos"] = df_alvo["cpf_limp"].isin(cpf_ambos)
    else:
        df_alvo["fez_curso1"] = False
        df_alvo["fez_curso2"] = False
        df_alvo["fez_ambos"] = False

    # --- Geração dos DataFrames de saída ---

    # 1) Resumo por revenda
    resumo_revenda = []
    for revenda in sorted(REVENDAS_ALVO):
        sub = df_alvo[df_alvo["grupo"] == revenda]
        if sub.empty:
            continue
        status_counts = sub["status"].value_counts().to_dict()
        resumo_revenda.append({
            "Revenda": revenda,
            "Total Participantes": len(sub),
            "Ativo": status_counts.get("Ativo", 0),
            "Inativo": status_counts.get("Inativo", 0),
            "Pré-Cadastrado": status_counts.get("Pré-Cadastrado", 0),
            "Indicado Moderado": status_counts.get("Indicado Moderado", 0),
            "Em Moderação": status_counts.get("Em Moderação", 0),
            "Reprovado": status_counts.get("Reprovado", 0),
            "Bloqueado": status_counts.get("Bloqueado", 0),
            "Deu Aceite": sub["deu_aceite"].sum(),
            "Não Deu Aceite": (~sub["deu_aceite"]).sum(),
            f"Fez Curso 1 ({curso1 or 'N/A'})": sub["fez_curso1"].sum() if curso1 else 0,
            f"Fez Curso 2 ({curso2 or 'N/A'})": sub["fez_curso2"].sum() if curso2 else 0,
            "Fez Ambos Cursos": sub["fez_ambos"].sum(),
            "Não Fez Ambos Cursos": (~sub["fez_ambos"]).sum(),
        })
    df_resumo_revenda = pd.DataFrame(resumo_revenda)

    # 2) Resumo por loja
    resumo_loja = []
    for (revenda, loja), sub in df_alvo.groupby(["grupo", "loja"], dropna=False):
        if pd.isna(loja):
            continue
        cnpj_loja = sub["cnpj_loja"].dropna().astype(str).iloc[0] if not sub["cnpj_loja"].dropna().empty else ""
        status_counts = sub["status"].value_counts().to_dict()
        resumo_loja.append({
            "Revenda": revenda,
            "Loja": loja,
            "CNPJ Loja": cnpj_loja,
            "Total Participantes": len(sub),
            "Ativo": status_counts.get("Ativo", 0),
            "Inativo": status_counts.get("Inativo", 0),
            "Pré-Cadastrado": status_counts.get("Pré-Cadastrado", 0),
            "Indicado Moderado": status_counts.get("Indicado Moderado", 0),
            "Em Moderação": status_counts.get("Em Moderação", 0),
            "Reprovado": status_counts.get("Reprovado", 0),
            "Bloqueado": status_counts.get("Bloqueado", 0),
            "Deu Aceite": sub["deu_aceite"].sum(),
            "Não Deu Aceite": (~sub["deu_aceite"]).sum(),
            "Fez Ambos Cursos": sub["fez_ambos"].sum(),
            "Não Fez Ambos Cursos": (~sub["fez_ambos"]).sum(),
        })
    df_resumo_loja = pd.DataFrame(resumo_loja)

    # 3) Detalhamento por participante
    cols_detalhe = [
        "cpf_limp", "nome", "cargo", "status", "grupo", "revenda_det", "loja",
        "cnpj_loja", "cnpj_revenda", "regional_loja", "regional_revenda",
        "data de aceite", "Ultimo Aceite Mensal", "data atualização",
        "deu_aceite", "fez_curso1", "fez_curso2", "fez_ambos",
    ]
    cols_detalhe = [c for c in cols_detalhe if c in df_alvo.columns]
    df_detalhe = df_alvo[cols_detalhe].copy()
    df_detalhe = df_detalhe.rename(columns={
        "cpf_limp": "CPF",
        "nome": "Nome",
        "cargo": "Cargo",
        "status": "Status",
        "grupo": "Revenda (Cadastro)",
        "revenda_det": "Revenda (Detalhamento)",
        "loja": "Loja",
        "cnpj_loja": "CNPJ Loja",
        "cnpj_revenda": "CNPJ Revenda",
        "regional_loja": "Regional Loja",
        "regional_revenda": "Regional Revenda",
        "data de aceite": "Data de Aceite",
        "Ultimo Aceite Mensal": "Último Aceite Mensal",
        "data atualização": "Data Atualização",
        "deu_aceite": "Deu Aceite",
        "fez_curso1": f"Fez Curso 1 ({curso1 or 'N/A'})",
        "fez_curso2": f"Fez Curso 2 ({curso2 or 'N/A'})",
        "fez_ambos": "Fez Ambos Cursos",
    })

    # --- Salva Excel ---
    hoje_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"relatorio_mm_guaibim_{hoje_str}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_resumo_revenda.to_excel(writer, sheet_name="Resumo_por_Revenda", index=False)
        df_resumo_loja.to_excel(writer, sheet_name="Resumo_por_Loja", index=False)
        df_detalhe.to_excel(writer, sheet_name="Detalhamento", index=False)

    logger.info(f"Relatório salvo em: {output_path}")

    # Log final
    logger.info("\nResumo:")
    for _, row in df_resumo_revenda.iterrows():
        logger.info(
            f"  {row['Revenda']}: {row['Total Participantes']} participantes, "
            f"{row['Deu Aceite']} aceites, {row['Fez Ambos Cursos']} fizeram ambos os treinamentos"
        )


if __name__ == "__main__":
    main()
