#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatórios MM Varejo/MM Atacado e Guaibim — atualizado (vendas e aceites de maio/2026)
--------------------------------------------------------------------------------------
Gera dois arquivos Excel separados:
  1) MM Varejo + MM Atacado
  2) Guaibim

Período de referência: 01/05/2026 a 16/06/2026

Colunas principais:
  • Quantidade por status
  • Quantidade de aceites mensais (referência: maio/2026)
  • CNPJ da loja / nome da loja
  • Quantidade que realizou os 2 treinamentos obrigatórios do mês
  • Quantidade que NÃO realizou os treinamentos
  • Quantidade de produtos vendidos por participante (maio/2026)
  • Data da venda
  • Data de inclusão do cadastro (tabela banco)
  • Aba de detalhamento por participante
  • Aba de vendas detalhadas

Regra especial de Pré-Cadastro:
  - Cruzar CPFs do cadastro com o arquivo do banco.
  - Para participantes com status "Pré-Cadastrado" no cadastro, manter apenas
    aqueles cujos CPFs também existem no banco.
  - Demais status são mantidos normalmente.

Fontes:
  - bases/Tabelas/cadastro.xlsx
  - validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx
  - bases/Tabelas/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx
  - bases/Tabelas/WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx  (aba MAIO_2026)
  - bases/Tabelas/Base_treinamentos.xlsx
  - bases/Vendas Processadas/2026_05.xlsx

Uso:
  python3 gerar_relatorios_mm_guaibim_atualizado.py
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES
# ---------------------------------------------------------------------------
BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")
TABELAS_DIR = BASE_DIR / "bases" / "Tabelas"
VENDAS_DIR = BASE_DIR / "bases" / "Vendas Processadas"
OUTPUT_DIR = BASE_DIR / "envio_relatorio" / "relatorios_gerados"
BANCO_PARTICIPANTES_PATH = Path(
    "/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx"
)

OUTPUT_DIR.mkdir(exist_ok=True)

DATA_INICIO = pd.Timestamp("2026-05-01")
DATA_FIM = pd.Timestamp("2026-06-16")
MES_REFERENCIA = "2026-05"
ABA_ACEITE = "MAIO_2026"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("relatorios_mm_guaibim")


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
    s = (
        str(cpf)
        .strip()
        .replace("'", "")
        .replace(".", "")
        .replace("-", "")
        .replace("/", "")
        .replace(" ", "")
    )
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(11)


def limpar_cnpj(cnpj):
    """Normaliza CNPJ para texto com 14 dígitos."""
    if pd.isna(cnpj):
        return None
    try:
        if isinstance(cnpj, float):
            cnpj = int(cnp)
    except Exception:
        pass
    s = (
        str(cnpj)
        .strip()
        .replace("'", "")
        .replace(".", "")
        .replace("-", "")
        .replace("/", "")
        .replace(" ", "")
    )
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(14)


def converter_data_venda(serie):
    """Converte coluna 'data venda' corretamente (data ou número serial Excel)."""
    numerico = pd.to_numeric(serie, errors="coerce")

    def _converter(val, num):
        if pd.isna(val):
            return pd.NaT
        if pd.notna(num) and 1 <= num <= 100000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(num))
        try:
            return pd.to_datetime(val, errors="coerce")
        except Exception:
            return pd.NaT

    return pd.Series(
        [_converter(v, n) for v, n in zip(serie, numerico)], index=serie.index
    )


def detectar_cursos_obrigatorios(df_trein, ano_mes):
    """Detecta os 2 cursos obrigatórios mais frequentes do mês."""
    mes_dt = pd.Period(ano_mes, freq="M")
    obr = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
        & (df_trein["Trilha"].str.contains("OBRIGAT", case=False, na=False))
    ].copy()

    if obr.empty:
        return None, None

    top = obr["Curso"].value_counts().head(2)
    return tuple(top.index.tolist()) if len(top) >= 2 else (top.index[0], None)


# ---------------------------------------------------------------------------
# CARGA DE DADOS
# ---------------------------------------------------------------------------
def carregar_dados():
    logger.info("=" * 70)
    logger.info("Carregando bases atualizadas")

    # --- Cadastro ---
    cadastro_path = TABELAS_DIR / "cadastro.xlsx"
    df_cad = pd.read_excel(cadastro_path)
    df_cad["cpf_limp"] = df_cad["cpf/cnpj"].apply(limpar_cpf)
    logger.info(f"Cadastro carregado: {len(df_cad):,} registros")

    # --- Banco ---
    df_banco = pd.read_excel(BANCO_PARTICIPANTES_PATH)
    df_banco["cpf_limp"] = df_banco["Cpf"].apply(limpar_cpf)
    df_banco["DataInclusao_dt"] = pd.to_datetime(
        df_banco["DataInclusao"], errors="coerce"
    )
    banco_clean = df_banco[["cpf_limp", "DataInclusao_dt"]].drop_duplicates(
        subset=["cpf_limp"], keep="first"
    )
    cpfs_no_banco = set(df_banco["cpf_limp"].dropna().unique())
    logger.info(f"Banco carregado: {len(df_banco):,} registros ({len(cpfs_no_banco):,} CPFs únicos)")

    # --- Regra de Pré-Cadastro ---
    mask_pre = df_cad["status"] == "Pré-Cadastrado"
    mask_pre_no_banco = mask_pre & ~df_cad["cpf_limp"].isin(cpfs_no_banco)
    qtd_removidos = mask_pre_no_banco.sum()
    if qtd_removidos > 0:
        logger.info(
            f"{qtd_removidos:,} Pré-Cadastrados removidos por não estarem no banco"
        )

    # Mantém todos os não-pré-cadastrados + pré-cadastrados que existem no banco
    df_cad = df_cad[~(mask_pre & ~df_cad["cpf_limp"].isin(cpfs_no_banco))].copy()
    df_cad = df_cad.merge(banco_clean, on="cpf_limp", how="left")
    logger.info(f"Cadastro após regra de Pré-Cadastro: {len(df_cad):,} registros")

    # --- Cadastro Revenda x Loja x Regional ---
    # O arquivo atualizado possui o cabeçalho deslocado em relação aos dados.
    # Leitura sem cabeçalho e mapeamento manual das colunas.
    det_path = TABELAS_DIR / "cadastro_revenda_loja_regional.xlsx"
    df_det = pd.read_excel(det_path, header=None)

    # Descarta a linha de cabeçalho original e aplica nomes corretos
    df_det = df_det.iloc[1:].reset_index(drop=True)
    colunas_corretas = [
        "Cpf",           # 0
        "Nome",          # 1
        "Cargo",         # 2
        "Celular",       # 3
        "Email",         # 4
        "Status",        # 5
        "Data_Inclusao_Extra",  # 6 (coluna vazia no arquivo)
        "Vazio1",        # 7
        "Flag",          # 8
        "Data_Inclusao", # 9
        "CNPJ_Revenda",  # 10
        "Revenda",       # 11
        "Regional_da_Revenda",  # 12
        "CNPJ_Loja",     # 13
        "Loja",          # 14
        "Regional_da_Loja",     # 15
    ]
    df_det.columns = colunas_corretas[: df_det.shape[1]]

    df_det["cpf_limp"] = df_det["Cpf"].apply(limpar_cpf)
    df_det["CNPJ_Loja"] = df_det["CNPJ_Loja"].apply(limpar_cnpj)
    df_det["CNPJ_Revenda"] = df_det["CNPJ_Revenda"].apply(limpar_cnpj)
    df_det = df_det.rename(
        columns={
            "Revenda": "revenda_det",
            "Loja": "loja",
            "CNPJ_Loja": "cnpj_loja",
            "CNPJ_Revenda": "cnpj_revenda",
            "Regional_da_Loja": "regional_loja",
            "Regional_da_Revenda": "regional_revenda",
        }
    )

    # Para revendas sem loja cadastrada, preenche com "Matriz" e CNPJ da revenda
    mask_sem_loja = df_det["loja"].isna() | (df_det["loja"].astype(str).str.strip() == "")
    df_det.loc[mask_sem_loja, "loja"] = "Matriz"
    df_det.loc[mask_sem_loja & df_det["cnpj_loja"].isna(), "cnpj_loja"] = df_det.loc[
        mask_sem_loja & df_det["cnpj_loja"].isna(), "cnpj_revenda"
    ]

    logger.info(f"Detalhamento revenda/loja carregado: {len(df_det):,} registros")

    # --- Aceites (mês de referência) ---
    aceite_path = TABELAS_DIR / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
    xl_aceite = pd.ExcelFile(aceite_path)
    if ABA_ACEITE not in xl_aceite.sheet_names:
        raise ValueError(f"Aba {ABA_ACEITE} não encontrada no arquivo de aceites. Abas: {xl_aceite.sheet_names}")

    df_aceite = pd.read_excel(aceite_path, sheet_name=ABA_ACEITE)
    df_aceite["cpf_limp"] = df_aceite["CPF"].apply(limpar_cpf)
    df_aceite["DataAceite"] = pd.to_datetime(df_aceite["DataAceite"], errors="coerce")
    aceitaram = set(df_aceite["cpf_limp"].dropna().unique())
    logger.info(f"Aceites ({ABA_ACEITE}): {len(aceitaram):,} CPFs únicos")

    # --- Treinamentos ---
    trein_path = TABELAS_DIR / "Base_treinamentos.xlsx"
    df_trein = pd.read_excel(trein_path)
    df_trein["cpf_limp"] = df_trein["CPF"].apply(limpar_cpf)
    df_trein["Conclusão"] = pd.to_datetime(df_trein["Conclusão"], errors="coerce")
    df_trein["Estado_norm"] = df_trein["Estado"].str.lower()
    logger.info(f"Treinamentos carregado: {len(df_trein):,} registros")

    # Detecta cursos obrigatórios de maio/2026
    curso1, curso2 = detectar_cursos_obrigatorios(df_trein, MES_REFERENCIA)
    if curso1 is None:
        logger.warning("Não foi possível detectar cursos obrigatórios de maio/2026")
    else:
        logger.info(f"Cursos obrigatórios detectados: '{curso1}' e '{curso2 or 'N/A'}'")

    # --- Vendas ---
    vendas_path = VENDAS_DIR / "2026_05.xlsx"
    df_vendas = pd.read_excel(vendas_path)
    if "CPF" in df_vendas.columns:
        df_vendas["cpf_limp"] = df_vendas["CPF"].apply(limpar_cpf)
    if "data venda" in df_vendas.columns:
        df_vendas["data_venda"] = converter_data_venda(df_vendas["data venda"])
    logger.info(f"Vendas maio/2026 carregadas: {len(df_vendas):,} registros")

    # Filtra vendas pelo período (01/05 a 16/06)
    df_vendas_periodo = df_vendas[
        (df_vendas["data_venda"] >= DATA_INICIO)
        & (df_vendas["data_venda"] <= DATA_FIM)
    ].copy()
    logger.info(f"Vendas no período: {len(df_vendas_periodo):,} registros")

    return {
        "cadastro": df_cad,
        "detalhamento": df_det,
        "aceite": df_aceite,
        "treinamentos": df_trein,
        "vendas": df_vendas_periodo,
        "cursos": (curso1, curso2),
    }


# ---------------------------------------------------------------------------
# PROCESSAMENTO POR GRUPO DE REVENDAS
# ---------------------------------------------------------------------------
def processar_grupo(df_cad, df_det, df_aceite, df_trein, df_vendas, cursos, revendas_alvo, nome_relatorio):
    """Processa um grupo de revendas e retorna os DataFrames de saída."""
    curso1, curso2 = cursos

    # Filtra revendas alvo
    df_alvo = df_cad[df_cad["grupo"].isin(revendas_alvo)].copy()
    logger.info(f"{nome_relatorio}: {len(df_alvo):,} registros após filtro de revenda")

    # Cruza com detalhamento de loja
    df_alvo = df_alvo.merge(
        df_det[
            [
                "cpf_limp",
                "revenda_det",
                "loja",
                "cnpj_loja",
                "cnpj_revenda",
                "regional_loja",
                "regional_revenda",
            ]
        ].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="left",
    )

    # Flag de aceite
    aceitaram = set(df_aceite["cpf_limp"].dropna().unique())
    df_alvo["deu_aceite"] = df_alvo["cpf_limp"].isin(aceitaram)

    # Flags de treinamento
    if curso1:
        trein_mes = df_trein[
            (df_trein["Conclusão"] >= DATA_INICIO)
            & (df_trein["Conclusão"] <= DATA_FIM)
            & (df_trein["Estado_norm"] == "concluido")
        ].copy()
        cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique())
        df_alvo["fez_curso1"] = df_alvo["cpf_limp"].isin(cpf_curso1)
    else:
        df_alvo["fez_curso1"] = False

    if curso2:
        trein_mes = df_trein[
            (df_trein["Conclusão"] >= DATA_INICIO)
            & (df_trein["Conclusão"] <= DATA_FIM)
            & (df_trein["Estado_norm"] == "concluido")
        ].copy()
        cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique())
        df_alvo["fez_curso2"] = df_alvo["cpf_limp"].isin(cpf_curso2)
    else:
        df_alvo["fez_curso2"] = False

    df_alvo["fez_ambos"] = df_alvo["fez_curso1"] & df_alvo["fez_curso2"]

    # Vendas por participante
    df_vendas_alvo = df_vendas[df_vendas["cpf_limp"].isin(df_alvo["cpf_limp"])].copy()
    vendas_por_cpf = (
        df_vendas_alvo.groupby("cpf_limp")
        .agg(
            total_produtos_vendidos=("Vendas", "sum"),
            primeira_data_venda=("data_venda", "min"),
            ultima_data_venda=("data_venda", "max"),
        )
        .reset_index()
    )
    df_alvo = df_alvo.merge(vendas_por_cpf, on="cpf_limp", how="left")
    df_alvo["total_produtos_vendidos"] = df_alvo["total_produtos_vendidos"].fillna(0)

    # -----------------------------------------------------------------------
    # RESUMO POR REVENDA
    # -----------------------------------------------------------------------
    resumo_revenda = []
    for revenda in sorted(revendas_alvo):
        sub = df_alvo[df_alvo["grupo"] == revenda]
        if sub.empty:
            continue
        status_counts = sub["status"].value_counts().to_dict()
        resumo_revenda.append(
            {
                "Revenda": revenda,
                "Total Participantes": len(sub),
                "Ativo": status_counts.get("Ativo", 0),
                "Inativo": status_counts.get("Inativo", 0),
                "Pré-Cadastrado": status_counts.get("Pré-Cadastrado", 0),
                "Indicado Moderado": status_counts.get("Indicado Moderado", 0),
                "Em Moderação": status_counts.get("Em Moderação", 0),
                "Reprovado": status_counts.get("Reprovado", 0),
                "Bloqueado": status_counts.get("Bloqueado", 0),
                "Deu Aceite": int(sub["deu_aceite"].sum()),
                "Não Deu Aceite": int((~sub["deu_aceite"]).sum()),
                f"Fez Curso 1 ({curso1 or 'N/A'})": int(sub["fez_curso1"].sum()) if curso1 else 0,
                f"Fez Curso 2 ({curso2 or 'N/A'})": int(sub["fez_curso2"].sum()) if curso2 else 0,
                "Fez Ambos Cursos": int(sub["fez_ambos"].sum()),
                "Não Fez Ambos Cursos": int((~sub["fez_ambos"]).sum()),
                "Participantes Com Venda": int((sub["total_produtos_vendidos"] > 0).sum()),
                "Total Produtos Vendidos": int(sub["total_produtos_vendidos"].sum()),
            }
        )
    df_resumo_revenda = pd.DataFrame(resumo_revenda)

    # -----------------------------------------------------------------------
    # RESUMO POR LOJA
    # -----------------------------------------------------------------------
    resumo_loja = []
    for (revenda, loja), sub in df_alvo.groupby(["grupo", "loja"], dropna=False):
        if pd.isna(loja):
            continue
        cnpj_loja = (
            sub["cnpj_loja"].dropna().astype(str).iloc[0]
            if not sub["cnpj_loja"].dropna().empty
            else ""
        )
        status_counts = sub["status"].value_counts().to_dict()
        resumo_loja.append(
            {
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
                "Deu Aceite": int(sub["deu_aceite"].sum()),
                "Não Deu Aceite": int((~sub["deu_aceite"]).sum()),
                "Fez Ambos Cursos": int(sub["fez_ambos"].sum()),
                "Não Fez Ambos Cursos": int((~sub["fez_ambos"]).sum()),
                "Participantes Com Venda": int((sub["total_produtos_vendidos"] > 0).sum()),
                "Total Produtos Vendidos": int(sub["total_produtos_vendidos"].sum()),
            }
        )
    df_resumo_loja = pd.DataFrame(resumo_loja)

    # -----------------------------------------------------------------------
    # DETALHAMENTO POR PARTICIPANTE
    # -----------------------------------------------------------------------
    cols_detalhe = [
        "cpf_limp",
        "nome",
        "cargo",
        "status",
        "grupo",
        "revenda_det",
        "loja",
        "cnpj_loja",
        "cnpj_revenda",
        "regional_loja",
        "regional_revenda",
        "data de aceite",
        "Ultimo Aceite Mensal",
        "data atualização",
        "DataInclusao_dt",
        "deu_aceite",
        "fez_curso1",
        "fez_curso2",
        "fez_ambos",
        "total_produtos_vendidos",
        "primeira_data_venda",
        "ultima_data_venda",
    ]
    cols_detalhe = [c for c in cols_detalhe if c in df_alvo.columns]
    df_detalhe = df_alvo[cols_detalhe].copy()

    df_detalhe = df_detalhe.rename(
        columns={
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
            "DataInclusao_dt": "Data Inclusão (Banco)",
            "deu_aceite": "Deu Aceite",
            "fez_curso1": f"Fez Curso 1 ({curso1 or 'N/A'})",
            "fez_curso2": f"Fez Curso 2 ({curso2 or 'N/A'})",
            "fez_ambos": "Fez Ambos Cursos",
            "total_produtos_vendidos": "Total Produtos Vendidos",
            "primeira_data_venda": "Primeira Data Venda",
            "ultima_data_venda": "Última Data Venda",
        }
    )

    # -----------------------------------------------------------------------
    # VENDAS DETALHADAS
    # -----------------------------------------------------------------------
    cols_vendas = ["cpf_limp", "Nome", "Revenda", "CNPJ", "data_venda", "Vendas", "SKU + produto"]
    cols_vendas = [c for c in cols_vendas if c in df_vendas_alvo.columns]
    df_vendas_detalhe = df_vendas_alvo[cols_vendas].copy()
    df_vendas_detalhe = df_vendas_detalhe.rename(
        columns={
            "cpf_limp": "CPF",
            "Nome": "Nome Participante",
            "Revenda": "Revenda (Venda)",
            "CNPJ": "CNPJ Loja (Venda)",
            "data_venda": "Data Venda",
            "Vendas": "Quantidade Vendida",
            "SKU + produto": "SKU + Produto",
        }
    )

    return {
        "resumo_revenda": df_resumo_revenda,
        "resumo_loja": df_resumo_loja,
        "detalhe": df_detalhe,
        "vendas": df_vendas_detalhe,
    }


# ---------------------------------------------------------------------------
# SALVAR RELATÓRIO
# ---------------------------------------------------------------------------
def salvar_relatorio(dfs, nome_arquivo):
    hoje_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"{nome_arquivo}_{hoje_str}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dfs["resumo_revenda"].to_excel(writer, sheet_name="Resumo_por_Revenda", index=False)
        dfs["resumo_loja"].to_excel(writer, sheet_name="Resumo_por_Loja", index=False)
        dfs["detalhe"].to_excel(writer, sheet_name="Detalhamento", index=False)
        if not dfs["vendas"].empty:
            dfs["vendas"].to_excel(writer, sheet_name="Vendas_Detalhadas", index=False)

    logger.info(f"Relatório salvo: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    dados = carregar_dados()

    # Relatório 1: MM Varejo + MM Atacado
    logger.info("\n" + "=" * 70)
    logger.info("Processando MM Varejo + MM Atacado")
    dfs_mm = processar_grupo(
        dados["cadastro"],
        dados["detalhamento"],
        dados["aceite"],
        dados["treinamentos"],
        dados["vendas"],
        dados["cursos"],
        {"MM Varejo", "MM Atacado"},
        "MM Varejo + MM Atacado",
    )
    path_mm = salvar_relatorio(dfs_mm, "relatorio_mm_varejo_atacado")

    # Relatório 2: Guaibim
    logger.info("\n" + "=" * 70)
    logger.info("Processando Guaibim")
    dfs_guaibim = processar_grupo(
        dados["cadastro"],
        dados["detalhamento"],
        dados["aceite"],
        dados["treinamentos"],
        dados["vendas"],
        dados["cursos"],
        {"Guaibim"},
        "Guaibim",
    )
    path_guaibim = salvar_relatorio(dfs_guaibim, "relatorio_guaibim")

    # Log final
    logger.info("\n" + "=" * 70)
    logger.info("RESUMO FINAL")
    logger.info(f"  MM Varejo + MM Atacado: {path_mm}")
    logger.info(f"  Guaibim: {path_guaibim}")

    for nome, dfs in [("MM Varejo + MM Atacado", dfs_mm), ("Guaibim", dfs_guaibim)]:
        logger.info(f"\n  {nome}:")
        for _, row in dfs["resumo_revenda"].iterrows():
            logger.info(
                f"    {row['Revenda']}: {row['Total Participantes']} participantes, "
                f"{row['Deu Aceite']} aceites, {row['Fez Ambos Cursos']} fizeram ambos os treinamentos, "
                f"{row['Participantes Com Venda']} com venda"
            )


if __name__ == "__main__":
    main()
