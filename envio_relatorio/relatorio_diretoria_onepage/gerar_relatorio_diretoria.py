#!/usr/bin/env python3
"""
Relatório de Diretoria One Page +TOP
------------------------------------
Gera Excel V1 e JSON para relatório mensal de diretoria.

Estrutura baseada na planilha Staff Zanatta:
- Revenda, Cadastro, Treinamento, Aceite, Investimento Mês, Não Investimento,
  Vendas (+ comparativos LM/L3M/YOY), Ranking.

Dados financeiros (Investimento Mês / Vendas) lidos da planilha consolidada
Campanha_+TOP. Indicadores operacionais calculados a partir das bases do projeto.

Uso:
    python3 gerar_relatorio_diretoria.py
"""

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# Adiciona envio_relatorio ao path para reutilizar funções
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from gerar_relatorio_semanal import (
    limpar_cpf as limpar_cpf_semanal,
    normalizar_revenda_hierarquia,
    nome_revenda_exibicao,
    read_excel_robusto,
    carregar_hierarquias,
    detectar_cursos_obrigatorios,
)

# Adiciona painel-ranking-mensal ao path
sys.path.insert(0, "/home/thamiresvieira/projetos/painel-ranking-mensal")
from gerar_dados_reais import normalizar_revenda as normalizar_revenda_vendas

from config import (
    MES_REF,
    ANO_REF,
    MES_INICIO_ACUMULADO,
    ANO_INICIO_ACUMULADO,
    META_CADASTRO,
    META_TREINAMENTO,
    META_ACEITE,
    VALOR_PONTO,
    HIERARQUIA_DIR,
    CADASTRO_FILE,
    TREINAMENTOS_FILE,
    ACEITES_FILE,
    CAMPANHA_FILE,
    CAMPANHA_BEMOL_FILE,
    CAMPANHA_FILES_ACUMULADO,
    VENDAS_DIR,
    OUTPUT_DIR,
    MESES_PT,
    COLUNAS_RELATORIO,
    COR_PRIMARIA,
    COR_PRIMARIA_CLARA,
    COR_CINZA,
    COR_BRANCO,
    COR_TEXTO,
    COR_VERDE,
    COR_AMARELO,
    COR_VERMELHO,
    NOME_EXIBICAO_REVENDA,
)
from utils import (
    limpar_cpf,
    extrair_sku_curto,
    classificar_farol,
    calcular_ranking,
    formatar_pct,
    formatar_moeda,
    formatar_inteiro,
    normalizar_revenda_campanha,
    iterar_meses,
    nome_aba_aceites,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
)
logger = logging.getLogger("relatorio_diretoria")


# -----------------------------------------------------------------------------
# CARREGAMENTO DE BASES
# -----------------------------------------------------------------------------
def carregar_cadastro():
    """Carrega cadastro.xlsx e normaliza CPF, status e regional."""
    logger.info(f"Carregando cadastro: {CADASTRO_FILE}")
    df = read_excel_robusto(CADASTRO_FILE, sheet_name=0)
    # Mapeia colunas para uppercase
    colunas_upper = {str(c).strip().upper(): c for c in df.columns}

    # Detecta coluna de CPF
    cpf_col = None
    for chave in ["CPF/CNPJ", "CPF", "CNPJ", "CPFCNPJ"]:
        if chave in colunas_upper:
            cpf_col = colunas_upper[chave]
            break
    if cpf_col is None:
        raise ValueError("Coluna de CPF/CNPJ não encontrada no cadastro")
    df["cpf_limp"] = df[cpf_col].apply(limpar_cpf_semanal)

    # Detecta coluna de status
    status_col = colunas_upper.get("STATUS")
    if status_col is None:
        raise ValueError("Coluna STATUS não encontrada no cadastro")
    df["status"] = df[status_col].astype(str).str.strip().str.upper()

    # Detecta coluna de revenda e regional
    rev_col = None
    for chave in ["GRUPO", "REVENDA", "GRUPO CLUSTER"]:
        if chave in colunas_upper:
            rev_col = colunas_upper[chave]
            break
    regional_col = colunas_upper.get("REGIONAL")

    if rev_col:
        df["revenda_cad"] = df[rev_col].astype(str).str.strip()
    else:
        df["revenda_cad"] = ""
    if regional_col:
        df["regional_cad"] = df[regional_col].astype(str).str.strip()
    else:
        df["regional_cad"] = ""

    cols = ["cpf_limp", "status"]
    if regional_col:
        cols.append("regional_cad")
    return df[cols].copy()


def carregar_treinamentos():
    """Carrega base de treinamentos e normaliza CPF."""
    logger.info(f"Carregando treinamentos: {TREINAMENTOS_FILE}")
    df = read_excel_robusto(TREINAMENTOS_FILE, sheet_name=0)
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf_semanal)
    df["Curso"] = df["Curso"].astype(str).str.strip()
    df["Trilha"] = df["Trilha"].astype(str).str.strip().str.upper()
    df["Estado"] = df["Estado"].astype(str).str.strip().str.lower()
    df["Conclusão"] = pd.to_datetime(df["Conclusão"], errors="coerce", dayfirst=True)
    return df


def carregar_aceites():
    """Carrega aba do mês de referência do arquivo de aceites."""
    logger.info(f"Carregando aceites: {ACEITES_FILE}")
    xl = pd.ExcelFile(ACEITES_FILE)
    nome_mes = MESES_PT[MES_REF].upper()
    aba_padrao = f"{nome_mes}_{ANO_REF}"
    if aba_padrao in xl.sheet_names:
        aba = aba_padrao
    else:
        candidatas = [a for a in xl.sheet_names if nome_mes[:3] in a.upper()]
        aba = candidatas[0] if candidatas else xl.sheet_names[-1]
    logger.info(f"Aba de aceites: {aba}")
    df = read_excel_robusto(ACEITES_FILE, sheet_name=aba)
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf_semanal)
    return set(df["cpf_limp"].dropna().unique())


def carregar_aceites_acumulado(ano_inicio, mes_inicio, ano_fim, mes_fim):
    """Carrega CPFs que realizaram aceite em pelo menos um mês do período."""
    logger.info(f"Carregando aceites acumulados: {ano_inicio}-{mes_inicio:02d} a {ano_fim}-{mes_fim:02d}")
    xl = pd.ExcelFile(ACEITES_FILE)
    cpfs_aceite = set()
    meses_processados = []

    for ano, mes in iterar_meses(ano_inicio, mes_inicio, ano_fim, mes_fim):
        aba_padrao = nome_aba_aceites(ano, mes, MESES_PT)
        if aba_padrao in xl.sheet_names:
            aba = aba_padrao
        else:
            candidatas = [a for a in xl.sheet_names if MESES_PT[mes][:3].upper() in a.upper()]
            aba = candidatas[0] if candidatas else None

        if aba is None:
            logger.warning(f"Aba de aceites não encontrada para {mes:02d}/{ano}")
            continue

        try:
            df = read_excel_robusto(ACEITES_FILE, sheet_name=aba)
            if "CPF" not in df.columns:
                logger.warning(f"Aba {aba} não possui coluna CPF")
                continue
            cpfs_mes = set(df["CPF"].apply(limpar_cpf_semanal).dropna().unique())
            cpfs_aceite.update(cpfs_mes)
            meses_processados.append(f"{mes:02d}/{ano}")
            logger.info(f"  Aceite {mes:02d}/{ano}: {len(cpfs_mes)} CPFs únicos")
        except Exception as e:
            logger.warning(f"Erro ao carregar aceite {mes:02d}/{ano}: {e}")

    logger.info(f"Aceites acumulados processados ({', '.join(meses_processados)}): {len(cpfs_aceite)} CPFs únicos")
    return cpfs_aceite


def carregar_treinamentos_acumulado(ano_inicio, mes_inicio, ano_fim, mes_fim, df_trein=None):
    """
    Retorna CPFs ativos que concluíram os 2 cursos obrigatórios em pelo menos um mês do período.
    Se df_trein for None, carrega a base completa.
    """
    logger.info(f"Carregando treinamentos acumulados: {ano_inicio}-{mes_inicio:02d} a {ano_fim}-{mes_fim:02d}")
    if df_trein is None:
        df_trein = carregar_treinamentos()

    cpfs_treinados = set()
    meses_processados = []

    for ano, mes in iterar_meses(ano_inicio, mes_inicio, ano_fim, mes_fim):
        ano_mes = f"{ano}-{mes:02d}"
        try:
            curso1, curso2, *_ = detectar_cursos_obrigatorios(df_trein, ano_mes)
            if not curso1 or not curso2:
                logger.warning(f"Cursos obrigatórios não detectados para {ano_mes}")
                continue

            trein_mes = df_trein[
                (df_trein["Conclusão"].dt.to_period("M") == pd.Period(ano_mes, freq="M"))
                & (df_trein["Estado"] == "concluido")
            ]
            cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique()) if curso1 else set()
            cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique()) if curso2 else set()
            cpf_ambos = cpf_curso1 & cpf_curso2
            cpfs_treinados.update(cpf_ambos)
            meses_processados.append(f"{mes:02d}/{ano}")
            logger.info(f"  Treinamento {mes:02d}/{ano}: {len(cpf_ambos)} CPFs concluíram ambos os cursos")
        except Exception as e:
            logger.warning(f"Erro ao processar treinamentos {mes:02d}/{ano}: {e}")

    logger.info(f"Treinamentos acumulados processados ({', '.join(meses_processados)}): {len(cpfs_treinados)} CPFs únicos")
    return cpfs_treinados


# -----------------------------------------------------------------------------
# CÁLCULO DE INDICADORES OPERACIONAIS
# -----------------------------------------------------------------------------
def _chave_merge(nome):
    """Cria chave de merge em uppercase sem acentos para comparar nomes de revenda."""
    if pd.isna(nome):
        return ""
    s = str(nome).strip().upper()
    s = (
        s.replace("Á", "A").replace("É", "E").replace("Í", "I")
         .replace("Ó", "O").replace("Ú", "U").replace("Ã", "A")
         .replace("Õ", "O").replace("Ç", "C").replace("Ê", "E")
         .replace("Â", "A").replace("Ô", "O")
    )
    s = re.sub(r"\s+", " ", s)
    return s


def calcular_indicadores(df_hier, df_cad, df_trein, aceites_cpfs):
    """Calcula % Cadastro, % Treinamento e % Aceite por revenda."""
    regional_col = None
    for c in df_cad.columns:
        if str(c).strip().upper() == "REGIONAL":
            regional_col = c
            break

    # Enriquece hierarquia com status do cadastro e regional
    df = df_hier.merge(df_cad, on="cpf_limp", how="left")
    df["ativo"] = df["status"].fillna("").str.upper() == "ATIVO"
    df["regional"] = df["regional_cad"].fillna("").astype(str).str.strip()

    # Cursos obrigatórios do mês
    ano_mes = f"{ANO_REF}-{MES_REF:02d}"
    curso1, curso2, nome1, nome2, sku1, sku2 = detectar_cursos_obrigatorios(df_trein, ano_mes)
    logger.info(f"Cursos obrigatórios {ano_mes}: {nome1} ({sku1}), {nome2} ({sku2})")

    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == pd.Period(ano_mes, freq="M"))
        & (df_trein["Estado"] == "concluido")
    ]
    cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique()) if curso1 else set()
    cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique()) if curso2 else set()
    cpf_ambos = cpf_curso1 & cpf_curso2

    resultados = []
    for revenda, g in df.groupby("revenda"):
        total_hier = g["cpf_limp"].nunique()
        ativos = g[g["ativo"]]["cpf_limp"].nunique()
        pct_cadastro = (ativos / total_hier * 100) if total_hier else 0

        cpfs_ativos = set(g[g["ativo"]]["cpf_limp"].unique())
        treinaram = cpfs_ativos & cpf_ambos
        aceitaram = cpfs_ativos & aceites_cpfs
        aptos = cpfs_ativos & cpf_ambos & aceites_cpfs

        pct_treinamento = (len(treinaram) / ativos * 100) if ativos else 0
        pct_aceite = (len(aceitaram) / ativos * 100) if ativos else 0

        regional = ""
        if "regional" in g.columns and not g["regional"].dropna().empty:
            regional = g["regional"].dropna().mode().iloc[0]
        elif "regional_curta" in g.columns and not g["regional_curta"].dropna().empty:
            regional = g["regional_curta"].dropna().mode().iloc[0]

        resultados.append({
            "revenda": revenda,
            "revenda_key": _chave_merge(revenda),
            "regional": regional,
            "total_hier": total_hier,
            "ativos": ativos,
            "pct_cadastro": pct_cadastro,
            "pct_treinamento": pct_treinamento,
            "pct_aceite": pct_aceite,
            "cpfs_ativos": cpfs_ativos,
            "cpfs_aptos": aptos,
        })

    df_ind = pd.DataFrame(resultados)
    return df_ind, (curso1, curso2, sku1, sku2)


# -----------------------------------------------------------------------------
# PLANILHA CAMPANHA +TOP (FONTE FINANCEIRA)
# -----------------------------------------------------------------------------
def carregar_campanha_referencia():
    """Carrega aba Resumo da planilha Campanha_+TOP e normaliza nomes de revenda."""
    logger.info(f"Carregando planilha Campanha: {CAMPANHA_FILE}")
    df = read_excel_robusto(CAMPANHA_FILE, sheet_name="Resumo")

    df = df.rename(columns={
        "Revenda": "revenda_orig",
        "Qtd Peças": "qtd_pecas",
        "Pontos Vendas (Vendedor)": "pontos_vendas",
        "Total": "total",
    })

    df["revenda"] = df["revenda_orig"].apply(normalizar_revenda_campanha)
    df["revenda_key"] = df["revenda"].apply(_chave_merge)

    for col in ["qtd_pecas", "pontos_vendas", "total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def carregar_campanha_bemol():
    """Carrega planilha específica da Bemol e retorna linha da aba Resumo."""
    if not CAMPANHA_BEMOL_FILE.exists():
        logger.warning(f"Planilha Bemol não encontrada: {CAMPANHA_BEMOL_FILE}")
        return None

    logger.info(f"Carregando planilha Bemol: {CAMPANHA_BEMOL_FILE}")
    df = read_excel_robusto(CAMPANHA_BEMOL_FILE, sheet_name="Resumo")

    # Renomeia colunas para padrão amigável
    df = df.rename(columns={
        "Revenda": "revenda_orig",
        "Qtd Peças": "qtd_pecas",
        "Pontos Vendas (Vendedor)": "pontos_vendas",
        "Total": "total",
        "Aniversário Junho(Julho - parcial)": "aniversario",
        "Aniversário Junho (Julho - parcial)": "aniversario",
    })

    df["revenda"] = "Bemol"
    df["revenda_key"] = _chave_merge("Bemol")

    for col in ["qtd_pecas", "pontos_vendas", "total"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def mesclar_campanha_com_bemol(df_campanha, df_bemol):
    """Mescla dados gerais da Campanha com os dados específicos da Bemol."""
    if df_bemol is None or df_bemol.empty:
        return df_campanha

    # Remove linha de Bemol da campanha geral (se existir)
    df_campanha = df_campanha[df_campanha["revenda_key"] != _chave_merge("Bemol")].copy()

    # Adiciona dados da Bemol
    df_mesclado = pd.concat([df_campanha, df_bemol], ignore_index=True)
    return df_mesclado


def carregar_campanha_acumulado(files_dict):
    """
    Carrega e soma planilhas Campanha +TOP de múltiplos meses.
    files_dict: dict {(ano, mes): Path}
    Retorna DataFrame com colunas revenda, revenda_key, qtd_pecas, pontos_vendas, total.
    """
    logger.info("Carregando Campanha +TOP acumulada")
    dfs = []
    for (ano, mes), path in sorted(files_dict.items()):
        if not path.exists():
            logger.warning(f"Arquivo não encontrado: {path}")
            continue
        try:
            df = read_excel_robusto(path, sheet_name="Resumo")
            df = df.rename(columns={
                "Revenda": "revenda_orig",
                "Qtd Peças": "qtd_pecas",
                "Pontos Vendas (Vendedor)": "pontos_vendas",
                "Total": "total",
            })
            df["revenda"] = df["revenda_orig"].apply(normalizar_revenda_campanha)
            df["revenda_key"] = df["revenda"].apply(_chave_merge)
            for col in ["qtd_pecas", "pontos_vendas", "total"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            dfs.append(df[["revenda", "revenda_key", "qtd_pecas", "pontos_vendas", "total"]])
            logger.info(f"  Campanha {mes:02d}/{ano}: {len(df)} revendas")
        except Exception as e:
            logger.warning(f"Erro ao carregar Campanha {mes:02d}/{ano}: {e}")

    if not dfs:
        raise ValueError("Nenhuma planilha Campanha +TOP foi carregada para o período acumulado")

    df_total = pd.concat(dfs, ignore_index=True)
    df_agg = df_total.groupby(["revenda_key", "revenda"], as_index=False).agg({
        "qtd_pecas": "sum",
        "pontos_vendas": "sum",
        "total": "sum",
    })
    logger.info(f"Campanha acumulada: {len(df_agg)} revendas, {df_agg['qtd_pecas'].sum():,.0f} peças, R$ {df_agg['total'].sum():,.2f}")
    return df_agg


def carregar_vendas_acumulado(ano_inicio, mes_inicio, ano_fim, mes_fim):
    """Carrega e soma vendas processadas de out/2025 a jun/2026 por revenda."""
    logger.info(f"Carregando vendas acumuladas: {ano_inicio}-{mes_inicio:02d} a {ano_fim}-{mes_fim:02d}")
    vendas_por_rev = {}
    meses_processados = []

    for ano, mes in iterar_meses(ano_inicio, mes_inicio, ano_fim, mes_fim):
        arquivo = VENDAS_DIR / f"{ano}_{mes:02d}.xlsx"
        if not arquivo.exists():
            logger.warning(f"Arquivo de vendas não encontrado: {arquivo}")
            continue
        try:
            df = read_excel_robusto(arquivo, sheet_name="Export")
            df["cpf_limp"] = df["CPF"].apply(limpar_cpf_semanal)
            df["revenda"] = (
                df["Revenda"]
                .astype(str)
                .str.strip()
                .apply(normalizar_revenda_vendas)
                .apply(normalizar_revenda_hierarquia)
            )
            agg = df.groupby("revenda")["Vendas"].sum()
            for rev, qtd in agg.items():
                rev_key = _chave_merge(rev)
                vendas_por_rev[rev_key] = vendas_por_rev.get(rev_key, 0) + qtd
            meses_processados.append(f"{mes:02d}/{ano}")
            logger.info(f"  Vendas {mes:02d}/{ano}: {agg.sum():,.0f} peças")
        except Exception as e:
            logger.warning(f"Erro ao carregar vendas {mes:02d}/{ano}: {e}")

    logger.info(f"Vendas acumuladas processadas ({', '.join(meses_processados)}): {sum(vendas_por_rev.values()):,.0f} peças")
    return vendas_por_rev
def carregar_vendas_por_revenda(ano, mes):
    """Carrega vendas processadas do mês e agrega quantidade por revenda."""
    arquivo = Path("/home/thamiresvieira/projetos/Programa_mais_top/bases/Vendas Processadas") / f"{ano}_{mes:02d}.xlsx"
    if not arquivo.exists():
        logger.warning(f"Arquivo de vendas não encontrado: {arquivo}")
        return {}

    logger.info(f"Carregando vendas comparativas: {arquivo}")
    df = read_excel_robusto(arquivo, sheet_name="Export")
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf_semanal)
    df["revenda"] = (
        df["Revenda"]
        .astype(str)
        .str.strip()
        .apply(normalizar_revenda_vendas)
        .apply(normalizar_revenda_hierarquia)
    )

    agg = df.groupby("revenda")["Vendas"].sum().reset_index()
    agg["revenda_key"] = agg["revenda"].apply(_chave_merge)
    return dict(zip(agg["revenda_key"], agg["Vendas"]))


def calcular_comparativos_vendas():
    """Calcula LM, L3M e YOY de vendas (quantidade) por revenda."""
    comparativos = {}

    # LM
    if MES_REF == 1:
        ano_lm, mes_lm = ANO_REF - 1, 12
    else:
        ano_lm, mes_lm = ANO_REF, MES_REF - 1
    comparativos["lm"] = carregar_vendas_por_revenda(ano_lm, mes_lm)

    # L3M: últimos 3 meses incluindo o atual
    meses_l3m = []
    for i in range(2, -1, -1):
        m = MES_REF - i
        a = ANO_REF
        if m <= 0:
            m += 12
            a -= 1
        meses_l3m.append((a, m))

    soma_l3m = {}
    n_meses = 0
    for a, m in meses_l3m:
        dados_m = carregar_vendas_por_revenda(a, m)
        if dados_m:
            n_meses += 1
            for rev, qtd in dados_m.items():
                soma_l3m[rev] = soma_l3m.get(rev, 0) + qtd
    comparativos["l3m"] = {r: v / n_meses for r, v in soma_l3m.items()} if n_meses else {}

    # YOY
    ano_yoy = ANO_REF - 1
    comparativos["yoy"] = carregar_vendas_por_revenda(ano_yoy, MES_REF)

    return comparativos


# -----------------------------------------------------------------------------
# INVESTIMENTO, POTENCIAL E RANKING
# -----------------------------------------------------------------------------
def calcular_investimento_potencial(df_ind, df_campanha):
    """Calcula investimento real (Total Campanha), potencial e não investimento."""
    # Mapeia Total e Qtd Peças da Campanha por revenda_key
    total_campanha = dict(zip(df_campanha["revenda_key"], df_campanha["total"]))
    pontos_vendas_campanha = dict(zip(df_campanha["revenda_key"], df_campanha["pontos_vendas"]))
    qtd_pecas_campanha = dict(zip(df_campanha["revenda_key"], df_campanha["qtd_pecas"]))

    # Calcula pontuação por CPF ativo apto usando apenas Pontos Vendas (Vendedor)
    pontuacao_por_cpf = {}
    for _, row in df_ind.iterrows():
        rev_key = row["revenda_key"]
        pts_vendas = pontos_vendas_campanha.get(rev_key, 0)
        aptos = row["cpfs_aptos"]
        if aptos and pts_vendas > 0:
            media_por_apto = pts_vendas / len(aptos)
            for cpf in aptos:
                pontuacao_por_cpf[cpf] = media_por_apto

    # Média geral por ativo apto (fallback)
    todos_aptos = set()
    for _, row in df_ind.iterrows():
        todos_aptos.update(row["cpfs_aptos"])
    media_geral = sum(pontuacao_por_cpf.get(cpf, 0) for cpf in todos_aptos) / len(todos_aptos) if todos_aptos else 0

    resultados = []
    for _, row in df_ind.iterrows():
        rev_key = row["revenda_key"]
        inv = total_campanha.get(rev_key, 0) * VALOR_PONTO
        vendas = qtd_pecas_campanha.get(rev_key, 0)

        aptos = row["cpfs_aptos"]
        if aptos:
            media_revenda = sum(pontuacao_por_cpf.get(cpf, 0) for cpf in aptos) / len(aptos)
        else:
            media_revenda = 0

        if media_revenda == 0 and media_geral > 0:
            media_revenda = media_geral

        base_potencial = row["total_hier"]
        potencial = base_potencial * media_revenda
        nao_investimento = max(0, potencial - inv)

        resultados.append({
            "revenda_key": rev_key,
            "investimento": inv,
            "potencial_full": potencial,
            "nao_investimento": nao_investimento,
            "vendas": vendas,
        })

    return pd.DataFrame(resultados)


def adicionar_ranking(df):
    """Adiciona coluna de ranking (0-3 pontos)."""
    df["ranking"] = df.apply(
        lambda r: calcular_ranking(r["pct_cadastro"], r["pct_treinamento"], r["pct_aceite"]),
        axis=1,
    )
    return df


# -----------------------------------------------------------------------------
# CONSOLIDAÇÃO
# -----------------------------------------------------------------------------
def consolidar_relatorio(df_ind, df_inv, comparativos, df_campanha):
    """Consolida DataFrame final do relatório."""
    df = df_ind.merge(df_inv, on="revenda_key", how="left")
    df = df.merge(df_campanha[["revenda_key", "revenda"]], on="revenda_key", how="left")

    # Usa nome da Campanha se disponível, senão o da hierarquia
    df["revenda"] = df["revenda_y"].fillna(df["revenda_x"])

    # Aplica formatação de exibição dos nomes de revenda (case-insensitive)
    def _formatar_nome_revenda(nome):
        s = str(nome).strip()
        for chave, valor in NOME_EXIBICAO_REVENDA.items():
            if s.lower() == chave.lower():
                return valor
        return s

    df["revenda"] = df["revenda"].apply(_formatar_nome_revenda)

    df["vendas_lm"] = df["revenda_key"].map(comparativos.get("lm", {})).fillna(0)
    df["vendas_l3m"] = df["revenda_key"].map(comparativos.get("l3m", {})).fillna(0)
    df["vendas_yoy"] = df["revenda_key"].map(comparativos.get("yoy", {})).fillna(0)

    df["var_lm_pct"] = df.apply(
        lambda r: ((r["vendas"] - r["vendas_lm"]) / r["vendas_lm"] * 100) if r["vendas_lm"] else 0.0,
        axis=1,
    )
    df["var_yoy_pct"] = df.apply(
        lambda r: ((r["vendas"] - r["vendas_yoy"]) / r["vendas_yoy"] * 100) if r["vendas_yoy"] else 0.0,
        axis=1,
    )

    df = adicionar_ranking(df)

    # Ordenação: ranking desc, depois investimento desc
    df = df.sort_values(["ranking", "investimento"], ascending=[False, False])

    # Colunas finais (inclui total_hier e ativos para ponderação no JSON/Excel)
    df_final = df[[
        "revenda", "regional", "pct_cadastro", "pct_treinamento", "pct_aceite",
        "investimento", "nao_investimento", "vendas", "vendas_lm", "var_lm_pct",
        "vendas_l3m", "vendas_yoy", "var_yoy_pct", "ranking", "total_hier", "ativos",
    ]].copy()

    df_final.columns = COLUNAS_RELATORIO + ["Total de CPFs na Hierarquia", "Ativos"]
    return df_final


def _ajustar_largura_colunas(ws):
    """Ajusta largura das colunas para o tamanho do conteúdo."""
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except Exception:
                pass
        # Adiciona margem de 2 caracteres
        ws.column_dimensions[col_letter].width = min(max_length + 2, 60)


# -----------------------------------------------------------------------------
# GERAÇÃO DO EXCEL
# -----------------------------------------------------------------------------
def gerar_excel(df_final, skus_obrigatorios):
    """Gera arquivo Excel formatado."""
    wb = Workbook()
    ws = wb.active
    ws.title = "One Page"

    # Aba One Page usa apenas as colunas do relatório oficial
    df_onepage = df_final[COLUNAS_RELATORIO].copy()

    # Título
    ws.merge_cells("A1:N1")
    ws["A1"] = f"Relatório Geral Programa + TOP - {MESES_PT[MES_REF]}/{ANO_REF}"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=COR_BRANCO)
    ws["A1"].fill = PatternFill(start_color=COR_PRIMARIA, end_color=COR_PRIMARIA, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # Cabeçalho
    header_fill = PatternFill(start_color=COR_PRIMARIA_CLARA, end_color=COR_PRIMARIA_CLARA, fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color=COR_TEXTO)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col_idx, col_name in enumerate(COLUNAS_RELATORIO, 1):
        cell = ws.cell(row=3, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Dados
    fills = {
        "verde": PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid"),
        "amarelo": PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid"),
        "vermelho": PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid"),
    }

    for row_idx, row in enumerate(df_onepage.itertuples(index=False), 4):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = Font(name="Arial", size=10, color=COR_TEXTO)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

            if col_idx == 1:  # Revenda
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 2:  # Regional
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx in (3, 4, 5):  # Percentuais
                cell.number_format = "0.0%"
                cell.value = value / 100 if pd.notna(value) else 0
                meta = [META_CADASTRO, META_TREINAMENTO, META_ACEITE][col_idx - 3]
                cor = classificar_farol(value, meta * 100)
                cell.fill = fills.get(cor, PatternFill())
            elif col_idx in (6, 7):  # Moeda
                cell.number_format = '"R$" #,##0.00'
            elif col_idx in (8, 9, 11, 12):  # Inteiros
                cell.number_format = "#,##0"
            elif col_idx in (10, 13):  # Variações
                cell.number_format = "0.0%"
                cell.value = value / 100 if pd.notna(value) else 0
                if pd.notna(value):
                    cor = "verde" if value >= 0 else "vermelho"
                    cell.fill = fills.get(cor, PatternFill())
            elif col_idx == 14:  # Ranking
                cell.font = Font(name="Arial", size=10, bold=True, color=COR_PRIMARIA)

    # Linha de total geral com percentuais ponderados
    total_row_idx = 4 + len(df_onepage)
    ws.merge_cells(start_row=total_row_idx, start_column=1, end_row=total_row_idx, end_column=2)
    total_cell = ws.cell(row=total_row_idx, column=1, value="TOTAL GERAL")
    total_cell.font = Font(name="Arial", size=10, bold=True, color=COR_TEXTO)
    total_cell.fill = PatternFill(start_color=COR_PRIMARIA_CLARA, end_color=COR_PRIMARIA_CLARA, fill_type="solid")
    total_cell.alignment = Alignment(horizontal="left", vertical="center")
    total_cell.border = thin_border

    # Cálculos ponderados
    total_hier_sum = df_final["Total de CPFs na Hierarquia"].sum()
    ativos_por_rev = df_final["Total de CPFs na Hierarquia"] * df_final["Cadastro (%)"] / 100
    ativos_sum = ativos_por_rev.sum()

    pct_cad_ponderado = (df_final["Cadastro (%)"] * df_final["Total de CPFs na Hierarquia"]).sum() / total_hier_sum if total_hier_sum else 0
    pct_trein_ponderado = (df_final["Treinamento (%)"] * ativos_por_rev).sum() / ativos_sum if ativos_sum else 0
    pct_aceite_ponderado = (df_final["Aceite (%)"] * ativos_por_rev).sum() / ativos_sum if ativos_sum else 0

    inv_total = df_onepage["Investimento Mês (R$)"].sum()
    nao_inv_total = df_onepage["Não Investimento (R$)"].sum()
    vendas_total = df_onepage["Vendas"].sum()
    vendas_lm_total = df_onepage["Vendas LM"].sum()
    vendas_l3m_total = df_onepage["Vendas L3M"].sum()
    vendas_yoy_total = df_onepage["Vendas YOY"].sum()

    var_lm_total = ((vendas_total - vendas_lm_total) / vendas_lm_total * 100) if vendas_lm_total else 0.0
    var_yoy_total = ((vendas_total - vendas_yoy_total) / vendas_yoy_total * 100) if vendas_yoy_total else 0.0

    totais = [
        None, None,  # Revenda + Regional (mescladas)
        pct_cad_ponderado,
        pct_trein_ponderado,
        pct_aceite_ponderado,
        inv_total,
        nao_inv_total,
        vendas_total,
        vendas_lm_total,
        var_lm_total,
        vendas_l3m_total,
        vendas_yoy_total,
        var_yoy_total,
        None,  # Ranking
    ]

    for col_idx, value in enumerate(totais[2:], 3):  # começa na coluna C (3)
        if value is None:
            continue
        cell = ws.cell(row=total_row_idx, column=col_idx, value=value)
        cell.font = Font(name="Arial", size=10, bold=True, color=COR_TEXTO)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

        if col_idx in (3, 4, 5):  # Percentuais
            cell.number_format = "0.0%"
            cell.value = value / 100 if pd.notna(value) else 0
            meta = [META_CADASTRO, META_TREINAMENTO, META_ACEITE][col_idx - 3]
            cor = classificar_farol(value, meta * 100)
            cell.fill = fills.get(cor, PatternFill())
        elif col_idx in (6, 7):  # Moeda
            cell.number_format = '"R$" #,##0.00'
        elif col_idx in (8, 9, 11, 12):  # Inteiros
            cell.number_format = "#,##0"
        elif col_idx in (10, 13):  # Variações
            cell.number_format = "0.0%"
            cell.value = value / 100 if pd.notna(value) else 0
            if pd.notna(value):
                cor = "verde" if value >= 0 else "vermelho"
                cell.fill = fills.get(cor, PatternFill())

    # Ajusta larguras
    widths = [22, 15, 14, 16, 12, 22, 22, 12, 12, 12, 12, 12, 12, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Congela painéis
    ws.freeze_panes = "A4"

    # Aba Dados (todas as colunas, incluindo total_hier)
    ws_dados = wb.create_sheet("Dados")
    for r in dataframe_to_rows(df_final, index=False, header=True):
        ws_dados.append(r)
    _ajustar_largura_colunas(ws_dados)

    # Aba Regras
    ws_regras = wb.create_sheet("Regras")
    regras = [
        ["Relatório de Diretoria +TOP", f"Período: {MESES_PT[MES_REF]}/{ANO_REF}"],
        ["", ""],
        ["Indicador", "Fórmula"],
        ["% Cadastro", "Ativos no cadastro / Total na hierarquia ativa"],
        ["% Treinamento", "Ativos que concluíram os 2 cursos obrigatórios / Base ativa"],
        ["% Aceite", "Ativos com aceite mensal / Base ativa"],
        ["Investimento Mês", "Total de pontuação da revenda na planilha Campanha +TOP (1 ponto = R$ 1,00)"],
        ["Potencial Full", "Total de CPFs na hierarquia ativa × média de pontuação por ativo apto (Campanha +TOP)"],
        ["Não Investimento", "Potencial Full - Investimento Mês"],
        ["Vendas", "Vendas Totais Bases Revendas (coluna da planilha Campanha +TOP)"],
        ["LM", "Mês anterior"],
        ["L3M", "Média dos últimos 3 meses"],
        ["YOY", "Mesmo mês do ano anterior"],
        ["Var LM (%)", "Variação percentual do mês atual vs mês anterior"],
        ["Var YOY (%)", "Variação percentual do mês atual vs mesmo mês do ano anterior"],
        ["Ranking", "0 a 3 pontos. 1 ponto para cada meta atingida: Cadastro ≥85%, Treinamento ≥70%, Aceite ≥70%."],
        ["", "Exemplo 1: 98% cadastro, 87% treinamento, 95% aceite = 3 pontos."],
        ["", "Exemplo 2: 90% cadastro, 50% treinamento, 95% aceite = 2 pontos."],
        ["", "Exemplo 3: 45% cadastro, 8% treinamento, 76% aceite = 1 ponto."],
        ["", "Exemplo 4: 45% cadastro, 8% treinamento, 60% aceite = 0 pontos."],
        ["", ""],
        ["Farol", "Verde = meta atingida | Amarelo = entre 50% e a meta | Vermelho = abaixo de 50%"],
        ["Farol Variação", "Verde = variação ≥ 0% | Vermelho = variação < 0%"],
        ["Média dos percentuais", "Média ponderada pelo Total de CPFs na Hierarquia Ativa de cada revenda."],
        ["SKUs obrigatórios", ", ".join([s for s in skus_obrigatorios if s]) if skus_obrigatorios else "nenhum"],
    ]
    for r in regras:
        ws_regras.append(r)
    _ajustar_largura_colunas(ws_regras)

    # Salva
    arquivo = OUTPUT_DIR / f"Relatorio_Programa_Geral_+TOP_{MESES_PT[MES_REF]}_{ANO_REF}_V1.xlsx"
    wb.save(arquivo)
    logger.info(f"Excel salvo em: {arquivo}")
    return arquivo


# -----------------------------------------------------------------------------
# GERAÇÃO DO JSON
# -----------------------------------------------------------------------------
def gerar_json(df_final):
    """Gera JSON para uso futuro na one page HTML."""
    dados = {
        "periodo": f"{ANO_REF}-{MES_REF:02d}",
        "label": f"{MESES_PT[MES_REF]}/{ANO_REF}",
        "metas": {
            "cadastro": META_CADASTRO,
            "treinamento": META_TREINAMENTO,
            "aceite": META_ACEITE,
        },
        "revendas": json.loads(df_final.to_json(orient="records")),
    }
    arquivo = OUTPUT_DIR / f"dados_{ANO_REF}-{MES_REF:02d}.json"
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    logger.info(f"JSON salvo em: {arquivo}")
    return arquivo


# -----------------------------------------------------------------------------
# VALIDAÇÕES
# -----------------------------------------------------------------------------
def validar_resultados(df_final, df_campanha):
    """Executa validações básicas."""
    logger.info("Iniciando validações...")
    valido = True

    # 1. Percentuais entre 0 e 100
    for col in ["Cadastro (%)", "Treinamento (%)", "Aceite (%)"]:
        if (df_final[col] < 0).any() or (df_final[col] > 100).any():
            logger.error(f"Percentuais fora do intervalo em {col}")
            valido = False

    # 2. Não investimento não negativo
    if (df_final["Não Investimento (R$)"] < 0).any():
        logger.error("Não Investimento negativo encontrado")
        valido = False

    # 3. Soma do investimento bate com total da Campanha
    total_campanha = df_campanha["total"].sum() * VALOR_PONTO
    total_investimento = df_final["Investimento Mês (R$)"].sum()
    if abs(total_campanha - total_investimento) > 0.01:
        logger.error(f"Divergência no investimento: campanha={total_campanha}, relatório={total_investimento}")
        valido = False
    else:
        logger.info(f"Investimento total validado: R$ {total_investimento:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))

    # 4. Soma das vendas bate com Qtd Peças da Campanha
    total_vendas_campanha = df_campanha["qtd_pecas"].sum()
    total_vendas_rel = df_final["Vendas"].sum()
    if abs(total_vendas_campanha - total_vendas_rel) > 0.01:
        logger.error(f"Divergência nas vendas: campanha={total_vendas_campanha}, relatório={total_vendas_rel}")
        valido = False
    else:
        logger.info(f"Vendas totais validadas: {total_vendas_rel:,.0f}")

    # 5. Total de revendas
    logger.info(f"Total de revendas no relatório: {len(df_final)}")

    if valido:
        logger.info("Todas as validações passaram.")
    else:
        logger.warning("Validação encontrou problemas.")
    return valido


# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
def main():
    logger.info(f"Gerando relatório de diretoria para {MESES_PT[MES_REF]}/{ANO_REF}")

    # Carrega bases
    df_hier, _ = carregar_hierarquias()
    df_cad = carregar_cadastro()
    df_trein = carregar_treinamentos()
    aceites_cpfs = carregar_aceites()

    # Calcula indicadores operacionais
    df_ind, cursos_info = calcular_indicadores(df_hier, df_cad, df_trein, aceites_cpfs)
    curso1, curso2, sku1, sku2 = cursos_info
    skus_obrigatorios = [s for s in [sku1, sku2] if s]

    # Carrega planilha Campanha +TOP (fonte financeira)
    df_campanha = carregar_campanha_referencia()

    # Carrega planilha específica da Bemol e mescla
    df_bemol = carregar_campanha_bemol()
    df_campanha = mesclar_campanha_com_bemol(df_campanha, df_bemol)

    # Calcula investimento, potencial e não investimento
    df_inv = calcular_investimento_potencial(df_ind, df_campanha)

    # Calcula comparativos de vendas
    comparativos = calcular_comparativos_vendas()

    # Consolida
    df_final = consolidar_relatorio(df_ind, df_inv, comparativos, df_campanha)

    # Valida
    validar_resultados(df_final, df_campanha)

    # Gera arquivos
    gerar_excel(df_final, skus_obrigatorios)
    gerar_json(df_final)

    logger.info("Relatório gerado com sucesso.")


if __name__ == "__main__":
    main()
