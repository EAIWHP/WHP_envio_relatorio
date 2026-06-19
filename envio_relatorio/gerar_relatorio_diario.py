#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório Diário Programa +TOP
------------------------------
Gera e envia por e-mail, todos os dias às 11:00, um resumo operacional com:
  • Cadastros (ativos / não ativos) por regional e por revenda
  • Treinamentos obrigatórios do mês (realizaram / não realizaram) por regional e revenda
  • Aceites mensais (aceitaram / não aceitaram) por regional
  • Insights automáticos no corpo do e-mail, prontos para cobrança das regionais

Fontes (pasta envio_relatorio/bases/):
  - cadastro.xlsx
  - Base_treinamentos.xlsx
  - WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx (ou arquivo com abas mensais)

Uso:
  python3 gerar_relatorio_diario.py
  python3 gerar_relatorio_diario.py --teste   # gera relatório local sem enviar e-mail

Agendamento (cron - Linux/WSL):
  0 11 * * * cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio && python3 gerar_relatorio_diario.py >> cron.log 2>&1
"""

import argparse
import base64
import io
import json
import logging
import os
import smtplib
import sys
from datetime import date, datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd

matplotlib.use("Agg")  # backend não interativo para cron/servers

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES GLOBAIS
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
LOG_DIR = BASE_DIR / "logs"
CONFIG_FILE = BASE_DIR / "config_email.json"

OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# Revendas excluídas/teste (conforme memória do projeto)
REVENDAS_EXCLUIR = {"EAI", "Whirlpool", "Novo Mundo", "ELETROMÓVEIS MARTINELLO"}

# Status que NÃO devem compor as métricas de cadastro (base operacional)
STATUS_CADASTRO_EXCLUIR = {"Inativo", "Bloqueado", "Reprovado", "Aguardando Aprovação"}

# Caminho da base do banco de participantes (usada para validar DataInclusao de Pré-Cadastrados)
BANCO_PARTICIPANTES_PATH = Path("/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx")

# Prazo em dias para considerar Pré-Cadastrado como Inativo (conforme regulamento)
PRAZO_PRE_CADASTRO_INATIVO = 90

# Filiais "B" que devem ser agrupadas com a revenda principal
FILIAIS_B = {
    "Becker B": "Becker",
    "Bemol B": "Bemol",
    "Colombo B": "Colombo",
    "Imperio B": "Imperio",
    "Koerich B": "Koerich",
    "Laser Eletro B": "Laser Eletro",
    "Multiloja B": "Multiloja",
    "Nosso Lar B": "Nosso Lar",
    "Ramsons B": "Ramsons",
    "Sipolatti B": "Sipolatti",
    "Taqi B": "Taqi",
    "Tele Rio B": "Tele Rio",
    "Armazem Mateus B": "Armazem Mateus",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"relatorio_{date.today():%Y%m%d}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("relatorio_top")


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ---------------------------------------------------------------------------
def nome_mes_pt_br(data_ref=None, ano_mes=None):
    """Retorna mês/ano em português (ex: Junho/2026)."""
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    if data_ref is None and ano_mes:
        data_ref = datetime.strptime(ano_mes, "%Y-%m")
    if data_ref is None:
        data_ref = date.today()
    return f"{meses[data_ref.month - 1]}/{data_ref.year}"


def limpar_cpf(cpf):
    """Normaliza CPF/CNPJ para texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    # Se vier como float (ex: 98470256068.0), converte para int primeiro
    try:
        if isinstance(cpf, float):
            cpf = int(cpf)
    except Exception:
        pass
    s = str(cpf).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(11)


def normalizar_revenda(grupo):
    """Agrupa filiais 'B' na revenda principal."""
    g = str(grupo).strip()
    return FILIAIS_B.get(g, g)


def regional_curta(regional):
    """Remove prefixo 'CONTA ' da regional."""
    r = str(regional).strip()
    return r.replace("CONTA ", "") if r.startswith("CONTA ") else r


def descobrir_aba_aceites(xl, ano_mes):
    """
    Procura a aba do mês de referência. Aceita nomes como 'JUN_2026', 'JUNHO_2026',
    '06_2026', '2026-06', etc. Se não encontrar, retorna a última aba disponível.
    """
    abas = xl.sheet_names
    ano, mes = ano_mes.split("-")
    mes_int = int(mes)
    mes_nome = {
        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ",
    }[mes_int]

    candidatos = [
        f"{mes_nome}_{ano}",
        f"{mes_nome.lower()}_{ano}",
        f"{ano}_{mes.zfill(2)}",
        f"{mes.zfill(2)}_{ano}",
        ano_mes,
        ano_mes.replace("-", "_"),
    ]

    for aba in abas:
        aba_limpa = aba.strip().replace(" ", "_").replace("-", "_")
        if aba_limpa.upper() in [c.upper() for c in candidatos]:
            return aba, False

    logger.warning(f"Aba de aceites para {ano_mes} não encontrada. Usando última aba disponível: {abas[-1]}")
    return abas[-1], True


def detectar_cursos_obrigatorios(df_trein, ano_mes):
    """
    Detecta os 2 cursos obrigatórios do mês a partir da trilha de conteúdo obrigatório.
    Retorna: (curso1, curso2, nome_curto1, nome_curto2)
    """
    mes_dt = pd.Period(ano_mes, freq="M")
    obr = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
        & (df_trein["Trilha"].str.contains("OBRIGAT", case=False, na=False))
    ].copy()

    if obr.empty:
        return None, None, None, None

    top = obr["Curso"].value_counts().head(2)
    if len(top) < 2:
        logger.warning(f"Foram encontrados apenas {len(top)} curso(s) obrigatório(s) para {ano_mes}")

    cursos = top.index.tolist()
    nomes_curtos = []
    for curso in cursos:
        partes = str(curso).split("|")
        nome = partes[-1].strip() if len(partes) > 1 else str(curso).strip()
        nomes_curtos.append(nome)

    return cursos[0] if len(cursos) > 0 else None, \
           cursos[1] if len(cursos) > 1 else None, \
           nomes_curtos[0] if len(nomes_curtos) > 0 else None, \
           nomes_curtos[1] if len(nomes_curtos) > 1 else None


def fig_to_base64(fig):
    """Converte figura matplotlib em string base64 para embutir no e-mail."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=120, bbox_inches="tight", facecolor="white")
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    plt.close(fig)
    return img_base64


def estilizar_tabela_html(df, destaque_coluna=None, destaque_menor_que_media=None, media=None):
    """Gera tabela HTML estilizada a partir de DataFrame."""
    if df.empty:
        return "<p><em>Sem dados para exibir.</em></p>"

    html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;">\n'
    html += "<thead><tr>"
    for col in df.columns:
        html += (
            f'<th style="border: 1px solid #cccccc; padding: 8px; background-color: #003366; '
            f'color: white; text-align: center;">{col}</th>'
        )
    html += "</tr></thead><tbody>\n"

    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = row[col]
            align = "right" if isinstance(val, (int, float)) else "left"

            # Destaca células abaixo da média, se solicitado
            bg = ""
            if destaque_coluna and col == destaque_coluna and destaque_menor_que_media is not None:
                try:
                    if float(val) < destaque_menor_que_media:
                        bg = 'background-color: #ffe6e6;'
                except (ValueError, TypeError):
                    pass

            html += f'<td style="border: 1px solid #cccccc; padding: 6px; text-align: {align}; {bg}">{val}</td>'
        html += "</tr>\n"

    html += "</tbody></table>"
    return html


# ---------------------------------------------------------------------------
# CARGA E PROCESSAMENTO DE BASES
# ---------------------------------------------------------------------------
def carregar_bases():
    """Carrega cadastro, treinamentos e aceites."""
    logger.info("Carregando bases...")

    # Cadastro
    cadastro_path = DATA_DIR / "cadastro.xlsx"
    xl_cad = pd.ExcelFile(cadastro_path)
    df_cad = pd.read_excel(cadastro_path, sheet_name=xl_cad.sheet_names[0])
    df_cad["cpf_limp"] = df_cad["cpf/cnpj"].apply(limpar_cpf)
    df_cad["revenda"] = df_cad["grupo"].apply(normalizar_revenda)
    df_cad["regional_curta"] = df_cad["regional"].apply(regional_curta)
    df_cad = df_cad[~df_cad["revenda"].isin(REVENDAS_EXCLUIR)].copy()
    # Pré-cadastrado há 90+ dias NA BASE DO BANCO é tratado como Inativo
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
        qtd_pre_inativo = mask_pre_inativo.sum()
        if qtd_pre_inativo > 0:
            df_cad.loc[mask_pre_inativo, "status"] = "Inativo"
            logger.info(f"{qtd_pre_inativo:,} registros 'Pré-Cadastrado' com 90+ dias de DataInclusao na base do banco convertidos para 'Inativo'")
        df_cad = df_cad.drop(columns=["DataInclusao_dt"])
    else:
        logger.warning(f"Base do banco não encontrada em {BANCO_PARTICIPANTES_PATH}. Regra de Pré-Cadastrado 90+ dias não aplicada.")
    # Remove status que não compõem a base operacional do relatório
    df_cad = df_cad[~df_cad["status"].isin(STATUS_CADASTRO_EXCLUIR)].copy()
    # Remove registros sem regional/revenda identificada
    df_cad = df_cad.dropna(subset=["regional_curta", "revenda"]).copy()
    df_cad = df_cad[df_cad["regional_curta"].str.lower() != "nan"].copy()
    df_cad = df_cad[df_cad["revenda"].str.lower() != "nan"].copy()
    logger.info(f"Cadastro: {len(df_cad):,} registros, {df_cad['cpf_limp'].nunique():,} CPFs únicos")

    # Treinamentos
    trein_path = DATA_DIR / "Base_treinamentos.xlsx"
    xl_trein = pd.ExcelFile(trein_path)
    df_trein = pd.read_excel(trein_path, sheet_name=xl_trein.sheet_names[0])
    df_trein["cpf_limp"] = df_trein["CPF"].apply(limpar_cpf)
    df_trein["Conclusão"] = pd.to_datetime(df_trein["Conclusão"], errors="coerce")
    df_trein = df_trein.merge(
        df_cad[["cpf_limp", "status", "regional_curta", "revenda"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="left",
    )
    logger.info(f"Treinamentos: {len(df_trein):,} registros")

    # Aceites — procura arquivo padrão ou qualquer arquivo com "Aceite" no nome
    aceite_path = DATA_DIR / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
    if not aceite_path.exists():
        candidatos = sorted(DATA_DIR.glob("*Aceite*.xlsx"))
        if candidatos:
            aceite_path = candidatos[-1]
            logger.info(f"Arquivo padrão de aceites não encontrado. Usando: {aceite_path.name}")
        else:
            raise FileNotFoundError("Arquivo de aceites não encontrado em envio_relatorio/bases/")

    xl_aceite = pd.ExcelFile(aceite_path)
    ano_mes_hoje = date.today().strftime("%Y-%m")
    aba_aceite, usou_ultima = descobrir_aba_aceites(xl_aceite, ano_mes_hoje)
    df_aceite = pd.read_excel(aceite_path, sheet_name=aba_aceite)
    df_aceite["cpf_limp"] = df_aceite["CPF"].apply(limpar_cpf)
    df_aceite["DataAceite"] = pd.to_datetime(df_aceite["DataAceite"], errors="coerce")
    df_aceite["mes_aceite"] = df_aceite["DataAceite"].dt.to_period("M")
    logger.info(f"Aceites: aba '{aba_aceite}' ({len(df_aceite):,} registros)")

    return {
        "cadastro": df_cad,
        "treinamentos": df_trein,
        "aceites": df_aceite,
        "aba_aceite": aba_aceite,
        "usou_ultima_aba": usou_ultima,
        "mes_referencia": ano_mes_hoje,
    }


# ---------------------------------------------------------------------------
# CÁLCULOS POR INDICADOR
# ---------------------------------------------------------------------------
def calcular_cadastros(df_cad):
    """Calcula cadastros por regional e por revenda."""
    # Por regional
    cad_reg = (
        df_cad.groupby("regional_curta")
        .agg(
            total=("cpf_limp", "nunique"),
            ativos=("status", lambda x: (x == "Ativo").sum()),
            nao_ativos=("status", lambda x: (x != "Ativo").sum()),
        )
        .reset_index()
    )
    cad_reg["pct_ativos"] = (cad_reg["ativos"] / cad_reg["total"] * 100).round(1)
    cad_reg = cad_reg.sort_values("pct_ativos", ascending=False)

    # Por revenda
    cad_rev = (
        df_cad.groupby(["regional_curta", "revenda"])
        .agg(
            total=("cpf_limp", "nunique"),
            ativos=("status", lambda x: (x == "Ativo").sum()),
            nao_ativos=("status", lambda x: (x != "Ativo").sum()),
        )
        .reset_index()
    )
    cad_rev["pct_ativos"] = (cad_rev["ativos"] / cad_rev["total"] * 100).round(1)
    cad_rev = cad_rev.sort_values("pct_ativos", ascending=True)

    return cad_reg, cad_rev


def calcular_treinamentos(df_trein, df_cad, ano_mes):
    """Calcula treinamentos por regional e por revenda (base = ativos)."""
    curso1, curso2, nome1, nome2 = detectar_cursos_obrigatorios(df_trein, ano_mes)

    if curso1 is None or curso2 is None:
        logger.warning("Não foi possível detectar os 2 cursos obrigatórios do mês")
        return None, None, (curso1, curso2, nome1, nome2)

    mes_dt = pd.Period(ano_mes, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ]

    cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique())
    cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique())
    cpf_ambos = cpf_curso1 & cpf_curso2

    logger.info(f"Treinamentos {ano_mes}: '{nome1}'={len(cpf_curso1)}, '{nome2}'={len(cpf_curso2)}, ambos={len(cpf_ambos)}")

    ativos = df_cad[df_cad["status"] == "Ativo"].copy()

    def calcular_grupo(grupo_df):
        cpfs = set(grupo_df["cpf_limp"].unique())
        realizaram = cpfs & cpf_ambos
        nao_realizaram = cpfs - cpf_ambos
        total = len(cpfs)
        pct = round(len(realizaram) / total * 100, 1) if total else 0
        return pd.Series({
            "total_ativos": int(total),
            "realizaram": int(len(realizaram)),
            "nao_realizaram": int(len(nao_realizaram)),
            "pct_realizaram": pct,
        })

    trein_reg = ativos.groupby("regional_curta").apply(calcular_grupo).reset_index()
    trein_reg = trein_reg.sort_values("pct_realizaram", ascending=False)

    trein_rev = ativos.groupby(["regional_curta", "revenda"]).apply(calcular_grupo).reset_index()
    trein_rev = trein_rev.sort_values("pct_realizaram", ascending=True)

    return trein_reg, trein_rev, (curso1, curso2, nome1, nome2)


def inferir_mes_da_aba(aba):
    """Tenta extrair ano-mês do nome da aba (ex: ABR_2026, JUN_2026)."""
    import re
    meses_map = {
        "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
        "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
    }
    aba_limpa = aba.strip().upper().replace("-", "_")
    # Procurar padrão MES_ANO
    match = re.search(r"([A-Z]{3})_(\d{4})", aba_limpa)
    if match:
        mes_nome, ano = match.groups()
        if mes_nome in meses_map:
            return pd.Period(f"{ano}-{meses_map[mes_nome]:02d}", freq="M")
    # Procurar padrão YYYY_MM
    match = re.search(r"(\d{4})_(\d{2})", aba_limpa)
    if match:
        return pd.Period(f"{match.group(1)}-{match.group(2)}", freq="M")
    return None


def calcular_aceites(df_aceite, df_cad, ano_mes, aba_aceite, usou_ultima_aba=False):
    """Calcula aceites mensais por regional (base = ativos)."""
    # Se usou a última aba disponível, inferir o mês a partir do nome da aba
    if usou_ultima_aba:
        mes_ref = inferir_mes_da_aba(aba_aceite)
        if mes_ref is None:
            mes_ref = df_aceite["mes_aceite"].max()
    else:
        mes_ref = pd.Period(ano_mes, freq="M")

    aceite_mes = df_aceite[df_aceite["mes_aceite"] == mes_ref]
    cpfs_aceitaram = set(aceite_mes["cpf_limp"].unique())
    logger.info(f"Aceites em {mes_ref}: {len(cpfs_aceitaram):,} CPFs únicos")

    ativos = df_cad[df_cad["status"] == "Ativo"].copy()

    aceite_reg = []
    for regional, gdf in ativos.groupby("regional_curta"):
        cpfs = set(gdf["cpf_limp"].unique())
        aceitaram = cpfs & cpfs_aceitaram
        nao_aceitaram = cpfs - cpfs_aceitaram
        total = len(cpfs)
        pct = round(len(aceitaram) / total * 100, 1) if total else 0
        aceite_reg.append({
            "regional": regional,
            "total_ativos": int(total),
            "aceitaram": int(len(aceitaram)),
            "nao_aceitaram": int(len(nao_aceitaram)),
            "pct_aceite": pct,
        })

    aceite_reg = pd.DataFrame(aceite_reg).sort_values("pct_aceite", ascending=False)
    return aceite_reg, mes_ref


# ---------------------------------------------------------------------------
# GRÁFICOS
# ---------------------------------------------------------------------------
def gerar_grafico_barras(df, x_col, y_col, titulo, cor="#003366", destaque_abaixo=None):
    """Gera gráfico de barras horizontal."""
    fig, ax = plt.subplots(figsize=(9, max(4, len(df) * 0.55)))
    fig.subplots_adjust(left=0.28)
    barras = ax.barh(df[x_col], df[y_col], color=cor)

    if destaque_abaixo is not None:
        for i, (val, bar) in enumerate(zip(df[y_col], barras)):
            if val < destaque_abaixo:
                bar.set_color("#d9534f")

    ax.set_xlabel("%")
    ax.set_title(titulo, fontsize=14, fontweight="bold", loc="left")
    ax.invert_yaxis()
    ax.grid(axis="x", linestyle="--", alpha=0.5)

    for i, v in enumerate(df[y_col]):
        ax.text(v + 1, i, f"{v}%", va="center", fontsize=10)

    return fig


def gerar_graficos(cad_reg, trein_reg, aceite_reg):
    """Gera os 3 gráficos principais e retorna dict base64."""
    graficos = {}

    if not cad_reg.empty:
        media_cad = cad_reg["pct_ativos"].mean()
        fig = gerar_grafico_barras(
            cad_reg.sort_values("pct_ativos", ascending=True),
            "regional_curta", "pct_ativos",
            f"Cadastros Ativos por Regional (média: {media_cad:.1f}%)",
            destaque_abaixo=media_cad,
        )
        graficos["cadastros"] = fig_to_base64(fig)

    if trein_reg is not None and not trein_reg.empty:
        media_trein = trein_reg["pct_realizaram"].mean()
        fig = gerar_grafico_barras(
            trein_reg.sort_values("pct_realizaram", ascending=True),
            "regional_curta", "pct_realizaram",
            f"Treinamentos Obrigatórios Realizados por Regional (média: {media_trein:.1f}%)",
            destaque_abaixo=media_trein,
        )
        graficos["treinamentos"] = fig_to_base64(fig)

    if not aceite_reg.empty:
        media_aceite = aceite_reg["pct_aceite"].mean()
        fig = gerar_grafico_barras(
            aceite_reg.sort_values("pct_aceite", ascending=True),
            "regional", "pct_aceite",
            f"Aceite Mensal por Regional (média: {media_aceite:.1f}%)",
            destaque_abaixo=media_aceite,
        )
        graficos["aceites"] = fig_to_base64(fig)

    return graficos


# ---------------------------------------------------------------------------
# INSIGHTS AUTOMÁTICOS
# ---------------------------------------------------------------------------
def gerar_insights(cad_reg, cad_rev, trein_reg, trein_rev, aceite_reg, dados):
    """Gera textos de insights prontos para o corpo do e-mail."""
    insights = []
    mes_nome = nome_mes_pt_br().lower()
    mes_aceite_nome = nome_mes_pt_br(ano_mes=str(dados["mes_aceite"]))

    # --- Cadastros ---
    if not cad_reg.empty:
        total_ativos = cad_reg["ativos"].sum()
        total_base = cad_reg["total"].sum()
        pct_geral = round(total_ativos / total_base * 100, 1)
        media_reg = cad_reg["pct_ativos"].mean()

        insights.append(
            f"Nossa performance em cadastros no Programa +TOP está em <strong>{pct_geral}%</strong> da base ativa. "
        )

        abaixo_media = cad_rev[cad_rev["pct_ativos"] < media_reg].sort_values("pct_ativos").head(10)
        if not abaixo_media.empty:
            linhas = [f"{r['revenda'].upper()} {r['pct_ativos']:.0f}% - {r['regional_curta']}" for _, r in abaixo_media.iterrows()]
            insights.append(
                "Ponto de atenção para as revendas que estão com o % de base menor do que a média do programa! "
                "Conto com o reforço das regionais para revertermos isso:<br>" + "<br>".join(linhas)
            )

    # --- Treinamentos ---
    if trein_reg is not None and not trein_reg.empty:
        pct_trein_geral = round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1)
        _, _, nome1, nome2 = dados["cursos_info"]
        insights.append(
            f"Os dois conteúdos de {mes_nome} já estão disponíveis para os vendedores. "
            f"No momento, apenas <strong>{pct_trein_geral}%</strong> da nossa base ativa fez os treinamentos.<br>"
            f"Conteúdo 1: {nome1}<br>Conteúdo 2: {nome2}"
        )

        media_trein = trein_reg["pct_realizaram"].mean()
        piores_rev = trein_rev[trein_rev["pct_realizaram"] < media_trein].sort_values("pct_realizaram").head(8)
        if not piores_rev.empty:
            linhas = [f"{r['revenda'].upper()} {r['pct_realizaram']:.0f}% - {r['regional_curta']}" for _, r in piores_rev.iterrows()]
            insights.append(
                "Alerta de treinamentos: regionais/revendas abaixo da média precisam de reforço:<br>" + "<br>".join(linhas)
            )

    # --- Aceites ---
    if not aceite_reg.empty:
        pct_aceite_geral = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1)
        insights.append(
            f"Sobre os aceites mensais de <strong>{mes_aceite_nome}</strong>, "
            f"<strong>{pct_aceite_geral}%</strong> dos participantes ativos deram aceite na campanha base."
        )

        media_aceite = aceite_reg["pct_aceite"].mean()
        abaixo_aceite = aceite_reg[aceite_reg["pct_aceite"] < media_aceite].sort_values("pct_aceite")
        if not abaixo_aceite.empty:
            linhas = [f"{r['regional']} {r['pct_aceite']:.0f}%" for _, r in abaixo_aceite.iterrows()]
            insights.append(
                "Regionais com aceite abaixo da média e que precisam de atenção:<br>" + "<br>".join(linhas)
            )

    return "<br><br>".join(insights)


# ---------------------------------------------------------------------------
# MONTAGEM DO E-MAIL
# ---------------------------------------------------------------------------
def montar_email_html(dados, graficos, teste=False):
    """Monta corpo do e-mail em HTML."""
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    aceite_reg = dados["aceites"]
    mes_nome = date.today().strftime("%B/%Y").capitalize()
    cursos_info = dados["cursos_info"]

    insights_html = gerar_insights(cad_reg, cad_rev, trein_reg, trein_rev, aceite_reg, dados)

    # Prepara tabelas
    tabela_cad_reg = estilizar_tabela_html(
        cad_reg.rename(columns={
            "regional_curta": "Regional",
            "total": "Total de Participantes",
            "ativos": "Usuários com +TOP",
            "nao_ativos": "Usuários sem +TOP",
            "pct_ativos": "% Ativos",
        })
    )

    tabela_cad_rev = estilizar_tabela_html(
        cad_rev.rename(columns={
            "regional_curta": "Regional",
            "revenda": "Revenda",
            "total": "Total",
            "ativos": "Usuários com +TOP",
            "nao_ativos": "Usuários sem +TOP",
            "pct_ativos": "% Ativos",
        }),
        destaque_coluna="% Ativos",
        destaque_menor_que_media=cad_reg["pct_ativos"].mean(),
    )

    trein_tabela = ""
    if trein_reg is not None and not trein_reg.empty:
        trein_tabela = estilizar_tabela_html(
            trein_reg.rename(columns={
                "regional_curta": "Regional",
                "total_ativos": "Base Ativa",
                "realizaram": "Realizado",
                "nao_realizaram": "Não Realizado",
                "pct_realizaram": "% Realizado",
            })
        )

    trein_rev_tabela = ""
    if trein_rev is not None and not trein_rev.empty:
        trein_rev_tabela = estilizar_tabela_html(
            trein_rev.rename(columns={
                "regional_curta": "Regional",
                "revenda": "Revenda",
                "total_ativos": "Base Ativa",
                "realizaram": "Realizado",
                "nao_realizaram": "Não Realizado",
                "pct_realizaram": "% Realizado",
            })
        )

    aceite_tabela = estilizar_tabela_html(
        aceite_reg.rename(columns={
            "regional": "Regional",
            "total_ativos": "Base Ativa",
            "aceitaram": "Aceitaram",
            "nao_aceitaram": "Não Aceitaram",
            "pct_aceite": "% Aceite",
        })
    )

    alerta_teste = "<p style='color:red;'><strong>[MODO TESTE - e-mail não enviado]</strong></p>" if teste else ""

    html = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; color: #333333; line-height: 1.5; }}
            h2 {{ color: #003366; border-bottom: 2px solid #003366; padding-bottom: 6px; }}
            h3 {{ color: #003366; margin-top: 24px; }}
            .insight {{ background-color: #f4f8fb; border-left: 4px solid #003366; padding: 12px; margin: 16px 0; }}
            .alerta {{ background-color: #fff3f3; border-left: 4px solid #d9534f; padding: 12px; margin: 16px 0; }}
            img {{ max-width: 100%; height: auto; margin: 12px 0; }}
        </style>
    </head>
    <body>
        {alerta_teste}
        <p>Bom dia, pessoal!<br>Tudo bem?</p>

        <p>Para melhorar nossos indicadores antes do final deste mês, compartilho atualização detalhada de cadastros, treinamentos e aceites:</p>

        <h2>CADASTROS</h2>
        <div class="insight">
            {insights_html.split('<br><br>')[0] if insights_html else ''}
        </div>

        <h3>Por Regional</h3>
        {tabela_cad_reg}
        <img src="cid:grafico_cadastros" alt="Gráfico Cadastros">

        <div class="alerta">
            {insights_html.split('<br><br>')[1] if '<br><br>' in insights_html and len(insights_html.split('<br><br>')) > 1 else ''}
        </div>

        <h3>Por Revenda</h3>
        {tabela_cad_rev}

        <h2>TREINAMENTOS</h2>
        <div class="insight">
            {insights_html.split('<br><br>')[2] if '<br><br>' in insights_html and len(insights_html.split('<br><br>')) > 2 else ''}
        </div>

        <h3>Por Regional</h3>
        {trein_tabela if trein_tabela else '<p><em>Não foi possível identificar os 2 treinamentos obrigatórios do mês.</em></p>'}
        <img src="cid:grafico_treinamentos" alt="Gráfico Treinamentos">

        <div class="alerta">
            {insights_html.split('<br><br>')[3] if '<br><br>' in insights_html and len(insights_html.split('<br><br>')) > 3 else ''}
        </div>

        <h3>Por Revenda</h3>
        {trein_rev_tabela if trein_rev_tabela else '<p><em>Sem dados de treinamentos.</em></p>'}

        <p><em>Lembrando que desde abril temos a nossa mecânica adicional, onde a cada 3 meses consecutivos que o vendedor concluir todos os treinamentos, ele ganha +100 pontos no programa.</em></p>

        <h2>ACEITES MENSAIS</h2>
        <div class="insight">
            {insights_html.split('<br><br>')[4] if '<br><br>' in insights_html and len(insights_html.split('<br><br>')) > 4 else ''}
        </div>

        <h3>Por Regional</h3>
        {aceite_tabela}
        <img src="cid:grafico_aceites" alt="Gráfico Aceites">

        <div class="alerta">
            {insights_html.split('<br><br>')[5] if '<br><br>' in insights_html and len(insights_html.split('<br><br>')) > 5 else ''}
        </div>

        <p>Att.,<br>Relatório Automático Programa +TOP</p>
    </body>
    </html>
    """
    return html


# ---------------------------------------------------------------------------
# ENVIO DE E-MAIL
# ---------------------------------------------------------------------------
def enviar_email(html_body, config, graficos, anexos=None, teste=False):
    """Envia e-mail HTML com imagens embutidas e anexos opcionais."""
    if teste:
        logger.info("MODO TESTE: e-mail não será enviado.")
        return False

    if not config:
        logger.error("Configuração de e-mail não encontrada. Verifique config_email.json")
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = config.get("assunto", f"Relatório Diário Programa +TOP - {date.today():%d/%m/%Y}")
    msg["From"] = formataddr((config.get("remetente_nome", "Relatório +TOP"), config["remetente_email"]))
    msg["To"] = ", ".join(config["destinatarios"])
    if config.get("cc"):
        msg["Cc"] = ", ".join(config["cc"])

    # Parte relacionada: HTML + imagens inline
    related = MIMEMultipart("related")
    related.attach(MIMEText(html_body, "html", _charset="utf-8"))

    # Imagens embutidas
    for cid, img_base64 in graficos.items():
        img_data = base64.b64decode(img_base64)
        mime_img = MIMEImage(img_data)
        mime_img.add_header("Content-ID", f"<grafico_{cid}>")
        mime_img.add_header("Content-Disposition", "inline", filename=f"grafico_{cid}.png")
        related.attach(mime_img)

    msg.attach(related)

    # Anexos
    if anexos:
        for caminho in anexos:
            with open(caminho, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {caminho.name}")
            msg.attach(part)

    # Envio
    try:
        server = smtplib.SMTP(config["smtp_host"], config["smtp_port"], timeout=30)
        server.starttls()
        server.login(config["remetente_email"], config["senha_app"])
        destinatarios = config["destinatarios"] + (config.get("cc") or [])
        server.sendmail(config["remetente_email"], destinatarios, msg.as_string())
        server.quit()
        logger.info(f"E-mail enviado com sucesso para: {', '.join(config['destinatarios'])}")
        return True
    except Exception as e:
        logger.error(f"Falha ao enviar e-mail: {e}")
        return False


def carregar_config_email():
    """Carrega configurações de e-mail do arquivo JSON.

    A senha de app pode vir da variável de ambiente SMTP_APP_PASSWORD,
    evitando deixar a senha salva no arquivo de configuração.
    """
    if not CONFIG_FILE.exists():
        return None
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
    env_senha = os.environ.get("SMTP_APP_PASSWORD")
    if env_senha:
        config["senha_app"] = env_senha
    return config


# ---------------------------------------------------------------------------
# EXPORTAÇÃO EXCEL
# ---------------------------------------------------------------------------
def salvar_relatorio_excel(dados, caminho):
    """Salva relatório consolidado em Excel com múltiplas abas."""
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    aceite_reg = dados["aceites"]

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        cad_reg.to_excel(writer, sheet_name="Cadastro_Regional", index=False)
        cad_rev.to_excel(writer, sheet_name="Cadastro_Revenda", index=False)
        if trein_reg is not None:
            trein_reg.to_excel(writer, sheet_name="Treinamento_Regional", index=False)
        if trein_rev is not None:
            trein_rev.to_excel(writer, sheet_name="Treinamento_Revenda", index=False)
        aceite_reg.to_excel(writer, sheet_name="Aceite_Regional", index=False)

        # Aba de resumo
        resumo = {
            "Indicador": [
                "Total de CPFs na base",
                "Total de CPFs ativos",
                "% Ativos geral",
                "Treinamentos realizaram (ambos cursos)",
                "% Treinamentos realizado (base ativos)",
                "Aceites mensais",
                "% Aceite mensal (base ativos)",
            ],
            "Valor": [
                cad_reg["total"].sum(),
                cad_reg["ativos"].sum(),
                round(cad_reg["ativos"].sum() / cad_reg["total"].sum() * 100, 1),
                trein_reg["realizaram"].sum() if trein_reg is not None else "N/A",
                round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1) if trein_reg is not None else "N/A",
                aceite_reg["aceitaram"].sum(),
                round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1),
            ],
        }
        pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)

    logger.info(f"Relatório Excel salvo em: {caminho}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Relatório Diário Programa +TOP")
    parser.add_argument("--teste", action="store_true", help="Gera relatório local sem enviar e-mail")
    args = parser.parse_args()

    logger.info("=" * 60)
    logger.info("Iniciando geração do relatório diário +TOP")
    logger.info(f"Modo: {'TESTE' if args.teste else 'PRODUÇÃO'}")

    try:
        bases = carregar_bases()
        df_cad = bases["cadastro"]
        df_trein = bases["treinamentos"]
        df_aceite = bases["aceites"]
        ano_mes = bases["mes_referencia"]

        # Cálculos
        cad_reg, cad_rev = calcular_cadastros(df_cad)
        trein_reg, trein_rev, cursos_info = calcular_treinamentos(df_trein, df_cad, ano_mes)
        aceite_reg, mes_aceite_ref = calcular_aceites(
            df_aceite, df_cad, ano_mes, bases["aba_aceite"], bases["usou_ultima_aba"]
        )

        dados = {
            "cadastros": (cad_reg, cad_rev),
            "treinamentos": (trein_reg, trein_rev),
            "aceites": aceite_reg,
            "cursos_info": cursos_info,
            "mes_aceite": mes_aceite_ref,
            "usou_ultima_aba": bases["usou_ultima_aba"],
        }

        # Gráficos
        graficos = gerar_graficos(cad_reg, trein_reg, aceite_reg)

        # Salvar Excel
        excel_path = OUTPUT_DIR / f"relatorio_top_{date.today():%Y%m%d}.xlsx"
        salvar_relatorio_excel(dados, excel_path)

        # Montar e-mail
        html = montar_email_html(dados, graficos, teste=args.teste)

        # Salvar cópia HTML
        html_path = OUTPUT_DIR / f"relatorio_top_{date.today():%Y%m%d}.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)

        # Enviar e-mail
        config = carregar_config_email()
        enviado = enviar_email(html, config, graficos, anexos=[excel_path], teste=args.teste)

        if args.teste:
            logger.info(f"Modo teste: relatório gerado em {excel_path} e {html_path}")
        elif enviado:
            logger.info("Relatório enviado com sucesso.")
        else:
            logger.error("Relatório gerado, mas não foi possível enviar o e-mail.")

    except Exception as e:
        logger.exception("Erro durante geração do relatório")
        sys.exit(1)


if __name__ == "__main__":
    main()
