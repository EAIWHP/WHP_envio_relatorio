#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório Semanal Programa +TOP
-------------------------------
Gera e envia por e-mail, toda semana, um resumo operacional com:
  • Cadastros (ativos / não ativos) por regional e por revenda
  • Treinamentos obrigatórios do mês (realizaram / não realizaram) por regional e revenda
  • Aceites mensais (aceitaram / não aceitaram) por regional
  • Comparação com a semana anterior (evolução, pontos de atenção)
  • Insights automáticos no corpo do e-mail
  • Link para o Drive com versão completa em Excel
  • Envio para a cliente + todos os regionais
  • E-mail individual por regional contendo apenas os dados daquela regional

Fontes (pasta envio_relatorio/bases/):
  - cadastro.xlsx
  - Base_treinamentos.xlsx
  - WHP_Aceite_Mensal_*.xlsx (arquivo com abas mensais)
  - emails_regionais.xlsx (Regional, Email) [opcional, para envio por regional]

Uso:
  python3 gerar_relatorio_semanal.py
  python3 gerar_relatorio_semanal.py --teste   # gera relatório local sem enviar e-mail

Agendamento (cron - Linux/WSL):
  # Toda segunda-feira às 11:00
  0 11 * * 1 cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio && python3 gerar_relatorio_semanal.py >> cron.log 2>&1
"""

import argparse
import base64
import io
import json
import logging
import os
import smtplib
import subprocess
import sys
from datetime import date, datetime, timedelta
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
SNAPSHOT_DIR = BASE_DIR / "snapshots"
CONFIG_FILE = BASE_DIR / "config_email.json"

OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
SNAPSHOT_DIR.mkdir(exist_ok=True)

# Revendas excluídas/teste (conforme memória do projeto)
REVENDAS_EXCLUIR = {"EAI", "Whirlpool", "Novo Mundo", "ELETROMÓVEIS MARTINELLO"}

# Status que NÃO devem compor as métricas de cadastro (base operacional)
STATUS_CADASTRO_EXCLUIR = {"Inativo", "Bloqueado", "Reprovado", "Aguardando Aprovação"}

# Caminho da base do banco de participantes (usada para validar DataInclusao de Pré-Cadastrados)
BANCO_PARTICIPANTES_PATH = Path("/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx")

# Pasta com as hierarquias mensais por revenda (fonte correta de CPF -> revenda/loja/CNPJ)
HIERARQUIA_DIR = DATA_DIR / "bases_cadastro_hierarquia"

# Prazo em dias para considerar Pré-Cadastrado como Inativo (conforme regulamento)
PRAZO_PRE_CADASTRO_INATIVO = 90

# Metas/objetivos ideais para os indicadores (%)
META_CADASTRO = 85.0
META_TREINAMENTOS = 70.0
META_ACEITES = 70.0

# Mapeamento de nomes de revenda para exibição (evita siglas)
NOME_REVENDA_EXIBICAO = {
    "CDA": "Casas da Água",
}

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

# Normalização de nomes de lojas/filiais para revenda principal (usada nas hierarquias)
MAPA_REVENDA_PRINCIPAL = {
    # Casas da Água
    "CASAS DA AGUA": "Casas da Água",
    "CDA": "Casas da Água",
    # Armazem Mateus
    "ELETRO MATEUS": "Armazem Mateus",
    "ARMAZEM MATEUS": "Armazem Mateus",
    # MM
    "MERCADOMOVEIS VAREJO": "MM Varejo",
    "MERCADOMOVEIS": "MM Atacado",
    # Ramsons
    "RAMSONS": "Ramsons",
    # Solar
    "SOLAR COMERCIO E AGROINDUSTRIA LTDA": "Solar",
    "SOLAR MOVEIS E ELETROS": "Solar Magazine",
    # Outros nomes exatos da hierarquia para normalizar
    "GAZIN ATACADO": "Gazin Atacado",
    "GAZIN ONLINE": "Gazin Online",
    "GAZIN": "Gazin Varejo",
    "LOJAS GUAIBIM": "Guaibim",
    "JMAHFUZ": "Jmahfuz",
    "IMPERIO": "Imperio",
    "LASER ELETRO": "Laser Eletro",
    "NOSSO LAR LOJAS DE DEPTOS LTDA": "Nosso Lar",
    "MILLENA MOVEIS": "Millena",
    "BECKER": "Becker",
    "BEMOL": "Bemol",
    "FORMOSA": "Formosa",
    "ESTRELA": "Estrela",
    "ANGELONI": "Angeloni",
    "SIPOLATTI": "Sipolatti",
    "TAQI": "Taqi",
    "TELE RIO": "Tele Rio",
    "TELERIO": "Tele Rio",
    "ZEMA": "Zema",
    "ELETROZEMA": "Zema",
    "KOERICH": "Koerich",
    "LEBES": "Lebes",
    "DREBES": "Lebes",
    "MULTILOJA": "Multiloja",
    "HORFRAN": "Multiloja",
    "MAGAZAN": "Magazan",
    "LOJAS BECKER LTDA": "Becker",
    "MOVEIS ESTRELA": "Estrela",
    "ZENIR": "Zenir",
}

# Mapeamento SKU/Código -> Nome descritivo do curso obrigatório do mês
# Atualizado conforme catálogo Whirlpool/Brastemp/Consul.
NOMES_CURSOS = {
    "CZD12": "Ar Condicionado Consul",
    "BRM46": "Geladeira Consul Inverse",
    "BMC29": "Micro-ondas Consul",
    "CRM44M": "Geladeira Consul Frost Free",
    "BRM44": "Geladeira Consul Frost Free",
    "CWN15": "Lavadora Consul",
    "CWN13": "Lavadora Consul",
    "BRE66": "Refrigerador Electrolux",
    "BRE57": "Refrigerador Electrolux",
    "BWJ14": "Lava-Louças Brastemp",
    "CRA30M": "Ar Condicionado Consul",
    "CFO4ZAB": "Fogão Consul",
    "BRM62": "Geladeira Brastemp",
    "CRM56M": "Geladeira Consul",
    "CRM53M": "Geladeira Consul",
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


def normalizar_revenda_hierarquia(nome_revenda):
    """
    Normaliza nome de revenda vindo das hierarquias para a revenda principal.
    Ex: 'Zenir Camocim' -> 'Zenir', 'ELETRO MATEUS' -> 'Armazem Mateus'.
    """
    if pd.isna(nome_revenda):
        return nome_revenda
    nome = str(nome_revenda).strip()
    # Remove acentos e espacos multiplos
    nome = nome.upper()
    import re
    nome = re.sub(r"\s+", " ", nome)
    nome = nome.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    nome = nome.replace("Â", "A").replace("Ê", "E").replace("Ô", "O")
    nome = nome.replace("Ã", "A").replace("Ç", "C")
    nome_upper = nome

    # Mapeamento direto
    if nome_upper in MAPA_REVENDA_PRINCIPAL:
        return MAPA_REVENDA_PRINCIPAL[nome_upper]

    # Heuristica: se comecar com nome de revenda conhecido, usa o nome principal
    for chave, principal in sorted(MAPA_REVENDA_PRINCIPAL.items(), key=lambda x: -len(x[0])):
        chave_limpa = re.sub(r"\s+", " ", chave.upper())
        if nome_upper.startswith(chave_limpa):
            return principal

    # Heuristica 2: se contiver nome de revenda conhecido (ex: SOLAR no meio)
    for chave, principal in sorted(MAPA_REVENDA_PRINCIPAL.items(), key=lambda x: -len(x[0])):
        chave_limpa = re.sub(r"\s+", " ", chave.upper())
        if chave_limpa in nome_upper:
            return principal

    return nome_revenda.strip() if isinstance(nome_revenda, str) else nome_revenda


def nome_revenda_exibicao(revenda):
    """Retorna nome amigavel da revenda para exibicao (ex: CDA -> Casas da Agua)."""
    rev = str(revenda).strip() if not pd.isna(revenda) else ""
    return NOME_REVENDA_EXIBICAO.get(rev, rev)


def regional_curta(regional):
    """Remove prefixo 'CONTA ' da regional."""
    r = str(regional).strip()
    return r.replace("CONTA ", "") if r.startswith("CONTA ") else r


def regional_title_case(regional):
    """Retorna nome da regional em Title Case (ex: COMPRA DIRETA -> Compra Direta)."""
    r = str(regional).strip()
    return r.title()



def read_excel_robusto(path, **kwargs):
    """
    Lê arquivo Excel robustamente, removendo regras de formatação/filtro
    que costumam quebrar o openpyxl em arquivos gerados por terceiros.
    """
    import zipfile
    import xml.etree.ElementTree as ET
    from io import BytesIO

    namespaces = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    }
    tags_remover = [
        "{%(main)s}dataValidations",
        "{%(main)s}conditionalFormatting",
        "{%(main)s}autoFilter",
        "{%(main)s}extLst",
    ]

    try:
        return pd.read_excel(path, **kwargs)
    except Exception as e:
        logger.warning(f"Falha ao ler {path.name} normalmente: {e}. Tentando leitura robusta...")

    try:
        with zipfile.ZipFile(path, "r") as zin:
            buffer = BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.startswith("xl/worksheets/sheet") and not item.filename.endswith(".rels"):
                        try:
                            root = ET.fromstring(data)
                            for tag_pattern in tags_remover:
                                tag = tag_pattern % namespaces
                                for elem in list(root.findall(tag)):
                                    root.remove(elem)
                            data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
                        except Exception as xml_err:
                            logger.warning(f"Não foi possível limpar XML de {item.filename}: {xml_err}")
                    zout.writestr(item, data)
        buffer.seek(0)
        return pd.read_excel(buffer, **kwargs)
    except Exception as e2:
        logger.error(f"Leitura robusta também falhou para {path.name}: {e2}")
        raise


def nome_revenda_exibicao(revenda):
    """Retorna nome amigável da revenda para exibição (ex: CDA -> Casas da Água)."""
    return NOME_REVENDA_EXIBICAO.get(str(revenda).strip(), str(revenda).strip())


def carregar_hierarquias():
    """
    Consolida todos os arquivos de hierarquia em bases_cadastro_hierarquia.
    Retorna DataFrame com CPF -> revenda/loja/cnpj/cargo/desligado.
    """
    if not HIERARQUIA_DIR.exists():
        logger.warning(f"Pasta de hierarquias não encontrada: {HIERARQUIA_DIR}")
        return None

    dfs = []
    mapeamento_colunas = {
        "REVENDA": "revenda",
        "COD LOJA": "cod_loja",
        "CODLOJA": "cod_loja",
        "CNPJ": "cnpj",
        "CPF": "cpf",
        "NOME": "nome_hier",
        "VENDEDOR": "vendedor",
        "GERENTE DE LOJA": "gerente_loja",
        "GERENTE REGIONAL": "gerente_regional",
        "DIRETOR": "diretor",
        "DESLIGADO": "desligado",
        "CARGO": "cargo_hier",
    }

    for arquivo in sorted(HIERARQUIA_DIR.glob("*.xlsx")):
        nome = arquivo.name.upper()
        # Ignora a base do banco e arquivos de sistema
        if "PROD_WHP" in nome or nome.startswith("~"):
            continue
        try:
            xl = read_excel_robusto(arquivo, sheet_name=None)
            primeira_aba = list(xl.keys())[0]
            df = xl[primeira_aba]
        except Exception as e:
            logger.warning(f"Não foi possível ler {arquivo.name}: {e}")
            continue

        df.columns = [str(c).strip().upper() for c in df.columns]
        rename = {c: mapeamento_colunas[c] for c in df.columns if c in mapeamento_colunas}
        df = df.rename(columns=rename)

        if "cpf" not in df.columns:
            logger.warning(f"{arquivo.name} não possui coluna CPF. Ignorado.")
            continue

        df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
        # Normaliza colunas de flag SIM/NÃO para maiúsculo sem acento
        for col in ["vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper().str.replace("Ã", "A")

        dfs.append(df)
        logger.info(f"Hierarquia {arquivo.name}: {len(df)} registros")

    if not dfs:
        logger.warning("Nenhum arquivo de hierarquia válido encontrado.")
        return None

    df_hier = pd.concat(dfs, ignore_index=True)

    # Normaliza nomes de revenda individualmente
    df_hier["revenda"] = df_hier["revenda"].apply(normalizar_revenda_hierarquia)

    # Normaliza CNPJ para numeros e cria raiz (8 primeiros digitos) para agrupar filiais/lojas
    if "cnpj" in df_hier.columns:
        df_hier["cnpj_limpo"] = df_hier["cnpj"].astype(str).str.replace(r"[^0-9]", "", regex=True)
        df_hier["cnpj_valido"] = df_hier["cnpj_limpo"].str.len() >= 8
        df_hier["cnpj_raiz"] = df_hier.apply(
            lambda r: r["cnpj_limpo"][:8] if r["cnpj_valido"] else None, axis=1
        )

        # Para cada grupo de CNPJ raiz valido, define a revenda principal
        def revenda_principal_do_grupo(grupo):
            nomes = grupo["revenda"].dropna().astype(str).str.strip().unique()
            # Prioriza nomes presentes nos valores do mapeamento
            for nome in nomes:
                if nome and str(nome).upper() in [v.upper() for v in MAPA_REVENDA_PRINCIPAL.values()]:
                    return nome
            # Fallback: nome mais curto
            nomes_validos = [n for n in nomes if n and str(n).lower() != "nan"]
            if nomes_validos:
                return min(nomes_validos, key=lambda x: len(str(x)))
            return grupo["revenda"].dropna().astype(str).str.strip().iloc[0] if not grupo["revenda"].dropna().empty else ""

        # Aplica agrupamento apenas para CNPJs validos
        df_validos = df_hier[df_hier["cnpj_valido"]].copy()
        if not df_validos.empty:
            raiz_para_revenda = df_validos.groupby("cnpj_raiz").apply(revenda_principal_do_grupo).to_dict()
            df_hier.loc[df_hier["cnpj_valido"], "revenda"] = df_hier.loc[df_hier["cnpj_valido"], "cnpj_raiz"].map(raiz_para_revenda)

        df_hier = df_hier.drop(columns=["cnpj_limpo", "cnpj_valido", "cnpj_raiz"])

    # Se um CPF aparecer em mais de uma revenda/loja, mantem a primeira ocorrencia
    df_hier = df_hier.drop_duplicates(subset=["cpf_limp"], keep="first")
    logger.info(f"Hierarquias consolidadas: {len(df_hier)} CPFs unicos")
    return df_hier


def mapeamento_revenda_regional(df_cad):
    """Cria mapeamento revenda -> regional a partir do cadastro base."""
    mapa = {}
    if "grupo" in df_cad.columns and "regional" in df_cad.columns:
        for _, row in df_cad.dropna(subset=["grupo", "regional"]).iterrows():
            rev = str(row["grupo"]).strip()
            reg = str(row["regional"]).strip()
            if rev and reg and rev.lower() != "nan" and reg.lower() != "nan":
                mapa[rev] = reg
    return mapa


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


def extrair_sku_do_curso(nome_curso):
    """Extrai o código/SKU do final do nome do curso, tratando '|' ou ' I ' como separadores."""
    s = str(nome_curso).strip()
    # Tenta separar por '|' ou ' I ' (usado em cursos antigos maiúsculos)
    for sep in ["|", " I "]:
        if sep in s:
            return s.split(sep)[-1].strip()
    return s


def detectar_cursos_obrigatorios(df_trein, ano_mes):
    """
    Detecta os 2 cursos obrigatórios do mês a partir da trilha de conteúdo obrigatório.
    Retorna: (curso1, curso2, nome_curto1, nome_curto2, sku1, sku2)
    """
    mes_dt = pd.Period(ano_mes, freq="M")
    obr = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
        & (df_trein["Trilha"].str.contains("OBRIGAT", case=False, na=False))
    ].copy()

    if obr.empty:
        return None, None, None, None, None, None

    top = obr["Curso"].value_counts().head(2)
    if len(top) < 2:
        logger.warning(f"Foram encontrados apenas {len(top)} curso(s) obrigatório(s) para {ano_mes}")

    cursos = top.index.tolist()
    nomes_curtos = []
    skus = []
    for curso in cursos:
        codigo = extrair_sku_do_curso(curso)
        skus.append(codigo)
        # Se houver mapeamento descritivo, usa; senão mantém o código
        nomes_curtos.append(NOMES_CURSOS.get(codigo, codigo))

    return (
        cursos[0] if len(cursos) > 0 else None,
        cursos[1] if len(cursos) > 1 else None,
        nomes_curtos[0] if len(nomes_curtos) > 0 else None,
        nomes_curtos[1] if len(nomes_curtos) > 1 else None,
        skus[0] if len(skus) > 0 else None,
        skus[1] if len(skus) > 1 else None,
    )


def fig_to_base64(fig):
    """Converte figura matplotlib em string base64 para embutir no e-mail."""
    try:
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=100, bbox_inches="tight", facecolor="white")
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")
        plt.close(fig)
        return img_base64
    except Exception as e:
        logger.warning(f"Falha ao converter gráfico para base64: {e}")
        plt.close(fig)
        return None


def identificar_destaque(df, coluna_metrica, coluna_nome, maior_melhor=True, min_base=10, coluna_base=None, meta=None):
    """Identifica a melhor revenda/regional, exigindo base mínima para evitar outliers.
    Se meta for informada, só destaca quem atingiu ou superou o objetivo."""
    if df is None or df.empty or coluna_metrica not in df.columns or coluna_nome not in df.columns:
        return None

    df_valido = df.copy()
    if coluna_base and coluna_base in df_valido.columns:
        df_valido = df_valido[df_valido[coluna_base] >= min_base]
    elif "total" in df_valido.columns:
        df_valido = df_valido[df_valido["total"] >= min_base]
    elif "total_ativos" in df_valido.columns:
        df_valido = df_valido[df_valido["total_ativos"] >= min_base]

    if df_valido.empty:
        df_valido = df.copy()

    # Quem atingiu a meta, se houver meta definida
    if meta is not None:
        df_valido = df_valido[df_valido[coluna_metrica] >= meta]
        if df_valido.empty:
            return None

    try:
        if maior_melhor:
            destaque = df_valido.sort_values(coluna_metrica, ascending=False).iloc[0]
        else:
            destaque = df_valido.sort_values(coluna_metrica, ascending=True).iloc[0]
    except Exception:
        return None

    return {
        "nome": str(destaque.get(coluna_nome, "N/A")),
        "valor": float(destaque.get(coluna_metrica, 0)),
        "regional": str(destaque.get("regional_curta", destaque.get("regional", "N/A"))),
        "base": int(destaque.get(coluna_base, destaque.get("total", destaque.get("total_ativos", 0)))),
    }


def farol_html(valor, meta, tamanho=16):
    """Retorna emoji de farol de acordo com o valor vs meta.
    Semáforo: verde (atingiu meta), amarelo (entre 70% e a meta),
    vermelho (abaixo de 70% da meta - critico).
    """
    if valor >= meta:
        return f'<span style="font-size:{tamanho}px;">🟢</span>'
    elif valor >= meta * 0.7:
        return f'<span style="font-size:{tamanho}px;">🟡</span>'
    else:
        return f'<span style="font-size:{tamanho}px;">🔴</span>'


def html_destaque(titulo, nome, regional, valor, sufixo="%", meta=None):
    """Gera conteúdo HTML de parabenização com farol e ordem regional -> revenda."""
    farol = farol_html(valor, meta, tamanho=18) if meta is not None else ""
    return (
        f"<div style='font-size:18px; font-weight:bold; margin-bottom:4px;'>"
        f"{farol} <strong>Parabéns.</strong></div>"
        f"<div style='font-size:15px; font-weight:bold; margin-bottom:10px;'>{titulo}</div>"
        f"A regional <strong>{regional}</strong> se destaca com a revenda <strong>{nome.upper()}</strong> "
        f"com <strong>{valor:.1f}{sufixo}</strong>."
    )


def formatar_inteiro(val):
    """Formata número como inteiro, tratando N/A e NaN."""
    if pd.isna(val) or val == "N/A":
        return "-"
    return f"{int(float(val)):,}".replace(",", ".")


def estilizar_tabela_html(df, destaque_coluna=None, destaque_menor_que_media=None,
                          formato_inteiro=True, semaforo_coluna=None, meta_semaforo=None):
    """
    Gera tabela HTML estilizada a partir de DataFrame.
    Suporta semaforo (verde/laranja/amarelo) para colunas percentuais e
    destaque de valores abaixo da media.
    """
    if df.empty:
        return "<p><em>Sem dados para exibir.</em></p>"

    html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;">\n'
    html += "<thead><tr>"
    for col in df.columns:
        html += (
            f'<th bgcolor="#ef4e22" style="border: 1px solid #cccccc; padding: 8px; background-color: #ef4e22; '
            f'color: white; text-align: center;">{col}</th>'
        )
    html += "</tr></thead><tbody>\n"

    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = row[col]
            is_num = isinstance(val, (int, float)) and not pd.isna(val)
            align = "right" if is_num else "left"

            # Semaforo para colunas percentuais (verde / amarelo / vermelho)
            bg = "background-color: #ffffff;"
            farol_celula = ""
            if semaforo_coluna and col == semaforo_coluna and is_num and meta_semaforo is not None:
                v = float(val)
                if v >= meta_semaforo:
                    bg = 'background-color: #d4edda; color: #155724;'  # verde
                    farol_celula = '🟢 '
                elif v >= meta_semaforo * 0.7:
                    bg = 'background-color: #fff3cd; color: #856404;'  # amarelo
                    farol_celula = '🟡 '
                else:
                    bg = 'background-color: #f8d7da; color: #721c24;'  # vermelho
                    farol_celula = '🔴 '

            # Destaca celulas abaixo da media, se solicitado (sobrepoe semaforo)
            if destaque_coluna and col == destaque_coluna and destaque_menor_que_media is not None and is_num:
                if float(val) < destaque_menor_que_media:
                    bg = 'background-color: #f8d7da; color: #721c24;'
                    farol_celula = '🔴 '

            # Formatacao
            if is_num and formato_inteiro:
                display = formatar_inteiro(val)
            elif is_num:
                display = f"{val:.1f}" if isinstance(val, float) else str(val)
            else:
                display = "" if pd.isna(val) else str(val)

            # Adiciona % nas colunas percentuais quando nao estiver presente
            if "%" in col and is_num and not str(display).endswith("%"):
                display = f"{display}%"

            # Prefixa farol quando houver
            display = f"{farol_celula}{display}" if farol_celula else display

            html += f'<td style="border: 1px solid #cccccc; padding: 6px; text-align: {align}; {bg}">{display}</td>'
        html += "</tr>\n"

    html += "</tbody></table>"
    return html


# ---------------------------------------------------------------------------
# CARGA E PROCESSAMENTO DE BASES
# ---------------------------------------------------------------------------
def carregar_bases():
    """Carrega cadastro, treinamentos, aceites, emails regionais e hierarquias."""
    logger.info("Carregando bases...")

    # ------------------------------------------------------------------
    # Cadastro base (status, nome, endereço e mapeamento regional)
    # ------------------------------------------------------------------
    cadastro_path = DATA_DIR / "cadastro.xlsx"
    xl_cad = pd.ExcelFile(cadastro_path)
    df_cad = pd.read_excel(cadastro_path, sheet_name=xl_cad.sheet_names[0])
    df_cad.columns = [c.strip().lower() for c in df_cad.columns]
    df_cad["cpf_limp"] = df_cad["cpf/cnpj"].apply(limpar_cpf)
    df_cad["revenda_original"] = df_cad["grupo"].astype(str).str.strip()
    df_cad["regional_original"] = df_cad["regional"].astype(str).str.strip()

    # Mapeamento revenda -> regional a partir do cadastro base
    mapa_regional = mapeamento_revenda_regional(df_cad)
    logger.info(f"Mapeamento revenda->regional: {len(mapa_regional)} revendas")

    # ------------------------------------------------------------------
    # Hierarquias (fonte correta de CPF -> revenda/loja/CNPJ/cargo)
    # ------------------------------------------------------------------
    df_hier = carregar_hierarquias()
    if df_hier is not None:
        cols_hier = ["cpf_limp", "revenda", "cod_loja", "cnpj", "nome_hier", "cargo_hier",
                     "vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado"]
        df_hier = df_hier[[c for c in cols_hier if c in df_hier.columns]].copy()
        df_cad = df_cad.merge(df_hier, on="cpf_limp", how="left")
        # Normaliza revenda da hierarquia para revenda principal
        df_hier["revenda"] = df_hier["revenda"].apply(normalizar_revenda_hierarquia)
        df_cad["revenda_hier_norm"] = df_cad["revenda"].apply(normalizar_revenda_hierarquia)
        # Se encontrou revenda na hierarquia, usa a normalizada; senao mantem a original
        df_cad["revenda"] = df_cad["revenda_hier_norm"].fillna(df_cad["revenda_original"].apply(normalizar_revenda))
        # Aplica nome amigavel de exibicao
        df_cad["revenda"] = df_cad["revenda"].apply(nome_revenda_exibicao)
        # Regional pela nova revenda, via mapeamento; fallback para original
        df_cad["regional"] = df_cad["revenda"].map(mapa_regional).fillna(df_cad["regional_original"])
        # Usa nome da hierarquia quando disponivel
        if "nome_hier" in df_cad.columns:
            df_cad["nome"] = df_cad["nome_hier"].fillna(df_cad["nome"])
        # Cargo da hierarquia (quando houver)
        if "cargo_hier" in df_cad.columns:
            df_cad["cargo"] = df_cad["cargo_hier"].fillna(df_cad["cargo"])
        logger.info(f"CPFs do cadastro com hierarquia encontrada: {df_cad['revenda'].notna().sum()}")
    else:
        df_cad["revenda"] = df_cad["revenda_original"].apply(nome_revenda_exibicao)
        df_cad["regional"] = df_cad["regional_original"]
        logger.warning("Hierarquias não carregadas. Usando cadastro base para revenda/regional.")

    df_cad["regional_curta"] = df_cad["regional"].apply(regional_curta)
    df_cad = df_cad[~df_cad["revenda"].isin(REVENDAS_EXCLUIR)].copy()

    # ------------------------------------------------------------------
    # Regra de Pré-Cadastrado 90+ dias na base do banco
    # ------------------------------------------------------------------
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
    df_cad = df_cad.dropna(subset=["regional_curta", "revenda"]).copy()
    df_cad = df_cad[df_cad["regional_curta"].str.lower() != "nan"].copy()
    df_cad = df_cad[df_cad["revenda"].str.lower() != "nan"].copy()
    logger.info(f"Cadastro: {len(df_cad):,} registros, {df_cad['cpf_limp'].nunique():,} CPFs únicos")

    # ------------------------------------------------------------------
    # Treinamentos
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # Aceites
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # Detalhamento (consolidado das hierarquias)
    # ------------------------------------------------------------------
    df_det = None
    if df_hier is not None:
        df_det = df_hier.copy()
        df_det = df_det[df_det["cpf_limp"].isin(df_cad["cpf_limp"].unique())].copy()
        # Cruza com cadastro para nome, cidade, uf, bairro, status e regional
        df_det = df_det.merge(
            df_cad[["cpf_limp", "nome", "cargo", "status", "cidade", "uf", "bairro", "regional_curta"]].drop_duplicates("cpf_limp"),
            on="cpf_limp",
            how="left",
        )
        # Padroniza colunas esperadas pela aba de detalhamento
        df_det = df_det.rename(columns={
            "regional_curta": "regional_da_loja",
            "revenda": "loja",
            "cnpj": "cnpj_loja",
        })
        # Mantém nome da hierarquia quando não houver no cadastro
        if "nome_hier" in df_det.columns and "nome" in df_det.columns:
            df_det["nome"] = df_det["nome"].fillna(df_det["nome_hier"])
        logger.info(f"Detalhamento: {len(df_det):,} registros")
    else:
        logger.warning("Hierarquias não carregadas. Aba de detalhamento não será gerada.")

    # ------------------------------------------------------------------
    # Emails regionais
    # ------------------------------------------------------------------
    emails_reg_path = DATA_DIR / "emails_regionais.xlsx"
    df_emails_reg = None
    if emails_reg_path.exists():
        df_emails_reg = pd.read_excel(emails_reg_path)
        df_emails_reg.columns = [c.strip().lower() for c in df_emails_reg.columns]
        if "regional" in df_emails_reg.columns and "email" in df_emails_reg.columns:
            df_emails_reg["regional"] = df_emails_reg["regional"].apply(regional_curta)
            logger.info(f"Emails regionais: {len(df_emails_reg)} registros")
        else:
            logger.warning("emails_regionais.xlsx deve ter colunas 'Regional' e 'Email'")
            df_emails_reg = None
    else:
        logger.warning("emails_regionais.xlsx não encontrado. Envio por regional será desabilitado.")

    return {
        "cadastro": df_cad,
        "treinamentos": df_trein,
        "aceites": df_aceite,
        "detalhamento": df_det,
        "aba_aceite": aba_aceite,
        "usou_ultima_aba": usou_ultima,
        "mes_referencia": ano_mes_hoje,
        "emails_regionais": df_emails_reg,
    }


# ---------------------------------------------------------------------------
# CÁLCULOS POR INDICADOR
# ---------------------------------------------------------------------------
def calcular_cadastros(df_cad):
    """Calcula cadastros por regional e por revenda."""
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
    cad_reg["regional_curta"] = cad_reg["regional_curta"].apply(regional_title_case)

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
    cad_rev = cad_rev.sort_values("pct_ativos", ascending=False)
    cad_rev["regional_curta"] = cad_rev["regional_curta"].apply(regional_title_case)

    return cad_reg, cad_rev


def calcular_treinamentos(df_trein, df_cad, ano_mes):
    """Calcula treinamentos por regional e por revenda (base = ativos).
    Retorna: (trein_reg, trein_rev, trein_por_curso, cursos_info)
    """
    curso1, curso2, nome1, nome2, sku1, sku2 = detectar_cursos_obrigatorios(df_trein, ano_mes)

    if curso1 is None or curso2 is None:
        logger.warning("Não foi possível detectar os 2 cursos obrigatórios do mês")
        return None, None, None, (curso1, curso2, nome1, nome2, sku1, sku2)

    mes_dt = pd.Period(ano_mes, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ]

    cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique())
    cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique())
    cpf_ambos = cpf_curso1 & cpf_curso2

    logger.info(f"Treinamentos {ano_mes}: '{nome1}' ({sku1})={len(cpf_curso1)}, '{nome2}' ({sku2})={len(cpf_curso2)}, ambos={len(cpf_ambos)}")

    ativos = df_cad[df_cad["status"] == "Ativo"].copy()

    def calcular_grupo(grupo_df, cpf_realizaram):
        cpfs = set(grupo_df["cpf_limp"].unique())
        realizaram = cpfs & cpf_realizaram
        nao_realizaram = cpfs - cpf_realizaram
        total = len(cpfs)
        pct = round(len(realizaram) / total * 100, 1) if total else 0
        return pd.Series({
            "total_ativos": int(total),
            "realizaram": int(len(realizaram)),
            "nao_realizaram": int(len(nao_realizaram)),
            "pct_realizaram": pct,
        })

    # Combinado (ambos os cursos)
    trein_reg = ativos.groupby("regional_curta").apply(lambda g: calcular_grupo(g, cpf_ambos)).reset_index()
    trein_reg = trein_reg.sort_values("pct_realizaram", ascending=False)
    trein_reg["regional_curta"] = trein_reg["regional_curta"].apply(regional_title_case)

    trein_rev = ativos.groupby(["regional_curta", "revenda"]).apply(lambda g: calcular_grupo(g, cpf_ambos)).reset_index()
    trein_rev = trein_rev.sort_values("pct_realizaram", ascending=False)
    trein_rev["regional_curta"] = trein_rev["regional_curta"].apply(regional_title_case)

    # Por curso individual
    def calcular_por_curso(cpf_curso):
        reg = ativos.groupby("regional_curta").apply(lambda g: calcular_grupo(g, cpf_curso)).reset_index()
        reg = reg.sort_values("pct_realizaram", ascending=False)
        reg["regional_curta"] = reg["regional_curta"].apply(regional_title_case)
        rev = ativos.groupby(["regional_curta", "revenda"]).apply(lambda g: calcular_grupo(g, cpf_curso)).reset_index()
        rev = rev.sort_values("pct_realizaram", ascending=False)
        rev["regional_curta"] = rev["regional_curta"].apply(regional_title_case)
        return reg, rev

    c1_reg, c1_rev = calcular_por_curso(cpf_curso1)
    c2_reg, c2_rev = calcular_por_curso(cpf_curso2)

    trein_por_curso = {
        "nome1": nome1,
        "nome2": nome2,
        "sku1": sku1,
        "sku2": sku2,
        "curso1_reg": c1_reg,
        "curso1_rev": c1_rev,
        "curso2_reg": c2_reg,
        "curso2_rev": c2_rev,
    }

    return trein_reg, trein_rev, trein_por_curso, (curso1, curso2, nome1, nome2, sku1, sku2)


def inferir_mes_da_aba(aba):
    """Tenta extrair ano-mês do nome da aba (ex: ABR_2026, JUN_2026)."""
    import re
    meses_map = {
        "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
        "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
    }
    aba_limpa = aba.strip().upper().replace("-", "_")
    match = re.search(r"([A-Z]{3})_(\d{4})", aba_limpa)
    if match:
        mes_nome, ano = match.groups()
        if mes_nome in meses_map:
            return pd.Period(f"{ano}-{meses_map[mes_nome]:02d}", freq="M")
    match = re.search(r"(\d{4})_(\d{2})", aba_limpa)
    if match:
        return pd.Period(f"{match.group(1)}-{match.group(2)}", freq="M")
    return None


def calcular_aceites(df_aceite, df_cad, ano_mes, aba_aceite, usou_ultima_aba=False):
    """Calcula aceites mensais por regional e por revenda (base = ativos)."""
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

    def calcular_grupo(grupo_df):
        cpfs = set(grupo_df["cpf_limp"].unique())
        aceitaram = cpfs & cpfs_aceitaram
        nao_aceitaram = cpfs - cpfs_aceitaram
        total = len(cpfs)
        pct = round(len(aceitaram) / total * 100, 1) if total else 0
        return pd.Series({
            "total_ativos": int(total),
            "aceitaram": int(len(aceitaram)),
            "nao_aceitaram": int(len(nao_aceitaram)),
            "pct_aceite": pct,
        })

    aceite_reg = (
        ativos.groupby("regional_curta")
        .apply(calcular_grupo)
        .reset_index()
        .rename(columns={"regional_curta": "regional"})
        .sort_values("pct_aceite", ascending=False)
    )
    aceite_reg["regional"] = aceite_reg["regional"].apply(regional_title_case)

    aceite_rev = (
        ativos.groupby(["regional_curta", "revenda"])
        .apply(calcular_grupo)
        .reset_index()
        .rename(columns={"regional_curta": "regional"})
        .sort_values("pct_aceite", ascending=False)
    )
    aceite_rev["regional"] = aceite_rev["regional"].apply(regional_title_case)

    return aceite_reg, aceite_rev, mes_ref


# ---------------------------------------------------------------------------
# SNAPSHOT SEMANAL
# ---------------------------------------------------------------------------
def salvar_snapshot(dados):
    """Salva snapshot dos indicadores atuais para comparação futura."""
    cad_reg, _ = dados["cadastros"]
    trein_reg, _ = dados["treinamentos"]
    aceite_reg = dados["aceites"]

    snapshot = {
        "data": date.today().isoformat(),
        "mes_referencia": dados["mes_referencia"],
        "cadastros": cad_reg.set_index("regional_curta")[["ativos", "total", "pct_ativos"]].to_dict("index"),
        "treinamentos": trein_reg.set_index("regional_curta")[["realizaram", "total_ativos", "pct_realizaram"]].to_dict("index") if trein_reg is not None else {},
        "aceites": aceite_reg.set_index("regional")[["aceitaram", "total_ativos", "pct_aceite"]].to_dict("index"),
    }

    snapshot_path = SNAPSHOT_DIR / f"snapshot_{date.today():%Y%m%d}.json"
    with open(snapshot_path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)
    logger.info(f"Snapshot salvo em: {snapshot_path}")


def carregar_snapshot_anterior():
    """Carrega o snapshot mais recente anterior ao dia de hoje."""
    snapshots = sorted(SNAPSHOT_DIR.glob("snapshot_*.json"))
    hoje_str = f"snapshot_{date.today():%Y%m%d}.json"
    anteriores = [s for s in snapshots if s.name != hoje_str]

    if not anteriores:
        return None

    snapshot_path = anteriores[-1]
    with open(snapshot_path, "r", encoding="utf-8") as f:
        return json.load(f)


def calcular_evolucao(snapshot_atual, snapshot_anterior):
    """Calcula evolução dos indicadores por regional em relação ao snapshot anterior."""
    if snapshot_anterior is None:
        return None

    evolucao = []
    for regional in snapshot_atual["cadastros"]:
        atual = snapshot_atual["cadastros"][regional]
        ant = snapshot_anterior["cadastros"].get(regional, {})
        evolucao.append({
            "regional": regional_title_case(regional),
            "ativos_atual": atual.get("ativos", 0),
            "ativos_anterior": ant.get("ativos", 0),
            "var_ativos": atual.get("ativos", 0) - ant.get("ativos", 0),
            "pct_atual": atual.get("pct_ativos", 0),
            "pct_anterior": ant.get("pct_ativos", 0),
            "var_pct": round(atual.get("pct_ativos", 0) - ant.get("pct_ativos", 0), 1),
        })

    return pd.DataFrame(evolucao).sort_values("var_pct", ascending=True)


# ---------------------------------------------------------------------------
# GRÁFICOS
# ---------------------------------------------------------------------------
def gerar_grafico_barras(df, x_col, y_col, titulo, cor="#ef4e22", meta=None):
    """
    Gera gráfico de barras horizontal ordenado do maior para o menor.
    As barras usam as cores do semáforo quando meta é informada:
      - verde: atingiu ou superou a meta
      - amarelo: entre 70% da meta e a meta
      - vermelho: abaixo de 70% da meta
    """
    if df.empty or y_col not in df.columns or x_col not in df.columns:
        logger.warning(f"Dados insuficientes para gerar gráfico: {titulo}")
        return None

    # Ordena do maior para o menor (igual à tabela)
    df = df.sort_values(y_col, ascending=False).reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(9, max(4, len(df) * 0.55)))
    fig.subplots_adjust(left=0.28)
    barras = ax.barh(df[x_col], df[y_col], color=cor)

    for i, (val, bar) in enumerate(zip(df[y_col], barras)):
        try:
            v = float(val)
            if meta is not None:
                if v >= meta:
                    bar.set_color("#2e7d32")  # verde
                elif v >= meta * 0.7:
                    bar.set_color("#f9a825")  # amarelo
                else:
                    bar.set_color("#d9534f")  # vermelho
        except (TypeError, ValueError):
            pass

    ax.set_xlabel("%")
    ax.set_title(titulo, fontsize=14, fontweight="bold", loc="left", color="#333333")
    ax.invert_yaxis()
    ax.set_xlim(0, 100)
    ax.grid(axis="x", linestyle="--", alpha=0.5)

    for i, v in enumerate(df[y_col]):
        try:
            ax.text(min(float(v) + 1, 98), i, f"{v}%", va="center", fontsize=10, color="#333333")
        except (TypeError, ValueError):
            pass

    plt.tight_layout()
    return fig


def gerar_graficos(cad_reg, trein_reg, aceite_reg):
    """Gera os 3 gráficos principais e retorna dict base64."""
    graficos = {}

    if not cad_reg.empty:
        media_cad = cad_reg["pct_ativos"].mean()
        fig = gerar_grafico_barras(
            cad_reg,
            "regional_curta", "pct_ativos",
            f"Cadastros Ativos por Regional (média: {media_cad:.1f}%)",
            meta=META_CADASTRO,
        )
        if fig:
            graficos["cadastros"] = fig_to_base64(fig)

    if trein_reg is not None and not trein_reg.empty:
        media_trein = trein_reg["pct_realizaram"].mean()
        fig = gerar_grafico_barras(
            trein_reg,
            "regional_curta", "pct_realizaram",
            f"Treinamentos Obrigatórios Realizados por Regional (média: {media_trein:.1f}%)",
            meta=META_TREINAMENTOS,
        )
        if fig:
            graficos["treinamentos"] = fig_to_base64(fig)

    if not aceite_reg.empty:
        media_aceite = aceite_reg["pct_aceite"].mean()
        fig = gerar_grafico_barras(
            aceite_reg,
            "regional", "pct_aceite",
            f"Aceite Mensal por Regional (média: {media_aceite:.1f}%)",
            meta=META_ACEITES,
        )
        if fig:
            graficos["aceites"] = fig_to_base64(fig)

    return graficos

    return graficos


# ---------------------------------------------------------------------------
# INSIGHTS AUTOMÁTICOS
# ---------------------------------------------------------------------------
def garantir_n_itens(df, coluna_metrica, mascara_filtro=None, limite=10, maior_melhor=True):
    """
    Garante que o DataFrame resultante tenha pelo menos `limite` itens.
    Prioriza itens que atendem ao `mascara_filtro` (ex: abaixo da média) e,
    se necessário, completa com os itens restantes mais extremos.
    """
    if df is None or df.empty:
        return df
    df = df.copy()
    if mascara_filtro is not None:
        selecionados = df[mascara_filtro].sort_values(coluna_metrica, ascending=not maior_melhor)
    else:
        selecionados = df.sort_values(coluna_metrica, ascending=not maior_melhor)

    if len(selecionados) >= limite:
        return selecionados.head(limite)

    ids_selecionados = set(selecionados.index)
    restantes = df[~df.index.isin(ids_selecionados)].sort_values(
        coluna_metrica, ascending=not maior_melhor
    )
    return pd.concat([selecionados, restantes]).head(limite)


def gerar_insights(cad_reg, cad_rev, trein_reg, trein_rev, aceite_reg, aceite_rev, evolucao, dados):
    """Gera textos de insights organizados por seção, com farol e linguagem ajustada."""
    mes_nome = nome_mes_pt_br().lower()
    mes_aceite_nome = nome_mes_pt_br(ano_mes=str(dados["mes_aceite"]))

    resultado = {
        "cadastros": {"intro": "", "alerta": "", "alerta_titulo": "", "alerta_subtitulo": "", "alerta_itens": None, "destaque": ""},
        "evolucao": "",
        "treinamentos": {"intro": "", "alerta": "", "alerta_titulo": "", "alerta_subtitulo": "", "alerta_itens": None, "destaque": ""},
        "aceites": {"intro": "", "alerta": "", "alerta_titulo": "", "alerta_subtitulo": "", "alerta_itens": None, "destaque": ""},
    }

    # --- Cadastros ---
    if not cad_reg.empty:
        total_ativos = int(cad_reg["ativos"].sum())
        total_base = int(cad_reg["total"].sum())
        pct_geral = round(total_ativos / total_base * 100, 1)
        media_reg = cad_reg["pct_ativos"].mean()
        farol_geral = farol_html(pct_geral, META_CADASTRO)

        resultado["cadastros"]["intro"] = (
            f"{farol_geral} Nosso objetivo ideal para a base de cadastros do Programa +TOP é de "
            f"<strong>{META_CADASTRO:.0f}%</strong>.<br>"
            f"Até o momento, estamos em <strong>{pct_geral}%</strong> da base ativa."
        )

        abaixo_media = garantir_n_itens(
            cad_rev, "pct_ativos",
            mascara_filtro=cad_rev["pct_ativos"] < media_reg,
            limite=10, maior_melhor=True
        )
        if not abaixo_media.empty:
            itens_df = abaixo_media[["regional_curta", "revenda", "pct_ativos"]].copy()
            itens_df["revenda"] = itens_df["revenda"].str.title()
            itens_df = itens_df.rename(columns={
                "regional_curta": "Regional",
                "revenda": "Revenda",
                "pct_ativos": "% Ativos",
            })
            resultado["cadastros"]["alerta_titulo"] = "Ponto de atenção: revendas com % de base menor do que a média do programa"
            resultado["cadastros"]["alerta_subtitulo"] = ""
            resultado["cadastros"]["alerta_itens"] = itens_df

        # Apenas 1 destaque geral (a melhor revenda do programa, que atingiu a meta)
        dest = identificar_destaque(cad_rev, "pct_ativos", "revenda", maior_melhor=True, min_base=20, meta=META_CADASTRO)
        if dest:
            resultado["cadastros"]["destaque"] = html_destaque(
                "Melhor % de cadastros ativos.", dest["nome"], dest["regional"], dest["valor"], meta=META_CADASTRO
            )

    # --- Evolucao semanal ---
    if evolucao is not None and not evolucao.empty:
        partes = []
        piores = evolucao[evolucao["var_pct"] < 0].sort_values("var_pct").head(5)
        melhores = evolucao[evolucao["var_pct"] > 0].sort_values("var_pct", ascending=False).head(5)

        if not piores.empty:
            linhas = [f"{r['regional']}: {r['var_pct']:+.1f}p.p. ({r['ativos_anterior']} -> {r['ativos_atual']} ativos)" for _, r in piores.iterrows()]
            partes.append("Regionais que <strong>perderam</strong> percentual de ativos.<br>" + "<br>".join(linhas))

        if not melhores.empty:
            linhas = [f"{r['regional']}: {r['var_pct']:+.1f}p.p. ({r['ativos_anterior']} -> {r['ativos_atual']} ativos)" for _, r in melhores.iterrows()]
            partes.append("Regionais que <strong>cresceram</strong> no percentual de ativos.<br>" + "<br>".join(linhas))

        if partes:
            resultado["evolucao"] = "".join(
                [f"<p style='margin:0 0 10px 0; line-height:1.5;'>{p}</p>" for p in partes]
            )

    # --- Treinamentos ---
    if trein_reg is not None and not trein_reg.empty:
        pct_trein_geral = round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1)
        _, _, nome1, nome2, sku1, sku2 = dados["cursos_info"]
        farol_trein = farol_html(pct_trein_geral, META_TREINAMENTOS)
        resultado["treinamentos"]["intro"] = (
            f"{farol_trein} Os dois conteúdos de {mes_nome} estão disponíveis para os vendedores ate o dia "
            f"{date.today().replace(day=30):%d/%m/%Y}.<br>"
            f"Nosso objetivo ideal para os treinamentos do Programa +TOP é de <strong>{META_TREINAMENTOS:.0f}%</strong>.<br>"
            f"Até o momento, estamos em <strong>{pct_trein_geral}%</strong> da base ativa.<br>"
            f"Conteúdo 1: {nome1} (SKU {sku1})<br>Conteúdo 2: {nome2} (SKU {sku2})"
        )

        media_trein = trein_reg["pct_realizaram"].mean()
        piores_rev = garantir_n_itens(
            trein_rev, "pct_realizaram",
            mascara_filtro=trein_rev["pct_realizaram"] < media_trein,
            limite=10, maior_melhor=True
        )
        if not piores_rev.empty:
            itens_df = piores_rev[["regional_curta", "revenda", "pct_realizaram"]].copy()
            itens_df["revenda"] = itens_df["revenda"].str.title()
            itens_df = itens_df.rename(columns={
                "regional_curta": "Regional",
                "revenda": "Revenda",
                "pct_realizaram": "% Realizado",
            })
            resultado["treinamentos"]["alerta_titulo"] = "Alerta de treinamentos: revendas abaixo da média precisam de reforço"
            resultado["treinamentos"]["alerta_subtitulo"] = f"A média geral do programa é de {media_trein:.1f}%."
            resultado["treinamentos"]["alerta_itens"] = itens_df

        # Apenas 1 destaque geral
        dest = identificar_destaque(trein_rev, "pct_realizaram", "revenda", maior_melhor=True, min_base=20, meta=META_TREINAMENTOS)
        if dest:
            resultado["treinamentos"]["destaque"] = html_destaque(
                "Melhor % de treinamentos concluídos.", dest["nome"], dest["regional"], dest["valor"], meta=META_TREINAMENTOS
            )

    # --- Aceites ---
    if not aceite_reg.empty:
        pct_aceite_geral = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1)
        farol_aceite = farol_html(pct_aceite_geral, META_ACEITES)
        resultado["aceites"]["intro"] = (
            f"{farol_aceite} Sobre os aceites mensais de <strong>{mes_aceite_nome}</strong>.<br>"
            f"Nosso objetivo ideal é de <strong>{META_ACEITES:.0f}%</strong>. "
            f"Ate o momento, <strong>{pct_aceite_geral}%</strong> dos participantes ativos deram aceite na campanha base."
        )

        media_aceite = aceite_reg["pct_aceite"].mean()
        if aceite_rev is not None and not aceite_rev.empty:
            abaixo_aceite = garantir_n_itens(
                aceite_rev, "pct_aceite",
                mascara_filtro=aceite_rev["pct_aceite"] < media_aceite,
                limite=10, maior_melhor=True
            )
            if not abaixo_aceite.empty:
                itens_df = abaixo_aceite[["regional", "revenda", "pct_aceite"]].copy()
                itens_df = itens_df.rename(columns={
                    "regional": "Regional",
                    "revenda": "Revenda",
                    "pct_aceite": "% Aceite",
                })
                resultado["aceites"]["alerta_titulo"] = "Revendas com aceite abaixo da média e que precisam de atenção"
                resultado["aceites"]["alerta_subtitulo"] = f"A média geral do programa é de {media_aceite:.1f}%."
                resultado["aceites"]["alerta_itens"] = itens_df

        # Apenas 1 destaque geral
        if aceite_rev is not None and not aceite_rev.empty:
            dest = identificar_destaque(aceite_rev, "pct_aceite", "revenda", maior_melhor=True, min_base=20, meta=META_ACEITES)
            if dest:
                resultado["aceites"]["destaque"] = html_destaque(
                    "Melhor % de aceite mensal.", dest["nome"], dest["regional"], dest["valor"], meta=META_ACEITES
                )

    return resultado


def gerar_insights_regional(cad_reg, cad_rev, trein_reg, trein_rev, aceite_reg, aceite_rev, evolucao, dados, regional_filtro):
    """Gera insights específicos para uma regional, organizados por seção."""
    mes_nome = nome_mes_pt_br().lower()
    mes_aceite_nome = nome_mes_pt_br(ano_mes=str(dados["mes_aceite"]))

    resultado = {
        "cadastros": {"intro": "", "alerta": "", "destaque": ""},
        "evolucao": "",
        "treinamentos": {"intro": "", "alerta": "", "destaque": ""},
        "aceites": {"intro": "", "alerta": "", "destaque": ""},
    }

    cad_reg_f = cad_reg[cad_reg["regional_curta"] == regional_filtro]
    cad_rev_f = cad_rev[cad_rev["regional_curta"] == regional_filtro]
    trein_reg_f = trein_reg[trein_reg["regional_curta"] == regional_filtro] if trein_reg is not None else None
    trein_rev_f = trein_rev[trein_rev["regional_curta"] == regional_filtro] if trein_rev is not None else None
    aceite_reg_f = aceite_reg[aceite_reg["regional"] == regional_filtro]
    aceite_rev_f = aceite_rev[aceite_rev["regional"] == regional_filtro] if aceite_rev is not None else None

    if not cad_reg_f.empty:
        row = cad_reg_f.iloc[0]
        media = cad_reg["pct_ativos"].mean()
        posição = (cad_reg["pct_ativos"] > row["pct_ativos"]).sum() + 1
        total = len(cad_reg)
        farol_reg = farol_html(row['pct_ativos'], META_CADASTRO)
        resultado["cadastros"]["intro"] = (
            f"{farol_reg} A regional <strong>{regional_filtro}</strong> está com <strong>{row['pct_ativos']:.1f}%</strong> de base ativa.<br>"
            f"Ela ocupa a <strong>{posição}ª posição</strong> entre {total} regionais (media geral: {media:.1f}%).<br>"
            f"Objetivo ideal: <strong>{META_CADASTRO:.0f}%</strong>."
        )

        abaixo_media = cad_rev_f[cad_rev_f["pct_ativos"] < media].sort_values("pct_ativos", ascending=False).head(10)
        if not abaixo_media.empty:
            linhas = [f"{r['revenda'].upper()} {r['pct_ativos']:.0f}%" for _, r in abaixo_media.iterrows()]
            resultado["cadastros"]["alerta"] = (
                "Revendas da regional abaixo da média geral do programa. Precisam de reforço.<br>" + "<br>".join(linhas)
            )

        dest = identificar_destaque(cad_rev_f, "pct_ativos", "revenda", maior_melhor=True, min_base=10, meta=META_CADASTRO)
        if dest:
            resultado["cadastros"]["destaque"] = html_destaque(
                "Melhor % de cadastros ativos.", dest["nome"], dest["regional"], dest["valor"], meta=META_CADASTRO
            )

    if evolucao is not None and not evolucao.empty:
        evo_f = evolucao[evolucao["regional"] == regional_filtro]
        if not evo_f.empty:
            r = evo_f.iloc[0]
            if r["var_pct"] > 0:
                resultado["evolucao"] = (
                    f"Na comparacao com a semana anterior, a regional <strong>cresceu {r['var_pct']:+.1f}p.p.</strong><br>"
                    f"({r['ativos_anterior']} -> {r['ativos_atual']} ativos)."
                )
            elif r["var_pct"] < 0:
                resultado["evolucao"] = (
                    f"Na comparacao com a semana anterior, a regional <strong>perdeu {r['var_pct']:+.1f}p.p.</strong><br>"
                    f"({r['ativos_anterior']} -> {r['ativos_atual']} ativos). Atenção para reverter essa queda."
                )
            else:
                resultado["evolucao"] = (
                    f"Na comparacao com a semana anterior, a regional manteve o percentual de ativos em {r['pct_atual']:.1f}%."
                )

    if trein_reg_f is not None and not trein_reg_f.empty:
        r = trein_reg_f.iloc[0]
        media_trein = trein_reg["pct_realizaram"].mean()
        _, _, nome1, nome2, sku1, sku2 = dados["cursos_info"]
        farol_trein_reg = farol_html(r['pct_realizaram'], META_TREINAMENTOS)
        resultado["treinamentos"]["intro"] = (
            f"{farol_trein_reg} Treinamentos de {mes_nome}: <strong>{r['pct_realizaram']:.1f}%</strong> da base ativa da regional concluiu ambos os cursos.<br>"
            f"Média geral do programa: {media_trein:.1f}%. Objetivo ideal: <strong>{META_TREINAMENTOS:.0f}%</strong>.<br>"
            f"Conteúdo 1: {nome1} (SKU {sku1})<br>Conteúdo 2: {nome2} (SKU {sku2})"
        )

        piores_rev = trein_rev_f[trein_rev_f["pct_realizaram"] < media_trein].sort_values("pct_realizaram", ascending=False).head(10)
        if not piores_rev.empty:
            linhas = [f"{r['revenda'].upper()} {r['pct_realizaram']:.0f}%" for _, r in piores_rev.iterrows()]
            resultado["treinamentos"]["alerta"] = (
                "Revendas da regional com treinamentos abaixo da media.<br>" + "<br>".join(linhas)
            )

        dest = identificar_destaque(trein_rev_f, "pct_realizaram", "revenda", maior_melhor=True, min_base=10, meta=META_TREINAMENTOS)
        if dest:
            resultado["treinamentos"]["destaque"] = html_destaque(
                "Melhor % de treinamentos concluídos.", dest["nome"], dest["regional"], dest["valor"], meta=META_TREINAMENTOS
            )

    if not aceite_reg_f.empty:
        r = aceite_reg_f.iloc[0]
        media_aceite = aceite_reg["pct_aceite"].mean()
        farol_aceite_reg = farol_html(r['pct_aceite'], META_ACEITES)
        resultado["aceites"]["intro"] = (
            f"{farol_aceite_reg} Aceite mensal de {mes_aceite_nome}: <strong>{r['pct_aceite']:.1f}%</strong> dos ativos da regional deram aceite.<br>"
            f"Média geral do programa: {media_aceite:.1f}%. Objetivo ideal: <strong>{META_ACEITES:.0f}%</strong>."
        )

        if aceite_rev_f is not None and not aceite_rev_f.empty:
            dest = identificar_destaque(aceite_rev_f, "pct_aceite", "revenda", maior_melhor=True, min_base=10, meta=META_ACEITES)
            if dest:
                resultado["aceites"]["destaque"] = html_destaque(
                    "Melhor % de aceite mensal.", dest["nome"], dest["regional"], dest["valor"], meta=META_ACEITES
                )

    return resultado


# ---------------------------------------------------------------------------
# MONTAGEM DO E-MAIL
# ---------------------------------------------------------------------------
def _renomear_cadastro_reg(df):
    return df.rename(columns={
        "regional_curta": "Regional",
        "total": "Total de Participantes",
        "ativos": "Usuários com +TOP",
        "nao_ativos": "Usuários sem +TOP",
        "pct_ativos": "% Ativos",
    })


def _renomear_cadastro_rev(df):
    return df.rename(columns={
        "regional_curta": "Regional",
        "revenda": "Revenda",
        "total": "Total",
        "ativos": "Usuários com +TOP",
        "nao_ativos": "Usuários sem +TOP",
        "pct_ativos": "% Ativos",
    })


def _renomear_trein_reg(df):
    return df.rename(columns={
        "regional_curta": "Regional",
        "total_ativos": "Base Ativa",
        "realizaram": "Realizado",
        "nao_realizaram": "Não Realizado",
        "pct_realizaram": "% Realizado",
    })


def _renomear_trein_rev(df):
    return df.rename(columns={
        "regional_curta": "Regional",
        "revenda": "Revenda",
        "total_ativos": "Base Ativa",
        "realizaram": "Realizado",
        "nao_realizaram": "Não Realizado",
        "pct_realizaram": "% Realizado",
    })


def _renomear_aceite_reg(df):
    return df.rename(columns={
        "regional": "Regional",
        "total_ativos": "Base Ativa",
        "aceitaram": "Aceitaram",
        "nao_aceitaram": "Não Aceitaram",
        "pct_aceite": "% Aceite",
    })


def _renomear_aceite_rev(df):
    return df.rename(columns={
        "regional": "Regional",
        "revenda": "Revenda",
        "total_ativos": "Base Ativa",
        "aceitaram": "Aceitaram",
        "nao_aceitaram": "Não Aceitaram",
        "pct_aceite": "% Aceite",
    })


def preparar_tabela_treinamentos_combinada(trein_por_curso, trein_ambos=None, nivel="regional"):
    """
    Gera uma única tabela com colunas para cada curso obrigatório.
    nivel='regional' agrupa por regional; nivel='revenda' inclui revenda.
    trein_ambos: DataFrame combinado (trein_reg ou trein_rev) com % de ambos os cursos.
    """
    if trein_por_curso is None:
        return None

    c1_key = "curso1_reg" if nivel == "regional" else "curso1_rev"
    c2_key = "curso2_reg" if nivel == "regional" else "curso2_rev"

    c1 = trein_por_curso.get(c1_key)
    c2 = trein_por_curso.get(c2_key)
    nome1 = trein_por_curso.get("nome1", "Curso 1")
    nome2 = trein_por_curso.get("nome2", "Curso 2")

    if c1 is None or c2 is None or c1.empty or c2.empty:
        return None

    chave = ["regional_curta"]
    if nivel == "revenda":
        chave.append("revenda")

    df1 = c1[chave + ["pct_realizaram"]].rename(columns={"pct_realizaram": f"% {nome1}"})
    df2 = c2[chave + ["pct_realizaram"]].rename(columns={"pct_realizaram": f"% {nome2}"})

    df = df1.merge(df2, on=chave, how="outer")

    # Adiciona % de ambos os cursos, se disponível
    if trein_ambos is not None and not trein_ambos.empty:
        cols = chave + ["pct_realizaram"]
        df_ambos = trein_ambos[cols].rename(columns={"pct_realizaram": "% Ambos"})
        df = df.merge(df_ambos, on=chave, how="left")

    df = df.sort_values(f"% {nome1}", ascending=False).fillna(0)

    # Renomeia chaves para exibição
    rename = {"regional_curta": "Regional"}
    if nivel == "revenda":
        rename["revenda"] = "Revenda"
    df = df.rename(columns=rename)

    return df


def preparar_tabelas(dados):
    """Prepara tabelas HTML renomeadas e formatadas."""
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    aceite_reg = dados["aceites"]
    aceite_rev = dados.get("aceites_rev")
    trein_por_curso = dados.get("treinamentos_por_curso")

    tabelas = {
        "cad_reg": estilizar_tabela_html(
            _renomear_cadastro_reg(cad_reg),
            semaforo_coluna="% Ativos",
            meta_semaforo=META_CADASTRO,
        ),
        "cad_rev": estilizar_tabela_html(
            _renomear_cadastro_rev(cad_rev.head(10)),
            destaque_coluna="% Ativos",
            destaque_menor_que_media=cad_reg["pct_ativos"].mean(),
            semaforo_coluna="% Ativos",
            meta_semaforo=META_CADASTRO,
        ),
        "trein_reg": "",
        "trein_rev": "",
        "trein_combinado_reg": "",
        "trein_combinado_rev": "",
        "aceite_reg": estilizar_tabela_html(
            _renomear_aceite_reg(aceite_reg),
            semaforo_coluna="% Aceite",
            meta_semaforo=META_ACEITES,
        ),
        "aceite_rev": "",
    }

    if trein_reg is not None and not trein_reg.empty:
        tabelas["trein_reg"] = estilizar_tabela_html(
            _renomear_trein_reg(trein_reg),
            semaforo_coluna="% Realizado",
            meta_semaforo=META_TREINAMENTOS,
        )
        tabelas["trein_rev"] = estilizar_tabela_html(
            _renomear_trein_rev(trein_rev.head(10)),
            semaforo_coluna="% Realizado",
            meta_semaforo=META_TREINAMENTOS,
        )

    if aceite_rev is not None and not aceite_rev.empty:
        tabelas["aceite_rev"] = estilizar_tabela_html(
            _renomear_aceite_rev(aceite_rev.head(10)),
            semaforo_coluna="% Aceite",
            meta_semaforo=META_ACEITES,
        )

    if trein_por_curso is not None:
        comb_reg = preparar_tabela_treinamentos_combinada(trein_por_curso, trein_ambos=trein_reg, nivel="regional")
        comb_rev = preparar_tabela_treinamentos_combinada(trein_por_curso, trein_ambos=trein_rev, nivel="revenda")
        if comb_reg is not None and not comb_reg.empty:
            tabelas["trein_combinado_reg"] = estilizar_tabela_html(
                comb_reg,
                semaforo_coluna="% Ambos",
                meta_semaforo=META_TREINAMENTOS,
            )
        if comb_rev is not None and not comb_rev.empty:
            tabelas["trein_combinado_rev"] = estilizar_tabela_html(
                comb_rev.head(10),
                semaforo_coluna="% Ambos",
                meta_semaforo=META_TREINAMENTOS,
            )
        # Mantem disponibilidade dos dados individuais, se necessario no futuro
        c1_reg = trein_por_curso.get("curso1_reg")
        c1_rev = trein_por_curso.get("curso1_rev")
        c2_reg = trein_por_curso.get("curso2_reg")
        c2_rev = trein_por_curso.get("curso2_rev")
        tabelas["trein_c1_reg"] = estilizar_tabela_html(_renomear_trein_reg(c1_reg)) if c1_reg is not None and not c1_reg.empty else ""
        tabelas["trein_c1_rev"] = estilizar_tabela_html(_renomear_trein_rev(c1_rev)) if c1_rev is not None and not c1_rev.empty else ""
        tabelas["trein_c2_reg"] = estilizar_tabela_html(_renomear_trein_reg(c2_reg)) if c2_reg is not None and not c2_reg.empty else ""
        tabelas["trein_c2_rev"] = estilizar_tabela_html(_renomear_trein_rev(c2_rev)) if c2_rev is not None and not c2_rev.empty else ""

    return tabelas


def _destaque_tom_ok_html(texto, imagens_kv=None, imagens_tom=None):
    """Gera card verde de destaque com Tom fazendo joinha ao lado, ajustado ao tamanho do texto."""
    if not texto:
        return ""
    imagens_tom = imagens_tom or {}
    imagens_kv = imagens_kv or {}
    img_base64 = imagens_tom.get("ok")
    cid = None
    if img_base64:
        cid = "tom_ok"
    elif "tom" in imagens_kv:
        cid = "kv_tom"
    tom_html = f'<img src="cid:{cid}" alt="Tom +TOP" width="60" style="display:block;">' if cid else ""
    return f"""
    <table cellpadding="0" cellspacing="0" border="0" style="margin:12px 0; width:auto; display:inline-table;">
      <tr>
        <td width="70" align="center" valign="middle" style="padding-right:6px;">
          {tom_html}
        </td>
        <td valign="middle" style="padding:6px 10px; background-color:#d4edda; border-radius:8px; border:1px solid #00a651; color:#155724; font-size:13px; line-height:1.35;">
          {texto}
        </td>
      </tr>
    </table>
    """


def _pontos_atencao_header_html():
    """Retorna cabeçalho destacado para a seção de pontos de atenção."""
def _balao_tom_html(texto, imagens_kv=None, imagens_tom=None, tipo_tom="tom", alinhamento="esquerda", largura_tom=90):
    """Cria balão de fala com imagem do Tom ao lado (Tom 'falando' o texto)."""
    imagens_kv = imagens_kv or {}
    imagens_tom = imagens_tom or {}

    # Prioriza imagem específica da pasta TOM, senão usa kv_tom
    img_base64 = imagens_tom.get(tipo_tom)
    cid = None
    if img_base64:
        cid = f"tom_{tipo_tom}"
    elif "tom" in imagens_kv:
        cid = "kv_tom"

    tom_html = f'<img src="cid:{cid}" alt="Tom +TOP" width="{largura_tom}" style="display:block;">' if cid else ""

    if alinhamento == "direita":
        return f"""
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;">
          <tr>
            <td valign="middle" style="padding:8px 12px; background-color:#ffffff; border-radius:12px; border:2px solid #00a651; color:#333333; font-size:15px; line-height:1.4;">
              {texto}
            </td>
            <td width="{largura_tom + 10}" align="center" valign="middle" style="padding-left:10px;">
              {tom_html}
            </td>
          </tr>
        </table>
        """
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;">
      <tr>
        <td width="{largura_tom + 10}" align="center" valign="middle" style="padding-right:10px;">
          {tom_html}
        </td>
        <td valign="middle" style="padding:8px 12px; background-color:#ffffff; border-radius:12px; border:2px solid #00a651; color:#333333; font-size:15px; line-height:1.4;">
          {texto}
        </td>
      </tr>
    </table>
    """


def _pontos_atencao_secao_html(titulo, subtitulo, tabela_html, imagens_kv=None, imagens_tom=None):
    """Monta bloco de pontos de atenção com título grande, subtítulo e tabela."""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
      <tr>
        <td style="padding:14px; background-color:#fff3cd; border-left:5px solid #ef4e22; color:#856404;">
          <div style="font-size:17px; font-weight:bold; margin-bottom:6px;">⚠️ {titulo}</div>
          <div style="font-size:14px; margin-bottom:12px;">{subtitulo}</div>
          {tabela_html}
        </td>
      </tr>
    </table>
    """



def _balao_insight_html(texto):
    """Gera balão branco com borda verde para insights, sem imagem do Tom."""
    if not texto:
        return ""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:12px 0;">
      <tr>
        <td style="padding:14px; background-color:#ffffff; border-radius:12px; border:2px solid #00a651; color:#333333; font-size:15px; line-height:1.5;">
          {texto}
        </td>
      </tr>
    </table>
    """


def _secao_html(titulo):
    """Retorna titulo de secao em tabela, compativel com Outlook."""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:28px;">
      <tr>
        <td style="color:#ef4e22; font-size:18px; font-weight:bold; padding-bottom:8px;">
          {titulo}
        </td>
      </tr>
    </table>
    """


def _card_html(conteudo, tipo="insight"):
    """Retorna card em tabela com borda lateral, compativel com Outlook."""
    if not conteudo:
        return ""

    cores = {
        "insight": ("#fff5f2", "#ef4e22"),
        "alerta": ("#fff8f0", "#ef4e22"),
        "parabens": ("#f0fff4", "#00a651"),
        "drive": ("#f0fff4", "#00a651"),
    }
    bg, borda = cores.get(tipo, ("#f4f4f4", "#999999"))

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;">
      <tr>
        <td width="4" bgcolor="{borda}" style="font-size:0; line-height:0;">&nbsp;</td>
        <td bgcolor="{bg}" style="padding:16px; color:#333333; font-size:16px; line-height:1.6;">
          {conteudo}
        </td>
      </tr>
    </table>
    """


def _img_html(cid, alt):
    if not cid:
        return ""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;">
      <tr>
        <td align="center">
          <img src="cid:{cid}" alt="{alt}" width="700" style="max-width:700px; width:100%; height:auto; display:block;">
        </td>
      </tr>
    </table>
    """


def carregar_imagens_kv():
    """Carrega imagens do KV +TOP, redimensiona e comprime para e-mail."""
    kv_dir = BASE_DIR / "ajustes_relatorio_envio" / "logo_kv" / "elementos_kv"
    imagens = {}
    arquivos = {
        "logo": (kv_dir / "_top_logo.png", 200),
        "tom": (kv_dir / "_top_tom_cubos1.png", 150),
        "fundo": (kv_dir / "_top_fundo.png", 700),
    }
    for nome, (path, max_width) in arquivos.items():
        if not path.exists():
            logger.warning(f"Imagem KV nao encontrada: {path}")
            continue
        try:
            from PIL import Image
            img = Image.open(path)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGBA")
            else:
                img = img.convert("RGB")
            if img.width > max_width:
                ratio = max_width / img.width
                new_size = (max_width, int(img.height * ratio))
                img = img.resize(new_size, Image.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="PNG", optimize=True)
            imagens[nome] = base64.b64encode(buffer.getvalue()).decode("utf-8")
        except Exception as e:
            logger.warning(f"Falha ao otimizar imagem {path.name}: {e}. Usando original.")
            with open(path, "rb") as f:
                imagens[nome] = base64.b64encode(f.read()).decode("utf-8")
    return imagens


def carregar_imagens_tom():
    """Carrega imagens específicas da pasta TOM, com fallback de nomes, otimizadas para e-mail."""
    tom_dir = BASE_DIR / "ajustes_relatorio_envio" / "TOM"
    if not tom_dir.exists():
        logger.warning(f"Pasta de imagens TOM não encontrada: {tom_dir}")
        return {}

    candidatos = {
        "apontando": ["TOM-APONTANDO.png", "TOM APONTANDO.png", "TOM-APONTANDO2.png"],
        "atencao": ["TOM_ATENÇÃO.png", "TOM ATENÇÃO.png", "TOM_ATENCAO.png", "TOM-ATENCAO.png"],
        "ok": ["TOM OK.png", "TOM-OK.png"],
        "incentivo": ["TOM-APONTANDO6.png", "TOM APONTANDO6.png", "TOM-VC_SABIA.png"],
    }

    imagens = {}
    for chave, nomes in candidatos.items():
        for nome in nomes:
            path = tom_dir / nome
            if path.exists():
                try:
                    from PIL import Image
                    img = Image.open(path)
                    if img.mode in ("RGBA", "P"):
                        img = img.convert("RGBA")
                    else:
                        img = img.convert("RGB")
                    # Redimensiona para largura máxima de 120px
                    max_width = 120
                    if img.width > max_width:
                        ratio = max_width / img.width
                        new_size = (max_width, int(img.height * ratio))
                        img = img.resize(new_size, Image.LANCZOS)
                    buffer = io.BytesIO()
                    img.save(buffer, format="PNG", optimize=True)
                    imagens[chave] = base64.b64encode(buffer.getvalue()).decode("utf-8")
                    break
                except Exception as e:
                    logger.warning(f"Falha ao otimizar imagem {path.name}: {e}")
        if chave not in imagens:
            logger.warning(f"Imagem TOM não encontrada para: {chave}")
    return imagens


def _metric_box(titulo, valor, meta=None):
    meta_html = f"{meta:.0f}%" if meta is not None else ""
    farol = farol_html(valor, meta, tamanho=38) if meta is not None else ""
    return f"""
    <td width="33%" align="center" valign="middle" bgcolor="#ffffff" style="padding:20px 24px; color:#333333; font-size:16px; font-weight:bold; border-radius:10px; border:3px solid #ef4e22;">
      <div style="font-size:16px; margin-bottom:6px; color:#ef4e22;">{titulo}</div>
      <div style="font-size:42px; margin-bottom:8px; color:#ef4e22; line-height:1; white-space:nowrap;">
        {farol}&nbsp;<strong>{valor}%</strong>
      </div>
      <div style="font-size:14px; color:#333333;">objetivo ideal <strong>{meta_html}</strong></div>
    </td>
    """


def _header_html(titulo, hoje, imagens_kv=None):
    """Header com identidade visual +TOP usando imagens do KV."""
    imagens_kv = imagens_kv or {}
    logo_cid = "kv_logo" if "logo" in imagens_kv else None
    tom_cid = "kv_tom" if "tom" in imagens_kv else None
    fundo_cid = "kv_fundo" if "fundo" in imagens_kv else None

    bg_style = f'background-image: url(cid:{fundo_cid}); background-size: cover; background-position: center;' if fundo_cid else 'background-color: #f5f5f5;'

    logo_html = f'<img src="cid:{logo_cid}" alt="Logo +TOP" width="180" style="display:block;">' if logo_cid else '<span style="font-size:24px; font-weight:bold;">+top</span>'
    tom_html = f'<img src="cid:{tom_cid}" alt="Tom +TOP" width="120" style="display:block;">' if tom_cid else ''

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff">
      <tr>
        <td style="padding:0; font-family:Arial, Helvetica, sans-serif; {bg_style}">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td style="padding:24px;" valign="middle">
                {logo_html}
              </td>
              <td align="right" style="padding:24px;" valign="bottom">
                {tom_html}
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td bgcolor="#ef4e22" style="padding:14px 24px; color:#ffffff; font-family:Arial, Helvetica, sans-serif;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td>
                <h1 style="margin:0; font-size:18px; font-weight:bold; color:#ffffff;">{titulo}</h1>
              </td>
              <td align="right">
                <p style="margin:0; font-size:12px; color:#ffffff;">{hoje}</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    """


def _destaques_cadastro_html(cad_rev, imagens_kv=None, imagens_tom=None):
    """Gera um único balão verde de destaque com revendas agrupadas por regional."""
    if cad_rev is None or cad_rev.empty:
        return ""

    atingiram = cad_rev[cad_rev["pct_ativos"] >= META_CADASTRO].copy()
    if atingiram.empty:
        return ""

    # Ordenar revendas dentro de cada regional do maior para o menor
    atingiram = atingiram.sort_values(["regional_curta", "pct_ativos"], ascending=[True, False])

    # Ordenar regionais pela média de % ativos (maior primeiro)
    ordem_regional = (
        atingiram.groupby("regional_curta")["pct_ativos"]
        .mean()
        .sort_values(ascending=False)
        .index.tolist()
    )

    imagens_tom = imagens_tom or {}
    imagens_kv = imagens_kv or {}
    img_base64 = imagens_tom.get("ok")
    cid = None
    if img_base64:
        cid = "tom_ok"
    elif "tom" in imagens_kv:
        cid = "kv_tom"
    tom_html = f'<img src="cid:{cid}" alt="Tom +TOP" width="100" style="display:block;">' if cid else ""

    blocos_regional = []
    for regional in ordem_regional:
        revendas_reg = atingiram[atingiram["regional_curta"] == regional]
        linhas_rev = []
        for _, r in revendas_reg.iterrows():
            linhas_rev.append(
                f"• <strong>{r['revenda']}</strong> — <strong>{r['pct_ativos']:.1f}%</strong>"
            )

        bloco = (
            f"<div style='margin-bottom:16px;'>"
            f"<div style='font-size:16px; font-weight:bold; margin-bottom:6px;'>"
            f"🟢 Regional <strong>{regional}</strong></div>"
            f"<div style='font-size:14px; line-height:1.7;'>"
            + "<br>".join(linhas_rev)
            + "</div></div>"
        )
        blocos_regional.append(bloco)

    conteudo = (
        f"<div style='font-size:18px; font-weight:bold; margin-bottom:4px;'>"
        f"🎉 Parabéns!</div>"
        f"<div style='font-size:15px; font-weight:bold; margin-bottom:12px;'>"
        f"Melhor % de cadastros ativos.</div>"
        f"<div style='font-size:14px; margin-bottom:12px;'>"
        f"As revendas abaixo atingiram ou superaram o objetivo de <strong>{META_CADASTRO:.0f}%</strong>:</div>"
        + "\n".join(blocos_regional)
    )

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0;">
      <tr>
        <td width="110" align="center" valign="middle" style="padding-right:10px;">
          {tom_html}
        </td>
        <td valign="middle" style="padding:14px 16px; background-color:#d4edda; border-radius:12px; border:1px solid #00a651; color:#155724; font-size:14px; line-height:1.6;">
          {conteudo}
        </td>
      </tr>
    </table>
    """


def montar_email_html(dados, graficos, tabelas, insights, link_drive, teste=False, regional_filtro=None):
    """Monta corpo do e-mail em HTML compativel com Gmail e Outlook."""
    hoje = date.today().strftime("%d/%m/%Y")
    titulo = f"Relatorio Semanal Programa +TOP — {regional_filtro}" if regional_filtro else "Relatorio Semanal Programa +TOP"

    alerta_teste = "<p style='color:#d9534f; font-weight:bold; margin:16px 0;'>[MODO TESTE - e-mail nao enviado]</p>" if teste else ""

    # Carrega imagens do KV e imagens específicas do Tom
    imagens_kv = carregar_imagens_kv()
    imagens_tom = carregar_imagens_tom()

    # Metricas principais
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    aceite_reg = dados["aceites"]

    pct_geral = round(cad_reg["ativos"].sum() / cad_reg["total"].sum() * 100, 1) if not cad_reg.empty else 0
    pct_trein = round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1) if trein_reg is not None and not trein_reg.empty else 0
    pct_aceite = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1) if not aceite_reg.empty else 0

    metricas = f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0;">
      <tr>
        <td align="center">
          <table cellpadding="0" cellspacing="8" border="0">
            <tr>
              {_metric_box('Cadastros', pct_geral, META_CADASTRO)}
              {_metric_box('Treinamentos', pct_trein, META_TREINAMENTOS)}
              {_metric_box('Aceites', pct_aceite, META_ACEITES)}
            </tr>
          </table>
          <p style="font-size:11px; color:#666666; margin-top:6px;">
            🟢 Atingiu a meta &nbsp;|&nbsp; 🟡 Entre 70% e a meta &nbsp;|&nbsp; 🔴 Abaixo de 70% da meta
          </p>
        </td>
      </tr>
    </table>
    """

    # Se for email por regional, filtra as tabelas
    if regional_filtro:
        tabelas_usar = {
            "cad_reg": estilizar_tabela_html(_renomear_cadastro_reg(cad_reg[cad_reg["regional_curta"] == regional_filtro])),
            "cad_rev": estilizar_tabela_html(_renomear_cadastro_rev(cad_rev[cad_rev["regional_curta"] == regional_filtro])),
            "trein_reg": "",
            "trein_rev": "",
            "trein_combinado_reg": "<p><em>Sem dados de treinamentos.</em></p>",
            "trein_combinado_rev": "<p><em>Sem dados de treinamentos.</em></p>",
            "trein_c1_reg": "",
            "trein_c1_rev": "",
            "trein_c2_reg": "",
            "trein_c2_rev": "",
            "aceite_reg": estilizar_tabela_html(_renomear_aceite_reg(aceite_reg[aceite_reg["regional"] == regional_filtro])),
            "aceite_rev": "",
        }
        trein_por_curso = dados.get("treinamentos_por_curso")
        trein_reg = dados["treinamentos"][0]
        trein_rev = dados["treinamentos"][1]
        if trein_por_curso is not None:
            comb_reg = preparar_tabela_treinamentos_combinada(trein_por_curso, trein_ambos=trein_reg, nivel="regional")
            comb_rev = preparar_tabela_treinamentos_combinada(trein_por_curso, trein_ambos=trein_rev, nivel="revenda")
            if comb_reg is not None and not comb_reg.empty:
                tabelas_usar["trein_combinado_reg"] = estilizar_tabela_html(
                    comb_reg[comb_reg["Regional"] == regional_filtro]
                )
            if comb_rev is not None and not comb_rev.empty:
                tabelas_usar["trein_combinado_rev"] = estilizar_tabela_html(
                    comb_rev[comb_rev["Regional"] == regional_filtro]
                )
    else:
        tabelas_usar = tabelas

    # Nomes dos cursos
    nome_curso1 = dados["cursos_info"][2] if dados.get("cursos_info") else "Curso 1"
    nome_curso2 = dados["cursos_info"][3] if dados.get("cursos_info") else "Curso 2"

    def subsecao_titulo(texto):
        return f"""
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px;">
          <tr><td style="color:#ef4e22; font-size:15px; font-weight:bold;">{texto}</td></tr>
        </table>
        """

    html = f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{titulo}</title>
    </head>
    <body style="margin:0; padding:20px; background-color:#f5f5f5; font-family:Arial, Helvetica, sans-serif; color:#333333; line-height:1.6;">
        <!--[if mso]>
        <table role="presentation" width="700" cellspacing="0" cellpadding="0" border="0" align="center">
        <tr><td>
        <![endif]-->
        <table role="presentation" width="100%" max-width="700" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:700px; width:100%; background-color:#ffffff;">
          <tr>
            <td>

              {_header_html(titulo, hoje, imagens_kv)}

              <!-- Conteudo -->
              <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff">
                <tr><td style="padding:24px;">

                  {alerta_teste}

                  {metricas}

                  {_secao_html("CADASTROS")}
                  {_balao_tom_html(
                      f"<p style='font-size:20px; margin:0 0 10px 0; line-height:1.4;'><strong>Nosso objetivo ideal para a base de cadastros do Programa +TOP é de <span style='font-size:28px; color:#ef4e22;'>{META_CADASTRO:.0f}%</span></strong></p>"
                      f"<p style='font-size:17px; margin:0; line-height:1.4;'>Até o momento, estamos com <strong>{pct_geral}% da base ativa</strong>.</p>",
                      imagens_kv=imagens_kv, imagens_tom=imagens_tom, tipo_tom="apontando", alinhamento="esquerda"
                  )}
                  {subsecao_titulo("Por Regional")}
                  {tabelas_usar['cad_reg']}
                  {_img_html("grafico_cadastros" if "cadastros" in graficos else None, "Gráfico Cadastros")}
                  {_destaques_cadastro_html(cad_rev, imagens_kv=imagens_kv, imagens_tom=imagens_tom)}
                  {_pontos_atencao_secao_html(
                      insights.get("cadastros", {}).get("alerta_titulo", ""),
                      insights.get("cadastros", {}).get("alerta_subtitulo", ""),
                      estilizar_tabela_html(insights.get("cadastros", {}).get("alerta_itens"), semaforo_coluna="% Ativos", meta_semaforo=META_CADASTRO) if insights.get("cadastros", {}).get("alerta_itens") is not None else "",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                  )}
                  {subsecao_titulo("TOP 10 revendas")}
                  {tabelas_usar['cad_rev']}

                  {_secao_html("TREINAMENTOS")}
                  {_balao_tom_html(
                      f"<p style='font-size:20px; margin:0 0 10px 0; line-height:1.4;'><strong>Nosso objetivo ideal para os treinamentos do Programa +TOP é de <span style='font-size:28px; color:#ef4e22;'>{META_TREINAMENTOS:.0f}%</span></strong>.</p>"
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4;'>Os dois conteúdos de <strong>{nome_mes_pt_br().lower()}</strong> estão disponíveis até <strong>{date.today().replace(day=30):%d/%m/%Y}</strong></p>"
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4;'><strong>Conteúdos:</strong><br>"
                      f"1. <strong>{nome_curso1}</strong> (SKU {dados['cursos_info'][4]})<br>"
                      f"2. <strong>{nome_curso2}</strong> (SKU {dados['cursos_info'][5]}).</p>"
                      f"<p style='font-size:17px; margin:0; line-height:1.4;'>Até o momento, estamos em <strong>{pct_trein}% da base ativa</strong>.</p>",
                      imagens_kv=imagens_kv, imagens_tom=imagens_tom, tipo_tom="apontando", alinhamento="esquerda"
                  )}
                  {subsecao_titulo("Por Regional")}
                  {tabelas_usar['trein_combinado_reg'] if tabelas_usar.get('trein_combinado_reg') else '<p><em>Não foi possível identificar os 2 treinamentos obrigatórios do mês.</em></p>'}
                  {_img_html("grafico_treinamentos" if "treinamentos" in graficos else None, "Gráfico Treinamentos")}
                  {_pontos_atencao_secao_html(
                      insights.get("treinamentos", {}).get("alerta_titulo", ""),
                      insights.get("treinamentos", {}).get("alerta_subtitulo", ""),
                      estilizar_tabela_html(insights.get("treinamentos", {}).get("alerta_itens"), semaforo_coluna="% Realizado", meta_semaforo=META_TREINAMENTOS) if insights.get("treinamentos", {}).get("alerta_itens") is not None else "",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                  )}
                  {_destaque_tom_ok_html(insights.get("treinamentos", {}).get("destaque", ""), imagens_kv=imagens_kv, imagens_tom=imagens_tom)}
                  {subsecao_titulo("Por Revenda")}
                  {tabelas_usar['trein_combinado_rev'] if tabelas_usar.get('trein_combinado_rev') else '<p><em>Sem dados de treinamentos.</em></p>'}

                  <p style="font-size:12px; color:#666666; font-style:italic; margin-top:8px;">
                    Lembrando que desde abril temos a mecânica adicional: a cada 3 meses consecutivos que o vendedor concluir todos os treinamentos, ele ganha +100 pontos no programa.
                  </p>

                  {_secao_html("ACEITES MENSAIS")}
                  {_balao_tom_html(
                      f"<p style='font-size:20px; margin:0 0 10px 0; line-height:1.4;'><strong>Nosso objetivo ideal para os aceites mensais de {nome_mes_pt_br(ano_mes=str(dados['mes_aceite']))} é de <span style='font-size:28px; color:#ef4e22;'>{META_ACEITES:.0f}%</span>.</strong></p>"
                      f"<p style='font-size:17px; margin:0; line-height:1.4;'>Até o momento, <strong>{pct_aceite}% dos participantes ativos deram aceite</strong>.</p>",
                      imagens_kv=imagens_kv, imagens_tom=imagens_tom, tipo_tom="apontando", alinhamento="esquerda"
                  )}
                  {subsecao_titulo("Por Regional")}
                  {tabelas_usar['aceite_reg']}
                  {_img_html("grafico_aceites" if "aceites" in graficos else None, "Gráfico Aceites")}
                  {_pontos_atencao_secao_html(
                      insights.get("aceites", {}).get("alerta_titulo", ""),
                      insights.get("aceites", {}).get("alerta_subtitulo", ""),
                      estilizar_tabela_html(insights.get("aceites", {}).get("alerta_itens"), semaforo_coluna="% Aceite", meta_semaforo=META_ACEITES) if insights.get("aceites", {}).get("alerta_itens") is not None else "",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                  )}
                  {_destaque_tom_ok_html(insights.get("aceites", {}).get("destaque", ""), imagens_kv=imagens_kv, imagens_tom=imagens_tom)}
                  {subsecao_titulo("TOP 10 revendas")}
                  {tabelas_usar['aceite_rev'] if tabelas_usar.get('aceite_rev') else '<p><em>Sem dados de aceites por revenda.</em></p>'}

                  <p style="margin-top:28px; font-size:16px; color:#155724; background-color:#d4edda; padding:14px 16px; border-radius:10px; border:1px solid #00a651; line-height:1.5;">
                    💪 <strong>Conto com o reforço das regionais para revertermos isso.</strong><br>
                    Juntos, vamos fortalecer ainda mais o Programa +TOP!
                  </p>

                  <p style="margin-top:24px; font-size:16px;">Em anexo <strong>base detalhada</strong>.</p>

                  <p style="margin-top:16px;">Att.,<br>Relatório Automático Programa +TOP</p>

                </td></tr>
              </table>

            </td>
          </tr>
        </table>
        <!--[if mso]>
        </td></tr></table>
        <![endif]-->
    </body>
    </html>
    """
    return html, imagens_kv, imagens_tom


# ---------------------------------------------------------------------------
# ENVIO DE E-MAIL
# ---------------------------------------------------------------------------
def enviar_email(html_body, config, graficos, destinatarios, anexos=None, teste=False, assunto=None, imagens_kv=None, imagens_tom=None):
    """Envia e-mail HTML com imagens embutidas e anexos opcionais."""
    if teste:
        logger.info("MODO TESTE: e-mail não será enviado.")
        return False

    if not config:
        logger.error("Configuração de e-mail não encontrada. Verifique config_email.json")
        return False

    # Filtra apenas e-mails válidos (strings não vazias)
    destinatarios = [str(d).strip() for d in destinatarios if d and not pd.isna(d) and str(d).strip()]
    if not destinatarios:
        logger.warning("Nenhum destinatário válido encontrado. E-mail não será enviado.")
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = assunto or config.get("assunto", f"Relatório Semanal Programa +TOP - {date.today():%d/%m/%Y}")
    msg["From"] = formataddr((config.get("remetente_nome", "Relatório +TOP"), config["remetente_email"]))
    msg["To"] = ", ".join(destinatarios)

    related = MIMEMultipart("related")
    related.attach(MIMEText(html_body, "html", _charset="utf-8"))

    for cid, img_base64 in graficos.items():
        img_data = base64.b64decode(img_base64)
        mime_img = MIMEImage(img_data)
        mime_img.add_header("Content-ID", f"<grafico_{cid}>")
        mime_img.add_header("Content-Disposition", "inline", filename=f"grafico_{cid}.png")
        related.attach(mime_img)

    # Anexa imagens do KV (logo, tom, fundo)
    if imagens_kv:
        for nome, img_base64 in imagens_kv.items():
            img_data = base64.b64decode(img_base64)
            mime_img = MIMEImage(img_data)
            mime_img.add_header("Content-ID", f"<kv_{nome}>")
            mime_img.add_header("Content-Disposition", "inline", filename=f"kv_{nome}.png")
            related.attach(mime_img)

    # Anexa imagens específicas do Tom
    if imagens_tom:
        for nome, img_base64 in imagens_tom.items():
            img_data = base64.b64decode(img_base64)
            mime_img = MIMEImage(img_data)
            mime_img.add_header("Content-ID", f"<tom_{nome}>")
            mime_img.add_header("Content-Disposition", "inline", filename=f"tom_{nome}.png")
            related.attach(mime_img)

    msg.attach(related)

    if anexos:
        for caminho in anexos:
            with open(caminho, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {caminho.name}")
            msg.attach(part)

    try:
        server = smtplib.SMTP(config["smtp_host"], config["smtp_port"], timeout=30)
        server.starttls()
        server.login(config["remetente_email"], config["senha_app"])
        server.sendmail(config["remetente_email"], destinatarios, msg.as_string())
        server.quit()
        logger.info(f"E-mail enviado com sucesso para: {', '.join(destinatarios)}")
        return True
    except Exception as e:
        logger.error(f"Falha ao enviar e-mail: {e}")
        return False


def carregar_config_email():
    """Carrega configurações de e-mail do arquivo JSON."""
    if not CONFIG_FILE.exists():
        return None
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
    env_senha = os.environ.get("SMTP_APP_PASSWORD")
    if env_senha:
        config["senha_app"] = env_senha

    required = ["smtp_host", "smtp_port", "remetente_email", "senha_app"]
    missing = [field for field in required if not config.get(field)]
    if missing:
        logger.warning(f"Configuração de e-mail incompleta. Campos ausentes: {missing}")

    return config


# ---------------------------------------------------------------------------
# EXPORTAÇÃO EXCEL
# ---------------------------------------------------------------------------
def nome_aba_excel(nome, sufixo=""):
    """Trunca nome para limite de 31 caracteres do Excel, preservando sufixo."""
    nome_limpo = str(nome).strip().replace("/", "-")
    if sufixo:
        max_nome = 31 - len(sufixo) - 1
        if len(nome_limpo) > max_nome:
            nome_limpo = nome_limpo[:max_nome]
        return f"{nome_limpo}_{sufixo}"
    return nome_limpo[:31]


def preparar_aba_detalhamento(df_det, dados=None, regional_filtro=None):
    """Prepara aba de detalhamento com campos solicitados, incluindo cursos realizados."""
    if df_det is None or df_det.empty:
        return None

    df = df_det.copy()
    if regional_filtro:
        df = df[df["regional_da_loja"].apply(regional_curta) == regional_filtro]

    colunas_map = {
        "regional_da_loja": "Regional",
        "loja": "Revenda",
        "cnpj_loja": "CNPJ",
        "nome": "Nome",
        "cargo": "Cargo",
        "status": "Status",
        "cidade": "Cidade",
        "uf": "UF",
        "bairro": "Bairro loja",
        "unnamed: 14": "Nome loja",
    }

    for col in colunas_map:
        if col not in df.columns:
            df[col] = None

    df_out = df[list(colunas_map.keys())].rename(columns=colunas_map)
    # Mantém cpf_limp apenas para cruzamento interno (não exporta)
    df_out["cpf_limp"] = df["cpf_limp"].values

    def limpar_cnpj(cnpj):
        if pd.isna(cnpj):
            return ""
        s = str(cnpj).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
        if "." in s:
            s = s.split(".")[0]
        return s.zfill(14)

    df_out["CNPJ"] = df_out["CNPJ"].apply(limpar_cnpj)
    # Força CNPJ como texto no Excel (evita notação científica)
    df_out["CNPJ"] = df_out["CNPJ"].apply(lambda x: f"'{x}" if x else x)

    # Adiciona colunas de treinamentos obrigatórios do mês
    if dados is not None:
        df_trein = dados.get("treinamentos_df")
        cursos_info = dados.get("cursos_info")
        mes_referencia = dados.get("mes_referencia")
        if df_trein is not None and cursos_info is not None and mes_referencia is not None:
            curso1, curso2, nome1, nome2, sku1, sku2 = cursos_info
            if curso1 and curso2:
                mes_dt = pd.Period(mes_referencia, freq="M")
                trein_mes = df_trein[
                    (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
                    & (df_trein["Estado"].str.lower() == "concluido")
                ].copy()

                for curso, nome, sku in [(curso1, nome1, sku1), (curso2, nome2, sku2)]:
                    if not curso:
                        continue
                    trein_curso = trein_mes[trein_mes["Curso"] == curso][["cpf_limp", "Conclusão"]].drop_duplicates("cpf_limp")
                    trein_curso["status_curso"] = "Realizado/Aprovado"
                    trein_curso["data_curso"] = trein_curso["Conclusão"].dt.strftime("%d/%m/%Y")

                    df_out = df_out.merge(
                        trein_curso[["cpf_limp", "status_curso", "data_curso"]],
                        left_on="cpf_limp",
                        right_on="cpf_limp",
                        how="left",
                    )
                    col_status = f"{nome} (SKU {sku}) - Status"
                    col_data = f"{nome} (SKU {sku}) - Conclusão"
                    df_out = df_out.rename(columns={
                        "status_curso": col_status,
                        "data_curso": col_data,
                    })
                    df_out[col_status] = df_out[col_status].fillna("Não realizado")
                    df_out[col_data] = df_out[col_data].fillna("")

            # Remove coluna interna de CPF antes de exportar
            if "cpf_limp" in df_out.columns:
                df_out = df_out.drop(columns=["cpf_limp"])

    return df_out


def salvar_relatorio_excel(dados, caminho):
    """Salva relatório consolidado em Excel com múltiplas abas."""
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    trein_por_curso = dados.get("treinamentos_por_curso")
    aceite_reg = dados["aceites"]
    df_det = dados.get("detalhamento")

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        _renomear_cadastro_reg(cad_reg).to_excel(writer, sheet_name="Cadastro_Regional", index=False)
        _renomear_cadastro_rev(cad_rev).to_excel(writer, sheet_name="Cadastro_Revenda", index=False)

        if trein_reg is not None:
            _renomear_trein_reg(trein_reg).to_excel(writer, sheet_name="Treinamento_Regional", index=False)
        if trein_rev is not None:
            _renomear_trein_rev(trein_rev).to_excel(writer, sheet_name="Treinamento_Revenda", index=False)

        if trein_por_curso is not None:
            c1_reg = trein_por_curso.get("curso1_reg")
            c1_rev = trein_por_curso.get("curso1_rev")
            c2_reg = trein_por_curso.get("curso2_reg")
            c2_rev = trein_por_curso.get("curso2_rev")
            nome1 = trein_por_curso.get("nome1", "Curso 1")
            nome2 = trein_por_curso.get("nome2", "Curso 2")
            if c1_reg is not None and not c1_reg.empty:
                _renomear_trein_reg(c1_reg).to_excel(writer, sheet_name=nome_aba_excel(nome1, "Reg"), index=False)
            if c1_rev is not None and not c1_rev.empty:
                _renomear_trein_rev(c1_rev).to_excel(writer, sheet_name=nome_aba_excel(nome1, "Rev"), index=False)
            if c2_reg is not None and not c2_reg.empty:
                _renomear_trein_reg(c2_reg).to_excel(writer, sheet_name=nome_aba_excel(nome2, "Reg"), index=False)
            if c2_rev is not None and not c2_rev.empty:
                _renomear_trein_rev(c2_rev).to_excel(writer, sheet_name=nome_aba_excel(nome2, "Rev"), index=False)

        _renomear_aceite_reg(aceite_reg).to_excel(writer, sheet_name="Aceite_Regional", index=False)

        df_det_out = preparar_aba_detalhamento(df_det, dados=dados)
        if df_det_out is not None:
            df_det_out.to_excel(writer, sheet_name="Detalhamento", index=False)

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
                int(cad_reg["total"].sum()),
                int(cad_reg["ativos"].sum()),
                round(cad_reg["ativos"].sum() / cad_reg["total"].sum() * 100, 1),
                int(trein_reg["realizaram"].sum()) if trein_reg is not None else "N/A",
                round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1) if trein_reg is not None else "N/A",
                int(aceite_reg["aceitaram"].sum()),
                round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1),
            ],
        }
        pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)

    logger.info(f"Relatório Excel salvo em: {caminho}")


def salvar_relatorio_excel_regional(dados, caminho, regional_filtro):
    """Salva relatório filtrado por regional em Excel."""
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    trein_por_curso = dados.get("treinamentos_por_curso")
    aceite_reg = dados["aceites"]
    df_det = dados.get("detalhamento")

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        _renomear_cadastro_reg(cad_reg[cad_reg["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name="Cadastro_Regional", index=False)
        _renomear_cadastro_rev(cad_rev[cad_rev["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name="Cadastro_Revenda", index=False)

        if trein_reg is not None:
            _renomear_trein_reg(trein_reg[trein_reg["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name="Treinamento_Regional", index=False)
        if trein_rev is not None:
            _renomear_trein_rev(trein_rev[trein_rev["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name="Treinamento_Revenda", index=False)

        if trein_por_curso is not None:
            c1_reg = trein_por_curso.get("curso1_reg")
            c1_rev = trein_por_curso.get("curso1_rev")
            c2_reg = trein_por_curso.get("curso2_reg")
            c2_rev = trein_por_curso.get("curso2_rev")
            nome1 = trein_por_curso.get("nome1", "Curso 1")
            nome2 = trein_por_curso.get("nome2", "Curso 2")
            if c1_reg is not None and not c1_reg.empty:
                _renomear_trein_reg(c1_reg[c1_reg["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name=nome_aba_excel(nome1, "Reg"), index=False)
            if c1_rev is not None and not c1_rev.empty:
                _renomear_trein_rev(c1_rev[c1_rev["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name=nome_aba_excel(nome1, "Rev"), index=False)
            if c2_reg is not None and not c2_reg.empty:
                _renomear_trein_reg(c2_reg[c2_reg["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name=nome_aba_excel(nome2, "Reg"), index=False)
            if c2_rev is not None and not c2_rev.empty:
                _renomear_trein_rev(c2_rev[c2_rev["regional_curta"] == regional_filtro]).to_excel(writer, sheet_name=nome_aba_excel(nome2, "Rev"), index=False)

        _renomear_aceite_reg(aceite_reg[aceite_reg["regional"] == regional_filtro]).to_excel(writer, sheet_name="Aceite_Regional", index=False)

        df_det_out = preparar_aba_detalhamento(df_det, dados=dados, regional_filtro=regional_filtro)
        if df_det_out is not None:
            df_det_out.to_excel(writer, sheet_name="Detalhamento", index=False)

    logger.info(f"Relatório regional ({regional_filtro}) Excel salvo em: {caminho}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Relatório Semanal Programa +TOP")
    parser.add_argument("--teste", action="store_true", help="Gera relatório local sem enviar e-mail")
    args = parser.parse_args()

    logger.info("=" * 60)
    logger.info("Iniciando geração do relatório semanal +TOP")
    logger.info(f"Modo: {'TESTE' if args.teste else 'PRODUÇÃO'}")

    try:
        bases = carregar_bases()
        df_cad = bases["cadastro"]
        df_trein = bases["treinamentos"]
        df_aceite = bases["aceites"]
        df_emails_reg = bases["emails_regionais"]
        df_det = bases["detalhamento"]
        ano_mes = bases["mes_referencia"]

        # Cálculos
        cad_reg, cad_rev = calcular_cadastros(df_cad)
        trein_reg, trein_rev, trein_por_curso, cursos_info = calcular_treinamentos(df_trein, df_cad, ano_mes)
        aceite_reg, aceite_rev, mes_aceite_ref = calcular_aceites(
            df_aceite, df_cad, ano_mes, bases["aba_aceite"], bases["usou_ultima_aba"]
        )

        dados = {
            "cadastros": (cad_reg, cad_rev),
            "treinamentos": (trein_reg, trein_rev),
            "treinamentos_df": df_trein,
            "treinamentos_por_curso": trein_por_curso,
            "aceites": aceite_reg,
            "aceites_rev": aceite_rev,
            "detalhamento": df_det,
            "cursos_info": cursos_info,
            "mes_aceite": mes_aceite_ref,
            "mes_referencia": ano_mes,
            "usou_ultima_aba": bases["usou_ultima_aba"],
        }

        # Snapshot e evolução
        snapshot_anterior = carregar_snapshot_anterior()
        salvar_snapshot(dados)
        evolucao = calcular_evolucao(
            {
                "cadastros": cad_reg.set_index("regional_curta")[["ativos", "total", "pct_ativos"]].to_dict("index"),
            },
            snapshot_anterior,
        )

        # Gráficos (apenas para email consolidado)
        graficos = gerar_graficos(cad_reg, trein_reg, aceite_reg)

        # Salvar Excel consolidado
        excel_path = OUTPUT_DIR / f"relatorio_top_{date.today():%Y%m%d}.xlsx"
        salvar_relatorio_excel(dados, excel_path)

        # Configurações
        config = carregar_config_email()
        link_drive = config.get("link_drive", "") if config else ""

        # Prepara tabelas e insights
        tabelas = preparar_tabelas(dados)
        insights = gerar_insights(cad_reg, cad_rev, trein_reg, trein_rev, aceite_reg, aceite_rev, evolucao, dados)

        # Montar e-mail consolidado
        html, imagens_kv, imagens_tom = montar_email_html(dados, graficos, tabelas, insights, link_drive, teste=args.teste)

        # Salvar cópia HTML
        html_path = OUTPUT_DIR / f"relatorio_top_{date.today():%Y%m%d}.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)

        # Enviar e-mail consolidado para cliente principal e cópias
        destinatarios = config.get("destinatarios", []) if config else []
        cc = config.get("cc", []) or []
        destinatarios_total = destinatarios + cc

        enviado_principal = enviar_email(
            html, config, graficos, destinatarios_total, anexos=[excel_path], teste=args.teste,
            imagens_kv=imagens_kv, imagens_tom=imagens_tom
        )

        if args.teste:
            logger.info(f"Modo teste: relatório gerado em {excel_path} e {html_path}")
        elif enviado_principal:
            logger.info("Relatório enviado com sucesso.")
        else:
            logger.error("Relatório gerado, mas não foi possível enviar o e-mail principal.")

    except Exception as e:
        logger.exception("Erro durante geração do relatório")
        sys.exit(1)


if __name__ == "__main__":
    main()
