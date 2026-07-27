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
import re
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

# Validador pré-envio (importado no escopo para evitar dependência circular)
from validar_envio_relatorio import ValidadorEnvioRelatorio

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

matplotlib.use("Agg")  # backend não interativo para cron/servers

# Padroniza fonte dos gráficos com Arial (única fonte permitida no email)
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = ["Arial"]
plt.rcParams["axes.titlesize"] = 14
plt.rcParams["axes.labelsize"] = 11
plt.rcParams["xtick.labelsize"] = 10
plt.rcParams["ytick.labelsize"] = 10
plt.rcParams["figure.titlesize"] = 14

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES GLOBAIS
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
LOG_DIR = BASE_DIR / "logs"
SNAPSHOT_DIR = BASE_DIR / "snapshots"
CONFIG_FILE = BASE_DIR / "config_email.json"

# Colunas percentuais no Excel que devem receber number_format 0.0%
PCT_COLS = {
    "% Ativos", "% Realizado", "% Aceite", "% Ambos",
}

# Cores do KV +TOP para formatação do Excel
COR_PRIMARIA = "EF4E22"       # laranja Whirlpool/+TOP
COR_PRIMARIA_CLARA = "FDE8E0" # laranja claro para cabeçalhos
COR_CINZA = "F2F2F2"          # cinza claro
COR_BRANCO = "FFFFFF"
COR_TEXTO = "333333"

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
SNAPSHOT_DIR.mkdir(exist_ok=True)

# Revendas excluídas/teste (conforme memória do projeto)
REVENDAS_EXCLUIR = {"EAI", "Whirlpool", "Novo Mundo", "ELETROMÓVEIS MARTINELLO"}

# Status que NÃO devem compor as métricas de cadastro (base operacional)
STATUS_CADASTRO_EXCLUIR = {"Inativo", "Bloqueado", "Reprovado", "Aguardando Aprovação"}

# Caminho da base do banco de participantes (usada para validar DataInclusao de Pré-Cadastrados)
BANCO_PARTICIPANTES_PATH = Path("/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx")

# Base com o nome real da loja/filial por CPF (coluna auxiliar Unnamed: 14)
LOJA_POR_CPF_PATH = DATA_DIR / "WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx"

# Pasta com as hierarquias mensais por revenda.
# (anteriormente hierarquia_rodrigo; agora bases_cadastro_hierarquia)
HIERARQUIA_DIR = DATA_DIR / "bases_cadastro_hierarquia"

# (DESATIVADO) A base consolidada não é mais utilizada. cadastro.xlsx é a base
# principal e a pasta de hierarquias enriquece os dados.
CONSOLIDADO_HIERARQUIA_CADASTRO_SITE_PATH = HIERARQUIA_DIR / "consolidado_hierarquia_com_cadastro_site.xlsx"

# Prazo em dias para considerar Pré-Cadastrado como Inativo (conforme regulamento)
PRAZO_PRE_CADASTRO_INATIVO = 90

# Metas/objetivos ideais para os indicadores (%)
META_CADASTRO = 85.0
META_TREINAMENTOS = 70.0
META_ACEITES = 70.0

# Mapeamento de nomes de revenda para exibição (evita siglas)
NOME_REVENDA_EXIBICAO = {
    "CDA": "Casas da Água",
    "Imperio": "Império",
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
    "LOJA SOLAR": "Solar",
    # Outros nomes exatos da hierarquia para normalizar
    "GAZIN ATACADO": "Gazin Atacado",
    "GAZIN ONLINE": "Gazin Online",
    "GAZIN": "Gazin Varejo",
    "LOJAS GUAIBIM": "Guaibim",
    "GUAIBIM": "Guaibim",
    "HAVAN": "Havan",
    "JMAHFUZ": "Jmahfuz",
    "IMPERIO": "Império",
    "LASER ELETRO": "Laser Eletro",
    "NOSSO LAR LOJAS DE DEPTOS LTDA": "Nosso Lar",
    "MILLENA MOVEIS": "Millena",
    "BECKER": "Becker",
    "BEMOL": "Bemol",
    "COLOMBO": "Colombo",
    "Colombo": "Colombo",
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
    "LIDER": "Líder",
    "LÍDER": "Líder",
    # Filial da Zenir (Filial 68, CNPJ raiz 41.426.966) que veio na hierarquia
    # sem o prefixo "Zenir" na coluna REVENDA — agrupada manualmente na Zenir.
    "MESSEJANA 2": "Zenir",
}

# Mapeamento SKU/Código -> Nome descritivo do curso obrigatório do mês
# Atualizado conforme catálogo Whirlpool/Brastemp/Consul.
NOMES_CURSOS = {
    "CZD12": "Ar-Condicionado Consul",
    "BRM46": "Geladeira Consul Inverse",
    "BMC29": "Micro-ondas Consul",
    "CRM44M": "Geladeira Consul Frost Free",
    "BRM44": "Geladeira Consul Frost Free",
    "CWN15": "Lavadora Consul",
    "CWN13": "Lavadora Consul",
    "BRE66": "Refrigerador Electrolux",
    "BRE57": "Refrigerador Electrolux",
    "BWJ14": "Lava-Louças Brastemp",
    "CRA30M": "Ar-Condicionado Consul",
    "CFO4ZAB": "Fogão Consul",
    "BRM62": "Geladeira Brastemp",
    "CRM56M": "Geladeira Consul",
    "CRM53M": "Geladeira Consul",
    "BRM52": "Geladeira Brastemp Frost Free Duplex",
    "CWN14": "Máquina de Lavar Consul",
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
    # Usa word boundaries para evitar falsos positivos (ex: GAZIN dentro de MAGAZINE)
    for chave, principal in sorted(MAPA_REVENDA_PRINCIPAL.items(), key=lambda x: -len(x[0])):
        chave_limpa = re.sub(r"\s+", " ", chave.upper())
        if re.search(r"\b" + re.escape(chave_limpa) + r"\b", nome_upper):
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
    r = r.title()
    # Ajustes específicos de exibição
    r = r.replace("Centro Norte", "Centro-Norte")
    return r



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
        return None, None

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

    # Seleciona arquivos: prefere versões *_corrigido.xlsx quando existirem
    arquivos_brutos = sorted(HIERARQUIA_DIR.glob("*.xlsx"))
    arquivos_corrigidos = {a.stem.replace("_corrigido", ""): a for a in arquivos_brutos if a.stem.endswith("_corrigido")}
    arquivos_usar = []
    for arquivo in arquivos_brutos:
        nome = arquivo.name.upper()
        if "PROD_WHP" in nome or nome.startswith("~") or "_corrigido" in nome:
            continue
        # Se existe versão corrigida, usa ela em vez do original
        if arquivo.stem in arquivos_corrigidos:
            arquivos_usar.append(arquivos_corrigidos[arquivo.stem])
        else:
            arquivos_usar.append(arquivo)
    # Remove duplicatas mantendo ordem
    arquivos_usar = list(dict.fromkeys(arquivos_usar))

    for arquivo in arquivos_usar:
        nome = arquivo.name.upper()
        # Ignora arquivos consolidados, de comparacao, resumo ou analise de CPFs repetidos
        if any(x in nome for x in ["CONSOLIDADO", "COMPARACAO", "RESUMO", "CPFS_REPETIDOS"]):
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

        # Revenda derivada do nome do arquivo (ex: "GUAIBIM - HIERARQUIA MAIO.xlsx" -> "Guaibim").
        # Usada quando o arquivo não traz a coluna REVENDA ou ela vem vazia.
        revenda_arquivo = normalizar_revenda_hierarquia(arquivo.stem.split("-")[0].strip())
        if "revenda" not in df.columns:
            df["revenda"] = revenda_arquivo
        else:
            rev_vazia = df["revenda"].isna() | (df["revenda"].astype(str).str.strip() == "")
            df.loc[rev_vazia, "revenda"] = revenda_arquivo

        # Correção: arquivo da Lider está com o nome MAGAZAN e a coluna REVENDA também como MAGAZAN.
        # Força a revenda como Líder para manter consistência.
        rev_arquivo_norm = (
            revenda_arquivo.upper()
            .replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
            .replace("Ã", "A").replace("Ç", "C")
        )
        if revenda_arquivo and rev_arquivo_norm in ("LIDER", "MAGAZAN"):
            df["revenda"] = "Líder"

        df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
        # Normaliza colunas de flag SIM/NÃO para maiúsculo sem acento
        for col in ["vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper().str.replace("Ã", "A")

        dfs.append(df)
        logger.info(f"Hierarquia {arquivo.name}: {len(df)} registros")

    if not dfs:
        logger.warning("Nenhum arquivo de hierarquia válido encontrado.")
        return None, None

    df_hier = pd.concat(dfs, ignore_index=True)

    # Normaliza nomes de revenda individualmente
    df_hier["revenda"] = df_hier["revenda"].apply(normalizar_revenda_hierarquia)

    # NOTA: agrupamento por CNPJ raiz foi removido porque, com a hierarquia como
    # base do indicador, a revenda deve ser a informada pela própria hierarquia.
    # O agrupamento estava causando dispersão de CPFs entre regionais distintas.

    # Regra de validação (Andressa/Thamires, 01/07/2026): SOMENTE quem está com
    # DESLIGADO = NÃO entra na base do indicador. SIM, FÉRIAS, BENEFÍCIO e valores
    # em branco são desconsiderados — não entram no total (denominador) nem nos
    # ativos. Filtro aplicado antes da dedupe. Todos os arquivos de hierarquia
    # possuem a coluna DESLIGADO (verificado em 01/07/2026).
    # Antes de filtrar, guardamos os CPFs de FÉRIAS por revenda para exibir a
    # quantidade e o % de férias nas tabelas de cadastro.
    df_ferias = None
    if "desligado" in df_hier.columns:
        antes_desligado = len(df_hier)
        deslig_norm = df_hier["desligado"].astype(str).str.strip().str.upper()
        df_ferias = df_hier[deslig_norm == "FÉRIAS"][["cpf_limp", "revenda"]].copy()
        df_hier = df_hier[deslig_norm.isin(["NAO", "NÃO"])].copy()
        logger.info(
            f"Filtro DESLIGADO=NÃO: {antes_desligado - len(df_hier)} registros removidos "
            f"(SIM / FÉRIAS / BENEFÍCIO / em branco)"
        )

    # Se um CPF aparecer em mais de uma revenda/loja, mantem a primeira ocorrencia
    df_hier = df_hier.drop_duplicates(subset=["cpf_limp"], keep="first")

    # Consolida férias: mantém apenas CPFs que NÃO estão na base ativa (NÃO),
    # para não contar duas vezes a mesma pessoa.
    if df_ferias is not None and not df_ferias.empty:
        df_ferias = df_ferias[~df_ferias["cpf_limp"].isin(set(df_hier["cpf_limp"]))]
        df_ferias = df_ferias.drop_duplicates(subset=["cpf_limp"], keep="first")
        logger.info(f"CPFs em férias na hierarquia: {len(df_ferias)}")

    logger.info(f"Hierarquias consolidadas: {len(df_hier)} CPFs unicos")
    return df_hier, df_ferias


def carregar_cadastro_consolidado_hierarquia():
    """
    Carrega a base consolidada 'consolidado_hierarquia_com_cadastro_site.xlsx',
    considerada a fonte correta para o total de cadastros do Programa +TOP.
    Retorna DataFrame padronizado com cpf_limp, nome, revenda, status, cargo etc.
    """
    if not CONSOLIDADO_HIERARQUIA_CADASTRO_SITE_PATH.exists():
        logger.warning(f"Base consolidada não encontrada: {CONSOLIDADO_HIERARQUIA_CADASTRO_SITE_PATH}")
        return None

    try:
        df = pd.read_excel(
            CONSOLIDADO_HIERARQUIA_CADASTRO_SITE_PATH,
            sheet_name="CPFs Distintos Hierarquia",
            dtype={"CPF_limpo": str, "CPF": str},
        )
    except Exception as e:
        logger.warning(f"Não foi possível ler a base consolidada: {e}")
        return None

    df.columns = [str(c).strip() for c in df.columns]

    # Renomeia para o padrão usado no restante do script
    df = df.rename(columns={
        "CPF_limpo": "cpf_limp",
        "CPF": "cpf",
        "Nome": "nome",
        "Revenda": "revenda",
        "Status": "status",
        "Cargo": "cargo",
        "Loja": "loja",
        "CNPJ_Loja": "cnpj_loja",
        "Cod_Loja_Hierarquia": "cod_loja",
        "CNPJ_Hierarquia": "cnpj_hierarquia",
        "Vendedor_Hierarquia": "vendedor_hierarquia",
        "Gerente_Loja_Hierarquia": "gerente_loja_hierarquia",
        "Gerente_Regional_Hierarquia": "gerente_regional_hierarquia",
        "Data_Inclusao": "data_inclusao",
    })

    # Garante CPF limpo com 11 dígitos (texto)
    df["cpf_limp"] = df["cpf_limp"].astype(str).str.replace(r"[^0-9]", "", regex=True)
    # Remove CPFs vazios/inválidos
    df = df[df["cpf_limp"].str.len().isin([11])].copy()

    # Normaliza revenda
    df["revenda"] = df["revenda"].apply(normalizar_revenda_hierarquia)
    df["revenda"] = df["revenda"].apply(nome_revenda_exibicao)

    # Mapeamento revenda -> regional a partir do cadastro base (fallback inclui CDA/Casas da Agua)
    cadastro_path = DATA_DIR / "cadastro.xlsx"
    mapa_regional = {}
    if cadastro_path.exists():
        try:
            xl_cad = pd.ExcelFile(cadastro_path)
            df_cad_orig = pd.read_excel(cadastro_path, sheet_name=xl_cad.sheet_names[0])
            df_cad_orig.columns = [c.strip().lower() for c in df_cad_orig.columns]
            if "grupo" in df_cad_orig.columns and "regional" in df_cad_orig.columns:
                for _, row in df_cad_orig.dropna(subset=["grupo", "regional"]).iterrows():
                    rev = str(row["grupo"]).strip()
                    reg = str(row["regional"]).strip()
                    if rev and reg and rev.lower() != "nan" and reg.lower() != "nan":
                        mapa_regional[rev] = reg
                # Adiciona mapeamento para Casas da Água via CDA
                if "CDA" in mapa_regional and "Casas da Água" not in mapa_regional:
                    mapa_regional["Casas da Água"] = mapa_regional["CDA"]
        except Exception as e:
            logger.warning(f"Não foi possível carregar mapeamento revenda->regional: {e}")

    # Aplica mapeamento considerando normalização
    df["regional"] = df["revenda"].apply(lambda r: regional_por_revenda(r, mapa_regional))
    df["regional_curta"] = df["regional"].apply(regional_curta)

    logger.info(f"Base consolidada hierarquia+cadastro: {len(df):,} registros, {df['cpf_limp'].nunique():,} CPFs únicos")
    return df


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


def regional_por_revenda(revenda, mapa_regional):
    """Busca regional no mapa de forma case-insensitiva e tratando CDA -> Casas da Água."""
    if pd.isna(revenda):
        return None
    rev = str(revenda).strip()
    if not rev or rev.lower() == "nan":
        return None
    if rev in mapa_regional:
        return mapa_regional[rev]
    rev_upper = rev.upper()
    for k, v in mapa_regional.items():
        if str(k).strip().upper() == rev_upper:
            return v
    # Fallback para Casas da Água via CDA (com ou sem acento)
    rev_lower = rev.lower().replace("á", "a").replace("ã", "a").replace("ç", "c")
    if rev_lower == "casas da agua" and "CDA" in mapa_regional:
        return mapa_regional["CDA"]
    return None


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
    mes_nome_completo = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARCO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
        7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
    }[mes_int]

    candidatos = [
        f"{mes_nome}_{ano}",
        f"{mes_nome.lower()}_{ano}",
        f"{mes_nome_completo}_{ano}",
        f"{mes_nome_completo.lower()}_{ano}",
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


def identificar_destaque(df, coluna_metrica, coluna_nome, maior_melhor=True, min_base=0, coluna_base=None, meta=None):
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
        return f'<span style="font-size:{tamanho}px; font-family:Arial;">🟢</span>'
    elif valor >= meta * 0.7:
        return f'<span style="font-size:{tamanho}px; font-family:Arial;">🟡</span>'
    else:
        return f'<span style="font-size:{tamanho}px; font-family:Arial;">🔴</span>'


def html_destaque(titulo, nome, regional, valor, sufixo="%", meta=None):
    """Gera conteúdo HTML de parabenização com farol e ordem regional -> revenda."""
    farol = farol_html(valor, meta, tamanho=18) if meta is not None else ""
    return (
        f"<div style='font-size:18px; font-weight:bold; margin-bottom:4px; font-family:Arial;'>"
        f"{farol} <strong>Parabéns.</strong></div>"
        f"<div style='font-size:15px; font-weight:bold; margin-bottom:10px; font-family:Arial;'>{titulo}</div>"
        f"<div style='font-family:Arial;'>A regional <strong>{regional}</strong> se destaca com a revenda <strong>{nome.upper()}</strong> "
        f"com <strong>{valor:.1f}{sufixo}</strong>.</div>"
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
        return "<p style='font-family:Arial;'><em>Sem dados para exibir.</em></p>"

    html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial; font-size: 13px;">\n'
    html += "<thead><tr>"
    for col in df.columns:
        html += (
            f'<th bgcolor="#ef4e22" style="border: 1px solid #cccccc; padding: 8px; background-color: #ef4e22; '
            f'color: white; text-align: center; font-family: Arial;">{col}</th>'
        )
    html += "</tr></thead><tbody>\n"

    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = row[col]
            is_num = pd.api.types.is_number(val) and not pd.isna(val)
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

            # Posiciona farol à direita do valor
            display = f"{display}{farol_celula}" if farol_celula else display

            html += f'<td style="border: 1px solid #cccccc; padding: 6px; text-align: {align}; font-family: Arial; {bg}">{display}</td>'
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
    df_cad_orig = pd.read_excel(cadastro_path, sheet_name=xl_cad.sheet_names[0])
    df_cad_orig.columns = [c.strip().lower() for c in df_cad_orig.columns]
    df_cad_orig["cpf_limp"] = df_cad_orig["cpf/cnpj"].apply(limpar_cpf)
    df_cad_orig["revenda_original"] = df_cad_orig["grupo"].astype(str).str.strip()
    df_cad_orig["regional_original"] = df_cad_orig["regional"].astype(str).str.strip()

    # Mapeamento revenda -> regional a partir do cadastro base
    mapa_regional = mapeamento_revenda_regional(df_cad_orig)
    logger.info(f"Mapeamento revenda->regional: {len(mapa_regional)} revendas")

    # ------------------------------------------------------------------
    # Hierarquias individuais (usadas para o detalhamento / aba de lojas)
    # ------------------------------------------------------------------
    df_hier = carregar_hierarquias()
    df_ferias_hier = None
    if isinstance(df_hier, tuple):
        df_hier, df_ferias_hier = df_hier

    # ------------------------------------------------------------------
    # Base consolidada DESATIVADA. cadastro.xlsx é a base principal e as
    # hierarquias da pasta hierarquia_rodrigo apenas enriquecem os dados.
    # ------------------------------------------------------------------
    df_cad_cons = None

    if df_cad_cons is not None:
        # Usa o consolidado como base principal de cadastros
        df_cad = df_cad_cons.copy()

        # Cruza com cadastro base para complementar dados (cidade, uf, bairro, telefone, email etc.)
        # sem sobrescrever as informações do consolidado, que é a fonte correta.
        cols_complementares = ["cpf_limp", "nome", "cargo", "status", "cidade", "uf", "bairro", "rua", "cep", "telefone", "celular", "email", "data de aceite", "lgpd"]
        cols_existentes = [c for c in cols_complementares if c in df_cad_orig.columns]
        df_comp = df_cad_orig[cols_existentes].drop_duplicates(subset=["cpf_limp"], keep="first")

        # Renomeia colunas que já existem no consolidado para sufixo _cadastro_base
        colunas_sobreposicao = [c for c in cols_existentes if c in df_cad.columns and c != "cpf_limp"]
        if colunas_sobreposicao:
            df_comp = df_comp.rename(columns={c: f"{c}_cadastro_base" for c in colunas_sobreposicao})
        df_cad = df_cad.merge(df_comp, on="cpf_limp", how="left")

        # Preenche campos vazios do consolidado com dados do cadastro base, mas status/nome/cargo do consolidado prevalecem
        for col in colunas_sobreposicao:
            if col in ["status", "nome", "cargo"]:
                # Consolidado é a fonte correta; cadastro base só preenche vazios
                df_cad[col] = df_cad[col].combine_first(df_cad.get(f"{col}_cadastro_base"))
            else:
                # Dados complementares: cadastro base preenche vazios do consolidado
                df_cad[col] = df_cad.get(f"{col}_cadastro_base").combine_first(df_cad[col])
            df_cad = df_cad.drop(columns=[f"{col}_cadastro_base"])

        # Garante colunas esperadas pelo restante do fluxo
        if "nome" not in df_cad.columns:
            df_cad["nome"] = None
        if "cargo" not in df_cad.columns:
            df_cad["cargo"] = None
        if "cidade" not in df_cad.columns:
            df_cad["cidade"] = None
        if "uf" not in df_cad.columns:
            df_cad["uf"] = None
        if "bairro" not in df_cad.columns:
            df_cad["bairro"] = None

        # Força regional a partir do mapeamento do cadastro base (o consolidado pode não ter)
        df_cad["regional"] = df_cad["revenda"].apply(lambda r: regional_por_revenda(r, mapa_regional))
        df_cad["regional_curta"] = df_cad["regional"].apply(regional_curta)

        logger.info(f"CPFs da base consolidada com regional mapeada: {df_cad['regional'].notna().sum():,}")
    else:
        # ------------------------------------------------------------------
        # Fallback: hierarquias individuais + cadastro base
        # ------------------------------------------------------------------
        df_cad = df_cad_orig.copy()
        if df_hier is not None:
            cols_hier = ["cpf_limp", "revenda", "cod_loja", "cnpj", "nome_hier", "cargo_hier",
                         "vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado"]
            df_hier = df_hier[[c for c in cols_hier if c in df_hier.columns]].copy()
            df_cad = df_cad.merge(df_hier, on="cpf_limp", how="left")
            # Revenda da hierarquia (normalizada) — usada apenas como fallback
            df_cad["revenda_hier_norm"] = df_cad["revenda"].apply(normalizar_revenda_hierarquia)
            # Cadastro é a base principal: a revenda vem do GRUPO do cadastro.
            # A hierarquia só preenche a revenda quando o grupo do cadastro estiver vazio.
            rev_cadastro = df_cad["revenda_original"].apply(normalizar_revenda)
            rev_cad_vazia = rev_cadastro.isna() | rev_cadastro.astype(str).str.strip().str.lower().isin(["", "nan"])
            df_cad["revenda"] = rev_cadastro.where(~rev_cad_vazia, df_cad["revenda_hier_norm"])
            # Aplica nome amigavel de exibicao
            df_cad["revenda"] = df_cad["revenda"].apply(nome_revenda_exibicao)
            # Regional pela revenda do cadastro, via mapeamento; fallback para original
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

    # Reconstrói o mapeamento CPF -> status APÓS todas as regras de cadastro
    # (regra dos 90 dias, exclusão de revendas e status operacionais excluídos).
    # Isso garante alinhamento com o painel ranking mensal.
    status_por_cpf_completo = (
        df_cad.dropna(subset=["cpf_limp", "status"])
        .drop_duplicates(subset=["cpf_limp"], keep="first")
        .set_index("cpf_limp")["status"]
        .to_dict()
    )
    logger.info(f"Status completo atualizado: {len(status_por_cpf_completo):,} CPFs")

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
    # Mês de referência = último mês com conclusões obrigatórias na base de
    # treinamentos (regra do projeto: usar o último mês disponível, não o mês do
    # calendário). Enquanto julho/2026 não for carregado, segue em junho/2026.
    mes_calendario = date.today().strftime("%Y-%m")
    ano_mes_hoje = mes_calendario
    try:
        obr_concl = df_trein[
            (df_trein["Estado"].astype(str).str.lower() == "concluido")
            & (df_trein["Trilha"].astype(str).str.contains("OBRIGAT", case=False, na=False))
            & (df_trein["Conclusão"].notna())
        ]
        if not obr_concl.empty:
            ano_mes_hoje = str(obr_concl["Conclusão"].dt.to_period("M").max())
            if ano_mes_hoje != mes_calendario:
                logger.info(
                    f"Mês de referência ajustado para {ano_mes_hoje} "
                    f"(último mês com treinamentos na base; calendário={mes_calendario})"
                )
    except Exception as e:
        logger.warning(f"Não foi possível derivar o mês de referência da base de treinamentos: {e}")
    aba_aceite, usou_ultima = descobrir_aba_aceites(xl_aceite, ano_mes_hoje)
    df_aceite = pd.read_excel(aceite_path, sheet_name=aba_aceite)
    df_aceite["cpf_limp"] = df_aceite["CPF"].apply(limpar_cpf)
    df_aceite["DataAceite"] = pd.to_datetime(df_aceite["DataAceite"], errors="coerce")
    df_aceite["mes_aceite"] = df_aceite["DataAceite"].dt.to_period("M")
    logger.info(f"Aceites: aba '{aba_aceite}' ({len(df_aceite):,} registros)")

    # ------------------------------------------------------------------
    # Detalhamento (consolidado das hierarquias ativas + férias como inativos)
    # ------------------------------------------------------------------
    df_det = None
    if df_hier is not None:
        # Mantém todos os CPFs da hierarquia ativa (DESLIGADO=NÃO), mesmo que
        # ainda não existam no cadastro da plataforma. Isso garante que a aba
        # Detalhamento reflita exatamente a base usada nos cálculos do e-mail.
        df_det = df_hier.copy()
        df_det["base_calculo_geral"] = True
        df_det["em_ferias"] = False

        # Adiciona CPFs em férias como inativos na base de detalhamento
        if df_ferias_hier is not None and not df_ferias_hier.empty:
            df_ferias_det = df_ferias_hier.copy()
            df_ferias_det["base_calculo_geral"] = True
            df_ferias_det["em_ferias"] = True
            # Garante as mesmas colunas básicas para concatenação
            for col in df_det.columns:
                if col not in df_ferias_det.columns:
                    df_ferias_det[col] = None
            df_det = pd.concat([df_det, df_ferias_det[df_det.columns]], ignore_index=True)

        # Cruza com cadastro para nome, cidade, uf, bairro, status e regional
        df_det = df_det.merge(
            df_cad[["cpf_limp", "nome", "cargo", "status", "cidade", "uf", "bairro", "regional_curta"]].drop_duplicates("cpf_limp"),
            on="cpf_limp",
            how="left",
        )

        # Traz o nome real da loja/filial a partir da base de cadastro x revenda x loja
        if LOJA_POR_CPF_PATH.exists():
            try:
                df_loja_real = pd.read_excel(LOJA_POR_CPF_PATH)
                df_loja_real.columns = [c.strip() for c in df_loja_real.columns]
                if "Cpf" in df_loja_real.columns and "Unnamed: 14" in df_loja_real.columns:
                    df_loja_real["cpf_limp"] = df_loja_real["Cpf"].apply(limpar_cpf)
                    df_loja_real = df_loja_real.rename(columns={"Unnamed: 14": "nome_loja_real"})
                    df_loja_real = (
                        df_loja_real.dropna(subset=["cpf_limp", "nome_loja_real"])
                        .drop_duplicates(subset=["cpf_limp"], keep="first")
                        [["cpf_limp", "nome_loja_real"]]
                    )
                    df_det = df_det.merge(df_loja_real, on="cpf_limp", how="left")
                    logger.info(
                        f"Nome real da loja mapeado para {df_det['nome_loja_real'].notna().sum():,} "
                        f"de {df_det['cpf_limp'].nunique():,} CPFs do detalhamento"
                    )
            except Exception as e:
                logger.warning(f"Não foi possível carregar nome real da loja: {e}")

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

    # ------------------------------------------------------------------
    # Últimas datas disponíveis em cada base (para exibição no e-mail/Excel)
    # ------------------------------------------------------------------
    data_ultimo_cadastro = None
    if "data de inclusão" in df_cad.columns:
        data_ultimo_cadastro = pd.to_datetime(df_cad["data de inclusão"], errors="coerce").max()

    data_ultimo_treinamento = None
    if "Conclusão" in df_trein.columns:
        df_trein_obr = df_trein[
            (df_trein["Estado"].astype(str).str.lower() == "concluido")
            & (df_trein["Trilha"].astype(str).str.contains("OBRIGAT", case=False, na=False))
            & (df_trein["Conclusão"].notna())
        ]
        if not df_trein_obr.empty:
            data_ultimo_treinamento = df_trein_obr["Conclusão"].max()

    data_ultimo_aceite = None
    if "DataAceite" in df_aceite.columns:
        data_ultimo_aceite = pd.to_datetime(df_aceite["DataAceite"], errors="coerce").max()

    return {
        "cadastro": df_cad,
        "treinamentos": df_trein,
        "aceites": df_aceite,
        "detalhamento": df_det,
        "hierarquia": df_hier,
        "ferias_hier": df_ferias_hier,
        "aba_aceite": aba_aceite,
        "usou_ultima_aba": usou_ultima,
        "mes_referencia": ano_mes_hoje,
        "emails_regionais": df_emails_reg,
        "status_completo": status_por_cpf_completo,
        "datas_ultimas": {
            "cadastro": data_ultimo_cadastro,
            "treinamento": data_ultimo_treinamento,
            "aceite": data_ultimo_aceite,
        },
    }


# ---------------------------------------------------------------------------
# CÁLCULOS POR INDICADOR
# ---------------------------------------------------------------------------
def enriquecer_hierarquia_com_regional(df_hier, df_cad):
    """
    Retorna DataFrame da hierarquia enriquecido com regional_curta.
    A regional é determinada pela revenda informada na própria hierarquia,
    usando a regional do cadastro como referência. Isso garante que todos os
    CPFs de uma mesma revenda fiquem na mesma regional, mesmo que no cadastro
    individual algum CPF esteja associado a outra revenda/regional.
    """
    if df_hier is None or df_hier.empty:
        return None

    cad_regional = df_cad[["cpf_limp", "regional_curta"]].drop_duplicates(subset=["cpf_limp"], keep="first")
    hier = df_hier[["cpf_limp", "revenda"]].drop_duplicates().copy()
    hier = hier.merge(cad_regional, on="cpf_limp", how="left")

    # Mapeamento revenda da hierarquia -> regional mais frequente no cadastro
    revenda_para_regional = (
        hier.dropna(subset=["regional_curta"])
        .groupby("revenda")["regional_curta"]
        .agg(lambda x: x.value_counts().index[0])
        .to_dict()
    )

    # Usa a regional da revenda da hierarquia para todos os CPFs da revenda
    hier["regional_curta"] = hier["revenda"].map(revenda_para_regional)

    return hier[hier["regional_curta"].notna()].copy()


def calcular_base_hierarquia(df_hier, df_cad):
    """
    Calcula o total de CPFs na hierarquia por regional e por revenda.
    Usado como denominador para os indicadores de aderência/cobertura.
    """
    hier = enriquecer_hierarquia_com_regional(df_hier, df_cad)
    if hier is None or hier.empty:
        return None, None

    hier_reg = (
        hier.groupby("regional_curta")["cpf_limp"]
        .nunique()
        .reset_index(name="total_hier")
    )
    hier_rev = (
        hier.groupby(["regional_curta", "revenda"])["cpf_limp"]
        .nunique()
        .reset_index(name="total_hier")
    )

    return hier_reg, hier_rev


def calcular_cadastros(df_cad, df_hier=None, df_ferias=None, status_completo=None):
    """
    Calcula cadastros por regional e por revenda.
    Quando a hierarquia está disponível, o denominador passa a ser o total de
    CPFs enviados pela revenda na hierarquia, refletindo a cobertura da base.
    CPFs marcados como DESLIGADO = FÉRIAS na hierarquia são incluídos na base e
    considerados inativos, independentemente do status na plataforma.

    Separação de status:
      - Ativos no +TOP: status "Ativo" no cadastro e NÃO em férias
      - Pré-Cadastro: status que não é Ativo, Inativo nem Bloqueado, e NÃO em férias
      - Inativos: status "Inativo" ou "Bloqueado" no cadastro, OU CPF em férias
    """
    hier_reg, hier_rev = calcular_base_hierarquia(df_hier, df_cad)

    if hier_reg is not None and hier_rev is not None:
        # Enriquece hierarquia com regional (cadastro + inferência por revenda)
        hier = enriquecer_hierarquia_com_regional(df_hier, df_cad)

        # Adiciona CPFs em férias à base, marcando-os como inativos
        if df_ferias is not None and not df_ferias.empty:
            ferias = df_ferias[["cpf_limp", "revenda"]].drop_duplicates().copy()
            ferias["em_ferias"] = True
            hier = pd.concat([hier, ferias], ignore_index=True)
            hier = hier.drop_duplicates(subset=["cpf_limp"], keep="first")

        hier["em_ferias"] = hier["em_ferias"].fillna(False)

        # Total da hierarquia; ativos = CPFs da hierarquia ativos na plataforma
        def _tipo_status(row):
            if row["em_ferias"]:
                return "inativo"
            status = status_completo.get(row["cpf_limp"]) if status_completo else None
            if status == "Ativo":
                return "ativo"
            if status in {"Inativo", "Bloqueado"}:
                return "inativo"
            return "pre_cadastro"

        hier["tipo_status"] = hier.apply(_tipo_status, axis=1)

        def _agg_hier(grupo_df):
            total = grupo_df["cpf_limp"].nunique()
            ativos = (grupo_df["tipo_status"] == "ativo").sum()
            inativos = (grupo_df["tipo_status"] == "inativo").sum()
            pre_cadastro = total - ativos - inativos
            return pd.Series({
                "total": total,
                "ativos": ativos,
                "pre_cadastro": pre_cadastro,
                "inativos": inativos,
                "pct_ativos": round(ativos / total * 100, 1) if total else 0,
            })

        # Por revenda (mantém nome da hierarquia)
        cad_rev = hier.groupby(["regional_curta", "revenda"]).apply(_agg_hier).reset_index()

        # Aplica nomes amigáveis de exibição para revendas
        cad_rev["revenda"] = cad_rev["revenda"].apply(nome_revenda_exibicao)

        cad_rev = cad_rev.sort_values("pct_ativos", ascending=False)
        cad_rev["regional_curta"] = cad_rev["regional_curta"].apply(regional_title_case)

        # Por regional (soma as revendas)
        cad_reg = cad_rev.groupby("regional_curta").agg(
            total=("total", "sum"),
            ativos=("ativos", "sum"),
            pre_cadastro=("pre_cadastro", "sum"),
            inativos=("inativos", "sum"),
        ).reset_index()
        cad_reg["pct_ativos"] = (cad_reg["ativos"] / cad_reg["total"] * 100).round(1)
        cad_reg = cad_reg.sort_values("pct_ativos", ascending=False)
        cad_reg["regional_curta"] = cad_reg["regional_curta"].apply(regional_title_case)

        # Ordena colunas
        cad_rev = cad_rev[["regional_curta", "revenda", "total", "ativos",
                           "pre_cadastro", "inativos", "pct_ativos"]]
        cad_reg = cad_reg[["regional_curta", "total", "ativos",
                           "pre_cadastro", "inativos", "pct_ativos"]]

        return cad_reg, cad_rev

    # Fallback: comportamento antigo (base = cadastros da plataforma)
    cad_reg = (
        df_cad.groupby("regional_curta")
        .agg(
            total=("cpf_limp", "nunique"),
            ativos=("status", lambda x: (x == "Ativo").sum()),
            pre_cadastro=("status", lambda x: (x != "Ativo").sum()),
        )
        .reset_index()
    )
    cad_reg["inativos"] = 0
    cad_reg["pct_ativos"] = (cad_reg["ativos"] / cad_reg["total"] * 100).round(1)
    cad_reg = cad_reg.sort_values("pct_ativos", ascending=False)
    cad_reg["regional_curta"] = cad_reg["regional_curta"].apply(regional_title_case)

    cad_rev = (
        df_cad.groupby(["regional_curta", "revenda"])
        .agg(
            total=("cpf_limp", "nunique"),
            ativos=("status", lambda x: (x == "Ativo").sum()),
            pre_cadastro=("status", lambda x: (x != "Ativo").sum()),
        )
        .reset_index()
    )
    cad_rev["inativos"] = 0
    cad_rev["pct_ativos"] = (cad_rev["ativos"] / cad_rev["total"] * 100).round(1)
    cad_rev = cad_rev.sort_values("pct_ativos", ascending=False)
    cad_rev["regional_curta"] = cad_rev["regional_curta"].apply(regional_title_case)

    # Sem hierarquia não há informação de férias — colunas zeradas por consistência
    for _df in (cad_reg, cad_rev):
        _df["ferias"] = 0
        _df["pct_ferias"] = 0.0

    return cad_reg, cad_rev


def calcular_treinamentos(df_trein, df_cad, ano_mes, df_hier=None):
    """
    Calcula treinamentos por regional e por revenda.
    Quando a hierarquia está disponível, a base de cálculo passa a ser o total
    de CPFs enviados pela revenda, medindo a cobertura dos treinamentos na
    hierarquia. Caso contrário, mantém o comportamento anterior (base = ativos).
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

    # Define a base de cálculo: hierarquia quando disponível, senão ativos
    if df_hier is not None and not df_hier.empty:
        base = enriquecer_hierarquia_com_regional(df_hier, df_cad)
        logger.info(f"Treinamentos usarão hierarquia como base: {base['cpf_limp'].nunique():,} CPFs")
    else:
        base = df_cad[df_cad["status"] == "Ativo"][["cpf_limp", "revenda", "regional_curta"]].drop_duplicates().copy()
        logger.info("Treinamentos usarão base ativa da plataforma (hierarquia indisponível)")

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
    trein_reg = base.groupby("regional_curta").apply(lambda g: calcular_grupo(g, cpf_ambos)).reset_index()
    trein_reg = trein_reg.sort_values("pct_realizaram", ascending=False)
    trein_reg["regional_curta"] = trein_reg["regional_curta"].apply(regional_title_case)

    trein_rev = base.groupby(["regional_curta", "revenda"]).apply(lambda g: calcular_grupo(g, cpf_ambos)).reset_index()
    trein_rev = trein_rev.sort_values("pct_realizaram", ascending=False)
    trein_rev["regional_curta"] = trein_rev["regional_curta"].apply(regional_title_case)
    trein_rev["revenda"] = trein_rev["revenda"].apply(nome_revenda_exibicao)

    # Por curso individual
    def calcular_por_curso(cpf_curso):
        reg = base.groupby("regional_curta").apply(lambda g: calcular_grupo(g, cpf_curso)).reset_index()
        reg = reg.sort_values("pct_realizaram", ascending=False)
        reg["regional_curta"] = reg["regional_curta"].apply(regional_title_case)
        rev = base.groupby(["regional_curta", "revenda"]).apply(lambda g: calcular_grupo(g, cpf_curso)).reset_index()
        rev = rev.sort_values("pct_realizaram", ascending=False)
        rev["regional_curta"] = rev["regional_curta"].apply(regional_title_case)
        rev["revenda"] = rev["revenda"].apply(nome_revenda_exibicao)
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

    return trein_reg, trein_rev, trein_por_curso, (curso1, curso2, nome1, nome2, sku1, sku2), base


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


def calcular_aceites(df_aceite, df_cad, ano_mes, aba_aceite, usou_ultima_aba=False, df_hier=None):
    """
    Calcula aceites mensais por regional e por revenda.
    Quando a hierarquia está disponível, a base de cálculo passa a ser o total
    de CPFs enviados pela revenda na hierarquia, medindo a cobertura de aceites.
    Caso contrário, mantém o comportamento anterior (base = ativos).
    """
    if usou_ultima_aba:
        mes_ref = inferir_mes_da_aba(aba_aceite)
        if mes_ref is None:
            mes_ref = df_aceite["mes_aceite"].max()
    else:
        mes_ref = pd.Period(ano_mes, freq="M")

    aceite_mes = df_aceite[df_aceite["mes_aceite"] == mes_ref].copy()
    cpfs_aceitaram = set(aceite_mes["cpf_limp"].unique())
    logger.info(f"Aceites em {mes_ref}: {len(cpfs_aceitaram):,} CPFs únicos")

    # Define a base de cálculo: hierarquia quando disponível, senão ativos
    if df_hier is not None and not df_hier.empty:
        base = enriquecer_hierarquia_com_regional(df_hier, df_cad)
        logger.info(f"Aceites usarão hierarquia como base: {base['cpf_limp'].nunique():,} CPFs")
    else:
        base = df_cad[df_cad["status"] == "Ativo"][["cpf_limp", "revenda", "regional_curta"]].drop_duplicates().copy()
        logger.info("Aceites usarão base ativa da plataforma (hierarquia indisponível)")

    # Aceitaram: cruza base com aceites
    base["aceitou"] = base["cpf_limp"].isin(cpfs_aceitaram)

    # Regional
    aceite_reg = (
        base.groupby("regional_curta")
        .agg(
            total_ativos=("cpf_limp", "nunique"),
            aceitaram=("aceitou", "sum"),
        )
        .reset_index()
    )
    aceite_reg["aceitaram"] = aceite_reg["aceitaram"].astype(int)
    aceite_reg["nao_aceitaram"] = (aceite_reg["total_ativos"] - aceite_reg["aceitaram"]).clip(lower=0)
    aceite_reg["pct_aceite"] = (aceite_reg["aceitaram"] / aceite_reg["total_ativos"] * 100).round(1)
    aceite_reg = aceite_reg.sort_values("pct_aceite", ascending=False)
    aceite_reg = aceite_reg.rename(columns={"regional_curta": "regional"})
    aceite_reg["regional"] = aceite_reg["regional"].apply(regional_title_case)

    # Revenda
    aceite_rev = (
        base.groupby(["regional_curta", "revenda"])
        .agg(
            total_ativos=("cpf_limp", "nunique"),
            aceitaram=("aceitou", "sum"),
        )
        .reset_index()
    )
    aceite_rev["aceitaram"] = aceite_rev["aceitaram"].astype(int)
    aceite_rev["nao_aceitaram"] = (aceite_rev["total_ativos"] - aceite_rev["aceitaram"]).clip(lower=0)
    aceite_rev["pct_aceite"] = aceite_rev.apply(
        lambda r: round(r["aceitaram"] / r["total_ativos"] * 100, 1) if r["total_ativos"] > 0 else 0,
        axis=1,
    )
    aceite_rev = aceite_rev.sort_values("pct_aceite", ascending=False)
    aceite_rev["regional"] = aceite_rev["regional_curta"].apply(regional_title_case)
    aceite_rev["revenda"] = aceite_rev["revenda"].apply(nome_revenda_exibicao)

    return aceite_reg, aceite_rev, mes_ref, base


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

    # Garante fonte Arial em todos os elementos do gráfico
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Arial"]

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
        media_cad = round(cad_reg["ativos"].sum() / cad_reg["total"].sum() * 100, 1)
        fig = gerar_grafico_barras(
            cad_reg,
            "regional_curta", "pct_ativos",
            f"Cadastros Ativos por Regional (média geral: {media_cad:.1f}%)",
            meta=META_CADASTRO,
        )
        if fig:
            graficos["cadastros"] = fig_to_base64(fig)

    if trein_reg is not None and not trein_reg.empty:
        media_trein = round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1)
        fig = gerar_grafico_barras(
            trein_reg,
            "regional_curta", "pct_realizaram",
            f"Treinamentos Obrigatórios Realizados por Regional (média geral: {media_trein:.1f}%)",
            meta=META_TREINAMENTOS,
        )
        if fig:
            graficos["treinamentos"] = fig_to_base64(fig)

    if not aceite_reg.empty:
        media_aceite = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1)
        fig = gerar_grafico_barras(
            aceite_reg,
            "regional", "pct_aceite",
            f"Aceite Mensal por Regional (média geral: {media_aceite:.1f}%)",
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
            f"{farol_geral} Nosso objetivo para a cobertura de cadastros do Programa +TOP é de "
            f"<strong>{META_CADASTRO:.0f}%</strong>.<br>"
            f"Até o momento, <strong>{pct_geral}%</strong> da hierarquia enviada pelas revendas está ativa na plataforma."
        )

        piores_cad = garantir_n_itens(
            cad_rev, "pct_ativos",
            mascara_filtro=None,
            limite=10, maior_melhor=False
        )
        if not piores_cad.empty:
            itens_df = piores_cad[["regional_curta", "revenda", "total", "ativos", "pct_ativos"]].copy()
            itens_df["revenda"] = itens_df["revenda"].str.title()
            itens_df = itens_df.rename(columns={
                "regional_curta": "Regional",
                "revenda": "Revenda",
                "total": "Total de participantes",
                "ativos": "Ativos no +TOP",
                "pct_ativos": "% Ativos",
            })
            resultado["cadastros"]["alerta_titulo"] = "Top 10 revendas com menor % de ativos"
            resultado["cadastros"]["alerta_subtitulo"] = ""
            resultado["cadastros"]["alerta_itens"] = itens_df

        # Apenas 1 destaque geral (a melhor revenda do programa, que atingiu a meta)
        dest = identificar_destaque(cad_rev, "pct_ativos", "revenda", maior_melhor=True, min_base=0, meta=META_CADASTRO)
        if dest:
            resultado["cadastros"]["destaque"] = html_destaque(
                "Melhor % de ativos.", dest["nome"], dest["regional"], dest["valor"], meta=META_CADASTRO
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
            f"{farol_trein} Os dois conteúdos de {mes_nome} ficaram disponíveis para os vendedores ate o dia "
            f"{date.today().replace(day=30):%d/%m/%Y}.<br>"
            f"Nosso objetivo para os treinamentos do Programa +TOP é de <strong>{META_TREINAMENTOS:.0f}%</strong>.<br>"
            f"Até o momento, estamos em <strong>{pct_trein_geral}%</strong> da hierarquia enviada pelas revendas.<br>"
            f"Conteúdo 1: {nome1} (SKU {sku1})<br>Conteúdo 2: {nome2} (SKU {sku2})"
        )

        piores_rev = garantir_n_itens(
            trein_rev[trein_rev["total_ativos"] > 0], "pct_realizaram",
            mascara_filtro=None,
            limite=10, maior_melhor=False
        )
        if not piores_rev.empty:
            itens_df = piores_rev[["regional_curta", "revenda", "total_ativos", "realizaram", "pct_realizaram"]].copy()
            itens_df["revenda"] = itens_df["revenda"].str.title()
            itens_df = itens_df.rename(columns={
                "regional_curta": "Regional",
                "revenda": "Revenda",
                "total_ativos": "Total de participantes",
                "realizaram": "Realizado",
                "pct_realizaram": "% Realizado",
            })
            resultado["treinamentos"]["alerta_titulo"] = "Top 10 revendas com menor % de treinamentos concluídos"
            resultado["treinamentos"]["alerta_subtitulo"] = ""
            resultado["treinamentos"]["alerta_itens"] = itens_df

        # Destaque coletivo é renderizado diretamente no template do e-mail
        # usando _destaques_meta_html; não precisa gerar texto único aqui.
        resultado["treinamentos"]["destaque"] = ""

    # --- Aceites ---
    if not aceite_reg.empty:
        media_aceite = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1)
        farol_aceite = farol_html(media_aceite, META_ACEITES)
        resultado["aceites"]["intro"] = (
            f"{farol_aceite} Sobre os aceites mensais de <strong>{mes_aceite_nome}</strong>.<br>"
            f"Nosso objetivo é de <strong>{META_ACEITES:.0f}%</strong>. "
            f"Ate o momento, <strong>{media_aceite}%</strong> da hierarquia enviada pelas revendas deu aceite na campanha base."
        )
        if aceite_rev is not None and not aceite_rev.empty:
            piores_aceite = garantir_n_itens(
                aceite_rev[aceite_rev["total_ativos"] > 0], "pct_aceite",
                mascara_filtro=None,
                limite=10, maior_melhor=False
            )
            if not piores_aceite.empty:
                itens_df = piores_aceite[["regional", "revenda", "total_ativos", "aceitaram", "pct_aceite"]].copy()
                itens_df = itens_df.rename(columns={
                    "regional": "Regional",
                    "revenda": "Revenda",
                    "total_ativos": "Total de participantes",
                    "aceitaram": "Aceitaram",
                    "pct_aceite": "% Aceite",
                })
                resultado["aceites"]["alerta_titulo"] = "Top 10 revendas com menor % de aceite"
                resultado["aceites"]["alerta_subtitulo"] = ""
                resultado["aceites"]["alerta_itens"] = itens_df

            # Destaque coletivo é renderizado diretamente no template do e-mail
            # usando _destaques_meta_html; não precisa gerar texto único aqui.
            resultado["aceites"]["destaque"] = ""

    return resultado


def gerar_insights_regional(
    cad_reg, cad_rev, trein_reg, trein_rev, aceite_reg, aceite_rev,
    evolucao, dados, regional_filtro, imagens_kv=None, imagens_tom=None
):
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
            f"{farol_reg} A regional <strong>{regional_filtro}</strong> está com <strong>{row['pct_ativos']:.1f}%</strong> de CPFs ativos na plataforma.<br>"
            f"Ela ocupa a <strong>{posição}ª posição</strong> entre {total} regionais (media geral: {media:.1f}%).<br>"
            f"Objetivo: <strong>{META_CADASTRO:.0f}%</strong>."
        )

        abaixo_media = cad_rev_f[cad_rev_f["pct_ativos"] < media].sort_values("pct_ativos", ascending=False).head(10)
        if not abaixo_media.empty:
            linhas = [f"{r['revenda'].upper()} {r['pct_ativos']:.0f}%" for _, r in abaixo_media.iterrows()]
            resultado["cadastros"]["alerta"] = (
                "Revendas da regional abaixo da média geral do programa. Precisam de reforço.<br>" + "<br>".join(linhas)
            )

        resultado["cadastros"]["destaque"] = _destaques_meta_html(
            cad_rev_f,
            meta=META_CADASTRO,
            col_pct="pct_ativos",
            col_regional="regional_curta",
            col_revenda="revenda",
            titulo="Parabéns!",
            subtitulo="Melhor % de ativos.",
            imagens_kv=imagens_kv,
            imagens_tom=imagens_tom,
            col_num="ativos",
            col_den="total",
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
            f"{farol_trein_reg} Treinamentos de {mes_nome}: <strong>{r['pct_realizaram']:.1f}%</strong> da hierarquia da regional concluiu ambos os cursos.<br>"
            f"Média geral do programa: {media_trein:.1f}%. Objetivo: <strong>{META_TREINAMENTOS:.0f}%</strong>.<br>"
            f"Conteúdo 1: {nome1} (SKU {sku1})<br>Conteúdo 2: {nome2} (SKU {sku2})"
        )

        piores_rev = trein_rev_f[trein_rev_f["pct_realizaram"] < media_trein].sort_values("pct_realizaram", ascending=False).head(10)
        if not piores_rev.empty:
            linhas = [f"{r['revenda'].upper()} {r['pct_realizaram']:.0f}%" for _, r in piores_rev.iterrows()]
            resultado["treinamentos"]["alerta"] = (
                "Revendas da regional com treinamentos abaixo da media.<br>" + "<br>".join(linhas)
            )

        resultado["treinamentos"]["destaque"] = _destaques_meta_html(
            trein_rev_f,
            meta=META_TREINAMENTOS,
            col_pct="pct_realizaram",
            col_regional="regional_curta",
            col_revenda="revenda",
            titulo="Parabéns!",
            subtitulo="Melhor % de treinamentos concluídos.",
            imagens_kv=imagens_kv,
            imagens_tom=imagens_tom,
            col_num="realizaram",
            col_den="total_ativos",
        )

    if not aceite_reg_f.empty:
        r = aceite_reg_f.iloc[0]
        media_aceite = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1)
        farol_aceite_reg = farol_html(r['pct_aceite'], META_ACEITES)
        resultado["aceites"]["intro"] = (
            f"{farol_aceite_reg} Aceite mensal de {mes_aceite_nome}: <strong>{r['pct_aceite']:.1f}%</strong> da hierarquia da regional deu aceite.<br>"
            f"Média geral do programa: {media_aceite:.1f}%. Objetivo: <strong>{META_ACEITES:.0f}%</strong>."
        )

        if aceite_rev_f is not None and not aceite_rev_f.empty:
            resultado["aceites"]["destaque"] = _destaques_meta_html(
                aceite_rev_f,
                meta=META_ACEITES,
                col_pct="pct_aceite",
                col_regional="regional",
                col_revenda="revenda",
                titulo="Parabéns!",
                subtitulo="Melhor % de aceite mensal.",
                imagens_kv=imagens_kv,
                imagens_tom=imagens_tom,
                col_num="aceitaram",
                col_den="total_ativos",
            )

    return resultado


# ---------------------------------------------------------------------------
# MONTAGEM DO E-MAIL
# ---------------------------------------------------------------------------
def _renomear_cadastro_reg(df):
    df = df.rename(columns={
        "regional_curta": "Regional",
        "total": "Total de participantes",
        "ativos": "Ativos no +TOP",
        "pre_cadastro": "Pré-Cadastro",
        "pct_ativos": "% Ativos",
    })
    cols = ["Regional", "Total de participantes", "Ativos no +TOP", "Pré-Cadastro", "% Ativos"]
    return df[[c for c in cols if c in df.columns]]


def _renomear_cadastro_rev(df):
    df = df.rename(columns={
        "regional_curta": "Regional",
        "revenda": "Revenda",
        "total": "Total de participantes",
        "ativos": "Ativos no +TOP",
        "pre_cadastro": "Pré-Cadastro",
        "pct_ativos": "% Ativos",
    })
    cols = ["Regional", "Revenda", "Total de participantes", "Ativos no +TOP", "Pré-Cadastro", "% Ativos"]
    return df[[c for c in cols if c in df.columns]]


def _renomear_trein_reg(df):
    return df.rename(columns={
        "regional_curta": "Regional",
        "total_ativos": "Total de participantes",
        "realizaram": "Realizado",
        "nao_realizaram": "Não Realizado",
        "pct_realizaram": "% Realizado",
    })


def _renomear_trein_rev(df):
    return df.rename(columns={
        "regional_curta": "Regional",
        "revenda": "Revenda",
        "total_ativos": "Total de participantes",
        "realizaram": "Realizado",
        "nao_realizaram": "Não Realizado",
        "pct_realizaram": "% Realizado",
    })


def _renomear_aceite_reg(df):
    return df.rename(columns={
        "regional": "Regional",
        "total_ativos": "Total de participantes",
        "aceitaram": "Aceitaram",
        "nao_aceitaram": "Não Aceitaram",
        "pct_aceite": "% Aceite",
    })


def _renomear_aceite_rev(df):
    df = df.drop(columns=["regional_curta"], errors="ignore")
    df = df.rename(columns={
        "regional": "Regional",
        "revenda": "Revenda",
        "total_ativos": "Total de participantes",
        "aceitaram": "Aceitaram",
        "nao_aceitaram": "Não Aceitaram",
        "pct_aceite": "% Aceite",
    })
    cols = ["Regional", "Revenda", "Total de participantes", "Aceitaram", "Não Aceitaram", "% Aceite"]
    return df[[c for c in cols if c in df.columns]]


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

    df = df.sort_values("% Ambos", ascending=False).fillna(0)

    # Reordena colunas: regional/revenda, curso1, curso2, % Ambos (semaforo sempre por ultimo)
    cols_exibicao = ["Regional"]
    if nivel == "revenda":
        cols_exibicao.append("Revenda")
    cols_exibicao.extend([f"% {nome1}", f"% {nome2}", "% Ambos"])

    # Renomeia chaves para exibição
    rename = {"regional_curta": "Regional"}
    if nivel == "revenda":
        rename["revenda"] = "Revenda"
    df = df.rename(columns=rename)

    return df[cols_exibicao]


def preparar_tabela_consolidada(cad_df, trein_df, aceite_df, nivel="regional"):
    """
    Consolida cadastros, treinamentos e aceites em uma única tabela.
    nivel='regional' agrupa por regional; nivel='revenda' inclui revenda.
    Usada principalmente para o Excel anexo (aba Resumo).
    """
    chaves = ["regional_curta"] if nivel == "regional" else ["regional_curta", "revenda"]

    df = cad_df.copy()

    if trein_df is not None and not trein_df.empty:
        trein_cols = [c for c in chaves + ["realizaram", "pct_realizaram"] if c in trein_df.columns]
        if "realizaram" in trein_df.columns:
            df = df.merge(trein_df[trein_cols], on=chaves, how="left")
            df = df.rename(columns={
                "realizaram": "Realizaram 2 cursos",
                "pct_realizaram": "% Realizaram 2 cursos",
            })
            df["Não realizaram 2 cursos"] = df["total"] - df["Realizaram 2 cursos"]
        else:
            df["Realizaram 2 cursos"] = 0
            df["Não realizaram 2 cursos"] = df["total"]
            df["% Realizaram 2 cursos"] = 0.0
    else:
        df["Realizaram 2 cursos"] = 0
        df["Não realizaram 2 cursos"] = df["total"]
        df["% Realizaram 2 cursos"] = 0.0

    if aceite_df is not None and not aceite_df.empty:
        aceite_chaves = ["regional"] if nivel == "regional" else ["regional", "revenda"]
        aceite_cols = [c for c in aceite_chaves + ["aceitaram", "pct_aceite"] if c in aceite_df.columns]
        if "aceitaram" in aceite_df.columns:
            df_aceite = aceite_df[aceite_cols].copy()
            df_aceite = df_aceite.rename(columns={"regional": "regional_curta"})
            df = df.merge(df_aceite, on=chaves, how="left")
            df = df.rename(columns={
                "aceitaram": "Aceitaram",
                "pct_aceite": "% Aceite",
            })
            df["Não aceitaram"] = df["total"] - df["Aceitaram"]
        else:
            df["Aceitaram"] = 0
            df["Não aceitaram"] = df["total"]
            df["% Aceite"] = 0.0
    else:
        df["Aceitaram"] = 0
        df["Não aceitaram"] = df["total"]
        df["% Aceite"] = 0.0

    for col in ["Realizaram 2 cursos", "Não realizaram 2 cursos", "Aceitaram", "Não aceitaram"]:
        df[col] = df[col].fillna(0).astype(int)
    for col in ["% Realizaram 2 cursos", "% Aceite"]:
        df[col] = df[col].fillna(0).round(1)

    rename = {
        "regional_curta": "Regional",
        "total": "Total de participantes",
        "ativos": "Ativos no +TOP",
        "pre_cadastro": "Pré-Cadastro",
        "pct_ativos": "% Ativos",
    }
    if nivel == "revenda":
        rename["revenda"] = "Revenda"

    df = df.rename(columns=rename)

    cols = ["Regional"]
    if nivel == "revenda":
        cols.append("Revenda")
    cols.extend([
        "Total de participantes", "Ativos no +TOP", "Pré-Cadastro", "% Ativos",
        "Realizaram 2 cursos", "Não realizaram 2 cursos", "% Realizaram 2 cursos",
        "Aceitaram", "Não aceitaram", "% Aceite",
    ])

    return df[cols].sort_values("% Ativos", ascending=False)


def preparar_tabela_base_treinamentos(cad_df, trein_df, nivel="regional"):
    """
    Tabela de treinamentos mostrando a base da hierarquia e quem realizou os 2 cursos.
    """
    chaves = ["regional_curta"] if nivel == "regional" else ["regional_curta", "revenda"]
    df = cad_df[chaves + ["total"]].copy()

    if trein_df is not None and not trein_df.empty:
        trein_cols = chaves + ["realizaram"]
        df = df.merge(trein_df[trein_cols], on=chaves, how="left")
    else:
        df["realizaram"] = 0

    df["realizaram"] = df["realizaram"].fillna(0).astype(int)
    df["nao_realizaram"] = df["total"] - df["realizaram"]
    df["pct_realizaram"] = (df["realizaram"] / df["total"] * 100).round(1)

    rename = {
        "regional_curta": "Regional",
        "total": "Total de participantes",
        "realizaram": "Realizado",
        "nao_realizaram": "Não Realizado",
        "pct_realizaram": "% Ambos",
    }
    if nivel == "revenda":
        rename["revenda"] = "Revenda"

    df = df.rename(columns=rename)
    cols = ["Regional"]
    if nivel == "revenda":
        cols.append("Revenda")
    cols.extend([
        "Total de participantes", "Realizado",
        "Não Realizado", "% Ambos"
    ])
    return df[cols].sort_values("% Ambos", ascending=False)


def preparar_tabela_base_aceites(cad_df, aceite_df, nivel="regional"):
    """
    Tabela de aceites mostrando a base da hierarquia e quem aceitou.
    """
    chaves = ["regional_curta"] if nivel == "regional" else ["regional_curta", "revenda"]
    df = cad_df[chaves + ["total"]].copy()

    if aceite_df is not None and not aceite_df.empty:
        aceite_chaves = ["regional"] if nivel == "regional" else ["regional", "revenda"]
        aceite_cols = aceite_chaves + ["aceitaram"]
        df_aceite = aceite_df[aceite_cols].copy()
        df_aceite = df_aceite.rename(columns={"regional": "regional_curta"})
        df = df.merge(df_aceite, on=chaves, how="left")
    else:
        df["aceitaram"] = 0

    df["aceitaram"] = df["aceitaram"].fillna(0).astype(int)
    df["nao_aceitaram"] = df["total"] - df["aceitaram"]
    df["pct_aceite"] = (df["aceitaram"] / df["total"] * 100).round(1)

    rename = {
        "regional_curta": "Regional",
        "total": "Total de participantes",
        "aceitaram": "Aceitaram",
        "nao_aceitaram": "Não Aceitaram",
        "pct_aceite": "% Aceite",
    }
    if nivel == "revenda":
        rename["revenda"] = "Revenda"

    df = df.rename(columns=rename)
    cols = ["Regional"]
    if nivel == "revenda":
        cols.append("Revenda")
    cols.extend([
        "Total de participantes", "Aceitaram", "Não Aceitaram", "% Aceite"
    ])
    return df[cols].sort_values("% Aceite", ascending=False)


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
        "consolidado_reg": "",
        "consolidado_rev": "",
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

    # Tabelas consolidadas com base da hierarquia (cadastro + treinamentos + aceites)
    # A consolidada por regional usa dados de treinamento/aceite; a por revenda fica apenas com cadastro
    cons_reg = preparar_tabela_consolidada(cad_reg, trein_reg, aceite_reg, nivel="regional")
    cons_rev = preparar_tabela_consolidada(cad_rev, None, None, nivel="revenda")
    tabelas["consolidado_reg"] = estilizar_tabela_html(
        cons_reg,
        semaforo_coluna="% Ativos",
        meta_semaforo=META_CADASTRO,
    )
    tabelas["consolidado_rev"] = estilizar_tabela_html(
        cons_rev.head(10),
        semaforo_coluna="% Ativos",
        meta_semaforo=META_CADASTRO,
    )

    # Tabelas de treinamentos com base da hierarquia
    trein_base_reg = preparar_tabela_base_treinamentos(cad_reg, trein_reg, nivel="regional")
    trein_base_rev = preparar_tabela_base_treinamentos(cad_rev, trein_rev, nivel="revenda")
    tabelas["trein_base_reg"] = estilizar_tabela_html(
        trein_base_reg,
        semaforo_coluna="% Ambos",
        meta_semaforo=META_TREINAMENTOS,
    )
    tabelas["trein_base_rev"] = estilizar_tabela_html(
        trein_base_rev.head(10),
        semaforo_coluna="% Ambos",
        meta_semaforo=META_TREINAMENTOS,
    )

    # Tabelas de aceites com base da hierarquia
    aceite_base_reg = preparar_tabela_base_aceites(cad_reg, aceite_reg, nivel="regional")
    aceite_base_rev = preparar_tabela_base_aceites(cad_rev, aceite_rev, nivel="revenda")
    tabelas["aceite_base_reg"] = estilizar_tabela_html(
        aceite_base_reg,
        semaforo_coluna="% Aceite",
        meta_semaforo=META_ACEITES,
    )
    tabelas["aceite_base_rev"] = estilizar_tabela_html(
        aceite_base_rev.head(10),
        semaforo_coluna="% Aceite",
        meta_semaforo=META_ACEITES,
    )

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
    """Gera card verde de destaque sem imagem do Tom (Tom mantido apenas no cabeçalho)."""
    if not texto:
        return ""
    return f"""
    <table cellpadding="0" cellspacing="0" border="0" style="margin:12px 0; width:100%; font-family:Arial;">
      <tr>
        <td valign="middle" style="padding:6px 10px; background-color:#d4edda; border-radius:8px; border:1px solid #00a651; color:#155724; font-size:13px; line-height:1.35; font-family:Arial;">
          {texto}
        </td>
      </tr>
    </table>
    """


def _pontos_atencao_header_html():
    """Retorna cabeçalho destacado para a seção de pontos de atenção."""


def _balao_tom_html(texto, imagens_kv=None, imagens_tom=None, tipo_tom="tom", alinhamento="esquerda", largura_tom=90):
    """Cria balão de destaque sem imagem do Tom (Tom mantido apenas no cabeçalho)."""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0; font-family:Arial;">
      <tr>
        <td valign="middle" style="padding:8px 12px; background-color:#ffffff; border-radius:12px; border:2px solid #00a651; color:#333333; font-size:15px; line-height:1.4; font-family:Arial;">
          {texto}
        </td>
      </tr>
    </table>
    """


def _pontos_atencao_secao_html(titulo, subtitulo, tabela_html, imagens_kv=None, imagens_tom=None):
    """Monta bloco de pontos de atenção com título grande, subtítulo e tabela."""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px; font-family:Arial;">
      <tr>
        <td style="padding:14px; background-color:#fff3cd; border-left:5px solid #ef4e22; color:#856404; font-family:Arial;">
          <div style="font-size:17px; font-weight:bold; margin-bottom:6px; font-family:Arial;">⚠️ {titulo}</div>
          <div style="font-size:14px; margin-bottom:12px; font-family:Arial;">{subtitulo}</div>
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
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:12px 0; font-family:Arial;">
      <tr>
        <td style="padding:14px; background-color:#ffffff; border-radius:12px; border:2px solid #00a651; color:#333333; font-size:15px; line-height:1.5; font-family:Arial;">
          {texto}
        </td>
      </tr>
    </table>
    """


def _secao_html(titulo):
    """Retorna titulo de secao em tabela, compativel com Outlook."""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:28px; font-family:Arial;">
      <tr>
        <td style="color:#ef4e22; font-size:18px; font-weight:bold; padding-bottom:8px; font-family:Arial;">
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
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0; font-family:Arial;">
      <tr>
        <td width="4" bgcolor="{borda}" style="font-size:0; line-height:0;">&nbsp;</td>
        <td bgcolor="{bg}" style="padding:16px; color:#333333; font-size:16px; line-height:1.6; font-family:Arial;">
          {conteudo}
        </td>
      </tr>
    </table>
    """


def _img_html(cid, alt):
    if not cid:
        return ""
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0; font-family:Arial;">
      <tr>
        <td align="center" style="font-family:Arial;">
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
        "logo": (kv_dir / "_top_logo.png", 220),
        "tom": (kv_dir / "_top_tom_cubos1.png", 120),
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
    farol = farol_html(valor, meta, tamanho=38) if meta is not None else ""
    return f"""
    <td width="33%" align="center" valign="middle" bgcolor="#ffffff" style="padding:20px 24px; color:#333333; font-size:16px; font-weight:bold; border-radius:10px; border:3px solid #ef4e22; font-family:Arial;">
      <div style="font-size:16px; margin-bottom:6px; color:#ef4e22; font-family:Arial;">{titulo}</div>
      <div style="font-size:42px; margin-bottom:8px; color:#ef4e22; line-height:1; white-space:nowrap; font-family:Arial;">
        {farol}&nbsp;<strong>{valor}%</strong>
      </div>
    </td>
    """


def _header_html(titulo, hoje, imagens_kv=None):
    """Header com identidade visual +TOP usando imagens do KV."""
    imagens_kv = imagens_kv or {}
    logo_cid = "kv_logo" if "logo" in imagens_kv else None
    tom_cid = "kv_tom" if "tom" in imagens_kv else None
    fundo_cid = "kv_fundo" if "fundo" in imagens_kv else None

    bg_style = f'background-image: url(cid:{fundo_cid}); background-size: cover; background-position: center;' if fundo_cid else 'background-color: #f5f5f5;'

    logo_html = f'<img src="cid:{logo_cid}" alt="Logo +TOP" width="220" style="display:block;">' if logo_cid else '<span style="font-size:24px; font-weight:bold; font-family:Arial;">+top</span>'
    # TOM posicionado à direita, próximo à faixa laranja (padding-bottom reduzido)
    tom_html = f'<img src="cid:{tom_cid}" alt="Tom +TOP" width="120" style="display:block;" align="bottom">' if tom_cid else ''

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="font-family:Arial;">
      <tr>
        <td style="padding:0; font-family:Arial; {bg_style}">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:Arial;">
            <tr>
              <td style="padding:24px 24px 0 24px; font-family:Arial;" valign="top">
                {logo_html}
              </td>
              <td align="right" style="padding:8px 24px 0 24px; font-family:Arial;" valign="bottom">
                {tom_html}
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td bgcolor="#ef4e22" style="padding:14px 24px; color:#ffffff; font-family:Arial;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:Arial;">
            <tr>
              <td style="font-family:Arial;">
                <h1 style="margin:0; font-size:18px; font-weight:bold; color:#ffffff; font-family:Arial;">{titulo}</h1>
              </td>
              <td align="right" style="font-family:Arial;">
                <p style="margin:0; font-size:12px; color:#ffffff; font-family:Arial;">{hoje}</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    """


def _destaques_meta_html(
    df,
    meta,
    col_pct,
    col_regional,
    col_revenda,
    titulo,
    subtitulo,
    imagens_kv=None,
    imagens_tom=None,
    col_num=None,
    col_den=None,
):
    """Gera um único balão verde de destaque com revendas agrupadas por regional.

    Se col_num e col_den forem informados, filtra pelo cálculo exato para evitar
    que valores arredondados entrem no destaque. Caso contrário, usa col_pct.
    """
    if df is None or df.empty:
        return ""

    df = df.copy()

    # Usa o % exato quando possível, para não incluir valores que arredondariam para a meta.
    if col_num and col_den and col_num in df.columns and col_den in df.columns:
        atingiram = df[(df[col_num] / df[col_den] * 100) >= meta].copy()
    else:
        atingiram = df[df[col_pct] >= meta].copy()

    if atingiram.empty:
        return ""

    # Ordenar revendas dentro de cada regional do maior para o menor
    atingiram = atingiram.sort_values([col_regional, col_pct], ascending=[True, False])

    # Ordenar regionais pela média do indicador (maior primeiro)
    ordem_regional = (
        atingiram.groupby(col_regional)[col_pct]
        .mean()
        .sort_values(ascending=False)
        .index.tolist()
    )

    imagens_tom = imagens_tom or {}
    imagens_kv = imagens_kv or {}
    # Tom removido do corpo do e-mail; mantido apenas no cabeçalho
    tom_html = ""

    blocos_regional = []
    for regional in ordem_regional:
        revendas_reg = atingiram[atingiram[col_regional] == regional]
        linhas_rev = []
        for _, r in revendas_reg.iterrows():
            linhas_rev.append(
                f"• <strong>{r[col_revenda]}</strong> — <strong>{r[col_pct]:.1f}%</strong>"
            )

        bloco = (
            f"<div style='margin-bottom:16px; font-family:Arial;'>"
            f"<div style='font-size:16px; font-weight:bold; margin-bottom:6px; font-family:Arial;'>"
            f"🟢 Regional <strong>{regional}</strong></div>"
            f"<div style='font-size:14px; line-height:1.7; font-family:Arial;'>"
            + "<br>".join(linhas_rev)
            + "</div></div>"
        )
        blocos_regional.append(bloco)

    conteudo = (
        f"<div style='font-size:18px; font-weight:bold; margin-bottom:4px; font-family:Arial;'>"
        f"🎉 {titulo}</div>"
        f"<div style='font-size:15px; font-weight:bold; margin-bottom:12px; font-family:Arial;'>"
        f"{subtitulo}</div>"
        f"<div style='font-size:14px; margin-bottom:12px; font-family:Arial;'>"
        f"As revendas abaixo atingiram ou superaram o objetivo mínimo de <strong>{meta:.0f}%</strong>:</div>"
        + "\n".join(blocos_regional)
    )

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0; font-family:Arial;">
      <tr>
        <td valign="middle" style="padding:14px 16px; background-color:#d4edda; border-radius:12px; border:1px solid #00a651; color:#155724; font-size:14px; line-height:1.6; font-family:Arial;">
          {conteudo}
        </td>
      </tr>
    </table>
    """


def _destaques_cadastro_html(cad_rev, imagens_kv=None, imagens_tom=None):
    """Wrapper para manter o destaque de cadastros com o comportamento atual."""
    return _destaques_meta_html(
        cad_rev,
        meta=META_CADASTRO,
        col_pct="pct_ativos",
        col_regional="regional_curta",
        col_revenda="revenda",
        titulo="Parabéns!",
        subtitulo="Melhor % de ativos.",
        imagens_kv=imagens_kv,
        imagens_tom=imagens_tom,
        col_num="ativos",
        col_den="total",
    )


def montar_email_html(dados, graficos, tabelas, insights, link_drive, teste=False, regional_filtro=None):
    """Monta corpo do e-mail em HTML compativel com Gmail e Outlook."""
    hoje = date.today().strftime("%d/%m/%Y")
    titulo = f"Relatório Semanal Programa +TOP — {regional_filtro}" if regional_filtro else "Relatório Semanal Programa +TOP"

    alerta_teste = "<p style='color:#d9534f; font-weight:bold; margin:16px 0; font-family:Arial;'>[MODO TESTE - e-mail nao enviado]</p>" if teste else ""

    # Período de análise e últimas datas das bases
    def fmt_dt(dt):
        if pd.isna(dt) or dt is None:
            return "N/A"
        if isinstance(dt, pd.Timestamp):
            return dt.strftime("%d/%m/%Y")
        if isinstance(dt, datetime):
            return dt.strftime("%d/%m/%Y")
        return str(dt)

    periodo_inicio, periodo_fim = dados.get("periodo_analise", (None, None))
    datas_ultimas = dados.get("datas_ultimas", {})
    periodo_texto = ""
    if periodo_inicio and periodo_fim:
        periodo_texto = f"""
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 16px 0; font-family:Arial;">
          <tr>
            <td style="padding:14px 16px; background-color:#f9f9f9; border-left:4px solid #ef4e22; border-radius:0 8px 8px 0; color:#333333; font-size:14px; line-height:1.6; font-family:Arial;">
              <strong>Período de análise:</strong> {periodo_inicio.strftime('%d/%m/%Y')} a {periodo_fim.strftime('%d/%m/%Y')}<br>
              <span style="font-size:12px; color:#666666; font-family:Arial;">
                Último dado de cadastros: <strong>{fmt_dt(datas_ultimas.get('cadastro'))}</strong> &nbsp;|&nbsp;
                Último dado de treinamentos: <strong>{fmt_dt(datas_ultimas.get('treinamento'))}</strong> &nbsp;|&nbsp;
                Último dado de aceites: <strong>{fmt_dt(datas_ultimas.get('aceite'))}</strong>
              </span>
            </td>
          </tr>
        </table>
        """

    # Carrega imagens do KV e imagens específicas do Tom
    imagens_kv = carregar_imagens_kv()
    imagens_tom = carregar_imagens_tom()

    # Metricas principais
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    aceite_reg = dados["aceites"]
    aceite_rev = dados.get("aceites_rev")

    pct_geral = round(cad_reg["ativos"].sum() / cad_reg["total"].sum() * 100, 1) if not cad_reg.empty else 0
    pct_trein = round(trein_reg["realizaram"].sum() / trein_reg["total_ativos"].sum() * 100, 1) if trein_reg is not None and not trein_reg.empty else 0
    pct_aceite = round(aceite_reg["aceitaram"].sum() / aceite_reg["total_ativos"].sum() * 100, 1) if not aceite_reg.empty else 0

    metricas = f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0; font-family:Arial;">
      <tr>
        <td align="center" style="font-family:Arial;">
          <table cellpadding="0" cellspacing="8" border="0" style="font-family:Arial;">
            <tr>
              {_metric_box('CADASTROS', pct_geral, META_CADASTRO)}
              {_metric_box('TREINAMENTOS', pct_trein, META_TREINAMENTOS)}
              {_metric_box('ACEITES', pct_aceite, META_ACEITES)}
            </tr>
          </table>
          <p style="font-size:11px; color:#666666; margin-top:6px; font-family:Arial;">
            🟢 Atingiu a meta mínima &nbsp;|&nbsp; 🟡 Entre 70% e a meta &nbsp;|&nbsp; 🔴 Abaixo de 70% da meta
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
            "consolidado_reg": "",
            "consolidado_rev": "",
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
        # Tabelas consolidadas filtradas por regional
        cons_reg_f = preparar_tabela_consolidada(
            cad_reg[cad_reg["regional_curta"] == regional_filtro],
            trein_reg[trein_reg["regional_curta"] == regional_filtro] if trein_reg is not None else None,
            aceite_reg[aceite_reg["regional"] == regional_filtro],
            nivel="regional"
        )
        cons_rev_f = preparar_tabela_consolidada(
            cad_rev[cad_rev["regional_curta"] == regional_filtro],
            trein_rev[trein_rev["regional_curta"] == regional_filtro] if trein_rev is not None else None,
            aceite_reg[aceite_reg["regional"] == regional_filtro],
            nivel="revenda"
        )
        tabelas_usar["consolidado_reg"] = estilizar_tabela_html(cons_reg_f)
        tabelas_usar["consolidado_rev"] = estilizar_tabela_html(cons_rev_f)

        # Tabelas base de treinamentos e aceites filtradas por regional
        trein_base_reg_f = preparar_tabela_base_treinamentos(
            cad_reg[cad_reg["regional_curta"] == regional_filtro],
            trein_reg[trein_reg["regional_curta"] == regional_filtro] if trein_reg is not None else None,
            nivel="regional"
        )
        trein_base_rev_f = preparar_tabela_base_treinamentos(
            cad_rev[cad_rev["regional_curta"] == regional_filtro],
            trein_rev[trein_rev["regional_curta"] == regional_filtro] if trein_rev is not None else None,
            nivel="revenda"
        )
        tabelas_usar["trein_base_reg"] = estilizar_tabela_html(trein_base_reg_f)
        tabelas_usar["trein_base_rev"] = estilizar_tabela_html(trein_base_rev_f)

        aceite_base_reg_f = preparar_tabela_base_aceites(
            cad_reg[cad_reg["regional_curta"] == regional_filtro],
            aceite_reg[aceite_reg["regional"] == regional_filtro],
            nivel="regional"
        )
        aceite_base_rev_f = preparar_tabela_base_aceites(
            cad_rev[cad_rev["regional_curta"] == regional_filtro],
            aceite_reg[aceite_reg["regional"] == regional_filtro],
            nivel="revenda"
        )
        tabelas_usar["aceite_base_reg"] = estilizar_tabela_html(aceite_base_reg_f)
        tabelas_usar["aceite_base_rev"] = estilizar_tabela_html(aceite_base_rev_f)
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
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px; font-family:Arial;">
          <tr><td style="color:#ef4e22; font-size:15px; font-weight:bold; font-family:Arial;">{texto}</td></tr>
        </table>
        """

    html = f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{titulo}</title>
    </head>
    <body style="margin:0; padding:20px; background-color:#f5f5f5; font-family:Arial; color:#333333; line-height:1.6;">
        <!--[if mso]>
        <table role="presentation" width="700" cellspacing="0" cellpadding="0" border="0" align="center">
        <tr><td>
        <![endif]-->
        <table role="presentation" width="100%" max-width="700" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:700px; width:100%; background-color:#ffffff; font-family:Arial;">
          <tr>
            <td style="font-family:Arial;">

              {_header_html(titulo, hoje, imagens_kv)}

              <!-- Conteudo -->
              <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="font-family:Arial;">
                <tr><td style="padding:24px; font-family:Arial;">

                  {alerta_teste}

                  {periodo_texto}

                  {metricas}

                  {_secao_html("CADASTROS")}
                  {_balao_tom_html(
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4; font-family:Arial;'><strong>Queremos levar o +TOP ainda mais longe! A nossa meta mínima é de <span style='color:#00a651; font-family:Arial;'>{META_CADASTRO:.0f}%</span> de <span style='white-space:nowrap; font-family:Arial;'>cadastros ativos</span> e, até esta semana, já alcançamos <span style='color:#ef4e22; font-family:Arial;'>{f'{pct_geral:.1f}'.replace('.', ',')}%</span> da base engajada.</strong></p>"
                      f"<p style='font-size:17px; margin:0; line-height:1.4; font-family:Arial;'><strong>Vamos juntos mobilizar as revendas para buscar o percentual restante!</strong></p>",
                      imagens_kv=imagens_kv, imagens_tom=imagens_tom, tipo_tom="apontando", alinhamento="esquerda"
                  )}
                  {subsecao_titulo("Por regional")}
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
                  {subsecao_titulo("Top 10 revendas com maior % de ativos")}
                  {tabelas_usar['cad_rev']}

                  {_secao_html("TREINAMENTOS")}
                  {_balao_tom_html(
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4; font-family:Arial;'><strong>Nossa meta é ter, no mínimo, <span style='color:#00a651; font-family:Arial;'>{META_TREINAMENTOS:.0f}%</span> dos participantes aprovados e capacitados nos 2 treinamentos do mês.</strong></p>"
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4; font-family:Arial;'><strong>Até o momento, apenas <span style='color:#ef4e22; font-family:Arial;'>{f'{pct_trein:.1f}'.replace('.', ',')}%</span> concluíram os cursos obrigatórios:</strong></p>"
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4; font-family:Arial;'>🔹 <strong>{nome_curso1}</strong> (SKU {dados['cursos_info'][4]})<br>"
                      f"🔹 <strong>{nome_curso2}</strong> (SKU {dados['cursos_info'][5]})</p>"
                      f"<p style='font-size:17px; margin:0; line-height:1.4; font-family:Arial;'><strong>Ainda temos um longo caminho até a meta! Garantir essa capacitação é fundamental para dominar o argumento de vendas dos vendedores para alavancar o nosso Sell Out.</strong></p>",
                      imagens_kv=imagens_kv, imagens_tom=imagens_tom, tipo_tom="apontando", alinhamento="esquerda"
                  )}
                  {subsecao_titulo("Por regional")}
                  {tabelas_usar['trein_base_reg'] if tabelas_usar.get('trein_base_reg') else '<p><em>Sem dados.</em></p>'}
                  {_img_html("grafico_treinamentos" if "treinamentos" in graficos else None, "Gráfico Treinamentos")}
                  {_pontos_atencao_secao_html(
                      insights.get("treinamentos", {}).get("alerta_titulo", ""),
                      insights.get("treinamentos", {}).get("alerta_subtitulo", ""),
                      estilizar_tabela_html(insights.get("treinamentos", {}).get("alerta_itens"), semaforo_coluna="% Realizado", meta_semaforo=META_TREINAMENTOS) if insights.get("treinamentos", {}).get("alerta_itens") is not None else "",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                  )}
                  {_destaques_meta_html(
                      trein_rev,
                      meta=META_TREINAMENTOS,
                      col_pct="pct_realizaram",
                      col_regional="regional_curta",
                      col_revenda="revenda",
                      titulo="Parabéns!",
                      subtitulo="Melhor % de treinamentos concluídos.",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                      col_num="realizaram",
                      col_den="total_ativos",
                  )}
                  {subsecao_titulo("Top 10 revendas com maior % de treinamentos realizados")}
                  {tabelas_usar['trein_base_rev'] if tabelas_usar.get('trein_base_rev') else '<p><em>Sem dados.</em></p>'}

                  <p style="font-size:12px; color:#666666; font-style:italic; margin-top:8px; font-family:Arial;">
                    Dados de treinamentos são sempre D-1.
                  </p>

                  {_secao_html("ACEITES MENSAIS")}
                  {_balao_tom_html(
                      f"<p style='font-size:17px; margin:0 0 10px 0; line-height:1.4; font-family:Arial;'><strong>Nosso objetivo é atingir <span style='color:#00a651; font-family:Arial;'>{META_ACEITES:.0f}%</span> de aceites mensais em {nome_mes_pt_br(ano_mes=str(dados['mes_aceite']))}.</strong></p>"
                      f"<p style='font-size:17px; margin:0; line-height:1.4; font-family:Arial;'><strong>Até o momento, <span style='color:#ef4e22; font-family:Arial;'>{f'{pct_aceite:.1f}'.replace('.', ',')}%</span> dos participantes realizaram o aceite no +TOP (validação mensal necessária para garantir os pontos do programa).</strong></p>",
                      imagens_kv=imagens_kv, imagens_tom=imagens_tom, tipo_tom="apontando", alinhamento="esquerda"
                  )}
                  {subsecao_titulo("Por regional")}
                  {tabelas_usar['aceite_reg']}
                  {_img_html("grafico_aceites" if "aceites" in graficos else None, "Gráfico Aceites")}
                  {_pontos_atencao_secao_html(
                      insights.get("aceites", {}).get("alerta_titulo", ""),
                      insights.get("aceites", {}).get("alerta_subtitulo", ""),
                      estilizar_tabela_html(insights.get("aceites", {}).get("alerta_itens"), semaforo_coluna="% Aceite", meta_semaforo=META_ACEITES) if insights.get("aceites", {}).get("alerta_itens") is not None else "",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                  )}
                  {_destaques_meta_html(
                      aceite_rev,
                      meta=META_ACEITES,
                      col_pct="pct_aceite",
                      col_regional="regional",
                      col_revenda="revenda",
                      titulo="Parabéns!",
                      subtitulo="Melhor % de aceite mensal.",
                      imagens_kv=imagens_kv,
                      imagens_tom=imagens_tom,
                      col_num="aceitaram",
                      col_den="total_ativos",
                  )}
                  {subsecao_titulo("Top 10 revendas com maior % de aceite")}
                  {tabelas_usar['aceite_rev'] if tabelas_usar.get('aceite_rev') else '<p><em>Sem dados de aceites por revenda.</em></p>'}

                  <p style="margin-top:28px; font-size:16px; color:#155724; background-color:#d4edda; padding:14px 16px; border-radius:10px; border:1px solid #00a651; line-height:1.5; font-family:Arial;">
                    💪 <strong>Contamos com a atuação de cada regional para virarmos esse jogo e atingirmos nossas metas!</strong><br>
                    Vamos juntos fazer do +TOP um sucesso ainda maior!
                  </p>

                  <p style="margin-top:24px; font-size:16px; font-family:Arial;">📋 No anexo, você encontra a <strong>base detalhada</strong> de todas as revendas participantes do Programa. Utilize essas informações para direcionar as ações com seus times.</p>

                  <p style="margin-top:16px; font-size:16px; font-family:Arial;">Abraços,<br><strong style="color:#00a651; font-family:Arial;">Time do +TOP</strong></p>

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
    """Prepara aba de detalhamento com campos auditáveis para replicar os cálculos do e-mail."""
    if df_det is None or df_det.empty:
        return None

    df = df_det.copy()
    if regional_filtro:
        df = df[df["regional_da_loja"].apply(regional_curta) == regional_filtro]

    # ------------------------------------------------------------------
    # Colunas cadastrais / hierarquia
    # ------------------------------------------------------------------
    colunas_base = [
        "cpf_limp",
        "regional_da_loja",
        "loja",
        "cnpj_loja",
        "cod_loja",
        "nome",
        "cargo",
        "status",
        "desligado",
        "cidade",
        "uf",
        "bairro",
        "nome_loja_real",
    ]
    for col in colunas_base:
        if col not in df.columns:
            df[col] = None

    df_out = df[colunas_base].copy()
    df_out = df_out.rename(columns={
        "regional_da_loja": "Regional",
        "loja": "Revenda",
        "cnpj_loja": "CNPJ",
        "cod_loja": "Código Loja",
        "nome": "Nome",
        "cargo": "Cargo",
        "status": "Status",
        "desligado": "Desligado",
        "cidade": "Cidade",
        "uf": "UF",
        "bairro": "Bairro loja",
        "nome_loja_real": "Nome loja",
    })

    # Preenche regional vazia a partir do mapeamento revenda -> regional do cadastro
    if dados is not None:
        df_cad = dados.get("cadastro_df")
        if df_cad is not None and "grupo" in df_cad.columns and "regional" in df_cad.columns:
            mapa_regional_det = (
                df_cad.dropna(subset=["grupo", "regional"])
                .drop_duplicates(subset=["grupo"], keep="first")
                .set_index("grupo")["regional"]
                .to_dict()
            )
            mapa_regional_det_norm = {
                normalizar_revenda_hierarquia(str(k).strip()): regional_title_case(str(v))
                for k, v in mapa_regional_det.items()
                if pd.notna(k) and pd.notna(v)
            }
            regional_preenchida = df_out["Revenda"].map(mapa_regional_det_norm)
            df_out["Regional"] = df_out["Regional"].fillna(regional_preenchida)

    def limpar_cnpj(cnpj):
        if pd.isna(cnpj):
            return ""
        s = str(cnpj).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
        if "." in s:
            s = s.split(".")[0]
        return s.zfill(14)

    df_out["CNPJ"] = df_out["CNPJ"].apply(limpar_cnpj)
    # Não adiciona apóstrofo; number_format '@' é aplicado em _formatar_celulas
    # para manter CPF/CNPJ como texto no Excel.

    # ------------------------------------------------------------------
    # Flags calculáveis
    # ------------------------------------------------------------------
    df_out["em_ferias"] = df["em_ferias"].fillna(False)
    base_calculo = df["base_calculo_geral"].fillna(True)
    # CPFs em férias são considerados inativos, mesmo que o cadastro esteja Ativo
    df_out["Status"] = np.where(df_out["em_ferias"], "Inativo", df_out["Status"])
    df_out["Ativo no +TOP?"] = np.where(
        (df_out["Status"].eq("Ativo")) & (base_calculo.eq(True)) & (~df_out["em_ferias"]),
        "Sim", "Não"
    )

    # Preenche campos vazios com informações disponíveis da hierarquia/cadastro
    df_out["Cargo"] = df_out["Cargo"].fillna("Não informado")
    status_preenchido = pd.Series(
        np.where(df_out["Ativo no +TOP?"].eq("Sim"), "Ativo", "Pré-Cadastrado"),
        index=df_out.index
    )
    df_out["Status"] = df_out["Status"].fillna(status_preenchido)
    df_out["Nome loja"] = df_out["Nome loja"].fillna(df_out["Revenda"])

    # Remove colunas auxiliares que não devem ir para a base final
    df_out = df_out.drop(columns=["Cargo na Hierarquia", "em_ferias"], errors="ignore")

    # ------------------------------------------------------------------
    # Aceite mensal
    # ------------------------------------------------------------------
    df_out["Aceite no Mês?"] = "Não"
    df_out["Data Aceite"] = ""
    if dados is not None:
        df_aceite = dados.get("aceites_df")
        mes_aceite_ref = dados.get("mes_aceite") or dados.get("mes_aceite_ref")
        if df_aceite is not None and mes_aceite_ref is not None:
            mes_dt = pd.Period(mes_aceite_ref, freq="M") if isinstance(mes_aceite_ref, str) else mes_aceite_ref
            aceite_mes = df_aceite[df_aceite["mes_aceite"] == mes_dt].copy()
            if not aceite_mes.empty:
                aceite_mes = aceite_mes.sort_values("DataAceite", ascending=False).drop_duplicates("cpf_limp")
                df_out = df_out.merge(
                    aceite_mes[["cpf_limp", "DataAceite"]],
                    on="cpf_limp",
                    how="left",
                )
                df_out["Aceite no Mês?"] = np.where(df_out["DataAceite"].notna(), "Sim", "Não")
                df_out["Data Aceite"] = pd.to_datetime(df_out["DataAceite"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                df_out = df_out.drop(columns=["DataAceite"], errors="ignore")

    # ------------------------------------------------------------------
    # Treinamentos obrigatórios do mês
    # ------------------------------------------------------------------
    colunas_trein = []
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

                flags_cursos = []
                for curso, nome, sku in [(curso1, nome1, sku1), (curso2, nome2, sku2)]:
                    if not curso:
                        continue
                    trein_curso = trein_mes[trein_mes["Curso"] == curso][["cpf_limp", "Conclusão"]].drop_duplicates("cpf_limp")
                    trein_curso["realizado"] = "Sim"
                    trein_curso["data_curso"] = trein_curso["Conclusão"].dt.strftime("%d/%m/%Y")

                    col_realizado = f"{nome} (SKU {sku}) - Realizado?"
                    col_status = f"{nome} (SKU {sku}) - Status"
                    col_data = f"{nome} (SKU {sku}) - Conclusão"
                    colunas_trein.extend([col_realizado, col_status, col_data])

                    df_out = df_out.merge(
                        trein_curso[["cpf_limp", "realizado", "data_curso"]],
                        on="cpf_limp",
                        how="left",
                    )
                    df_out = df_out.rename(columns={
                        "realizado": col_realizado,
                        "data_curso": col_data,
                    })
                    df_out[col_realizado] = df_out[col_realizado].fillna("Não")
                    df_out[col_status] = np.where(df_out[col_realizado].eq("Sim"), "Realizado/Aprovado", "Não realizado")
                    df_out[col_data] = df_out[col_data].fillna("")
                    flags_cursos.append(df_out[col_realizado].eq("Sim"))

                if len(flags_cursos) == 2:
                    df_out["Realizou Ambos os Cursos?"] = np.where(
                        flags_cursos[0] & flags_cursos[1], "Sim", "Não"
                    )

    # ------------------------------------------------------------------
    # Mês de referência
    # ------------------------------------------------------------------
    df_out["Mês de Referência"] = dados.get("mes_referencia", "") if dados else ""

    # ------------------------------------------------------------------
    # Formatação final: remove CPF da base (dado sensível)
    # ------------------------------------------------------------------
    df_out = df_out.drop(columns=["cpf_limp"], errors="ignore")

    # ------------------------------------------------------------------
    # Reordena colunas
    # ------------------------------------------------------------------
    colunas_inicio = [
        "Regional", "Revenda", "CNPJ", "Código Loja",
        "Nome", "Cargo",
        "Status", "Ativo no +TOP?", "Desligado",
        "Cidade", "UF", "Bairro loja", "Nome loja",
        "Aceite no Mês?", "Data Aceite",
    ]
    colunas_fim = ["Realizou Ambos os Cursos?", "Mês de Referência"]

    colunas_existentes = [c for c in colunas_inicio if c in df_out.columns]
    colunas_existentes += [c for c in colunas_trein if c in df_out.columns]
    colunas_existentes += [c for c in colunas_fim if c in df_out.columns]
    # Garante que colunas não listadas também sejam mantidas
    colunas_existentes += [c for c in df_out.columns if c not in colunas_existentes]
    df_out = df_out[[c for c in colunas_existentes if c in df_out.columns]]

    return df_out


def _preparar_porcentagens(df):
    """Converte colunas percentuais de 0-100 para 0-1 para number_format do Excel."""
    df = df.copy()
    for col in df.columns:
        if col in PCT_COLS and pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col] / 100.0
    return df


def _formatar_celulas(ws, df, start_row, start_col):
    """Aplica number_format percentual e texto em colunas de CPF/CNPJ."""
    for col_idx, col in enumerate(df.columns, start_col + 1):
        letter = get_column_letter(col_idx)
        is_pct = col in PCT_COLS
        is_text_id = str(col).upper() in {"CPF", "CNPJ"}
        header_row = start_row + 1
        for r in range(header_row + 1, header_row + 1 + len(df)):
            cell = ws[f"{letter}{r}"]
            if is_pct and isinstance(cell.value, (int, float)):
                cell.number_format = "0.0%"
            elif is_text_id:
                cell.number_format = "@"


def _ajustar_largura_aba(ws):
    """Ajusta largura das colunas com base no conteúdo (máx. 60)."""
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = max(
            (len(str(cell.value or "")) for cell in ws[letter]),
            default=0,
        )
        ws.column_dimensions[letter].width = min(max_len + 2, 60)


def _escrever_tabela(writer, df, sheet_name, startrow=0, startcol=0):
    """Escreve DataFrame no Excel aplicando formatação de % e ajuste de largura."""
    df_out = _preparar_porcentagens(df)
    df_out.to_excel(
        writer, sheet_name=sheet_name, index=False,
        startrow=startrow, startcol=startcol,
    )
    ws = writer.sheets[sheet_name]
    _formatar_celulas(ws, df_out, startrow, startcol)
    _ajustar_largura_aba(ws)
    return startrow + len(df_out) + 1


def _escrever_secao_resumo(writer, df, sheet_name, titulo, startrow=0, startcol=0):
    """
    Escreve uma seção visual na aba Resumo: título estilizado + tabela formatada.
    Retorna a próxima linha disponível.
    """
    # Garante que a aba existe
    if sheet_name not in writer.sheets:
        writer.book.create_sheet(sheet_name)
    ws = writer.sheets[sheet_name]

    # Título da seção
    ws.cell(row=startrow + 1, column=startcol + 1, value=titulo)
    titulo_cell = ws.cell(row=startrow + 1, column=startcol + 1)
    titulo_cell.font = Font(name="Calibri", size=14, bold=True, color=COR_BRANCO)
    titulo_cell.fill = PatternFill(start_color=COR_PRIMARIA, end_color=COR_PRIMARIA, fill_type="solid")
    titulo_cell.alignment = Alignment(horizontal="left", vertical="center")

    # Mescla células do título (até a última coluna do DataFrame)
    n_cols = len(df.columns)
    if n_cols > 1:
        ws.merge_cells(
            start_row=startrow + 1, start_column=startcol + 1,
            end_row=startrow + 1, end_column=startcol + n_cols
        )

    # Escreve a tabela abaixo do título
    df_out = _preparar_porcentagens(df)
    table_startrow = startrow + 1
    df_out.to_excel(
        writer, sheet_name=sheet_name, index=False,
        startrow=table_startrow, startcol=startcol,
        header=True,
    )

    # Formata cabeçalhos da tabela
    header_row = table_startrow + 1
    for col_idx in range(startcol + 1, startcol + n_cols + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.font = Font(name="Calibri", size=11, bold=True, color=COR_TEXTO)
        cell.fill = PatternFill(start_color=COR_PRIMARIA_CLARA, end_color=COR_PRIMARIA_CLARA, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # Formata células de dados
    for row_idx in range(header_row + 1, header_row + 1 + len(df_out)):
        for col_idx, col in enumerate(df_out.columns, startcol + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if col in PCT_COLS else "left", vertical="center")
            if col in PCT_COLS and isinstance(cell.value, (int, float)):
                cell.number_format = "0.0%"
            elif str(col).upper() in {"CPF", "CNPJ"}:
                cell.number_format = "@"

    _ajustar_largura_aba(ws)
    return table_startrow + len(df_out) + 2


def _listar_participantes_treinamento(dados, regional_filtro=None):
    """Retorna DataFrames de participantes que não fizeram nenhum curso e que fizeram apenas 1."""
    base = dados.get("base_trein")
    df_cad = dados.get("cadastro_df")
    df_trein = dados.get("treinamentos_df")
    df_hier = dados.get("hierarquia")
    cursos_info = dados.get("cursos_info")
    mes_ref = dados.get("mes_referencia")

    if base is None or df_cad is None or df_trein is None or cursos_info is None:
        return None, None

    curso1, curso2, nome1, nome2, sku1, sku2 = cursos_info
    if not curso1 or not curso2:
        return None, None

    base_cols = ["cpf_limp", "regional_curta", "revenda"]
    base = base[base_cols].drop_duplicates("cpf_limp").copy()
    base = base.merge(
        df_cad[["cpf_limp", "nome"]].drop_duplicates("cpf_limp"),
        on="cpf_limp", how="left",
    )
    # Fallback: nome da hierarquia quando não houver no cadastro
    if df_hier is not None and "nome_hier" in df_hier.columns:
        base = base.merge(
            df_hier[["cpf_limp", "nome_hier"]].drop_duplicates("cpf_limp"),
            on="cpf_limp", how="left",
        )
        base["nome"] = base["nome"].fillna(base["nome_hier"])
    if regional_filtro:
        base = base[base["regional_curta"] == regional_filtro]

    mes_dt = pd.Period(mes_ref, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ]
    c1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique())
    c2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique())

    base["fez_c1"] = base["cpf_limp"].isin(c1)
    base["fez_c2"] = base["cpf_limp"].isin(c2)
    base["n_cursos"] = base["fez_c1"].astype(int) + base["fez_c2"].astype(int)

    cols_out = {"nome": "Nome", "regional_curta": "Regional", "revenda": "Revenda"}

    df_0 = base[base["n_cursos"] == 0][list(cols_out.keys())].rename(columns=cols_out).copy()

    df_1 = base[base["n_cursos"] == 1][list(cols_out.keys()) + ["fez_c1", "fez_c2"]].copy()
    df_1["Curso Realizado"] = df_1.apply(lambda r: nome1 if r["fez_c1"] else nome2, axis=1)
    df_1 = df_1.drop(columns=["fez_c1", "fez_c2"]).rename(columns=cols_out)

    return df_0, df_1


def _listar_participantes_aceite(dados, regional_filtro=None):
    """Retorna DataFrames de participantes que aceitaram e que não aceitaram."""
    base = dados.get("base_aceite")
    df_cad = dados.get("cadastro_df")
    df_hier = dados.get("hierarquia")

    if base is None or df_cad is None:
        return None, None

    base = base[["cpf_limp", "regional_curta", "revenda", "aceitou"]].drop_duplicates("cpf_limp").copy()
    base = base.merge(
        df_cad[["cpf_limp", "nome"]].drop_duplicates("cpf_limp"),
        on="cpf_limp", how="left",
    )
    # Fallback: nome da hierarquia quando não houver no cadastro
    if df_hier is not None and "nome_hier" in df_hier.columns:
        base = base.merge(
            df_hier[["cpf_limp", "nome_hier"]].drop_duplicates("cpf_limp"),
            on="cpf_limp", how="left",
        )
        base["nome"] = base["nome"].fillna(base["nome_hier"])
    if regional_filtro:
        base = base[base["regional_curta"] == regional_filtro]

    cols_out = {"nome": "Nome", "regional_curta": "Regional", "revenda": "Revenda"}

    df_sim = base[base["aceitou"]][list(cols_out.keys())].rename(columns=cols_out).copy()

    df_nao = base[~base["aceitou"]][list(cols_out.keys())].rename(columns=cols_out).copy()

    return df_sim, df_nao


def _salvar_relatorio_excel_core(dados, caminho, regional_filtro=None):
    """Salva relatório consolidado ou regional em Excel com múltiplas abas formatadas."""
    cad_reg, cad_rev = dados["cadastros"]
    trein_reg, trein_rev = dados["treinamentos"]
    trein_por_curso = dados.get("treinamentos_por_curso")
    aceite_reg = dados["aceites"]
    aceite_rev = dados.get("aceites_rev")
    df_det = dados.get("detalhamento")
    cursos_info = dados.get("cursos_info")

    # Filtros por regional quando aplicável
    if regional_filtro:
        cad_reg_f = cad_reg[cad_reg["regional_curta"] == regional_filtro]
        cad_rev_f = cad_rev[cad_rev["regional_curta"] == regional_filtro]
        trein_reg_f = trein_reg[trein_reg["regional_curta"] == regional_filtro] if trein_reg is not None else None
        trein_rev_f = trein_rev[trein_rev["regional_curta"] == regional_filtro] if trein_rev is not None else None
        aceite_reg_f = aceite_reg[aceite_reg["regional"] == regional_filtro]
        aceite_rev_f = aceite_rev[aceite_rev["regional"] == regional_filtro] if aceite_rev is not None else None
    else:
        cad_reg_f, cad_rev_f, trein_reg_f, trein_rev_f = cad_reg, cad_rev, trein_reg, trein_rev
        aceite_reg_f, aceite_rev_f = aceite_reg, aceite_rev

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        # ------------------------------------------------------------------
        # ABA RESUMO (primeira)
        # ------------------------------------------------------------------
        resumo_linha = 0

        # Período de análise e últimas datas das bases
        periodo_inicio, periodo_fim = dados.get("periodo_analise", (None, None))
        datas_ultimas = dados.get("datas_ultimas", {})

        def fmt_dt_excel(dt):
            if pd.isna(dt) or dt is None:
                return "N/A"
            if isinstance(dt, pd.Timestamp):
                return dt.strftime("%d/%m/%Y")
            if isinstance(dt, datetime):
                return dt.strftime("%d/%m/%Y")
            return str(dt)

        df_periodo = pd.DataFrame({
            "Informação": [
                "Período de análise",
                "Último dado de cadastros",
                "Último dado de treinamentos",
                "Último dado de aceites",
            ],
            "Valor": [
                f"{periodo_inicio.strftime('%d/%m/%Y')} a {periodo_fim.strftime('%d/%m/%Y')}" if periodo_inicio and periodo_fim else "N/A",
                fmt_dt_excel(datas_ultimas.get("cadastro")),
                fmt_dt_excel(datas_ultimas.get("treinamento")),
                fmt_dt_excel(datas_ultimas.get("aceite")),
            ],
        })
        resumo_linha = _escrever_secao_resumo(
            writer, df_periodo, "Resumo", "Período e atualização dos dados", startrow=resumo_linha
        )

        # Indicadores gerais de cadastro (sem divisão por regional)
        total_participantes = int(cad_reg_f["total"].sum())
        ativos = int(cad_reg_f["ativos"].sum())
        pre_cadastro = int(cad_reg_f["pre_cadastro"].sum())
        pct_ativos_total = round(ativos / total_participantes * 100, 1) if total_participantes else 0

        resumo_dados = {
            "Indicador": [
                "Total de participantes",
                "Ativos no +TOP",
                "Pré-Cadastro",
                "% Ativos no total",
            ],
            "Valor": [
                total_participantes,
                ativos,
                pre_cadastro,
                pct_ativos_total,
            ],
        }
        df_resumo = pd.DataFrame(resumo_dados)
        # Formata % como texto para não aplicar number_format 0.0% nesses indicadores
        df_resumo["Valor"] = df_resumo.apply(
            lambda r: f"{r['Valor']:.1f}%" if r["Indicador"] in ["% Ativos no total"] else r["Valor"],
            axis=1,
        )
        resumo_linha = _escrever_secao_resumo(
            writer, df_resumo, "Resumo", "Indicadores Gerais", startrow=resumo_linha
        )

        # Seções por revenda
        resumo_linha = _escrever_secao_resumo(
            writer, _renomear_cadastro_rev(cad_rev_f), "Resumo", "Cadastro", startrow=resumo_linha
        )

        if trein_rev_f is not None:
            resumo_linha = _escrever_secao_resumo(
                writer, _renomear_trein_rev(trein_rev_f), "Resumo", "Treinamentos", startrow=resumo_linha
            )

        if aceite_rev_f is not None:
            resumo_linha = _escrever_secao_resumo(
                writer, _renomear_aceite_rev(aceite_rev_f), "Resumo", "Aceites", startrow=resumo_linha
            )

        # ------------------------------------------------------------------
        # DEMAIS ABAS
        # ------------------------------------------------------------------
        _escrever_tabela(writer, _renomear_cadastro_reg(cad_reg_f), "Cadastro_Regional")
        _escrever_tabela(writer, _renomear_cadastro_rev(cad_rev_f), "Cadastro_Revenda")

        if trein_reg_f is not None:
            _escrever_tabela(writer, _renomear_trein_reg(trein_reg_f), "Treinamento_Regional")
        if trein_rev_f is not None:
            _escrever_tabela(writer, _renomear_trein_rev(trein_rev_f), "Treinamento_Revenda")

        _escrever_tabela(writer, _renomear_aceite_reg(aceite_reg_f), "Aceite_Regional")

        df_det_out = preparar_aba_detalhamento(df_det, dados=dados, regional_filtro=regional_filtro)
        if df_det_out is not None:
            _escrever_tabela(writer, df_det_out, "Detalhamento")

    logger.info(f"Relatório Excel salvo em: {caminho}")


def salvar_relatorio_excel(dados, caminho):
    """Wrapper para salvar relatório consolidado em Excel."""
    _salvar_relatorio_excel_core(dados, caminho)


def salvar_relatorio_excel_regional(dados, caminho, regional_filtro):
    """Wrapper para salvar relatório filtrado por regional em Excel."""
    _salvar_relatorio_excel_core(dados, caminho, regional_filtro=regional_filtro)



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
        cad_reg, cad_rev = calcular_cadastros(df_cad, bases.get("hierarquia"), bases.get("ferias_hier"), bases.get("status_completo"))
        trein_reg, trein_rev, trein_por_curso, cursos_info, base_trein = calcular_treinamentos(
            df_trein, df_cad, ano_mes, bases.get("hierarquia")
        )
        aceite_reg, aceite_rev, mes_aceite_ref, base_aceite = calcular_aceites(
            df_aceite, df_cad, ano_mes, bases["aba_aceite"], bases["usou_ultima_aba"],
            df_hier=bases.get("hierarquia"),
        )

        # -------------------------------------------------
        # VALIDAÇÃO PRÉ-ENVIO
        # -------------------------------------------------
        try:
            validador = ValidadorEnvioRelatorio(
                cadastro_df=df_cad,
                hierarquia_df=bases.get("hierarquia"),
                ferias_df=bases.get("ferias_hier"),
                treinamentos_df=df_trein,
                aceites_df=df_aceite,
                cad_reg=cad_reg,
                cad_rev=cad_rev,
                trein_reg=trein_reg,
                trein_rev=trein_rev,
                aceite_reg=aceite_reg,
                aceite_rev=aceite_rev,
                base_trein=base_trein,
                base_aceite=base_aceite,
                mes_referencia=ano_mes,
                mes_aceite=mes_aceite_ref,
                cursos_info=cursos_info,
                output_dir=OUTPUT_DIR,
                log_dir=LOG_DIR,
                logger_validador=logger,
            )
            ok, relatorio_validacao = validador.executar(gerar_excel=True)
            if not ok:
                logger.error("Validação pré-envio encontrou erros CRÍTICOS. Envio bloqueado.")
                if relatorio_validacao.get("caminho_excel"):
                    logger.error(f"Relatório de validação: {relatorio_validacao['caminho_excel']}")
                sys.exit(1)

            for alerta in relatorio_validacao.get("alertas", []):
                logger.warning(f"[ALERTA] {alerta.regra}: {alerta.mensagem}")
            logger.info("Validação pré-envio concluída: OK")
        except Exception as e:
            logger.exception("Erro ao executar validador pré-envio")
            sys.exit(1)
        # -------------------------------------------------

        # Datas de corte/última atualização das bases
        datas_ultimas = bases.get("datas_ultimas", {})
        data_fim_periodo = date.today()  # data do relatório como fim do período
        data_inicio_periodo = datetime.strptime(ano_mes, "%Y-%m").date().replace(day=1)

        dados = {
            "cadastros": (cad_reg, cad_rev),
            "cadastro_df": df_cad,
            "treinamentos": (trein_reg, trein_rev),
            "treinamentos_df": df_trein,
            "treinamentos_por_curso": trein_por_curso,
            "base_trein": base_trein,
            "aceites": aceite_reg,
            "aceites_df": df_aceite,
            "aceites_rev": aceite_rev,
            "base_aceite": base_aceite,
            "detalhamento": df_det,
            "cursos_info": cursos_info,
            "mes_aceite": mes_aceite_ref,
            "mes_referencia": ano_mes,
            "usou_ultima_aba": bases["usou_ultima_aba"],
            "hierarquia": bases.get("hierarquia"),
            "periodo_analise": (data_inicio_periodo, data_fim_periodo),
            "datas_ultimas": datas_ultimas,
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
        excel_path = OUTPUT_DIR / f"base_detalhada_relatorio_semanal_programa_+TOP_{date.today():%d%m%Y}.xlsx"
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
