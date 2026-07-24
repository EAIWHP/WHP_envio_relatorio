#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório Regional Nordeste — Rhus
----------------------------------
Gera uma planilha Excel com visão geral da regional CONTA NORDESTE e detalhamento
por revenda, contemplando:
  • Cadastro (total, ativos, pré-cadastrados, inativos, % ativos)
  • Treinamentos obrigatórios do mês de referência
  • Aceites mensais
  • Acessos à plataforma
  • Vendas (último mês disponível)
  • Pontuação (créditos, resgates, saldo)

Usa as mesmas bases e regras de negócio do envio_relatorio semanal:
  - CPF/CNPJ sempre como texto
  - Hierarquia filtrada por DESLIGADO = NÃO
  - Status Inativo/Bloqueado/Reprovado/Aguardando Aprovação excluídos
  - Filiais "B" agrupadas na revenda principal

Saída:
  relatorios_gerados/Relatorio_Nordeste_Rhus_DDMMYYYY.xlsx

Uso:
  python3 gerar_relatorio_nordeste_rhus.py
"""

import logging
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from gerar_relatorio_semanal import (
    limpar_cpf,
    normalizar_revenda,
    normalizar_revenda_hierarquia,
    nome_revenda_exibicao,
    regional_curta,
    mapeamento_revenda_regional,
    regional_por_revenda,
    carregar_hierarquias,
    detectar_cursos_obrigatorios,
    read_excel_robusto,
    FILIAIS_B,
    REVENDAS_EXCLUIR,
    STATUS_CADASTRO_EXCLUIR,
    META_CADASTRO,
    META_TREINAMENTOS,
    META_ACEITES,
    BANCO_PARTICIPANTES_PATH,
    PRAZO_PRE_CADASTRO_INATIVO,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
OUTPUT_DIR.mkdir(exist_ok=True)

VENDAS_DIR = BASE_DIR / "Vendas Processadas"

# Revendas que compõem a regional Nordeste (conforme cadastro.xlsx)
REVENDAS_NORDESTE = {
    "Armazem Mateus", "Armazem Mateus B",
    "Guaibim",
    "Imperio", "Imperio B",
    "Laser Eletro", "Laser Eletro B",
    "Millena",
    "Nosso Lar", "Nosso Lar B",
    "Solar Magazine",
    "Zenir",
}

# Normalização visual das revendas da Nordeste
NOME_EXIBICAO = {
    "Armazem Mateus": "Armazém Mateus",
    "Imperio": "Império",
    "Laser Eletro": "Laser Eletro",
    "Solar Magazine": "Solar Magazine",
}

# Mapeamento adicional para nomes que vêm na base de vendas em maiúsculo/curto
MAPA_VENDAS_PARA_EXIBICAO = {
    "NOSSO LAR": "Nosso Lar",
    "LASER": "Laser Eletro",
    "MILLENA": "Millena",
    "SOLAR": "Solar",
    "SOLAR MAGAZINE": "Solar Magazine",
}

COR_PRIMARIA = "EF4E22"
COR_PRIMARIA_CLARA = "FDE8E0"
COR_CINZA = "F2F2F2"
COR_TEXTO = "333333"

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def revenda_exibicao(revenda):
    """Nome amigável para exibição da revenda."""
    rev = str(revenda).strip()
    # Mapeamento direto de nomes que vêm na base de vendas
    if rev.upper() in MAPA_VENDAS_PARA_EXIBICAO:
        return MAPA_VENDAS_PARA_EXIBICAO[rev.upper()]
    # Agrupa filial B
    rev = normalizar_revenda(rev)
    return NOME_EXIBICAO.get(rev, rev)


def is_nordeste(series):
    """Verifica se a regional é a CONTA NORDESTE (suporta Series)."""
    return series.astype(str).str.upper().str.contains("NORDESTE", na=False)


def carregar_cadastro_nordeste():
    """Carrega cadastro, aplica regras de negócio e filtra regional Nordeste."""
    cadastro_path = DATA_DIR / "cadastro.xlsx"
    xl_cad = pd.ExcelFile(cadastro_path)
    df_cad = pd.read_excel(cadastro_path, sheet_name=xl_cad.sheet_names[0])
    df_cad.columns = [c.strip().lower() for c in df_cad.columns]

    df_cad["cpf_limp"] = df_cad["cpf/cnpj"].apply(limpar_cpf)
    df_cad["revenda_original"] = df_cad["grupo"].astype(str).str.strip()
    df_cad["regional_original"] = df_cad["regional"].astype(str).str.strip()

    # Mapeamento revenda -> regional
    mapa_regional = mapeamento_revenda_regional(df_cad)

    # Hierarquias
    df_hier = carregar_hierarquias()
    df_ferias_hier = None
    if isinstance(df_hier, tuple):
        df_hier, df_ferias_hier = df_hier

    if df_hier is not None:
        cols_hier = ["cpf_limp", "revenda", "cod_loja", "cnpj", "nome_hier", "cargo_hier",
                     "vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado"]
        df_hier = df_hier[[c for c in cols_hier if c in df_hier.columns]].copy()
        df_cad = df_cad.merge(df_hier, on="cpf_limp", how="left")
        df_cad["revenda_hier_norm"] = df_cad["revenda"].apply(normalizar_revenda_hierarquia)
        rev_cadastro = df_cad["revenda_original"].apply(normalizar_revenda)
        rev_cad_vazia = rev_cadastro.isna() | rev_cadastro.astype(str).str.strip().str.lower().isin(["", "nan"])
        df_cad["revenda"] = rev_cadastro.where(~rev_cad_vazia, df_cad["revenda_hier_norm"])
        df_cad["revenda"] = df_cad["revenda"].apply(nome_revenda_exibicao)
        df_cad["regional"] = df_cad["revenda"].map(mapa_regional).fillna(df_cad["regional_original"])
        if "nome_hier" in df_cad.columns:
            df_cad["nome"] = df_cad["nome_hier"].fillna(df_cad["nome"])
        if "cargo_hier" in df_cad.columns:
            df_cad["cargo"] = df_cad["cargo_hier"].fillna(df_cad["cargo"])
    else:
        df_cad["revenda"] = df_cad["revenda_original"].apply(nome_revenda_exibicao)
        df_cad["regional"] = df_cad["regional_original"]

    df_cad["regional_curta"] = df_cad["regional"].apply(regional_curta)

    # Regra Pré-Cadastrado 90+ dias na base do banco
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
        if mask_pre_inativo.sum() > 0:
            df_cad.loc[mask_pre_inativo, "status"] = "Inativo"
        df_cad = df_cad.drop(columns=["DataInclusao_dt"])

    # Remove revendas excluídas (teste/excluídas), mas mantém TODOS os status
    df_cad = df_cad[~df_cad["revenda"].isin(REVENDAS_EXCLUIR)].copy()
    df_cad = df_cad.dropna(subset=["regional_curta", "revenda"]).copy()

    # Filtra apenas regional Nordeste
    df_cad = df_cad[is_nordeste(df_cad["regional"])].copy()

    # Agrupa filiais B
    df_cad["revenda_principal"] = df_cad["revenda"].apply(normalizar_revenda)
    df_cad["revenda_exibicao"] = df_cad["revenda_principal"].apply(revenda_exibicao)

    logger.info(f"Cadastro Nordeste: {len(df_cad):,} registros, {df_cad['cpf_limp'].nunique():,} CPFs únicos")
    return df_cad


def carregar_treinamentos(df_cad):
    """Carrega base de treinamentos e cruza com cadastro Nordeste."""
    trein_path = DATA_DIR / "Base_treinamentos.xlsx"
    xl_trein = pd.ExcelFile(trein_path)
    df_trein = pd.read_excel(trein_path, sheet_name=xl_trein.sheet_names[0])
    df_trein["cpf_limp"] = df_trein["CPF"].apply(limpar_cpf)
    df_trein["Conclusão"] = pd.to_datetime(df_trein["Conclusão"], errors="coerce")
    df_trein = df_trein.merge(
        df_cad[["cpf_limp", "status", "regional_curta", "revenda_principal", "revenda_exibicao"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="inner",
    )
    return df_trein


def carregar_aceites(df_cad, ano_mes):
    """Carrega aba de aceites do mês de referência e cruza com cadastro Nordeste."""
    aceite_path = DATA_DIR / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
    if not aceite_path.exists():
        candidatos = sorted(DATA_DIR.glob("*Aceite*.xlsx"))
        if not candidatos:
            raise FileNotFoundError("Arquivo de aceites não encontrado em envio_relatorio/bases/")
        aceite_path = candidatos[-1]

    xl_aceite = pd.ExcelFile(aceite_path)
    aba_aceite = _descobrir_aba_aceites(xl_aceite, ano_mes)
    df_aceite = pd.read_excel(aceite_path, sheet_name=aba_aceite)
    df_aceite["cpf_limp"] = df_aceite["CPF"].apply(limpar_cpf)
    df_aceite["DataAceite"] = pd.to_datetime(df_aceite["DataAceite"], errors="coerce")
    df_aceite = df_aceite.merge(
        df_cad[["cpf_limp", "revenda_exibicao"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="inner",
    )
    return df_aceite, aba_aceite


def _descobrir_aba_aceites(xl, ano_mes):
    """Procura aba do mês de referência nos aceites."""
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
            return aba

    logger.warning(f"Aba de aceites para {ano_mes} não encontrada. Usando última aba: {abas[-1]}")
    return abas[-1]


def carregar_acessos(df_cad, data_inicio=None, data_fim=None):
    """Carrega base de acessos e cruza com cadastro Nordeste."""
    acessos_path = DATA_DIR / "acessos.xlsx"
    xl = pd.ExcelFile(acessos_path)
    df = pd.read_excel(acessos_path, sheet_name=xl.sheet_names[0])
    df.columns = [c.strip() for c in df.columns]
    df["cpf_limp"] = df["Cpf/Cnpj"].apply(limpar_cpf)
    df["Data Acesso"] = pd.to_datetime(df["Data Acesso"], errors="coerce")

    if data_inicio is not None:
        df = df[df["Data Acesso"] >= pd.Timestamp(data_inicio)].copy()
    if data_fim is not None:
        df = df[df["Data Acesso"] <= pd.Timestamp(data_fim)].copy()

    df = df.merge(
        df_cad[["cpf_limp", "revenda_exibicao"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="inner",
    )
    return df


def carregar_vendas(df_cad):
    """Carrega o arquivo de vendas mais recente e filtra revendas da Nordeste."""
    arquivos = sorted(VENDAS_DIR.glob("2*.xlsx"))
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo de vendas encontrado em 'Vendas Processadas/'")

    # Último arquivo por nome (ex: 2026_06.xlsx)
    vendas_path = arquivos[-1]
    df = pd.read_excel(vendas_path)
    df.columns = [c.strip() for c in df.columns]

    # Normaliza CPF
    if "CPF" in df.columns:
        df["cpf_limp"] = df["CPF"].apply(limpar_cpf)
    else:
        raise ValueError("Coluna CPF não encontrada na base de vendas")

    # Normaliza revenda da base de vendas
    df["revenda_vendas"] = df["Revenda"].astype(str).str.strip().apply(normalizar_revenda_hierarquia)
    df["revenda_exibicao"] = df["revenda_vendas"].apply(revenda_exibicao)

    # Filtra apenas revendas da regional Nordeste
    revendas_nordeste = set(df_cad["revenda_exibicao"].unique())
    df = df[df["revenda_exibicao"].isin(revendas_nordeste)].copy()

    return df, vendas_path.stem


def carregar_pontuacao(df_cad):
    """Carrega aba 'PONTOS POR CPF' e cruza com cadastro Nordeste."""
    pontos_path = DATA_DIR / "pontos_por_cpf_.xlsx"
    df = pd.read_excel(pontos_path, sheet_name="PONTOS POR CPF")
    df.columns = [c.strip() for c in df.columns]
    df["cpf_limp"] = df["CPFCNPJ"].apply(limpar_cpf)
    df = df.merge(
        df_cad[["cpf_limp", "revenda_exibicao"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="inner",
    )
    return df


def calcular_cadastro(df_cad):
    """Retorna DataFrame de cadastro por revenda + total regional."""
    cols = ["revenda_exibicao", "status"]
    df = df_cad.drop_duplicates(subset=["cpf_limp"], keep="first").copy()

    resumo = []
    for revenda in sorted(df["revenda_exibicao"].unique()):
        sub = df[df["revenda_exibicao"] == revenda]
        total = len(sub)
        ativos = (sub["status"] == "Ativo").sum()
        pre_cad = (sub["status"] == "Pré-Cadastrado").sum()
        inativos = (sub["status"] == "Inativo").sum()
        pct_ativos = ativos / total if total > 0 else 0
        resumo.append({
            "Revenda": revenda,
            "Total": total,
            "Ativos": ativos,
            "Pré-Cadastrados": pre_cad,
            "Inativos": inativos,
            "% Ativos": pct_ativos,
        })

    # Total regional
    total = len(df)
    ativos = (df["status"] == "Ativo").sum()
    pre_cad = (df["status"] == "Pré-Cadastrado").sum()
    inativos = (df["status"] == "Inativo").sum()
    pct_ativos = ativos / total if total > 0 else 0
    resumo.append({
        "Revenda": "TOTAL REGIONAL",
        "Total": total,
        "Ativos": ativos,
        "Pré-Cadastrados": pre_cad,
        "Inativos": inativos,
        "% Ativos": pct_ativos,
    })

    return pd.DataFrame(resumo)


def calcular_treinamentos(df_trein, df_cad, ano_mes):
    """Retorna DataFrame de treinamentos por revenda + total regional.
    Base = todos os CPFs do cadastro. Quem não aparece na base de treinamentos
    é considerado como não realizado."""
    curso1, curso2, nome1, nome2, sku1, sku2 = detectar_cursos_obrigatorios(df_trein, ano_mes)

    # Base = TODOS os CPFs da revenda na base de cadastro (todos os status)
    cpfs_base = set(df_cad["cpf_limp"].unique())

    # CPFs que concluíram cada curso no mês
    mes_dt = pd.Period(ano_mes, freq="M")
    concluintes_c1 = set()
    concluintes_c2 = set()
    if curso1:
        concluintes_c1 = set(df_trein[
            (df_trein["Curso"] == curso1)
            & (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
            & (df_trein["Estado"].str.lower() == "concluido")
        ]["cpf_limp"].unique())
    if curso2:
        concluintes_c2 = set(df_trein[
            (df_trein["Curso"] == curso2)
            & (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
            & (df_trein["Estado"].str.lower() == "concluido")
        ]["cpf_limp"].unique())

    resumo = []
    for revenda in sorted(df_cad["revenda_exibicao"].unique()):
        cpfs_rev = set(df_cad[df_cad["revenda_exibicao"] == revenda]["cpf_limp"].unique())
        base = len(cpfs_rev)
        c1 = len(cpfs_rev & concluintes_c1)
        c2 = len(cpfs_rev & concluintes_c2)
        ambos = len((cpfs_rev & concluintes_c1) & (cpfs_rev & concluintes_c2))
        pct_ambos = ambos / base if base > 0 else 0
        resumo.append({
            "Revenda": revenda,
            "Base": base,
            f"Concluintes {nome1 or sku1 or 'Curso 1'}": c1,
            f"Concluintes {nome2 or sku2 or 'Curso 2'}": c2,
            "Concluintes Ambos": ambos,
            "% Ambos": pct_ambos,
        })

    # Total regional
    base = len(cpfs_base)
    c1 = len(cpfs_base & concluintes_c1)
    c2 = len(cpfs_base & concluintes_c2)
    ambos = len((cpfs_base & concluintes_c1) & (cpfs_base & concluintes_c2))
    pct_ambos = ambos / base if base > 0 else 0
    resumo.append({
        "Revenda": "TOTAL REGIONAL",
        "Base": base,
        f"Concluintes {nome1 or sku1 or 'Curso 1'}": c1,
        f"Concluintes {nome2 or sku2 or 'Curso 2'}": c2,
        "Concluintes Ambos": ambos,
        "% Ambos": pct_ambos,
    })

    info = {
        "curso1": nome1 or sku1 or "Curso 1",
        "curso2": nome2 or sku2 or "Curso 2",
    }
    return pd.DataFrame(resumo), info


def calcular_aceites(df_aceite, df_cad):
    """Retorna DataFrame de aceites por revenda + total regional."""
    # Base = TODOS os CPFs da revenda (todos os status)
    cpfs_base = set(df_cad["cpf_limp"].unique())
    cpfs_aceite = set(df_aceite["cpf_limp"].unique())

    resumo = []
    for revenda in sorted(df_cad["revenda_exibicao"].unique()):
        cpfs_rev = set(df_cad[df_cad["revenda_exibicao"] == revenda]["cpf_limp"].unique())
        base = len(cpfs_rev)
        aceitaram = len(cpfs_rev & cpfs_aceite)
        pct = aceitaram / base if base > 0 else 0
        resumo.append({
            "Revenda": revenda,
            "Base": base,
            "Aceitaram": aceitaram,
            "Não Aceitaram": base - aceitaram,
            "% Aceite": pct,
        })

    # Total regional
    base = len(cpfs_base)
    aceitaram = len(cpfs_base & cpfs_aceite)
    pct = aceitaram / base if base > 0 else 0
    resumo.append({
        "Revenda": "TOTAL REGIONAL",
        "Base": base,
        "Aceitaram": aceitaram,
        "Não Aceitaram": base - aceitaram,
        "% Aceite": pct,
    })

    return pd.DataFrame(resumo)


def calcular_acessos(df_acessos, df_cad):
    """Retorna DataFrame de acessos por revenda + total regional."""
    resumo = []
    for revenda in sorted(df_cad["revenda_exibicao"].unique()):
        sub = df_acessos[df_acessos["revenda_exibicao"] == revenda]
        total_acessos = len(sub)
        usuarios_unicos = sub["cpf_limp"].nunique()
        resumo.append({
            "Revenda": revenda,
            "Total de Acessos": total_acessos,
            "Usuários Únicos": usuarios_unicos,
        })

    # Total regional
    total_acessos = len(df_acessos)
    usuarios_unicos = df_acessos["cpf_limp"].nunique()
    resumo.append({
        "Revenda": "TOTAL REGIONAL",
        "Total de Acessos": total_acessos,
        "Usuários Únicos": usuarios_unicos,
    })

    return pd.DataFrame(resumo)


def calcular_vendas(df_vendas):
    """Retorna DataFrame de vendas por revenda + total regional."""
    resumo = []
    for revenda in sorted(df_vendas["revenda_exibicao"].unique()):
        sub = df_vendas[df_vendas["revenda_exibicao"] == revenda]
        total_vendas = sub["Vendas"].sum()
        total_pontos = sub["Pontuação"].sum()
        usuarios_unicos = sub["cpf_limp"].nunique()
        resumo.append({
            "Revenda": revenda,
            "Total de Vendas": int(total_vendas),
            "Total de Pontos": float(total_pontos),
            "Usuários com Venda": usuarios_unicos,
        })

    # Total regional
    total_vendas = df_vendas["Vendas"].sum()
    total_pontos = df_vendas["Pontuação"].sum()
    usuarios_unicos = df_vendas["cpf_limp"].nunique()
    resumo.append({
        "Revenda": "TOTAL REGIONAL",
        "Total de Vendas": int(total_vendas),
        "Total de Pontos": float(total_pontos),
        "Usuários com Venda": usuarios_unicos,
    })

    return pd.DataFrame(resumo)


def calcular_pontuacao(df_pontos):
    """Retorna DataFrame de pontuação por revenda + total regional."""
    # Limpa nomes de colunas
    df_pontos.columns = [c.strip() for c in df_pontos.columns]
    # Renomeia colunas para padronizar
    df_pontos = df_pontos.rename(columns={"Créditos": "Creditos", "EXPIRADO": "Expirado", "Saldo Atual": "Saldo Atual"})
    for col in ["Creditos", "Resgates", "Expirado", "Saldo Atual"]:
        df_pontos[col] = pd.to_numeric(df_pontos[col], errors="coerce").fillna(0)

    resumo = []
    for revenda in sorted(df_pontos["revenda_exibicao"].unique()):
        sub = df_pontos[df_pontos["revenda_exibicao"] == revenda]
        creditos = sub["Creditos"].sum()
        resgates = sub["Resgates"].sum()
        expirado = sub["Expirado"].sum()
        saldo = sub["Saldo Atual"].sum()
        usuarios = sub["cpf_limp"].nunique()
        resumo.append({
            "Revenda": revenda,
            "Participantes": usuarios,
            "Créditos": creditos,
            "Resgates": resgates,
            "Expirado": expirado,
            "Saldo Atual": saldo,
        })

    # Total regional
    creditos = df_pontos["Creditos"].sum()
    resgates = df_pontos["Resgates"].sum()
    expirado = df_pontos["Expirado"].sum()
    saldo = df_pontos["Saldo Atual"].sum()
    usuarios = df_pontos["cpf_limp"].nunique()
    resumo.append({
        "Revenda": "TOTAL REGIONAL",
        "Participantes": usuarios,
        "Créditos": creditos,
        "Resgates": resgates,
        "Expirado": expirado,
        "Saldo Atual": saldo,
    })

    return pd.DataFrame(resumo)


def criar_resumo_performance(df_cadastro, df_trein, df_aceite, df_acessos, df_vendas, df_pontos,
                             info_trein, mes_ref, mes_vendas):
    """Cria aba de resumo com KPIs gerais da regional."""
    # Cadastro
    total_cad = len(df_cadastro[df_cadastro["Revenda"] == "TOTAL REGIONAL"])
    row_cad = df_cadastro[df_cadastro["Revenda"] == "TOTAL REGIONAL"].iloc[0]

    # Treinamento
    row_trein = df_trein[df_trein["Revenda"] == "TOTAL REGIONAL"].iloc[0]

    # Aceite
    row_aceite = df_aceite[df_aceite["Revenda"] == "TOTAL REGIONAL"].iloc[0]

    # Acessos
    row_acessos = df_acessos[df_acessos["Revenda"] == "TOTAL REGIONAL"].iloc[0]

    # Vendas
    row_vendas = df_vendas[df_vendas["Revenda"] == "TOTAL REGIONAL"].iloc[0]

    # Pontuação
    row_pontos = df_pontos[df_pontos["Revenda"] == "TOTAL REGIONAL"].iloc[0]

    dados = [
        ["Indicador", "Valor", "Observação"],
        ["Regional", "CONTA NORDESTE", ""],
        ["Mês de referência", mes_ref, "Cadastro / Treinamento / Aceites"],
        ["Mês de vendas", mes_vendas, "Último mês disponível na base"],
        ["Total de revendas", df_cadastro[df_cadastro["Revenda"] != "TOTAL REGIONAL"].shape[0], ""],
        ["", "", ""],
        ["CADASTRO", "", ""],
        ["Total de participantes", int(row_cad["Total"]), ""],
        ["Ativos", int(row_cad["Ativos"]), ""],
        ["% Ativos", f"{row_cad['% Ativos']*100:.1f}%", f"Meta mínima: {META_CADASTRO:.0f}%"],
        ["Pré-Cadastrados", int(row_cad["Pré-Cadastrados"]), ""],
        ["Inativos", int(row_cad["Inativos"]), ""],
        ["", "", ""],
        ["TREINAMENTOS", "", ""],
        ["Curso 1", info_trein["curso1"], ""],
        ["Curso 2", info_trein["curso2"], ""],
        ["Base", int(row_trein["Base"]), ""],
        ["Concluintes ambos", int(row_trein["Concluintes Ambos"]), ""],
        ["% Ambos", f"{row_trein['% Ambos']*100:.1f}%", f"Meta mínima: {META_TREINAMENTOS:.0f}%"],
        ["", "", ""],
        ["ACEITES", "", ""],
        ["Base", int(row_aceite["Base"]), ""],
        ["Aceitaram", int(row_aceite["Aceitaram"]), ""],
        ["% Aceite", f"{row_aceite['% Aceite']*100:.1f}%", f"Meta mínima: {META_ACEITES:.0f}%"],
        ["", "", ""],
        ["ACESSOS", "", ""],
        ["Total de acessos", int(row_acessos["Total de Acessos"]), ""],
        ["Usuários únicos", int(row_acessos["Usuários Únicos"]), ""],
        ["", "", ""],
        ["VENDAS", "", ""],
        ["Total de vendas", int(row_vendas["Total de Vendas"]), ""],
        ["Total de pontos vendas", float(row_vendas["Total de Pontos"]), ""],
        ["Usuários com venda", int(row_vendas["Usuários com Venda"]), ""],
        ["", "", ""],
        ["PONTUAÇÃO", "", ""],
        ["Participantes com saldo", int(row_pontos["Participantes"]), ""],
        ["Créditos acumulados", float(row_pontos["Créditos"]), ""],
        ["Resgates", float(row_pontos["Resgates"]), ""],
        ["Expirado", float(row_pontos["Expirado"]), ""],
        ["Saldo Atual", float(row_pontos["Saldo Atual"]), ""],
    ]
    return pd.DataFrame(dados[1:], columns=dados[0])


# Colunas que devem ser formatadas como moeda (R$)
COLUNAS_MOEDA = {
    "Total de Pontos", "Créditos", "Resgates", "Expirado", "Saldo Atual",
    "Pontos Vendas", "Créditos Acumulados", "Total de pontos vendas",
    "Créditos acumulados",
}


def salvar_excel_formatado(saida, abas):
    """Salva DataFrames em arquivo Excel com formatação padronizada."""
    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        for nome_aba, df in abas.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)
            ws = writer.sheets[nome_aba]

            # Cabeçalho
            header_fill = PatternFill(start_color=COR_PRIMARIA, end_color=COR_PRIMARIA, fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

            # Largura das colunas e alinhamento
            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        val = str(cell.value) if cell.value is not None else ""
                        max_length = max(max_length, len(val))
                        header_val = str(ws.cell(1, cell.column).value)
                        # Formatação percentual
                        if "%" in header_val:
                            cell.number_format = "0.0%"
                            cell.alignment = Alignment(horizontal="right")
                        # Formatação moeda
                        elif any(moeda in header_val for moeda in COLUNAS_MOEDA):
                            cell.number_format = '"R$" #,##0.00'
                            cell.alignment = Alignment(horizontal="right")
                        # Destaque para linha TOTAL REGIONAL
                        elif isinstance(cell.value, str) and cell.value == "TOTAL REGIONAL":
                            cell.font = Font(bold=True)
                            cell.fill = PatternFill(start_color=COR_PRIMARIA_CLARA, end_color=COR_PRIMARIA_CLARA, fill_type="solid")
                        else:
                            cell.border = thin_border
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 45)

            # Ajustes específicos por aba
            if nome_aba == "Resumo":
                indicadores_moeda = {
                    "Total de pontos vendas", "Créditos acumulados", "Resgates", "Expirado", "Saldo Atual"
                }
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
                    indicador_cell = row[0]
                    if indicador_cell.value in indicadores_moeda:
                        valor_cell = ws.cell(indicador_cell.row, 2)
                        valor_cell.number_format = '"R$" #,##0.00'
                        valor_cell.alignment = Alignment(horizontal="right")

            # Congela primeira linha
            ws.freeze_panes = "A2"

    logger.info(f"Relatório salvo em: {saida}")


def main():
    logger.info("Iniciando geração do relatório da regional Nordeste")

    # Cadastro
    df_cad = carregar_cadastro_nordeste()

    # Mês de referência: último mês com treinamentos obrigatórios na base
    df_trein_full = carregar_treinamentos(df_cad)
    obr_concl = df_trein_full[
        (df_trein_full["Estado"].astype(str).str.lower() == "concluido")
        & (df_trein_full["Trilha"].astype(str).str.contains("OBRIGAT", case=False, na=False))
        & (df_trein_full["Conclusão"].notna())
    ]
    if not obr_concl.empty:
        ano_mes_ref = str(obr_concl["Conclusão"].dt.to_period("M").max())
    else:
        ano_mes_ref = date.today().strftime("%Y-%m")
    mes_ref_nome = f"{ano_mes_ref.split('-')[1]}/{ano_mes_ref.split('-')[0]}"
    logger.info(f"Mês de referência: {mes_ref_nome}")

    # Aceites
    df_aceite, aba_aceite = carregar_aceites(df_cad, ano_mes_ref)

    # Acessos (mês de referência)
    inicio_mes = pd.Timestamp(f"{ano_mes_ref}-01")
    # Último dia do mês
    if inicio_mes.month == 12:
        fim_mes = pd.Timestamp(f"{inicio_mes.year}-12-31")
    else:
        fim_mes = pd.Timestamp(f"{inicio_mes.year}-{inicio_mes.month + 1:02d}-01") - pd.Timedelta(days=1)
    df_acessos = carregar_acessos(df_cad, data_inicio=inicio_mes, data_fim=fim_mes)

    # Vendas
    df_vendas, nome_vendas = carregar_vendas(df_cad)
    mes_vendas = f"{nome_vendas.split('_')[1]}/{nome_vendas.split('_')[0]}"

    # Pontuação
    df_pontos = carregar_pontuacao(df_cad)

    # Cálculos
    df_cadastro = calcular_cadastro(df_cad)
    df_trein, info_trein = calcular_treinamentos(df_trein_full, df_cad, ano_mes_ref)
    df_aceite = calcular_aceites(df_aceite, df_cad)
    df_acessos_raw = df_acessos.copy()
    df_acessos = calcular_acessos(df_acessos, df_cad)
    df_vendas_raw = df_vendas.copy()
    df_vendas = calcular_vendas(df_vendas)
    df_pontos_raw = df_pontos.copy()
    df_pontos = calcular_pontuacao(df_pontos)
    df_resumo = criar_resumo_performance(
        df_cadastro, df_trein, df_aceite, df_acessos, df_vendas, df_pontos,
        info_trein, mes_ref_nome, mes_vendas
    )

    # Geração do detalhamento
    df_det = df_cad[["cpf_limp", "nome", "cargo", "status", "revenda_exibicao", "regional_curta"]].copy()
    df_det = df_det.rename(columns={
        "cpf_limp": "CPF",
        "revenda_exibicao": "Revenda",
        "regional_curta": "Regional",
    })

    # Flag aceite no mês
    cpfs_aceite = set(df_aceite["cpf_limp"].unique()) if "cpf_limp" in df_aceite.columns else set()
    df_det["Aceite no Mês"] = df_det["CPF"].isin(cpfs_aceite).map({True: "Sim", False: "Não"})

    # Flags treinamento
    mes_dt = pd.Period(ano_mes_ref, freq="M")
    cursos = df_trein_full[
        (df_trein_full["Estado"].astype(str).str.lower() == "concluido")
        & (df_trein_full["Trilha"].astype(str).str.contains("OBRIGAT", case=False, na=False))
        & (df_trein_full["Conclusão"].dt.to_period("M") == mes_dt)
    ]
    cpfs_c1 = set()
    cpfs_c2 = set()
    if info_trein["curso1"] != "Curso 1":
        cpfs_c1 = set(cursos[cursos["Curso"].str.contains(info_trein["curso1"].split()[-1] if " " in info_trein["curso1"] else info_trein["curso1"], case=False, na=False)]["cpf_limp"].unique())
    if info_trein["curso2"] != "Curso 2":
        cpfs_c2 = set(cursos[cursos["Curso"].str.contains(info_trein["curso2"].split()[-1] if " " in info_trein["curso2"] else info_trein["curso2"], case=False, na=False)]["cpf_limp"].unique())
    df_det[f"Concluiu {info_trein['curso1']}"] = df_det["CPF"].isin(cpfs_c1).map({True: "Sim", False: "Não"})
    df_det[f"Concluiu {info_trein['curso2']}"] = df_det["CPF"].isin(cpfs_c2).map({True: "Sim", False: "Não"})

    # Acessos
    acessos_por_cpf = df_acessos_raw.groupby("cpf_limp").size().to_dict() if not df_acessos_raw.empty and "cpf_limp" in df_acessos_raw.columns else {}
    df_det["Qtd Acessos no Mês"] = df_det["CPF"].map(acessos_por_cpf).fillna(0).astype(int)

    # Vendas
    vendas_por_cpf = df_vendas_raw.groupby("cpf_limp").agg({"Vendas": "sum", "Pontuação": "sum"}).to_dict("index") if not df_vendas_raw.empty and "cpf_limp" in df_vendas_raw.columns else {}
    df_det["Vendas no Mês"] = df_det["CPF"].map(lambda x: vendas_por_cpf.get(x, {}).get("Vendas", 0)).fillna(0).astype(int)
    df_det["Pontos Vendas"] = df_det["CPF"].map(lambda x: vendas_por_cpf.get(x, {}).get("Pontuação", 0)).fillna(0)

    # Pontuação
    df_pontos_raw.columns = [c.strip() for c in df_pontos_raw.columns]
    df_pontos_raw = df_pontos_raw.rename(columns={"Créditos": "Creditos", "EXPIRADO": "Expirado"})
    for col in ["Creditos", "Resgates", "Expirado", "Saldo Atual"]:
        df_pontos_raw[col] = pd.to_numeric(df_pontos_raw[col], errors="coerce").fillna(0)
    pontos_por_cpf = df_pontos_raw.set_index("cpf_limp")[["Creditos", "Resgates", "Saldo Atual"]].to_dict("index") if not df_pontos_raw.empty and "cpf_limp" in df_pontos_raw.columns else {}
    df_det["Créditos Acumulados"] = df_det["CPF"].map(lambda x: pontos_por_cpf.get(x, {}).get("Creditos", 0)).fillna(0)
    df_det["Resgates"] = df_det["CPF"].map(lambda x: pontos_por_cpf.get(x, {}).get("Resgates", 0)).fillna(0)
    df_det["Saldo Atual"] = df_det["CPF"].map(lambda x: pontos_por_cpf.get(x, {}).get("Saldo Atual", 0)).fillna(0)

    saida = OUTPUT_DIR / f"Relatorio_Nordeste_Rhus_{date.today():%d%m%Y}.xlsx"
    abas = {
        "Resumo": df_resumo,
        "Cadastro_Revenda": df_cadastro,
        "Treinamento_Revenda": df_trein,
        "Aceite_Revenda": df_aceite,
        "Acessos_Revenda": df_acessos,
        "Vendas_Revenda": df_vendas,
        "Pontuacao_Revenda": df_pontos,
        "Detalhamento": df_det,
    }

    salvar_excel_formatado(saida, abas)
    print(f"\n✅ Relatório gerado com sucesso: {saida}")
    print(f"   Regional: CONTA NORDESTE")
    print(f"   Mês ref.: {mes_ref_nome}  |  Vendas: {mes_vendas}")


if __name__ == "__main__":
    main()
