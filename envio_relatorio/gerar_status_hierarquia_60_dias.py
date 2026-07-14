#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_status_hierarquia_60_dias.py

Gera um arquivo Excel com uma aba por revenda, utilizando apenas:
- Hierarquias da pasta bases/bases_cadastro_hierarquia
- Cadastro da plataforma em bases/cadastro.xlsx

NÃO utiliza o arquivo de banco PROD_WHP.

Cada aba contém:
- Dados da hierarquia (todos os registros, incluindo desligados/férias/benefício)
- Flag de situação na hierarquia: Desligado, Férias, Benefício, Ativo ou Vazio
- Flag indicando se o CPF está na base de cadastro +TOP
- Status do cadastro
- Data de aceite do cadastro
- Data de corte utilizada
- Dias desde a data de aceite até a data de corte
- Flag indicando se passou de 60 dias

A data de corte padrão é 02/07/2026, mas pode ser alterada na variável DATA_CORTE.
"""

import logging
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "bases"
HIERARQUIA_DIR = DATA_DIR / "bases_cadastro_hierarquia"
SAIDA_DIR = BASE_DIR / "relatorios_gerados"
SAIDA_DIR.mkdir(exist_ok=True)

CADASTRO_FILE = DATA_DIR / "cadastro.xlsx"
BANCO_FILE = Path("/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx")
SAIDA_FILE = SAIDA_DIR / f"status_hierarquia_60_dias_{date.today():%Y%m%d}.xlsx"

# Data de corte para cálculo dos 60 dias (pode ser alterada manualmente)
DATA_CORTE = date(2026, 7, 2)

MAPEAMENTO_REVENDA = {
    "ANGELONI": "Angeloni",
    "ARMAZEM MATEUS": "Armazem Mateus",
    "BECKER": "Becker",
    "BEMOL": "Bemol",
    "CASAS DA AGUA": "Casas da Água",
    "COLOMBO": "Colombo",
    "ESTRELA": "Estrela",
    "FORMOSA": "Formosa",
    "GAZIN ATACADO": "Gazin Atacado",
    "GAZIN ONLINE": "Gazin Online",
    "GAZIN VAREJO": "Gazin Varejo",
    "GUAIBIM": "Guaibim",
    "HAVAN": "Havan",
    "IMPERIO": "Imperio",
    "JMAHFUZ": "Jmahfuz",
    "KOERICH": "Koerich",
    "LASER": "Laser Eletro",
    "LEBES": "Lebes",
    "LOJAS SOLAR": "Solar",
    "LOJA SOLAR": "Solar",
    "MAGAZAN": "Magazan",
    "MILLENA": "Millena",
    "MM ATACADO": "MM Atacado",
    "MM VAREJO": "MM Varejo",
    "MULTILOJA": "Multiloja",
    "NOSSO LAR": "Nosso Lar",
    "RAMSONS": "Ramsons",
    "SIPOLATTI": "Sipolatti",
    "SOLAR MAGAZINE": "Solar Magazine",
    "TAQI": "Taqi",
    "TELE RIO": "Tele Rio",
    "ZEMA": "Zema",
    "ZENIR": "Zenir",
}


def limpar_cpf(cpf):
    """Normaliza CPF em texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    cpf_str = str(cpf).strip().replace("'", "")
    cpf_limpo = re.sub(r"[^0-9]", "", cpf_str)
    if not cpf_limpo:
        return None
    return cpf_limpo.zfill(11)


def formatar_cpf(cpf):
    """Formata CPF com pontos e traço para leitura humana."""
    if not cpf or len(cpf) != 11:
        return cpf
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


def normalizar_desligado(valor):
    """Normaliza o valor da coluna DESLIGADO."""
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


def classificar_situacao(valor):
    """Classifica a situação do participante na hierarquia."""
    v = normalizar_desligado(valor)
    if v in ("", "NAN", "NONE"):
        return "Vazio"
    if v in ("NÃO", "NAO", "N"):
        return "Ativo"
    if v in ("SIM", "S"):
        return "Desligado"
    if "FÉRIAS" in v or "FERIAS" in v:
        return "Férias"
    if "BENEFICIO" in v or "BENEFÍCIO" in v:
        return "Benefício"
    return "Outro"


def descobrir_arquivos_hierarquia():
    """Retorna lista de arquivos de hierarquia a processar (preferência corrigidos)."""
    todos = sorted(HIERARQUIA_DIR.glob("*.xlsx"))

    ignorar = {
        "consolidado", "comparacao", "resumo", "cpfs_repetidos", "banco_hierarquia",
        "_limpa", "GUAIBIM_HIERARQUIA_LIMPA",
    }
    candidatos = []
    for f in todos:
        nome_upper = f.name.upper()
        if "HIERARQUIA" not in nome_upper:
            continue
        if any(ign.upper() in nome_upper for ign in ignorar):
            continue
        candidatos.append(f)

    corrigidos = {f for f in candidatos if "_CORRIGIDO" in f.name.upper()}
    bases_corrigidas = set()
    for c in corrigidos:
        base = re.sub(r"_corrigido", "", c.stem, flags=re.IGNORECASE)
        bases_corrigidas.add(base.lower())

    selecionados = []
    for f in candidatos:
        if "_CORRIGIDO" in f.name.upper():
            selecionados.append(f)
        else:
            base = f.stem.lower()
            if base not in bases_corrigidas:
                selecionados.append(f)

    return sorted(selecionados)


def extrair_nome_revenda(nome_arquivo):
    """Extrai o nome da revenda a partir do nome do arquivo."""
    nome = nome_arquivo.replace(".xlsx", "").strip().upper()
    nome = re.sub(r"_CORRIGIDO$", "", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\s*\(\d+\)$", "", nome)
    nome = nome.strip()

    for chave, valor in MAPEAMENTO_REVENDA.items():
        if nome == chave:
            return valor
    for chave, valor in MAPEAMENTO_REVENDA.items():
        if chave in nome or nome in chave:
            return valor
    return nome.title()


def carregar_hierarquia(arquivo):
    """Carrega um arquivo de hierarquia individual (todos os registros)."""
    logger.info(f"Lendo hierarquia: {arquivo.name}")
    try:
        df = pd.read_excel(arquivo, dtype=str)
    except Exception as e:
        logger.warning(f"Erro ao ler {arquivo.name}: {e}")
        return None

    df.columns = [str(c).strip().upper() for c in df.columns]

    mapeamento = {
        "REVENDA": "revenda",
        "COD LOJA": "cod_loja",
        "CODLOJA": "cod_loja",
        "CNPJ": "cnpj",
        "CPF": "cpf",
        "NOME": "nome",
        "VENDEDOR": "vendedor",
        "GERENTE DE LOJA": "gerente_loja",
        "GERENTE REGIONAL": "gerente_regional",
        "DIRETOR": "diretor",
        "DESLIGADO": "desligado",
        "CARGO": "cargo",
    }
    df = df.rename(columns=mapeamento)

    for col in ["revenda", "cod_loja", "cnpj", "cpf", "nome", "desligado"]:
        if col not in df.columns:
            df[col] = None

    df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()

    if df.empty:
        return None

    df["situacao_hierarquia"] = df["desligado"].apply(classificar_situacao)
    df["flag_desligado"] = df["situacao_hierarquia"].apply(lambda x: "Sim" if x == "Desligado" else "Não")
    df["flag_ferias"] = df["situacao_hierarquia"].apply(lambda x: "Sim" if x == "Férias" else "Não")
    df["flag_beneficio"] = df["situacao_hierarquia"].apply(lambda x: "Sim" if x == "Benefício" else "Não")
    df["flag_ativo"] = df["situacao_hierarquia"].apply(lambda x: "Sim" if x == "Ativo" else "Não")

    df["revenda_arquivo"] = extrair_nome_revenda(arquivo.name)
    return df


def carregar_cadastro():
    """Carrega o cadastro da plataforma."""
    logger.info(f"Lendo cadastro: {CADASTRO_FILE.name}")
    df = pd.read_excel(CADASTRO_FILE, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    df["cpf_limp"] = df["cpf/cnpj"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()
    df["status"] = df["status"].fillna("").astype(str).str.strip().str.title()
    df["grupo"] = df["grupo"].fillna("").astype(str).str.strip()
    df["grupo_mapeado"] = df["grupo"].replace({"CDA": "Casas da Água"})

    # Converte data de aceite para datetime
    df["data de aceite"] = pd.to_datetime(df["data de aceite"], errors="coerce")

    logger.info(f"Cadastro: {len(df):,} registros")
    return df


def carregar_banco():
    """Carrega a base do banco PROD_WHP usada no envio_relatorio."""
    logger.info(f"Lendo banco: {BANCO_FILE.name}")
    df = pd.read_excel(BANCO_FILE, dtype={"Cpf": str, "Ativo": str})

    df = df.rename(columns={
        "Nome": "revenda_banco",
        "Nome.1": "nome_banco",
        "Cpf": "cpf",
        "Ativo": "ativo_banco",
        "DataInclusao": "data_inclusao",
        "DataBloqueio": "data_bloqueio",
        "MotivoBloqueio": "motivo_bloqueio",
        "OrigemCadastro": "origem_cadastro",
        "Matricula": "matricula_banco",
        "Email": "email_banco",
    })

    df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()
    df["data_inclusao"] = pd.to_datetime(df["data_inclusao"], errors="coerce")
    df["ativo_banco"] = df["ativo_banco"].apply(
        lambda x: "Sim" if str(x).strip() in ("1", "True", "SIM", "Sim", "true") else "Não"
    )

    # Mantém apenas o primeiro registro por CPF
    df = df.drop_duplicates(subset=["cpf_limp"], keep="first").copy()

    logger.info(f"Banco: {len(df):,} CPFs únicos")
    return df


def processar_revenda(df_hier, df_cad, df_banco):
    """Processa uma revenda e retorna DataFrame com flags e status."""
    revenda = df_hier["revenda_arquivo"].iloc[0]

    # Base: hierarquia (todos os registros)
    cols_hier = [
        "cpf_limp", "cpf", "nome", "revenda", "revenda_arquivo", "cod_loja", "cnpj",
        "cargo", "vendedor", "gerente_loja", "gerente_regional", "diretor",
        "desligado", "situacao_hierarquia", "flag_desligado", "flag_ferias",
        "flag_beneficio", "flag_ativo"
    ]
    cols_hier_presentes = [c for c in cols_hier if c in df_hier.columns]
    df = df_hier[cols_hier_presentes].copy()

    df["cpf_formatado"] = df["cpf_limp"].apply(formatar_cpf)

    # Cruza com cadastro
    df_cad_merge = df_cad[[
        "cpf_limp", "status", "nome", "grupo", "grupo_mapeado", "regional",
        "cargo", "matricula", "email", "data de aceite"
    ]].rename(columns={
        "status": "status_cadastro",
        "nome": "nome_cadastro",
        "grupo": "grupo_cadastro",
        "regional": "regional_cadastro",
        "cargo": "cargo_cadastro",
        "matricula": "matricula_cadastro",
        "email": "email_cadastro",
        "data de aceite": "data_aceite_cadastro",
    }).drop_duplicates(subset=["cpf_limp"], keep="first")

    df = df.merge(df_cad_merge, on="cpf_limp", how="left")
    df["na_base_cadastro"] = df["status_cadastro"].notna().map({True: "Sim", False: "Não"})
    # Preenche status vazio como "Não Cadastrado"
    df["status"] = df["status_cadastro"].fillna("Não Cadastrado")

    # Cruza com banco (base usada no envio_relatorio)
    df_banco_merge = df_banco[[
        "cpf_limp", "nome_banco", "ativo_banco", "data_inclusao",
        "data_bloqueio", "motivo_bloqueio", "origem_cadastro", "matricula_banco"
    ]].drop_duplicates(subset=["cpf_limp"], keep="first")

    df = df.merge(df_banco_merge, on="cpf_limp", how="left")
    df["na_base_banco"] = df["data_inclusao"].notna().map({True: "Sim", False: "Não"})

    # Data de corte e cálculo de 60 dias a partir da DataInclusao do banco
    df["data_corte"] = pd.to_datetime(DATA_CORTE)
    df["dias_desde_inclusao"] = (df["data_corte"] - df["data_inclusao"]).dt.days
    df["passou_60_dias"] = df["dias_desde_inclusao"].apply(
        lambda x: "Sim" if pd.notna(x) and x > 60 else "Não"
    )

    # Renomeia gerente_regional da hierarquia para regional_hierarquia
    if "gerente_regional" in df.columns:
        df["regional_hierarquia"] = df["gerente_regional"]

    # Reordena e seleciona colunas finais
    colunas_finais = [
        "cpf_limp", "nome", "revenda", "revenda_arquivo", "cod_loja", "cnpj",
        "cargo", "vendedor", "gerente_loja", "regional_hierarquia", "desligado",
        "na_base_cadastro", "status", "cargo_cadastro",
        "na_base_banco", "data_inclusao", "data_corte", "dias_desde_inclusao", "passou_60_dias",
    ]
    colunas_finais = [c for c in colunas_finais if c in df.columns]
    df = df[colunas_finais].copy()

    return revenda, df


def regional_curta(regional):
    """Remove prefixo 'CONTA ' da regional."""
    r = str(regional).strip()
    return r.replace("CONTA ", "") if r.startswith("CONTA ") else r


def regional_title_case(regional):
    """Retorna nome da regional em Title Case (ex: COMPRA DIRETA -> Compra Direta)."""
    r = str(regional).strip()
    return r.title() if r and r.lower() != "nan" else r


def mapeamento_revenda_regional(df_cad):
    """Cria dicionário revenda -> regional a partir do cadastro base."""
    mapa = {}
    if "grupo" in df_cad.columns and "regional" in df_cad.columns:
        for _, row in df_cad.dropna(subset=["grupo", "regional"]).iterrows():
            rev = str(row["grupo"]).strip()
            reg = str(row["regional"]).strip()
            if rev and reg and rev.lower() != "nan" and reg.lower() != "nan":
                mapa[rev] = reg
    # Fallback para Casas da Água via CDA
    if "CDA" in mapa and "Casas da Água" not in mapa:
        mapa["Casas da Água"] = mapa["CDA"]
    return mapa


def regional_por_revenda_mapa(revenda, mapa_regional):
    """Busca regional no mapa de forma case-insensitiva."""
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
    return None


def calcular_resumos(df_consolidado, mapa_regional):
    """
    Calcula resumos por revenda e por regional a partir do consolidado.
    O consolidado já contém apenas CPFs únicos globais com DESLIGADO=NÃO.
    """
    df = df_consolidado.copy()
    df["regional"] = df["revenda_arquivo"].apply(lambda r: regional_por_revenda_mapa(r, mapa_regional))
    df["ativo"] = df["status"].astype(str).str.strip().str.title() == "Ativo"

    # Resumo por revenda
    resumo_rev = (
        df.groupby(["regional", "revenda_arquivo"])
        .agg(total=("cpf_limp", "count"), ativos=("ativo", "sum"))
        .reset_index()
    )
    resumo_rev["ativos"] = resumo_rev["ativos"].astype(int)
    resumo_rev["inativos"] = resumo_rev["total"] - resumo_rev["ativos"]
    resumo_rev["pct_ativos"] = (resumo_rev["ativos"] / resumo_rev["total"] * 100).round(1)
    resumo_rev = resumo_rev.rename(columns={
        "revenda_arquivo": "Revenda",
        "regional": "Regional",
        "total": "Total participantes",
        "ativos": "Ativos no +TOP",
        "inativos": "Inativos no +TOP",
        "pct_ativos": "% Ativos",
    })
    resumo_rev = resumo_rev[[
        "Regional", "Revenda", "Total participantes", "Ativos no +TOP",
        "Inativos no +TOP", "% Ativos"
    ]]

    # Resumo por regional
    resumo_reg = (
        resumo_rev.groupby("Regional")
        .agg(
            total=("Total participantes", "sum"),
            ativos=("Ativos no +TOP", "sum"),
            inativos=("Inativos no +TOP", "sum"),
        )
        .reset_index()
    )
    resumo_reg["% Ativos"] = (resumo_reg["ativos"] / resumo_reg["total"] * 100).round(1)
    resumo_reg = resumo_reg.rename(columns={
        "Regional": "Regional",
        "total": "Total participantes",
        "ativos": "Ativos no +TOP",
        "inativos": "Inativos no +TOP",
    })
    resumo_reg = resumo_reg[[
        "Regional", "Total participantes", "Ativos no +TOP",
        "Inativos no +TOP", "% Ativos"
    ]]
    resumo_reg = resumo_reg.sort_values("% Ativos", ascending=False)

    # Normaliza nomes das regionais (remove "CONTA " e aplica Title Case)
    resumo_reg["Regional"] = resumo_reg["Regional"].apply(regional_curta).apply(regional_title_case)
    resumo_rev["Regional"] = resumo_rev["Regional"].apply(regional_curta).apply(regional_title_case)

    return resumo_reg, resumo_rev


def main():
    logger.info("Iniciando geração de status hierarquia 60 dias...")
    logger.info(f"Data de corte: {DATA_CORTE:%d/%m/%Y}")
    logger.info("Bases utilizadas: hierarquias + cadastro + banco PROD_WHP")

    arquivos = descobrir_arquivos_hierarquia()
    logger.info(f"{len(arquivos)} arquivos de hierarquia selecionados")

    df_cad = carregar_cadastro()
    df_banco = carregar_banco()

    todos_rev = []

    for arquivo in arquivos:
        df_hier = carregar_hierarquia(arquivo)
        if df_hier is None or df_hier.empty:
            continue

        revenda, df_rev = processar_revenda(df_hier, df_cad, df_banco)
        todos_rev.append(df_rev)

        logger.info(
            f"  {revenda}: {len(df_rev)} registros (bruto) | "
            f"na base cadastro: {(df_rev['na_base_cadastro'] == 'Sim').sum()} | "
            f"na base banco: {(df_rev['na_base_banco'] == 'Sim').sum()} | "
            f"passou 60 dias: {(df_rev['passou_60_dias'] == 'Sim').sum()}"
        )

    if not todos_rev:
        raise ValueError("Nenhuma revenda foi processada")

    # Consolida todos os CPFs de todas as revendas e aplica deduplicação global,
    # mantendo a primeira ocorrência. A ordem dos arquivos já é alfabética
    # (descobrir_arquivos_hierarquia retorna sorted), o que garante que um CPF
    # que aparece em mais de uma hierarquia fique na primeira revenda — mesmo
    # comportamento usado no envio_relatorio/email semanal.
    #
    # IMPORTANTE: filtra DESLIGADO=NÃO *antes* da dedupe, igual ao email. CPFs
    # marcados como SIM, FÉRIAS, BENEFÍCIO ou em branco são descartados antes de
    # definir a revenda "dona" do CPF. Sem esse filtro, um CPF que está SIM na
    # primeira revenda e NÃO numa revenda posterior acaba ficando na primeira
    # aba como desligado, diminuindo o total da regional correta.
    df_consolidado = pd.concat(todos_rev, ignore_index=True)

    antes_filtro = len(df_consolidado)
    deslig_norm = df_consolidado["desligado"].astype(str).str.strip().str.upper().str.replace("Ã", "A")
    df_consolidado = df_consolidado[deslig_norm.isin(["NAO", "NÃO"])].copy()
    logger.info(
        f"Filtro DESLIGADO=NÃO (antes da deduplicação): "
        f"{antes_filtro - len(df_consolidado)} registros removidos"
    )

    antes = len(df_consolidado)
    df_consolidado = df_consolidado.drop_duplicates(subset=["cpf_limp"], keep="first")
    removidos = antes - len(df_consolidado)
    logger.info(
        f"Deduplicação global por CPF: {removidos} registros removidos — "
        f"{len(df_consolidado)} CPFs únicos restantes"
    )

    abas_por_revenda = {}
    for revenda_arquivo, df_rev in df_consolidado.groupby("revenda_arquivo"):
        nome_aba = revenda_arquivo[:31]
        contador = 1
        nome_base = nome_aba
        while nome_aba in abas_por_revenda:
            sufixo = f"_{contador}"
            nome_aba = nome_base[:31 - len(sufixo)] + sufixo
            contador += 1

        # Remove a coluna auxiliar antes de salvar
        df_rev = df_rev.drop(columns=["revenda_arquivo"], errors="ignore")
        abas_por_revenda[nome_aba] = df_rev

    if not abas_por_revenda:
        raise ValueError("Nenhuma revenda foi processada")

    # Cria mapa de revenda -> regional a partir do cadastro
    mapa_regional = mapeamento_revenda_regional(df_cad)
    resumo_regional, resumo_revenda = calcular_resumos(df_consolidado, mapa_regional)
    logger.info(f"Resumo por regional calculado: {len(resumo_regional)} regionais")

    logger.info(f"Salvando arquivo em: {SAIDA_FILE.name}")
    with pd.ExcelWriter(SAIDA_FILE, engine="openpyxl") as writer:
        # Abas de resumo primeiro
        resumo_regional.to_excel(writer, sheet_name="Resumo_Regional", index=False)
        resumo_revenda.to_excel(writer, sheet_name="Resumo_Revenda", index=False)

        # Depois as abas por revenda
        for nome_aba, df_rev in abas_por_revenda.items():
            df_rev.to_excel(writer, sheet_name=nome_aba, index=False)

    print("\n" + "=" * 80)
    print("ARQUIVO GERADO")
    print("=" * 80)
    print(f"Arquivo: {SAIDA_FILE}")
    print(f"Data de corte: {DATA_CORTE:%d/%m/%Y}")
    print(f"Total de abas (revendas): {len(abas_por_revenda)}")
    print(f"Abas de resumo: Resumo_Regional, Resumo_Revenda")
    print("-" * 80)
    print("Resumo por Regional:")
    print(resumo_regional.to_string(index=False))
    print("=" * 80)


if __name__ == "__main__":
    main()
