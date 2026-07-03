#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_validacao_todas_revendas.py

Gera a base de validação CPF cadastro x hierarquia para TODAS as revendas,
utilizando os arquivos de hierarquia da pasta bases_cadastro_hierarquia.

A estrutura de saída (abas e colunas) é idêntica à do arquivo
validacao_cpf_cadastro_x_hierarquia.xlsx, mas consolidando todas as revendas:
    - Resumo Geral
    - Por Status
    - Em Ambos
    - Só na Plataforma
    - Só na Hierarquia
    - Cruzamento Completo

Regras:
- Usa preferencialmente arquivos *_corrigido.xlsx; ignora o original correspondente.
- Considera apenas CPFs com DESLIGADO = NÃO (quando houver coluna).
- Cruza com cadastro.xlsx por CPF (texto, 11 dígitos).
- CDA no cadastro é mapeado para Casas da Água.
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
HIERARQUIA_DIR = DATA_DIR / "bases_cadastro_hierarquia"
SAIDA_DIR = BASE_DIR / "relatorios_gerados"
SAIDA_DIR.mkdir(exist_ok=True)

CADASTRO_FILE = DATA_DIR / "cadastro.xlsx"
EMAIL_FILE = SAIDA_DIR / "email_validacao_todas_revendas.txt"
VALIDACAO_FILE = SAIDA_DIR / "validacao_cpf_cadastro_x_hierarquia_todas_revendas.xlsx"

# Mapeamento de nomes de arquivo de hierarquia -> grupo no cadastro
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

# Colunas finais esperadas no arquivo de validação (mesma ordem do original)
COLUNAS_FINAIS = [
    "nome", "cpf/cnpj", "matricula", "cargo", "status", "regional", "grupo",
    "cnpj grupo", "rua", "bairro", "cidade", "uf", "cep", "telefone",
    "celular", "email", "data de aceite", "lgpd", "ultimo aceite mensal",
    "data atualização", "cpf_limp",
    "revenda_hierarquia", "cod_loja_hierarquia", "cnpj_hierarquia",
    "nome_hierarquia", "vendedor_hierarquia", "gerente_loja_hierarquia",
    "gerente_regional_hierarquia", "desligado_hierarquia",
    "_merge", "cpf_formatado", "na_plataforma", "na_hierarquia",
]


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


def descobrir_arquivos_hierarquia():
    """Retorna lista de arquivos de hierarquia a processar (preferência corrigidos)."""
    todos = sorted(HIERARQUIA_DIR.glob("*.xlsx"))

    # Apenas arquivos individuais de revenda (evita consolidados, resumos, comparacoes etc.)
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

    # Preferência por arquivos _corrigido
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
            # Só adiciona o original se não existir corrigido equivalente
            if base not in bases_corrigidas:
                selecionados.append(f)

    return sorted(selecionados)


def extrair_nome_revenda(nome_arquivo):
    """Extrai o nome da revenda a partir do nome do arquivo."""
    nome = nome_arquivo.replace(".xlsx", "").strip().upper()
    # Remove sufixos
    nome = re.sub(r"_CORRIGIDO$", "", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\s*\(\d+\)$", "", nome)
    nome = nome.strip()

    # Tenta match exato no mapeamento
    for chave, valor in MAPEAMENTO_REVENDA.items():
        if nome == chave:
            return valor

    # Match parcial
    for chave, valor in MAPEAMENTO_REVENDA.items():
        if chave in nome or nome in chave:
            return valor

    # Fallback: retorna o nome do arquivo capitalizado
    return nome.title()


def carregar_hierarquia(arquivo):
    """Carrega um arquivo de hierarquia individual."""
    logger.info(f"Lendo hierarquia: {arquivo.name}")
    try:
        df = pd.read_excel(arquivo, dtype=str)
    except Exception as e:
        logger.warning(f"Erro ao ler {arquivo.name}: {e}")
        return None

    df.columns = [str(c).strip().upper() for c in df.columns]

    # Renomeia colunas conhecidas
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

    # Garante colunas mínimas
    for col in ["revenda", "cod_loja", "cnpj", "cpf", "nome", "desligado"]:
        if col not in df.columns:
            df[col] = None

    df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()

    # Normaliza desligado
    df["desligado"] = df["desligado"].fillna("").astype(str).str.strip().str.upper()

    # Se tiver coluna DESLIGADO, filtra apenas NÃO
    tem_desligado = "desligado" in df.columns and df["desligado"].notna().any()
    if tem_desligado:
        antes = len(df)
        df = df[df["desligado"].isin(["NÃO", "NAO", "N"])].copy()
        depois = len(df)
        logger.info(f"  Filtro DESLIGADO=NÃO: {antes} -> {depois}")

    if df.empty:
        return None

    df["arquivo_origem"] = arquivo.name
    df["revenda_arquivo"] = extrair_nome_revenda(arquivo.name)

    return df


def carregar_cadastro():
    """Carrega o cadastro completo da plataforma."""
    logger.info(f"Lendo cadastro: {CADASTRO_FILE.name}")
    df = pd.read_excel(CADASTRO_FILE, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    df["cpf_limp"] = df["cpf/cnpj"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()
    df["status"] = df["status"].fillna("").astype(str).str.strip().str.title()
    df["grupo"] = df["grupo"].fillna("").astype(str).str.strip()
    # Mapeia CDA -> Casas da Água para cruzamento com hierarquia
    df["grupo_mapeado"] = df["grupo"].replace({"CDA": "Casas da Água"})
    logger.info(f"Cadastro: {len(df):,} registros, {df['cpf_limp'].nunique():,} CPFs únicos")
    return df


def criar_cruzamento_revenda(df_hier, df_cad):
    """
    Cria o DataFrame de cruzamento para uma revenda com a mesma estrutura
    de colunas do arquivo validacao_cpf_cadastro_x_hierarquia.xlsx.
    """
    revenda = df_hier["revenda_arquivo"].iloc[0]
    cpfs_hier = set(df_hier["cpf_limp"].unique())

    # Cadastros da mesma revenda no cadastro.xlsx (usando grupo mapeado)
    df_cad_rev = df_cad[df_cad["grupo_mapeado"] == revenda].copy()

    # Prepara hierarquia para merge (colunas com sufixo _hierarquia)
    cols_hier = ["cpf_limp", "cod_loja", "cnpj", "nome", "vendedor",
                 "gerente_loja", "gerente_regional", "desligado"]
    cols_hier_presentes = [c for c in cols_hier if c in df_hier.columns]
    df_hier_merge = df_hier[cols_hier_presentes].rename(columns={
        "nome": "nome_hierarquia",
        "cod_loja": "cod_loja_hierarquia",
        "cnpj": "cnpj_hierarquia",
        "vendedor": "vendedor_hierarquia",
        "gerente_loja": "gerente_loja_hierarquia",
        "gerente_regional": "gerente_regional_hierarquia",
        "desligado": "desligado_hierarquia",
    })
    # Adiciona revenda_hierarquia
    df_hier_merge["revenda_hierarquia"] = revenda

    # Merge outer
    df_cruzamento = df_cad_rev.merge(
        df_hier_merge,
        on="cpf_limp",
        how="outer",
        indicator=True,
    )

    # Flags
    df_cruzamento["na_plataforma"] = df_cruzamento["_merge"].isin(["both", "left_only"])
    df_cruzamento["na_hierarquia"] = df_cruzamento["_merge"].isin(["both", "right_only"])
    df_cruzamento["cpf_formatado"] = df_cruzamento["cpf_limp"].apply(formatar_cpf)

    # Garante todas as colunas finais
    for col in COLUNAS_FINAIS:
        if col not in df_cruzamento.columns:
            df_cruzamento[col] = None

    return df_cruzamento[COLUNAS_FINAIS].copy()


def calcular_resumo_revenda(df_hier, df_cad):
    """Calcula métricas resumidas por revenda."""
    revenda = df_hier["revenda_arquivo"].iloc[0]
    cpfs_hier = set(df_hier["cpf_limp"].unique())
    total_hierarquia = len(cpfs_hier)

    df_cad_rev = df_cad[df_cad["grupo_mapeado"] == revenda].copy()
    total_cadastro = df_cad_rev["cpf_limp"].nunique()
    df_cad_rev["na_hierarquia"] = df_cad_rev["cpf_limp"].isin(cpfs_hier)

    cadastro_na_hierarquia = df_cad_rev[df_cad_rev["na_hierarquia"]]["cpf_limp"].nunique()
    cadastro_fora_hierarquia = total_cadastro - cadastro_na_hierarquia
    hierarquia_fora_plataforma = total_hierarquia - cadastro_na_hierarquia

    def contar(status, na_hier):
        return int(((df_cad_rev["status"] == status) & (df_cad_rev["na_hierarquia"] == na_hier)).sum())

    pct_ativos = (contar("Ativo", True) / total_hierarquia * 100) if total_hierarquia else 0

    return {
        "revenda": revenda,
        "arquivo": df_hier["arquivo_origem"].iloc[0],
        "total_hierarquia": total_hierarquia,
        "total_cadastro": total_cadastro,
        "cadastro_na_hierarquia": cadastro_na_hierarquia,
        "cadastro_fora_hierarquia": cadastro_fora_hierarquia,
        "hierarquia_fora_plataforma": hierarquia_fora_plataforma,
        "fora_ativos": contar("Ativo", False),
        "fora_inativos": contar("Inativo", False),
        "fora_reprovados": contar("Reprovado", False),
        "dentro_ativos": contar("Ativo", True),
        "dentro_inativos": contar("Inativo", True),
        "dentro_reprovados": contar("Reprovado", True),
        "pct_ativos": round(pct_ativos, 1),
    }


def gerar_resumo_geral(df_resumo):
    """Gera DataFrame Resumo Geral consolidado."""
    return pd.DataFrame({
        "Indicador": [
            "Total de revendas processadas",
            "Total na hierarquia (soma)",
            "Total cadastrados na plataforma (soma)",
            "Cadastrados na plataforma e na hierarquia (soma)",
            "Cadastrados na plataforma e FORA da hierarquia (soma)",
            "Na hierarquia e FORA da plataforma (soma)",
        ],
        "Quantidade": [
            len(df_resumo),
            df_resumo["total_hierarquia"].sum(),
            df_resumo["total_cadastro"].sum(),
            df_resumo["cadastro_na_hierarquia"].sum(),
            df_resumo["cadastro_fora_hierarquia"].sum(),
            df_resumo["hierarquia_fora_plataforma"].sum(),
        ],
    })


def gerar_por_status(df_cruzamento_total):
    """Gera DataFrame Por Status consolidado."""
    df_cad = df_cruzamento_total[df_cruzamento_total["_merge"].isin(["both", "left_only"])].copy()
    df_cad["na_hierarquia"] = df_cad["_merge"] == "both"
    tabela = pd.crosstab(
        df_cad["status"],
        df_cad["na_hierarquia"].map({True: "Faz parte da hierarquia", False: "Não faz parte da hierarquia"}),
        margins=True,
    ).reset_index()
    return tabela


def gerar_email(resumo_geral, data_corte=None):
    """Gera email/resumo consolidado com as métricas de todas as revendas."""
    if data_corte is None:
        data_corte = date.today().strftime("%d/%m/%Y")

    linhas = []
    for _, row in resumo_geral.iterrows():
        linhas.append(
            f"{row['revenda']:<20} | Hier: {row['total_hierarquia']:>5} | Cad: {row['total_cadastro']:>5} | "
            f"Cad na Hier: {row['cadastro_na_hierarquia']:>5} | Ativos na Hier: {row['dentro_ativos']:>5} | "
            f"% Ativos: {row['pct_ativos']:>5}%"
        )

    email = f"""Prezados,

Segue a validação de CPF cadastro x hierarquia para TODAS as revendas do Programa +TOP.

**Data de corte da análise:** {data_corte}
**Base de hierarquia utilizada:** arquivos de HIERARQUIA MAIO (DESLIGADO = NÃO, quando aplicável)

---

**Resumo por revenda:**

{'Revenda':<20} | {'Hier':>5} | {'Cad':>5} | {'Cad na Hier':>11} | {'Ativos na Hier':>14} | {'% Ativos':>8}
{'-' * 90}
"""
    email += "\n".join(linhas)
    email += f"""

---

**Destaques — Guaibim (referência):**
- Total na hierarquia: {resumo_geral[resumo_geral['revenda'] == 'Guaibim']['total_hierarquia'].values[0]}
- Cadastrados na plataforma: {resumo_geral[resumo_geral['revenda'] == 'Guaibim']['total_cadastro'].values[0]}
- Cadastrados na hierarquia: {resumo_geral[resumo_geral['revenda'] == 'Guaibim']['cadastro_na_hierarquia'].values[0]}
- Ativos na hierarquia: {resumo_geral[resumo_geral['revenda'] == 'Guaibim']['dentro_ativos'].values[0]}
- % Ativos: {resumo_geral[resumo_geral['revenda'] == 'Guaibim']['pct_ativos'].values[0]}%

A base de validação completa está em anexo, com as abas: Resumo Geral, Por Status, Em Ambos, Só na Plataforma, Só na Hierarquia e Cruzamento Completo.

Atenciosamente,
"""

    EMAIL_FILE.write_text(email, encoding="utf-8")
    logger.info(f"Email salvo em: {EMAIL_FILE.name}")
    return email


def main():
    logger.info("Iniciando validação de todas as revendas...")

    arquivos = descobrir_arquivos_hierarquia()
    logger.info(f"{len(arquivos)} arquivos de hierarquia selecionados")

    df_cad = carregar_cadastro()

    resumos = []
    cruzamentos = []
    cruzamentos_por_revenda = {}

    for arquivo in arquivos:
        df_hier = carregar_hierarquia(arquivo)
        if df_hier is None or df_hier.empty:
            continue

        df_cruz = criar_cruzamento_revenda(df_hier, df_cad)
        cruzamentos.append(df_cruz)

        resumo = calcular_resumo_revenda(df_hier, df_cad)
        resumos.append(resumo)

        # Guarda cruzamento completo por revenda para aba individual
        nome_aba = resumo["revenda"][:31]
        cruzamentos_por_revenda[nome_aba] = df_cruz.copy()

        logger.info(
            f"  {resumo['revenda']}: hier={resumo['total_hierarquia']}, "
            f"cad={resumo['total_cadastro']}, cad_na_hier={resumo['cadastro_na_hierarquia']}, "
            f"ativos={resumo['dentro_ativos']} ({resumo['pct_ativos']}%)"
        )

    if not resumos:
        raise ValueError("Nenhuma revenda foi processada")

    df_resumo = pd.DataFrame(resumos)
    df_resumo = df_resumo.sort_values("total_hierarquia", ascending=False)

    # Consolida cruzamentos
    df_cruzamento_total = pd.concat(cruzamentos, ignore_index=True)

    # Separa grupos
    df_ambos = df_cruzamento_total[df_cruzamento_total["_merge"] == "both"].copy()
    df_somente_cadastro = df_cruzamento_total[df_cruzamento_total["_merge"] == "left_only"].copy()
    df_somente_hierarquia = df_cruzamento_total[df_cruzamento_total["_merge"] == "right_only"].copy()

    # Resumos
    df_resumo_geral = gerar_resumo_geral(df_resumo)
    df_por_status = gerar_por_status(df_cruzamento_total)

    # Salva Excel
    logger.info(f"Salvando validação em: {VALIDACAO_FILE.name}")
    with pd.ExcelWriter(VALIDACAO_FILE, engine="openpyxl") as writer:
        df_resumo_geral.to_excel(writer, sheet_name="Resumo Geral", index=False)
        df_por_status.to_excel(writer, sheet_name="Por Status", index=False)
        df_ambos.to_excel(writer, sheet_name="Em Ambos", index=False)
        df_somente_cadastro.to_excel(writer, sheet_name="Só na Plataforma", index=False)
        df_somente_hierarquia.to_excel(writer, sheet_name="Só na Hierarquia", index=False)
        df_cruzamento_total.to_excel(writer, sheet_name="Cruzamento Completo", index=False)

        # Uma aba por revenda (cruzamento completo)
        for nome_aba, df_rev in cruzamentos_por_revenda.items():
            df_rev.to_excel(writer, sheet_name=nome_aba, index=False)
            logger.info(f"  Aba gerada: {nome_aba} ({len(df_rev)} registros)")

    email = gerar_email(df_resumo)

    print("\n" + "=" * 90)
    print("VALIDAÇÃO DE TODAS AS REVENDAS GERADA")
    print("=" * 90)
    print(f"Revendas processadas: {len(df_resumo)}")
    print(f"\n{df_resumo.to_string(index=False)}")
    print(f"\nAbas geradas: Resumo Geral, Por Status, Em Ambos, Só na Plataforma, Só na Hierarquia, Cruzamento Completo")
    print("\n" + "-" * 90)
    print(email)
    print("-" * 90)


if __name__ == "__main__":
    main()
