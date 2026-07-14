#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_validacao_guaibim_email.py

Gera a base de validação CPF cadastro x hierarquia para a revenda Guaibim,
utilizada como referência para validar as métricas do relatório +TOP.

Regras aplicadas:
- Hierarquia da Guaibim (MAIO) com DESLIGADO = NÃO.
- Cadastro da plataforma +TOP (cadastro.xlsx) filtrado por grupo Guaibim.
- Cruzamento único por CPF (texto, 11 dígitos).
- Separação entre participantes na plataforma que fazem / não fazem parte da hierarquia.
- Contagem por status (Ativo / Inativo / Reprovado).

Saídas:
1. Arquivo Excel de validação em relatorios_gerados/validacao_cpf_cadastro_x_hierarquia.xlsx
2. Cópia limpa da hierarquia Guaibim em bases/bases_cadastro_hierarquia/GUAIBIM_HIERARQUIA_LIMPA.xlsx
3. Arquivo de texto com o email formatado em relatorios_gerados/email_validacao_guaibim.txt
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

HIERARQUIA_FILE = HIERARQUIA_DIR / "GUAIBIM - HIERARQUIAS MAIO_corrigido.xlsx"
CADASTRO_FILE = DATA_DIR / "cadastro.xlsx"
EMAIL_FILE = SAIDA_DIR / "email_validacao_guaibim.txt"
VALIDACAO_FILE = SAIDA_DIR / "validacao_cpf_cadastro_x_hierarquia.xlsx"
LIMPA_FILE = HIERARQUIA_DIR / "GUAIBIM_HIERARQUIA_LIMPA.xlsx"


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


def carregar_hierarquia_limpo():
    """Carrega a hierarquia Guaibim já limpa (sem filtros/colunas ocultos)."""
    logger.info(f"Lendo hierarquia: {HIERARQUIA_FILE.name}")
    df = pd.read_excel(HIERARQUIA_FILE, dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Normaliza colunas
    df = df.rename(columns={
        "REVENDA": "revenda",
        "CODLOJA": "cod_loja",
        "CNPJ": "cnpj",
        "CPF": "cpf",
        "NOME": "nome",
        "VENDEDOR": "vendedor",
        "GERENTE DE LOJA": "gerente_loja",
        "GERENTE REGIONAL": "gerente_regional",
        "DIRETOR": "diretor",
        "DESLIGADO": "desligado",
    })

    df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()

    # Garante DESLIGADO preenchido
    df["desligado"] = df["desligado"].fillna("").astype(str).str.strip().str.upper()

    logger.info(f"Hierarquia: {len(df):,} registros, {df['cpf_limp'].nunique():,} CPFs únicos")
    return df


def carregar_cadastro_guaibim():
    """Carrega o cadastro da plataforma e filtra a revenda Guaibim."""
    logger.info(f"Lendo cadastro: {CADASTRO_FILE.name}")
    df = pd.read_excel(CADASTRO_FILE, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]

    df["cpf_limp"] = df["cpf/cnpj"].apply(limpar_cpf)
    df = df[df["cpf_limp"].notna()].copy()

    # Filtra Guaibim
    df_gua = df[df["grupo"].str.contains("GUAIBIM", case=False, na=False)].copy()

    # Normaliza status
    df_gua["status"] = df_gua["status"].fillna("").astype(str).str.strip().str.title()

    logger.info(f"Cadastro Guaibim: {len(df_gua):,} registros, {df_gua['cpf_limp'].nunique():,} CPFs únicos")
    return df_gua


def gerar_validacao(df_hier, df_cad):
    """Gera a base de validação cruzando cadastro e hierarquia."""
    cpfs_hier = set(df_hier["cpf_limp"].unique())

    # Adiciona flag na hierarquia
    df_hier = df_hier.copy()
    df_hier["na_plataforma"] = df_hier["cpf_limp"].isin(set(df_cad["cpf_limp"]))

    # Cruzamento completo: cadastro + dados da hierarquia
    df_cruzamento = df_cad.merge(
        df_hier[[
            "cpf_limp", "revenda", "cod_loja", "cnpj", "nome",
            "vendedor", "gerente_loja", "gerente_regional", "desligado"
        ]].rename(columns={
            "revenda": "revenda_hierarquia",
            "cod_loja": "cod_loja_hierarquia",
            "cnpj": "cnpj_hierarquia",
            "nome": "nome_hierarquia",
            "vendedor": "vendedor_hierarquia",
            "gerente_loja": "gerente_loja_hierarquia",
            "gerente_regional": "gerente_regional_hierarquia",
            "desligado": "desligado_hierarquia",
        }),
        on="cpf_limp",
        how="outer",
        indicator=True,
    )

    df_cruzamento["cpf_formatado"] = df_cruzamento["cpf_limp"].apply(formatar_cpf)
    df_cruzamento["na_plataforma"] = df_cruzamento["_merge"].isin(["both", "left_only"])
    df_cruzamento["na_hierarquia"] = df_cruzamento["_merge"].isin(["both", "right_only"])

    # Separa grupos
    df_ambos = df_cruzamento[df_cruzamento["_merge"] == "both"].copy()
    df_somente_cadastro = df_cruzamento[df_cruzamento["_merge"] == "left_only"].copy()
    df_somente_hierarquia = df_cruzamento[df_cruzamento["_merge"] == "right_only"].copy()

    # Resumo por status e presença na hierarquia
    resumo_status = df_cad.copy()
    resumo_status["na_hierarquia"] = resumo_status["cpf_limp"].isin(cpfs_hier)
    tabela_status = pd.crosstab(
        resumo_status["status"],
        resumo_status["na_hierarquia"].map({True: "Faz parte da hierarquia", False: "Não faz parte da hierarquia"}),
        margins=True,
    ).reset_index()

    # Resumo geral
    total_hierarquia = df_hier["cpf_limp"].nunique()
    total_cadastro = df_cad["cpf_limp"].nunique()
    cadastro_na_hierarquia = df_cad[df_cad["cpf_limp"].isin(cpfs_hier)]["cpf_limp"].nunique()
    cadastro_fora_hierarquia = total_cadastro - cadastro_na_hierarquia
    hierarquia_fora_plataforma = total_hierarquia - cadastro_na_hierarquia

    resumo_geral = pd.DataFrame({
        "Indicador": [
            "Total na hierarquia (DESLIGADO=NÃO)",
            "Total cadastrados na plataforma +TOP",
            "Cadastrados na plataforma e na hierarquia",
            "Cadastrados na plataforma e FORA da hierarquia",
            "Na hierarquia e FORA da plataforma",
        ],
        "Quantidade": [
            total_hierarquia,
            total_cadastro,
            cadastro_na_hierarquia,
            cadastro_fora_hierarquia,
            hierarquia_fora_plataforma,
        ],
    })

    return {
        "cruzamento_completo": df_cruzamento,
        "ambos": df_ambos,
        "somente_cadastro": df_somente_cadastro,
        "somente_hierarquia": df_somente_hierarquia,
        "tabela_status": tabela_status,
        "resumo_geral": resumo_geral,
        "total_hierarquia": total_hierarquia,
        "total_cadastro": total_cadastro,
        "cadastro_na_hierarquia": cadastro_na_hierarquia,
        "cadastro_fora_hierarquia": cadastro_fora_hierarquia,
        "hierarquia_fora_plataforma": hierarquia_fora_plataforma,
    }


def salvar_hierarquia_limpa(df_hier):
    """Salva uma cópia da hierarquia garantidamente limpa (sem filtros/colunas ocultos)."""
    logger.info(f"Salvando hierarquia limpa em: {LIMPA_FILE.name}")
    df_saida = df_hier[[
        "revenda", "cod_loja", "cnpj", "cpf", "nome",
        "vendedor", "gerente_loja", "gerente_regional", "desligado",
    ]].copy()
    df_saida.columns = [
        "REVENDA", "CODLOJA", "CNPJ", "CPF", "NOME",
        "VENDEDOR", "GERENTE DE LOJA", "GERENTE REGIONAL", "DESLIGADO",
    ]

    with pd.ExcelWriter(LIMPA_FILE, engine="openpyxl") as writer:
        df_saida.to_excel(writer, sheet_name="Hierarquia", index=False)
        ws = writer.sheets["Hierarquia"]
        # Remove autofiltros e congela painéis
        ws.auto_filter.ref = None
        ws.freeze_panes = None

    logger.info(f"Hierarquia limpa salva: {len(df_saida):,} registros")


def salvar_validacao(resultados):
    """Salva o arquivo Excel de validação."""
    logger.info(f"Salvando validação em: {VALIDACAO_FILE.name}")

    with pd.ExcelWriter(VALIDACAO_FILE, engine="openpyxl") as writer:
        resultados["resumo_geral"].to_excel(writer, sheet_name="Resumo Geral", index=False)
        resultados["tabela_status"].to_excel(writer, sheet_name="Por Status", index=False)
        resultados["ambos"].to_excel(writer, sheet_name="Em Ambos", index=False)
        resultados["somente_cadastro"].to_excel(writer, sheet_name="Só na Plataforma", index=False)
        resultados["somente_hierarquia"].to_excel(writer, sheet_name="Só na Hierarquia", index=False)
        resultados["cruzamento_completo"].to_excel(writer, sheet_name="Cruzamento Completo", index=False)

    logger.info("Validação salva com sucesso")


def gerar_email(resultados, data_corte=None):
    """Gera o texto do email com os cálculos da Guaibim."""
    if data_corte is None:
        data_corte = date.today().strftime("%d/%m/%Y")

    # Contagem detalhada por status
    df_cad = resultados["somente_cadastro"]
    df_ambos = resultados["ambos"]

    def contar_status(df, status):
        return int((df["status"].str.lower() == status.lower()).sum())

    fora_ativos = contar_status(df_cad, "Ativo")
    fora_inativos = contar_status(df_cad, "Inativo")
    fora_reprovados = contar_status(df_cad, "Reprovado")

    dentro_ativos = contar_status(df_ambos, "Ativo")
    dentro_inativos = contar_status(df_ambos, "Inativo")
    dentro_reprovados = contar_status(df_ambos, "Reprovado")

    email = f"""Prezados,

Segue a validação da base da revenda **Guaibim**, utilizada como referência para o relatório +TOP.

**Data de corte da análise:** {data_corte}
**Base de hierarquia utilizada:** GUAIBIM - HIERARQUIAS MAIO (DESLIGADO = NÃO)

---

Na hierarquia enviada pela Revenda, temos:
{resultados['total_hierarquia']:,} participantes

Na base de cadastro da plataforma do +TOP, temos:
{resultados['total_cadastro']:,} cadastros

Sendo que:
{resultados['cadastro_fora_hierarquia']:,} participantes que estão na plataforma, não fazem parte da hierarquia
{fora_ativos:,} ativos
{fora_inativos:,} inativos
{fora_reprovados:,} reprovados

{resultados['cadastro_na_hierarquia']:,} participantes que estão na plataforma, fazem parte da hierarquia
{dentro_ativos:,} ativos
{dentro_inativos:,} inativos
{dentro_reprovados:,} reprovados

---

**Conclusão:**
Apenas os participantes com status **Ativo** na plataforma +TOP e que também estão na hierarquia enviada pela revenda compõem a base ativa operacional para o indicador.

- Total na hierarquia: {resultados['total_hierarquia']:,}
- Ativos na plataforma + dentro da hierarquia: {dentro_ativos:,}
- Percentual de ativos sobre a hierarquia: {(dentro_ativos / resultados['total_hierarquia'] * 100):.1f}%

Desta forma, não podemos afirmar que a base de Guaibim está 100% ativa, uma vez que {resultados['hierarquia_fora_plataforma']:,} participantes da hierarquia não possuem cadastro ativo na plataforma +TOP.

Por favor, podem verificar.
Anexo os arquivos que analisei.

Atenciosamente,
"""

    EMAIL_FILE.write_text(email, encoding="utf-8")
    logger.info(f"Email salvo em: {EMAIL_FILE.name}")
    return email


def main():
    logger.info("Iniciando geração da validação Guaibim...")

    df_hier = carregar_hierarquia_limpo()
    df_cad = carregar_cadastro_guaibim()

    resultados = gerar_validacao(df_hier, df_cad)

    salvar_hierarquia_limpa(df_hier)
    salvar_validacao(resultados)
    email = gerar_email(resultados)

    print("\n" + "=" * 60)
    print("VALIDAÇÃO GUAIBIM GERADA")
    print("=" * 60)
    print(f"Hierarquia total:          {resultados['total_hierarquia']:,}")
    print(f"Cadastros plataforma:      {resultados['total_cadastro']:,}")
    print(f"Cadastros na hierarquia:   {resultados['cadastro_na_hierarquia']:,}")
    print(f"Cadastros fora hierarquia: {resultados['cadastro_fora_hierarquia']:,}")
    print(f"Hierarquia fora plataforma:{resultados['hierarquia_fora_plataforma']:,}")
    print(f"\nArquivos gerados:")
    print(f"  {VALIDACAO_FILE}")
    print(f"  {LIMPA_FILE}")
    print(f"  {EMAIL_FILE}")
    print("\n" + "-" * 60)
    print(email)
    print("-" * 60)


if __name__ == "__main__":
    main()
