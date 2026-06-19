#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório de Pré-Cadastro por Revenda
-------------------------------------
Consulta o arquivo do banco PROD_WHP_Participante_Banco e traz participantes
com data de inclusão entre 01/05/2026 e a data de referência (ontem),
separados por revenda:
  - MM Varejo + MM Atacado (arquivo único)
  - Guaibim (arquivo separado)

Inclui:
  • Dados do participante (CPF, nome, status, data de inclusão)
  • Flag de pré-cadastro (< 90 dias de inclusão)
  • Aceite mensal mais recente disponível
  • Treinamentos realizados
  • CNPJ e nome da loja
  • Vendas (último mês disponível na base de vendas)
  • Aba resumo com quantidade por status
"""

import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES
# ---------------------------------------------------------------------------
DATA_INICIO = pd.Timestamp("2026-05-01")
DATA_FIM = pd.Timestamp("2026-06-16")  # ontem em relação a 17/06/2026
PRAZO_PRE_CADASTRO_DIAS = 90

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
OUTPUT_DIR.mkdir(exist_ok=True)

BANCO_PATH = Path("/home/thamiresvieira/projetos/validacoes_precadastro/PROD_WHP_Participante_Banco_v2_16062026.xlsx")
CADASTRO_PATH = BASE_DIR / "bases" / "cadastro.xlsx"
CADASTRO_REV_LOJA_PATH = BASE_DIR / "bases" / "cadastro_revenda_loja_regional.xlsx"
DIST_PDV_PATH = Path("/home/thamiresvieira/projetos/Programa_mais_top/bases/Tabelas/WHP_Distribuidor_x_PDV.xlsx")
ACEITES_PATH = BASE_DIR / "bases" / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
TREINAMENTOS_PATH = BASE_DIR / "bases" / "Base_treinamentos.xlsx"
VENDAS_PATH = Path("/home/thamiresvieira/projetos/Programa_mais_top/bases/Vendas Processadas/2026_04.xlsx")

REVENDAS_MM = {"MM Varejo", "MM Atacado"}
REVENDAS_GUAIBIM = {"Guaibim"}


def normalizar_cpf(cpf):
    """Normaliza CPF para texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    s = str(cpf).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(11)


def normalizar_cnpj(cnpj):
    """Normaliza CNPJ para texto com 14 dígitos."""
    if pd.isna(cnpj):
        return None
    s = str(cnpj).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(14)


def carregar_banco():
    """Carrega base do banco e normaliza CPF."""
    df = pd.read_excel(BANCO_PATH)
    df["cpf_str"] = df["Cpf"].apply(normalizar_cpf)
    df["DataInclusao"] = pd.to_datetime(df["DataInclusao"], errors="coerce")
    df["Ativo"] = df["Ativo"].map({1: "Sim", 0: "Não"})
    return df


def carregar_cadastro():
    """Carrega cadastro principal e normaliza CPF."""
    df = pd.read_excel(CADASTRO_PATH, dtype={"cpf/cnpj": str})
    df["cpf_str"] = df["cpf/cnpj"].astype(str).str.strip().str.replace("'", "", regex=False).str.replace(r"[^0-9]", "", regex=True).str.zfill(11)
    return df


def carregar_cadastro_revenda_loja():
    """Carrega mapeamento CPF -> loja/CNPJ."""
    df = pd.read_excel(CADASTRO_REV_LOJA_PATH)
    df["cpf_str"] = df["Cpf"].apply(normalizar_cpf)
    df["cnpj_loja_str"] = df["CNPJ_Loja"].apply(normalizar_cnpj)
    # Pega o primeiro CNPJ/Loja encontrado por CPF
    return df.drop_duplicates(subset=["cpf_str"], keep="first")[["cpf_str", "CNPJ_Loja", "Loja", "cnpj_loja_str"]]


def carregar_dist_pdv():
    """Carrega mapeamento CNPJ loja -> nome da loja."""
    df = pd.read_excel(DIST_PDV_PATH)
    df["cnpj_loja_str"] = df["CNPJ_LOJA"].apply(normalizar_cnpj)
    return df.drop_duplicates(subset=["cnpj_loja_str"], keep="first")[["cnpj_loja_str", "NOME_LOJA"]]


def carregar_aceites():
    """Carrega aceites mensais e identifica o mês mais recente disponível."""
    xl = pd.ExcelFile(ACEITES_PATH)
    abas = xl.sheet_names
    # Ordem natural das abas já está cronológica no arquivo
    ultima_aba = abas[-1]
    df = pd.read_excel(ACEITES_PATH, sheet_name=ultima_aba)
    df["cpf_str"] = df["CPF"].apply(normalizar_cpf)
    df["Aceite_Mensal"] = "Sim"
    df["Mes_Aceite_Mensal"] = ultima_aba
    return df[["cpf_str", "Aceite_Mensal", "Mes_Aceite_Mensal", "DataAceite"]].drop_duplicates(subset=["cpf_str"], keep="first")


def carregar_treinamentos():
    """Carrega treinamentos e agrega por CPF."""
    df = pd.read_excel(TREINAMENTOS_PATH)
    df["cpf_str"] = df["CPF"].apply(normalizar_cpf)
    # Quantidade de treinamentos concluídos por CPF
    concluidos = df[df["Estado"].astype(str).str.lower() == "concluido"].groupby("cpf_str").agg(
        Qtd_Treinamentos=("Curso", "nunique"),
        Treinamentos_Realizados=("Curso", lambda x: "; ".join(sorted(x.unique()))),
        Ultimo_Treinamento=("Conclusão", "max"),
    ).reset_index()
    concluidos["Realizou_Treinamento"] = "Sim"
    return concluidos


def carregar_vendas():
    """Carrega vendas do último mês disponível e agrega por CPF."""
    df = pd.read_excel(VENDAS_PATH)
    df["cpf_str"] = df["CPF"].apply(normalizar_cpf)
    vendas = df.groupby("cpf_str").agg(
        Qtd_Vendas=("Vendas", "sum"),
        Total_Pontuacao=("Pontuação", "sum"),
        Ultima_Venda=("data venda", "max"),
    ).reset_index()
    vendas["Teve_Venda"] = "Sim"
    return vendas


def construir_relatorio(df_base, nome_arquivo, titulo_aba):
    """Gera o arquivo Excel com detalhamento e resumo."""
    output_path = OUTPUT_DIR / nome_arquivo

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Aba detalhamento
        df_base.to_excel(writer, sheet_name="Detalhamento", index=False)

        # Aba resumo
        resumo = df_base["Status"].value_counts().reset_index()
        resumo.columns = ["Status", "Quantidade"]
        resumo.loc[len(resumo)] = ["TOTAL GERAL", df_base.shape[0]]

        resumo_adicional = pd.DataFrame({
            "Indicador": [
                "Total de participantes",
                "Ativos",
                "Pré-Cadastrados",
                "Pré-Cadastro (< 90 dias)",
                "Aceite Mensal (mês mais recente)",
                "Realizou Treinamento",
                "Teve Venda (2026-04)",
            ],
            "Quantidade": [
                df_base.shape[0],
                (df_base["Status"] == "Ativo").sum(),
                (df_base["Status"] == "Pré-Cadastrado").sum(),
                (df_base["Pre_Cadastro_90_dias"] == "Sim").sum(),
                (df_base["Aceite_Mensal"] == "Sim").sum(),
                (df_base["Realizou_Treinamento"] == "Sim").sum(),
                (df_base["Teve_Venda"] == "Sim").sum(),
            ],
        })

        resumo_adicional.to_excel(writer, sheet_name="Resumo", index=False)

    print(f"✅ Relatório salvo: {output_path}")
    print(f"   Total de participantes: {df_base.shape[0]}")
    print(resumo.to_string(index=False))
    print()


def main():
    print("=" * 70)
    print("Relatório de Pré-Cadastro - MM Varejo/Atacado e Guaibim")
    print(f"Período de inclusão: {DATA_INICIO.date()} a {DATA_FIM.date()}")
    print("=" * 70)
    print()

    # Carrega bases
    df_banco = carregar_banco()
    df_cadastro = carregar_cadastro()
    df_rev_loja = carregar_cadastro_revenda_loja()
    df_dist_pdv = carregar_dist_pdv()
    df_aceites = carregar_aceites()
    df_treinamentos = carregar_treinamentos()
    df_vendas = carregar_vendas()

    # Cruza banco com cadastro para obter grupo/status
    print(f"   df_banco cpf_str dtype: {df_banco['cpf_str'].dtype}")
    print(f"   df_cadastro cpf_str dtype: {df_cadastro['cpf_str'].dtype}")
    print(f"   df_cadastro grupo dtype: {df_cadastro['grupo'].dtype}")
    print(f"   df_banco cpf_str amostra: {df_banco['cpf_str'].head(3).tolist()}")
    print(f"   df_cadastro cpf_str amostra: {df_cadastro['cpf_str'].head(3).tolist()}")
    print(f"   Interseção CPFs: {len(set(df_banco['cpf_str'].dropna()).intersection(set(df_cadastro['cpf_str'].dropna())))}")
    df = df_banco.merge(
        df_cadastro[["cpf_str", "status", "grupo", "regional", "cnpj grupo"]],
        on="cpf_str",
        how="left",
    )
    print(f"   Após merge banco+cadastro: {len(df)} registros")
    print(f"   status notna: {df['status'].notna().sum()}")
    print(f"   grupo notna: {df['grupo'].notna().sum()}")

    # Filtra por data de inclusão
    print(f"   Antes do filtro de data: {len(df)} registros")
    print(f"   DataInclusao dtype: {df['DataInclusao'].dtype}")
    filtro_data = (df["DataInclusao"] >= DATA_INICIO) & (df["DataInclusao"] <= DATA_FIM)
    print(f"   Filtro de data: {filtro_data.sum()} registros")
    df = df[filtro_data].copy()
    print(f"   Após filtro de data: {len(df)} registros")
    print(f"   Grupos no filtro: {df['grupo'].value_counts().head(10).to_dict()}")

    if df.empty:
        print("Nenhum participante encontrado no período e revendas especificadas.")
        return

    # Calcula dias de inclusão e flag de pré-cadastro
    df["Dias_Desde_Inclusao"] = (DATA_FIM - df["DataInclusao"]).dt.days
    df["Pre_Cadastro_90_dias"] = df["Dias_Desde_Inclusao"].apply(lambda x: "Sim" if x < PRAZO_PRE_CADASTRO_DIAS else "Não")

    # Adiciona loja/CNPJ
    df = df.merge(df_rev_loja, on="cpf_str", how="left")
    df = df.merge(df_dist_pdv, on="cnpj_loja_str", how="left")

    # Adiciona aceites
    df = df.merge(df_aceites, on="cpf_str", how="left")
    df["Aceite_Mensal"] = df["Aceite_Mensal"].fillna("Não")
    df["Mes_Aceite_Mensal"] = df["Mes_Aceite_Mensal"].fillna("-")
    df["DataAceite"] = pd.to_datetime(df["DataAceite"], errors="coerce")

    # Adiciona treinamentos
    df = df.merge(df_treinamentos, on="cpf_str", how="left")
    df["Realizou_Treinamento"] = df["Realizou_Treinamento"].fillna("Não")
    df["Qtd_Treinamentos"] = df["Qtd_Treinamentos"].fillna(0).astype(int)
    df["Treinamentos_Realizados"] = df["Treinamentos_Realizados"].fillna("-")

    # Adiciona vendas
    df = df.merge(df_vendas, on="cpf_str", how="left")
    df["Teve_Venda"] = df["Teve_Venda"].fillna("Não")
    df["Qtd_Vendas"] = df["Qtd_Vendas"].fillna(0)
    df["Total_Pontuacao"] = df["Total_Pontuacao"].fillna(0)

    # Renomeia colunas para exibição
    df = df.rename(columns={
        "Cpf": "CPF",
        "Nome": "Nome_Participante",
        "Ativo": "Ativo_Banco",
        "status": "Status",
        "grupo": "Revenda",
        "regional": "Regional",
        "DataInclusao": "Data_Inclusao",
        "DataAceiteRegulamento": "Data_Aceite_Regulamento",
        "DataAceiteLGPD": "Data_Aceite_LGPD",
        "cnpj grupo": "CNPJ_Grupo",
        "Loja": "Loja_Cadastro",
        "NOME_LOJA": "Nome_Loja",
        "DataAceite": "Data_Aceite_Mensal",
    })

    # Seleciona e ordena colunas finais
    colunas_finais = [
        "CPF",
        "Nome_Participante",
        "Status",
        "Ativo_Banco",
        "Revenda",
        "Regional",
        "Data_Inclusao",
        "Dias_Desde_Inclusao",
        "Pre_Cadastro_90_dias",
        "CNPJ_Grupo",
        "CNPJ_Loja",
        "Loja_Cadastro",
        "Nome_Loja",
        "Data_Aceite_Regulamento",
        "Data_Aceite_LGPD",
        "Aceite_Mensal",
        "Mes_Aceite_Mensal",
        "Data_Aceite_Mensal",
        "Realizou_Treinamento",
        "Qtd_Treinamentos",
        "Treinamentos_Realizados",
        "Ultimo_Treinamento",
        "Teve_Venda",
        "Qtd_Vendas",
        "Total_Pontuacao",
        "Ultima_Venda",
    ]

    # Garante que todas as colunas existam
    for col in colunas_finais:
        if col not in df.columns:
            df[col] = None

    df = df[colunas_finais].sort_values(["Revenda", "Status", "Nome_Participante"])

    # Separa em dois relatórios
    df_mm = df[df["Revenda"].isin(REVENDAS_MM)].copy()
    df_guaibim = df[df["Revenda"].isin(REVENDAS_GUAIBIM)].copy()

    print("--- MM Varejo + MM Atacado ---")
    if not df_mm.empty:
        construir_relatorio(df_mm, "relatorio_precadastro_MM_20260617.xlsx", "MM Varejo/Atacado")
    else:
        print("Nenhum participante encontrado para MM Varejo/Atacado no período.")

    print("--- Guaibim ---")
    if not df_guaibim.empty:
        construir_relatorio(df_guaibim, "relatorio_precadastro_Guaibim_20260617.xlsx", "Guaibim")
    else:
        print("Nenhum participante encontrado para Guaibim no período.")


if __name__ == "__main__":
    main()
