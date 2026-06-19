#!/usr/bin/env python3
"""
Script de análise completa das bases do Programa +TOP
Extrai: colunas, tipos, amostras, estatísticas e relacionamentos
"""

import pandas as pd
import os
from pathlib import Path
import json

BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")
OUTPUT_FILE = BASE_DIR / "analise_bases_completa.json"

# Mapeamento das bases principais
BASES = {
    "cadastro": BASE_DIR / "bases/Tabelas/cadastro.xlsx",
    "acessos": BASE_DIR / "bases/Tabelas/acessos.xlsx",
    "pontos_por_cpf": BASE_DIR / "bases/Tabelas/pontos_por_cpf_.xlsx",
    "resgates_detalhado": BASE_DIR / "bases/Tabelas/resgates_detalhado.xlsx",
    "treinamentos": BASE_DIR / "bases/Tabelas/Base_treinamentos.xlsx",
    "aceites": BASE_DIR / "bases/Tabelas/WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx",
    "chamados": BASE_DIR / "bases/Tabelas/Chamados_catalogo.xlsx",
    "fale_conosco": BASE_DIR / "bases/Tabelas/Fale_Conosco.xlsx",
    "enquete": BASE_DIR / "bases/Tabelas/enquete.xlsx",
    "ocorrencias": BASE_DIR / "bases/Tabelas/OcorrênciasTOPSAC.xlsx",
    "tab_nome_revendas": BASE_DIR / "bases/Tabelas/tab_nome_revendas.xlsx",
    "distribuidor_pdv": BASE_DIR / "bases/Tabelas/WHP_Distribuidor_x_PDV.xlsx",
    "cadastro_revenda": BASE_DIR / "bases/Tabelas/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx",
    "aptos": BASE_DIR / "bases/Tabelas/Aptos_Abril_2026_v4.xlsx",
    "lideranca_aptos": BASE_DIR / "bases/Tabelas/Lideranca_Aptos_Abril_2026_v4.xlsx",
    "vendas_amostra": BASE_DIR / "bases/Vendas Processadas/2026_04.xlsx",
    "base_tratada": BASE_DIR / "bases/Tabelas/base_tratada.xlsx",
}


def analisar_base(nome, caminho):
    """Analisa uma base e retorna dicionário com metadados"""
    print(f"\n{'='*70}")
    print(f"Analisando: {nome}")
    print(f"Arquivo: {caminho}")
    print(f"{'='*70}")

    if not caminho.exists():
        print(f"  ❌ Arquivo não encontrado")
        return None

    try:
        # Descobrir abas
        xl = pd.ExcelFile(caminho)
        abas = xl.sheet_names
        print(f"  Abas: {abas}")

        resultado = {
            "nome": nome,
            "arquivo": str(caminho),
            "tamanho_mb": round(caminho.stat().st_size / (1024*1024), 2),
            "abas": {},
        }

        for aba in abas:
            print(f"\n  --- Aba: '{aba}' ---")
            df = pd.read_excel(caminho, sheet_name=aba)

            # Estatísticas básicas
            print(f"    Linhas: {len(df):,}")
            print(f"    Colunas: {len(df.columns)}")
            print(f"    Colunas: {list(df.columns)}")

            # Tipos de dados
            tipos = df.dtypes.astype(str).to_dict()
            print(f"    Tipos: {tipos}")

            # Amostra de dados (primeiras 3 linhas)
            amostra = df.head(3).to_dict(orient='records')

            # Valores únicos para colunas categóricas (máx 20 valores)
            categorias = {}
            for col in df.columns:
                if df[col].dtype == 'object' or df[col].dtype.name == 'category':
                    unicos = df[col].dropna().unique()
                    if len(unicos) <= 50:
                        categorias[col] = {
                            "unicos": len(unicos),
                            "valores": unicos[:20].tolist() if len(unicos) <= 20 else unicos[:20].tolist() + [f"... e mais {len(unicos)-20}"]
                        }
                # Também verificar colunas numéricas com poucos valores únicos
                elif df[col].nunique() <= 20:
                    categorias[col] = {
                        "unicos": df[col].nunique(),
                        "valores": df[col].dropna().unique().tolist()
                    }

            # Estatísticas numéricas
            numericas = df.describe().to_dict() if df.select_dtypes(include=['number']).shape[1] > 0 else {}

            # Valores nulos
            nulos = df.isnull().sum().to_dict()
            nulos_pct = (df.isnull().sum() / len(df) * 100).round(2).to_dict()

            # Colunas de data
            datas = {}
            for col in df.columns:
                if 'data' in col.lower() or 'date' in col.lower():
                    try:
                        dt_col = pd.to_datetime(df[col], errors='coerce')
                        if dt_col.notna().sum() > 0:
                            datas[col] = {
                                "min": str(dt_col.min()),
                                "max": str(dt_col.max()),
                                "nao_nulos": int(dt_col.notna().sum())
                            }
                    except:
                        pass

            resultado["abas"][aba] = {
                "linhas": len(df),
                "colunas": list(df.columns),
                "tipos": tipos,
                "amostra": amostra,
                "categorias": categorias,
                "estatisticas_numericas": numericas,
                "nulos": nulos,
                "nulos_pct": nulos_pct,
                "colunas_data": datas,
            }

        return resultado

    except Exception as e:
        print(f"  ❌ Erro: {e}")
        return {"nome": nome, "erro": str(e)}


def main():
    resultados = {}

    for nome, caminho in BASES.items():
        resultado = analisar_base(nome, caminho)
        if resultado:
            resultados[nome] = resultado

    # Salvar resultados em JSON
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(resultados, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n\n{'='*70}")
    print(f"Análise completa salva em: {OUTPUT_FILE}")
    print(f"{'='*70}")

    # Resumo executivo
    print("\n📊 RESUMO DAS BASES ANALISADAS:")
    print("-" * 70)
    for nome, res in resultados.items():
        if "erro" not in res:
            for aba_nome, aba_data in res["abas"].items():
                print(f"  {nome:20s} | {aba_nome:20s} | {aba_data['linhas']:>8,} linhas | {len(aba_data['colunas'])} colunas")
        else:
            print(f"  {nome:20s} | ERRO: {res['erro'][:50]}")


if __name__ == "__main__":
    main()
