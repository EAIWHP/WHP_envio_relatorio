#!/usr/bin/env python3
"""
Script para converter campos CPF de número para texto (com zeros à esquerda)
em todas as bases do Programa +TOP
"""

import pandas as pd
import shutil
from pathlib import Path
import os

BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")
BACKUP_DIR = BASE_DIR / "bases/Old/backup_cpf_texto_20250608"

# Criar diretório de backup
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

def backup_file(src_path):
    """Faz backup do arquivo original"""
    dst = BACKUP_DIR / src_path.name
    shutil.copy2(src_path, dst)
    print(f"  💾 Backup: {dst}")

def formatar_cpf(cpf_valor):
    """Converte CPF para texto com 11 dígitos (zeros à esquerda)"""
    if pd.isna(cpf_valor):
        return None
    try:
        # Remover qualquer formatação existente (pontos, traços, espaços)
        cpf_limpo = str(cpf_valor).strip()
        cpf_limpo = cpf_limpo.replace('.', '').replace('-', '').replace('/', '').replace(' ', '')
        # Remover parte decimal se for float (ex: 123456789.0 → 123456789)
        if '.' in cpf_limpo:
            cpf_limpo = cpf_limpo.split('.')[0]
        # Preencher com zeros à esquerda até 11 dígitos
        return cpf_limpo.zfill(11)
    except:
        return str(cpf_valor)

def processar_excel(caminho, colunas_cpf, abas=None):
    """Processa um arquivo Excel convertendo CPFs para texto"""
    print(f"\n{'='*60}")
    print(f"📁 Processando: {caminho.name}")
    print(f"{'='*60}")

    # Backup
    backup_file(caminho)

    # Ler arquivo
    xl = pd.ExcelFile(caminho)
    abas_para_processar = abas if abas else xl.sheet_names

    # Dicionário para guardar DataFrames modificados
    dfs = {}

    for aba in xl.sheet_names:
        df = pd.read_excel(caminho, sheet_name=aba)

        if aba in abas_para_processar:
            print(f"  📄 Aba '{aba}': {len(df)} linhas")

            for col in colunas_cpf:
                if col in df.columns:
                    # Contar quantos eram nulos antes
                    nulos_antes = df[col].isna().sum()

                    # Converter para texto formatado
                    df[col] = df[col].apply(formatar_cpf)

                    # Contar quantos ficaram nulos depois
                    nulos_depois = df[col].isna().sum()

                    # Mostrar amostra
                    amostra = df[col].dropna().head(3).tolist()
                    print(f"     Coluna '{col}': {len(df) - nulos_depois} preenchidos")
                    print(f"        Amostra: {amostra}")

            dfs[aba] = df
        else:
            print(f"  📄 Aba '{aba}': pulada (não na lista de processamento)")
            dfs[aba] = df

    # Salvar arquivo modificado
    with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
        for aba, df in dfs.items():
            df.to_excel(writer, sheet_name=aba, index=False)

    print(f"  ✅ Arquivo salvo: {caminho}")
    return True

# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

print("="*60)
print("CONVERTENDO CPFs DE NÚMERO PARA TEXTO")
print("="*60)
print(f"Backup em: {BACKUP_DIR}")

# 1. ACESSOS
processar_excel(
    BASE_DIR / "bases/Tabelas/acessos.xlsx",
    colunas_cpf=["Cpf/Cnpj"]
)

# 2. PONTOS POR CPF (todas as abas que têm CPF)
processar_excel(
    BASE_DIR / "bases/Tabelas/pontos_por_cpf_.xlsx",
    colunas_cpf=["CPF", "CPFCNPJ", "DOCUMENTO"],
    abas=["CONSOLIDADO_RESGATES", "PONTOS POR CPF", "VALES", "LOJA VIRTUAL", "PAGAMENTO DE CONTAS", "RECARGA DE CELULAR"]
)

# 3. RESGATES DETALHADO
processar_excel(
    BASE_DIR / "bases/Tabelas/resgates_detalhado.xlsx",
    colunas_cpf=["CPF"]
)

# 4. TREINAMENTOS
processar_excel(
    BASE_DIR / "bases/Tabelas/Base_treinamentos.xlsx",
    colunas_cpf=["CPF", "Cnpj Distribuidor", "Cnpj PDV"]
)

# 5. VENDAS PROCESSADAS (todos os arquivos)
print(f"\n{'='*60}")
print("📁 Processando: Vendas Processadas (todos os arquivos)")
print(f"{'='*60}")
vendas_dir = BASE_DIR / "bases/Vendas Processadas"
for arquivo in sorted(vendas_dir.glob("*.xlsx")):
    if arquivo.name.startswith("OLD") or arquivo.name.startswith("~"):
        continue
    try:
        processar_excel(arquivo, colunas_cpf=["CPF"])
    except Exception as e:
        print(f"  ❌ Erro em {arquivo.name}: {e}")

# 6. ENQUETE
processar_excel(
    BASE_DIR / "bases/Tabelas/enquete.xlsx",
    colunas_cpf=["cpf"]
)

# 7. ACEITES (todas as abas)
processar_excel(
    BASE_DIR / "bases/Tabelas/WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx",
    colunas_cpf=["CPF"]
)

# 8. CADASTRO REVENDA
processar_excel(
    BASE_DIR / "bases/Tabelas/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx",
    colunas_cpf=["Cpf", "CNPJ_Revenda", "CNPJ_Loja"]
)

print(f"\n{'='*60}")
print("✅ CONVERSÃO CONCLUÍDA")
print(f"{'='*60}")
print(f"Backups salvos em: {BACKUP_DIR}")
