#!/usr/bin/env python3
"""
Análise profunda das bases do Programa +TOP
- Valores únicos das colunas categóricas chave
- Cruzamentos entre bases (integridade referencial)
- Estatísticas de qualidade de dados
"""

import pandas as pd
from pathlib import Path

BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")

# Ler todas as bases principais em memória
print("=" * 80)
print("CARREGANDO BASES EM MEMÓRIA...")
print("=" * 80)

# 1. Cadastro
print("\n📋 Carregando cadastro...")
df_cadastro = pd.read_excel(BASE_DIR / "bases/Tabelas/cadastro.xlsx")
df_cadastro.columns = df_cadastro.columns.str.strip().str.lower()
print(f"  Linhas: {len(df_cadastro):,} | Colunas: {list(df_cadastro.columns)}")

# 2. Acessos
print("\n📋 Carregando acessos...")
df_acessos = pd.read_excel(BASE_DIR / "bases/Tabelas/acessos.xlsx")
df_acessos.columns = df_acessos.columns.str.strip()
print(f"  Linhas: {len(df_acessos):,} | Colunas: {list(df_acessos.columns)}")

# 3. Pontos por CPF
print("\n📋 Carregando pontos_por_cpf...")
xl_pontos = pd.ExcelFile(BASE_DIR / "bases/Tabelas/pontos_por_cpf_.xlsx")
print(f"  Abas: {xl_pontos.sheet_names}")
df_pontos = pd.read_excel(BASE_DIR / "bases/Tabelas/pontos_por_cpf_.xlsx", sheet_name="PONTOS POR CPF")
df_pontos.columns = df_pontos.columns.str.strip()
print(f"  Linhas: {len(df_pontos):,} | Colunas: {list(df_pontos.columns)}")

# 4. Resgates
print("\n📋 Carregando resgates...")
df_resgates = pd.read_excel(BASE_DIR / "bases/Tabelas/resgates_detalhado.xlsx")
df_resgates.columns = df_resgates.columns.str.strip()
print(f"  Linhas: {len(df_resgates):,} | Colunas: {list(df_resgates.columns)}")

# 5. Treinamentos
print("\n📋 Carregando treinamentos...")
df_trein = pd.read_excel(BASE_DIR / "bases/Tabelas/Base_treinamentos.xlsx")
df_trein.columns = df_trein.columns.str.strip()
print(f"  Linhas: {len(df_trein):,} | Colunas: {list(df_trein.columns)}")

# 6. Aceites
print("\n📋 Carregando aceites (último mês - ABR_2026)...")
df_aceites = pd.read_excel(BASE_DIR / "bases/Tabelas/WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx", sheet_name="ABR_2026")
df_aceites.columns = df_aceites.columns.str.strip()
print(f"  Linhas: {len(df_aceites):,} | Colunas: {list(df_aceites.columns)}")

# 7. Chamados
print("\n📋 Carregando chamados...")
df_chamados = pd.read_excel(BASE_DIR / "bases/Tabelas/Chamados_catalogo.xlsx")
df_chamados.columns = df_chamados.columns.str.strip()
print(f"  Linhas: {len(df_chamados):,} | Colunas: {list(df_chamados.columns)}")

# 8. Vendas
print("\n📋 Carregando vendas (2026_04)...")
df_vendas = pd.read_excel(BASE_DIR / "bases/Vendas Processadas/2026_04.xlsx")
df_vendas.columns = df_vendas.columns.str.strip().str.lower()
print(f"  Linhas: {len(df_vendas):,} | Colunas: {list(df_vendas.columns)}")

# 9. Cadastro Revenda
print("\n📋 Carregando cadastro_revenda...")
df_cad_rev = pd.read_excel(BASE_DIR / "bases/Tabelas/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx")
df_cad_rev.columns = df_cad_rev.columns.str.strip()
print(f"  Linhas: {len(df_cad_rev):,} | Colunas: {list(df_cad_rev.columns)}")

# 10. Aptos
print("\n📋 Carregando aptos...")
df_aptos = pd.read_excel(BASE_DIR / "bases/Tabelas/Aptos_Abril_2026_v4.xlsx")
df_aptos.columns = df_aptos.columns.str.strip().str.lower()
print(f"  Linhas: {len(df_aptos):,} | Colunas: {list(df_aptos.columns)}")

# 11. Ocorrências
print("\n📋 Carregando ocorrências...")
df_ocor = pd.read_excel(BASE_DIR / "bases/Tabelas/OcorrênciasTOPSAC.xlsx")
df_ocor.columns = df_ocor.columns.str.strip()
print(f"  Linhas: {len(df_ocor):,} | Colunas: {list(df_ocor.columns)}")

# 12. Enquete
print("\n📋 Carregando enquete...")
df_enquete = pd.read_excel(BASE_DIR / "bases/Tabelas/enquete.xlsx")
df_enquete.columns = df_enquete.columns.str.strip().str.lower()
print(f"  Linhas: {len(df_enquete):,} | Colunas: {list(df_enquete.columns)}")

# ============================================
# ANÁLISE 1: VALORES ÚNICOS DAS COLUNAS CHAVE
# ============================================

print("\n" + "=" * 80)
print("1. VALORES ÚNICOS DAS COLUNAS CATEGÓRICAS CHAVE")
print("=" * 80)

print("\n--- CADASTRO ---")
print(f"  Status: {df_cadastro['status'].dropna().unique()}")
print(f"  Cargos: {df_cadastro['cargo'].dropna().unique()}")
print(f"  Regionais: {sorted(df_cadastro['regional'].dropna().unique())}")
print(f"  Total grupos (revendas): {df_cadastro['grupo'].nunique()}")

print("\n--- ACESSOS ---")
print(f"  Status: {df_acessos['Status'].dropna().unique()}")
print(f"  Cargos: {df_acessos['Cargo'].dropna().unique()}")
print(f"  Regionais: {sorted(df_acessos['Regional'].dropna().unique())}")
print(f"  Grupos Cluster: {df_acessos['Grupo Cluster'].nunique()} únicos")
print(f"  Páginas (top 10): {df_acessos['Pagina'].value_counts().head(10).to_dict()}")

print("\n--- TREINAMENTOS ---")
print(f"  Estados: {df_trein['Estado'].dropna().unique()}")
print(f"  Cursos (top 10): {df_trein['Curso'].value_counts().head(10).to_dict()}")
print(f"  Trilhas: {df_trein['Trilha'].dropna().unique()}")
print(f"  Distribuidores (top 10): {df_trein['Distribuidor'].value_counts().head(10).to_dict()}")
print(f"  Progresso - únicos: {sorted(df_trein['Progresso'].dropna().unique())}")

print("\n--- VENDAS (2026_04) ---")
print(f"  Categorias: {sorted(df_vendas['categoria'].dropna().unique())}")
print(f"  Subcategorias: {sorted(df_vendas['subcategoria'].dropna().unique())}")
print(f"  Regionais: {sorted(df_vendas['regional'].dropna().unique())}")
print(f"  Revendas: {df_vendas['revenda'].nunique()} únicas")
print(f"  SUPERTOP: {sorted(df_vendas['supertop'].dropna().unique())}")
print(f"  SKU + produto (top 5): {df_vendas['sku + produto'].value_counts().head(5).to_dict()}")

print("\n--- PONTOS POR CPF ---")
print(f"  Total CPFs: {df_pontos['CPFCNPJ'].nunique()}")
print(f"  Créditos - min: {df_pontos['Créditos'].min():.0f}, max: {df_pontos['Créditos'].max():.0f}")
print(f"  Saldo Atual - min: {df_pontos['Saldo Atual'].min():.0f}, max: {df_pontos['Saldo Atual'].max():.0f}")
print(f"  CPFs com saldo negativo: {(df_pontos['Saldo Atual'] < 0).sum()}")

print("\n--- RESGATES ---")
print(f"  Produtos (top 10): {df_resgates['PRODUTO'].value_counts().head(10).to_dict()}")
print(f"  Valor - min: {df_resgates['VALOR'].min():.2f}, max: {df_resgates['VALOR'].max():.2f}")
print(f"  Período: {df_resgates['DATA RESGATE'].min()} a {df_resgates['DATA RESGATE'].max()}")

print("\n--- CHAMADOS ---")
print(f"  Assuntos: {df_chamados['ASSUNTO'].value_counts().head(15).to_dict()}")
print(f"  Resolução média (dias): {df_chamados['SLA - DIA'].mean():.1f}")

print("\n--- APTOS ---")
print(f"  KPIs: {df_aptos['kpi'].dropna().unique()}")
print(f"  Produtos (top 10): {df_aptos['produto'].value_counts().head(10).to_dict()}")
print(f"  Treinamento: {df_aptos['treinamento'].dropna().unique()}")
print(f"  Meses: {sorted(df_aptos['mes'].dropna().unique())}")
print(f"  Anos: {sorted(df_aptos['ano'].dropna().unique())}")

print("\n--- OCORRÊNCIAS ---")
print(f"  Solicitações: {df_ocor['+TOP - SOLICITAÇÃO'].value_counts().head(10).to_dict()}")
print(f"  Motivos: {df_ocor['+TOP - MOTIVO'].value_counts().head(10).to_dict()}")
print(f"  Status: {df_ocor['Status (sem tempo decorrido)'].dropna().unique()}")

print("\n--- ENQUETE ---")
print(f"  Cargos: {df_enquete['cargo'].dropna().unique()}")
print(f"  Tipos: {df_enquete['tipo'].dropna().unique()}")
print(f"  Perguntas: {df_enquete['pergunta'].dropna().unique()}")
print(f"  Respostas: {df_enquete['resposta'].dropna().unique()}")

print("\n--- CADASTRO REVENDA ---")
print(f"  Cargos: {df_cad_rev['Cargo'].dropna().unique()}")
print(f"  Status: {df_cad_rev['Status'].dropna().unique()}")
print(f"  Regionais da Revenda: {sorted(df_cad_rev['Regional_da_Revenda'].dropna().unique())}")


# ============================================
# ANÁLISE 2: CRUZAMENTOS ENTRE BASES
# ============================================

print("\n" + "=" * 80)
print("2. CRUZAMENTOS ENTRE BASES (INTEGRIDADE REFERENCIAL)")
print("=" * 80)

# Normalizar CPFs para texto (sem pontos)
df_cadastro['cpf_str'] = df_cadastro['cpf/cnpj'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_acessos['cpf_str'] = df_acessos['Cpf/Cnpj'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_pontos['cpf_str'] = df_pontos['CPFCNPJ'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_resgates['cpf_str'] = df_resgates['CPF'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_trein['cpf_str'] = df_trein['CPF'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_aceites['cpf_str'] = df_aceites['CPF'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_vendas['cpf_str'] = df_vendas['cpf'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_aptos['cpf_str'] = df_aptos['cpf'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_cad_rev['cpf_str'] = df_cad_rev['Cpf'].astype(str).str.replace(r'[^0-9]', '', regex=True)
df_enquete['cpf_str'] = df_enquete['cpf'].astype(str).str.replace(r'[^0-9]', '', regex=True)

# Cadastro vs Acessos
print("\n--- CADASTRO vs ACESSOS ---")
cpf_cad = set(df_cadastro['cpf_str'].dropna())
cpf_acess = set(df_acessos['cpf_str'].dropna())
print(f"  CPFs em cadastro: {len(cpf_cad):,}")
print(f"  CPFs em acessos: {len(cpf_acess):,}")
print(f"  CPFs em comum: {len(cpf_cad & cpf_acess):,}")
print(f"  CPFs em cadastro mas NÃO em acessos: {len(cpf_cad - cpf_acess):,}")
print(f"  CPFs em acessos mas NÃO em cadastro: {len(cpf_acess - cpf_cad):,}")

# Cadastro vs Pontos
print("\n--- CADASTRO vs PONTOS_POR_CPF ---")
cpf_pontos = set(df_pontos['cpf_str'].dropna())
print(f"  CPFs em pontos_por_cpf: {len(cpf_pontos):,}")
print(f"  CPFs em comum (cadastro x pontos): {len(cpf_cad & cpf_pontos):,}")
print(f"  CPFs em cadastro mas SEM pontos: {len(cpf_cad - cpf_pontos):,}")
print(f"  CPFs com pontos mas NÃO em cadastro: {len(cpf_pontos - cpf_cad):,}")

# Cadastro vs Resgates
print("\n--- CADASTRO vs RESGATES ---")
cpf_resgates = set(df_resgates['cpf_str'].dropna())
print(f"  CPFs em resgates: {len(cpf_resgates):,}")
print(f"  CPFs em comum (cadastro x resgates): {len(cpf_cad & cpf_resgates):,}")
print(f"  CPFs em cadastro mas SEM resgates: {len(cpf_cad - cpf_resgates):,}")

# Cadastro vs Treinamentos
print("\n--- CADASTRO vs TREINAMENTOS ---")
cpf_trein = set(df_trein['cpf_str'].dropna())
print(f"  CPFs em treinamentos: {len(cpf_trein):,}")
print(f"  CPFs em comum (cadastro x treinamentos): {len(cpf_cad & cpf_trein):,}")
print(f"  CPFs em cadastro mas SEM treinamentos: {len(cpf_cad - cpf_trein):,}")

# Cadastro vs Aceites
print("\n--- CADASTRO vs ACEITES (ABR_2026) ---")
cpf_aceites = set(df_aceites['cpf_str'].dropna())
print(f"  CPFs com aceite em Abr/2026: {len(cpf_aceites):,}")
print(f"  CPFs em comum (cadastro x aceites): {len(cpf_cad & cpf_aceites):,}")
print(f"  CPFs em cadastro ATIVO mas SEM aceite: {len(cpf_cad - cpf_aceites):,}")

# Cadastro vs Vendas
print("\n--- CADASTRO vs VENDAS (2026_04) ---")
cpf_vendas = set(df_vendas['cpf_str'].dropna())
print(f"  CPFs em vendas (Abr/2026): {len(cpf_vendas):,}")
print(f"  CPFs em comum (cadastro x vendas): {len(cpf_cad & cpf_vendas):,}")
print(f"  CPFs em cadastro mas SEM vendas: {len(cpf_cad - cpf_vendas):,}")

# Cadastro vs Aptos
print("\n--- CADASTRO vs APTOS ---")
cpf_aptos = set(df_aptos['cpf_str'].dropna())
print(f"  CPFs em aptos: {len(cpf_aptos):,}")
print(f"  CPFs em comum (cadastro x aptos): {len(cpf_cad & cpf_aptos):,}")

# Cadastro vs Cadastro Revenda
print("\n--- CADASTRO vs CADASTRO_REVENDA ---")
cpf_cad_rev_set = set(df_cad_rev['cpf_str'].dropna())
print(f"  CPFs em cadastro_revenda: {len(cpf_cad_rev_set):,}")
print(f"  CPFs em comum: {len(cpf_cad & cpf_cad_rev_set):,}")
print(f"  CPFs em cadastro mas NÃO em cadastro_revenda: {len(cpf_cad - cpf_cad_rev_set):,}")

# Cadastro vs Enquete
print("\n--- CADASTRO vs ENQUETE ---")
cpf_enquete = set(df_enquete['cpf_str'].dropna())
print(f"  CPFs em enquete: {len(cpf_enquete):,}")
print(f"  CPFs em comum (cadastro x enquete): {len(cpf_cad & cpf_enquete):,}")


# ============================================
# ANÁLISE 3: ESTATÍSTICAS DE QUALIDADE
# ============================================

print("\n" + "=" * 80)
print("3. ESTATÍSTICAS DE QUALIDADE DE DADOS")
print("=" * 80)

# Cadastro
print("\n--- QUALIDADE: CADASTRO ---")
print(f"  CPFs duplicados: {df_cadastro['cpf/cnpj'].duplicated().sum()}")
print(f"  Nomes nulos: {df_cadastro['nome'].isnull().sum()}")
print(f"  Status nulos: {df_cadastro['status'].isnull().sum()}")
print(f"  Cargo nulos: {df_cadastro['cargo'].isnull().sum()}")
print(f"  Regional nulos: {df_cadastro['regional'].isnull().sum()}")
print(f"  Grupo nulos: {df_cadastro['grupo'].isnull().sum()}")
print(f"  Data de aceite nula: {df_cadastro['data de aceite'].isnull().sum()}")
print(f"  Data de aceite preenchida: {df_cadastro['data de aceite'].notna().sum()}")

# Análise de status
status_counts = df_cadastro['status'].value_counts()
print(f"\n  Distribuição de Status:")
for s, c in status_counts.items():
    print(f"    {s}: {c:,} ({c/len(df_cadastro)*100:.1f}%)")

# Análise de cargo
cargo_counts = df_cadastro['cargo'].value_counts()
print(f"\n  Distribuição de Cargos:")
for c, n in cargo_counts.items():
    print(f"    {c}: {n:,} ({n/len(df_cadastro)*100:.1f}%)")

# Vendas
print("\n--- QUALIDADE: VENDAS (2026_04) ---")
print(f"  Vendas nulas: {df_vendas['vendas'].isnull().sum()}")
print(f"  Pontuação nula: {df_vendas['pontuação'].isnull().sum()}")
print(f"  CPF nulo: {df_vendas['cpf'].isnull().sum()}")
print(f"  SKU + produto nulo: {df_vendas['sku + produto'].isnull().sum()}")
print(f"  Vendas = 0: {(df_vendas['vendas'] == 0).sum()}")
print(f"  Pontuação = 0: {(df_vendas['pontuação'] == 0).sum()}")
print(f"  Registros com Vendas > 0 e Pontuação = 0: {((df_vendas['vendas'] > 0) & (df_vendas['pontuação'] == 0)).sum()}")
print(f"  Registros com Vendas = 0 e Pontuação > 0: {((df_vendas['vendas'] == 0) & (df_vendas['pontuação'] > 0)).sum()}")

# Análise SUPERTOP
print(f"\n  SUPERTOP:")
print(f"    Registros SUPERTOP: {(df_vendas['supertop'] == 1).sum():,}")
print(f"    Registros não SUPERTOP: {(df_vendas['supertop'] == 0).sum():,}")
print(f"    Registros SUPERTOP nulo: {df_vendas['supertop'].isnull().sum():,}")

# Pontos
print("\n--- QUALIDADE: PONTOS_POR_CPF ---")
print(f"  CPFs duplicados: {df_pontos['CPFCNPJ'].duplicated().sum()}")
print(f"  Créditos = 0: {(df_pontos['Créditos'] == 0).sum()}")
print(f"  Resgates = 0: {(df_pontos['Resgates'] == 0).sum()}")
print(f"  Expirado = 0: {(df_pontos['EXPIRADO '] == 0).sum()}")
print(f"  Saldo = 0: {(df_pontos['Saldo Atual'] == 0).sum()}")
print(f"  Saldo > 0: {(df_pontos['Saldo Atual'] > 0).sum()}")
print(f"  Saldo < 0: {(df_pontos['Saldo Atual'] < 0).sum()}")

# Treinamentos
print("\n--- QUALIDADE: TREINAMENTOS ---")
print(f"  CPFs únicos: {df_trein['CPF'].nunique()}")
print(f"  Cursos únicos: {df_trein['Curso'].nunique()}")
print(f"  Progresso = 100: {(df_trein['Progresso'] == 100).sum():,}")
print(f"  Progresso < 100: {(df_trein['Progresso'] < 100).sum():,}")
print(f"  Estado = 'Concluído': {(df_trein['Estado'] == 'Concluído').sum():,}")
print(f"  Estado = 'Em andamento': {(df_trein['Estado'] == 'Em andamento').sum():,}")
print(f"  Estado = 'Não iniciado': {(df_trein['Estado'] == 'Não iniciado').sum():,}")
print(f"  Estados únicos: {df_trein['Estado'].dropna().unique()}")

# Aceites
print("\n--- QUALIDADE: ACEITES (ABR_2026) ---")
print(f"  CPFs duplicados: {df_aceites['CPF'].duplicated().sum()}")
print(f"  Revendas únicas: {df_aceites['Revenda'].nunique()}")
print(f"  DataAceite nula: {df_aceites['DataAceite'].isnull().sum()}")

# Aptos
print("\n--- QUALIDADE: APTOS ---")
print(f"  CPFs únicos: {df_aptos['cpf'].nunique()}")
print(f"  Revendas únicas: {df_aptos['revenda'].nunique()}")
print(f"  KPIs: {df_aptos['kpi'].dropna().unique()}")
print(f"  Treinamento: {df_aptos['treinamento'].dropna().unique()}")

# Análise temporal
print("\n--- ANÁLISE TEMPORAL ---")
print(f"  Acessos: {df_acessos['Data Acesso'].min()} a {df_acessos['Data Acesso'].max()}")
print(f"  Resgates: {df_resgates['DATA RESGATE'].min()} a {df_resgates['DATA RESGATE'].max()}")
print(f"  Chamados: {df_chamados['DATA'].min()} a {df_chamados['DATA'].max()}")

# Anos/meses das vendas
print(f"\n  Vendas (último arquivo 2026_04):")
print(f"    Data Venda: {df_vendas['data venda'].min()} a {df_vendas['data venda'].max()}")
print(f"    Mês: {df_vendas['mês'].unique()}")

print("\n" + "=" * 80)
print("ANÁLISE PROFUNDA CONCLUÍDA")
print("=" * 80)
