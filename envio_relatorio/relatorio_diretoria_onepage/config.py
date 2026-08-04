"""
Configurações do Relatório de Diretoria One Page +TOP.
"""

from pathlib import Path

# Período de referência
# Indicadores operacionais (Cadastro/Treinamento/Aceite) usam o mês abaixo.
# Dados financeiros (Investimento/Vendas) permanecem de junho/2026 porque a
# planilha Campanha_+TOP_Julho_2026 ainda não foi disponibilizada.
MES_REF = 7
ANO_REF = 2026

# Período acumulado (out/2025 a jun/2026)
# Nota: dados financeiros (Campanha +TOP) só existem a partir de abr/2026.
MES_INICIO_ACUMULADO = 10
ANO_INICIO_ACUMULADO = 2025

# Metas dos indicadores
META_CADASTRO = 0.85
META_TREINAMENTO = 0.70
META_ACEITE = 0.70

# Pesos do ranking (1 ponto por meta atingida)
PONTOS_POR_META = 1

# 1 ponto = R$ 1,00
VALOR_PONTO = 1.0

# Cores do KV +TOP / Whirlpool
COR_PRIMARIA = "EF4E22"       # laranja Whirlpool
COR_PRIMARIA_CLARA = "FDE8E0"
COR_CINZA = "F2F2F2"
COR_BRANCO = "FFFFFF"
COR_TEXTO = "333333"
COR_VERDE = "2E7D32"
COR_AMARELO = "F9A825"
COR_VERMELHO = "C62828"

# Farol de KPIs
FAIXA_VERDE = 1.0    # atingiu meta
FAIXA_AMARELO = 0.5  # entre 50% e a meta

# Caminhos das bases
BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")
ENVIO_RELATORIO_DIR = BASE_DIR / "envio_relatorio"
HIERARQUIA_DIR = ENVIO_RELATORIO_DIR / "bases" / "bases_cadastro_hierarquia"
CADASTRO_FILE = ENVIO_RELATORIO_DIR / "bases" / "cadastro.xlsx"
TREINAMENTOS_FILE = ENVIO_RELATORIO_DIR / "bases" / "Base_treinamentos.xlsx"
ACEITES_FILE = ENVIO_RELATORIO_DIR / "bases" / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
VENDAS_DIR = BASE_DIR / "bases" / "Vendas Processadas"

# Planilha de referência consolidada (Campanha +TOP)
CAMPANHA_FILE = Path(__file__).resolve().parent / "Campanha_+TOP_Junho_2026 (2).xlsx"

# Planilhas Campanha históricas para versão acumulada (dados disponíveis)
CAMPANHA_FILES_ACUMULADO = {
    (2026, 4): BASE_DIR / "bases" / "Historico hierarquias" / "Bases de Abril_26" / "APURAÇÃO" / "Campanha_+TOP_Abril_2026.xlsx",
    (2026, 5): BASE_DIR / "bases" / "Historico hierarquias" / "Bases de Maio_26" / "APURAÇÃO - PONTUAÇÕES CREDITADAS" / "Campanha_+TOP_Maio_2026.xlsx",
    (2026, 6): Path(__file__).resolve().parent / "Campanha_+TOP_Junho_2026 (2).xlsx",
}

# Planilha específica da Bemol
CAMPANHA_BEMOL_FILE = Path(__file__).resolve().parent / "Campanha_+TOP_BEMOL_Junho_2026.xlsx"

# Caminho de saída
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Mapeamento de meses
MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

# Mapeamento de nomes de revenda da planilha Campanha -> nome padrão do projeto
MAPEAMENTO_REVENDA_CAMPANHA = {
    "MAGAZAN": "Líder",
    "CDA": "Casas da Água",
    "MM ATACADO": "MM Atacado",
    "MM VAREJO": "MM Varejo",
    "SOLAR": "Solar",
    "SOLAR MAGAZINE": "Solar Magazine",
    "ARMAZÉM MATEUS": "Armazem Mateus",
    "IMPÉRIO": "Império",
    "LASER ELETRO": "Laser Eletro",
}

# Mapeamento de nomes de revenda para exibição final (title case uniforme)
NOME_EXIBICAO_REVENDA = {
    "Angeloni": "Angeloni",
    "Armazem Mateus": "Armazém Mateus",
    "Armazém Mateus": "Armazém Mateus",
    "Becker": "Becker",
    "Bemol": "Bemol",
    "Casas da Água": "Casas da Água",
    "Colombo": "Colombo",
    "Estrela": "Estrela",
    "Formosa": "Formosa",
    "Gazin Atacado": "Gazin Atacado",
    "Gazin Online": "Gazin Online",
    "Gazin Varejo": "Gazin Varejo",
    "Guaibim": "Guaibim",
    "Havan": "Havan",
    "Império": "Império",
    "Imperio": "Império",
    "Jmahfuz": "Jmahfuz",
    "Koerich": "Koerich",
    "Laser Eletro": "Laser Eletro",
    "Lebes": "Lebes",
    "Líder": "Líder",
    "Lider": "Líder",
    "MM Atacado": "MM Atacado",
    "MM Varejo": "MM Varejo",
    "Millena": "Millena",
    "Multiloja": "Multiloja",
    "Nosso Lar": "Nosso Lar",
    "Ramsons": "Ramsons",
    "Sipolatti": "Sipolatti",
    "Solar": "Solar",
    "Solar Magazine": "Solar Magazine",
    "Taqi": "Taqi",
    "Tele Rio": "Tele Rio",
    "Zema": "Zema",
    "Zenir": "Zenir",
}

# Colunas finais do relatório (estrutura Staff Zanatta)
COLUNAS_RELATORIO = [
    "Revenda",
    "Regional",
    "Cadastro (%)",
    "Treinamento (%)",
    "Aceite (%)",
    "Investimento Mês (R$)",
    "Não Investimento (R$)",
    "Vendas",
    "Vendas LM",
    "Var LM (%)",
    "Vendas L3M",
    "Vendas YOY",
    "Var YOY (%)",
    "Ranking",
]
