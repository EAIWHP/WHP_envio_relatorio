"""
Utilitários para o Relatório de Diretoria One Page +TOP.
"""

import pandas as pd
from pathlib import Path

from config import (
    META_CADASTRO,
    META_TREINAMENTO,
    META_ACEITE,
    FAIXA_VERDE,
    FAIXA_AMARELO,
    VALOR_PONTO,
    MAPEAMENTO_REVENDA_CAMPANHA,
)


def normalizar_revenda_campanha(nome):
    """Normaliza nome de revenda vindo da planilha Campanha para o padrão do projeto."""
    if pd.isna(nome):
        return nome
    s = str(nome).strip().upper()
    s = (
        s.replace("Á", "A").replace("É", "E").replace("Í", "I")
         .replace("Ó", "O").replace("Ú", "U").replace("Ã", "A")
         .replace("Õ", "O").replace("Ç", "C").replace("Ê", "E")
    )
    if s in MAPEAMENTO_REVENDA_CAMPANHA:
        return MAPEAMENTO_REVENDA_CAMPANHA[s]
    return str(nome).strip()


def limpar_cpf(cpf):
    """Normaliza CPF/CNPJ para texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    try:
        if isinstance(cpf, float):
            cpf = int(cpf)
    except Exception:
        pass
    s = str(cpf).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(11)


def extrair_sku_curto(sku_produto):
    """Extrai o SKU curto da coluna 'SKU + produto', removendo descrição após hífen."""
    if pd.isna(sku_produto):
        return ""
    s = str(sku_produto).strip()
    # Remove tudo após o primeiro hífen seguido de letra/espaço (ex: 'BRM46MK-Refrigerador...')
    if "-" in s:
        s = s.split("-")[0].strip()
    return s


def classificar_farol(valor, meta):
    """Classifica o farol de um indicador percentual."""
    if pd.isna(valor):
        return "cinza"
    if valor >= meta * FAIXA_VERDE:
        return "verde"
    if valor >= meta * FAIXA_AMARELO:
        return "amarelo"
    return "vermelho"


def calcular_ranking(pct_cadastro, pct_treinamento, pct_aceite):
    """Calcula ranking 0-3 (1 ponto por meta atingida)."""
    pontos = 0
    if pd.notna(pct_cadastro) and pct_cadastro >= META_CADASTRO * 100:
        pontos += 1
    if pd.notna(pct_treinamento) and pct_treinamento >= META_TREINAMENTO * 100:
        pontos += 1
    if pd.notna(pct_aceite) and pct_aceite >= META_ACEITE * 100:
        pontos += 1
    return pontos


def formatar_pct(valor):
    """Formata percentual com 1 casa decimal."""
    if pd.isna(valor):
        return "N/A"
    return f"{valor:.1f}%"


def formatar_moeda(valor):
    """Formata valor em reais."""
    if pd.isna(valor):
        return "N/A"
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def formatar_inteiro(valor):
    """Formata número inteiro."""
    if pd.isna(valor):
        return "N/A"
    return f"{int(valor):,}".replace(",", ".")


def carregar_excel_robusto(path, **kwargs):
    """Wrapper para pd.read_excel com tratamento de erro básico."""
    try:
        return pd.read_excel(path, **kwargs)
    except Exception as e:
        raise RuntimeError(f"Erro ao ler {path}: {e}")


def iterar_meses(ano_inicio, mes_inicio, ano_fim, mes_fim):
    """Gera lista de tuplas (ano, mes) no intervalo fechado."""
    meses = []
    ano, mes = ano_inicio, mes_inicio
    while (ano < ano_fim) or (ano == ano_fim and mes <= mes_fim):
        meses.append((ano, mes))
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
    return meses


def nome_aba_aceites(ano, mes, meses_pt=None):
    """Tenta retornar o nome provável da aba de aceites para um determinado mês/ano."""
    if meses_pt is None:
        from config import MESES_PT
        meses_pt = MESES_PT
    return f"{meses_pt[mes].upper()}_{ano}"
