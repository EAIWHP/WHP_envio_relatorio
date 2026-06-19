"""
Utilitário para leitura correta das bases de vendas do Programa +TOP.

Problema conhecido:
- O arquivo 2025_12.xlsx tem a coluna 'data venda' salva como número serial do Excel
  (ex: 45992 = 01/12/2025). O pandas lê como int64 e pd.to_datetime() interpreta
  como nanosegundos desde 1970, gerando datas em 1970.
- O arquivo 2026_02.xlsx tem a coluna 'data venda' vazia.

Este utilitário NÃO altera os arquivos originais e NÃO impacta os relatórios de
produção. Serve apenas para análises pontuais.
"""

import pandas as pd
from pathlib import Path


def limpar_cpf(cpf):
    if pd.isna(cpf):
        return None
    s = str(cpf).strip().replace("'", "").replace(".", "").replace("-", "").replace("/", "")
    if "." in s:
        s = s.split(".")[0]
    return s.zfill(11)


def converter_data_venda(serie):
    """
    Converte a coluna 'data venda' corretamente.
    - Se for número serial do Excel (ex: 45992), converte para data.
    - Se já for data, mantém.
    - Se for vazio/NaT, mantém NaT.
    """
    # Tenta converter para numérico (número serial do Excel)
    numerico = pd.to_numeric(serie, errors='coerce')

    def _converter(val, num):
        if pd.isna(val):
            return pd.NaT
        if pd.notna(num):
            # Se o valor numérico for razoável para data Excel (entre 1 e 100000)
            if 1 <= num <= 100000:
                return pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(num))
        # Tenta converter como string/data normal
        try:
            return pd.to_datetime(val, errors='coerce')
        except Exception:
            return pd.NaT

    return pd.Series([_converter(v, n) for v, n in zip(serie, numerico)], index=serie.index)


def ler_vendas(caminho):
    """Lê um arquivo de vendas e converte CPF e data venda corretamente."""
    df = pd.read_excel(caminho)
    if 'CPF' in df.columns:
        df['cpf_limp'] = df['CPF'].apply(limpar_cpf)
    if 'data venda' in df.columns:
        df['data_venda'] = converter_data_venda(df['data venda'])
    return df


def ler_todas_vendas(diretorio, padrao='*.xlsx'):
    """Lê todos os arquivos de vendas de um diretório."""
    diretorio = Path(diretorio)
    arquivos = sorted(diretorio.glob(padrao))
    dfs = []
    for arquivo in arquivos:
        if arquivo.name.startswith('~') or arquivo.name.startswith('OLD'):
            continue
        df = ler_vendas(arquivo)
        df['nome_arquivo'] = arquivo.name
        dfs.append(df)
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    return pd.DataFrame()


if __name__ == '__main__':
    VENDAS_DIR = Path('/home/thamiresvieira/projetos/Programa_mais_top/bases/Vendas Processadas')
    df = ler_todas_vendas(VENDAS_DIR)
    print(f"Total de registros lidos: {len(df)}")
    print(f"\nDistribuição de anos na data_venda:")
    print(df['data_venda'].dt.year.value_counts().sort_index().to_string())
    print(f"\nRegistros sem data: {df['data_venda'].isna().sum()}")
