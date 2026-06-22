#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validar_hierarquia_banco.py

Compila e valida as bases de hierarquia (pasta bases_cadastro_hierarquia)
contra a base do banco PROD_WHP_Participante_Banco_v2_*.xlsx.

Gera um arquivo Excel com abas de validação:
    - Hierarquia Consolidada
    - Banco
    - Cruzamento Completo
    - Só na Hierarquia
    - Só no Banco
    - Resumo por Revenda
    - Duplicados na Hierarquia
    - Duplicados no Banco
"""

import logging
import re
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
HIERARQUIA_DIR = BASE_DIR / "bases" / "bases_cadastro_hierarquia"
SAIDA_DIR = HIERARQUIA_DIR


def limpar_cpf(cpf):
    """Normaliza CPF em texto com 11 dígitos."""
    if pd.isna(cpf):
        return None
    cpf_str = str(cpf).strip().replace("'", "")
    # Remove tudo que não for dígito
    cpf_limpo = re.sub(r"[^0-9]", "", cpf_str)
    if not cpf_limpo:
        return None
    # Preenche com zeros à esquerda para garantir 11 dígitos
    return cpf_limpo.zfill(11)


def limpar_cnpj(cnpj):
    """Normaliza CNPJ em texto com 14 dígitos."""
    if pd.isna(cnpj):
        return None
    cnpj_str = str(cnpj).strip()
    cnpj_limpo = re.sub(r"[^0-9]", "", cnpj_str)
    if not cnpj_limpo:
        return None
    return cnpj_limpo.zfill(14)


def normalizar_revenda(nome):
    """Normaliza o nome da revenda (título, remove espaços extras)."""
    if pd.isna(nome):
        return None
    return str(nome).strip().title()


def descobrir_banco():
    """Localiza o arquivo de banco na pasta de hierarquias."""
    candidatos = sorted(HIERARQUIA_DIR.glob("PROD_WHP_Participante_Banco_v2_*.xlsx"))
    if not candidatos:
        raise FileNotFoundError(
            f"Base do banco não encontrada em: {HIERARQUIA_DIR}"
        )
    return candidatos[-1]


def ler_excel_robusto(arquivo, dtype=str):
    """Tenta ler o Excel com pandas; em caso de erro de formatação/filtro,
    cria uma cópia temporária removendo conditionalFormatting e autoFilter,
    e faz a leitura com openpyxl."""
    try:
        return pd.read_excel(arquivo, dtype=dtype)
    except (ValueError, Exception) as e:
        logger.warning(f"Falha padrão ao ler {arquivo.name}: {e}. Tentando reparar XML...")

    import tempfile
    import zipfile

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        extracted = tmp_path / "extracted"
        extracted.mkdir()

        with zipfile.ZipFile(arquivo, "r") as z:
            z.extractall(extracted)

        sheet_path = extracted / "xl" / "worksheets" / "sheet1.xml"
        if sheet_path.exists():
            content = sheet_path.read_text(encoding="utf-8")
            # Remove formatação condicional e autofiltros (causam erros no openpyxl)
            content = re.sub(
                r"<conditionalFormatting[^/]*>.*?</conditionalFormatting>",
                "",
                content,
                flags=re.DOTALL,
            )
            content = re.sub(
                r"<autoFilter[^/]*>.*?</autoFilter>",
                "",
                content,
                flags=re.DOTALL,
            )
            sheet_path.write_text(content, encoding="utf-8")

        novo_xlsx = tmp_path / "arquivo_reparado.xlsx"
        with zipfile.ZipFile(novo_xlsx, "w", zipfile.ZIP_DEFLATED) as z:
            for f in extracted.rglob("*"):
                if f.is_file():
                    arcname = f.relative_to(extracted).as_posix()
                    z.write(f, arcname)

        return pd.read_excel(novo_xlsx, dtype=dtype)


def carregar_hierarquias():
    """Consolida todos os arquivos de hierarquia em um único DataFrame."""
    dfs = []
    arquivos = sorted([
        f for f in HIERARQUIA_DIR.glob("*.xlsx")
        if "PROD_WHP" not in f.name.upper()
        and not f.name.startswith("validacao_hierarquia_banco_")
    ])
    logger.info(f"{len(arquivos)} arquivos de hierarquia encontrados")

    mapeamento_colunas = {
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

    for arquivo in arquivos:
        logger.info(f"Lendo {arquivo.name} ...")
        try:
            df = ler_excel_robusto(arquivo, dtype=str)
        except Exception as e:
            logger.warning(f"Erro ao ler {arquivo.name}: {e}")
            continue

        if df.empty:
            logger.warning(f"{arquivo.name} está vazio")
            continue

        # Renomeia colunas ignorando case e espaços extras
        df.columns = [str(c).strip().upper() for c in df.columns]
        df_renomeado = df.rename(columns=mapeamento_colunas)

        # Garante colunas mínimas
        for col in ["revenda", "cod_loja", "cnpj", "cpf", "nome"]:
            if col not in df_renomeado.columns:
                df_renomeado[col] = None

        # Origem do arquivo
        df_renomeado["arquivo_origem"] = arquivo.name

        # Normaliza CPF e CNPJ
        df_renomeado["cpf_limp"] = df_renomeado["cpf"].apply(limpar_cpf)
        df_renomeado["cnpj_limp"] = df_renomeado["cnpj"].apply(limpar_cnpj)

        # Revenda normalizada
        df_renomeado["revenda_norm"] = df_renomeado["revenda"].apply(normalizar_revenda)

        # Colunas de cargo/categoria
        for col in ["vendedor", "gerente_loja", "gerente_regional", "diretor", "cargo", "desligado"]:
            if col not in df_renomeado.columns:
                df_renomeado[col] = None
            else:
                df_renomeado[col] = df_renomeado[col].astype(str).str.strip().str.upper()

        dfs.append(df_renomeado)

    if not dfs:
        raise ValueError("Nenhuma hierarquia válida foi carregada")

    df_hier = pd.concat(dfs, ignore_index=True)

    # Colunas finais
    colunas_finais = [
        "revenda", "revenda_norm", "cod_loja", "cnpj", "cnpj_limp",
        "cpf", "cpf_limp", "nome", "cargo",
        "vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado",
        "arquivo_origem",
    ]
    colunas_presentes = [c for c in colunas_finais if c in df_hier.columns]
    df_hier = df_hier[colunas_presentes].copy()

    logger.info(f"Hierarquia consolidada: {len(df_hier):,} registros")
    return df_hier


def carregar_banco(caminho_banco):
    """Carrega a base do banco de participantes."""
    logger.info(f"Lendo base do banco: {caminho_banco.name} ...")
    df = pd.read_excel(
        caminho_banco,
        dtype={"Cpf": str, "Ativo": str},
    )

    # Renomeia colunas principais
    df = df.rename(columns={
        "Nome": "revenda_banco",
        "Nome.1": "nome",
        "Cpf": "cpf",
        "Ativo": "ativo",
        "DataInclusao": "data_inclusao",
        "DataAceiteRegulamento": "data_aceite_regulamento",
        "DataBloqueio": "data_bloqueio",
        "MotivoBloqueio": "motivo_bloqueio",
        "OrigemCadastro": "origem_cadastro",
        "Matricula": "matricula",
        "Email": "email",
        "Telefone": "telefone",
        "Celular": "celular",
    })

    # Garante colunas mínimas
    for col in ["revenda_banco", "nome", "cpf", "ativo", "data_inclusao"]:
        if col not in df.columns:
            df[col] = None

    # Normaliza CPF
    df["cpf_limp"] = df["cpf"].apply(limpar_cpf)

    # Normaliza ativo
    df["ativo_str"] = df["ativo"].apply(
        lambda x: "Sim" if str(x).strip() in ("1", "True", "SIM", "Sim", "true") else ("Não" if pd.notna(x) else None)
    )

    # Datas
    for col in ["data_inclusao", "data_aceite_regulamento", "data_bloqueio"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    colunas_finais = [
        "revenda_banco", "nome", "cpf", "cpf_limp", "ativo", "ativo_str",
        "data_inclusao", "data_aceite_regulamento", "data_bloqueio", "motivo_bloqueio",
        "origem_cadastro", "matricula", "email", "telefone", "celular",
    ]
    colunas_presentes = [c for c in colunas_finais if c in df.columns]
    df = df[colunas_presentes].copy()

    logger.info(f"Base do banco: {len(df):,} registros")
    return df


def gerar_resumo_revenda(df_hier, df_banco):
    """Gera resumo por revenda comparando hierarquia e banco."""
    hier_por_revenda = df_hier.groupby("revenda_norm").agg(
        total_hierarquia=("cpf_limp", "nunique"),
        total_registros_hier=("cpf_limp", "size"),
    ).reset_index()

    banco_por_revenda = df_banco.groupby("revenda_banco").agg(
        total_banco=("cpf_limp", "nunique"),
    ).reset_index()

    resumo = hier_por_revenda.merge(
        banco_por_revenda,
        left_on="revenda_norm",
        right_on="revenda_banco",
        how="outer",
    )
    resumo["revenda"] = resumo["revenda_norm"].fillna(resumo["revenda_banco"])
    resumo = resumo.drop(columns=["revenda_norm", "revenda_banco"])
    resumo = resumo[["revenda", "total_hierarquia", "total_banco", "total_registros_hier"]]
    resumo = resumo.sort_values("revenda").fillna(0)
    resumo["total_hierarquia"] = resumo["total_hierarquia"].astype(int)
    resumo["total_banco"] = resumo["total_banco"].astype(int)
    resumo["total_registros_hier"] = resumo["total_registros_hier"].astype(int)
    resumo["diferenca"] = resumo["total_hierarquia"] - resumo["total_banco"]

    return resumo


def main():
    caminho_banco = descobrir_banco()

    df_hier = carregar_hierarquias()
    df_banco = carregar_banco(caminho_banco)

    # Remove CPFs nulos
    df_hier_valid = df_hier[df_hier["cpf_limp"].notna()].copy()
    df_banco_valid = df_banco[df_banco["cpf_limp"].notna()].copy()

    logger.info(f"Hierarquia com CPF válido: {len(df_hier_valid):,}")
    logger.info(f"Banco com CPF válido: {len(df_banco_valid):,}")

    # CPFs únicos
    cpfs_hier = set(df_hier_valid["cpf_limp"].unique())
    cpfs_banco = set(df_banco_valid["cpf_limp"].unique())

    cpfs_somente_hier = cpfs_hier - cpfs_banco
    cpfs_somente_banco = cpfs_banco - cpfs_hier
    cpfs_ambos = cpfs_hier & cpfs_banco

    logger.info(f"CPFs só na hierarquia: {len(cpfs_somente_hier):,}")
    logger.info(f"CPFs só no banco: {len(cpfs_somente_banco):,}")
    logger.info(f"CPFs em ambos: {len(cpfs_ambos):,}")

    # Cruzamento: hierarquia à esquerda + dados do banco
    df_cruzamento = df_hier_valid.merge(
        df_banco_valid[[
            "cpf_limp", "nome", "ativo_str", "data_inclusao",
            "data_aceite_regulamento", "data_bloqueio", "motivo_bloqueio",
            "origem_cadastro", "matricula", "email",
        ]].rename(columns={"nome": "nome_banco"}),
        on="cpf_limp",
        how="left",
        suffixes=("", "_banco"),
    )

    # Só na hierarquia
    df_somente_hier = df_hier_valid[df_hier_valid["cpf_limp"].isin(cpfs_somente_hier)].copy()

    # Só no banco
    df_somente_banco = df_banco_valid[df_banco_valid["cpf_limp"].isin(cpfs_somente_banco)].copy()

    # Duplicados na hierarquia
    dup_hier = df_hier_valid[df_hier_valid.duplicated(subset=["cpf_limp"], keep=False)].copy()
    dup_hier = dup_hier.sort_values(["cpf_limp", "arquivo_origem"])

    # Duplicados no banco
    dup_banco = df_banco_valid[df_banco_valid.duplicated(subset=["cpf_limp"], keep=False)].copy()
    dup_banco = dup_banco.sort_values(["cpf_limp", "data_inclusao"])

    # Resumo por revenda
    df_resumo = gerar_resumo_revenda(df_hier_valid, df_banco_valid)

    # Gera arquivo de saída
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = SAIDA_DIR / f"validacao_hierarquia_banco_{timestamp}.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df_hier_valid.to_excel(writer, sheet_name="Hierarquia Consolidada", index=False)
        df_banco_valid.to_excel(writer, sheet_name="Banco", index=False)
        df_cruzamento.to_excel(writer, sheet_name="Cruzamento Completo", index=False)
        df_somente_hier.to_excel(writer, sheet_name="Só na Hierarquia", index=False)
        df_somente_banco.to_excel(writer, sheet_name="Só no Banco", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo por Revenda", index=False)
        dup_hier.to_excel(writer, sheet_name="Duplicados na Hierarquia", index=False)
        dup_banco.to_excel(writer, sheet_name="Duplicados no Banco", index=False)

    logger.info(f"Arquivo de validação gerado: {arquivo_saida}")
    print(f"\nArquivo gerado: {arquivo_saida}")
    print(f"Total hierarquia: {len(df_hier_valid):,}")
    print(f"Total banco: {len(df_banco_valid):,}")
    print(f"Só na hierarquia: {len(df_somente_hier):,}")
    print(f"Só no banco: {len(df_somente_banco):,}")
    print(f"Duplicados hierarquia: {len(dup_hier):,}")
    print(f"Duplicados banco: {len(dup_banco):,}")


if __name__ == "__main__":
    main()
