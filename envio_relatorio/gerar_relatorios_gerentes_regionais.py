#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera relatórios individuais para cada Gerente Regional usando a hierarquia como fonte.

Regra de responsabilidade:
  • O gerente regional é identificado pela coluna GERENTE REGIONAL = SIM na hierarquia.
  • Um gerente regional pode ser responsável por várias lojas.
  • A equipe do gerente = todos os CPFs das lojas onde ele está marcado como GERENTE REGIONAL = SIM.

Cada relatório contém:
  • Aba Resumo: dados do gerente regional + indicadores operacionais.
  • Aba Detalhamento: equipe do gerente com CPF, status de cadastro e treinamentos obrigatórios do mês.

Os arquivos são salvos em:
  Programa_mais_top/relatorios_gerentes_regionais/

Uso:
  python3 gerar_relatorios_gerentes_regionais.py
"""

import logging
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import numpy as np

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from gerar_relatorio_semanal import (
    carregar_bases,
    detectar_cursos_obrigatorios,
    normalizar_revenda_hierarquia,
    read_excel_robusto,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR.parent / "relatorios_gerentes_regionais"
HIERARQUIA_DIR = BASE_DIR / "bases" / "bases_cadastro_hierarquia"


def formatar_mes_referencia(mes_referencia):
    """Converte '2026-06' para 'Junho/2026'."""
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
    }
    p = pd.Period(mes_referencia, freq="M")
    return f"{meses_pt[p.month]}/{p.year}"


def formatar_mes_referencia_curto(mes_referencia):
    """Converte '2026-06' para 'Junho 26'."""
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
    }
    p = pd.Period(mes_referencia, freq="M")
    return f"{meses_pt[p.month]} {str(p.year)[-2:]}"


def formatar_percentual(valor):
    """Formata decimal como percentual com uma casa (ex: 0.429 -> 42.9%)."""
    if pd.isna(valor):
        return "0.0%"
    return f"{valor * 100:.1f}%"


def limpar_cpf(x):
    """Remove caracteres não numéricos do CPF e retorna como string."""
    if pd.isna(x):
        return None
    s = re.sub(r"[^0-9]", "", str(x).strip())
    return s if s else None


def formatar_cpf(cpf):
    """Garante que o CPF tenha 11 dígitos como texto, tratando floats do Excel."""
    if cpf is None or pd.isna(cpf):
        return None
    s = str(cpf).strip()
    # Remove sufixo .0 de floats lidos do Excel antes de remover não-numéricos
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"[^0-9]", "", s)
    return s.zfill(11) if s else None


def chave_loja(row):
    """Cria chave única de loja usando COD LOJA quando existir, senão CNPJ."""
    revenda = str(row.get("revenda", "")).strip()
    cod_loja = row.get("cod_loja")
    cnpj = row.get("cnpj")
    if pd.notna(cod_loja) and str(cod_loja).strip() not in ("", "nan", "None"):
        return f"{revenda}|{cod_loja}"
    if pd.notna(cnpj) and str(cnpj).strip() not in ("", "nan", "None"):
        return f"{revenda}|{cnpj}"
    return None


def carregar_hierarquia_completa():
    """
    Carrega todos os arquivos de hierarquia SEM deduplicar CPFs.
    Mantém todas as ocorrências de gerentes regionais em várias lojas.
    """
    if not HIERARQUIA_DIR.exists():
        logger.warning(f"Pasta de hierarquias não encontrada: {HIERARQUIA_DIR}")
        return None

    mapeamento_colunas = {
        "REVENDA": "revenda",
        "COD LOJA": "cod_loja",
        "CODLOJA": "cod_loja",
        "CNPJ": "cnpj",
        "CPF": "cpf",
        "NOME": "nome_hier",
        "VENDEDOR": "vendedor",
        "GERENTE DE LOJA": "gerente_loja",
        "GERENTE REGIONAL": "gerente_regional",
        "DIRETOR": "diretor",
        "DESLIGADO": "desligado",
        "CARGO": "cargo_hier",
    }

    dfs = []
    arquivos_brutos = sorted(HIERARQUIA_DIR.glob("*.xlsx"))
    arquivos_corrigidos = {a.stem.replace("_corrigido", ""): a for a in arquivos_brutos if a.stem.endswith("_corrigido")}
    arquivos_usar = []
    for arquivo in arquivos_brutos:
        nome = arquivo.name.upper()
        if "PROD_WHP" in nome or nome.startswith("~") or "_corrigido" in nome:
            continue
        if arquivo.stem in arquivos_corrigidos:
            arquivos_usar.append(arquivos_corrigidos[arquivo.stem])
        else:
            arquivos_usar.append(arquivo)
    arquivos_usar = list(dict.fromkeys(arquivos_usar))

    for arquivo in arquivos_usar:
        nome = arquivo.name.upper()
        if any(x in nome for x in ["CONSOLIDADO", "COMPARACAO", "RESUMO", "CPFS_REPETIDOS"]):
            continue
        try:
            xl = read_excel_robusto(arquivo, sheet_name=None)
            primeira_aba = list(xl.keys())[0]
            df = xl[primeira_aba]
        except Exception as e:
            logger.warning(f"Não foi possível ler {arquivo.name}: {e}")
            continue

        df.columns = [str(c).strip().upper() for c in df.columns]
        rename = {c: mapeamento_colunas[c] for c in df.columns if c in mapeamento_colunas}
        df = df.rename(columns=rename)

        if "cpf" not in df.columns:
            logger.warning(f"{arquivo.name} não possui coluna CPF. Ignorado.")
            continue

        revenda_arquivo = normalizar_revenda_hierarquia(arquivo.stem.split("-")[0].strip())
        if "revenda" not in df.columns:
            df["revenda"] = revenda_arquivo
        else:
            rev_vazia = df["revenda"].isna() | (df["revenda"].astype(str).str.strip() == "")
            df.loc[rev_vazia, "revenda"] = revenda_arquivo

        df["cpf_limp"] = df["cpf"].apply(formatar_cpf)

        for col in ["vendedor", "gerente_loja", "gerente_regional", "diretor", "desligado"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper().str.replace("Ã", "A")

        dfs.append(df)
        logger.info(f"Hierarquia {arquivo.name}: {len(df)} registros")

    if not dfs:
        logger.warning("Nenhum arquivo de hierarquia válido encontrado.")
        return None

    df_hier = pd.concat(dfs, ignore_index=True)
    df_hier["revenda"] = df_hier["revenda"].apply(normalizar_revenda_hierarquia)

    # Aplica filtro DESLIGADO = NÃO (mesma regra do envio_relatorio)
    # Se um mesmo CPF tiver qualquer registro desligado (SIM/FÉRIAS/BENEFÍCIO),
    # remove TODOS os registros desse CPF para não aparecer nos relatórios.
    if "desligado" in df_hier.columns:
        antes = len(df_hier)
        deslig_norm = df_hier["desligado"].astype(str).str.strip().str.upper().str.replace("Ã", "A")
        cpfs_desligados = set(df_hier.loc[~deslig_norm.isin(["NAO", "NÃO"]), "cpf_limp"].dropna().unique())
        df_hier = df_hier[~df_hier["cpf_limp"].isin(cpfs_desligados)].copy()
        logger.info(f"Filtro DESLIGADO=NÃO: {antes - len(df_hier)} registros removidos ({len(cpfs_desligados)} CPFs desligados)")

    return df_hier


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    logger.info("Carregando bases...")

    bases = carregar_bases()
    df_cad = bases["cadastro"]
    df_trein = bases["treinamentos"]
    mes_referencia = bases["mes_referencia"]

    # Garante compatibilidade de CPF entre cadastro e hierarquia
    df_cad["cpf_limp"] = df_cad["cpf_limp"].apply(formatar_cpf)

    # Carrega hierarquia completa sem deduplicação
    df_hier = carregar_hierarquia_completa()
    if df_hier is None or df_hier.empty:
        logger.warning("Nenhuma hierarquia carregada.")
        return

    # Normaliza colunas de loja
    def _norm_cod_loja(x):
        if pd.isna(x):
            return None
        s = str(x).strip().replace(".0", "")
        if s.lower() in ("", "nan"):
            return None
        return s

    def _norm_cnpj(x):
        if pd.isna(x):
            return None
        s = re.sub(r"[^0-9]", "", str(x).strip())
        return s if s else None

    df_hier["cod_loja"] = df_hier["cod_loja"].apply(_norm_cod_loja)
    df_hier["cnpj"] = df_hier["cnpj"].apply(_norm_cnpj)
    df_hier["chave_loja"] = df_hier.apply(chave_loja, axis=1)

    # ------------------------------------------------------------------
    # Identifica gerentes regionais e todas as lojas de cada um
    # ------------------------------------------------------------------
    mask_gerente = df_hier["gerente_regional"].astype(str).str.strip().str.upper() == "SIM"
    df_gerentes_registros = df_hier[mask_gerente].copy()

    if df_gerentes_registros.empty:
        logger.warning("Nenhum gerente regional encontrado na hierarquia (GERENTE REGIONAL = SIM).")
        return

    logger.info(f"{len(df_gerentes_registros)} registros de gerente regional encontrados na hierarquia")

    # Agrupa por CPF do gerente, coletando todas as lojas
    def _coletar_lojas(grupo):
        lojas = set()
        for _, row in grupo.iterrows():
            if pd.notna(row["cod_loja"]):
                lojas.add(("cod_loja", row["cod_loja"]))
            elif pd.notna(row["cnpj"]):
                lojas.add(("cnpj", row["cnpj"]))
        return list(lojas)

    df_gerentes = df_gerentes_registros.groupby("cpf_limp").agg({
        "nome_hier": "first",
        "revenda": "first",
        "chave_loja": lambda x: sorted(set(x.dropna())),
    }).reset_index()

    # Cruza com cadastro para nome do gerente
    df_gerentes = df_gerentes.merge(
        df_cad[["cpf_limp", "nome"]].drop_duplicates("cpf_limp"),
        on="cpf_limp",
        how="left",
    )
    df_gerentes["nome"] = df_gerentes["nome"].fillna(df_gerentes["nome_hier"])

    logger.info(f"{len(df_gerentes)} gerentes regionais únicos")

    # ------------------------------------------------------------------
    # Prepara dados de treinamento
    # ------------------------------------------------------------------
    curso1, curso2, nome1, nome2, sku1, sku2 = detectar_cursos_obrigatorios(
        df_trein, mes_referencia
    )

    mes_dt = pd.Period(mes_referencia, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ]
    cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique()) if curso1 else set()
    cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique()) if curso2 else set()

    mes_curto = formatar_mes_referencia_curto(mes_referencia)
    nome_col1 = f"Treinamento - Conteúdo Obrigatório | {mes_curto} | {sku1}"
    nome_col2 = f"Treinamento - Conteúdo Obrigatório | {mes_curto} | {sku2}"

    # ------------------------------------------------------------------
    # Gera um relatório por gerente regional
    # ------------------------------------------------------------------
    arquivos_gerados = []

    for _, gerente in df_gerentes.iterrows():
        cpf_gerente = gerente["cpf_limp"]
        revenda = gerente["revenda"]
        chaves_loja = gerente["chave_loja"] if isinstance(gerente["chave_loja"], list) else []

        # Equipe = todos os CPFs das lojas onde o gerente é regional
        mask_equipe = df_hier["chave_loja"].isin(chaves_loja)

        if not mask_equipe.any():
            logger.warning(
                f"Nenhum CPF encontrado para as lojas do gerente {gerente['nome']} ({revenda})"
            )
            continue

        equipe_hier = df_hier[mask_equipe].copy()

        # Cruza com cadastro
        equipe = equipe_hier.merge(
            df_cad[["cpf_limp", "nome", "cargo", "status"]].drop_duplicates("cpf_limp"),
            on="cpf_limp",
            how="left",
        )
        equipe["nome"] = equipe["nome"].fillna(equipe["nome_hier"])
        equipe["cargo"] = equipe["cargo"].fillna(equipe.get("cargo_hier"))
        equipe["status"] = equipe["status"].fillna("Não cadastrado")

        # Remove o próprio gerente regional da equipe (não conta para os indicadores)
        equipe = equipe[equipe["cpf_limp"] != cpf_gerente].copy()

        # Flags de treinamento
        equipe[nome_col1] = np.where(equipe["cpf_limp"].isin(cpf_curso1), "Sim", "Não")
        equipe[nome_col2] = np.where(equipe["cpf_limp"].isin(cpf_curso2), "Sim", "Não")
        equipe["Fez os 2 Treinamentos"] = np.where(
            equipe["cpf_limp"].isin(cpf_curso1) & equipe["cpf_limp"].isin(cpf_curso2),
            "Sim", "Não"
        )

        mes_legivel = formatar_mes_referencia(mes_referencia)
        equipe["Mês de Referência"] = mes_legivel

        equipe = equipe.rename(columns={
            "cpf_limp": "CPF",
            "revenda": "Revenda",
            "cod_loja": "Código da Loja",
            "cnpj": "CNPJ",
            "nome": "Nome",
            "cargo": "Cargo",
            "status": "Status Cadastro",
        })

        # Normaliza CPF para garantir consistência entre Resumo, Lojas e Detalhamento
        equipe["CPF"] = equipe["CPF"].apply(formatar_cpf)

        # Remove vínculos sem CPF válido
        equipe = equipe[equipe["CPF"].notna()].copy()

        # Calcula indicadores do Resumo sobre CPFs distintos
        cpfs_distintos = equipe.drop_duplicates(subset=["CPF"])
        total_equipe = len(cpfs_distintos)
        ativos = (cpfs_distintos["Status Cadastro"] == "Ativo").sum()
        pct_ativos = ativos / total_equipe if total_equipe > 0 else 0
        fez_2 = (cpfs_distintos["Fez os 2 Treinamentos"] == "Sim").sum()
        fez_curso1 = (cpfs_distintos[nome_col1] == "Sim").sum()
        fez_curso2 = (cpfs_distintos[nome_col2] == "Sim").sum()
        pct_treinados_ativos = (fez_2 / ativos) if ativos > 0 else 0
        pct_treinados_total = (fez_2 / total_equipe) if total_equipe > 0 else 0

        # Indicadores por loja
        indicadores_lojas = []
        for chave in chaves_loja:
            parte = equipe[equipe["chave_loja"] == chave]
            if parte.empty:
                continue
            cpfs_loja = parte.drop_duplicates(subset=["CPF"])
            cod_loja_val = cpfs_loja["Código da Loja"].dropna().astype(str).unique()
            cnpj_val = cpfs_loja["CNPJ"].dropna().astype(str).unique()
            total_loja = len(cpfs_loja)
            ativos_loja = (cpfs_loja["Status Cadastro"] == "Ativo").sum()
            fez2_loja = (cpfs_loja["Fez os 2 Treinamentos"] == "Sim").sum()
            fez1_loja = (cpfs_loja[nome_col1] == "Sim").sum()
            fez2curso_loja = (cpfs_loja[nome_col2] == "Sim").sum()
            indicadores_lojas.append({
                "Código da Loja": ", ".join(cod_loja_val) if len(cod_loja_val) else None,
                "CNPJ": ", ".join(cnpj_val) if len(cnpj_val) else None,
                "Total de CPFs": total_loja,
                "Ativos no Cadastro": ativos_loja,
                "% Ativos": formatar_percentual(ativos_loja / total_loja if total_loja > 0 else 0),
                "Fez os 2 Treinamentos": fez2_loja,
                nome_col1: fez1_loja,
                nome_col2: fez2curso_loja,
            })

        df_lojas = pd.DataFrame(indicadores_lojas)
        df_lojas = df_lojas.sort_values(["Código da Loja", "CNPJ"]).reset_index(drop=True)

        mes_nome = formatar_mes_referencia(mes_referencia).split("/")[0]

        # Lojas do gerente para o Resumo (lista)
        cod_lojas = sorted({str(c).strip() for c in equipe["Código da Loja"].dropna().unique() if str(c).strip() not in ("", "nan", "None")})
        cnpjs = sorted({str(c).strip() for c in equipe["CNPJ"].dropna().unique() if str(c).strip() not in ("", "nan", "None")})
        lojas_str = ", ".join(cod_lojas) if cod_lojas else ", ".join(cnpjs)

        df_resumo = pd.DataFrame({
            "Indicador": [
                "Regional",
                "CPF do Regional",
                "Revenda",
                "Loja(s) do Regional",
                "Mês de Referência",
                "Quantidade de Lojas",
                "Total da Equipe",
                "Ativos no Cadastro",
                "% Ativos",
                f"Fez os 2 Treinamentos de {mes_nome}",
                "% Treinados (sobre ativos)",
                "% Treinados (sobre total da equipe)",
                f"Fez 'Conteúdo Obrigatório | {mes_curto} | {sku1}'",
                f"Fez 'Conteúdo Obrigatório | {mes_curto} | {sku2}'",
            ],
            "Valor": [
                gerente["nome"],
                gerente["cpf_limp"],
                revenda,
                lojas_str,
                mes_legivel,
                len(df_lojas),
                total_equipe,
                ativos,
                formatar_percentual(pct_ativos),
                fez_2,
                formatar_percentual(pct_treinados_ativos),
                formatar_percentual(pct_treinados_total),
                fez_curso1,
                fez_curso2,
            ],
        })

        colunas_det = [
            "Revenda", "Código da Loja", "CNPJ", "CPF", "Nome", "Cargo",
            "Status Cadastro", nome_col1, nome_col2,
            "Fez os 2 Treinamentos", "Mês de Referência",
        ]
        df_det = equipe[colunas_det].copy()
        df_det = df_det.sort_values(["Nome"])

        # Garante que CPF/CNPJ sejam salvos como texto no Excel
        df_det["CPF"] = df_det["CPF"].apply(lambda x: str(int(x)).zfill(11) if pd.notna(x) else x)
        df_det["CNPJ"] = df_det["CNPJ"].apply(lambda x: str(int(x)).zfill(14) if pd.notna(x) else x)

        df_lojas["CNPJ"] = df_lojas["CNPJ"].apply(lambda x: formatar_cpf(x) if pd.notna(x) else x)

        # Nome do arquivo
        nome_revenda = re.sub(r"[^\w\s-]", "", str(revenda)).strip().replace(" ", "_")
        nome_limpo = re.sub(r"[^\w\s-]", "", str(gerente["nome"])).strip()
        partes_nome = nome_limpo.split()
        primeiro_nome = partes_nome[0] if partes_nome else ""
        primeiro_sobrenome = partes_nome[1] if len(partes_nome) > 1 else ""
        nome_arquivo = "_".join([p for p in [primeiro_nome, primeiro_sobrenome] if p]).replace(" ", "_")

        data_str = date.today().strftime("%d_%m_%Y")
        caminho = OUTPUT_DIR / (
            f"relatorio_regional_{nome_revenda.lower()}_{nome_arquivo.lower()}_{data_str}.xlsx"
        )

        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
            df_lojas.to_excel(writer, sheet_name="Lojas", index=False)
            df_det.to_excel(writer, sheet_name="Detalhamento", index=False)

            for sheet_name, ws in writer.sheets.items():
                for col in ws.columns:
                    max_length = max(len(str(cell.value or "")) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

                # Formata colunas de CPF/CNPJ como texto
                if sheet_name == "Detalhamento":
                    for cell in ws["D"]:
                        cell.number_format = "@"
                    for cell in ws["C"]:
                        cell.number_format = "@"
                elif sheet_name == "Resumo":
                    # CPF do Regional está na célula B3
                    ws["B3"].number_format = "@"
                elif sheet_name == "Lojas":
                    for cell in ws["B"]:
                        cell.number_format = "@"

        arquivos_gerados.append(caminho.name)
        logger.info(
            f"Gerado: {caminho.name} ({len(df_det)} vínculos, {total_equipe} CPFs distintos, "
            f"{len(df_lojas)} lojas) - {gerente['nome']} | {revenda}"
        )

    logger.info(f"Total de relatórios gerados: {len(arquivos_gerados)}")
    logger.info(f"Pasta: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
