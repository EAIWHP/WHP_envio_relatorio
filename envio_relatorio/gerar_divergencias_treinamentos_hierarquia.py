import pandas as pd
from pathlib import Path
import re

BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/bases")
OUTPUT_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/divergencias_treinamentos_vs_hierarquia")
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "divergencias_treinamentos_vs_hierarquia.xlsx"


def limpar_cpf(cpf):
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


def normalizar_texto(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    s = s.replace("Â", "A").replace("Ê", "E").replace("Ô", "O")
    s = s.replace("Ã", "A").replace("Ç", "C")
    return s


def revenda_do_nome_arquivo(stem):
    nome = stem.split("-")[0].strip()
    nome = normalizar_texto(nome)
    return nome.title()


def agrupar_revenda(nome_revenda):
    nome = str(nome_revenda).strip()
    nome_norm = normalizar_texto(nome)
    nome_agrupado = re.sub(r"\s+[A-Z]$", "", nome_norm).strip()
    if nome_agrupado:
        return nome_agrupado.title()
    return nome.title()


def carregar_todas_hierarquias():
    hier_dir = BASE_DIR / "hierarquia_rodrigo"
    registros = []
    for arquivo in sorted(hier_dir.glob("*.xlsx")):
        if arquivo.name.startswith("~"):
            continue
        try:
            df = pd.read_excel(arquivo, sheet_name=0)
        except Exception as e:
            print(f"Ignorado {arquivo.name}: {e}")
            continue

        df.columns = [str(c).strip().upper() for c in df.columns]
        mapeamento = {
            "REVENDA": "revenda_col", "COD LOJA": "cod_loja", "CODLOJA": "cod_loja",
            "CNPJ": "cnpj", "CPF": "cpf", "NOME": "nome_hier",
            "VENDEDOR": "vendedor", "GERENTE DE LOJA": "gerente_loja",
            "GERENTE REGIONAL": "gerente_regional", "DIRETOR": "diretor",
            "DESLIGADO": "desligado", "CARGO": "cargo_hier",
        }
        rename = {c: mapeamento[c] for c in df.columns if c in mapeamento}
        df = df.rename(columns=rename)

        if "cpf" not in df.columns:
            continue

        rev_arquivo = revenda_do_nome_arquivo(arquivo.stem)
        if "revenda_col" not in df.columns:
            df["revenda_col"] = rev_arquivo
        else:
            rev_vazia = df["revenda_col"].isna() | (df["revenda_col"].astype(str).str.strip() == "")
            df.loc[rev_vazia, "revenda_col"] = rev_arquivo

        df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
        df["revenda_col_norm"] = df["revenda_col"].apply(normalizar_texto).str.title()
        df["revenda_col_agrup"] = df["revenda_col_norm"].apply(agrupar_revenda)
        df["arquivo"] = arquivo.name

        if "desligado" in df.columns:
            df["desligado"] = df["desligado"].astype(str).str.strip().str.upper().str.replace("Ã", "A")
        else:
            df["desligado"] = "NÃO INFORMADO"

        registros.append(df)
        print(f"Carregado {arquivo.name}: {len(df)} registros")

    if not registros:
        raise ValueError("Nenhuma hierarquia carregada")

    df_hier = pd.concat(registros, ignore_index=True)
    return df_hier


def carregar_treinamentos(mes_ref="2026-06"):
    trein_path = BASE_DIR / "Base_treinamentos.xlsx"
    df = pd.read_excel(trein_path, sheet_name=0)
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf)
    df["Conclusão"] = pd.to_datetime(df["Conclusão"], errors="coerce")
    df["mes_conclusao"] = df["Conclusão"].dt.to_period("M")

    # Considera apenas o mês de referência
    mes_period = pd.Period(mes_ref, freq="M")
    df_mes = df[df["mes_conclusao"] == mes_period].copy()

    # Revenda = Distribuidor
    df_mes["Revenda"] = df_mes["Distribuidor"]
    df_mes["Revenda_norm"] = df_mes["Revenda"].apply(normalizar_texto).str.title()
    df_mes["Revenda_agrup"] = df_mes["Revenda_norm"].apply(agrupar_revenda)

    print(f"\nBase de treinamentos: mês {mes_ref} ({len(df_mes):,} registros, {df_mes['cpf_limp'].nunique():,} CPFs únicos)")
    return df_mes


def ajustar_larguras(ws):
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 50)


def main():
    print("Carregando hierarquias...")
    df_hier = carregar_todas_hierarquias()
    print(f"Total hierarquias consolidadas: {len(df_hier):,} registros")

    df_trein = carregar_treinamentos("2026-06")

    # Revendas que não devem compor o relatório de divergências
    REVENDAS_EXCLUIR = {"EAI", "WHIRLPOOL", "NOVO MUNDO", "ELETROMOVEIS MARTINELLO", "ELETROMÓVEIS MARTINELLO"}
    antes_exclusao = len(df_trein)
    df_trein = df_trein[
        ~df_trein["Revenda_norm"].apply(normalizar_texto).isin(REVENDAS_EXCLUIR)
    ].copy()
    print(f"Revendas excluídas: {antes_exclusao - len(df_trein):,} registros removidos")

    # Remove CPFs duplicados dentro da mesma revenda (mantém um registro por CPF/revenda)
    df_trein_unicos = df_trein.drop_duplicates(subset=["cpf_limp", "Revenda_agrup"], keep="first").copy()
    print(f"CPFs únicos por revenda após deduplicação: {len(df_trein_unicos):,}")

    # Para cada CPF, monta lista de revendas encontradas na hierarquia (agrupadas)
    hier_por_cpf = (
        df_hier.groupby("cpf_limp")
        .agg(
            revendas_hier=("revenda_col_agrup", lambda x: " | ".join(sorted(set(x.astype(str))))),
            desligados=("desligado", lambda x: " | ".join(sorted(set(x.astype(str))))),
            arquivos=("arquivo", lambda x: " | ".join(sorted(set(x)))),
            nome_hier=("nome_hier", "first"),
        )
        .reset_index()
    )

    df_merged = df_trein_unicos.merge(hier_por_cpf, on="cpf_limp", how="left")

    def classificar_divergencia(row):
        rev_trein = str(row["Revenda_agrup"]).strip().lower()
        rev_hier = str(row["revendas_hier"]).strip().lower()

        if pd.isna(row["revendas_hier"]):
            return "NÃO ENCONTRADO NA HIERARQUIA"

        if rev_trein == rev_hier:
            return "OK"

        if rev_trein in rev_hier.split(" | "):
            return "OK"

        return "DIVERGENTE"

    df_merged["situacao"] = df_merged.apply(classificar_divergencia, axis=1)

    # Resumo agrupado
    resumo = df_merged.groupby(["Revenda_agrup", "situacao"]).size().unstack(fill_value=0)
    for col in ["OK", "DIVERGENTE", "NÃO ENCONTRADO NA HIERARQUIA"]:
        if col not in resumo.columns:
            resumo[col] = 0
    resumo["Total"] = resumo.sum(axis=1)
    resumo = resumo.reset_index()

    revendas_divergentes = resumo[
        (resumo["DIVERGENTE"] > 0) | (resumo["NÃO ENCONTRADO NA HIERARQUIA"] > 0)
    ]["Revenda_agrup"].tolist()

    print(f"\nRevendas agrupadas com divergência: {len(revendas_divergentes)}")
    for r in sorted(revendas_divergentes):
        qtd = resumo[resumo["Revenda_agrup"] == r][["DIVERGENTE", "NÃO ENCONTRADO NA HIERARQUIA"]].sum(axis=1).iloc[0]
        print(f"  - {r}: {int(qtd)} divergências")

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        ajustar_larguras(writer.sheets["Resumo"])

        for revenda in sorted(revendas_divergentes):
            df_rev = df_merged[
                (df_merged["Revenda_agrup"] == revenda)
                & (df_merged["situacao"] != "OK")
            ].copy()

            cols = [
                "cpf_limp", "Participante", "Revenda", "Revenda_agrup",
                "Curso", "Conclusão", "Trilha", "Estado",
                "revendas_hier", "desligados", "arquivos", "nome_hier", "situacao"
            ]
            cols = [c for c in cols if c in df_rev.columns]
            df_rev = df_rev[cols]
            df_rev = df_rev.sort_values(["situacao", "Participante"])

            aba_nome = re.sub(r'[\\/*?:\[\]]', '_', revenda)[:31]
            df_rev.to_excel(writer, sheet_name=aba_nome, index=False)
            ajustar_larguras(writer.sheets[aba_nome])

    print(f"\nArquivo gerado: {OUTPUT_FILE}")
    print(f"Total de abas: {len(revendas_divergentes) + 1} (Resumo + revendas agrupadas)")


if __name__ == "__main__":
    main()
