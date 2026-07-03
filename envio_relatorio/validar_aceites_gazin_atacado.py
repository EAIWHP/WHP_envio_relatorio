import pandas as pd
from pathlib import Path
import re

BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/bases")


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


def normalizar_revenda_hierarquia(nome_revenda):
    if pd.isna(nome_revenda):
        return nome_revenda
    nome = str(nome_revenda).strip().upper()
    nome = re.sub(r"\s+", " ", nome)
    nome = nome.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    nome = nome.replace("Â", "A").replace("Ê", "E").replace("Ô", "O")
    nome = nome.replace("Ã", "A").replace("Ç", "C")
    # Mapeamento simples para Gazin Atacado
    if "GAZIN" in nome and "ATACADO" in nome:
        return "Gazin Atacado"
    if "GAZIN" in nome and "ONLINE" in nome:
        return "Gazin Online"
    if "GAZIN" in nome and "VAREJO" in nome:
        return "Gazin Varejo"
    return nome.strip()


def descobrir_aba_aceites(xl, ano_mes):
    abas = xl.sheet_names
    ano, mes = ano_mes.split("-")
    mes_int = int(mes)
    mes_nome = {
        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ",
    }[mes_int]
    mes_nome_completo = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARCO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
        7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
    }[mes_int]
    candidatos = [
        f"{mes_nome}_{ano}", f"{mes_nome.lower()}_{ano}",
        f"{mes_nome_completo}_{ano}", f"{mes_nome_completo.lower()}_{ano}",
        f"{ano}_{mes.zfill(2)}", f"{mes.zfill(2)}_{ano}",
        ano_mes, ano_mes.replace("-", "_"),
    ]
    for aba in abas:
        aba_limpa = aba.strip().replace(" ", "_").replace("-", "_")
        if aba_limpa.upper() in [c.upper() for c in candidatos]:
            return aba, False
    return abas[-1], True


def carregar_hierarquia_gazin_atacado():
    arquivo = BASE_DIR / "hierarquia_rodrigo" / "GAZIN ATACADO - HIERARQUIA MAIO.xlsx"
    df = pd.read_excel(arquivo, sheet_name=0)
    df.columns = [str(c).strip().upper() for c in df.columns]

    mapeamento_colunas = {
        "REVENDA": "revenda", "COD LOJA": "cod_loja", "CODLOJA": "cod_loja",
        "CNPJ": "cnpj", "CPF": "cpf", "NOME": "nome_hier",
        "VENDEDOR": "vendedor", "GERENTE DE LOJA": "gerente_loja",
        "GERENTE REGIONAL": "gerente_regional", "DIRETOR": "diretor",
        "DESLIGADO": "desligado", "CARGO": "cargo_hier",
    }
    rename = {c: mapeamento_colunas[c] for c in df.columns if c in mapeamento_colunas}
    df = df.rename(columns=rename)

    revenda_arquivo = "Gazin Atacado"
    if "revenda" not in df.columns:
        df["revenda"] = revenda_arquivo
    else:
        rev_vazia = df["revenda"].isna() | (df["revenda"].astype(str).str.strip() == "")
        df.loc[rev_vazia, "revenda"] = revenda_arquivo

    df["cpf_limp"] = df["cpf"].apply(limpar_cpf)
    df["revenda_norm"] = df["revenda"].apply(normalizar_revenda_hierarquia)

    if "desligado" in df.columns:
        df["desligado"] = df["desligado"].astype(str).str.strip().str.upper().str.replace("Ã", "A")

    return df


def carregar_aceites(ano_mes="2026-06"):
    aceite_path = BASE_DIR / "WHP_Aceite_Mensal_OUT_NOV_DEZ_2025_JAN_FEV_2026.xlsx"
    xl = pd.ExcelFile(aceite_path)
    aba, usou_ultima = descobrir_aba_aceites(xl, ano_mes)
    df = pd.read_excel(aceite_path, sheet_name=aba)
    df["cpf_limp"] = df["CPF"].apply(limpar_cpf)
    if "DataAceite" in df.columns:
        df["DataAceite"] = pd.to_datetime(df["DataAceite"], errors="coerce")
        df["mes_aceite"] = df["DataAceite"].dt.to_period("M")
    return df, aba, usou_ultima


def main():
    print("=" * 80)
    print("VALIDAÇÃO ACEITES - GAZIN ATACADO")
    print("=" * 80)

    # 1. Hierarquia bruta
    df_hier = carregar_hierarquia_gazin_atacado()
    print(f"\n1. Hierarquia bruta (GAZIN ATACADO - HIERARQUIA MAIO.xlsx):")
    print(f"   Total de registros: {len(df_hier):,}")
    print(f"   CPFs únicos: {df_hier['cpf_limp'].nunique():,}")
    print(f"   Colunas: {list(df_hier.columns)}")
    if "desligado" in df_hier.columns:
        print(f"\n   Distribuição DESLIGADO:")
        print(df_hier["desligado"].value_counts(dropna=False).to_string())

    # 2. Hierarquia filtrada (DESLIGADO=NÃO) - base do relatório
    if "desligado" in df_hier.columns:
        deslig_norm = df_hier["desligado"].astype(str).str.strip().str.upper().str.replace("Ã", "A")
        df_hier_ativos = df_hier[deslig_norm.isin(["NAO", "NÃO"])].copy()
        print(f"\n2. Hierarquia filtrada (DESLIGADO=NÃO):")
        print(f"   Registros: {len(df_hier_ativos):,}")
        print(f"   CPFs únicos: {df_hier_ativos['cpf_limp'].nunique():,}")
    else:
        df_hier_ativos = df_hier.copy()

    # 3. Base de aceites
    df_aceite, aba, usou_ultima = carregar_aceites("2026-06")
    print(f"\n3. Base de aceites:")
    print(f"   Aba usada: '{aba}' (usou última={usou_ultima})")
    print(f"   Total de registros na aba: {len(df_aceite):,}")
    print(f"   CPFs únicos na aba: {df_aceite['cpf_limp'].nunique():,}")
    print(f"   Colunas: {list(df_aceite.columns)}")

    # Mês de referência
    if "mes_aceite" in df_aceite.columns:
        print(f"\n   Distribuição por mês de aceite:")
        print(df_aceite["mes_aceite"].value_counts(dropna=False).sort_index().to_string())
        mes_ref = pd.Period("2026-06", freq="M")
        aceite_mes = df_aceite[df_aceite["mes_aceite"] == mes_ref].copy()
        print(f"\n   Aceites em {mes_ref}: {len(aceite_mes):,} registros, {aceite_mes['cpf_limp'].nunique():,} CPFs únicos")
    else:
        aceite_mes = df_aceite.copy()

    cpfs_aceitaram = set(aceite_mes["cpf_limp"].unique())

    # 4. Cruzamentos
    cpfs_hier_ativos = set(df_hier_ativos["cpf_limp"].unique())
    cpfs_hier_bruta = set(df_hier["cpf_limp"].unique())

    intersecao_ativos_aceites = cpfs_hier_ativos & cpfs_aceitaram
    intersecao_bruta_aceites = cpfs_hier_bruta & cpfs_aceitaram
    aceites_fora_hier_bruta = cpfs_aceitaram - cpfs_hier_bruta
    aceites_fora_hier_ativos = cpfs_aceitaram - cpfs_hier_ativos

    print(f"\n4. CRUZAMENTOS:")
    print(f"   a) CPFs únicos na base de aceites (aba '{aba}'): {len(cpfs_aceitaram):,}")
    print(f"   b) CPFs únicos na hierarquia bruta: {len(cpfs_hier_bruta):,}")
    print(f"   c) CPFs únicos na hierarquia DESLIGADO=NÃO: {len(cpfs_hier_ativos):,}")
    print(f"   d) Aceites ∩ Hierarquia DESLIGADO=NÃO: {len(intersecao_ativos_aceites):,}")
    print(f"   e) Aceites ∩ Hierarquia bruta: {len(intersecao_bruta_aceites):,}")
    print(f"   f) Aceites FORA da hierarquia bruta: {len(aceites_fora_hier_bruta):,}")
    print(f"   g) Aceites FORA da hierarquia DESLIGADO=NÃO: {len(aceites_fora_hier_ativos):,}")

    # 5. Diferença entre intersecao_ativos_aceites e intersecao_bruta_aceites
    diff = intersecao_ativos_aceites - intersecao_bruta_aceites
    diff2 = intersecao_bruta_aceites - intersecao_ativos_aceites
    print(f"\n5. DIFERENÇAS ENTRE INTERSEÇÕES:")
    print(f"   (d) - (e): {len(diff)} CPFs")
    print(f"   (e) - (d): {len(diff2)} CPFs")

    if diff2:
        print(f"\n   CPFs que aceitaram e aparecem na hierarquia bruta, mas NÃO na hierarquia DESLIGADO=NÃO:")
        for cpf in sorted(diff2)[:20]:
            info = df_hier[df_hier["cpf_limp"] == cpf][["cpf_limp", "nome_hier", "desligado"]].drop_duplicates()
            print(f"     {cpf} - {info.to_dict('records')}")

    # 6. CPFs na base de aceites que não aparecem na hierarquia bruta
    if aceites_fora_hier_bruta:
        print(f"\n6. CPFs na base de aceites que NÃO aparecem na hierarquia bruta (primeiros 30):")
        for cpf in sorted(aceites_fora_hier_bruta)[:30]:
            info = aceite_mes[aceite_mes["cpf_limp"] == cpf][["cpf_limp", "Nome" if "Nome" in aceite_mes.columns else "CPF", "DataAceite"]].drop_duplicates()
            print(f"     {cpf} - {info.to_dict('records')}")

    # 7. Verifica se há CPFs duplicados na base de aceites
    duplicados_aceite = aceite_mes[aceite_mes.duplicated(subset=["cpf_limp"], keep=False)]
    print(f"\n7. DUPLICIDADES NA BASE DE ACEITES:")
    print(f"   Registros duplicados por CPF: {len(duplicados_aceite):,}")
    print(f"   CPFs duplicados: {duplicados_aceite['cpf_limp'].nunique():,}")
    if len(duplicados_aceite) > 0:
        print("   Exemplos de CPFs duplicados:")
        for cpf in duplicados_aceite["cpf_limp"].unique()[:10]:
            infos = aceite_mes[aceite_mes["cpf_limp"] == cpf][["cpf_limp", "Nome" if "Nome" in aceite_mes.columns else "CPF", "DataAceite"]].to_dict("records")
            print(f"     {cpf}: {infos}")

    # 8. Cálculo final do relatório
    base = df_hier_ativos[["cpf_limp", "revenda_norm"]].drop_duplicates().copy()
    base["aceitou"] = base["cpf_limp"].isin(cpfs_aceitaram)
    total = base["cpf_limp"].nunique()
    aceitaram = base["aceitou"].sum()
    nao_aceitaram = total - aceitaram
    pct = round(aceitaram / total * 100, 1) if total else 0
    print(f"\n8. RESULTADO FINAL DO RELATÓRIO (base = hierarquia DESLIGADO=NÃO):")
    print(f"   Revenda: Gazin Atacado")
    print(f"   Total na hierarquia: {total}")
    print(f"   Aceitaram: {aceitaram}")
    print(f"   Não aceitaram: {nao_aceitaram}")
    print(f"   % Aceite: {pct}%")


if __name__ == "__main__":
    main()
