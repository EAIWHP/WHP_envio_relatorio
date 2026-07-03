import pandas as pd
from pathlib import Path
import re
import sys

sys.path.insert(0, "/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio")

from gerar_relatorio_semanal import carregar_bases

OUTPUT_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/validacao_calculos_email")
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "cpfs_treinamento_fora_hierarquia.xlsx"


def normalizar_texto(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    s = s.replace("Â", "A").replace("Ê", "E").replace("Ô", "O")
    s = s.replace("Ã", "A").replace("Ç", "C")
    return s


def agrupar_revenda(nome_revenda):
    nome = str(nome_revenda).strip()
    nome_norm = normalizar_texto(nome)
    nome_agrupado = re.sub(r"\s+[A-Z]$", "", nome_norm).strip()
    if nome_agrupado:
        return nome_agrupado.title()
    return nome.title()


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
    print("Carregando bases usando a mesma rotina do relatório semanal...")
    dados = carregar_bases()

    df_trein = dados["treinamentos"]
    df_hier = dados["hierarquia"]
    mes_referencia = dados["mes_referencia"]

    print(f"Mês de referência: {mes_referencia}")

    # CPFs na hierarquia (DESLIGADO=NÃO)
    cpfs_hier = set(df_hier["cpf_limp"].unique())
    print(f"Total de CPFs na hierarquia (DESLIGADO=NÃO): {len(cpfs_hier):,}")

    # Filtra treinamentos do mês de referência
    mes_dt = pd.Period(mes_referencia, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ].copy()

    print(f"Total de registros de treinamentos em {mes_referencia}: {len(trein_mes):,}")
    print(f"CPFs únicos em treinamentos em {mes_referencia}: {trein_mes['cpf_limp'].nunique():,}")

    # CPFs em treinamentos que NÃO estão na hierarquia
    trein_mes["na_hierarquia"] = trein_mes["cpf_limp"].isin(cpfs_hier)
    fora_hier = trein_mes[~trein_mes["na_hierarquia"]].copy()

    print(f"\nCPFs em treinamentos que não estão na hierarquia: {fora_hier['cpf_limp'].nunique():,}")
    print(f"Registros desses CPFs: {len(fora_hier):,}")

    if fora_hier.empty:
        print("Nenhum CPF fora da hierarquia encontrado.")
        return

    # Revenda = Distribuidor, agrupando filiais
    fora_hier["Revenda"] = fora_hier["Distribuidor"]
    fora_hier["Revenda_agrup"] = fora_hier["Revenda"].apply(agrupar_revenda)

    # Resumo por revenda
    resumo = fora_hier.groupby("Revenda_agrup").agg(
        cpfs_unicos=("cpf_limp", "nunique"),
        registros=("cpf_limp", "size"),
    ).reset_index().sort_values("cpfs_unicos", ascending=False)

    print(f"\nRevendas com CPFs fora da hierarquia: {len(resumo)}")
    print(resumo.to_string(index=False))

    # Gera Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        ajustar_larguras(writer.sheets["Resumo"])

        for revenda in sorted(resumo["Revenda_agrup"].tolist()):
            df_rev = fora_hier[fora_hier["Revenda_agrup"] == revenda].copy()

            cols = ["cpf_limp", "Participante", "Revenda", "Revenda_agrup", "Curso", "Conclusão", "Trilha", "Estado", "Distribuidor", "PDV"]
            cols = [c for c in cols if c in df_rev.columns]
            df_rev = df_rev[cols]
            df_rev = df_rev.sort_values(["Participante", "Curso", "Conclusão"])

            aba_nome = re.sub(r'[\\/*?:\[\]]', '_', revenda)[:31]
            df_rev.to_excel(writer, sheet_name=aba_nome, index=False)
            ajustar_larguras(writer.sheets[aba_nome])

    print(f"\nArquivo gerado: {OUTPUT_FILE}")
    print(f"Total de abas: {len(resumo) + 1} (Resumo + revendas)")


if __name__ == "__main__":
    main()
