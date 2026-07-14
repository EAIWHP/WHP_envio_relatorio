import pandas as pd
from pathlib import Path
import re
import sys

sys.path.insert(0, "/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio")

from gerar_relatorio_semanal import (
    carregar_bases,
    calcular_aceites,
    calcular_treinamentos,
)

OUTPUT_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/validacao_calculos_email")
OUTPUT_DIR.mkdir(exist_ok=True)


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

    df_aceite = dados["aceites"]
    df_cad = dados["cadastro"]
    df_trein = dados["treinamentos"]
    df_hier = dados["hierarquia"]
    aba_aceite = dados["aba_aceite"]
    mes_referencia = dados["mes_referencia"]

    print(f"Mês de referência: {mes_referencia}")
    print(f"Aba de aceites: {aba_aceite}")

    # ---------------------------------------------------------------------
    # ACEITES
    # ---------------------------------------------------------------------
    print("\n" + "="*70)
    print("GERANDO VALIDAÇÃO DE ACEITES")
    print("="*70)

    aceite_reg, aceite_rev, mes_ref, base_aceite = calcular_aceites(
        df_aceite, df_cad, mes_referencia, aba_aceite,
        usou_ultima_aba=False, df_hier=df_hier
    )

    # Resumo por revenda
    resumo_aceites = aceite_rev[["revenda", "total_ativos", "aceitaram", "nao_aceitaram", "pct_aceite"]].copy()
    resumo_aceites = resumo_aceites.sort_values("pct_aceite", ascending=False)

    # Detalhes: base completa com flag aceitou
    base_aceite_det = base_aceite.copy()
    cpfs_aceitaram = set(df_aceite[df_aceite["mes_aceite"] == mes_ref]["cpf_limp"].unique())
    base_aceite_det["aceitou"] = base_aceite_det["cpf_limp"].isin(cpfs_aceitaram)

    # Adiciona nome do participante da hierarquia e do cadastro
    nome_hier = df_hier[["cpf_limp", "nome_hier"]].drop_duplicates(subset=["cpf_limp"], keep="first")
    nome_cad = df_cad[["cpf_limp", "nome"]].drop_duplicates(subset=["cpf_limp"], keep="first")

    base_aceite_det = base_aceite_det.merge(nome_hier, on="cpf_limp", how="left")
    base_aceite_det = base_aceite_det.merge(nome_cad, on="cpf_limp", how="left")
    base_aceite_det["Nome_Participante"] = base_aceite_det["nome_hier"].fillna(base_aceite_det["nome"])

    # Adiciona revenda original da hierarquia (antes de enriquecimento com cadastro)
    revenda_hier = df_hier[["cpf_limp", "revenda"]].drop_duplicates(subset=["cpf_limp"], keep="first")
    revenda_hier = revenda_hier.rename(columns={"revenda": "hierarquia_revenda"})
    base_aceite_det = base_aceite_det.merge(revenda_hier, on="cpf_limp", how="left")

    aceite_info = df_aceite[["cpf_limp", "Nome", "Revenda", "DataAceite"]].drop_duplicates(subset=["cpf_limp"], keep="first")

    base_aceite_det = base_aceite_det.merge(aceite_info, on="cpf_limp", how="left")

    # Preenche Nome (base de aceites) com Nome_Participante quando vazio
    base_aceite_det["Nome"] = base_aceite_det["Nome"].fillna(base_aceite_det["Nome_Participante"])

    base_aceitaram = base_aceite_det[base_aceite_det["aceitou"]].copy()
    base_nao_aceitaram = base_aceite_det[~base_aceite_det["aceitou"]].copy()

    # Reordena colunas
    colunas_desejadas = ["cpf_limp", "Nome", "revenda", "hierarquia_revenda", "regional_curta", "aceitou", "Revenda", "DataAceite"]
    colunas_existentes = [c for c in colunas_desejadas if c in base_aceite_det.columns]
    outras_colunas = [c for c in base_aceite_det.columns if c not in colunas_existentes]
    base_aceite_det = base_aceite_det[colunas_existentes + outras_colunas]
    base_aceitaram = base_aceitaram[colunas_existentes + outras_colunas]
    base_nao_aceitaram = base_nao_aceitaram[colunas_existentes + outras_colunas]

    # Base de aceites do mês
    aceite_mes = df_aceite[df_aceite["mes_aceite"] == mes_ref].copy()

    output_aceites = OUTPUT_DIR / "validacao_aceites_email.xlsx"
    with pd.ExcelWriter(output_aceites, engine="openpyxl") as writer:
        resumo_aceites.to_excel(writer, sheet_name="Resumo por Revenda", index=False)
        ajustar_larguras(writer.sheets["Resumo por Revenda"])

        base_aceite_det.to_excel(writer, sheet_name="Base Hierarquia", index=False)
        ajustar_larguras(writer.sheets["Base Hierarquia"])

        base_aceitaram.to_excel(writer, sheet_name="Aceitaram", index=False)
        ajustar_larguras(writer.sheets["Aceitaram"])

        base_nao_aceitaram.to_excel(writer, sheet_name="Nao Aceitaram", index=False)
        ajustar_larguras(writer.sheets["Nao Aceitaram"])

        aceite_mes.to_excel(writer, sheet_name="Base Aceites Completa", index=False)
        ajustar_larguras(writer.sheets["Base Aceites Completa"])

    print(f"Arquivo gerado: {output_aceites}")
    print(f"Total de revendas no resumo: {len(resumo_aceites)}")
    print("\nTop 5 revendas (% aceite):")
    print(resumo_aceites.head().to_string(index=False))

    # ---------------------------------------------------------------------
    # TREINAMENTOS
    # ---------------------------------------------------------------------
    print("\n" + "="*70)
    print("GERANDO VALIDAÇÃO DE TREINAMENTOS")
    print("="*70)

    trein_reg, trein_rev, trein_por_curso, info_cursos, base_trein = calcular_treinamentos(
        df_trein, df_cad, mes_referencia, df_hier=df_hier
    )

    curso1, curso2, nome1, nome2, sku1, sku2 = info_cursos

    # Resumo por revenda
    resumo_trein = trein_rev[["revenda", "total_ativos", "realizaram", "nao_realizaram", "pct_realizaram"]].copy()
    resumo_trein = resumo_trein.sort_values("pct_realizaram", ascending=False)

    # Adiciona colunas por curso individual
    c1_rev = trein_por_curso["curso1_rev"][["revenda", "realizaram"]].rename(columns={"realizaram": f"curso1_{sku1}"})
    c2_rev = trein_por_curso["curso2_rev"][["revenda", "realizaram"]].rename(columns={"realizaram": f"curso2_{sku2}"})
    resumo_trein = resumo_trein.merge(c1_rev, on="revenda", how="left")
    resumo_trein = resumo_trein.merge(c2_rev, on="revenda", how="left")

    # Detalhes
    mes_dt = pd.Period(mes_referencia, freq="M")
    trein_mes = df_trein[
        (df_trein["Conclusão"].dt.to_period("M") == mes_dt)
        & (df_trein["Estado"].str.lower() == "concluido")
    ].copy()

    cpf_curso1 = set(trein_mes[trein_mes["Curso"] == curso1]["cpf_limp"].unique())
    cpf_curso2 = set(trein_mes[trein_mes["Curso"] == curso2]["cpf_limp"].unique())
    cpf_ambos = cpf_curso1 & cpf_curso2

    base_trein_det = base_trein.copy()
    base_trein_det["curso1"] = base_trein_det["cpf_limp"].isin(cpf_curso1)
    base_trein_det["curso2"] = base_trein_det["cpf_limp"].isin(cpf_curso2)
    base_trein_det["ambos"] = base_trein_det["cpf_limp"].isin(cpf_ambos)

    # Adiciona nome do participante da hierarquia e do cadastro
    nome_hier = df_hier[["cpf_limp", "nome_hier"]].drop_duplicates(subset=["cpf_limp"], keep="first")
    nome_cad = df_cad[["cpf_limp", "nome"]].drop_duplicates(subset=["cpf_limp"], keep="first")

    base_trein_det = base_trein_det.merge(nome_hier, on="cpf_limp", how="left")
    base_trein_det = base_trein_det.merge(nome_cad, on="cpf_limp", how="left")
    base_trein_det["Nome"] = base_trein_det["nome_hier"].fillna(base_trein_det["nome"])

    # Lista de cursos obrigatórios feitos por CPF (para exibição correta)
    cursos_por_cpf = (
        trein_mes[trein_mes["Curso"].isin([curso1, curso2])]
        .groupby("cpf_limp")["Curso"]
        .apply(lambda x: " | ".join(sorted(set(x.astype(str)))))
        .reset_index(name="cursos_obrigatorios_feitos")
    )
    base_trein_det = base_trein_det.merge(cursos_por_cpf, on="cpf_limp", how="left")

    # Informação do participante (apenas nome, sem curso enganoso)
    participante_info = trein_mes[["cpf_limp", "Participante"]].drop_duplicates(subset=["cpf_limp"], keep="first")
    base_trein_det = base_trein_det.merge(participante_info, on="cpf_limp", how="left")

    # Preenche Participante com Nome quando vazio
    base_trein_det["Participante"] = base_trein_det["Participante"].fillna(base_trein_det["Nome"])

    base_realizaram = base_trein_det[base_trein_det["ambos"]].copy()
    base_nao_realizaram = base_trein_det[~base_trein_det["ambos"]].copy()

    # Reordena colunas para facilitar leitura
    colunas_desejadas = ["cpf_limp", "Participante", "Nome", "revenda", "regional_curta", "curso1", "curso2", "ambos", "cursos_obrigatorios_feitos"]
    colunas_existentes = [c for c in colunas_desejadas if c in base_trein_det.columns]
    outras_colunas = [c for c in base_trein_det.columns if c not in colunas_existentes]
    base_trein_det = base_trein_det[colunas_existentes + outras_colunas]
    base_realizaram = base_realizaram[colunas_existentes + outras_colunas]
    base_nao_realizaram = base_nao_realizaram[colunas_existentes + outras_colunas]

    # Detalhamento completo dos registros dos cursos obrigatórios
    detalhe_cursos_obrigatorios = trein_mes[trein_mes["Curso"].isin([curso1, curso2])].copy()

    df_cursos = pd.DataFrame({
        "curso": [curso1, curso2],
        "nome_curto": [nome1, nome2],
        "sku": [sku1, sku2],
        "cpfs_concluintes": [len(cpf_curso1), len(cpf_curso2)],
    })

    output_trein = OUTPUT_DIR / "validacao_treinamentos_email.xlsx"
    with pd.ExcelWriter(output_trein, engine="openpyxl") as writer:
        resumo_trein.to_excel(writer, sheet_name="Resumo por Revenda", index=False)
        ajustar_larguras(writer.sheets["Resumo por Revenda"])

        base_trein_det.to_excel(writer, sheet_name="Base Hierarquia", index=False)
        ajustar_larguras(writer.sheets["Base Hierarquia"])

        base_realizaram.to_excel(writer, sheet_name="Realizaram Ambos", index=False)
        ajustar_larguras(writer.sheets["Realizaram Ambos"])

        base_nao_realizaram.to_excel(writer, sheet_name="Nao Realizaram", index=False)
        ajustar_larguras(writer.sheets["Nao Realizaram"])

        df_cursos.to_excel(writer, sheet_name="Cursos Obrigatorios", index=False)
        ajustar_larguras(writer.sheets["Cursos Obrigatorios"])

        detalhe_cursos_obrigatorios.to_excel(writer, sheet_name="Detalhe Cursos Obrig", index=False)
        ajustar_larguras(writer.sheets["Detalhe Cursos Obrig"])

        trein_mes.to_excel(writer, sheet_name="Base Treinamentos Completa", index=False)
        ajustar_larguras(writer.sheets["Base Treinamentos Completa"])

    print(f"\nArquivo gerado: {output_trein}")
    print(f"Total de revendas no resumo: {len(resumo_trein)}")
    print(f"Cursos obrigatórios: {curso1} ({sku1}) e {curso2} ({sku2})")
    print(f"CPFs concluintes: curso1={len(cpf_curso1)}, curso2={len(cpf_curso2)}, ambos={len(cpf_ambos)}")
    print("\nTop 5 revendas (% treinamento):")
    print(resumo_trein.head().to_string(index=False))

    print("\n" + "="*70)
    print("VALIDAÇÃO CONCLUÍDA")
    print("="*70)


if __name__ == "__main__":
    main()
