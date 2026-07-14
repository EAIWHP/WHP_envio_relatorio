"""
Gera arquivo Excel com resumo de treinamentos por regional e revenda.
Mostra quem realizou e quem não realizou os 2 cursos obrigatórios do mês.
"""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from gerar_relatorio_semanal import carregar_bases, calcular_treinamentos


def main():
    bases = carregar_bases()
    df_cad = bases["cadastro"]
    df_trein = bases["treinamentos"]
    ano_mes = bases["mes_referencia"]

    trein_reg, trein_rev, trein_por_curso, cursos_info = calcular_treinamentos(df_trein, df_cad, ano_mes)

    if trein_reg is None or trein_rev is None:
        print("Não foi possível calcular os treinamentos.")
        return

    nome1, nome2, sku1, sku2 = cursos_info[2], cursos_info[3], cursos_info[4], cursos_info[5]

    # Ajusta nomes para exibição
    trein_reg = trein_reg.rename(columns={
        "regional_curta": "Regional",
        "total_ativos": "Base Ativa",
        "realizaram": "Realizaram Ambos",
        "nao_realizaram": "Não Realizaram",
        "pct_realizaram": "% Realizado",
    })

    trein_rev = trein_rev.rename(columns={
        "regional_curta": "Regional",
        "revenda": "Revenda",
        "total_ativos": "Base Ativa",
        "realizaram": "Realizaram Ambos",
        "nao_realizaram": "Não Realizaram",
        "pct_realizaram": "% Realizado",
    })

    # Ordena
    trein_reg = trein_reg.sort_values("% Realizado", ascending=False)
    trein_rev = trein_rev.sort_values(["Regional", "% Realizado"], ascending=[True, False])

    # Dados por curso individual
    c1_reg, c1_rev, c2_reg, c2_rev = None, None, None, None
    if trein_por_curso:
        c1_reg = trein_por_curso["curso1_reg"].rename(columns={
            "regional_curta": "Regional",
            "total_ativos": "Base Ativa",
            "realizaram": f"Realizaram {nome1}",
            "nao_realizaram": f"Não Realizaram {nome1}",
            "pct_realizaram": f"% {nome1}",
        })
        c2_reg = trein_por_curso["curso2_reg"].rename(columns={
            "regional_curta": "Regional",
            "total_ativos": "Base Ativa",
            "realizaram": f"Realizaram {nome2}",
            "nao_realizaram": f"Não Realizaram {nome2}",
            "pct_realizaram": f"% {nome2}",
        })
        c1_rev = trein_por_curso["curso1_rev"].rename(columns={
            "regional_curta": "Regional",
            "revenda": "Revenda",
            "total_ativos": "Base Ativa",
            "realizaram": f"Realizaram {nome1}",
            "nao_realizaram": f"Não Realizaram {nome1}",
            "pct_realizaram": f"% {nome1}",
        })
        c2_rev = trein_por_curso["curso2_rev"].rename(columns={
            "regional_curta": "Regional",
            "revenda": "Revenda",
            "total_ativos": "Base Ativa",
            "realizaram": f"Realizaram {nome2}",
            "nao_realizaram": f"Não Realizaram {nome2}",
            "pct_realizaram": f"% {nome2}",
        })

    output_path = Path(__file__).parent / "relatorios_gerados" / "contagem_treinamentos_por_regional_e_revenda.xlsx"
    output_path.parent.mkdir(exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        trein_reg.to_excel(writer, sheet_name="Ambos Cursos - Regional", index=False)
        trein_rev.to_excel(writer, sheet_name="Ambos Cursos - Revenda", index=False)
        if c1_reg is not None:
            c1_reg.to_excel(writer, sheet_name=f"Curso1 - Regional", index=False)
            c1_rev.to_excel(writer, sheet_name=f"Curso1 - Revenda", index=False)
        if c2_reg is not None:
            c2_reg.to_excel(writer, sheet_name=f"Curso2 - Regional", index=False)
            c2_rev.to_excel(writer, sheet_name=f"Curso2 - Revenda", index=False)

    print(f"Arquivo salvo em: {output_path}")


if __name__ == "__main__":
    main()
