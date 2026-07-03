"""
Gera arquivo Excel com contagem de cadastros ativos e base total
por regional e por revenda.
"""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from gerar_relatorio_semanal import carregar_bases, calcular_cadastros


def main():
    bases = carregar_bases()
    df_cad = bases["cadastro"]
    cad_reg, cad_rev = calcular_cadastros(df_cad)

    # Ajusta nomes para exibição
    cad_reg = cad_reg.rename(columns={
        "regional_curta": "Regional",
        "total": "Total de Participantes",
        "ativos": "Usuários com +TOP",
        "nao_ativos": "Usuários sem +TOP",
        "pct_ativos": "% Ativos",
    })

    cad_rev = cad_rev.rename(columns={
        "regional_curta": "Regional",
        "revenda": "Revenda",
        "total": "Total de Participantes",
        "ativos": "Usuários com +TOP",
        "nao_ativos": "Usuários sem +TOP",
        "pct_ativos": "% Ativos",
    })

    # Ordena
    cad_reg = cad_reg.sort_values("% Ativos", ascending=False)
    cad_rev = cad_rev.sort_values(["Regional", "% Ativos"], ascending=[True, False])

    output_path = Path(__file__).parent / "relatorios_gerados" / "contagem_cadastros_ativos_por_regional_e_revenda.xlsx"
    output_path.parent.mkdir(exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        cad_reg.to_excel(writer, sheet_name="Por Regional", index=False)
        cad_rev.to_excel(writer, sheet_name="Por Revenda", index=False)

    print(f"Arquivo salvo em: {output_path}")


if __name__ == "__main__":
    main()
