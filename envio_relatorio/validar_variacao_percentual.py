"""
validar_variacao_percentual.py

Gera um arquivo Excel de validação do cálculo da coluna Variação % do relatório semanal +TOP.
Mostra, por regional e por revenda, os valores do mês atual, do mês anterior e a variação percentual calculada.

Uso:
    cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
    python3 validar_variacao_percentual.py
"""

import sys
from pathlib import Path

import pandas as pd
import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gerar_relatorio_semanal import (
    carregar_bases,
    calcular_cadastros,
    calcular_treinamentos,
    calcular_aceites,
    carregar_snapshot_mes_anterior,
    complementar_snapshot_por_revenda,
    mapeamento_revenda_regional,
    _normalizar_chave_regional,
    _normalizar_chave_revenda,
)

OUTPUT_DIR = Path(__file__).resolve().parent / "relatorios_gerados"


def _buscar_no_snapshot(snap, tipo, chave, col):
    """Busca valor no snapshot já normalizado."""
    if not snap:
        return 0
    snap_tipo = snap.get(tipo, {})
    for k, v in snap_tipo.items():
        if isinstance(k, str) and "|" in k:
            reg, rev = k.split("|", 1)
            norm = f"{_normalizar_chave_regional(reg)}|{_normalizar_chave_revenda(rev)}"
        else:
            norm = _normalizar_chave_regional(k)
        if norm == chave:
            return float(v.get(col, 0))
    return 0


def _var_pct(atual, anterior):
    if anterior == 0:
        return 0.0 if atual == 0 else np.nan
    return round((atual - anterior) / anterior * 100, 1)


def _buscar_regional_no_snapshot(snap, tipo, reg):
    """Busca valor no snapshot regional já normalizado."""
    if not snap:
        return {}
    snap_tipo = snap.get(tipo, {})
    reg_norm = _normalizar_chave_regional(reg)
    for k, v in snap_tipo.items():
        if _normalizar_chave_regional(k) == reg_norm:
            return v
    return {}


def _build_regional(cad_reg, trein_reg, aceite_reg, snapshot_anterior):
    rows = []
    for _, r in cad_reg.iterrows():
        reg = r["regional_curta"]

        cad_ant = _buscar_regional_no_snapshot(snapshot_anterior, "cadastros", reg)
        trein_ant = _buscar_regional_no_snapshot(snapshot_anterior, "treinamentos", reg)
        aceite_ant = _buscar_regional_no_snapshot(snapshot_anterior, "aceites", reg)

        ativos_ant = float(cad_ant.get("ativos", 0))
        realizaram_ant = float(trein_ant.get("realizaram", 0))
        aceitaram_ant = float(aceite_ant.get("aceitaram", 0))

        trein_row = trein_reg[trein_reg["regional_curta"] == reg]
        aceite_row = aceite_reg[aceite_reg["regional"] == reg]

        rows.append({
            "Regional": reg,
            "Cadastro_Atual": int(r["ativos"]),
            "Cadastro_Anterior": int(ativos_ant),
            "Cadastro_Variacao_%": _var_pct(r["ativos"], ativos_ant),
            "Treinamento_Atual": int(trein_row["realizaram"].sum()) if not trein_row.empty else 0,
            "Treinamento_Anterior": int(realizaram_ant),
            "Treinamento_Variacao_%": _var_pct(trein_row["realizaram"].sum() if not trein_row.empty else 0, realizaram_ant),
            "Aceite_Atual": int(aceite_row["aceitaram"].sum()) if not aceite_row.empty else 0,
            "Aceite_Anterior": int(aceitaram_ant),
            "Aceite_Variacao_%": _var_pct(aceite_row["aceitaram"].sum() if not aceite_row.empty else 0, aceitaram_ant),
        })
    return pd.DataFrame(rows)


def _build_revenda(cad_rev, trein_rev, aceite_rev, snapshot_anterior):
    rows = []
    for _, r in cad_rev.iterrows():
        reg = r["regional_curta"]
        rev = r["revenda"]
        chave = f"{_normalizar_chave_regional(reg)}|{_normalizar_chave_revenda(rev)}"

        ativos_ant = _buscar_no_snapshot(snapshot_anterior, "cadastros_rev", chave, "ativos")
        realizaram_ant = _buscar_no_snapshot(snapshot_anterior, "treinamentos_rev", chave, "realizaram")
        aceitaram_ant = _buscar_no_snapshot(snapshot_anterior, "aceites_rev", chave, "aceitaram")

        trein_row = trein_rev[
            (trein_rev["regional_curta"].apply(_normalizar_chave_regional) == _normalizar_chave_regional(reg)) &
            (trein_rev["revenda"].apply(_normalizar_chave_revenda) == _normalizar_chave_revenda(rev))
        ]
        aceite_row = aceite_rev[
            (aceite_rev["regional_curta"].apply(_normalizar_chave_regional) == _normalizar_chave_regional(reg)) &
            (aceite_rev["revenda"].apply(_normalizar_chave_revenda) == _normalizar_chave_revenda(rev))
        ]

        rows.append({
            "Regional": reg,
            "Revenda": rev,
            "Cadastro_Atual": int(r["ativos"]),
            "Cadastro_Anterior": int(ativos_ant),
            "Cadastro_Variacao_%": _var_pct(r["ativos"], ativos_ant),
            "Treinamento_Atual": int(trein_row["realizaram"].sum()) if not trein_row.empty else 0,
            "Treinamento_Anterior": int(realizaram_ant),
            "Treinamento_Variacao_%": _var_pct(trein_row["realizaram"].sum() if not trein_row.empty else 0, realizaram_ant),
            "Aceite_Atual": int(aceite_row["aceitaram"].sum()) if not aceite_row.empty else 0,
            "Aceite_Anterior": int(aceitaram_ant),
            "Aceite_Variacao_%": _var_pct(aceite_row["aceitaram"].sum() if not aceite_row.empty else 0, aceitaram_ant),
        })
    return pd.DataFrame(rows)


def main():
    print("Carregando bases...")
    bases = carregar_bases()
    df_cad = bases["cadastro"]
    df_trein = bases["treinamentos"]
    df_aceite = bases["aceites"]
    ano_mes = bases["mes_referencia"]

    print("Calculando indicadores atuais...")
    cad_reg, cad_rev = calcular_cadastros(df_cad, bases.get("hierarquia"), bases.get("ferias_hier"), bases.get("status_completo"))
    trein_reg, trein_rev, _, _, _ = calcular_treinamentos(df_trein, df_cad, ano_mes, bases.get("hierarquia"))
    aceite_reg, aceite_rev, _, _ = calcular_aceites(
        df_aceite, df_cad, ano_mes, bases["aba_aceite"], bases["usou_ultima_aba"], df_hier=bases.get("hierarquia")
    )

    print("Carregando snapshot do mês anterior...")
    snapshot_anterior = carregar_snapshot_mes_anterior(ano_mes)
    if snapshot_anterior:
        mapa_regional = mapeamento_revenda_regional(df_cad)
        snapshot_anterior = complementar_snapshot_por_revenda(snapshot_anterior, ano_mes, mapa_regional)
        print(f"  Cadastros regional: {len(snapshot_anterior.get('cadastros', {}))}")
        print(f"  Cadastros revenda: {len(snapshot_anterior.get('cadastros_rev', {}))}")
        print(f"  Treinamentos revenda: {len(snapshot_anterior.get('treinamentos_rev', {}))}")
        print(f"  Aceites revenda: {len(snapshot_anterior.get('aceites_rev', {}))}")

    print("Montando abas de validação...")
    snap = snapshot_anterior or {}
    df_reg = _build_regional(cad_reg, trein_reg, aceite_reg, snap)
    df_rev = _build_revenda(cad_rev, trein_rev, aceite_rev, snap)

    # Aba de resumo explicativo
    df_regras = pd.DataFrame({
        "Item": [
            "Fórmula da variação",
            "Base cadastros",
            "Base treinamentos",
            "Base aceites",
            "Snapshot anterior",
            "Complementação por revenda",
            "Normalização regional",
            "Normalização revenda",
            "Valores ausentes/anterior zerado",
            "Setas/formato",
        ],
        "Descrição": [
            "((quantidade_atual - quantidade_anterior) / quantidade_anterior) * 100, arredondado em 1 casa decimal",
            "Quantidade de ativos (status Ativo na hierarquia/cadastro)",
            "Quantidade de CPFs que realizaram ambos os cursos obrigatórios do mês",
            "Quantidade de CPFs que deram aceite no mês",
            "Snapshot mais recente do mês anterior filtrado pela data real (snap['data'])",
            "Se o snapshot não tiver dados por revenda, usa o relatório Excel do mês anterior e a base histórica de aceite mensal",
            "Remove hífens/espaços extras: 'Centro Norte' == 'Centro-Norte'",
            "Aplica nome_revenda_exibicao + uppercase: 'Gazin Varejo' == 'GAZIN VAREJO'",
            "Quando anterior = 0 e atual > 0: NaN (formatado como 0.0%); quando ambos zero: 0.0%",
            "Setas simples à direita: positivo ↑, negativo ↓, zero/sem dados →",
        ],
    })

    output_path = OUTPUT_DIR / f"validacao_variacao_percentual_{ano_mes}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_regras.to_excel(writer, sheet_name="Regras", index=False)
        df_reg.to_excel(writer, sheet_name="Por_Regional", index=False)
        df_rev.to_excel(writer, sheet_name="Por_Revenda", index=False)

    print(f"\nArquivo de validação salvo em: {output_path}")


if __name__ == "__main__":
    main()
