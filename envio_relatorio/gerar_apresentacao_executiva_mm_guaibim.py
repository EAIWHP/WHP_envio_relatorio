#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Apresentação Executiva — MM Varejo/MM Atacado e Guaibim
------------------------------------------------------
Versão enxuta e visual para compartilhamento com cliente.
Foco em KPIs, comparativos e insights estratégicos.

Uso:
  python3 gerar_apresentacao_executiva_mm_guaibim.py
"""

import base64
import logging
import sys
from datetime import datetime
from io import BytesIO
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gerar_apresentacao_mm_guaibim import ler_relatorios, fig_to_base64, CORES_STATUS

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES
# ---------------------------------------------------------------------------
BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")
OUTPUT_DIR = BASE_DIR / "envio_relatorio" / "relatorios_gerados"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("apresentacao_executiva_mm_guaibim")


CORES_GRUPOS = {
    "MM Atacado": "#1a5276",
    "MM Varejo": "#2874a6",
    "Guaibim": "#1abc9c",
}


def grafico_comparativo_kpis(df_combined):
    """Gráfico comparativo de KPIs absolutos entre MM Atacado, MM Varejo e Guaibim."""
    metricas = ["Total Participantes", "Ativo", "Deu Aceite", "Fez Ambos Cursos", "Participantes Com Venda"]
    labels = ["Participantes", "Ativos", "Aceites", "Treinamentos", "Com Venda"]

    revendas = df_combined["Revenda"].tolist()
    cores = [CORES_GRUPOS.get(r, "#7f8c8d") for r in revendas]

    fig, ax = plt.subplots(figsize=(14, 7))
    x = range(len(labels))
    width = 0.25

    for i, revenda in enumerate(revendas):
        vals = [df_combined[df_combined["Revenda"] == revenda][m].values[0] for m in metricas]
        bars = ax.bar([p + width * (i - 1) for p in x], vals, width, label=revenda, color=cores[i])
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., h, f"{int(h):,}", ha="center", va="bottom", fontsize=8, fontweight="bold")

    ax.set_ylabel("Quantidade")
    ax.set_title("Comparativo de indicadores principais", fontsize=16, pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=11)
    ax.legend(fontsize=11)

    max_val = max(df_combined[metricas].max().max(), 1)
    ax.set_ylim(0, max_val * 1.15)
    return fig_to_base64(fig)


def grafico_taxas_comparativo(df_combined):
    """Comparativo de taxas percentuais entre as 3 revendas."""
    taxas_nomes = ["Ativos", "Aceite", "Treinamentos", "Com Venda"]
    metricas = ["Ativo", "Deu Aceite", "Fez Ambos Cursos", "Participantes Com Venda"]

    revendas = df_combined["Revenda"].tolist()
    cores = [CORES_GRUPOS.get(r, "#7f8c8d") for r in revendas]

    fig, ax = plt.subplots(figsize=(14, 7))
    x = range(len(taxas_nomes))
    width = 0.25

    for i, revenda in enumerate(revendas):
        row = df_combined[df_combined["Revenda"] == revenda].iloc[0]
        total = row["Total Participantes"]
        vals = [row[m] / total * 100 for m in metricas]
        bars = ax.bar([p + width * (i - 1) for p in x], vals, width, label=revenda, color=cores[i])
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., h + 1, f"{h:.1f}%", ha="center", va="bottom", fontsize=8, fontweight="bold")

    ax.set_ylabel("Percentual (%)")
    ax.set_title("Taxas de engajamento (%)" , fontsize=16, pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(taxas_nomes, fontsize=11)
    ax.legend(fontsize=11)
    ax.set_ylim(0, 100)

    return fig_to_base64(fig)


def grafico_rosca_status(df, titulo):
    """Gráfico de rosca com legenda para status."""
    status_cols = ["Ativo", "Inativo", "Pré-Cadastrado", "Indicado Moderado", "Em Moderação", "Reprovado", "Bloqueado"]
    valores = {c: df[c].sum() for c in status_cols if c in df.columns and df[c].sum() > 0}

    fig, ax = plt.subplots(figsize=(8, 8))
    cores = [CORES_STATUS.get(k, "#7f8c8d") for k in valores.keys()]
    wedges, texts, autotexts = ax.pie(
        valores.values(),
        labels=None,
        autopct=lambda pct: f"{pct:.1f}%" if pct >= 4 else "",
        colors=cores,
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.5, edgecolor="white"),
    )
    ax.set_title(titulo, pad=20, fontsize=14)
    plt.setp(autotexts, size=10, weight="bold")

    legend_labels = [f"{k}: {v} ({v/sum(valores.values())*100:.1f}%)" for k, v in valores.items()]
    ax.legend(wedges, legend_labels, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=10)

    return fig_to_base64(fig)


def gerar_insights_executivos(df_combined):
    """Gera insights enxutos e estratégicos para 3 revendas."""
    def get_row(revenda):
        return df_combined[df_combined["Revenda"] == revenda].iloc[0]

    mm_atacado = get_row("MM Atacado")
    mm_varejo = get_row("MM Varejo")
    guaibim = get_row("Guaibim")

    total_atacado = int(mm_atacado["Total Participantes"])
    total_varejo = int(mm_varejo["Total Participantes"])
    total_gui = int(guaibim["Total Participantes"])

    ativo_atacado = int(mm_atacado["Ativo"])
    ativo_varejo = int(mm_varejo["Ativo"])
    ativo_gui = int(guaibim["Ativo"])

    pct_ativo_atacado = ativo_atacado / total_atacado * 100
    pct_ativo_varejo = ativo_varejo / total_varejo * 100
    pct_ativo_gui = ativo_gui / total_gui * 100

    aceite_atacado = int(mm_atacado["Deu Aceite"])
    aceite_varejo = int(mm_varejo["Deu Aceite"])
    aceite_gui = int(guaibim["Deu Aceite"])

    pct_aceite_atacado = aceite_atacado / total_atacado * 100
    pct_aceite_varejo = aceite_varejo / total_varejo * 100
    pct_aceite_gui = aceite_gui / total_gui * 100

    treino_atacado = int(mm_atacado["Fez Ambos Cursos"])
    treino_varejo = int(mm_varejo["Fez Ambos Cursos"])
    treino_gui = int(guaibim["Fez Ambos Cursos"])

    pct_treino_atacado = treino_atacado / total_atacado * 100
    pct_treino_varejo = treino_varejo / total_varejo * 100
    pct_treino_gui = treino_gui / total_gui * 100

    vendas_atacado = int(mm_atacado["Total Produtos Vendidos"])
    vendas_varejo = int(mm_varejo["Total Produtos Vendidos"])
    vendas_gui = int(guaibim["Total Produtos Vendidos"])

    com_venda_atacado = int(mm_atacado["Participantes Com Venda"])
    com_venda_varejo = int(mm_varejo["Participantes Com Venda"])
    com_venda_gui = int(guaibim["Participantes Com Venda"])

    inativo_atacado = int(mm_atacado["Inativo"])
    inativo_varejo = int(mm_varejo["Inativo"])
    inativo_gui = int(guaibim["Inativo"])
    pct_inativo_atacado = inativo_atacado / total_atacado * 100
    pct_inativo_varejo = inativo_varejo / total_varejo * 100
    pct_inativo_gui = inativo_gui / total_gui * 100

    insights = [
        {
            "titulo": "Base e ativação",
            "texto": f"A <strong>MM Varejo</strong> concentra a maior parte da base ({total_varejo:,} participantes), seguida por <strong>MM Atacado</strong> ({total_atacado}) e <strong>Guaibim</strong> ({total_gui}). As taxas de ativos são {pct_ativo_varejo:.1f}% (MM Varejo), {pct_ativo_atacado:.1f}% (MM Atacado) e {pct_ativo_gui:.1f}% (Guaibim)."
        },
        {
            "titulo": "Aceite mensal",
            "texto": f"No mês de referência, <strong>MM Atacado</strong> lidera em aceite ({pct_aceite_atacado:.1f}%), seguida de <strong>MM Varejo</strong> ({pct_aceite_varejo:.1f}%) e <strong>Guaibim</strong> ({pct_aceite_gui:.1f}%)."
        },
        {
            "titulo": "Treinamentos obrigatórios",
            "texto": f"A adesão aos dois treinamentos de maio/2026 foi {pct_treino_atacado:.1f}% na MM Atacado, {pct_treino_varejo:.1f}% na MM Varejo e apenas {pct_treino_gui:.1f}% na Guaibim — ponto de atenção prioritário."
        },
        {
            "titulo": "Vendas",
            "texto": f"A <strong>MM Varejo</strong> vendeu {vendas_varejo:,} produtos ({com_venda_varejo} participantes), a <strong>MM Atacado</strong> vendeu {vendas_atacado:,} ({com_venda_atacado} participantes) e a <strong>Guaibim</strong> vendeu {vendas_gui:,} ({com_venda_gui} participantes)."
        },
        {
            "titulo": "Inativos — alerta Guaibim",
            "texto": f"A Guaibim apresenta {pct_inativo_gui:.1f}% de participantes inativos, bem acima da MM Atacado ({pct_inativo_atacado:.1f}%) e MM Varejo ({pct_inativo_varejo:.1f}%), indicando necessidade de ação de reengajamento."
        },
    ]

    recomendacoes = [
        "<strong>Guaibim — Treinamentos:</strong> priorizar comunicação direcionada para os cursos obrigatórios de maio/2026, dado o baixo índice de conclusão.",
        "<strong>Guaibim — Reengajamento:</strong> a alta taxa de inativos sugere uma campanha de reativação de cadastros.",
        "<strong>MM Varejo — Pré-cadastros:</strong> com o maior volume de pré-cadastrados, acompanhar a conversão para ativo dentro dos 90 dias evita perda futura.",
        "<strong>MM Atacado:</strong> manter o bom desempenho em aceite e treinamentos, replicando as práticas nas outras revendas.",
    ]

    return insights, recomendacoes


def montar_html_executivo(df_combined, graficos, insights, recomendacoes):
    """Monta o HTML executivo final."""
    def get_val(revenda, col):
        return int(df_combined[df_combined["Revenda"] == revenda][col].values[0])

    total_atacado = get_val("MM Atacado", "Total Participantes")
    total_varejo = get_val("MM Varejo", "Total Participantes")
    total_gui = get_val("Guaibim", "Total Participantes")

    ativo_atacado = get_val("MM Atacado", "Ativo")
    ativo_varejo = get_val("MM Varejo", "Ativo")
    ativo_gui = get_val("Guaibim", "Ativo")

    aceite_atacado = get_val("MM Atacado", "Deu Aceite")
    aceite_varejo = get_val("MM Varejo", "Deu Aceite")
    aceite_gui = get_val("Guaibim", "Deu Aceite")

    vendas_atacado = get_val("MM Atacado", "Total Produtos Vendidos")
    vendas_varejo = get_val("MM Varejo", "Total Produtos Vendidos")
    vendas_gui = get_val("Guaibim", "Total Produtos Vendidos")

    insights_html = "\n".join(
        f"""<div class="insight-card">
            <div class="insight-title">{i['titulo']}</div>
            <p>{i['texto']}</p>
        </div>""" for i in insights
    )
    recomendacoes_html = "\n".join(f"<li>{r}</li>" for r in recomendacoes)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Apresentação Executiva — MM Atacado, MM Varejo e Guaibim</title>
    <style>
        :root {{
            --mm-atacado: #1a5276;
            --mm-varejo: #2874a6;
            --guaibim: #1abc9c;
            --dark: #2c3e50;
            --danger: #e74c3c;
            --warning: #f39c12;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #eef1f5;
            color: var(--dark);
            line-height: 1.6;
        }}
        .slide {{
            max-width: 1200px;
            min-height: 700px;
            margin: 40px auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.10);
            padding: 60px;
            page-break-after: always;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .capa {{
            text-align: center;
            background: linear-gradient(135deg, var(--mm-atacado) 0%, var(--mm-varejo) 50%, var(--guaibim) 100%);
            color: white;
        }}
        .capa h1 {{ font-size: 3.2em; margin-bottom: 24px; font-weight: 700; }}
        .capa p {{ font-size: 1.4em; opacity: 0.92; margin-bottom: 12px; }}
        .capa .data {{ margin-top: 50px; font-size: 1em; opacity: 0.75; }}
        .logo-area {{ font-size: 0.95em; opacity: 0.8; margin-top: 20px; }}
        h2 {{
            color: var(--mm-atacado);
            font-size: 2.2em;
            margin-bottom: 32px;
            font-weight: 700;
        }}
        h3 {{ color: var(--mm-varejo); margin: 24px 0 16px; font-size: 1.3em; }}
        .subtitle {{
            color: #666;
            font-size: 1.1em;
            margin-bottom: 32px;
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin: 32px 0;
        }}
        .kpi {{
            background: white;
            border-radius: 16px;
            padding: 24px 16px;
            text-align: center;
            box-shadow: 0 4px 16px rgba(0,0,0,0.06);
            border-top: 5px solid #ccc;
        }}
        .kpi.atacado {{ border-top-color: var(--mm-atacado); }}
        .kpi.varejo {{ border-top-color: var(--mm-varejo); }}
        .kpi.guaibim {{ border-top-color: var(--guaibim); }}
        .kpi .valor {{ font-size: 2.2em; font-weight: 700; color: var(--dark); }}
        .kpi .rotulo {{ font-size: 0.9em; color: #666; margin-top: 10px; line-height: 1.3; }}
        .kpi .revenda-label {{
            font-size: 0.8em;
            font-weight: 700;
            text-transform: uppercase;
            margin-bottom: 8px;
            color: #888;
        }}
        .three-col {{
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 24px;
            align-items: center;
        }}
        .two-col {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 40px;
            align-items: center;
        }}
        .grafico {{
            text-align: center;
            margin: 20px 0;
        }}
        .grafico img {{
            max-width: 100%;
            height: auto;
            border-radius: 12px;
        }}
        .insights-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-top: 24px;
        }}
        .insight-card {{
            background: #f8fafc;
            border-left: 5px solid var(--guaibim);
            border-radius: 12px;
            padding: 24px;
        }}
        .insight-card.alert {{ border-left-color: var(--danger); background: #fdf2f2; }}
        .insight-card.warning {{ border-left-color: var(--warning); background: #fef9e7; }}
        .insight-title {{
            font-weight: 700;
            color: var(--mm-atacado);
            margin-bottom: 10px;
            font-size: 1.1em;
        }}
        .recomendacoes {{
            background: #fef9e7;
            border-radius: 16px;
            padding: 36px;
            margin-top: 24px;
            border-left: 6px solid var(--warning);
        }}
        .recomendacoes h3 {{ color: #b7791f; margin-top: 0; }}
        .recomendacoes ul {{ margin-left: 22px; }}
        .recomendacoes li {{ margin-bottom: 14px; font-size: 1.05em; }}
        .footer {{
            text-align: center;
            color: #888;
            font-size: 0.9em;
            margin-top: 40px;
        }}
        .badge {{
            display: inline-block;
            background: var(--guaibim);
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            margin-bottom: 12px;
        }}
        @media (max-width: 1000px) {{
            .kpi-grid {{ grid-template-columns: 1fr 1fr; }}
            .three-col {{ grid-template-columns: 1fr; }}
            .two-col {{ grid-template-columns: 1fr; }}
            .insights-grid {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>

<div class="slide capa">
    <h1>Programa +TOP</h1>
    <p>Visão Executiva — MM Atacado / MM Varejo / Guaibim</p>
    <p>Período: 01/05/2026 a 16/06/2026</p>
    <div class="data">{datetime.now().strftime('%d de %B de %Y')}</div>
    <div class="logo-area">Baseado nos relatórios de cadastro, banco, aceites e vendas de maio/2026.</div>
</div>

<div class="slide">
    <div class="badge">Resumo Executivo</div>
    <h2>Indicadores principais</h2>
    <p class="subtitle">Comparativo entre MM Atacado, MM Varejo e Guaibim no período analisado.</p>
    <div class="kpi-grid">
        <div class="kpi atacado">
            <div class="revenda-label">MM Atacado</div>
            <div class="valor">{total_atacado:,}</div>
            <div class="rotulo">Participantes</div>
        </div>
        <div class="kpi varejo">
            <div class="revenda-label">MM Varejo</div>
            <div class="valor">{total_varejo:,}</div>
            <div class="rotulo">Participantes</div>
        </div>
        <div class="kpi guaibim">
            <div class="revenda-label">Guaibim</div>
            <div class="valor">{total_gui:,}</div>
            <div class="rotulo">Participantes</div>
        </div>

        <div class="kpi atacado">
            <div class="revenda-label">MM Atacado</div>
            <div class="valor">{aceite_atacado}</div>
            <div class="rotulo">Aceites Maio/2026</div>
        </div>
        <div class="kpi varejo">
            <div class="revenda-label">MM Varejo</div>
            <div class="valor">{aceite_varejo}</div>
            <div class="rotulo">Aceites Maio/2026</div>
        </div>
        <div class="kpi guaibim">
            <div class="revenda-label">Guaibim</div>
            <div class="valor">{aceite_gui}</div>
            <div class="rotulo">Aceites Maio/2026</div>
        </div>

        <div class="kpi atacado">
            <div class="revenda-label">MM Atacado</div>
            <div class="valor">{vendas_atacado:,}</div>
            <div class="rotulo">Produtos Vendidos</div>
        </div>
        <div class="kpi varejo">
            <div class="revenda-label">MM Varejo</div>
            <div class="valor">{vendas_varejo:,}</div>
            <div class="rotulo">Produtos Vendidos</div>
        </div>
        <div class="kpi guaibim">
            <div class="revenda-label">Guaibim</div>
            <div class="valor">{vendas_gui:,}</div>
            <div class="rotulo">Produtos Vendidos</div>
        </div>

        <div class="kpi atacado">
            <div class="revenda-label">MM Atacado</div>
            <div class="valor">{ativo_atacado}</div>
            <div class="rotulo">Ativos</div>
        </div>
        <div class="kpi varejo">
            <div class="revenda-label">MM Varejo</div>
            <div class="valor">{ativo_varejo}</div>
            <div class="rotulo">Ativos</div>
        </div>
        <div class="kpi guaibim">
            <div class="revenda-label">Guaibim</div>
            <div class="valor">{ativo_gui}</div>
            <div class="rotulo">Ativos</div>
        </div>
    </div>

    <div class="grafico">
        <img src="data:image/png;base64,{graficos['comparativo_kpis']}" alt="Comparativo de KPIs">
    </div>
</div>

<div class="slide">
    <div class="badge">Engajamento</div>
    <h2>Taxas de engajamento</h2>
    <p class="subtitle">Percentual de participantes ativos, com aceite, treinamentos concluídos e com venda.</p>
    <div class="grafico">
        <img src="data:image/png;base64,{graficos['taxas_comparativo']}" alt="Taxas de engajamento">
    </div>
</div>

<div class="slide">
    <div class="badge">Status dos Cadastros</div>
    <h2>Distribuição de status</h2>
    <div class="three-col">
        <div class="grafico">
            <img src="data:image/png;base64,{graficos['atacado_status_pie']}" alt="Status MM Atacado">
        </div>
        <div class="grafico">
            <img src="data:image/png;base64,{graficos['varejo_status_pie']}" alt="Status MM Varejo">
        </div>
        <div class="grafico">
            <img src="data:image/png;base64,{graficos['gui_status_pie']}" alt="Status Guaibim">
        </div>
    </div>
</div>

<div class="slide">
    <div class="badge">Insights</div>
    <h2>Principais análises</h2>
    <div class="insights-grid">
        {insights_html}
    </div>
</div>

<div class="slide">
    <div class="badge">Próximos Passos</div>
    <h2>Recomendações estratégicas</h2>
    <div class="recomendacoes">
        <ul>
            {recomendacoes_html}
        </ul>
    </div>
    <div class="footer">
        Fonte: relatórios gerados em {datetime.now().strftime('%d/%m/%Y')}<br>
        Arquivos base: cadastro, PROD_WHP_Participante_Banco_v2_16062026, aceites e vendas de maio/2026.
    </div>
</div>

</body>
</html>"""

    return html


def main():
    logger.info("=" * 60)
    logger.info("Gerando apresentação executiva — 3 revendas")

    dados = ler_relatorios()

    # Combina os dados de revenda em um único DataFrame com 3 linhas
    df_combined = dados["mm"]["revenda"].copy()
    df_combined = pd.concat([df_combined, dados["guaibim"]["revenda"]], ignore_index=True)

    logger.info("Criando gráficos executivos...")
    graficos = {}
    graficos["comparativo_kpis"] = grafico_comparativo_kpis(df_combined)
    graficos["taxas_comparativo"] = grafico_taxas_comparativo(df_combined)
    graficos["atacado_status_pie"] = grafico_rosca_status(
        df_combined[df_combined["Revenda"] == "MM Atacado"], "MM Atacado"
    )
    graficos["varejo_status_pie"] = grafico_rosca_status(
        df_combined[df_combined["Revenda"] == "MM Varejo"], "MM Varejo"
    )
    graficos["gui_status_pie"] = grafico_rosca_status(
        df_combined[df_combined["Revenda"] == "Guaibim"], "Guaibim"
    )

    logger.info("Gerando insights executivos...")
    insights, recomendacoes = gerar_insights_executivos(df_combined)

    logger.info("Montando HTML executivo...")
    html = montar_html_executivo(df_combined, graficos, insights, recomendacoes)

    hoje_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"apresentacao_executiva_mm_guaibim_{hoje_str}.html"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    logger.info(f"Apresentação executiva salva em: {output_path}")


if __name__ == "__main__":
    main()
