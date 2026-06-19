#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Apresentação de gráficos e insights — MM Varejo/MM Atacado e Guaibim
--------------------------------------------------------------------
Gera um arquivo HTML com análise visual dos relatórios gerados por
`gerar_relatorios_mm_guaibim_atualizado.py`.

Uso:
  python3 gerar_apresentacao_mm_guaibim.py
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

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES
# ---------------------------------------------------------------------------
BASE_DIR = Path("/home/thamiresvieira/projetos/Programa_mais_top")
INPUT_DIR = BASE_DIR / "envio_relatorio" / "relatorios_gerados"
OUTPUT_DIR = INPUT_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("apresentacao_mm_guaibim")

CORES_STATUS = {
    "Ativo": "#2ecc71",
    "Inativo": "#e74c3c",
    "Pré-Cadastrado": "#f39c12",
    "Indicado Moderado": "#9b59b6",
    "Em Moderação": "#3498db",
    "Reprovado": "#95a5a6",
    "Bloqueado": "#34495e",
}


def fig_to_base64(fig):
    """Converte uma figura matplotlib para string base64 (PNG)."""
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def ler_relatorios():
    """Localiza os relatórios mais recentes e carrega as abas principais."""
    arquivos = sorted(INPUT_DIR.glob("relatorio_mm_varejo_atacado_*.xlsx"))
    if not arquivos:
        raise FileNotFoundError("Relatório MM Varejo + MM Atacado não encontrado")
    path_mm = arquivos[-1]

    arquivos = sorted(INPUT_DIR.glob("relatorio_guaibim_*.xlsx"))
    if not arquivos:
        raise FileNotFoundError("Relatório Guaibim não encontrado")
    path_guaibim = arquivos[-1]

    logger.info(f"Lendo {path_mm.name}")
    logger.info(f"Lendo {path_guaibim.name}")

    return {
        "mm": {
            "revenda": pd.read_excel(path_mm, sheet_name="Resumo_por_Revenda"),
            "loja": pd.read_excel(path_mm, sheet_name="Resumo_por_Loja"),
            "detalhe": pd.read_excel(path_mm, sheet_name="Detalhamento"),
        },
        "guaibim": {
            "revenda": pd.read_excel(path_guaibim, sheet_name="Resumo_por_Revenda"),
            "loja": pd.read_excel(path_guaibim, sheet_name="Resumo_por_Loja"),
            "detalhe": pd.read_excel(path_guaibim, sheet_name="Detalhamento"),
        },
    }


def grafico_barras_status(df):
    """Gráfico de barras empilhadas por status."""
    status_cols = ["Ativo", "Inativo", "Pré-Cadastrado", "Indicado Moderado", "Em Moderação", "Reprovado", "Bloqueado"]
    status_cols = [c for c in status_cols if c in df.columns]

    fig, ax = plt.subplots(figsize=(10, 6))
    bottom = pd.Series([0] * len(df), index=df.index)

    for col in status_cols:
        ax.bar(df["Revenda"], df[col], bottom=bottom, label=col, color=CORES_STATUS.get(col, "#7f8c8d"))
        bottom += df[col]

    ax.set_ylabel("Quantidade de participantes")
    ax.set_title("Distribuição de status por revenda")
    ax.legend(title="Status", bbox_to_anchor=(1.05, 1), loc="upper left")
    ax.set_ylim(0, bottom.max() * 1.15)
    for i, rev in enumerate(df["Revenda"]):
        ax.text(i, bottom.iloc[i] + bottom.max() * 0.02, str(bottom.iloc[i]), ha="center", fontweight="bold")

    return fig_to_base64(fig)


def grafico_pizza_status(df, titulo):
    """Gráfico de rosca (donut) da distribuição de status com legenda."""
    status_cols = ["Ativo", "Inativo", "Pré-Cadastrado", "Indicado Moderado", "Em Moderação", "Reprovado", "Bloqueado"]
    valores = {c: df[c].sum() for c in status_cols if c in df.columns and df[c].sum() > 0}

    fig, ax = plt.subplots(figsize=(9, 9))
    cores = [CORES_STATUS.get(k, "#7f8c8d") for k in valores.keys()]
    wedges, texts, autotexts = ax.pie(
        valores.values(),
        labels=None,
        autopct=lambda pct: f"{pct:.1f}%" if pct >= 3 else "",
        colors=cores,
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.5, edgecolor="white"),
    )
    ax.set_title(titulo, pad=20, fontsize=14)
    plt.setp(autotexts, size=10, weight="bold")

    # Legenda com quantidade
    legend_labels = [f"{k}: {v} ({v/sum(valores.values())*100:.1f}%)" for k, v in valores.items()]
    ax.legend(wedges, legend_labels, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

    return fig_to_base64(fig)


def grafico_metricas(df, titulo):
    """Gráfico comparativo de aceite e treinamentos."""
    metricas = ["Deu Aceite", "Fez Ambos Cursos", "Participantes Com Venda"]
    negativas = ["Não Deu Aceite", "Não Fez Ambos Cursos"]

    fig, ax = plt.subplots(figsize=(10, 6))
    x = range(len(df))
    width = 0.25

    for i, m in enumerate(metricas):
        ax.bar([p + width * i for p in x], df[m], width=width, label=m)

    ax.set_xticks([p + width for p in x])
    ax.set_xticklabels(df["Revenda"])
    ax.set_ylabel("Quantidade")
    ax.set_title(titulo)
    ax.legend()
    ax.set_ylim(0, max(df[metricas].max()) * 1.2)

    for i, m in enumerate(metricas):
        for j, v in enumerate(df[m]):
            ax.text(j + width * i, v + max(df[metricas].max()) * 0.02, str(int(v)), ha="center", fontsize=8)

    return fig_to_base64(fig)


def grafico_top_lojas(df_loja, revenda_filtro, titulo, n=10, excluir_matriz=False):
    """Top N lojas por total de produtos vendidos."""
    sub = df_loja[df_loja["Revenda"] == revenda_filtro].copy()
    if sub.empty:
        return None
    if excluir_matriz:
        sub = sub[sub["Loja"].astype(str).str.lower() != "matriz"]
        titulo = titulo + " (exceto Matriz)"
    sub = sub[sub["Total Produtos Vendidos"] > 0]
    sub = sub.sort_values("Total Produtos Vendidos", ascending=False).head(n)
    if sub.empty:
        return None

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(sub["Loja"], sub["Total Produtos Vendidos"], color="#1abc9c")
    ax.invert_yaxis()
    ax.set_xlabel("Total de produtos vendidos")
    ax.set_title(titulo)
    for i, v in enumerate(sub["Total Produtos Vendidos"]):
        ax.text(v + sub["Total Produtos Vendidos"].max() * 0.01, i, str(int(v)), va="center", fontsize=9)

    return fig_to_base64(fig)


def grafico_vendas_status(df_detalhe, titulo):
    """Participantes com venda vs sem venda por status."""
    df = df_detalhe.copy()
    df["Com Venda"] = df["Total Produtos Vendidos"] > 0
    resumo = df.groupby("Status")["Com Venda"].agg(["sum", "size"]).reset_index()
    resumo["Sem Venda"] = resumo["size"] - resumo["sum"]
    resumo = resumo.rename(columns={"sum": "Com Venda"})

    fig, ax = plt.subplots(figsize=(10, 6))
    x = range(len(resumo))
    width = 0.35
    ax.bar([p - width/2 for p in x], resumo["Com Venda"], width=width, label="Com Venda", color="#2ecc71")
    ax.bar([p + width/2 for p in x], resumo["Sem Venda"], width=width, label="Sem Venda", color="#e74c3c")

    ax.set_xticks(x)
    ax.set_xticklabels(resumo["Status"], rotation=45, ha="right")
    ax.set_ylabel("Quantidade")
    ax.set_title(titulo)
    ax.legend()

    return fig_to_base64(fig)


def gerar_insights(dados):
    """Gera insights textuais a partir dos dados."""
    mm_rev = dados["mm"]["revenda"]
    gui_rev = dados["guaibim"]["revenda"]
    mm_det = dados["mm"]["detalhe"]
    gui_det = dados["guaibim"]["detalhe"]

    insights = []

    # Comparativo geral
    total_mm = mm_rev["Total Participantes"].sum()
    total_gui = gui_rev["Total Participantes"].sum()
    ativo_mm = mm_rev["Ativo"].sum()
    ativo_gui = gui_rev["Ativo"].sum()
    pct_ativo_mm = ativo_mm / total_mm * 100
    pct_ativo_gui = ativo_gui / total_gui * 100

    insights.append(
        f"A base da <strong>MM Varejo + MM Atacado</strong> reúne {total_mm:,} participantes, "
        f"enquanto a <strong>Guaibim</strong> tem {total_gui:,}. "
        f"A taxa de ativos é de <strong>{pct_ativo_mm:.1f}%</strong> na MM e "
        f"<strong>{pct_ativo_gui:.1f}%</strong> na Guaibim."
    )

    # Status pré-cadastrado
    pre_mm = mm_rev["Pré-Cadastrado"].sum()
    pre_gui = gui_rev["Pré-Cadastrado"].sum()
    insights.append(
        f"Há <strong>{pre_mm}</strong> pré-cadastrados na MM e <strong>{pre_gui}</strong> na Guaibim. "
        f"Esses participantes foram mantidos por estarem presentes na tabela do banco."
    )

    # Aceites
    aceite_mm = mm_rev["Deu Aceite"].sum()
    aceite_gui = gui_rev["Deu Aceite"].sum()
    pct_aceite_mm = aceite_mm / total_mm * 100
    pct_aceite_gui = aceite_gui / total_gui * 100
    insights.append(
        f"No mês de referência (maio/2026), <strong>{aceite_mm}</strong> participantes da MM deram aceite "
        f"({pct_aceite_mm:.1f}%), contra <strong>{aceite_gui}</strong> da Guaibim ({pct_aceite_gui:.1f}%)."
    )

    # Treinamentos
    treino_mm = mm_rev["Fez Ambos Cursos"].sum()
    treino_gui = gui_rev["Fez Ambos Cursos"].sum()
    pct_treino_mm = treino_mm / total_mm * 100
    pct_treino_gui = treino_gui / total_gui * 100
    insights.append(
        f"Quanto aos dois treinamentos obrigatórios de maio/2026, "
        f"<strong>{treino_mm}</strong> da MM concluíram ambos ({pct_treino_mm:.1f}%), "
        f"enquanto na Guaibim foram apenas <strong>{treino_gui}</strong> ({pct_treino_gui:.1f}%)."
    )

    # Vendas
    vendas_mm = mm_rev["Total Produtos Vendidos"].sum()
    vendas_gui = gui_rev["Total Produtos Vendidos"].sum()
    com_venda_mm = mm_rev["Participantes Com Venda"].sum()
    com_venda_gui = gui_rev["Participantes Com Venda"].sum()
    insights.append(
        f"Em vendas, a MM totalizou <strong>{vendas_mm:,}</strong> produtos vendidos "
        f"({com_venda_mm} participantes com ao menos uma venda). "
        f"A Guaibim vendeu <strong>{vendas_gui:,}</strong> produtos ({com_venda_gui} participantes com venda)."
    )

    # Lojas
    mm_loja = dados["mm"]["loja"]
    if not mm_loja.empty:
        top_loja = mm_loja.sort_values("Total Produtos Vendidos", ascending=False).iloc[0]
        insights.append(
            f"A loja da MM com maior volume de vendas foi <strong>{top_loja['Loja']}</strong>, "
            f"com <strong>{int(top_loja['Total Produtos Vendidos']):,}</strong> produtos vendidos."
        )

    # Recomendações
    recomendacoes = []
    if pct_treino_gui < 10:
        recomendacoes.append(
            "A Guaibim apresenta baixa adesão aos treinamentos obrigatórios. "
            "Recomenda-se ação de comunicação/lembrte para os cursos de maio/2026."
        )
    if pct_aceite_gui < pct_aceite_mm:
        recomendacoes.append(
            "A taxa de aceite da Guaibim está abaixo da MM. Vale rever o envio de comunicados e prazo de aceite."
        )
    if pre_mm > 100:
        recomendacoes.append(
            "A MM possui um volume elevado de pré-cadastrados. Acompanhar a conversão para ativo "
            "dentro dos 90 dias é fundamental para evitar inativação automática."
        )

    return insights, recomendacoes


def montar_html(dados, graficos, insights, recomendacoes):
    """Monta o arquivo HTML final."""
    mm_rev = dados["mm"]["revenda"]
    gui_rev = dados["guaibim"]["revenda"]

    total_mm = int(mm_rev["Total Participantes"].sum())
    total_gui = int(gui_rev["Total Participantes"].sum())
    vendas_mm = int(mm_rev["Total Produtos Vendidos"].sum())
    vendas_gui = int(gui_rev["Total Produtos Vendidos"].sum())

    insights_html = "\n".join(f"<li>{i}</li>" for i in insights)
    recomendacoes_html = "\n".join(f"<li>{r}</li>" for r in recomendacoes)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Apresentação MM Varejo/MM Atacado e Guaibim</title>
    <style>
        :root {{
            --primary: #1a5276;
            --secondary: #2874a6;
            --accent: #1abc9c;
            --light: #f8f9fa;
            --dark: #2c3e50;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f0f2f5;
            color: var(--dark);
            line-height: 1.6;
        }}
        .slide {{
            max-width: 1200px;
            margin: 40px auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            padding: 48px;
            page-break-after: always;
        }}
        .capa {{
            text-align: center;
            padding: 80px 48px;
            background: linear-gradient(135deg, var(--primary), var(--secondary));
            color: white;
        }}
        .capa h1 {{ font-size: 2.8em; margin-bottom: 20px; }}
        .capa p {{ font-size: 1.3em; opacity: 0.9; }}
        .capa .data {{ margin-top: 40px; font-size: 1em; opacity: 0.8; }}
        h2 {{
            color: var(--primary);
            font-size: 2em;
            margin-bottom: 24px;
            border-bottom: 3px solid var(--accent);
            padding-bottom: 12px;
        }}
        h3 {{ color: var(--secondary); margin: 24px 0 12px; }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin: 32px 0;
        }}
        .kpi {{
            background: var(--light);
            border-radius: 12px;
            padding: 24px;
            text-align: center;
            border-left: 5px solid var(--accent);
        }}
        .kpi .valor {{ font-size: 2.2em; font-weight: bold; color: var(--primary); }}
        .kpi .rotulo {{ font-size: 0.95em; color: #666; margin-top: 8px; }}
        .grafico {{
            text-align: center;
            margin: 32px 0;
        }}
        .grafico img {{
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        }}
        .insights {{
            background: #eaf2f8;
            border-radius: 12px;
            padding: 28px;
            margin: 24px 0;
        }}
        .insights ul, .recomendacoes ul {{
            margin-left: 20px;
            margin-top: 12px;
        }}
        .insights li, .recomendacoes li {{
            margin-bottom: 10px;
        }}
        .recomendacoes {{
            background: #fef9e7;
            border-radius: 12px;
            padding: 28px;
            margin: 24px 0;
            border-left: 5px solid #f39c12;
        }}
        .two-col {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 24px;
        }}
        @media (max-width: 900px) {{
            .two-col {{ grid-template-columns: 1fr; }}
        }}
        .footer {{
            text-align: center;
            color: #888;
            font-size: 0.85em;
            margin-top: 40px;
        }}
    </style>
</head>
<body>

<div class="slide capa">
    <h1>Programa +TOP</h1>
    <p>Análise comparativa — MM Varejo / MM Atacado e Guaibim</p>
    <p>Período: 01/05/2026 a 16/06/2026</p>
    <div class="data">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
</div>

<div class="slide">
    <h2>Visão Geral</h2>
    <div class="kpi-grid">
        <div class="kpi">
            <div class="valor">{total_mm:,}</div>
            <div class="rotulo">Participantes MM Varejo + MM Atacado</div>
        </div>
        <div class="kpi">
            <div class="valor">{total_gui:,}</div>
            <div class="rotulo">Participantes Guaibim</div>
        </div>
        <div class="kpi">
            <div class="valor">{vendas_mm:,}</div>
            <div class="rotulo">Produtos vendidos — MM</div>
        </div>
        <div class="kpi">
            <div class="valor">{vendas_gui:,}</div>
            <div class="rotulo">Produtos vendidos — Guaibim</div>
        </div>
    </div>
    <div class="insights">
        <h3>Principais insights</h3>
        <ul>
            {insights_html}
        </ul>
    </div>
</div>

<div class="slide">
    <h2>MM Varejo + MM Atacado</h2>
    <h3>Distribuição de status</h3>
    <div class="grafico"><img src="data:image/png;base64,{graficos['mm_status_bar']}" alt="Status MM"></div>
    <div class="two-col">
        <div class="grafico"><img src="data:image/png;base64,{graficos['mm_status_pie']}" alt="Status MM Pizza"></div>
        <div class="grafico"><img src="data:image/png;base64,{graficos['mm_metricas']}" alt="Métricas MM"></div>
    </div>
</div>

<div class="slide">
    <h2>MM Varejo + MM Atacado — Vendas e Lojas</h2>
    <div class="two-col">
        <div>
            <h3>Vendas por status</h3>
            <div class="grafico"><img src="data:image/png;base64,{graficos['mm_vendas_status']}" alt="Vendas por status MM"></div>
        </div>
        <div>
            <h3>Top 10 lojas em vendas</h3>
            <div class="grafico"><img src="data:image/png;base64,{graficos['mm_top_lojas']}" alt="Top lojas MM"></div>
        </div>
    </div>
</div>

<div class="slide">
    <h2>Guaibim</h2>
    <h3>Distribuição de status</h3>
    <div class="grafico"><img src="data:image/png;base64,{graficos['gui_status_pie']}" alt="Status Guaibim Pizza"></div>
    <div class="two-col">
        <div class="grafico"><img src="data:image/png;base64,{graficos['gui_metricas']}" alt="Métricas Guaibim"></div>
        <div class="grafico"><img src="data:image/png;base64,{graficos['gui_vendas_status']}" alt="Vendas por status Guaibim"></div>
    </div>
</div>

<div class="slide">
    <h2>Recomendações</h2>
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
    logger.info("Gerando apresentação de gráficos e insights")

    dados = ler_relatorios()

    logger.info("Criando gráficos...")
    graficos = {}
    graficos["mm_status_bar"] = grafico_barras_status(dados["mm"]["revenda"])
    graficos["mm_status_pie"] = grafico_pizza_status(dados["mm"]["revenda"], "MM Varejo + MM Atacado — Status")
    graficos["mm_metricas"] = grafico_metricas(dados["mm"]["revenda"], "Aceite, treinamentos e vendas — MM")
    graficos["mm_vendas_status"] = grafico_vendas_status(dados["mm"]["detalhe"], "Vendas por status — MM")
    graficos["mm_top_lojas"] = grafico_top_lojas(
        dados["mm"]["loja"], "MM Varejo", "Top 10 lojas MM Varejo — produtos vendidos", n=10, excluir_matriz=True
    )

    graficos["gui_status_pie"] = grafico_pizza_status(dados["guaibim"]["revenda"], "Guaibim — Status")
    graficos["gui_metricas"] = grafico_metricas(dados["guaibim"]["revenda"], "Aceite, treinamentos e vendas — Guaibim")
    graficos["gui_vendas_status"] = grafico_vendas_status(dados["guaibim"]["detalhe"], "Vendas por status — Guaibim")

    logger.info("Gerando insights...")
    insights, recomendacoes = gerar_insights(dados)

    logger.info("Montando HTML...")
    html = montar_html(dados, graficos, insights, recomendacoes)

    hoje_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"apresentacao_mm_guaibim_{hoje_str}.html"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    logger.info(f"Apresentação salva em: {output_path}")


if __name__ == "__main__":
    main()
