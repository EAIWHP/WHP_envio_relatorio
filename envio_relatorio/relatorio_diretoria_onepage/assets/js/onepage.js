/**
 * Relatório Geral Programa +TOP · One Page
 * Consome output/dados_YYYY-MM.json gerado pelo script Python.
 */

const METAS = {
  cadastro: 0.85,
  treinamento: 0.70,
  aceite: 0.70,
};

const FAROL_LIMITES = {
  amarelo: 0.5, // 50% da meta
};

// Mapeamento de exibição dos nomes de revenda (title case uniforme)
const NOME_EXIBICAO_REVENDA = {
  "Angeloni": "Angeloni",
  "Armazem Mateus": "Armazém Mateus",
  "Armazém Mateus": "Armazém Mateus",
  "Becker": "Becker",
  "Bemol": "Bemol",
  "Casas da Água": "Casas da Água",
  "Colombo": "Colombo",
  "Estrela": "Estrela",
  "Formosa": "Formosa",
  "Gazin Atacado": "Gazin Atacado",
  "Gazin Online": "Gazin Online",
  "Gazin Varejo": "Gazin Varejo",
  "Guaibim": "Guaibim",
  "Havan": "Havan",
  "Império": "Império",
  "Imperio": "Império",
  "Jmahfuz": "Jmahfuz",
  "Koerich": "Koerich",
  "Laser Eletro": "Laser Eletro",
  "Lebes": "Lebes",
  "Líder": "Líder",
  "Lider": "Líder",
  "MM Atacado": "MM Atacado",
  "MM Varejo": "MM Varejo",
  "Millena": "Millena",
  "Multiloja": "Multiloja",
  "Nosso Lar": "Nosso Lar",
  "Ramsons": "Ramsons",
  "Sipolatti": "Sipolatti",
  "Solar": "Solar",
  "Solar Magazine": "Solar Magazine",
  "Taqi": "Taqi",
  "Tele Rio": "Tele Rio",
  "Zema": "Zema",
  "Zenir": "Zenir",
};

let dados = null;
let revendasFiltradas = [];
let sortConfig = { coluna: "Ranking", direcao: "desc" };

async function carregarDados() {
  try {
    const response = await fetch(`./output/dados_${dados ? dados.periodo : '2026-07'}.json`);
    if (!response.ok) throw new Error('HTTP ' + response.status);
    dados = await response.json();
    // Aplica formatação de exibição nos nomes de revenda
    dados.revendas.forEach(r => {
      r.Revenda = NOME_EXIBICAO_REVENDA[r.Revenda] || r.Revenda;
    });
    inicializar();
  } catch (e) {
    console.warn('Não foi possível carregar JSON externo.', e);
    document.body.innerHTML = `
      <div class="app">
        <div class="section">
          <h2>⚠️ Não foi possível carregar os dados</h2>
          <p>Abra esta página via servidor local para carregar o JSON atualizado:</p>
          <pre style="background:#f4f4f4;padding:12px;border-radius:8px;">cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/relatorio_diretoria_onepage
python3 -m http.server 8080</pre>
          <p>Depois acesse: <a href="http://localhost:8080">http://localhost:8080</a></p>
        </div>
      </div>
    `;
  }
}

function inicializar() {
  document.getElementById('titulo-periodo').textContent = dados.label;
  document.getElementById('badge-periodo').textContent = dados.label;
  document.getElementById('data-geracao').textContent = new Date().toLocaleDateString('pt-BR');

  revendasFiltradas = [...dados.revendas];

  renderizarKPIs();
  renderizarScorecards();
  renderizarAlertas();
  renderizarGraficos();
  popularFiltros();
  renderizarTabela();
  configurarEventos();
  configurarOrdenacao();
}

function formatarMoeda(valor) {
  if (valor === null || valor === undefined || isNaN(valor)) return '-';
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(valor);
}

function formatarPct(valor) {
  if (valor === null || valor === undefined || isNaN(valor)) return '-';
  return valor.toFixed(1).replace('.', ',') + '%';
}

function formatarInteiro(valor) {
  if (valor === null || valor === undefined || isNaN(valor)) return '-';
  return new Intl.NumberFormat('pt-BR').format(Math.round(valor));
}

function classificarFarol(valor, meta) {
  if (valor === null || valor === undefined || isNaN(valor)) return { cor: 'cinza', label: '-' };
  if (valor >= meta * 100) return { cor: 'verde', label: 'Meta' };
  if (valor >= meta * 100 * FAROL_LIMITES.amarelo) return { cor: 'amarelo', label: 'Atenção' };
  return { cor: 'vermelho', label: 'Crítico' };
}

function calcularMedia(coluna) {
  const valores = dados.revendas.map(r => r[coluna]).filter(v => v !== null && v !== undefined && !isNaN(v));
  if (!valores.length) return 0;
  return valores.reduce((a, b) => a + b, 0) / valores.length;
}

function calcularMediaPonderada(colunaPercentual, colunaPeso) {
  let somaPonderada = 0;
  let somaPesos = 0;
  dados.revendas.forEach(r => {
    const pct = r[colunaPercentual];
    const peso = r[colunaPeso];
    if (pct !== null && pct !== undefined && !isNaN(pct) && peso > 0) {
      somaPonderada += pct * peso;
      somaPesos += peso;
    }
  });
  return somaPesos > 0 ? somaPonderada / somaPesos : 0;
}

function renderizarKPIs() {
  const investimentoTotal = dados.revendas.reduce((acc, r) => acc + (r['Investimento Mês (R$)'] || 0), 0);
  const naoInvestimentoTotal = dados.revendas.reduce((acc, r) => acc + (r['Não Investimento (R$)'] || 0), 0);
  const vendasTotal = dados.revendas.reduce((acc, r) => acc + (r['Vendas'] || 0), 0);
  const revendasCount = dados.revendas.length;
  const rankingMedio = calcularMedia('Ranking');

  const kpis = [
    { label: 'Investimento Mês', value: formatarMoeda(investimentoTotal), sub: 'Total investido' },
    { label: 'Não Investimento', value: formatarMoeda(naoInvestimentoTotal), sub: 'Potencial não realizado' },
    { label: 'Vendas', value: formatarInteiro(vendasTotal), sub: 'Qtd Peças' },
    { label: 'Revendas', value: revendasCount, sub: 'No relatório' },
    { label: 'Ranking Médio', value: rankingMedio.toFixed(2).replace('.', ','), sub: 'Máximo 3 pontos' },
  ];

  const container = document.getElementById('kpis-container');
  container.innerHTML = kpis.map(k => `
    <div class="kpi-card">
      <span class="kpi-label">${k.label}</span>
      <span class="kpi-value">${k.value}</span>
      <span class="kpi-sub">${k.sub}</span>
    </div>
  `).join('');
}

function renderizarScorecards() {
  const mediaCadastro = calcularMediaPonderada('Cadastro (%)', 'Total de CPFs na Hierarquia');
  const mediaTreinamento = calcularMediaPonderada('Treinamento (%)', 'Ativos');
  const mediaAceite = calcularMediaPonderada('Aceite (%)', 'Ativos');

  const scorecards = [
    { title: 'Cadastro', valor: mediaCadastro, meta: METAS.cadastro },
    { title: 'Treinamento', valor: mediaTreinamento, meta: METAS.treinamento },
    { title: 'Aceite', valor: mediaAceite, meta: METAS.aceite },
  ];

  const container = document.getElementById('scorecards-container');
  container.innerHTML = scorecards.map(s => {
    const farol = classificarFarol(s.valor, s.meta);
    const pctMeta = Math.min((s.valor / (s.meta * 100)) * 100, 100);
    return `
      <div class="scorecard">
        <div class="scorecard-header">
          <span class="scorecard-title">${s.title}</span>
          <span class="farol ${farol.cor}"><span class="farol-dot"></span>${farol.label}</span>
        </div>
        <div class="scorecard-value">${formatarPct(s.valor)}</div>
        <div class="scorecard-bar">
          <div class="scorecard-bar-fill ${farol.cor}" style="width: ${pctMeta}%"></div>
        </div>
        <div class="scorecard-meta">Meta: ${formatarPct(s.meta * 100)}</div>
      </div>
    `;
  }).join('');
}

function renderizarAlertas() {
  const abaixo50 = dados.revendas.filter(r =>
    r['Cadastro (%)'] < 50 || r['Treinamento (%)'] < 50 || r['Aceite (%)'] < 50
  );
  const semRanking = dados.revendas.filter(r => r['Ranking'] === 0);
  const maiorNaoInvestimento = [...dados.revendas]
    .sort((a, b) => b['Não Investimento (R$)'] - a['Não Investimento (R$)'])
    .slice(0, 3);

  const container = document.getElementById('alertas-container');
  container.innerHTML = `
    <p><strong>${abaixo50.length} revendas</strong> estão com pelo menos um indicador abaixo de 50%.</p>
    <p><strong>${semRanking.length} revendas</strong> não atingiram nenhuma das 3 metas (0 pontos).</p>
    <p><strong>Top 3 não investimento:</strong> ${maiorNaoInvestimento.map(r => `${r.Revenda} (${formatarMoeda(r['Não Investimento (R$)'])})`).join(', ')}.</p>
  `;
}

function renderizarGraficos() {
  renderizarGraficoInvestimento();
  renderizarGraficoNaoInvestimento();
}

function renderizarGraficoInvestimento() {
  const top10 = [...dados.revendas]
    .sort((a, b) => b['Investimento Mês (R$)'] - a['Investimento Mês (R$)'])
    .slice(0, 10);

  const max = Math.max(...top10.map(r => r['Investimento Mês (R$)']));

  const container = document.getElementById('chart-investimento');
  container.innerHTML = top10.map(r => {
    const h = Math.max((r['Investimento Mês (R$)'] / max) * 160, 4);
    return `
      <div class="chart-bar-item">
        <span class="chart-bar-value">${formatarInteiro(r['Investimento Mês (R$)'])}</span>
        <div class="chart-bar-fill investimento" style="height: ${h}px"></div>
        <span class="chart-bar-label">${r.Revenda}</span>
      </div>
    `;
  }).join('');
}

function renderizarGraficoNaoInvestimento() {
  const top10 = [...dados.revendas]
    .sort((a, b) => b['Não Investimento (R$)'] - a['Não Investimento (R$)'])
    .slice(0, 10);

  const max = Math.max(...top10.map(r => r['Não Investimento (R$)']), 1);

  const container = document.getElementById('chart-nao-investimento');
  container.innerHTML = top10.map(r => {
    const h = Math.max((r['Não Investimento (R$)'] / max) * 160, 4);
    return `
      <div class="chart-bar-item">
        <span class="chart-bar-value">${formatarInteiro(r['Não Investimento (R$)'])}</span>
        <div class="chart-bar-fill nao-investimento" style="height: ${h}px"></div>
        <span class="chart-bar-label">${r.Revenda}</span>
      </div>
    `;
  }).join('');
}

function popularFiltros() {
  const regionais = [...new Set(dados.revendas.map(r => r.Regional).filter(Boolean))].sort();
  const select = document.getElementById('filtro-regional');
  select.innerHTML = '<option value="">Todas</option>' + regionais.map(r => `<option value="${r}">${r}</option>`).join('');
}

function filtrarRevendas() {
  const busca = document.getElementById('filtro-revenda').value.toLowerCase();
  const regional = document.getElementById('filtro-regional').value;
  const ranking = document.getElementById('filtro-ranking').value;

  revendasFiltradas = dados.revendas.filter(r => {
    const matchBusca = !busca || r.Revenda.toLowerCase().includes(busca);
    const matchRegional = !regional || r.Regional === regional;
    const matchRanking = ranking === '' || String(r.Ranking) === ranking;
    return matchBusca && matchRegional && matchRanking;
  });

  ordenarDados();
  renderizarTabela();
}

function ordenarDados() {
  const coluna = sortConfig.coluna;
  const dir = sortConfig.direcao === 'asc' ? 1 : -1;

  revendasFiltradas.sort((a, b) => {
    let va = a[coluna];
    let vb = b[coluna];

    if (va === null || va === undefined) va = dir === 1 ? Infinity : -Infinity;
    if (vb === null || vb === undefined) vb = dir === 1 ? Infinity : -Infinity;

    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();

    if (va < vb) return -1 * dir;
    if (va > vb) return 1 * dir;
    return 0;
  });
}

function renderizarTabela() {
  const tbody = document.getElementById('tabela-body');
  tbody.innerHTML = revendasFiltradas.map(r => {
    const farolCadastro = classificarFarol(r['Cadastro (%)'], METAS.cadastro);
    const farolTreinamento = classificarFarol(r['Treinamento (%)'], METAS.treinamento);
    const farolAceite = classificarFarol(r['Aceite (%)'], METAS.aceite);
    const varLm = r['Var LM (%)'];
    const varYoy = r['Var YOY (%)'];

    return `
      <tr ${r.Ranking === 3 ? 'class="linha-destaque"' : ''}>
        <td><span class="ranking-badge rank-${r.Ranking}">${r.Ranking}</span></td>
        <td><strong>${r.Revenda}</strong></td>
        <td>${r.Regional || '-'}</td>
        <td class="numeric">
          <span class="farol ${farolCadastro.cor}"><span class="farol-dot"></span>${formatarPct(r['Cadastro (%)'])}</span>
        </td>
        <td class="numeric">
          <span class="farol ${farolTreinamento.cor}"><span class="farol-dot"></span>${formatarPct(r['Treinamento (%)'])}</span>
        </td>
        <td class="numeric">
          <span class="farol ${farolAceite.cor}"><span class="farol-dot"></span>${formatarPct(r['Aceite (%)'])}</span>
        </td>
        <td class="numeric">${formatarMoeda(r['Investimento Mês (R$)'])}</td>
        <td class="numeric">${formatarMoeda(r['Não Investimento (R$)'])}</td>
        <td class="numeric">${formatarInteiro(r['Vendas'])}</td>
        <td class="numeric ${varLm > 0 ? 'positiva' : varLm < 0 ? 'negativa' : 'neutra'}">${varLm !== null && varLm !== undefined ? (varLm > 0 ? '+' : '') + formatarPct(varLm) : '-'}</td>
        <td class="numeric ${varYoy > 0 ? 'positiva' : varYoy < 0 ? 'negativa' : 'neutra'}">${varYoy !== null && varYoy !== undefined ? (varYoy > 0 ? '+' : '') + formatarPct(varYoy) : '-'}</td>
      </tr>
    `;
  }).join('');
}

function configurarEventos() {
  document.getElementById('filtro-revenda').addEventListener('input', filtrarRevendas);
  document.getElementById('filtro-regional').addEventListener('change', filtrarRevendas);
  document.getElementById('filtro-ranking').addEventListener('change', filtrarRevendas);
  document.getElementById('btn-limpar').addEventListener('click', () => {
    document.getElementById('filtro-revenda').value = '';
    document.getElementById('filtro-regional').value = '';
    document.getElementById('filtro-ranking').value = '';
    filtrarRevendas();
  });
  document.getElementById('btn-exportar-csv').addEventListener('click', exportarCSV);
}

function configurarOrdenacao() {
  const ths = document.querySelectorAll('#tabela-revendas th[data-sort]');
  ths.forEach(th => {
    th.addEventListener('click', () => {
      const coluna = th.dataset.sort;
      if (sortConfig.coluna === coluna) {
        sortConfig.direcao = sortConfig.direcao === 'asc' ? 'desc' : 'asc';
      } else {
        sortConfig = { coluna, direcao: 'desc' };
      }
      ths.forEach(t => t.classList.remove('sorted-asc', 'sorted-desc'));
      th.classList.add(sortConfig.direcao === 'asc' ? 'sorted-asc' : 'sorted-desc');
      ordenarDados();
      renderizarTabela();
    });
  });
}

function exportarCSV() {
  const colunas = [
    'Ranking', 'Revenda', 'Regional', 'Cadastro (%)', 'Treinamento (%)', 'Aceite (%)',
    'Investimento Mês (R$)', 'Não Investimento (R$)', 'Vendas', 'Vendas LM', 'Var LM (%)',
    'Vendas L3M', 'Vendas YOY', 'Var YOY (%)', 'Total de CPFs na Hierarquia'
  ];
  let csv = colunas.join(';') + '\n';
  revendasFiltradas.forEach(r => {
    const row = colunas.map(c => {
      let v = r[c];
      if (v === null || v === undefined) return '';
      if (typeof v === 'number') return String(v).replace('.', ',');
      return String(v);
    });
    csv += row.join(';') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `relatorio_geral_programa_top_${dados.periodo}.csv`;
  link.click();
}

// Inicialização
carregarDados();
