# Relatório de Diretoria One Page +TOP

Geração de relatório mensal one page para apresentação na Staff do Zanatta e na WConnection.

## Estrutura

```
relatorio_diretoria_onepage/
├── gerar_relatorio_diretoria.py   # Script principal
├── config.py                       # Constantes e caminhos
├── utils.py                        # Helpers
├── Campanha_+TOP_Junho_2026 (2).xlsx  # Planilha consolidada de referência
├── index.html                      # One page HTML
├── assets/
│   ├── css/onepage.css             # Estilos
│   └── js/onepage.js               # Lógica da página
├── output/                         # Arquivos gerados
│   ├── Relatorio_Diretoria_+TOP_Junho_2026_V1.xlsx
│   └── dados_2026-06.json
└── README.md
```

## Como usar

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/relatorio_diretoria_onepage
python3 gerar_relatorio_diretoria.py
```

O script gera:
- `output/Relatorio_Diretoria_+TOP_Junho_2026_V1.xlsx` — Excel V1 formatado.
- `output/dados_2026-06.json` — dados para a one page HTML.

## Estrutura do relatório (Staff Zanatta)

| Coluna | Descrição |
|---|---|
| Revenda | Nome da revenda |
| Regional | Regional da revenda |
| Cadastro (%) | Ativos no cadastro / Total na hierarquia ativa |
| Treinamento (%) | Ativos que concluíram os 2 cursos obrigatórios / Base ativa |
| Aceite (%) | Ativos com aceite mensal / Base ativa |
| Investimento Mês (R$) | Total de pontuação da revenda na planilha Campanha +TOP (1 ponto = R$ 1,00) |
| Não Investimento (R$) | Potencial Full - Investimento Mês |
| Vendas | Qtd Peças vendidas no mês (coluna da planilha Campanha +TOP) |
| Vendas LM | Mês anterior |
| Var LM (%) | Variação percentual do mês atual vs mês anterior |
| Vendas L3M | Média dos últimos 3 meses |
| Vendas YOY | Mesmo mês do ano anterior |
| Var YOY (%) | Variação percentual do mês atual vs mesmo mês do ano anterior |
| Ranking | 1 ponto por meta atingida (Cadastro ≥85%, Treinamento ≥70%, Aceite ≥70%). Máximo 3 pontos. |

## Cálculos

### Investimento Mês
Total de pontos da revenda na planilha `Campanha_+TOP_Junho_2026 (2).xlsx` (coluna **Total**), convertido em reais (1 ponto = R$ 1,00).

### Vendas
Coluna **Vendas Totais Bases Revendas** da planilha Campanha +TOP.

### Potencial Full
```
Total de CPFs na hierarquia ativa × média de Pontos Vendas (Vendedor) por CPF ativo apto
```

Se a revenda não tiver CPFs aptos, usa a média geral de pontos por apto de todas as revendas.

### Não Investimento
```
max(0, Potencial Full - Investimento Mês)
```

## Fontes de dados

- Planilha consolidada: `relatorio_diretoria_onepage/Campanha_+TOP_Junho_2026 (2).xlsx`
- Hierarquia ativa: `envio_relatorio/bases/bases_cadastro_hierarquia/`
- Cadastro: `envio_relatorio/bases/cadastro.xlsx`
- Treinamentos: `envio_relatorio/bases/Base_treinamentos.xlsx`
- Aceites: `envio_relatorio/bases/WHP_Aceite_Mensal_*.xlsx`
- Vendas comparativas: `bases/Vendas Processadas/`

## Farol de KPIs

- **Verde:** meta atingida.
- **Amarelo:** entre 50% e a meta.
- **Vermelho:** abaixo de 50%.

## Próximos passos

- [x] Criar a one page HTML reutilizando o visual do `painel-ranking-mensal`.
- [x] Adicionar gráficos e filtros interativos.
- [x] Refazer cálculos de acordo com a planilha Campanha +TOP.
- [x] Alinhar estrutura com a planilha Staff Zanatta.
- [ ] Validar cálculo do Potencial Full / Não Investimento.
- [ ] Automatizar a detecção do mês de referência.
- [ ] Adicionar série histórica/evolução mensal.
- [ ] Tornar o nome do arquivo da planilha Campanha dinâmico por mês.
