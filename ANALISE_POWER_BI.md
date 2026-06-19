# 📊 Análise Completa do Dashboard Power BI — Programa +TOP

**Arquivo:** `Dashboard_TOP_V7.pbix`  
**Tamanho:** 42 MB  
**Páginas:** 10  
**Tabelas no modelo:** 30  
**Medidas DAX:** 79  
**Data da análise:** 2026-06-08  

---

# PARTE 1: ESTRUTURA DO RELATÓRIO

## 1.1 Páginas do Dashboard

| # | Nome da Página | Visualizações | Tipos Principais | Foco |
|---|---------------|---------------|-----------------|------|
| 1 | **PERFORMANCE VENDAS** | 65 | pivotTable (5), lineChart (1), lineClusteredColumnComboChart (2), slicer (5) | Vendas, pontuação, variações, análise por revenda/regional/SKU/categoria |
| 2 | **CADASTROS-ACESSOS** | 84 | pivotTable (7), tableEx (3), gauge (1), treemap (1), scatterChart (1), lineChart (1) | Participantes, status, acessos, navegação, aceite |
| 3 | **RESUMO PONTOS** | 57 | cardVisual (5), pivotTable (4), lineClusteredColumnComboChart (2) | Créditos, resgates, expirados, saldo |
| 4 | **PONTUACAO INDIVIDUAL** | 42 | cardVisual (5), pivotTable (1) | Detalhe por participante (CPF/nome) |
| 5 | **RESGATES DETALHADO** | 49 | pivotTable (4), cardVisual (4), lineChart (1), hundredPercentStackedColumnChart (1) | Pedidos, participante, CPF, valor, data, produto |
| 6 | **CHAMADOS GERAL** | 58 | pivotTable (2), tableEx (1), pieChart (1), lineStackedColumnComboChart (1), lineClusteredColumnComboChart (2) | Total, status, motivos, submotivos |
| 7 | **SUPERTOP** | 83 | pivotTable (4), tableEx (1), lineClusteredColumnComboChart (2), lineChart (1) | Vendas e pontos SUPERTOP |
| 8 | **POTENCIAL** | 49 | pivotTable (2) | Potencial de pontuação e aproveitamento |
| 9 | **TREINAMENTOS** | 51 | tableEx (2), lineChart (1), barChart (1) | Obrigatórios, concluintes, taxa de conclusão |
| 10 | **ENQUETE** | 46 | columnChart (1), lineStackedColumnComboChart (1), hundredPercentStackedColumnChart (1) | Avaliação da plataforma |

**Total de visualizações:** 584 (incluindo shapes, textboxes e imagens)  
**Visualizações de dados úteis:** ~230

---

## 1.2 Tabelas do Modelo de Dados (30)

### Tabelas Fato (transacionais)
| # | Tabela | Descrição | Base Correspondente |
|---|--------|-----------|---------------------|
| 1 | **06_Base_Vendas** | Vendas de produtos participantes | `bases/Vendas Processadas/AAAA_MM.xlsx` |
| 2 | **02_Cadastro** | Cadastro mestre de participantes | `bases/Tabelas/cadastro.xlsx` |
| 3 | **03_Acessos** | Log de acessos à plataforma | `bases/Tabelas/acessos.xlsx` |
| 4 | **tab_pontos_por_cpf** | Pontos por participante | `bases/Tabelas/pontos_por_cpf_.xlsx` (aba PONTOS POR CPF) |
| 5 | **tab_resgates_detalhado** | Resgates detalhados | `bases/Tabelas/resgates_detalhado.xlsx` |
| 6 | **chamados** | Chamados ao suporte | `bases/Tabelas/Chamados_catalogo.xlsx` |
| 7 | **bases_potencial** | Potencial de pontuação | `bases/Tabelas/bases_potencial/` |
| 8 | **Base_treinamentos** | Treinamentos concluídos | `bases/Tabelas/Base_treinamentos.xlsx` |
| 9 | **enquete** | Pesquisa de satisfação | `bases/Tabelas/enquete.xlsx` |
| 10 | **tab_chamados_sac_plataforma** | Chamados SAC/Plataforma | `bases/Tabelas/base_tratada.xlsx` (provável) |
| 11 | **Cadastros_revenda_PDV** | Cadastro com vínculo revenda/loja | `bases/Tabelas/WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx` |

### Tabelas Dimensão (lookups)
| # | Tabela | Descrição | Base Correspondente |
|---|--------|-----------|---------------------|
| 12 | **dCalendario** | Calendário de datas (vendas) | Gerada no Power Query |
| 13 | **dCalendario_VCAD** | Calendário de datas (cadastros/acessos) | Gerada no Power Query |
| 14 | **dCalendario_Resgates** | Calendário de datas (resgates) | Gerada no Power Query |
| 15 | **dCalendario_Conclusao** | Calendário de datas (treinamentos) | Gerada no Power Query |
| 16 | **dCalendario_aceites** | Calendário de datas (aceites) | Gerada no Power Query |
| 17 | **dCalendar_resgate** | Calendário alternativo (resgates) | Gerada no Power Query |
| 18 | **dCalendar_ChamadosOcorrencias** | Calendário (chamados/ocorrências) | Gerada no Power Query |
| 19 | **dProduto** | Dimensão de produtos (SKU, Categoria, Subcategoria) | Gerada no Power Query ou tabela auxiliar |
| 20 | **tab_nome_revenda** | DE/PARA de nomes de revendas | `bases/Tabelas/tab_nome_revendas.xlsx` |

### Tabelas de Aceites (mensais)
| # | Tabela | Mês | Base Correspondente |
|---|--------|-----|---------------------|
| 21 | **OUT_2025** | Outubro/2025 | `WHP_Aceite_Mensal_*.xlsx` (aba OUT_2025) |
| 22 | **NOV_2025** | Novembro/2025 | `WHP_Aceite_Mensal_*.xlsx` (aba NOV_2025) |
| 23 | **JAN_2026** | Janeiro/2026 | `WHP_Aceite_Mensal_*.xlsx` (aba JAN_2026) |
| 24 | **FEV_2026** | Fevereiro/2026 | `WHP_Aceite_Mensal_*.xlsx` (aba FEV_2026) |
| 25 | **MAR_2026** | Março/2026 | `WHP_Aceite_Mensal_*.xlsx` (aba MAR_2026) |
| 26 | **ABR_2026** | Abril/2026 | `WHP_Aceite_Mensal_*.xlsx` (aba ABR_2026) |

### Tabelas Auxiliares
| # | Tabela | Descrição |
|---|--------|-----------|
| 27 | **01_medidas** / **01_Medidas** | Tabela de medidas DAX calculadas |
| 28 | **Fale_Conosco** | Contatos/Fale Conosco | `bases/Tabelas/Fale_Conosco.xlsx` |
| 29 | **Tabela_Status_Treinamento** | Tabela auxiliar de status de treinamento |
| 30 | **DEZ_2025** | Dezembro/2025 (aceite) |

---

## 1.3 Medidas DAX (79 identificadas)

### 📈 Medidas de Vendas
| Medida | Contexto de Uso |
|--------|----------------|
| Total de Vendas | Performance Vendas, SUPERTOP |
| Total de Vendas SuperTOP | Performance Vendas, SUPERTOP |
| Vendas Último Mês | Performance Vendas |
| Vendas SuperTOP Último Mês | SUPERTOP |
| Vendas Mês Anterior ao Último | Performance Vendas |
| Vendas Mês Anterior ao Último SUPERTOP | SUPERTOP |
| Vendas Último Mês (Ano Anterior) | Performance Vendas (YoY) |
| Vendas Último Mês (Ano Anterior) SUPERTOP | SUPERTOP (YoY) |
| Variação Vendas Último Mês % | Performance Vendas |
| Variação Vendas YoY % (Último Mês) | Performance Vendas |
| Variação Vendas Último Mês SUPERTOP % | SUPERTOP |
| Variação Vendas YoY % (Último Mês) SUPERTOP | SUPERTOP |
| Valor Crescimento Absoluto | Performance Vendas |
| Valor Crescimento Absoluto SUPERTOP | SUPERTOP |
| % Representatividade no Crescimento | Performance Vendas |
| % Representatividade no Crescimento SUPERTOP | SUPERTOP |
| Vendas Incentivadas | SUPERTOP |
| Pontos de Venda Último Mês | Performance Vendas, SUPERTOP, Potencial |
| Pontos de Venda Último Mês SUPERTOP | SUPERTOP |
| Total de Pontos de Venda | Performance Vendas |
| Total de Pontos de Venda SUPERTOP | SUPERTOP |

### 📊 Medidas Agregadas por Período
| Medida | Contexto |
|--------|----------|
| Total Vendas Janeiro 26 | Performance Vendas |
| Total Vendas Fevereiro 26 | Performance Vendas |
| Total Vendas Marco 26 | Performance Vendas |
| Total de Vendas SuperTOP Janeiro 2026 | SUPERTOP |
| Total de Vendas SuperTOP Fevereiro 2026 | SUPERTOP |
| Total de Vendas SuperTOP Marco 2026 | SUPERTOP |
| Total de Pontos de Venda Janeiro 2026 | Performance Vendas |
| Total de Pontos de Venda Fevereiro 2026 | Performance Vendas |
| Total de Pontos de Venda Marco 2026 | Performance Vendas |
| % SuperTOP Jan/26 | Performance Vendas |
| % SuperTOP Fev/26 | Performance Vendas |
| % SuperTOP Mar/26 | Performance Vendas |
| Total de Vendas Out25 a Mar26 | Performance Vendas |
| Total de Vendas SuperTOP Out25 a Mar26 | SUPERTOP |
| Total de Pontos de Venda Out25 a Mar26 | Performance Vendas |
| % Vendas SuperTOP Out25 a Mar26 | Performance Vendas |

### 👥 Medidas de Cadastro
| Medida | Contexto |
|--------|----------|
| N Participantes | Cadastros-Acessos |
| Taxa de Ativação | Cadastros-Acessos |
| Nº Participantes | Cadastros-Acessos |
| N° Participantes Ativos | Cadastros-Acessos |
| Nº Participantes Ativos Acumulados | Cadastros-Acessos |
| % participantes ativos | Cadastros-Acessos |
| % participantes inativos | Cadastros-Acessos |
| % Pre Cadastro | Cadastros-Acessos |
| % Indicado Moderado | Cadastros-Acessos |
| % Reprovado | Cadastros-Acessos |
| % Em moderacao | Cadastros-Acessos |
| % Aguardando aprovação | Cadastros-Acessos |
| % Total Geral Status | Cadastros-Acessos |
| Cadastrados Out25 a Mar26 | Performance Vendas |
| Cadastrados Ativos Out25 a Mar26 | Performance Vendas |
| % Cadastrados Ativos Out25 a Mar26 | Performance Vendas |

### 🎯 Medidas de Acessos
| Medida | Contexto |
|--------|----------|
| N° Acessos Totais | Cadastros-Acessos |
| N° Acessos Únicos | Cadastros-Acessos |
| Frequência de Acessos | Cadastros-Acessos |

### 💰 Medidas de Pontos
| Medida | Contexto |
|--------|----------|
| Total Pontos Creditados | Resumo Pontos, Pontuação Individual |
| Total Pontos Resgatados | Resumo Pontos, Pontuação Individual |
| Total Pontos Expirados | Resumo Pontos, Pontuação Individual |
| Saldo Total Pontos | Resumo Pontos, Pontuação Individual |
| % Consumo de Pontos | Resumo Pontos, Pontuação Individual |
| Participantes Pontuados | Resumo Pontos |
| Participantes com Saldo | Resumo Pontos |
| Qtd Participantes Pontos | Resumo Pontos |

### 🎁 Medidas de Resgates
| Medida | Contexto |
|--------|----------|
| Total Resgates | Resgates Detalhado |
| Total Resgates Detalhado | Resgates Detalhado |
| Total Resgates VALOR | Resgates Detalhado |
| Qtd Participantes com Resgates | Resgates Detalhado |
| Média Valor por Pedido | Resgates Detalhado |
| Média Valor Resgatado por Pessoa | Resgates Detalhado |

### 📞 Medidas de Chamados
| Medida | Contexto |
|--------|----------|
| Total Chamados | Chamados Geral |
| Chamados Em Andamento | Chamados Geral |
| Chamados Resolvidos | Chamados Geral |
| Total Contatos | Chamados Geral |
| Contatos Plataforma | Chamados Geral |

### 🏆 Medidas de SUPERTOP
| Medida | Contexto |
|--------|----------|
| Representatividade % SuperTOP | SUPERTOP |

---

# PARTE 2: MAPEAMENTO BASES ↔ POWER BI

## 2.1 Tabelas Fato: Base vs. Tabela PBI

| Base Física | Tabela PBI | Chave de Ligação | Status |
|-------------|-----------|-----------------|--------|
| `cadastro.xlsx` | `02_Cadastro` | `cpf/cnpj` (texto) | ✅ OK |
| `acessos.xlsx` | `03_Acessos` | `Cpf/Cnpj` (int → perde zeros) | ⚠️ Problema conhecido |
| `pontos_por_cpf_.xlsx` | `tab_pontos_por_cpf` | `CPFCNPJ` (int) | ⚠️ Problema conhecido |
| `resgates_detalhado.xlsx` | `tab_resgates_detalhado` | `CPF` (int) | ⚠️ Problema conhecido |
| `Base_treinamentos.xlsx` | `Base_treinamentos` | `CPF` (int) | ⚠️ Problema conhecido |
| `Chamados_catalogo.xlsx` | `chamados` | `DOCUMENTO` | ✅ OK |
| `enquete.xlsx` | `enquete` | `cpf` (int) | ⚠️ Problema conhecido |
| `Vendas Processadas/*.xlsx` | `06_Base_Vendas` | `CPF` (int), `CNPJ` | ⚠️ CPF como int |
| `WHP_Aceite_Mensal_*.xlsx` | `OUT_2025` a `ABR_2026` | `CPF` (int) | ⚠️ Problema conhecido |
| `base_tratada.xlsx` | `tab_chamados_sac_plataforma` | `cpf` (int) | ⚠️ Problema conhecido |
| `WHP_PROD_Cadastro_Revenda_*.xlsx` | `Cadastros_revenda_PDV` | `Cpf` (int) | ⚠️ Problema conhecido |

## 2.2 Tabelas Dimensão

| Tabela PBI | Base Física | Função |
|-----------|------------|--------|
| `tab_nome_revenda` | `tab_nome_revendas.xlsx` | DE/PARA de nomes de revendas |
| `dProduto` | Tabela auxiliar/EAN/SKU | Dimensão de produtos |
| `dCalendario*` | Geradas no Power Query | 6 calendários para diferentes contextos de data |

## 2.3 Tabela de Medidas

| Tabela | Função |
|--------|--------|
| `01_medidas` / `01_Medidas` | Centraliza todas as medidas DAX calculadas (79 medidas) |

> ⚠️ **Nota:** Há duas tabelas de medidas (`01_Medidas` e `01_medidas`) — provavelmente uma é antiga e a outra é a atual. Algumas visualizações ainda referenciam a antiga.

---

# PARTE 3: RELACIONAMENTOS DO MODELO

## 3.1 Relacionamentos Identificados (via DiagramLayout)

```
06_Base_Vendas ──[SKU]──► dProduto
06_Base_Vendas ──[CPF]──► 02_Cadastro
06_Base_Vendas ──[Data]──► dCalendario

02_Cadastro ──[grupo]──► tab_nome_revenda
02_Cadastro ──[CPF]──► Base_treinamentos
02_Cadastro ──[CPF]──► tab_chamados_sac_plataforma
02_Cadastro ──[CPF]──► Cadastros_revenda_PDV

02_Cadastro ──[CPF]──► tab_pontos_por_cpf
02_Cadastro ──[CPF]──► tab_resgates_detalhado
02_Cadastro ──[CPF]──► enquete
02_Cadastro ──[CPF]──► 03_Acessos

02_Cadastro ──[CPF]──► OUT_2025/NOV_2025/DEZ_2025/JAN_2026/FEV_2026/MAR_2026/ABR_2026

03_Acessos ──[Data Acesso]──► dCalendario_VCAD
tab_resgates_detalhado ──[DATA RESGATE]──► dCalendario_Resgates / dCalendar_resgate
chamados ──[DATA]──► dCalendar_ChamadosOcorrencias
tab_chamados_sac_plataforma ──[data]──► dCalendar_ChamadosOcorrencias
Base_treinamentos ──[Conclusão]──► dCalendario_Conclusao
enquete ──[dataInclusao]──► dCalendario (ou dimensão própria)

bases_potencial ──[revenda]──► tab_nome_revenda
bases_potencial ──[Data]──► dCalendario

Cadastros_revenda_PDV ──[CNPJ_Revenda]──► Distribuidor_PDV (se existir)
```

## 3.2 Análise dos Calendários (6 tabelas de data!)

| Calendário | Contexto | Base Relacionada |
|-----------|----------|-----------------|
| `dCalendario` | Vendas | `06_Base_Vendas[data venda]` |
| `dCalendario_VCAD` | Cadastros/Acessos | `03_Acessos[Data Acesso]` |
| `dCalendario_Resgates` | Resgates | `tab_resgates_detalhado[DATA RESGATE]` |
| `dCalendar_resgate` | Resgates (alternativo) | `tab_resgates_detalhado[DATA RESGATE]` |
| `dCalendario_Conclusao` | Treinamentos | `Base_treinamentos[Conclusão]` |
| `dCalendario_aceites` | Aceites | `Base_aceites[DataAceite]` |
| `dCalendar_ChamadosOcorrencias` | Chamados/Ocorrências | `chamados[DATA]`, `tab_chamados_sac_plataforma[data]` |

> ⚠️ **Observação:** Há duas tabelas de calendário para resgates (`dCalendario_Resgates` e `dCalendar_resgate`) — uma pode ser redundante.

---

# PARTE 4: ANÁLISE CRÍTICA — Diferenças e Problemas

## 4.1 Problemas Confirmados

### 🔴 P1: CPF como Número em Todas as Bases (exceto Cadastro)

| Base Física | Campo CPF no PBI | Tipo | Problema |
|------------|-----------------|------|----------|
| `cadastro.xlsx` | `cpf/cnpj` | **Texto** ✅ | Preserva zeros |
| `acessos.xlsx` | `Cpf/Cnpj` | Int | ❌ Perde zeros |
| `pontos_por_cpf_.xlsx` | `CPFCNPJ` | Int | ❌ Perde zeros |
| `resgates_detalhado.xlsx` | `CPF` | Int | ❌ Perde zeros |
| `Base_treinamentos.xlsx` | `CPF` | Int | ❌ Perde zeros |
| `Vendas Processadas/*.xlsx` | `CPF` | Int | ❌ Perde zeros |
| `enquete.xlsx` | `cpf` | Int | ❌ Perde zeros |

**Impacto:** O Power BI provavelmente normaliza os CPFs (removendo zeros) em todas as tabelas exceto no cadastro, o que pode causar:
- Participantes com CPF começando em zero não aparecem nos relacionamentos
- Contagem incorreta de CPFs distintos
- Dados de participantes "sumindo" nos cruzamentos

### 🔴 P2: Duas Tabelas de Medidas

- `01_Medidas` (capitalizado) — usada em algumas visualizações
- `01_medidas` (minúsculo) — usada em outras visualizações

**Impacto:** Risco de inconsistência se as medidas estiverem duplicadas ou desatualizadas em uma das tabelas.

### 🔴 P3: Duas Tabelas de Calendário para Resgates

- `dCalendario_Resgates`
- `dCalendar_resgate`

**Impacto:** Redundância. Uma pode ser desnecessária.

### 🟡 P4: Aceites como Tabelas Separadas (não consolidadas)

Cada mês de aceite é uma tabela separada no modelo:
- `OUT_2025`, `NOV_2025`, `DEZ_2025`, `JAN_2026`, `FEV_2026`, `MAR_2026`, `ABR_2026`

**Impacto:** O modelo cresce a cada mês. Ideal seria consolidar em uma única tabela `Base_aceites` com coluna de mês.

### 🟡 P5: Base de Treinamentos sem Participantes "Não Concluídos"

Como confirmado na análise da base, a `Base_treinamentos` só tem quem concluiu (Progresso = 100%).

**Impacto no PBI:** O dashboard não consegue mostrar quem está "em andamento" ou "não iniciou". A taxa de conclusão é calculada por exclusão (Ativos - Concluídos = Não concluíram).

### 🟡 P6: Campanha SUPERTOP

Como informado pelo usuário, nem todos os meses têm produtos SUPERTOP.

**Impacto no PBI:** A página SUPERTOP pode ficar vazia ou com zeros em meses sem investimento (ABR/2026, MAI/2026).

### 🟢 P7: BLACKLIST não está no modelo

A base `BLACKLIST 110526.xlsx` (revendas que não participam de campanhas extras) **não aparece** nas 30 tabelas do modelo.

**Impacto:** O dashboard pode estar considerando todas as revendas nas campanhas extras, sem filtrar as da Blacklist. Isso pode distorcer KPIs de campanhas extras/sazonais.

---

## 4.2 Inconsistências Bases ↔ PBI

| Base Física | O que a base tem | O que o PBI mostra | Diferença |
|------------|-----------------|-------------------|-----------|
| `cadastro.xlsx` | 21,144 participantes (7 status) | Filtra por status | Não mostra Pré-Cadastrados, Reprovados, etc. em KPIs principais |
| `pontos_por_cpf_.xlsx` | 27,388 CPFs (inclui edição anterior) | Total créditos, saldo | Dashboard não distingue pontos da edição anterior vs. atual |
| `Base_treinamentos.xlsx` | 65,013 conclusões, 137 cursos | Taxa de conclusão | Não mostra progresso parcial (30%, 50%, etc.) |
| `Vendas Processadas/*.xlsx` | Dados desde jan/2024 | YoY, variações | Mistura dados de edição anterior sem distinção explícita |

---

# PARTE 5: AUTO-PERGUNTAS E RESPOSTAS (POWER BI)

## P1: Por que há 6 calendários no modelo?
**R:** Cada calendário serve a um contexto de data diferente:
- `dCalendario` → datas de vendas
- `dCalendario_VCAD` → datas de acessos/cadastros
- `dCalendario_Resgates` / `dCalendar_resgate` → datas de resgates (redundantes)
- `dCalendario_Conclusao` → datas de conclusão de treinamentos
- `dCalendario_aceites` → datas de aceite
- `dCalendar_ChamadosOcorrencias` → datas de chamados

Isso é uma prática comum em modelos PBI para isolar filtros de data por contexto, mas pode ser otimizado.

## P2: Por que a página SUPERTOP tem 83 visualizações?
**R:** É a página com mais visualizações. Provavelmente porque:
- Mostra dados de SUPERTOP por mês (Jan, Fev, Mar, etc.)
- Tem muitos KPIs similares à página Performance Vendas, mas filtrados para SUPERTOP
- Em meses sem SUPERTOP (como ABR/2026), os visuais ficam vazios

## P3: Como o PBI trata o CPF como chave de relacionamento?
**R:** O cadastro tem CPF como texto (preserva zeros). Todas as outras bases têm CPF como número (perdem zeros). O PBI provavelmente:
1. Normaliza o CPF do cadastro (remove zeros) OU
2. Converte os CPFs numéricos para texto e preenche com zeros

Se a normalização não for feita corretamente, participantes com CPF começando em zero "somem" dos relacionamentos.

## P4: Por que há duas tabelas de medidas (`01_Medidas` e `01_medidas`)?
**R:** Provavelmente uma renomeação que não foi aplicada em todas as visualizações. A tabela antiga (`01_Medidas`) ainda é referenciada por alguns visuais, enquanto a nova (`01_medidas`) é usada pela maioria.

## P5: A BLACKLIST está sendo usada no PBI?
**R:** Não. A tabela `BLACKLIST 110526.xlsx` **não aparece** entre as 30 tabelas do modelo. Isso significa que o dashboard **não filtra** as revendas da Blacklist nas campanhas extras. Se houver campanhas extras sendo analisadas no PBI, os dados podem estar incorretos.

## P6: Como é calculada a "Taxa de Ativação" no PBI?
**R:** Provavelmente: `DIVIDE([N° Participantes Ativos], [Nº Participantes])`
- Numerador: CPFs com status = "Ativo"
- Denominador: Total de CPFs no cadastro
- Resultado: 63.3% (13.375 / 21.144)

## P7: Como é calculada a "Variação Vendas YoY %"?
**R:** Probável fórmula DAX:
```dax
Variação Vendas YoY % = 
DIVIDE(
    [Vendas Último Mês] - [Vendas Último Mês (Ano Anterior)],
    [Vendas Último Mês (Ano Anterior)],
    0
)
```
Usa o calendário `dCalendario` para navegar entre anos.

## P8: Por que a página de Treinamentos não tem medidas da tabela `01_medidas`?
**R:** A página Treinamentos usa campos calculados diretamente da tabela `Base_treinamentos`:
- `Base_treinamentos.% Treinamentos concluidos`
- `Base_treinamentos.Ativos com Curso Obrigatório Concluído`

Essas são provavelmente colunas calculadas no Power Query ou medidas criadas na própria tabela, não na tabela `01_medidas`.

## P9: Como o PBI diferencia campanha base vs. campanhas extras?
**R:** Pelo que foi identificado, o PBI **não diferencia** explicitamente. A base de vendas (`06_Base_Vendas`) traz todas as vendas, e a diferenciação seria feita por:
- Produtos SUPERTOP (quando existe)
- Aceites específicos (não implementado no PBI)
- BLACKLIST (não está no modelo)

A campanha base é o "padrão", e campanhas extras apareceriam como filtros ou páginas adicionais.

## P10: O que acontece em Setembro/2025 no PBI?
**R:** A base de vendas de set/2025 não existe. No PBI:
- A página Performance Vendas provavelmente mostra "em branco" ou zero para set/2025
- Gráficos de tendência mostram uma queda/"buraco" nesse mês
- YoY de set/2026 (quando chegar) não terá comparação com set/2025

---

# PARTE 6: INSIGHTS DO POWER BI

## Insight 1: BLACKLIST Ausente no Modelo
A Blacklist (`BLACKLIST 110526.xlsx`) não está no modelo PBI. Se o dashboard analisa campanhas extras, os dados incluem revendas que deveriam ser excluídas.

## Insight 2: CPF como Número em 10 das 11 Tabelas Fato
Apenas o cadastro tem CPF como texto. Todas as outras bases têm CPF como número. Isso requer normalização no Power Query.

## Insight 3: 6 Calendários podem ser Consolidados
A maioria dos calendários poderia ser unificada em um único calendário mestre com múltiplas relações ativas/inativas.

## Insight 4: Aceites Mensais como Tabelas Separadas
Cada mês de aceite é uma tabela separada. Isso aumenta o modelo a cada mês. Ideal: consolidar em uma tabela única com coluna de mês/ano.

## Insight 5: Duas Tabelas de Medidas = Risco de Inconsistência
A existência de `01_Medidas` e `01_medidas` pode causar confusão e manutenção duplicada.

---

# PARTE 7: RECOMENDAÇÕES

### Prioridade Alta
1. **Adicionar BLACKLIST ao modelo** — Criar relação com Cadastro/Vendas para filtrar revendas em campanhas extras
2. **Consolidar calendários** — Reduzir de 6 para 1-2 calendários
3. **Unificar tabelas de medidas** — Eliminar `01_Medidas` (antiga) e manter apenas `01_medidas`

### Prioridade Média
4. **Consolidar aceites mensais** — Criar uma única tabela `Base_aceites` com coluna de mês
5. **Remover calendário de resgates redundante** — Eliminar `dCalendar_resgate` ou `dCalendario_Resgates`

### Prioridade Baixa
6. **Documentar medidas DAX** — Exportar as 79 expressões DAX para documentação
7. **Revisar relacionamentos de CPF** — Verificar se a normalização de zeros está funcionando corretamente

---

*Documento gerado em 2026-06-08 após análise completa do arquivo Dashboard_TOP_V7.pbix*
*Analista virtual: Claude Code*
