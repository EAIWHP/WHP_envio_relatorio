# Programa Mais Top - Documentação do Projeto

**Projeto:** Whirlpool - Programa + TOP  
**Caminho:** `/home/thamiresvieira/projetos/Programa_mais_top`  
**Ferramenta:** Microsoft Power BI  
**Última atualização documentação:** 08/06/2026 (v1.3)  
**Data registro:** 2026-06-03  

**Responsáveis:**
- Técnico: Matheus Carmo; Thamires Vieira
- Negócio: Fauzi Naufal; Matheus Carmo
- Stakeholders: Rhuscaya; Laís Calazans; Andressa Borini

---

## ⚠️ IMPORTANTE - Regras de Negócio

Este projeto possui **regras de negócio específicas** que devem ser seguidas rigorosamente:

### Regras Fundamentais

| Regra | Descrição |
|-------|-----------|
| **CPF/CNPJ** | **SEMPRE tratar como TEXTO** — preservar zeros à esquerda |
| **Último mês** | Mês da última data disponível na base de vendas (não usa mês atual do calendário) |
| **Mês anterior** | Mês imediatamente anterior ao último mês disponível |
| **YoY** | Comparação último mês com mesmo mês do ano anterior |
| **Participante ativo** | Cadastro com status = "Ativo" |
| **Elegibilidade** | Ativo + aceite válido + regras de treinamento quando aplicável |
| **Treinamento concluído** | CPF aparece na Base_treinamentos. **CPF ausente = "Não fez treinamento"** |
| **Aceite** | Preferencialmente validado pela Base_aceites. Campo `data de aceite` da 02_Cadastro ainda não está completamente validado |
| **SUPERTOP** | Produto marcado como SUPERTOP na base de vendas. **Campanha sazonal — nem todos os meses têm produtos SUPERTOP** (depende do orçamento/investimento da Whirlpool no mês). Ex: ABR e MAI/2026 não tiveram SUPERTOP |
| **BLACKLIST (Campanhas Extras)** | Revendas na Blacklist participam da **campanha base** mas **NÃO participam das campanhas extras/sazonais**. Motivos principais: E-COMMERCE, Televendas, Vendas Atacado. Atualizada mensalmente pelo time comercial |
| **Setembro/2025** | ❌ **BASE INEXISTENTE — cliente nunca enviou, nunca será recebida.** Lacuna permanente. NUNCA interpolar ou estimar |

### Regras do Regulamento — Programa +TOP

As regras abaixo foram levantadas na análise cruzada regulamento vs. dashboard (2026-06-03) e **confirmadas para documentação em 2026-06-08**.

#### Status de Cadastro

| Status | Regra do Regulamento | Como está no Dashboard |
|--------|---------------------|----------------------|
| **Ativo** | Cadastro confirmado + aceite válido | ✅ Tratado como "Ativo" |
| **Inativo** | Após 3 meses sem acesso ou pré-cadastro sem ativação em 90 dias | ✅ Tratado como "Inativo" |
| **Somente Catálogo** | Após 2 meses sem acesso — não acumula, só resgata | ⚠️ **Mapeado para "Inativo"** (sistema simplifica) |
| **Pré-cadastrado** | Até 90 dias para ativar, depois vira "Inativo" | ⚠️ Regra documentada mas transição não é automática no dashboard |

#### Pontuações por Cargo

| Cargo | Pontuação de Vendas | Pontuação Adicional | No Dashboard? |
|-------|--------------------|---------------------|---------------|
| **Vendedor** | Pontos por produtos vendidos | 10 pts aniversário | ✅ Vendas / ❌ Aniversário |
| **Gerente** | 10% da pontuação da equipe de vendedores | 10 pts aniversário | ❌ 10% / ❌ Aniversário |
| **Gerente Regional** | 5% da pontuação dos vendedores das lojas atendidas | 10 pts aniversário | ❌ 5% / ❌ Aniversário |
| **Gestor da Informação** | 300 pts por envio correto no prazo | 300 pts bônus trimestral (mar/jun/set/dez) + 10 pts aniversário | ❌ Tudo |

> ⚠️ **Importante:** As pontuações de aniversário, GI, Gerente e Gerente Regional **não estão no dashboard**. São creditadas apenas na apuração final da Whirlpool.

#### Valor e Expiração dos Pontos

| Regra | Detalhe |
|-------|---------|
| **Valor do ponto** | R$ 1,00 |
| **Validade** | 12 meses após a data de crédito |
| **Expiração automática** | Todo 2º dia útil de cada mês |

#### Formas de Resgate

Os pontos podem ser trocados por:
- Transferência bancária
- PIX
- Cartão
- Lojas virtuais
- Pagamento de contas
- Recargas de celular
- Vouchers

#### Prazos Importantes

| Prazo | Regra |
|-------|-------|
| **Resgate após desligamento** | 90 dias após inativação da revenda |
| **Contestação de pontuação** | 7 dias corridos a partir da data de crédito |
| **Envio de dados (Gestor da Informação)** | Até o 5º dia útil do mês subsequente *(apuração do time de engenharia de dados)* |

---

### Fontes de Dados e Frequências

| Base | Fonte | Frequência | Armazenamento |
|------|-------|------------|--------------|
| Vendas | Rodrigo Moraes | Mensal (5º dia útil) | SharePoint |
| Cadastros | Plataforma + TOP | Diária | SharePoint |
| Acessos | Plataforma + TOP | Diária | SharePoint |
| Aceites | Rodrigo Moraes | Mensal | SharePoint |
| Pontuação | Rodrigo/Cauê Webster | Mensal | SharePoint |
| Resgates | Cauê Webster | Mensal | SharePoint |
| Chamados | Andressa Borini | Semanal (segunda) | SharePoint |
| SUPERTOP | Rodrigo Moraes | Mensal | SharePoint |
| Potencial | Matheus/Rodrigo | Mensal | SharePoint |
| Treinamentos | Plataforma + TOP | Diária | SharePoint |
| Enquete | Plataforma + TOP | Diária | SharePoint |

---

## Estrutura de Pastas

```
Programa_mais_top/
├── bases/
│   ├── Old/                    → Backups e versões antigas
│   ├── Tabelas/                → Bases principais do sistema
│   │   ├── ARQUIVOS_AUXILIARES/
│   │   ├── BASES ABR/
│   │   ├── Bases_antigas/
│   │   │   ├── 01_JAN_26/
│   │   │   ├── 02_FEV_26/
│   │   │   ├── 03_MAR_26/
│   │   │   ├── 04_ABR_26/
│   │   │   └── 2025/
│   │   ├── Bases_processamento_mensal/
│   │   ├── Relatorios pontuais/
│   │   │   └── Validacao cadastros ABR/
│   │   └── bases_potencial/
│   └── Vendas Processadas/     → Dados mensais (2024-01 a 2026-04)
├── bases_recorrentes/          → Bases baixadas automaticamente da plataforma
│   └── 030626/
└── power_bi/                   → Dashboards Power BI
    └── Dashboard_TOP_V7.pbix
```

---

## Bases Principais

| Base | Arquivo | Frequência |
|------|---------|------------|
| Vendas | `bases/Vendas Processadas/AAAA_MM.xlsx` | Mensal (jan/2024 a abr/2026) |
| Cadastro | `bases/Tabelas/cadastro.xlsx` | Recorrente |
| Acessos | `bases/Tabelas/acessos.xlsx` | Recorrente |
| Pontos por CPF | `bases/Tabelas/pontos_por_cpf_.xlsx` | Atualizado |
| Resgates | `bases/Tabelas/resgates_detalhado.xlsx` | Atualizado |
| Chamados | `bases/Tabelas/Chamados_catalogo.xlsx` | Atualizado |
| Aceite | `bases/Tabelas/WHP_Aceite_Mensal_*.xlsx` | Mensal |
| Treinamentos | `bases/Tabelas/Base_treinamentos.xlsx` | Atualizado |
| Aptos | `bases/Tabelas/Aptos_*.xlsx` | Mensal |
| Liderança Aptos | `bases/Tabelas/Lideranca_Aptos_*.xlsx` | Mensal |
| Enquete | `bases/Tabelas/enquete.xlsx` | Eventual |
| Fale Conosco | `bases/Tabelas/Fale_Conosco.xlsx` | Atualizado |
| Ocorrências | `bases/Tabelas/OcorrênciasTOPSAC.xlsx` | Atualizado |
| Distribuidor x PDV | `bases/Tabelas/WHP_Distribuidor_x_PDV.xlsx` | Estático |
| Cadastro Revenda | `bases/Tabelas/WHP_PROD_Cadastro_Revenda_*.xlsx` | Estático |

---

## Modelo de Dados - Relacionamentos

```
06_Base_Vendas ----SKU----> dProduto
06_Base_Vendas ----CPF----> 02_Cadastro
06_Base_Vendas ----Data----> dCalendario
02_Cadastro ----grupo----> tab_nome_revenda
02_Cadastro ----CPF----> Base_treinamentos
02_Cadastro ----CPF----> Base_aceites
02_Cadastro ----CPF----> Tab_resgates_detalhado
02_Cadastro ----CPF----> tab_pontos_por_cpf
02_Cadastro ----CPF----> 03_Acessos
02_Cadastro ----CPF----> enquete
02_Cadastro ----CPF----> OcorrênciasTOPSAC
bases_potencial ----revenda----> tab_nome_revenda
bases_potencial ----Data----> dCalendario
```

**Chave principal:** CPF/CNPJ (sempre como texto)

---

## KPIs Principais

### Vendas
- Vendas Totais (último mês)
- Vendas SUPERTOP (último mês)
- Usuários com venda (DISTINCTCOUNT CPF)
- TOP SKU (ranking TOP 1)
- Variação % último mês vs anterior
- Variação YoY %

### Cadastros/Acessos
- Acessos únicos (DISTINCTCOUNT CPF/CNPJ)
- Navegação (COUNTROWS)
- Taxa de ativação (Ativos / Total)
- Taxa de acesso ativos (Acessaram / Ativos)

### Pontuação
- Total créditos (SUM desde jan/2024)
- Total resgatados
- Pontos expirados
- Saldo atual

### Resgates
- Total de pedidos
- Participantes distintos
- Média valor resgatado

### SUPERTOP
- Usuários com venda SUPERTOP
- Total Pontos SUPERTOP

### Potencial
- Potencial de pontos
- Aproveitamento %

### Treinamentos
- Cadastros ativos
- Ativos concluintes
- Taxa de conclusão %

### Enquete
- Taxa de excelente %
- Taxa de oportunidade (Ruim) %

---

## Páginas do Dashboard

| Página | Foco |
|--------|------|
| Performance Vendas | Vendas, pontuação, variações, análise por revenda/regional/SKU/categoria |
| Cadastros-Acessos | Participantes, status, acessos, navegação, aceite |
| Resumo-Pontos | Créditos, resgates, expirados, saldo |
| Pontuação Individual | Detalhe por participante (CPF/nome) |
| Resgates Detalhado | Pedidos, participante, CPF, valor, data, produto |
| Chamados Geral | Total, status, motivos, submotivos |
| SUPERTOP | Vendas e pontos SUPERTOP |
| Potencial | Potencial de pontuação e aproveitamento |
| Treinamentos | Obrigatórios, concluintes, taxa de conclusão |
| Enquete | Avaliação da plataforma |

---

## Processo de Atualização

1. **Recebimento** — Conferir fontes no SharePoint/e-mail
2. **Validação estrutural** — Colunas, tipos, CPF/CNPJ
3. **Processamento Python** — Pipeline validação, pontuação, consolidação (vendas)
4. **Atualização SharePoint** — Substituir mantendo nomenclatura
5. **Atualização PBI** — Atualizar consultas, validar erros
6. **Validação KPIs** — Comparar totais com origem
7. **Publicação** — Publicar no workspace/app

---

## Validações Críticas

- [ ] Vendas: comparar soma Vendas e Pontuação com consolidado
- [ ] Cadastros: validar CPFs distintos, status Ativo, sem duplicidades
- [ ] Acessos: comparar com relatório da plataforma
- [ ] Aceites: CPFs com aceite, datas, CPF com 11 dígitos
- [ ] Pontuação: créditos, resgates, expirados, saldo
- [ ] Resgates: total pedidos, soma valor, participantes distintos
- [ ] Chamados: total, status, datas abertura/encerramento
- [ ] Treinamentos: CPFs concluintes, cursos obrigatórios, progresso 100%
- [ ] Enquete: total respostas, distribuição Excelente/Ótimo/Bom/Regular/Ruim
- [ ] Filtros: testar Revenda, Regional, SKU, Categoria, CNPJ, Período, Status, Cargo

---

## Problemas Comuns

| Problema | Ação |
|----------|------|
| Arquivo mensal não recebido | Registrar pendência, solicitar, documentar impacto |
| Coluna ausente/renomeada | Ajustar Power Query ou solicitar reenvio no layout padrão |
| CPF/CNPJ perdeu zeros | Garantir tipo texto, normalizar com apenas números |
| Erro de relacionamento | Verificar chaves, cardinalidade, direção, duplicidades |
| Indicador diferente da origem | Revisar filtros, período, status, aceite, duplicidades |
| Atualização falha PBI Service | Conferir credenciais, caminho SharePoint, permissões |
| Revenda nomes duplicados | Atualizar tabela DE/PARA tab_nome_revenda |
| Treinamentos não mostram quem não fez | Criar flag: cadastros ativos com aceite - CPFs ausentes = "Não fez" |

---

## Notas de Desenvolvimento

- Sempre preservar backups na pasta `bases/Old/` antes de sobrescrever
- CPF/CNPJ **sempre como texto** — nunca número
- **Data Venda Ajustada** = `data venda` após limpeza/formato data (feito pelo pipeline Python)
- **bases_recorrentes/** = downloads automáticos da plataforma; NÃO sobrescreve `bases/Tabelas/`; guarda histórico
- **Código Python de vendas** existe mas **não está em uso no momento**
- **Aceite mensal** (`Base_aceites`) = aceite na plataforma para estar apto. **Aceite cadastro** = aceite nos regulamentos da campanha base. Existem campanha base + extras (sazonais)
- Base de setembro/2025 não existe — marcar lacuna em análises
- Último mês é dinâmico (última data na base), não o mês atual do calendário
- Tabela DE/PARA `tab_nome_revenda` deve ser mantida atualizada
- Variáveis DAX usam `Data Venda Ajustada` (não `data venda` diretamente)

---

## 🛠️ Ajustes no Power Query (Documentados em 2026-06-08)

### CPF/CNPJ — Converter para Texto
**Status:** ⚠️ Pendente de aplicação no Power Query

| Base | Campo CPF | Tipo Atual | Ação no Power Query |
|------|-----------|-----------|---------------------|
| `acessos.xlsx` | `Cpf/Cnpj` | Número | `Text.From([Cpf/Cnpj])` + `Text.PadStart(..., 11, "0")` |
| `pontos_por_cpf_.xlsx` | `CPFCNPJ` | Número | `Text.From([CPFCNPJ])` + `Text.PadStart(..., 11, "0")` |
| `resgates_detalhado.xlsx` | `CPF` | Número | `Text.From([CPF])` + `Text.PadStart(..., 11, "0")` |
| `Base_treinamentos.xlsx` | `CPF` | Número | `Text.From([CPF])` + `Text.PadStart(..., 11, "0")` |
| `Vendas Processadas/*.xlsx` | `CPF` | Número | `Text.From([CPF])` + `Text.PadStart(..., 11, "0")` |
| `enquete.xlsx` | `cpf` | Número | `Text.From([cpf])` + `Text.PadStart(..., 11, "0")` |
| `WHP_Aceite_Mensal_*.xlsx` | `CPF` | Número | `Text.From([CPF])` + `Text.PadStart(..., 11, "0")` |
| `Cadastros_revenda_PDV` | `Cpf` | Número | `Text.From([Cpf])` + `Text.PadStart(..., 11, "0")` |

> **Raciocínio:** O cadastro (`02_Cadastro[cpf/cnpj]`) já é texto e preserva zeros. Todas as outras bases importam CPF como número, perdendo zeros à esquerda. Isso quebra relacionamentos no modelo para CPFs como `00123456789` (vindo como `123456789` nas outras tabelas).
>
> **Fórmula Power Query sugerida:**
> ```powerquery
> = Table.TransformColumns(Source, {{"CPF", each Text.PadStart(Text.From(_), 11, "0"), type text}})
> ```

### Calendários
> **Decisão do usuário (2026-06-08):** NÃO consolidar calendários. Manter as 6 tabelas de calendário como estão. Não excluir nenhuma.

### Tabelas de Medidas
> **Decisão do usuário (2026-06-08):** NÃO excluir nenhuma tabela de medidas. Aguardar revisão do usuário para decidir sobre `01_Medidas` vs `01_medidas`.

### Aceites Mensais
> **Decisão do usuário (2026-06-08):** Manter aceites mensais em abas/planilhas separadas. Criar histórico com todos os meses. Não consolidar em arquivo único.

### BLACKLIST
> **Decisão do usuário (2026-06-08):** BLACKLIST (`BLACKLIST 110526.xlsx`) é apenas para documentação — **não integrar ao modelo PBI**. O dashboard não utiliza essa informação.

---

## ⚠️ Análise do Regulamento (2026-06-03 / Atualizado 2026-06-08)

O regulamento oficial do Programa +TOP foi analisado e cruzado com a documentação técnica do dashboard. **Foram encontradas 27 inconsistências**.

### Decisões do Usuário (2026-06-08)

- **13 regras** foram documentadas na seção "Regras do Regulamento — Programa +TOP" acima
- **14 regras** foram descartadas ou já existem no dashboard
- Documento completo de decisões: `DECISOES_INCONSISTENCIAS_REGULAMENTO.md`

### 🚨 Descoberta Crítica (mantida)

> **O dashboard NÃO reflete a pontuação completa que a Whirlpool utiliza.**
>
> A pontuação real (usada para pagamento/resgate) inclui:
> - Vendas dos participantes ✅ (está no dashboard)
> - Pontuação de aniversário (10 pts) ❌ **NÃO está no dashboard**
> - Pontuação do Gestor da Informação (300 pts) ❌ **NÃO está no dashboard**
> - Pontuação do Gerente Regional (5%) ❌ **NÃO está no dashboard**
> - Pontuação do Gerente (10%) ❌ **NÃO está no dashboard**
>
> O dashboard serve para **acompanhamento operacional**, mas **NÃO** representa a pontuação final de pagamento.

---

*Documento consolidado do arquivo: Documentacao_Dashboard_TOP_Formatada_Modelo_Profissional.docx*  
*Mantido pelo Claude Code. Atualizado em cada sessão de trabalho.*
