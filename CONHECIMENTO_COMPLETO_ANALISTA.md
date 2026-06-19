# 🧠 Conhecimento Completo — Programa +TOP
## Análise Aprofundada das Bases e Regras de Negócio
**Data da análise:** 2026-06-08  
**Atualizado em:** 2026-06-08 (correções após revisão do usuário)  
**Metodologia:** Leitura estrutural de 17 bases + cruzamentos + validação contra regras de negócio

---

# PARTE 1: MAPEAMENTO DAS BASES

## 1.1 Resumo Geral das Bases

| # | Base | Registros | Colunas | Chave Primária | Frequência |
|---|------|-----------|---------|---------------|------------|
| 1 | **cadastro** | 21,144 | 19 | cpf/cnpj | Diária (plataforma) |
| 2 | **acessos** | 749,858 | 8 | - (log) | Diária (plataforma) |
| 3 | **pontos_por_cpf** | 27,388 | 6 | CPFCNPJ | Mensal |
| 4 | **resgates_detalhado** | 142,882 | 6 | Nº RESGATE | Mensal |
| 5 | **treinamentos** | 65,013 | 13 | - (CPF + Curso) | Diária (plataforma) |
| 6 | **aceites** | 7 abas (out/2025 a abr/2026) | 4-5 | CPF + mês | Mensal |
| 7 | **chamados** | 19,388 | 10 | NÚMERO | Semanal |
| 8 | **fale_conosco** | 5,035 | 120* | - | Eventual |
| 9 | **enquete** | 1,454 | 8 | - | Eventual |
| 10 | **ocorrências** | 7,240 | 39 | N.º | Eventual |
| 11 | **tab_nome_revendas** | 82 | 2 | REVENDAS_NAME_RAW | Estático |
| 12 | **distribuidor_pdv** | 3,039 | 4 | - | Estático |
| 13 | **cadastro_revenda** | 20,558 | 12 | Cpf | Estático |
| 14 | **aptos** | 64,296 | 13 | - | Mensal |
| 15 | **lideranca_aptos** | 1,089 | 8 | - | Mensal |
| 16 | **vendas** (amostra 2026_04) | 61,152 | 14 | - | Mensal (27 arquivos) |
| 17 | **base_tratada** | 7,667 | 9 | - | - |
| 18 | **blacklist** | 48 | 5 | REVENDA / CNPJ | Mensal (atualização comercial) |

*Fale Conosco tem 120 colunas mas apenas 13 preenchidas (resto é lixo de importação)

---

## 1.2 Detalhamento Base por Base

### 📋 1. CADASTRO (`cadastro.xlsx`)

**Propósito:** Cadastro mestre de todos os participantes do programa +TOP  
**Fonte:** Plataforma +TOP (download automático)  
**Chave:** `cpf/cnpj` (sempre como TEXTO, com zeros à esquerda)  
**Volume:** 21,144 participantes

**Colunas (19):**
| Coluna | Tipo | Descrição | Regra de Negócio |
|--------|------|-----------|-----------------|
| `nome` | object | Nome completo do participante | - |
| `cpf/cnpj` | str | CPF/CNPJ (texto com zeros) | **SEMPRE texto** — zeros à esquerda |
| `matricula` | str | Matrícula na revenda | - |
| `cargo` | str | Cargo no programa | 6 cargos possíveis |
| `status` | str | Status no programa | 7 status possíveis |
| `regional` | str | Regional comercial Whirlpool | 6 regionais |
| `grupo` | str | Nome da revenda (cluster) | 48 grupos |
| `cnpj grupo` | str | CNPJ da revenda matriz | - |
| `rua` | object | Endereço | - |
| `bairro` | object | Bairro | - |
| `cidade` | object | Cidade | - |
| `uf` | object | Estado | - |
| `cep` | float | CEP | - |
| `telefone` | float | Telefone fixo | - |
| `celular` | float | Celular | - |
| `email` | str | E-mail | - |
| `data de aceite` | datetime | Data do aceite do regulamento | Usada para elegibilidade |
| `LGPD` | datetime | Data de consentimento LGPD | - |
| `data atualização` | object | Data da última atualização | - |

**Distribuição de Status:**
| Status | Quantidade | % |
|--------|-----------|---|
| Ativo | 13,375 | 63.3% |
| Indicado Moderado | 3,612 | 17.1% |
| Pré-Cadastrado | 2,614 | 12.4% |
| Reprovado | 1,126 | 5.3% |
| Inativo | 411 | 1.9% |
| Em Moderação | 5 | 0.0% |
| Aguardando Aprovação | 1 | 0.0% |

**Distribuição de Cargos:**
| Cargo | Quantidade | % |
|-------|-----------|---|
| Vendedor | 17,148 | 81.1% |
| Gerente de Loja | 3,604 | 17.0% |
| Gerente Regional | 185 | 0.9% |
| Embaixador | 113 | 0.5% |
| Gestor da Informação | 63 | 0.3% |
| Master | 31 | 0.1% |

**🔴 PROBLEMAS IDENTIFICADOS:**
- Data de aceite nula em 7,490 registros (35.4%) — **ESTE É O PROBLEMA**
- A regra de elegibilidade diz "Ativo + aceite válido", mas 35% não têm data de aceite
- Cadastro_Revenda (atualização/complemento): 12,066 CPFs do cadastro não aparecem nesta base — possíveis razões: sem loja associada, vínculo desatualizado, ou status diferente
- 6 registros sem regional/grupo

**🔗 CONEXÕES:**
- CPF → Vendas, Acessos, Pontos, Resgates, Treinamentos, Aceites, Aptos, Enquete
- Grupo → tab_nome_revendas (DE/PARA de nomes)
- Status → Regra de elegibilidade (Ativo = elegível)
- Cargo → Regras de pontuação diferenciadas

---

### 📋 2. ACESSOS (`acessos.xlsx`)

**Propósito:** Log de todos os acessos à plataforma +TOP  
**Fonte:** Plataforma +TOP (download automático)  
**Volume:** 749,858 registros  
**Período:** 19/10/2025 a 03/06/2026

**Colunas (8):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `Data Acesso` | datetime | Data/hora do acesso |
| `Regional` | str | Regional comercial |
| `Grupo Cluster` | str | Nome da revenda (48 grupos) |
| `Nome` | object | Nome do participante |
| `Cpf/Cnpj` | int64 | CPF/CNPJ (⚠️ como número — perde zeros!) |
| `Cargo` | str | Cargo no programa |
| `Status` | str | Status no programa |
| `Pagina` | str | Página acessada |

**Páginas mais acessadas (top 10):**
| Página | Acessos | % |
|--------|---------|---|
| Home | 412,263 | 55.0% |
| Resgatar | 85,856 | 11.5% |
| Campanhas | 59,467 | 7.9% |
| Treinamentos | 57,033 | 7.6% |
| Campanha Extra | 47,180 | 6.3% |
| Cadastro | 29,847 | 4.0% |
| Meu Desempenho | 20,313 | 2.7% |
| Sobre a campanha | 10,715 | 1.4% |
| Produtos participantes | 9,618 | 1.3% |
| Indicadores Detalhes | 5,388 | 0.7% |

**🔴 PROBLEMAS IDENTIFICADOS:**
- CPF como int64 → perde zeros à esquerda → dificulta join com cadastro
- Apenas 5,933 CPFs em comum com cadastro (de 21,144)
- 8,097 CPFs em acessos que NÃO estão no cadastro → dados de participantes que podem ter saído
- Período começa em 19/10/2025 — exatamente a data de início do programa atual!

**🔗 CONEXÕES:**
- CPF/Cnpj (com zeros perdidos) → Cadastro (tratar como texto)
- Data Acesso → Último mês dinâmico do dashboard
- Pagina → KPI de navegação
- Status → Taxa de acesso de ativos

---

### 📋 3. PONTOS_POR_CPF (`pontos_por_cpf_.xlsx`)

**Propósito:** Consolidado de pontos por participante  
**Fonte:** Time de dados (Rodrigo/Cauê)  
**Volume:** 27,388 CPFs  
**Abas:** 6 (CONSOLIDADO_RESGATES, PONTOS POR CPF, VALES, LOJA VIRTUAL, PAGAMENTO DE CONTAS, RECARGA DE CELULAR)

**Aba "PONTOS POR CPF" (principal):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `Nome` | object | Nome do participante |
| `CPFCNPJ` | int64 | CPF (⚠️ como número) |
| `Créditos` | float | Total de pontos creditados |
| `Resgates` | float | Total de pontos resgatados |
| `EXPIRADO` | float | Pontos expirados |
| `Saldo Atual` | float | Saldo disponível |

**Distribuição de Saldo:**
| Faixa de Saldo | Quantidade |
|----------------|-----------|
| Saldo = 0 | 9,178 |
| Saldo > 0 | 18,210 |
| Saldo < 0 | 0 |

**🔴 PROBLEMAS IDENTIFICADOS:**
- Apenas 6,759 CPFs em comum com cadastro (de 21,144)
- 20,629 CPFs com pontos que NÃO estão no cadastro atual → dados de edição anterior!
- Confirma que dados de pontos misturam edições do programa
- CPF como número → perde zeros

**🔗 CONEXÕES:**
- CPFCNPJ → Cadastro (com normalização)
- Créditos/Resgates/Expirado → KPIs de pontuação
- Saldo Atual → Disponibilidade para resgate
- Outras abas → Formas de resgate detalhadas

---

### 📋 4. RESGATES_DETALHADO (`resgates_detalhado.xlsx`)

**Propósito:** Detalhamento de todos os resgates realizados  
**Fonte:** Cauê Webster  
**Volume:** 142,882 resgates  
**Período:** 05/03/2024 a 30/04/2026

**Colunas (6):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `Nº RESGATE` | int64 | ID do resgate |
| `PARTICIPANTE` | object | Nome |
| `CPF` | int64 | CPF (⚠️ como número) |
| `VALOR` | float | Valor em pontos (R$ 1,00 = 1 ponto) |
| `DATA RESGATE` | datetime | Data do resgate |
| `PRODUTO` | str | Produto resgatado |

**Top 10 produtos resgatados:**
| Produto | Resgates |
|---------|---------|
| Pix | 96,949 |
| Transferência | 20,341 |
| Vale Presente | 14,170 |
| Recarga | 6,128 |
| Conta de Luz | 2,855 |
| Loja | 1,580 |
| Pagamento de Contas | 469 |
| Vales | 147 |
| Outros | variados |

**🔴 PROBLEMAS IDENTIFICADOS:**
- Apenas 5,455 CPFs em comum com cadastro
- 13,259 CPFs com resgates que NÃO estão no cadastro → edição anterior
- Período começa em mar/2024 (edição anterior)

**🔗 CONEXÕES:**
- CPF → Cadastro
- VALOR → R$ 1,00 por ponto (regra do regulamento)
- DATA RESGATE → Prazo de expiração (12 meses)
- PRODUTO → Formas de resgate documentadas

---

### 📋 5. TREINAMENTOS (`Base_treinamentos.xlsx`)

**Propósito:** Registro de conclusão de treinamentos obrigatórios  
**Fonte:** Plataforma +TOP  
**Volume:** 65,013 registros  
**Chave:** CPF + Curso

**Colunas (13):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `Participante` | object | Nome |
| `CPF` | int64 | CPF (⚠️ como número) |
| `UsuarioZoomId` | int64 | ID na plataforma |
| `Trilha` | str | Trilha de treinamento |
| `Curso` | str | Nome do curso (137 cursos únicos) |
| `Inicio` | datetime | Data de início |
| `Conclusão` | datetime | Data de conclusão |
| `Progresso` | int64 | % de progresso |
| `Estado` | str | Status do treinamento |
| `Distribuidor` | str | Nome do distribuidor |
| `Cnpj Distribuidor` | int64 | CNPJ do distribuidor |
| `PDV` | object | Nome do PDV |
| `Cnpj PDV` | float | CNPJ do PDV |

**ESTATÍSTICAS CRÍTICAS:**
- **100% dos registros têm Progresso = 100** → base só registra conclusões!
- Estado = "Concluido" (sem acento) em 100% dos casos
- 137 cursos únicos
- 5,642 CPFs únicos fizeram treinamentos

**🔴 PROBLEMAS IDENTIFICADOS:**
- Base só tem quem CONCLUIU — quem não concluiu NÃO aparece
- Para saber quem NÃO fez: Cadastro Ativos com Aceite - CPFs em treinamentos = "Não fez"
- Progresso sempre 100% → não dá para saber quem está "em andamento"
- Não há nota/avaliação → não é possível validar regra dos 80% de acertos
- Apenas 2,494 CPFs em comum com cadastro

**🔗 CONEXÕES:**
- CPF → Cadastro
- Progresso/Estado → Elegibilidade (treinamento obrigatório)
- Trilha/Curso → KPI de taxa de conclusão
- Distribuidor/PDV → Hierarquia revenda

---

### 📋 6. ACEITES (`WHP_Aceite_Mensal_*.xlsx`)

**Propósito:** Registro mensal de aceite à campanha na plataforma  
**Fonte:** Rodrigo Moraes / Time de dados  
**Abas:** 7 (OUT_2025, NOV_2025, DEZ_2025, JAN_2026, FEV_2026, MAR_2026, ABR_2026)

**Colunas por aba:**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `Nome` | object | Nome |
| `CPF` | int64 | CPF (⚠️ como número) |
| `Revenda` | str | Nome da revenda |
| `DataAceite` | str | Data do aceite (como texto!) |
| `DataNascimento` | float/str | Data de nascimento (a partir de MAR_2026) |

**Volume por mês:**
| Mês | Aceites | Revendas |
|-----|---------|----------|
| OUT_2025 | 4,889 | - |
| NOV_2025 | 7,687 | - |
| DEZ_2025 | 7,524 | - |
| JAN_2026 | 7,631 | - |
| FEV_2026 | 4,664 | - |
| MAR_2026 | 7,987 | - |
| ABR_2026 | 8,021 | 46 |

**🔴 PROBLEMAS IDENTIFICADOS:**
- DataAceite como TEXTO (não datetime) → pode ter formatos inconsistentes
- Apenas 3,346 CPFs em comum com cadastro (de 21,144)
- Flutuação grande entre meses (FEV_2026 com apenas 4,664 aceites)
- A partir de MAR_2026 inclui DataNascimento → pode ser usada para pontuação de aniversário

**🔗 CONEXÕES:**
- CPF → Cadastro
- DataAceite → Elegibilidade mensal
- Aceite = condição obrigatória para pontuação (regulamento)
- Diferença de aceites entre meses → sazonalidade ou mudança de regra

---

### 📋 7. VENDAS (`Vendas Processadas/AAAA_MM.xlsx`)

**Propósito:** Registro de vendas de produtos participantes  
**Fonte:** Rodrigo Moraes (time de dados)  
**Volume (amostra 2026_04):** 61,152 registros  
**Total arquivos:** 27 (jan/2024 a abr/2026)  
**Faltante:** 2025_09 (setembro/2025) — BASE INEXISTENTE

**Colunas (14):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `Regional` | float | Regional comercial |
| `Revenda` | str | Nome da revenda |
| `CNPJ` | float | CNPJ da revenda/loja |
| `CPF` | int64 | CPF do vendedor (⚠️ como número) |
| `Nome` | str | Nome do vendedor |
| `CATEGORIA` | float | Categoria do produto |
| `SUBCATEGORIA` | float | Subcategoria |
| `SKU + produto` | str | Descrição completa do produto |
| `data venda` | datetime | Data da venda |
| `Vendas` | float | Quantidade vendida |
| `Pontuação` | float | Pontos gerados |
| `Mês` | str | Mês/Ano da venda |
| `SUPERTOP` | float | Flag SUPERTOP |
| `SKU` | float | Código SKU |

**ANÁLISE CRÍTICA DAS VENDAS:**
- 7,820 registros (12.8%) têm pontuação = 0 mesmo com vendas > 0 → produtos que não pontuam?
- 9 registros com vendas = 0 → provavelmente registros de ajuste
- 30 registros com SKU + produto nulo → dados incompletos
- SUPERTOP: todos os valores são nulos no arquivo 2026_04! → **Explicação: nem todos os meses têm produtos SUPERTOP. É uma campanha sazonal que depende do orçamento/investimento da Whirlpool no mês**

**🔴 PROBLEMAS IDENTIFICADOS:**
- Apenas 2,365 CPFs em comum com cadastro (de 21,144)
- Isso significa que apenas ~11% dos cadastrados venderam em Abr/2026
- CPF como número → perde zeros
- SUPERTOP nulo em 2026_04 → **normal: não houve investimento em SUPERTOP neste mês (ABR/2026)**
- Dados de jan/2024 a set/2025 = edição anterior → regras diferentes

**🔗 CONEXÕES:**
- CPF → Cadastro (dados do vendedor)
- CNPJ → Distribuidor_PDV (hierarquia)
- SKU + produto → dProduto (dimensão produto)
- data venda → dCalendario (dimensão tempo)
- Pontuação → KPI de pontuação por vendas
- SUPERTOP → KPI específico de produtos SUPERTOP

---

### 📋 8. APTOS (`Aptos_Abril_2026_v4.xlsx`)

**Propósito:** Base de aptos à pontuação (vendedores elegíveis)  
**Fonte:** Time de dados / apuração mensal  
**Volume:** 64,296 registros

**Colunas (13):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `ano` | int64 | Ano da apuração |
| `mes` | str | Mês da apuração |
| `revenda` | str | Nome da revenda |
| `cnpj` | float | CNPJ da revenda |
| `participante` | str | Nome do participante |
| `cpf` | int64 | CPF (⚠️ como número) |
| `gi` | str | Nome do Gestor da Informação |
| `cpf_gi` | float | CPF do GI |
| `kpi` | str | Tipo de KPI (SITE ou GI) |
| `produto` | str | Produto vendido |
| `quantidade` | float | Quantidade |
| `pontos` | float | Pontos gerados |
| `treinamento` | str | Fez treinamento (sim/não) |

**ESTATÍSTICAS:**
- 6,058 CPFs únicos
- 33 revendas
- KPI: SITE (vendas diretas) ou GI (envio do gestor)
- Treinamento: sim/não
- Inclui GI e CPF_GI → permite ligar venda ao gestor responsável

**🔴 PROBLEMAS IDENTIFICADOS:**
- Apenas 2,546 CPFs em comum com cadastro
- 6 meses de dados (provavelmente)

**🔗 CONEXÕES:**
- CPF → Cadastro
- CPF_GI → Cadastro (para pontuação do GI)
- CNPJ → Distribuidor_PDV
- KPI → Tipo de pontuação
- Treinamento → Elegibilidade

---

### 📋 9. LIDERANÇA_APTOS (`Lideranca_Aptos_Abril_2026_v4.xlsx`)

**Propósito:** Pontuação de liderança (Gerentes e Gerentes Regionais)  
**Fonte:** Time de dados / apuração mensal  
**Volume:** 1,089 registros

**Colunas (8):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `ano` | int64 | Ano |
| `mes` | str | Mês |
| `revenda` | str | Revenda |
| `cnpj` | float | CNPJ |
| `participante` | str | Nome |
| `cpf` | int64 | CPF |
| `kpi` | str | Tipo de KPI |
| `pontos_final` | float | Pontos finais |

**🔗 CONEXÕES:**
- CPF → Cadastro (filtrar cargos de liderança)
- Pontos_final → 10% (Gerente) ou 5% (Gerente Regional) da equipe
- KPI → Tipo de pontuação de liderança

---

### 📋 10. CHAMADOS (`Chamados_catalogo.xlsx`)

**Propósito:** Registro de chamados ao suporte  
**Fonte:** Andressa Borini  
**Volume:** 19,388 chamados  
**Período:** 01/03/2024 a 31/03/2026

**Colunas (10):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `NÚMERO` | int64 | ID do chamado |
| `DATA` | datetime | Data de abertura |
| `DOCUMENTO` | object | CPF/CNPJ (como texto) |
| `PARTICIPANTE` | object | Nome |
| `CAMPANHA` | str | Campanha |
| `ASSUNTO` | str | Motivo do chamado |
| `Resolução` | datetime | Data de resolução |
| `SLA - META` | float | Meta de SLA (dias) |
| `SLA - DIA` | int64 | Dias para resolução |
| `SLA - HORAS` | int64 | Horas para resolução |

**Top 15 Assuntos:**
| Assunto | Quantidade |
|---------|-----------|
| Problemas no resgate | 3,618 |
| Problemas no Cadastro | 2,740 |
| Dúvidas Sobre a Campanha | 2,471 |
| Informações de pontuação | 1,828 |
| Problemas de acesso | 1,783 |
| Dúvidas Sobre Resgate | 1,489 |
| Dúvidas de Cadastro | 1,429 |
| Problemas de acesso ao Site | 851 |
| Alteração de Dados Cadastrais | 532 |
| Dúvidas Sobre Treinamentos | 456 |
| Dúvidas Sobre Pontuação | 436 |
| Exclusão de Cadastro | 333 |
| Solicitação de Recuperação de senha | 259 |
| Dúvidas Sobre Ranking | 225 |
| Inclusão de Vendas | 124 |

**MÉDIA DE SLA:** 3.3 dias

**🔗 CONEXÕES:**
- DOCUMENTO → Cadastro (CPF do participante)
- Assunto → Categorias de problemas da plataforma
- SLA → KPI de qualidade do atendimento

---

### 📋 11. OCORRÊNCIAS (`OcorrênciasTOPSAC.xlsx`)

**Propósito:** Ocorrências/pendências de pontuação  
**Fonte:** Time de dados / SAC  
**Volume:** 7,240 registros

**Colunas relevantes (de 39):**
| Coluna | Descrição |
|--------|-----------|
| `N.º` | ID da ocorrência |
| `Data/Hora abertura` | Data |
| `Cliente` | Nome do participante |
| `+TOP - CARGO` | Cargo |
| `CPF / CNPJ sem pontuação` | CPF |
| `+TOP - SOLICITAÇÃO` | Tipo de solicitação |
| `+TOP - MOTIVO` | Motivo |
| `+TOP - SUBMOTIVO` | Submotivo |
| `+TOP - NOME DA REDE` | Revenda |
| `Status (sem tempo decorrido)` | Status |

**Solicitações principais:**
| Solicitação | Quantidade |
|-------------|-----------|
| Alteração/Inclusão de Cadastro | 2,242 |
| Contestação de Pontuação | 1,466 |
| Alteração de Dados Cadastrais | 1,132 |
| Consulta | 651 |
| Exclusão de Cadastro | 431 |

**Status:**
| Status | Quantidade |
|--------|-----------|
| Em atendimento | ~50% |
| Encerrado | ~50% |

**🔗 CONEXÕES:**
- CPF → Cadastro, Vendas, Pontos
- Status → Processo de contestação (prazo: 7 dias corridos)
- Motivo → Categorias de problema

---

### 📋 12. ENQUETE (`enquete.xlsx`)

**Propósito:** Pesquisa de satisfação com a plataforma  
**Fonte:** Plataforma +TOP  
**Volume:** 1,454 respostas

**Colunas (8):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `nome` | str | Nome do participante |
| `participante` | str | Nome (repetido) |
| `cpf` | int64 | CPF |
| `cargo` | str | Cargo |
| `tipo` | str | Tipo ("Enquete") |
| `pergunta` | str | Pergunta |
| `resposta` | str | Resposta |
| `dataInclusao` | datetime | Data da resposta |

**Pergunta:** "O que você está achando do novo site do +TOP?"

**Respostas possíveis:**
- Excelente
- Ótimo
- Bom
- Regular
- Ruim

**🔗 CONEXÕES:**
- CPF → Cadastro
- Resposta → KPI: Taxa de Excelente % / Taxa de Oportunidade (Ruim) %

---

### 📋 13. TABELAS AUXILIARES

#### tab_nome_revendas (`tab_nome_revendas.xlsx`)
- 82 registros, 2 colunas
- **REVENDAS_NAME_RAW** → nome original (vem da base)
- **REVENDA_TRATADO** → nome padronizado
- Função: DE/PARA para unificar nomes de revendas que vêm com grafias diferentes

#### distribuidor_pdv (`WHP_Distribuidor_x_PDV.xlsx`)
- 3,039 registros, 4 colunas
- CNPJ_REVENDA, NOME_REVENDA, CNPJ_LOJA, NOME_LOJA
- Função: Hierarquia CNPJ revenda → CNPJ loja (PDV)

#### cadastro_revenda (`WHP_PROD_Cadastro_Revenda_x_Loja_x_Regional.xlsx`)
- 20,558 registros, 12 colunas
- **Propósito:** Atualização e complemento da base principal de cadastro
- Adiciona informações mais atualizadas de:
  - Quais e quantas lojas fazem parte de cada revenda
  - CNPJs das revendas e lojas
  - Endereços, emails e outros dados de contato
- Colunas: Cpf, Nome, Cargo, Celular, Email, Status, CNPJ_Revenda, Revenda, Regional_da_Revenda, CNPJ_Loja, Loja, Regional_da_Loja
- **DIFERENÇA para cadastro.xlsx:** Esta base vem da plataforma com vínculo revenda/loja + dados atualizados
- Nem todos os CPFs do cadastro principal aparecem aqui (12,066 CPFs do cadastro não estão nesta base) — possíveis razões: participantes sem loja associada, vínculo desatualizado, ou status diferente

---

### 📋 14. BLACKLIST (`BLACKLIST 110526.xlsx`)

**Propósito:** Lista de revendas que participam da **campanha base** mas **NÃO participam das campanhas extras/sazonais**  
**Fonte:** Time comercial Whirlpool (atualização manual)  
**Volume:** 48 registros, 34 revendas únicas  
**Atualização:** Abril/2026 (dados atualizados entre 15/04 e 22/04)  
**Chave:** REVENDA / CNPJ

**Colunas (5):**
| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `REVENDA` | str | Nome da revenda (34 nomes únicos) |
| `CNPJ` | str | CNPJ da revenda (formato variado — com pontuação, sem pontuação, múltiplos CNPJs por linha) |
| `PARTICIPANTES DO E-COMMERCE` | str | Motivo da exclusão das campanhas extras |
| `Data da ultima Atualização` | str | Data da última atualização da informação |
| `Unnamed: 4` | str | Observações adicionais (esporádico) |

**Motivos de exclusão das campanhas extras:**
| Motivo | Quantidade | Descrição |
|--------|-----------|-----------|
| **E-COMMERCE** | 8 | Vendas online não entram nas campanhas extras |
| **e-commerce + televendas não fazem parte** | 3 | Canal online + televendas excluídos |
| **Televendas + e-commerce** | 3 | Canal televendas + online excluídos |
| **NÃO TEM** | 3 | Revenda não possui esse canal |
| **Tele-vendas** | 2 | Canal televendas excluído |
| **TELEVENDAS** | 2 | Canal televendas excluído |
| **Televendas + e-commerce não fazem parte** | 2 | Ambos os canais excluídos |
| **Vendem Atacado** | 2 | Vendas no atacado excluídas |
| **Vendas Online** | 1 | Vendas online excluídas |
| **Outros** | 4 | Diversos (não trabalham com e-commerce, etc.) |

**CNPJs na Blacklist:**
- 18 CNPJs preenchidos (30 registros com "-" ou sem CNPJ)
- Formatos variados: com pontuação (`00.607.587/0026-50`), sem pontuação (`2869763000875`), múltiplos CNPJs por linha
- **⚠️ Padronização necessária:** CNPJs precisam ser normalizados para comparação (remover pontos, traços, espaços)

**Revendas que NÃO fazem parte do Programa (campanha base):**
| Revenda | Motivo |
|---------|--------|
| MAGAZAN | E-COMMERCE, não faz parte do Programa |
| DREBES E CIA LTDA | E-COMMERCE não faz parte do Programa |
| SOLAR MAGAZINE | E-commerce não faz parte do programa |

**🔗 CONEXÕES:**
- REVENDA / CNPJ → Cadastro (filtrar participantes dessas revendas)
- REVENDA / CNPJ → Vendas (excluir dessas revendas nas campanhas extras)
- CNPJ → Distribuidor_PDV (match por CNPJ normalizado)

**🔴 PROBLEMAS IDENTIFICADOS:**
- CNPJs em formatos inconsistentes (com/sem pontuação, múltiplos por célula)
- Algumas revendas listadas sem CNPJ (só nome)
- Atualização manual — risco de desatualização
- Diferenciação entre "não participa de campanhas extras" e "não faz parte do programa" não está explícita na estrutura

---

# PARTE 2: RELACIONAMENTOS ENTRE BASES

## 2.1 Diagrama de Relacionamentos

```
┌─────────────────┐     CPF      ┌──────────────────┐
│   02_Cadastro   │◄────────────►│ 06_Base_Vendas   │
│   (21,144)      │              │ (mensal)         │
└────────┬────────┘              └────────┬─────────┘
         │                                │
         │ CPF                    SKU     │
         │                                ▼
         │                       ┌──────────────────┐
         │                       │    dProduto      │
         ▼                       └──────────────────┘
┌─────────────────┐
│ Base_treinamentos│
│   (65,013)      │
└─────────────────┘
         ▲
         │ CPF
         │
┌────────┴────────┐              ┌──────────────────┐
│   02_Cadastro   │◄────────────►│ Base_aceites     │
│   (21,144)      │     CPF      │ (mensal/abas)    │
└────────┬────────┘              └──────────────────┘
         │
         │ CPF
         ▼
┌─────────────────┐              ┌──────────────────┐
│tab_pontos_por_  │◄────────────►│Tab_resgates_     │
│     cpf         │     CPF      │  detalhado       │
│   (27,388)      │              │  (142,882)       │
└─────────────────┘              └──────────────────┘
         ▲
         │ CPF
         │
┌────────┴────────┐              ┌──────────────────┐
│   02_Cadastro   │◄────────────►│  03_Acessos      │
│   (21,144)      │     CPF      │  (749,858)       │
└────────┬────────┘              └──────────────────┘
         │
         │ CPF
         ▼
┌─────────────────┐              ┌──────────────────┐
│    enquete      │              │ OcorrênciasTOPSAC│
│   (1,454)       │              │   (7,240)        │
└─────────────────┘              └──────────────────┘

┌─────────────────┐     grupo    ┌──────────────────┐
│   02_Cadastro   │◄────────────►│ tab_nome_revenda │
│   (21,144)      │              │    (82)          │
└─────────────────┘              └──────────────────┘

┌─────────────────┐     revenda  ┌──────────────────┐
│ bases_potencial │◄────────────►│ tab_nome_revenda │
│   (mensal)      │              │    (82)          │
└─────────────────┘              └──────────────────┘
```

## 2.2 Tabela de Cruzamentos (Integridade Referencial)

| Base A | Base B | Em Comum | Apenas em A | Apenas em B | Observação |
|--------|--------|----------|-------------|-------------|------------|
| Cadastro (21,144) | Acessos (14,030 CPFs) | 5,933 | 15,211 | 8,097 | Muitos CPFs em acessos não estão no cadastro atual |
| Cadastro (21,144) | Pontos (27,388 CPFs) | 6,759 | 14,385 | 20,629 | **Mais CPFs em pontos do que no cadastro! Edição anterior** |
| Cadastro (21,144) | Resgates (18,714 CPFs) | 5,455 | 15,689 | 13,259 | Resgates de edição anterior |
| Cadastro (21,144) | Treinamentos (5,642) | 2,494 | 18,650 | 3,148 | Apenas 12% dos cadastrados fizeram treinamento |
| Cadastro (21,144) | Aceites Abr/2026 (8,021) | 3,346 | 17,798 | 4,675 | Apenas 16% dos cadastrados aceitaram em Abr/2026 |
| Cadastro (21,144) | Vendas Abr/2026 (5,696) | 2,365 | 18,779 | 3,331 | Apenas 11% dos cadastrados venderam em Abr/2026 |
| Cadastro (21,144) | Aptos (6,058) | 2,546 | 18,598 | 3,512 | - |
| Cadastro (21,144) | Cadastro_Revenda (20,558) | 9,078 | 12,066 | 11,480 | Atualização/complemento — nem todos têm vínculo loja |
| Cadastro (21,144) | Enquete (1,448) | 629 | 20,515 | 819 | - |

## 2.3 Análise dos Cruzamentos

### 🚨 DESCoberta CRÍTICA 1: Cadastro_Revenda como Complemento (não base paralela)
A base `cadastro_revenda` é uma **atualização/complemento** da base principal `cadastro`:
- Adiciona informações mais atualizadas de lojas, CNPJs, endereços, emails
- Nem todos os CPFs do cadastro principal aparecem aqui (12,066 não estão)
- Possíveis razões: participantes sem loja associada, vínculo desatualizado, ou status diferente
- **Não são bases concorrentes — a cadastro_revenda enriquece a cadastro com vínculo revenda→loja**

### 🚨 DESCoberta CRÍTICA 2: Dados de Edição Anterior
Mais CPFs em pontos (27,388) e resgates (18,714) do que no cadastro atual (21,144):
- Pontos: 20,629 CPFs com pontos que NÃO estão no cadastro atual
- Resgates: 13,259 CPFs com resgates que NÃO estão no cadastro atual
- **Isso confirma que a base de pontos e resgates mistura dados da edição anterior (jan/2024 a out/2025) com a edição atual (out/2025 em diante)**

### 🚨 DESCoberta CRÍTICA 3: Taxa de Engajamento (Base Total vs Ativos)
**Base total (21,144 cadastrados):**
- 16% aceitaram em Abr/2026 | 11% venderam em Abr/2026 | 12% fizeram treinamento

**Apenas ATIVOS (13,375):**
- 24.8% aceitaram em Abr/2026 | 17.6% venderam em Abr/2026 | 18.1% fizeram treinamento

> ⚠️ **Importante:** Para vendas e aceites, consideramos apenas cadastros ATIVOS. A análise por status total é mantida para entender o funil completo (Pré-Cadastrado → Ativo → Aceite → Venda).

---

# PARTE 3: REGRAS DE NEGÓCIO vs BASES

## 3.1 Mapeamento Regra → Base → Campo

| # | Regra de Negócio | Base | Campo | Validação |
|---|-----------------|------|-------|-----------|
| 1 | CPF como texto | Todas | cpf/cnpj | ⚠️ Cadastro usa texto, mas outras bases usam int |
| 2 | Participante ativo = status "Ativo" | Cadastro | status | ✅ 13,375 ativos (63.3%) |
| 3 | Elegibilidade = Ativo + aceite válido | Cadastro + Aceites | status + DataAceite | ⚠️ 35% sem data de aceite no cadastro |
| 4 | Treinamento concluído = CPF na base | Treinamentos | CPF + Estado | ✅ Só tem quem concluiu |
| 5 | CPF ausente em treinamentos = "Não fez" | Cadastro + Treinamentos | - | ✅ Pode ser calculado |
| 6 | Aceite mensal = Base_aceites | Aceites | CPF + DataAceite | ✅ Por mês |
| 7 | SUPERTOP = produto marcado (campanha sazonal) | Vendas | SUPERTOP | ✅ Nulo em ABR/2026 = não houve investimento neste mês |
| 8 | Pontuação por vendas | Vendas | Pontuação | ✅ Preenchido |
| 9 | Último mês = última data na base de vendas | Vendas | data venda | ✅ Abr/2026 |
| 10 | Setembro/2025 = base inexistente | - | - | ✅ Não existe arquivo |
| 11 | Status "Somente Catálogo" → "Inativo" | Cadastro | status | ✅ Mapeado |
| 12 | Pré-cadastrado → Inativo em 90 dias | Cadastro | status + data | ⚠️ Não automatizado |
| 13 | Pontuação aniversário (10 pts) | - | - | ❌ Não está em nenhuma base |
| 14 | Pontuação GI (300 pts) | Aptos | gi + pontos | ⚠️ Não separado no dashboard |
| 15 | Pontuação Gerente (10%) | Lideranca_Aptos | pontos_final | ⚠️ Parcial |
| 16 | Pontuação Gerente Regional (5%) | Lideranca_Aptos | pontos_final | ⚠️ Parcial |
| 17 | Valor do ponto = R$ 1,00 | Resgates | VALOR | ✅ Confirma 1:1 |
| 18 | Expiração: 12 meses + 2º dia útil | Pontos | EXPIRADO | ⚠️ Campo existe, regra não documentada |
| 19 | Acesso mensal obrigatório | Acessos | Data Acesso | ✅ Base recebida mensalmente |
| 20 | Formas de resgate | Resgates + Pontos | PRODUTO + abas | ✅ Documentado |
| 21 | Prazo resgate após desligamento: 90 dias | - | - | ⚠️ Não rastreável |
| 22 | Prazo contestação: 7 dias | Ocorrências | Data/Hora abertura | ⚠️ Não validado |

## 3.2 Hierarquia de Dados

### Fluxo de uma Venda → Pontuação:
```
Vendedor vende produto
    ↓
Rodrigo (time de dados) recebe e processa base de vendas
    ↓
Base de vendas (mensal) → SharePoint
    ↓
Power BI lê base de vendas + Cadastro + Aceites + Treinamentos
    ↓
Filtros de elegibilidade aplicados:
   - Status = Ativo
   - Aceite válido no mês
   - Treinamento concluído (quando aplicável)
    ↓
Pontuação calculada (vendas × pontos por produto)
    ↓
Dashboard mostra vendas e pontuação
    ↓
Apuração final Whirlpool inclui:
   - Vendas (do dashboard)
   + Pontuação aniversário (10 pts)
   + Pontuação GI (300 pts)
   + Pontuação Gerente (10%)
   + Pontuação Gerente Regional (5%)
```

### Fluxo de Resgate:
```
Participante acessa plataforma → página "Resgatar"
    ↓
Escolhe produto (PIX, transferência, vale, etc.)
    ↓
Solicita resgate na plataforma
    ↓
Cauê Webster processa resgate
    ↓
Base de resgates atualizada (mensal)
    ↓
Power BI atualiza saldo
```

### Fluxo de Treinamento:
```
Participante acessa plataforma → página "Treinamentos"
    ↓
Faz curso obrigatório
    ↓
Plataforma registra conclusão (progresso 100%)
    ↓
Base de treinamentos atualizada (diária)
    ↓
Power BI verifica se CPF está na base
    ↓
CPF ausente = "Não fez treinamento"
```

---

# PARTE 4: AUTO-PERGUNTAS E RESPOSTAS (TREINAMENTO)

## 4.1 Perguntas sobre Estrutura das Bases

**P1: Por que existem duas bases de cadastro (cadastro.xlsx e cadastro_revenda.xlsx)?**
**R:** A `cadastro_revenda.xlsx` é uma **atualização e complemento** da base principal `cadastro.xlsx`. Ela adiciona informações mais atualizadas de:
- Quais e quantas lojas fazem parte de cada revenda
- CNPJs das revendas e lojas
- Endereços, emails e outros dados de contato

Nem todos os CPFs do cadastro principal aparecem na cadastro_revenda (12,066 não estão). Possíveis razões: participantes sem loja associada, vínculo desatualizado, ou status diferente. No Power BI, ambas são usadas — a cadastro como base mestre e a cadastro_revenda para enriquecer com vínculo revenda→loja.

**P2: Por que a base de pontos tem mais CPFs (27,388) do que o cadastro (21,144)?**
**R:** A base de pontos mistura dados da edição anterior do programa (jan/2024 a set/2025) com a edição atual (out/2025 em diante). O programa atual começou em 20/10/2025, mas muitos participantes da edição anterior ainda têm saldo de pontos e resgates. No dashboard, isso é tratado como "Total créditos desde jan/2024".

**P3: Por que a base de acessos tem 749 mil registros mas apenas 14,030 CPFs únicos?**
**R:** A base de acessos é um LOG — cada linha é um acesso a uma página. Um participante pode acessar várias vezes por dia. São 14,030 CPFs únicos que acessaram a plataforma no período (out/2025 a jun/2026), totalizando 749,858 acessos. A média é de ~53 acessos por CPF no período.

**P4: Por que a base de treinamentos só tem registros com Progresso = 100?**
**R:** A plataforma só exporta para a base quem CONCLUIU o treinamento. Quem está "em andamento" ou "não iniciou" não aparece. Isso significa que para saber quem NÃO fez o treinamento, preciso fazer: CPFs em Cadastro Ativo com Aceite - CPFs em Treinamentos = "Não fez".

**P5: Por que o campo SUPERTOP está nulo em todas as linhas de 2026_04?**
**R:** Nem todos os meses têm produtos SUPERTOP. É uma **campanha sazonal** que depende do orçamento/investimento da Whirlpool no mês. Em ABR/2026 (e MAI/2026), a Whirlpool não investiu em produtos SUPERTOP — o dinheiro foi direcionado para outros pontos do programa. Quando há SUPERTOP no mês, o campo é preenchido na base de vendas. Quando não há, fica nulo ou zerado.

## 4.2 Perguntas sobre Regras de Negócio

**P6: Como é calculada a elegibilidade de um participante?**
**R:** A elegibilidade requer 3 condições:
1. Status = "Ativo" na base de cadastro
2. Aceite válido no mês (CPF presente na aba correspondente de Base_aceites)
3. Treinamento concluído (CPF presente na Base_treinamentos) — quando aplicável

⚠️ **Problema:** 35% dos cadastrados não têm data de aceite no cadastro.xlsx, mas a elegibilidade deve ser verificada pela Base_aceites mensal.

**P7: Como a pontuação do Gerente (10%) é calculada?**
**R:** O regulamento diz que o Gerente recebe 10% da pontuação da sua equipe de vendedores. Na prática, isso é calculado na apuração da Whirlpool (não no dashboard). A base `Lideranca_Aptos` contém os pontos finais de liderança, que já incluem esse percentual. O dashboard mostra esses pontos na página de liderança, mas a regra exata do cálculo não está documentada no Power BI.

**P8: Qual a diferença entre aceite de cadastro e aceite mensal?**
**R:** Dois aceites diferentes:
- **Aceite de cadastro** (campo `data de aceite` no cadastro.xlsx): Aceite do regulamento do programa feito no momento do cadastro. Só precisa ser feito uma vez.
- **Aceite mensal** (Base_aceites): Aceite à campanha do mês na plataforma. Precisa ser renovado mensalmente para estar "apto".

**P9: Por que o dashboard mostra uma pontuação diferente da apuração da Whirlpool?**
**R:** O dashboard só mostra:
- Vendas dos participantes

A apuração da Whirlpool inclui ADICIONALMENTE:
- Pontuação de aniversário (10 pts)
- Pontuação do Gestor da Informação (300 pts)
- Pontuação do Gerente (10% da equipe)
- Pontuação do Gerente Regional (5% da equipe)

Essas pontuações extras são calculadas pelo time de engenharia de dados da Whirlpool e NÃO estão disponíveis nas bases do dashboard.

**P10: Como identificar participantes "Somente Catálogo" no dashboard?**
**R:** O regulamento define "Somente Catálogo" como status após 2 meses sem acesso (não acumula pontos, só resgata). Mas a plataforma simplifica e manda direto para "Inativo". No dashboard, esses participantes aparecem como "Inativo". A distinção não existe na prática.

## 4.3 Perguntas sobre Qualidade de Dados

**P11: Por que tantos CPFs no cadastro não têm data de aceite?**
**R:** 7,490 registros (35.4%) têm data de aceite nula no cadastro.xlsx. Possíveis causas:
- Participantes pré-cadastrados que nunca aceitaram o regulamento
- Participantes de edição anterior que migraram sem aceite
- Campo não obrigatório na plataforma
- Importação de dados legados sem esse campo

**P12: Por que a base de aceites tem apenas 8,021 CPFs em Abr/2026?**
**R:** O aceite mensal é feito pelo participante na plataforma. Nem todos os cadastrados fazem o aceite todo mês. Isso é normal — apenas quem quer participar da campanha do mês faz o aceite. A variação entre meses (4,664 em Fev/2026 vs 8,021 em Abr/2026) pode indicar sazonalidade, mudança de regra ou problema na coleta.

**P13: Por que a base de vendas tem 7,820 registros com pontuação = 0?**
**R:** São produtos vendidos que NÃO geram pontos. Pode ser:
- Produtos fora da lista de produtos participantes
- Produtos de marcas não Whirlpool
- Produtos com pontuação = 0 no SKU
- Ajustes/estornos

**P14: Como tratar CPFs com zeros à esquerda quando algumas bases usam int?**
**R:** A regra diz "CPF SEMPRE como texto". Quando a base vem como int (perde zeros), preciso:
1. Converter para string
2. Preencher com zeros à esquerda até 11 dígitos
3. Fazer o join com a base de cadastro (que já tem como texto)

Exemplo: CPF 123456789 como int → "00123456789" como str

**P15: Por que existem CPFs em acessos que não estão no cadastro?**
**R:** 8,097 CPFs em acessos que não estão no cadastro atual. Possíveis causas:
- Participantes que saíram do programa (inativados/excluídos)
- Participantes de edição anterior que ainda acessam
- CPFs de teste/administrativos
- Erro de importação

## 4.4 Perguntas sobre KPIs e Métricas

**P16: Como calcular a taxa de ativação?**
**R:** Taxa de Ativação = Cadastros Ativos / Total de Cadastros
= 13,375 / 21,144 = 63.3%

**P17: Como calcular a taxa de conclusão de treinamentos?**
**R:** Taxa de Conclusão = Cadastros Ativos com Aceite que fizeram treinamento / Cadastros Ativos com Aceite
= (CPFs em Treinamentos ∩ Cadastro Ativo ∩ Aceites) / (Cadastro Ativo ∩ Aceites)

⚠️ Não posso fazer simplesmente Treinamentos / Cadastro porque nem todos os cadastrados precisam fazer treinamento (aplicável por campanha).

**P18: Como calcular o YoY (Year over Year)?**
**R:** YoY compara o último mês disponível com o mesmo mês do ano anterior.
- Último mês: Abr/2026 (última base de vendas)
- Mês anterior: Mar/2026
- YoY: Abr/2026 vs Abr/2025

**P19: Como calcular o aproveitamento de potencial?**
**R:** Aproveitamento % = Pontos Alcançados / Pontos Potenciais
A base `bases_potencial` tem o potencial de pontuação por revenda/mês. O aproveitamento é a divisão do que foi atingido pelo que era possível.

**P20: Como identificar o TOP SKU?**
**R:** TOP SKU = SKU com maior quantidade vendida (ou maior pontuação) no período filtrado. É um ranking dinâmico que muda conforme os filtros de revenda, regional, período aplicados.

## 4.5 Perguntas sobre Processos

**P21: Qual o processo de atualização mensal?**
**R:**
1. Receber bases no SharePoint/e-mail (Rodrigo: vendas, aceites, pontos, resgates, supertop)
2. Validar estrutura (colunas, tipos, CPF como texto)
3. Processar vendas (pipeline Python quando em uso)
4. Atualizar SharePoint com novos arquivos mantendo nomenclatura
5. Atualizar Power BI (atualizar consultas, validar erros)
6. Validar KPIs (comparar totais com origem)
7. Publicar no workspace/app

**P22: Quem é responsável por cada base?**
**R:**
| Base | Responsável | Frequência |
|------|------------|------------|
| Vendas | Rodrigo Moraes | Mensal (5º dia útil) |
| Cadastros | Plataforma +TOP | Diária |
| Acessos | Plataforma +TOP | Diária |
| Aceites | Rodrigo Moraes | Mensal |
| Pontuação | Rodrigo/Cauê Webster | Mensal |
| Resgates | Cauê Webster | Mensal |
| Chamados | Andressa Borini | Semanal (segunda) |
| SUPERTOP | Rodrigo Moraes | Mensal |
| Potencial | Matheus/Rodrigo | Mensal |
| Treinamentos | Plataforma +TOP | Diária |
| Enquete | Plataforma +TOP | Diária |

**P23: O que fazer quando o arquivo mensal não chega?**
**R:**
1. Registrar pendência
2. Solicitar ao responsável (Rodrigo)
3. Documentar impacto no dashboard
4. NUNCA interpolar ou estimar (principalmente set/2025 que nunca chegará)

**P24: Como validar se os KPIs do Power BI estão corretos?**
**R:**
1. Comparar soma de Vendas e Pontuação com o consolidado do time de dados
2. Validar CPFs distintos no cadastro (sem duplicidades)
3. Comparar acessos com relatório da plataforma
4. Verificar CPFs com aceite, datas e formato de 11 dígitos
5. Validar créditos, resgates, expirados e saldo
6. Testar filtros (Revenda, Regional, SKU, Categoria, CNPJ, Período, Status, Cargo)

**P25: Quais são os problemas recorrentes?**
**R:**
| Problema | Ação |
|----------|------|
| Arquivo mensal não recebido | Registrar, solicitar, documentar impacto |
| Coluna ausente/renomeada | Ajustar Power Query ou solicitar reenvio |
| CPF/CNPJ perdeu zeros | Garantir tipo texto, normalizar |
| Erro de relacionamento | Verificar chaves, cardinalidade, duplicidades |
| Indicador diferente da origem | Revisar filtros, período, status, aceite |
| Revenda nomes duplicados | Atualizar tabela DE/PARA tab_nome_revenda |
| Treinamentos não mostram quem não fez | Criar flag: Ativos com Aceite - CPFs em treinamentos |

---

# PARTE 5: INSIGHTS E RECOMENDAÇÕES

## 5.1 Insights de Dados

### Insight 1: Conversão de Cadastro → Venda (Base Total vs Ativos)
**Base total (21,144):**
- 5,696 venderam em Abr/2026 (26.9%)
- 8,021 aceitaram em Abr/2026 (37.9%)

**Apenas ATIVOS (13,375):**
- 2,356 venderam em Abr/2026 (17.6%)
- 3,320 aceitaram em Abr/2026 (24.8%)

> Para vendas e aceites, consideramos apenas ATIVOS. A análise da base total ajuda a entender o funil completo.

### Insight 2: Dados de Edição Anterior (Confirmado)
- O programa atual começou em **20/10/2025**
- Dados anteriores (jan/2024 a set/2025) são da **edição anterior** do +TOP
- Pontos: 20,629 CPFs (75%) são de participantes da edição anterior
- Resgates: 13,259 CPFs (70%) são de participantes da edição anterior
- **Nota:** Isso é esperado — participantes da edição anterior mantêm saldo e histórico de resgates

### Insight 3: Cadastro_Revenda como Complemento (não duplicidade)
- A base `cadastro_revenda` (20,558) é uma **atualização/complemento** da base principal (21,144)
- Adiciona informações mais atualizadas de lojas, CNPJs, endereços, emails
- 12,066 CPFs do cadastro não aparecem na cadastro_revenda — possíveis razões: sem loja associada, vínculo desatualizado, ou status diferente

### Insight 4: Treinamentos com Baixa Adesão
**Base total:** 5,642 CPFs fizeram treinamentos (de ~21k) = 27%
**Apenas ATIVOS:** 2,416 CPFs fizeram treinamentos (de 13,375) = 18.1%
- **Recomendação:** Investigar se treinamento é realmente obrigatório para pontuação

### Insight 5: Home Domina os Acessos
- 55% dos acessos são à Home
- Página "Resgatar" é a segunda com 11.5%
- **Recomendação:** A plataforma está sendo usada principalmente para consulta e resgate

## 5.2 Checklist de Validação para Novas Bases

- [ ] CPF/CNPJ como texto (11 dígitos para CPF, 14 para CNPJ)
- [ ] Colunas no layout padrão (nomes e ordem corretos)
- [ ] Sem linhas em branco ou duplicadas
- [ ] Data no formato correto
- [ ] Valores numéricos sem caracteres especiais
- [ ] Revendas com nomes padronizados (ou atualizar DE/PARA)
- [ ] Comparação com base anterior (variações esperadas)

---

*Documento gerado em 2026-06-08 após análise completa de 17 bases e cruzamento com regulamento do Programa +TOP.*
*Analista virtual: Claude Code (simulação de analista sênior de dados)*
