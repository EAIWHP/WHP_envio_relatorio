# Análise: Regulamento vs Documentação Técnica do Dashboard

**Data da análise:** 2026-06-03  
**Analista:** Claude Code  
**Fonte do regulamento:** Termos e Condições do Programa +TOP e Uso da Plataforma  
**Fonte da documentação técnica:** Documentacao_Dashboard_TOP_Formatada_Modelo_Profissional.docx

---

## RESUMO EXECUTIVO

Foram identificadas **27 inconsistências, gaps ou regras não documentadas** no dashboard/documentação técnica quando comparadas ao regulamento oficial do Programa +TOP.

| Severidade | Quantidade | Significado |
|------------|-----------|-------------|
| 🔴 **Crítica** | 8 | Regras de negócio obrigatórias pelo regulamento que não estão no dashboard |
| 🟡 **Alta** | 10 | Regras importantes que afetam cálculos ou validações |
| 🟢 **Média** | 9 | Melhorias de documentação ou controles que devem existir |

---

## 🔴 INCONSISTÊNCIAS CRÍTICAS (Regras obrigatórias do regulamento AUSENTES no dashboard)

### IC-01: Status de Cadastro "Somente Catálogo" — NÃO DOCUMENTADO

**O que o regulamento diz (item 2.4):**
> "Caso o Participante não realize ao menos um login por mês na Plataforma (...) por 02 (dois) meses consecutivos, o perfil do Participante será alterado para **Somente Catálogo**, o que fará com que ele não possa acumular novos pontos, sendo assegurado a ele o direito de resgatar seus pontos."

**O que a documentação técnica diz:**
- Status documentados: Ativo, Inativo, Aguardando aprovação, Em moderação, Indicado Moderado, Pré-cadastrado, Reprovado
- **"Somente Catálogo" NÃO APARECE** no dicionário de dados nem nas regras de negócio

**Impacto:** O dashboard pode estar classificando participantes "Somente Catálogo" como "Ativo" ou "Inativo", o que distorce KPIs de ativação, pontuação e elegibilidade.

**Ação necessária:** Adicionar "Somente Catálogo" como status válido no modelo de dados, no dicionário de dados, e nas regras de KPIs.

---

### IC-02: Status "Pré-cadastrado" com regra de inativação em 90 dias — NÃO DOCUMENTADO

**O que o regulamento diz (item 2.2.3):**
> "Caso o Participante não ative seu cadastro na plataforma por prazo superior a **90 (noventa) dias**, contados da data de seu cadastro no sistema (...) o seu acesso/perfil será automaticamente alterado para **inativo**"

**O que a documentação técnica diz:**
- O status "Pré-cadastrado" existe no dicionário
- Mas **não há nenhuma regra** que documente a transição automática para "Inativo" após 90 dias

**Impacto:** Participantes pré-cadastrados há mais de 90 dias podem estar sendo contados erroneamente em KPIs.

**Ação necessária:** Documentar a regra de transição e verificar se o dashboard trata esses casos corretamente.

---

### IC-03: Pontuação de aniversário (10 pontos) — NÃO DOCUMENTADA

**O que o regulamento diz:**
- Item 4.5: Vendedor faz jus a **10 pontos no mês de aniversário**
- Item 5.3: Gerente recebe **10 pontos no mês de aniversário**
- Item 6.2: Gerente Regional recebe **10 pontos no mês de aniversário**
- Item 7.3.4: Gestor da Informação faz jus a **10 pontos no mês de aniversário**

**O que a documentação técnica diz:**
- **Nenhuma menção** a pontuação de aniversário nas regras de negócio, KPIs ou fórmulas DAX

**Impacto:** Os 10 pontos de aniversário podem não estar sendo creditados no cálculo de pontuação total.

**Ação necessária:** Verificar se há uma flag/mechanismo que aplica esses pontos. Se não houver, precisa ser implementado.

---

### IC-04: Pontuação do Gestor da Informação (300 pontos + bônus trimestral) — NÃO DOCUMENTADA

**O que o regulamento diz (item 7.3):**
> "O Gestor da Informação receberá **300 (trezentos) pontos** sempre que fizer o envio correto das informações, dentro do modelo padrão e dentro do prazo"

> "Se, por 03 (três) meses consecutivos, o Gestor da Informação enviar corretamente (...) receberá **300 (trezentos) pontos adicionais**"

> "A pontuação trimestral será creditada nos meses de **março, junho, setembro e dezembro**"

**O que a documentação técnica diz:**
- **Nenhuma menção** a pontuação específica de Gestor da Informação
- A base de vendas aparentemente computa pontos apenas para vendedores

**Impacto:** Gestores da Informação podem não estar recebendo pontuação correta no sistema.

**Ação necessária:** Verificar se existe uma tabela ou mecanismo que registra esses pontos. Se não, precisa ser criado.

---

### IC-05: Pontuação de Gerentes (10% da equipe) — REGRA PARCIALMENTE DOCUMENTADA

**O que o regulamento diz (item 5.1):**
> "Os Gerentes Participantes (...) receberão **10% (dez por cento) da pontuação alcançada pela sua equipe de Vendedores Participantes**"

**O que a documentação técnica diz:**
- O item "cálculo de pontuação de vendedores e líderes" é mencionado como tratamento da base de vendas
- Mas **não há detalhamento** de como esses 10% são calculados nem como aparecem no dashboard

**Impacto:** Pode haver dupla contagem (vendedor pontua + gerente recebe 10%) ou a pontuação do gerente pode não estar sendo separada corretamente.

**Ação necessária:** Documentar como a pontuação de gerentes é calculada e armazenada na base de vendas.

---

### IC-06: Pontuação de Gerentes Regionais (5% da equipe) — NÃO DOCUMENTADA

**O que o regulamento diz (item 6.1):**
> "Os Gerentes Regionais (...) receberão **5% (cinco por cento) da pontuação alcançada** no mês, atribuída aos Vendedores Participantes das lojas atendidas por ele"

**O que a documentação técnica diz:**
- **Nenhuma menção** a pontuação de Gerentes Regionais

**Impacto:** Gerentes Regionais podem não estar recebendo pontuação correta.

**Ação necessária:** Verificar se existe mecanismo de cálculo e, se não, implementar.

---

### IC-07: Valor do ponto = R$ 1,00 — NÃO DOCUMENTADO

**O que o regulamento diz (item 3.4):**
> "o valor de cada ponto conquistado pelo Participante neste Programa é de **R$1,00 (um real)**"

**O que a documentação técnica diz:**
- **Nenhuma menção** ao valor monetário do ponto
- Isso é importante para o relatório de resgates e análise de custo/benefício

**Impacto:** Análises financeiras podem estar incorretas se baseadas em outro valor.

**Ação necessária:** Documentar o valor do ponto e verificar se é usado em cálculos de resgate.

---

### IC-08: Prazo de validade dos pontos (12 meses + expiração automática) — REGRA INCOMPLETA

**O que o regulamento diz (item 9.1):**
> "o Participante terá até **12 (doze) meses** após a data de crédito da pontuação para resgatar os pontos acumulados"

> "expirando-se automaticamente todo **2º dia útil de cada mês**"

**O que a documentação técnica diz:**
- O campo "Expirado" existe em `tab_pontos_por_cpf`
- Mas **não há documentação** sobre como a expiração é calculada (12 meses) nem quando ocorre (2º dia útil)

**Impacto:** O cálculo de pontos expirados pode estar incorreto se não seguir essas regras.

**Ação necessária:** Documentar a regra de expiração e verificar se o cálculo atual está correto.

---

## 🟡 INCONSISTÊNCIAS ALTAS (Regras importantes com gaps)

### IA-01: Regra de desempate de campanhas adicionais — NÃO DOCUMENTADA

**O que o regulamento diz (item 10.4):**
> Ordem de desempate:
> 1. Maior número de unidades vendidas de produtos supertops
> 2. Maior número de unidades de produtos participantes
> 3. Maior quantidade de treinamentos obrigatórios realizados
> 4. Data mais antiga de cadastro
> 5. Maior quantidade de produtos vendidos na ordem: refrigeradores, lavadoras, fogões, microondas

**O que a documentação técnica diz:**
- **Nenhuma menção** a regras de desempate
- As campanhas adicionais são mencionadas como "sob demanda" e "sem demanda ativa no momento"

**Impacto:** Se houver campanhas adicionais no futuro, não há como aplicar o desempate corretamente.

---

### IA-02: Regra de 80% de acertos em treinamentos para pontuação extra — NÃO DOCUMENTADA

**O que o regulamento diz (item 4.4):**
> "tendo necessariamente de acertar, no mínimo, **80% (oitenta por cento)** da avaliação de cada treinamento. O não atingimento da meta mínima de acertos impossibilita ao Vendedor Participante o recebimento de pontos extras"

**O que a documentação técnica diz:**
- Treinamentos são documentados como "cursos obrigatórios concluídos"
- Mas **não há menção** à nota mínima de 80%
- O dashboard atual verifica apenas "progresso/conclusão", não nota/avaliação

**Impacto:** Vendedores com conclusão mas nota < 80% podem estar sendo considerados elegíveis para pontuação extra incorretamente.

---

### IA-03: Participação só vale após cadastro + aceite + acesso mensal — REGRA PARCIAL

**O que o regulamento diz (item 2.3):**
> "a sua pontuação somente valerá a partir da realização do seu cadastro no Site, a sua posterior ativação com confirmação pela Whirlpool. Para ser elegível ao recebimento de qualquer pontuação, o Participante deverá, ainda, **acessar a Plataforma ao menos uma vez ao mês** para validar a sua participação"

**O que a documentação técnica diz:**
- Documenta "Elegibilidade: Ativo + aceite válido + regras de treinamento"
- Mas **não inclui** a regra de "acesso mensal à plataforma" como critério de elegibilidade

**Impacto:** Participantes ativos com aceite mas sem acesso mensal podem estar sendo considerados elegíveis incorretamente.

---

### IA-04: Vendas só contam a partir do mês de cadastro — NÃO DOCUMENTADO

**O que o regulamento diz (item 2.5):**
> "Os Participantes que forem admitidos pelas Revendas Participantes após o início do Programa terão seus pontos computados **somente a partir do mês em que** a própria Revenda Participante fizer o encaminhamento das informações (...) bem como a partir de quando for concluído o cadastro do Participante"

**O que a documentação técnica diz:**
- **Nenhuma menção** a essa regra de corte temporal
- A base de vendas parece considerar todas as vendas independentemente da data de cadastro

**Impacto:** Vendas de participantes que ainda não estavam cadastrados podem estar sendo contabilizadas.

---

### IA-05: Pontos não são pagos em dinheiro — NÃO DOCUMENTADO

**O que o regulamento diz (item 3.2):**
> "Em nenhuma hipótese haverá **pagamento em dinheiro** ao Participante ou conversão dos pontos em qualquer outro benefício"

**O que a documentação técnica diz:**
- **Nenhuma menção** a essa restrição
- Importante para auditoria e relatórios financeiros

---

### IA-06: Prazo de resgate após desligamento (90 dias) — NÃO DOCUMENTADO

**O que o regulamento diz (item 2.4.2):**
> "o Participante terá o prazo improrrogável de até **90 (noventa) dias** para o resgate dos pontos acumulados por ele após a inativação da Revenda Participante"

**O que a documentação técnica diz:**
- **Nenhuma menção** a esse prazo de resgate
- Importante para cálculo de saldo e relatórios de pontos expirados

---

### IA-07: Prazo de contestação de pontuação (7 dias) — NÃO DOCUMENTADO

**O que o regulamento diz (item 9.6):**
> "O prazo para contestação de qualquer pontuação é de **07 (sete) dias corridos** contados da data do seu crédito"

**O que a documentação técnica diz:**
- **Nenhuma menção** a prazo de contestação
- Importe para processo de validação e chamados

---

### IA-08: Campanhas adicionais requerem aceite específico — NÃO DOCUMENTADO

**O que o regulamento diz (item 10.2):**
> "Para participação de campanhas adicionais, as pessoas elegíveis (...) deverão acessar o Site e **aceitar a respectiva Campanha**"

**O que a documentação técnica diz:**
- Menção genérica a "Aceites campanhas extra" como "sob demanda"
- Mas **não há estrutura** no modelo de dados para separar aceites de campanha base vs. campanhas adicionais

---

### IA-09: Gestor da Informação — regras de envio até 5º dia útil — NÃO TOTALMENTE ALINHADO

**O que o regulamento diz (item 7.2):**
> "enviar corretamente as bases de Vendedores Participantes e informações sobre as vendas realizadas no respectivo mês (...) até no máximo o **5º (quinto) dia útil do mês subsequente**"

**O que a documentação técnica diz:**
> "Envio pelos Gestores da Informação até o 5º dia útil"

Mas a pontuação do Gestor (300 pontos) depende desse prazo, e isso **não está no dashboard**.

---

### IA-10: Cada participante só pode ter UM cargo — NÃO DOCUMENTADO

**O que o regulamento diz (item 2.1.1):**
> "Cada funcionário da Revenda Participante somente poderá ser cadastrado como ocupante de **um único cargo**, sendo que, em nenhuma hipótese, poderá o funcionário (...) ter mais de um cargo cadastrado no Programa"

**O que a documentação técnica diz:**
- **Nenhuma validação** documentada para impedir CPF com mais de um cargo
- Isso pode causar duplicidade na base de cadastro

---

## 🟢 INCONSISTÊNCIAS MÉDIAS (Melhorias de documentação/controle)

### IM-01: Vigência do programa (20/10/2025 a 31/12/2026) — NÃO DOCUMENTADA

**O que o regulamento diz (item 3.1):**
> "Este Programa terá duração de **20 de outubro de 2025 a 31 de dezembro de 2026**"

**O que a documentação técnica diz:**
- **Nenhuma menção** à vigência do programa
- O dashboard tem dados desde jan/2024, mas o programa atual começou em out/2025

**Pergunta:** Os dados de jan/2024 a set/2025 são de uma edição anterior do programa?

---

### IM-02: Site oficial do programa (https://programamaistop.com.br) — NÃO DOCUMENTADO

**O que o regulamento diz:**
> "Site https://programamaistop.com.br"

**O que a documentação técnica diz:**
- Referências genéricas à "Plataforma + TOP"
- Mas **não há URL** documentada

---

### IM-03: Central de Atendimento (0800 780 0606) — NÃO DOCUMENTADO

**O que o regulamento diz:**
> "0800 780 0606 ou www.programamaistop.com.br/FaleConosco/"

**O que a documentação técnica diz:**
- Base "Fale_Conosco" existe, mas **não há menção** ao canal de atendimento oficial

---

### IM-04: Revenda pode declinar participação a qualquer tempo — NÃO DOCUMENTADO

**O que o regulamento diz (item 1, parágrafo 2):**
> "a Revenda Participante é (e sempre será) livre para aderir e para **declinar da sua participação no Programa, a qualquer tempo**"

**O que a documentação técnica diz:**
- **Nenhuma menção** a revendas que saíram do programa
- Isso pode afetar análises históricas (vendas de revendas que já saíram)

---

### IM-05: Responsabilidade da Revenda pela veracidade dos dados — NÃO DOCUMENTADO

**O que o regulamento diz (item 1.2.1):**
> "A responsabilidade pelo envio e veracidade das informações das vendas realizadas (...) é **exclusivamente da Revenda Participante**"

**O que a documentação técnica diz:**
- **Nenhuma menção** a essa responsabilidade
- Importante para auditoria e validação de dados

---

### IM-06: Resgate parcial ou integral — NÃO DOCUMENTADO

**O que o regulamento diz (item 8.5):**
> "O resgate dos pontos poderá ser **parcial ou integral**"

**O que a documentação técnica diz:**
- **Nenhuma menção** a resgates parciais
- Pode afetar o cálculo de saldo e análise de resgates

---

### IM-07: Indisponibilidade de prêmio no resgate — NÃO DOCUMENTADO

**O que o regulamento diz (item 8.4):**
> "No caso de indisponibilidade do Prêmio escolhido (...) caberá a ele (...) selecionar outro Prêmio que esteja disponível"

**O que a documentação técnica diz:**
- **Nenhuma menção** a essa regra no processo de resgate

---

### IM-08: Saldo disponível até 20º dia útil do mês de apuração — NÃO DOCUMENTADO

**O que o regulamento diz (item 8.3):**
> "Os saldos atualizados dos pontos estarão disponíveis para o Participante até o **20º (vigésimo) dia útil do mês de apuração**, desde que a Revenda Participante forneça as informações até o 15º (décimo quinto) dia deste mesmo mês"

**O que a documentação técnica diz:**
- **Nenhuma menção** a esse prazo de disponibilização de saldo

---

### IM-09: Política de Privacidade e LGPD — NÃO DOCUMENTADA

**O que o regulamento diz:**
> Seção completa de Política de Privacidade com direitos do titular, cookies, compartilhamento, retenção de dados

**O que a documentação técnica diz:**
- **Nenhuma menção** a LGPD, privacidade ou retenção de dados
- Item 1.3 trata de controladoras independentes de dados

---

## 📋 SÍNTESE DAS PERGUNTAS AO USUÁRIO

### Perguntas sobre Status de Cadastro

1. **"Somente Catálogo"**: Esse status existe na base `02_Cadastro`? Se sim, como o dashboard atualmente o classifica (Ativo/Inativo)?

2. **"Pré-cadastrado" → "Inativo" após 90 dias**: Existe algum job/processamento que faz essa transição automaticamente, ou é manual?

### Perguntas sobre Pontuação

3. **Pontuação de aniversário (10 pontos)**: Existe algum mecanismo que credita os 10 pontos de aniversário? Onde isso aparece na base de vendas/pontuação?

4. **Pontuação de Gestor da Informação (300 pontos)**: Existe uma tabela/base separada para os pontos do Gestor da Informação? Como eles entram no `tab_pontos_por_cpf`?

5. **Pontuação de Gerente (10%) e Gerente Regional (5%)**: Como esses percentuais são calculados e onde são armazenados? São incluídos na `06_Base_Vendas` ou em outra tabela?

6. **Valor do ponto = R$ 1,00**: O dashboard usa esse valor em algum cálculo? (ex: análise de custo dos resgates)

### Perguntas sobre Regras de Elegibilidade

7. **Acesso mensal obrigatório**: O dashboard considera o acesso mensal como critério de elegibilidade para pontuação? Ou apenas o status "Ativo" + aceite?

8. **80% de acertos em treinamentos**: A base `Base_treinamentos` tem coluna de nota/avaliação? O dashboard verifica apenas conclusão ou também a nota?

### Perguntas sobre Campanhas Adicionais

9. **Campanhas adicionais**: Existe alguma estrutura no modelo de dados para campanhas adicionais (tabela de aceites de campanhas, por exemplo)? O item 10.4 menciona "sob demanda" e "sem demanda ativa no momento".

### Perguntas sobre Dados Históricos

10. **Vigência do programa**: O programa atual é de 20/10/2025 a 31/12/2026. Mas o dashboard tem dados desde jan/2024. Esses dados anteriores são de uma edição anterior do +TOP? Devem ser tratados de forma diferente?

### Perguntas sobre Validação

11. **CPF com mais de um cargo**: Existe validação na base `02_Cadastro` para impedir que um mesmo CPF tenha mais de um cargo?

12. **Revendas que saíram do programa**: Como identificar no dashboard se uma revenda declinou da participação? Existe algum flag ou data de saída?

---

## ✅ ALINHAMENTOS CONFIRMADOS (Regulamento = Documentação Técnica)

| Regra | Status |
|-------|--------|
| Cargos permitidos: vendedor, gerente, gerente regional, gestor da informação | ✅ OK |
| Aceite mensal para estar apto | ✅ OK (via Base_aceites) |
| Campanha base + campanhas extras (sazonais) | ✅ OK |
| Inatividade após 2 meses sem acesso, inativo após 3 meses | ⚠️ Parcial (status "Inativo" existe, mas "Somente Catálogo" não) |
| Pontuação por vendas de produtos participantes | ✅ OK |
| SUPERTOP (produtos com pontuação maior) | ✅ OK |
| Resgates via catálogo de prêmios | ✅ OK |
| Expiração de pontos | ⚠️ Parcial (campo existe, mas regra de 12 meses + 2º dia útil não documentada) |
| CPF/CNPJ como texto | ✅ OK |
| Gestor da Informação envia dados até 5º dia útil | ✅ OK |

---

*Análise gerada automaticamente pelo Claude Code em 2026-06-03.*
*Próxima ação: aguardar respostas do usuário para atualizar documentação e corrigir inconsistências.*
