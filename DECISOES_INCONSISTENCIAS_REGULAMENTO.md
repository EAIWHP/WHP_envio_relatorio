# Decisões do Usuário — Inconsistências Regulamento vs Dashboard

**Data das decisões:** 2026-06-08  
**Analista:** Thamires Vieira  
**Base:** Análise de 27 inconsistências identificadas em 2026-06-03

---

## RESUMO DAS DECISÕES

| Severidade | Total | Documentar | Descartar/Não Aplicar |
|------------|-------|-----------|----------------------|
| 🔴 Críticas | 8 | **8** | 0 |
| 🟡 Altas | 10 | **3** | 7 |
| 🟢 Médias | 9 | **2** | 7 |
| **Total** | **27** | **13** | **14** |

---

## 🔴 CRÍTICAS — DECISÕES

### IC-01: Status "Somente Catálogo" ✅ ADICIONAR AO DICIONÁRIO
- **Decisão:** Adicionar ao dicionário de dados do dashboard
- **Observação:** O sistema da plataforma simplifica — após 2 meses sem acesso vai direto para "Inativo". O dashboard não distingue "Somente Catálogo" do "Inativo", ambos são tratados igual. Mas o status deve constar no dicionário para registro.
- **Status:** Documentar como status conhecido do regulamento, mapeado para "Inativo" no dashboard

### IC-02: Pré-cadastrado → Inativo em 90 dias ✅ DOCUMENTAR
- **Decisão:** Documentar essa regra
- **Ação:** Incluir no CLAUDE.md e no dicionário de dados que após 90 dias sem ativação, o status vira "Inativo"

### IC-03: Pontuação de aniversário (10 pts) ✅ DOCUMENTAR
- **Decisão:** Documentar essa regra
- **Observação:** NÃO está no dashboard. É creditada apenas na apuração final enviada para a Whirlpool. Não aparece em nenhuma base ou no dashboard.
- **Ação:** Documentar no CLAUDE.md como regra do regulamento que NÃO está refletida no dashboard

### IC-04: Pontuação Gestor da Informação (300 pts + bônus trimestral) ✅ DOCUMENTAR
- **Decisão:** Adicionar essa regra à documentação
- **Observação:** NÃO é utilizada em nenhuma parte do dashboard ou bases. Pertence à apuração da Whirlpool.
- **Ação:** Documentar no CLAUDE.md como regra do regulamento que NÃO está refletida no dashboard

### IC-05: Pontuação Gerente (10% da equipe) ✅ DOCUMENTAR EM 100%
- **Decisão:** Documentar corretamente em 100%
- **Ação:** Incluir no CLAUDE.md a regra completa de cálculo dos 10% sobre a pontuação da equipe de vendedores

### IC-06: Pontuação Gerente Regional (5% da equipe) ✅ DOCUMENTAR
- **Decisão:** Documentar
- **Observação:** NÃO está no dashboard nem nas bases. Pertence à apuração da Whirlpool.
- **Ação:** Documentar no CLAUDE.md como regra do regulamento que NÃO está refletida no dashboard

### IC-07: Valor do ponto = R$ 1,00 ✅ DOCUMENTAR
- **Decisão:** Documentar
- **Ação:** Incluir no CLAUDE.md o valor de R$ 1,00 por ponto

### IC-08: Expiração de pontos (12 meses + 2º dia útil) ✅ DOCUMENTAR
- **Decisão:** Documentar essa regra
- **Ação:** Incluir no CLAUDE.md a regra completa: 12 meses de validade + expiração automática no 2º dia útil de cada mês

---

## 🟡 ALTAS — DECISÕES

### IA-01: Critérios de desempate de campanhas ❌ NÃO DOCUMENTAR
- **Decisão:** Não é importante para o presente momento e documento
- **Justificativa:** Critérios de desempate dependem de fatores específicos que não são utilizados no dashboard

### IA-02: 80% de acertos em treinamentos ❌ NÃO DOCUMENTAR
- **Decisão:** Manter como está — só trazer quem concluiu
- **Justificativa:** A informação de concluiu ou não concluiu vem diretamente da plataforma. Não temos um campo na base que diz a porcentagem do que já foi e o que falta concluir. Quem não estiver na base de treinamentos = não concluiu.

### IA-03: Acesso mensal obrigatório ✅ JÁ ESTÁ NO DASHBOARD
- **Decisão:** Já está contemplado
- **Justificativa:** Todos os meses recebemos uma base da equipe da plataforma com os nomes dos participantes, datas de acesso à plataforma e aceite mensal. Verificamos através dessa base.

### IA-04: Vendas só a partir do mês de cadastro ❌ NÃO DOCUMENTAR
- **Decisão:** Não faz sentido trazer essa info
- **Justificativa:** Não temos informações de vendas realizadas antes do cadastro

### IA-05: Pontos não pagos em dinheiro ✅ DOCUMENTAR (com correção)
- **Decisão:** Documentar as formas de resgate existentes
- **Observação:** Os pontos podem ser trocados por: transferência bancária, PIX, cartão, lojas virtuais, pagamento de contas, recargas de celular e vouchers
- **Ação:** Incluir no CLAUDE.md as formas de resgate disponíveis

### IA-06: Prazo de resgate após desligamento (90 dias) ✅ DOCUMENTAR
- **Decisão:** Documentar esse prazo
- **Ação:** Incluir no CLAUDE.md

### IA-07: Prazo de contestação de pontuação (7 dias) ✅ DOCUMENTAR
- **Decisão:** Documentar esse prazo de contestação
- **Ação:** Incluir no CLAUDE.md

### IA-08: Aceite específico para campanhas adicionais ❌ NÃO DOCUMENTAR
- **Decisão:** Não precisamos dessa informação por enquanto
- **Justificativa:** Não verificamos campanhas extras dentro desse projeto

### IA-09: Gestor da Informação — envio até 5º dia útil ❌ NÃO DOCUMENTAR
- **Decisão:** Não precisamos no dashboard por enquanto
- **Justificativa:** Essa data limite só é utilizada para a apuração mensal realizada pelo time de engenharia de dados

### IA-10: CPF com um único cargo ❌ NÃO DOCUMENTAR
- **Decisão:** Não aplicar no momento
- **Justificativa:** Não temos validação documentada para impedir CPF com múltiplos cargos

---

## 🟢 MÉDIAS — DECISÕES

### IM-01: Vigência do programa (20/10/2025 a 31/12/2026) ✅ MANTER DATAS ATUAIS
- **Decisão:** Manter as datas que estão no dashboard
- **Justificativa:** O dashboard tem dados desde jan/2024 por conter dados de edição anterior

### IM-02: Site oficial (programamaistop.com.br) ✅ CONFIRMADO
- **Decisão:** Site oficial confirmado

### IM-03: Central de Atendimento ❌ NÃO TEMOS
- **Decisão:** Não documentar
- **Justificativa:** Não temos o número da central

### IM-04: Revenda pode sair a qualquer tempo ❌ NÃO DOCUMENTAR
- **Decisão:** Não documentar
- **Justificativa:** Não temos o histórico de quando a revenda entrou e quando saiu

### IM-05: Responsabilidade da revenda pela veracidade ❌ NÃO DOCUMENTAR
- **Decisão:** Não documentar
- **Justificativa:** Não temos um modo oficial de passar responsabilidade das informações

### IM-06: Resgate parcial ou integral ✅ JÁ EXISTE
- **Decisão:** Já temos essa informação
- **Justificativa:** Temos na base de resgates detalhado e no dashboard, porém é detalhado por revenda e participantes

### IM-07: Indisponibilidade de prêmio no resgate ❌ NÃO DOCUMENTAR
- **Decisão:** Não documentar
- **Justificativa:** Não temos histórico da indisponibilidade dos prêmios

### IM-08: Saldo disponível até 20º dia útil ✅ JÁ EXISTE
- **Decisão:** Já temos essa informação
- **Justificativa:** Temos nas bases e no dashboard

### IM-09: Política de Privacidade / LGPD ❌ NÃO DOCUMENTAR
- **Decisão:** Não precisamos adicionar à documentação por enquanto

---

## AÇÕES PENDENTES DE DOCUMENTAÇÃO

### Para incluir no CLAUDE.md:

1. **Status de Cadastro** — Adicionar ao dicionário:
   - "Somente Catálogo": após 2 meses sem acesso (mapeado para "Inativo" no dashboard)
   - "Pré-cadastrado" → "Inativo" após 90 dias sem ativação

2. **Pontuações não refletidas no dashboard** (apuração Whirlpool):
   - Aniversário: 10 pts no mês de aniversário (todos os cargos)
   - Gestor da Informação: 300 pts por envio correto + 300 pts bônus trimestral (mar/jun/set/dez)
   - Gerente: 10% da pontuação da equipe de vendedores
   - Gerente Regional: 5% da pontuação dos vendedores das lojas atendidas

3. **Valor do ponto**: R$ 1,00

4. **Expiração de pontos**: 12 meses de validade, expirando no 2º dia útil de cada mês

5. **Formas de resgate**: Transferência bancária, PIX, cartão, lojas virtuais, pagamento de contas, recargas de celular, vouchers

6. **Prazo de resgate após desligamento**: 90 dias

7. **Prazo de contestação de pontuação**: 7 dias corridos a partir do crédito

---

*Documento gerado em 2026-06-08 com base nas respostas do usuário.*
