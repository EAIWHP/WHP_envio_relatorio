# RESUMO — Sessão de 12/06/2026

## O que foi feito

Criamos um script que gera e envia **semanalmente** o relatório do Programa +TOP por e-mail.

O relatório contém:
- Cadastros (ativos / não ativos) por regional e por revenda
- Treinamentos obrigatórios do mês por regional e por revenda
- Aceites mensais por regional
- Comparação com a semana anterior
- Insights prontos no corpo do e-mail
- Gráficos e tabelas
- Link para o Drive com planilha completa

## Para quem o e-mail é enviado

1. Cliente principal
2. Todos os regionais (e-mail geral com dados de todos)
3. Cada regional individualmente (e-mail só com os dados daquela regional)

## Arquivos criados / alterados

| Arquivo | O que é |
|---------|---------|
| `gerar_relatorio_semanal.py` | Script principal |
| `config_email.json` | Configuração de e-mail, destinatários e link do Drive |
| `config_email.json.example` | Exemplo de configuração |
| `bases/emails_regionais.xlsx` | E-mails de cada regional |
| `instalar_cron.sh` | Instala o envio automático toda segunda às 11h |
| `README.md` | Documentação técnica |
| `README_GUIA.md` | Guia passo a passo para não programadores |
| `RESUMO_SESSAO.md` | Este arquivo |

## Funcionalidades principais

- Detecta automaticamente os 2 treinamentos obrigatórios do mês
- Agrupa filiais "B" na revenda principal
- Desconsidera revendas de teste/excluídas
- Gera Excel completo + Excel por regional
- Salva snapshot para comparar semana com semana
- Gera HTML do e-mail para conferência

## Aba "Detalhamento" do Excel

A planilha completa (a que vai no Drive) tem uma aba chamada **Detalhamento** com:

- Regional
- Revenda
- CNPJ
- Nome
- Cargo
- Status
- Cidade
- UF
- Bairro loja
- Nome loja

## O que você precisa preencher antes de ligar o automático

### 1. `config_email.json`

Abra o arquivo e preencha:
- `senha_app` → senha de app do Gmail
- `destinatarios` → e-mail da cliente
- `cc` → seu e-mail
- `link_drive` → link da pasta do Drive

Como gerar senha de app do Gmail:
1. Acesse https://myaccount.google.com/apppasswords
2. Login com o e-mail que enviará
3. Escolha "Outro" e nomeie "Relatorio +TOP"
4. Copie a senha de 16 caracteres

### 2. `bases/emails_regionais.xlsx`

Preencha o e-mail de cada regional:

| Regional | Email |
|----------|-------|
| CENTRO NORTE | ... |
| COMPRA DIRETA | ... |
| NORDESTE | ... |
| SUDESTE | ... |
| SUL | ... |

### 3. Testar

Rode no terminal:

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
source .venv/bin/activate
python3 gerar_relatorio_semanal.py --teste
```

Isso gera os arquivos sem enviar e-mail.

### 4. Instalar o envio automático

```bash
./instalar_cron.sh
```

Isso agenda o envio toda **segunda-feira às 11:00**.

## Comandos úteis

```bash
# Entrar na pasta
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio

# Ativar ambiente virtual
source .venv/bin/activate

# Testar sem enviar e-mail
python3 gerar_relatorio_semanal.py --teste

# Enviar de verdade
python3 gerar_relatorio_semanal.py

# Ver agendamento
crontab -l
```

## Observações

- Os treinamentos de junho/2026 aparecem como `CZD12` e `BRM46` porque são os códigos da base. Se quiser nomes descritivos (ex: "Ar Condicionado Consul"), é só preencher o dicionário `NOMES_CURSOS` no início do script.
- A base `cadastro_revenda_loja_regional.xlsx` foi substituída pela versão correta do projeto principal.
- Toda a documentação detalhada está em `README_GUIA.md`.

## Próximos passos sugeridos

1. Preencher `config_email.json`
2. Preencher `bases/emails_regionais.xlsx`
3. Rodar `--teste` para conferir os números
4. Configurar o cron com `./instalar_cron.sh`
5. Subir o Excel gerado para a pasta do Drive
