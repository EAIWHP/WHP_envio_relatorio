# Relatório Semanal Programa +TOP

Este diretório contém o script automático de envio do relatório semanal de cadastros, treinamentos e aceites do Programa +TOP.

## O que o relatório envia

O e-mail semanal contém:

1. **Cadastros** — por regional e por revenda
   - Total de participantes
   - Usuários **com +TOP** (status `Ativo`)
   - Usuários **sem +TOP** (demais status)
   - Percentual de ativos

2. **Treinamentos obrigatórios do mês atual** — por regional e por revenda
   - Detecta automaticamente os 2 cursos obrigatórios do mês e mostra os nomes descritivos (ex: Ar Condicionado Consul, Geladeira Consul Inverse)
   - Participantes que **realizaram ambos os treinamentos** no mês
   - Participantes que **não realizaram**
   - Percentual de realização
   - **Também exibe o desempenho de cada curso individualmente** (por regional e por revenda)

3. **Aceites mensais** — por regional **e por revenda**
   - Participantes que deram aceite no mês
   - Participantes que não deram aceite
   - Percentual de aceite

4. **Evolução semanal**
   - Comparação com a semana anterior
   - Regionais que cresceram ou perderam percentual de ativos

5. **Insights automáticos** no corpo do e-mail, destacando regionais/revendas abaixo da média e **parabenizando a revenda destaque de cada regional** em cadastros, treinamentos e aceites.

6. **Gráficos de acompanhamento** por indicador, embutidos no e-mail.

7. **Relatório completo em Excel anexo** ao e-mail.

## Destinatários

- E-mail principal para a cliente (configurável em `config_email.json`, campo `destinatarios`)
- Cópia para e-mails adicionais (campo `cc`)
- Todos os destinatários recebem o mesmo e-mail com o relatório completo em anexo

## Estrutura esperada das bases

Coloque os arquivos dentro de `envio_relatorio/bases/`:

- `cadastro.xlsx` — base de cadastros
- `Base_treinamentos.xlsx` — base de treinamentos
- `WHP_Aceite_Mensal_*.xlsx` — aceites mensais (pode ser substituído por outro arquivo com abas mensais)
- `cadastro_revenda_loja_regional.xlsx` — detalhamento por participante, revenda e loja
- `emails_regionais.xlsx` — mapeamento Regional → Email (para envio individual)

> O script detecta automaticamente a aba do mês atual nos aceites. Se não encontrar, usa a última aba disponível e registra aviso no log.

## Configuração do e-mail

1. O arquivo `config_email.json` está pré-configurado para envio pelo e-mail corporativo:
   - **Remetente:** `thamires.vieira@eaimkt.com.br`
   - **SMTP:** Outlook/Microsoft 365 (`smtp.office365.com`, porta `587`)

2. Para produção, altere o campo `destinatarios` para o e-mail da cliente (pode manter seu e-mail em `cc` para acompanhar).

3. Configure `link_drive` com o link da pasta do Drive onde será salva a planilha completa.

4. Configure `destinatarios_regionais` como `true` se quiser enviar também para todos os regionais.

5. Preencha `bases/emails_regionais.xlsx` com os e-mails de cada regional.

6. Você precisa gerar uma **Senha de app** do Outlook/Microsoft 365:
   - Acesse https://account.activedirectory.windowsazure.com/AppPasswords.aspx
   - Ou peça ao administrador de TI para criar uma senha de app
   - Gere uma senha para o app `Relatorio +TOP`
   - Copie a senha gerada

7. Escolha uma das formas de informar a senha:

   **Opção A — variável de ambiente (mais seguro):**
   ```bash
   export SMTP_APP_PASSWORD="xxxx xxxx xxxx xxxx"
   ```

   **Opção B — arquivo `config_email.json`:**
   ```bash
   nano config_email.json
   # altere o campo "senha_app" para a senha de app
   ```

> ⚠️ Nunca use a senha normal do e-mail. Use obrigatoriamente a "Senha de app".

## Configuração do upload automático no Google Drive

O relatório pode enviar os arquivos gerados automaticamente para uma pasta no Google Drive usando o `rclone`.

1. Execute o script de configuração (apenas uma vez):
   ```bash
   bash setup_rclone.sh
   ```

2. Siga as instruções na tela para fazer login com a conta do Google e criar o remote `gdrive`.

3. O script criará a pasta `relatorio_semanal` no Drive e exibirá o link. Cole esse link no campo `link_drive` do `config_email.json`.

4. Teste o upload:
   ```bash
   python3 upload_drive.py
   ```

A partir da próxima execução do relatório, os arquivos serão enviados automaticamente. Se o upload falhar, o relatório ainda é enviado por e-mail e o erro é registrado no log.

---

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio

# Modo teste: gera relatório Excel e HTML local, mas não envia e-mail
python3 gerar_relatorio_semanal.py --teste

# Envio de teste (gera e envia o e-mail)
python3 gerar_relatorio_semanal.py
```

## Agendamento automático (cron)

Para enviar toda **segunda-feira às 11:00**, edite o cron do seu usuário:

```bash
crontab -e
```

Adicione a linha abaixo. **Importante:** informe a senha de app via variável de ambiente para não deixá-la exposta no comando:

```cron
0 11 * * 1 export SMTP_APP_PASSWORD="xxxx xxxx xxxx xxxx"; cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio && /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/.venv/bin/python3 gerar_relatorio_semanal.py >> /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/logs/cron.log 2>&1
```

Ou use o instalador automático:

```bash
./instalar_cron.sh
```

## Saídas geradas

- `relatorios_gerados/relatorio_top_YYYYMMDD.xlsx` — relatório completo em Excel
- `relatorios_gerados/relatorio_top_YYYYMMDD.html` — cópia do e-mail em HTML
- `relatorios_gerados/relatorio_top_YYYYMMDD_REGIONAL.xlsx` — relatório individual por regional
- `logs/relatorio_YYYYMMDD.log` — log de execução
- `snapshots/snapshot_YYYYMMDD.json` — snapshot dos indicadores para comparação semanal

A planilha completa contém as abas:

- `Cadastro_Regional` / `Cadastro_Revenda`
- `Treinamento_Regional` / `Treinamento_Revenda` (ambos os cursos)
- `<Nome do Curso>_Reg` / `<Nome do Curso>_Rev` (um par de abas para cada curso obrigatório do mês, com o nome real do treinamento)
- `Aceite_Regional`
- `Detalhamento` (inclui colunas de status e data de conclusão dos cursos obrigatórios)
- `Resumo`

## Guia passo a passo para não programadores

Consulte o arquivo `README_GUIA.md` para instruções detalhadas de instalação, configuração e rotina semanal.

## Regras de negócio aplicadas

- CPF/CNPJ sempre tratado como texto, preservando zeros à esquerda
- Status de cadastro `Inativo`, `Bloqueado` e `Reprovado` são desconsiderados nas métricas
- Filiais com sufixo `B` são agrupadas na revenda principal
- Revendas de teste/excluídas (`EAI`, `Whirlpool`, `Novo Mundo`, `ELETROMÓVEIS MARTINELLO`) são desconsideradas
- Treinamentos consideram os 2 cursos obrigatórios do mês, tanto combinados quanto individualmente
- Nomes dos cursos são mapeados por SKU no dicionário `NOMES_CURSOS`
- SKU do produto é exibido junto ao nome do conteúdo do curso
- Base de cálculo para treinamentos e aceites: participantes com status `Ativo`
- Metas ideais: Cadastro 85%, Treinamentos 70%, Aceites 70%
- Destaques por regional: melhor revenda em cadastros, treinamentos e aceites é parabenizada no e-mail, desde que tenha atingido a meta ideal

## Suporte

Em caso de erro, verifique:

1. Se as bases estão no caminho correto (`envio_relatorio/bases/`)
2. Se `config_email.json` está preenchido corretamente
3. Os logs em `logs/`
4. O guia em `README_GUIA.md`
