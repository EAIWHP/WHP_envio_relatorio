# 📘 Guia Passo a Passo — Relatório Semanal Programa +TOP

**Para quem é este guia:** qualquer pessoa que precise instalar, configurar e manter o envio do relatório semanal, mesmo sem saber programação.

**Tempo estimado:** 20 minutos na primeira vez; 5 minutos por semana depois.

---

## 1. O que este programa faz?

Toda semana, automaticamente, ele:

1. Lê as bases de cadastro, treinamentos e aceites do Programa +TOP.
2. Calcula indicadores por regional e por revenda.
3. Calcula os treinamentos obrigatórios do mês, tanto combinados quanto por curso individual.
4. Compara com a semana anterior.
5. Monta um e-mail bonito com gráficos, tabelas e insights prontos.
6. Destaca a melhor revenda de cada regional, parabenizando no corpo do e-mail.
7. Envia o e-mail para:
   - A cliente (e-mail principal)
   - Todos os regionais (e-mail geral, com dados de todas as regionais)
   - Cada regional individualmente (e-mail apenas com os dados daquela regional)
8. Envia os arquivos gerados automaticamente para uma pasta no Google Drive.
9. Salva uma planilha Excel completa para consulta.

---

## 2. O que você precisa ter antes de começar?

- Computador com **Windows + WSL** ou **Linux**.
- **Python 3** instalado.
- Acesso ao e-mail que fará o envio (Outlook/Microsoft 365 — thamires.vieira@eaimkt.com.br).
- As bases atualizadas dentro da pasta `envio_relatorio/bases/`.
- O link da pasta do **Drive** onde será salva a versão completa do relatório.

---

## 3. Instalação pela primeira vez

### 3.1. Abra o terminal

- No WSL: abra o aplicativo **Ubuntu** ou **Terminal**.
- No Linux: abra o **Terminal**.

### 3.2. Vá até a pasta do projeto

Digite o comando abaixo e pressione **Enter**:

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
```

> 💡 **Dica:** você pode copiar o comando e colar no terminal com o botão direito do mouse.

### 3.3. Crie o ambiente virtual

O ambiente virtual é uma "caixinha" isolada onde instalamos as bibliotecas do programa, sem bagunçar o computador.

```bash
python3 -m venv .venv
```

Se der erro, pode ser que o Python não tenha o módulo `venv`. Nesse caso, instale com:

```bash
sudo apt update
sudo apt install -y python3-venv python3-pip
```

> Será solicitada a senha do computador.

### 3.4. Ative o ambiente virtual

```bash
source .venv/bin/activate
```

Você verá algo como `(.venv)` no início da linha do terminal. Isso significa que o ambiente virtual está ligado.

### 3.5. Instale as bibliotecas necessárias

Ainda dentro da pasta `envio_relatorio`, digite:

```bash
pip install pandas openpyxl matplotlib
```

Espere terminar. Você verá várias linhas de instalação. Se aparecer algo como `Successfully installed ...`, deu certo.

### 3.6. Teste se o script funciona

```bash
python3 gerar_relatorio_semanal.py --teste
```

Se tudo estiver certo, você verá mensagens de "INFO" e, no final:

```
Modo teste: relatório gerado em ...
```

Isso significa que o relatório foi gerado, mas nenhum e-mail foi enviado.

---

## 4. Configurar o envio de e-mail

### 4.1. Abra o arquivo de configuração

```bash
nano config_email.json
```

Ou abra pelo Explorador de Arquivos: `envio_relatorio/config_email.json`.

### 4.2. Preencha os campos

```json
{
  "smtp_host": "smtp.office365.com",
  "smtp_port": 587,
  "remetente_email": "thamires.vieira@eaimkt.com.br",
  "remetente_nome": "Relatório Semanal +TOP",
  "senha_app": "SUA_SENHA_DE_APP_AQUI",
  "destinatarios": [
    "email.da.cliente@exemplo.com"
  ],
  "cc": [
    "thamires.vieira@eaimkt.com.br"
  ],
  "destinatarios_regionais": true,
  "assunto": "Relatório Semanal Programa +TOP",
  "link_drive": "https://drive.google.com/drive/folders/SEU_LINK_DO_DRIVE_AQUI"
}
```

| Campo | O que colocar |
|-------|---------------|
| `remetente_email` | O e-mail que vai enviar os relatórios (thamires.vieira@eaimkt.com.br). |
| `senha_app` | A "Senha de app" do Outlook/Microsoft 365 (veja abaixo como gerar). |
| `destinatarios` | E-mail principal da cliente. |
| `cc` | Seu e-mail e/ou outros e-mails que devem receber cópia. |
| `destinatarios_regionais` | `true` para enviar também para todos os regionais. `false` para enviar só para a cliente. |
| `link_drive` | Link da pasta do Drive onde ficará a planilha completa. |

> ⚠️ **NUNCA use a senha normal do e-mail.** Sempre use a "Senha de app".

### 4.3. Como gerar a Senha de app do Outlook / Microsoft 365

1. Acesse: https://account.activedirectory.windowsazure.com/AppPasswords.aspx
2. Faça login com o e-mail **thamires.vieira@eaimkt.com.br**.
3. Clique em **"Novo"** ou **"Criar senha de app"**.
4. Dê um nome: `Relatorio +TOP`.
5. Copie a senha gerada.
6. Cole no campo `senha_app` do `config_email.json`.

> Se o link acima não funcionar, peça ao administrador de TI da empresa para criar uma senha de app no portal de administração do Microsoft 365.

### 4.4. Configure os e-mails dos regionais

Abra ou edite o arquivo:

```bash
nano bases/emails_regionais.xlsx
```

> No Windows, é mais fácil abrir pelo Excel.

Preencha com o nome da regional e o e-mail de cada uma:

| Regional | Email |
|----------|-------|
| CENTRO NORTE | regional.centronorte@exemplo.com |
| COMPRA DIRETA | regional.compradireta@exemplo.com |
| NORDESTE | regional.nordeste@exemplo.com |
| SUDESTE | regional.sudeste@exemplo.com |
| SUL | regional.sul@exemplo.com |

Se alguma regional não tiver e-mail, deixe em branco. O programa pula automaticamente.

---

## 5. Configurar o upload automático no Google Drive

O relatório agora pode enviar os arquivos automaticamente para uma pasta no Google Drive.

### 5.1. Configure o rclone (apenas uma vez)

No terminal, dentro da pasta `envio_relatorio`, execute:

```bash
bash setup_rclone.sh
```

Siga as instruções na tela. O script vai:
1. Instalar o `rclone` (se não estiver instalado).
2. Abrir uma configuração onde você faz login com a conta do Google.
3. Criar a pasta `relatorio_semanal` no Drive.
4. Mostrar o link da pasta para você colar no `config_email.json`.

### 5.2. Atualize o link no `config_email.json`

Substitua o valor de `link_drive` pelo link da pasta criada:

```json
"link_drive": "https://drive.google.com/drive/folders/1aBcD2EfGhIjKlMnOpQr3StUvWxYz4"
```

### 5.3. Teste o upload

```bash
python3 upload_drive.py
```

Se aparecer "Upload concluído com sucesso", tudo está certo.

A planilha completa enviada para o Drive contém as seguintes abas:

- **Cadastro_Regional** — resumo de cadastros por regional
- **Cadastro_Revenda** — resumo de cadastros por revenda
- **Treinamento_Regional** — resumo de treinamentos obrigatórios (ambos os cursos) por regional
- **Treinamento_Revenda** — resumo de treinamentos obrigatórios (ambos os cursos) por revenda
- **Trein_Curso1_Reg** — resumo do 1º curso obrigatório por regional
- **Trein_Curso1_Rev** — resumo do 1º curso obrigatório por revenda
- **Trein_Curso2_Reg** — resumo do 2º curso obrigatório por regional
- **Trein_Curso2_Rev** — resumo do 2º curso obrigatório por revenda
- **Aceite_Regional** — resumo de aceites por regional
- **Detalhamento** — lista completa dos participantes com: Regional, Revenda, CNPJ, Nome, Cargo, Status, Cidade, UF, Bairro loja, Nome loja
- **Resumo** — indicadores gerais do programa

A aba **Detalhamento** é a que os regionais usam para fazer a análise completa por loja e por vendedor.

Os arquivos enviados automaticamente para o Drive são:

```
relatorios_gerados/relatorio_top_YYYYMMDD.xlsx
relatorios_gerados/relatorio_top_YYYYMMDD_REGIONAL.xlsx
relatorios_gerados/relatorio_top_YYYYMMDD.html
```

> 💡 **Fallback manual:** se o rclone não estiver configurado, você pode continuar enviando o arquivo manualmente. O relatório será enviado por e-mail normalmente e apenas avisará no log que o upload automático não foi possível.

---

---

## 6. Agendar o envio automático (toda semana)

### 6.1. Escolha o dia e horário

O exemplo abaixo envia toda **segunda-feira às 11:00**.

Se quiser outro dia, troque o número:

| Número | Dia |
|--------|-----|
| 0 | Domingo |
| 1 | Segunda |
| 2 | Terça |
| 3 | Quarta |
| 4 | Quinta |
| 5 | Sexta |
| 6 | Sábado |

### 6.2. Abra o agendador do sistema

```bash
crontab -e
```

Na primeira vez, pode pedir para escolher um editor. Escolha o **nano** (opção 1, geralmente).

### 6.3. Adicione a linha de agendamento

Role até o final do arquivo e cole esta linha:

```cron
0 11 * * 1 cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio && /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/.venv/bin/python3 gerar_relatorio_semanal.py >> /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio/logs/cron.log 2>&1
```

> 💡 Se você não criou o ambiente virtual, use `python3` no lugar do caminho completo do `.venv/bin/python3`.

Salve o arquivo:

- No nano: pressione `Ctrl + O`, depois `Enter`, depois `Ctrl + X`.

### 6.4. Confira se o agendamento foi salvo

```bash
crontab -l
```

A linha que você colou deve aparecer.

---

## 7. Rotina semanal

Siga estes passos toda semana, antes do horário agendado:

### 7.1. Atualize as bases

Certifique-se de que os arquivos abaixo estão atualizados dentro de `envio_relatorio/bases/`:

- `cadastro.xlsx`
- `Base_treinamentos.xlsx`
- `WHP_Aceite_Mensal_*.xlsx` (ou o arquivo de aceites do mês)
- `cadastro_revenda_loja_regional.xlsx` — base com dados detalhados por loja/revenda/regional

> Dica: mantenha sempre o nome do arquivo de aceites começando com `WHP_Aceite_Mensal_`.

### 7.2. Verifique se há nova aba de aceites

O programa detecta automaticamente a aba do mês atual. Se não encontrar, ele usa a última aba disponível e avisa no log.

### 7.3. Faça um teste manual (opcional, mas recomendado)

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
source .venv/bin/activate
python3 gerar_relatorio_semanal.py --teste
```

Isso gera o relatório sem enviar e-mail. Verifique se os números estão certos.

### 7.4. Envie manualmente, se necessário

Se precisar enviar fora do horário agendado:

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
source .venv/bin/activate
python3 gerar_relatorio_semanal.py
```

> ⚠️ Isso envia e-mails de verdade. Use com cuidado.

### 7.5. Verifique o upload no Drive

Se o rclone estiver configurado, o upload dos arquivos para o Drive acontece automaticamente.

Confira na pasta do Drive se os arquivos do dia apareceram:

```
relatorio_top_YYYYMMDD.xlsx
relatorio_top_YYYYMMDD_REGIONAL.xlsx
relatorio_top_YYYYMMDD.html
```

Se o upload automático não funcionar, o log em `logs/relatorio_YYYYMMDD.log` mostrará o aviso. Nesse caso, envie o Excel principal manualmente para a pasta do Drive.

> 📌 Veja a seção 5 para configurar ou testar o upload automático.

---

## 8. Onde verificar se deu certo?

### 8.1. Logs

Os logs ficam em:

```
envio_relatorio/logs/relatorio_YYYYMMDD.log
```

Abra o arquivo do dia para ver mensagens de erro ou confirmação de envio.

### 8.2. Relatórios gerados

Ficam em:

```
envio_relatorio/relatorios_gerados/
```

- `relatorio_top_YYYYMMDD.xlsx` — relatório completo
- `relatorio_top_YYYYMMDD.html` — cópia do e-mail
- `relatorio_top_YYYYMMDD_REGIONAL.xlsx` — relatório por regional

### 8.3. Snapshots

Ficam em:

```
envio_relatorio/snapshots/snapshot_YYYYMMDD.json
```

São usados para comparar uma semana com a outra.

---

## 9. Problemas comuns

### O e-mail não chegou

1. Verifique o log em `logs/relatorio_YYYYMMDD.log`.
2. Confira se a `senha_app` está correta no `config_email.json`.
3. Confira se os e-mails dos destinatários estão corretos.
4. Verifique se o remetente não está na caixa de spam.

### Os números do relatório estão errados

1. Confira se as bases estão atualizadas.
2. Verifique se o arquivo de aceites tem a aba do mês correto.
3. Confira se os CPFs estão como texto nas bases.

### Apareceu "Aba de aceites não encontrada"

O programa usou a última aba disponível. Adicione a aba do mês atual no arquivo de aceites ou renomeie para o padrão `JUN_2026`.

### O script não roda

1. Verifique se o ambiente virtual está ativado (`source .venv/bin/activate`).
2. Verifique se as bibliotecas estão instaladas (`pip install pandas openpyxl matplotlib`).
3. Leia a mensagem de erro no terminal.

---

## 10. Resumo dos comandos mais usados

```bash
# Entrar na pasta
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio

# Ativar ambiente virtual
source .venv/bin/activate

# Testar (não envia e-mail)
python3 gerar_relatorio_semanal.py --teste

# Enviar de verdade
python3 gerar_relatorio_semanal.py

# Ver agendamento
crontab -l

# Editar agendamento
crontab -e
```

---

## 11. Quem criou e mantém?

- **Responsáveis técnicos:** Matheus Carmo; Thamires Vieira
- **Responsáveis de negócio:** Fauzi Naufal; Matheus Carmo

Em caso de dúvidas, consulte as memórias do projeto em:

```
/home/thamiresvieira/.claude/projects/-home-thamiresvieira-projetos/memory/
```
