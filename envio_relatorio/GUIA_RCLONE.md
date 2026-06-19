# Guia de Configuração do Google Drive (rclone)

Este guia explica como configurar o upload automático dos relatórios para uma pasta no Google Drive.

---

## 1. Abra o terminal Ubuntu (WSL)

Use o mesmo terminal onde você roda os comandos do Python.

---

## 2. Vá até a pasta do projeto

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
```

---

## 3. Rode o script de configuração

```bash
bash setup_rclone.sh
```

---

## 4. Se der erro de `unzip`

Se aparecer a mensagem:

```
None of the supported tools for extracting zip archives (unzip 7z busybox) were found.
```

Instale o `unzip`:

```bash
sudo apt update
sudo apt install -y unzip
```

Depois rode o script de configuração novamente:

```bash
bash setup_rclone.sh
```

---

## 5. Siga o passo a passo do rclone config

Quando o script pedir para pressionar Enter, aperte **Enter**.

Vai abrir o `rclone config`. Responda assim:

| Pergunta | Resposta |
|---|---|
| `n/s/q>` | **n** |
| `name>` | **gdrive** |
| `Storage>` | **drive** |
| `client_id>` | pressione **Enter** (deixe em branco) |
| `client_secret>` | pressione **Enter** (deixe em branco) |
| `scope>` | **1** |
| `root_folder_id>` | pressione **Enter** |
| `service_account_file>` | pressione **Enter** |
| `Edit advanced config?` | **n** |
| `Use auto config?` | **y** |

---

## 6. Faça login no Google

O navegador vai abrir automaticamente.

1. Faça login com a conta do Google.
2. Clique em **Permitir** para autorizar o rclone.
3. Feche a aba do navegador e volte para o terminal.

---

## 7. Pegue o link da pasta no Drive

O script vai criar a pasta `relatorio_semanal` no Drive e mostrar o link, por exemplo:

```bash
Link: https://drive.google.com/drive/folders/1aBcD2EfGhIjKlMnOpQr3StUvWxYz4
```

Copie esse link.

---

## 8. Atualize o `config_email.json`

Abra o arquivo:

```bash
nano config_email.json
```

Substitua:

```json
"link_drive": "https://drive.google.com/drive/folders/SEU_LINK_DO_DRIVE_AQUI"
```

Pelo link que você copiou. Exemplo:

```json
"link_drive": "https://drive.google.com/drive/folders/1aBcD2EfGhIjKlMnOpQr3StUvWxYz4"
```

Salve no nano:

- Pressione **Ctrl + O**
- Pressione **Enter**
- Pressione **Ctrl + X**

---

## 9. Teste o upload

```bash
python3 upload_drive.py
```

Se aparecer a mensagem **"Upload concluído com sucesso"**, tudo está configurado.

A partir da próxima vez que o relatório semanal for gerado, os arquivos serão enviados automaticamente para o Drive.

---

## Problemas comuns

### Escolhi a opção errada no Storage

Se aparecer algo sobre FTP, host, etc., pressione **Ctrl + C** para sair e rode `bash setup_rclone.sh` de novo.

### Navegador não abre

Se você estiver em um ambiente sem navegador, digite **n** em `Use auto config?`. O terminal vai mostrar um link para você copiar e colar no navegador de outro computador.

### Upload falha depois de configurado

Verifique se o link no `config_email.json` está correto e se a pasta `relatorio_semanal` existe no Drive.
