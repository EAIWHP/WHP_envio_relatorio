# Guia: Configurar Envio do Relatório +TOP pelo E-mail Corporativo

**E-mail corporativo:** `thamires.vieira@eaimkt.com.br`  
**Provedor:** Outlook / Microsoft 365  
**Arquivo de configuração:** `envio_relatorio/config_email.json`  

---

## 1. O que já foi feito

O script do relatório semanal (`gerar_relatorio_semanal.py`) já foi configurado para enviar pelo servidor SMTP do Outlook:

- **Servidor SMTP:** `smtp.office365.com`
- **Porta:** `587`
- **Criptografia:** STARTTLS
- **Remetente:** `thamires.vieira@eaimkt.com.br`
- **Destinatários e CC:** conforme definido no `config_email.json`

O que ainda falta é a **autenticação correta**, porque sua conta tem **MFA (autenticação em 2 fatores)** habilitada.

---

## 2. Por que não funciona com sua senha normal

Com o MFA ativo, o Outlook/Office 365 **não aceita sua senha de login diretamente** em aplicativos que usam SMTP. Isso é uma segurança da Microsoft.

Você precisa de uma **Senha de App** (App Password) ou que o TI habilite uma forma de autenticação alternativa.

---

## 3. O que pedir para o TI

Envie esse texto para o TI:

> **Assunto:** Configuração de envio automático de e-mail via SMTP — thamires.vieira@eaimkt.com.br
>
> Olá,
>
> Preciso configurar um script Python para enviar um relatório semanal automaticamente usando o e-mail corporativo `thamires.vieira@eaimkt.com.br`.
>
> O script usa as seguintes configurações SMTP:
> - Servidor: `smtp.office365.com`
> - Porta: `587`
> - Criptografia: STARTTLS
>
> Como a conta tem MFA habilitado, preciso de uma das alternativas abaixo:
>
> **Opção 1 (preferencial):** Gerar uma **Senha de App** (App Password) para essa conta, exclusiva para esse envio automatizado.
>
> **Opção 2:** Habilitar **SMTP AUTH** na conta e confirmar se posso usar autenticação básica (e-mail + senha/app password).
>
> **Opção 3:** Se a política da empresa não permitir Senha de App nem SMTP AUTH, precisarei de orientação para configurar autenticação moderna (OAuth2) no script.
>
> Também peço para confirmar:
> - Se é necessário liberar algum IP ou rede para envio via SMTP.
> - Se há restrições de SPF/DKIM para envio automatizado por esse e-mail.
>
> Obrigada!

---

## 4. Como gerar a Senha de App você mesma (se a empresa permitir)

Se o TI informar que você pode gerar a senha de app diretamente:

1. Acesse: https://account.microsoft.com/security
2. Faça login com `thamires.vieira@eaimkt.com.br`
3. Vá em **"Segurança" → "Avançado" → "Senhas de app"**
4. Clique em **"Criar nova senha de app"**
5. Dê um nome fácil de identificar, exemplo: `Relatorio_Semanal_TOP`
6. Copie a senha de 16 caracteres gerada (ela aparece apenas uma vez)
7. Cole no arquivo `config_email.json`, no campo `senha_app`

Exemplo:

```json
{
  "smtp_host": "smtp.office365.com",
  "smtp_port": 587,
  "remetente_email": "thamires.vieira@eaimkt.com.br",
  "remetente_nome": "Relatório Semanal +TOP",
  "senha_app": "abcd efgh ijkl mnop",
  "destinatarios": ["thamires.vieira@eaimkt.com.br"],
  "cc": ["viteixethami@gmail.com", "lais.calazans@eaimkt.com.br"],
  "destinatarios_regionais": false,
  "assunto": "Relatório Semanal Programa +TOP",
  "link_drive": "https://eaimkt1.sharepoint.com/:f:/s/DadosPessoais-Sensveis/IgApaQeEhY0ZQJb9ZGpAmXMRAQLbd2EUpCTlbKdusQOGK5Q?e=0Ogbon"
}
```

---

## 5. Como testar após colocar a senha

Abra o terminal no diretório do projeto:

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/envio_relatorio
```

Rode em modo teste (gera o relatório, mas **não envia** o e-mail):

```bash
python3 gerar_relatorio_semanal.py --teste
```

Se não der erro, rode em produção para enviar o e-mail de verdade:

```bash
python3 gerar_relatorio_semanal.py
```

---

## 6. Erros comuns e o que fazer

| Erro | Causa provável | Solução |
|------|---------------|---------|
| `SMTP AUTH extension not supported` ou `authentication failed` | Senha normal com MFA ativo | Usar Senha de App |
| `535 5.7.139 Authentication unsuccessful` | SMTP AUTH desativado pela empresa | Pedir ao TI para habilitar |
| `530 5.7.57 SMTP; Client was not authenticated` | Autenticação moderna obrigatória | Precisa configurar OAuth2 no script |
| E-mail cai no spam | Falta configuração SPF/DKIM | Verificar com TI |

---

## 7. Segurança: protegendo a senha

Hoje a senha fica salva em texto no arquivo `config_email.json`. Para deixar mais seguro no servidor, o script já consegue ler de uma variável de ambiente:

```bash
export SMTP_APP_PASSWORD="sua_senha_de_app"
```

Depois de configurar a senha de app, podemos ajustar o script para usar preferencialmente a variável de ambiente, evitando deixar a senha fixa no arquivo.

---

## 8. Próximos passos resumidos

1. **Aguardar resposta do TI** sobre Senha de App ou SMTP AUTH.
2. **Inserir a senha** no `config_email.json`.
3. **Rodar o teste:** `python3 gerar_relatorio_semanal.py --teste`
4. **Rodar o envio real:** `python3 gerar_relatorio_semanal.py`
5. **(Opcional) Ajustar segurança** para usar variável de ambiente.

---

*Guia gerado em 17/06/2026 para configuração do envio do Relatório Semanal +TOP pelo e-mail corporativo.*
