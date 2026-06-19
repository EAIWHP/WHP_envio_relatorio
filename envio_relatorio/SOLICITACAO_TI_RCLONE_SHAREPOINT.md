# Solicitação de Aprovação — Aplicativo rclone no SharePoint

**Para:** Administrador de TI / Suporte Microsoft 365  
**Assunto:** Solicitação de autorização do aplicativo rclone para upload automático no SharePoint

---

## Resumo

Solicito autorização para o aplicativo **rclone** acessar o site do SharePoint abaixo, para fins de upload automático de relatórios semanais gerados por script Python.

---

## Detalhes técnicos

| Item | Informação |
|---|---|
| Aplicativo | rclone |
| Tipo de acesso | OneDrive / SharePoint (OAuth) |
| Conta do usuário | thamires.vieira@eaimkt.com.br |
| Site do SharePoint | https://eaimkt1.sharepoint.com/sites/DadosPessoais-Sensveis |
| Pasta destino | relatorio_semanal |
| Finalidade | Upload automático de relatórios gerados semanalmente |

---

## O que o rclone precisa fazer

- Acessar o site do SharePoint indicado
- Criar a pasta `relatorio_semanal` (caso não exista)
- Enviar arquivos Excel (.xlsx) e HTML para dentro dessa pasta

---

## Mensagem de erro atual

> "rclone precisa de permissão para acesso aos recursos em sua organização que apenas o administrador pode conceder."

---

## Como autorizar

### Opção 1 — Via portal de administração do Microsoft Entra (Azure AD)

1. Acesse o portal de administração do Microsoft Entra / Azure Active Directory
2. Vá em **Aplicativos empresariais** ou **Enterprise applications**
3. Localize o aplicativo **rclone**
4. Conceda consentimento administrativo (**Grant admin consent**)
5. Garanta que o usuário `thamires.vieira@eaimkt.com.br` tenha permissão para usar o aplicativo

### Opção 2 — No momento do login OAuth

Durante a configuração do rclone, quando a tela de permissão aparecer, o administrador deve:

1. Fazer login com uma conta administrativa
2. Marcar a opção **"Consentir em nome da organização"**
3. Clicar em **Aceitar**

---

## Informações adicionais

- **Documentação oficial do rclone para OneDrive/SharePoint:** https://rclone.org/onedrive/
- **Tipo de permissão necessária:** Delegada / Delegated (leitura e escrita em document libraries do SharePoint)
- **Cliente OAuth:** rclone usa o cliente OAuth padrão do projeto (aplicativo de código aberto)

---

## Contato

Caso seja necessário mais informações, podemos marcar uma rápida chamada para explicar o fluxo.

Att.,  
Thamires Vieira
