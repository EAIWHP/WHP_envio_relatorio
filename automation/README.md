# 🤖 Automação de Download - Programa Mais Top

Script Python com Playwright para automatizar o download das bases **cadastro** e **acessos** do site [programamaistop.com.br](https://programamaistop.com.br).

---

## 📁 Estrutura

```
automation/
├── README.md              # Este arquivo
├── config.py              # Configuracoes (URLs, caminhos, credenciais)
├── download_bases.py      # Script principal de automacao
├── requirements.txt       # Dependencias Python
├── setup.py               # Script de instalacao
├── .env.example           # Exemplo de arquivo de credenciais
└── screenshots/           # Screenshots de debug (gerado automaticamente)
```

As bases baixadas sao salvas em:
```
../bases_recorrentes/DDMMYY/
cadastro_YYYY-MM-DDTHH_MM_SS.csv
acessos_YYYY-MM-DDTHH_MM_SS.csv
```

---

## 🚀 Instalação

### 1. Instale as dependências

```bash
cd /home/thamiresvieira/projetos/Programa_mais_top/automation
python setup.py
```

Ou manualmente:
```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Configure as credenciais

```bash
cp .env.example .env
```

Edite o arquivo `.env`:
```bash
MAISTOP_USER=seu_usuario
MAISTOP_PASS=sua_senha
```

> 💡 **Dica**: O script também aceita variáveis de ambiente do sistema.
> No Windows: `set MAISTOP_USER=seu_usuario`
> No Linux/Mac: `export MAISTOP_USER=seu_usuario`

---

## ▶️ Uso

### Modo padrão (headless - sem abrir navegador)
```bash
python download_bases.py
```

### Modo com navegador visível (para primeira execução/debug)
```bash
python download_bases.py --visible
```

### Modo com screenshots de cada etapa
```bash
python download_bases.py --visible --screenshot
```

---

## 🔧 Configuração obrigatória antes do primeiro uso

O script vem com **seletores CSS genéricos** que provavelmente **não vão funcionar** no site real. Você precisa atualizá-los inspecionando o site.

### Como descobrir os seletores corretos:

1. Abra o site no Chrome/Edge
2. Aperte **F12** para abrir as DevTools
3. Clique no ícone 🔍 (seletor de elemento)
4. Clique no campo/botão que quer identificar
5. Veja o HTML e copie o seletor CSS

### Seletores que precisam ser ajustados:

No arquivo `download_bases.py`, atualize as listas em `realizar_login()`:

```python
SELECTORES_LOGIN = {
    "campo_usuario": [
        'input[name="username"]',   # <-- ajuste conforme o site
        # ... adicione mais seletores
    ],
    "campo_senha": [
        'input[type="password"]',   # <-- ajuste conforme o site
    ],
    "botao_login": [
        'button[type="submit"]',    # <-- ajuste conforme o site
    ],
}
```

E em `navegar_para_relatorios()`:
```python
SELECTORES_RELATORIOS = {
    "menu_relatorios": [
        'a:has-text("Relatórios")',  # <-- ajuste conforme o site
    ],
}
```

E nos métodos `baixar_cadastro()` e `baixar_acessos()`:
```python
seletores = [
    'a:has-text("Cadastro")',  # <-- ajuste conforme o site
]
```

---

## 📅 Agendamento Automático

### Linux (cron)

Edite o crontab:
```bash
crontab -e
```

Adicione para rodar diariamente às 9h:
```cron
0 9 * * * cd /home/thamiresvieira/projetos/Programa_mais_top/automation && /usr/bin/python3 download_bases.py >> /tmp/maistop_cron.log 2>&1
```

### Windows (Agendador de Tarefas)

1. Abra o "Agendador de Tarefas"
2. Crie uma nova tarefa básica
3. Ação: Iniciar um programa
4. Programa: `python.exe`
5. Argumentos: `download_bases.py`
6. Iniciar em: `C:\caminho\para\Programa_mais_top\automation`

---

## 📋 Fluxo Completo do Script

```
1. Inicia navegador Chromium (headless ou visível)
2. Acessa https://programamaistop.com.br
3. Preenche usuário e senha
4. Clica em Entrar/Login
5. Navega para a seção de Relatórios
6. Clica em download de Cadastro
7. Clica em download de Acessos
8. Move os arquivos para bases_recorrentes/DDMMYY/
9. Encerra navegador
```

---

## ⚠️ Possíveis Problemas

| Problema | Solução |
|----------|---------|
| Site retorna 403/Forbidden | O site pode ter proteção anti-bot (Cloudflare). Tente o modo `--visible`. |
| Seletor não encontrado | Atualize os seletores CSS inspecionando o site real com F12. |
| Download não inicia | O site pode abrir o download em uma nova aba. Ajuste o código para lidar com popups. |
| Timeout no login | O site pode ter carregamento lento. Aumente os timeouts em `config.py`. |
| Navegador não inicia | Execute `playwright install chromium` novamente. |

---

## 🔄 Integração com Power BI

Após o download automático, você ainda precisará:
1. Abrir o `Dashboard_TOP_V7.pbix` no Power BI Desktop
2. Ir em **Transformar Dados** → Atualizar fontes para a nova pasta
3. **Fechar e Aplicar**
4. **Salvar**
5. **Publicar** no Power BI Service

> Em uma versão futura, este processo também pode ser automatizado via Power BI REST API ou um script PowerShell com o Power BI Desktop.

---

## 📝 Log de Execução

O script exibe logs coloridos no terminal. Para salvar em arquivo:
```bash
python download_bases.py 2>&1 | tee automation_$(date +%Y%m%d).log
```
