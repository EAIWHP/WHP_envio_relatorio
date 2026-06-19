# Programa +TOP — Relatório Semanal

Automação do relatório semanal do **Programa +TOP Whirlpool**, com geração de HTML/e-mail e Excel de detalhamento.

## 📁 Estrutura principal

```
Programa_mais_top/
├── envio_relatorio/
│   ├── gerar_relatorio_semanal.py   # Script principal
│   ├── config_email.example.json     # Modelo de configuração de e-mail
│   ├── bases/                        # Bases de entrada (não versionadas)
│   ├── relatorios_gerados/           # HTML/XLSX gerados (não versionados)
│   ├── logs/                         # Logs de execução (não versionados)
│   └── snapshots/                    # Snapshots semanais (não versionados)
├── CLAUDE.md                         # Documentação do projeto
└── README.md                         # Este arquivo
```

## 🚀 Como usar

1. Copie o arquivo de configuração:
   ```bash
   cd envio_relatorio
   cp config_email.json.example config_email.json
   ```

2. Edite o `config_email.json` com a senha de app e destinatários corretos.

3. Certifique-se de que as bases estão na pasta `envio_relatorio/bases/`:
   - `cadastro.xlsx`
   - `Base_treinamentos.xlsx`
   - `WHP_Aceite_Mensal_*.xlsx`
   - `emails_regionais.xlsx` (opcional)

4. Execute em modo de teste (não envia e-mail):
   ```bash
   python3 gerar_relatorio_semanal.py --teste
   ```

5. Execute em produção:
   ```bash
   python3 gerar_relatorio_semanal.py
   ```

## 🌿 Fluxo de versionamento

Este projeto segue um fluxo com branches protegidas:

- **`main`** → código estável e aprovado para produção
- **`develop`** → desenvolvimento e testes
- **`feature/nome-da-alteracao`** → cada nova funcionalidade ou ajuste

Nunca faça alterações diretamente na `main`. Sempre crie uma branch a partir da `develop` e abra uma Pull Request.

## 🔒 Dados sensíveis

- `config_email.json` está no `.gitignore` e **nunca deve ser commitado**.
- Bases `.xlsx`/`.csv` também estão no `.gitignore` por conterem dados pessoais.

## 🛠️ Dependências

Principais bibliotecas utilizadas:

- pandas
- openpyxl
- matplotlib
- Pillow

Para instalar:

```bash
pip install pandas openpyxl matplotlib Pillow
```

## 📅 Agendamento sugerido

Executar toda segunda-feira às 11:00 via cron:

```cron
0 11 * * 1 cd /caminho/do/projeto/Programa_mais_top/envio_relatorio && python3 gerar_relatorio_semanal.py >> cron.log 2>&1
```

## 🔒 Proteção da branch `main`

A branch `main` está protegida no GitHub. Isso significa que:

- Não é possível fazer push direto na `main`
- Toda alteração deve passar por uma Pull Request
- O desenvolvimento acontece na branch `develop` ou em branches `feature/*`

Isso garante que o código em produção seja sempre revisado e estável.

## 📧 Contato

Relatório desenvolvido e mantido por Thamires Vieira.
