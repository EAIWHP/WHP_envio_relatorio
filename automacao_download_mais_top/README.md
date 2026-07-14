# Automação de Download — Programa +TOP

Script de teste para fazer login no site https://programamaistop.com.br e baixar as bases necessárias para o relatório semanal.

## ⚠️ Aviso

Este é um projeto separado de teste. Ele **não altera** o projeto `envio_relatorio` que já está funcionando.

## 🔧 Instalação

```bash
cd automacao_download_mais_top
pip install -r requirements.txt
playwright install chromium
```

## 🚀 Como usar

### Localmente

```bash
export MAIS_TOP_USUARIO="seu_usuario"
export MAIS_TOP_SENHA="sua_senha"
export HEADLESS="false"  # opcional: abre o navegador para ver
python3 download_bases.py
```

### No GitHub Actions

O workflow `.github/workflows/teste_download_mais_top.yml` pode ser executado manualmente com os secrets:
- `MAIS_TOP_USUARIO`
- `MAIS_TOP_SENHA`

## 📁 Saída

- Screenshots são salvos em `downloads/`
- Arquivos baixados também vão para `downloads/`

## 📝 Status

Em fase de teste. O script atual faz login e mapeia elementos da página. A lógica de download precisa ser ajustada conforme a estrutura real do site.
