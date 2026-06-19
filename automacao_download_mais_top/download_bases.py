#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automação de teste para download de bases do Programa +TOP.

Este script faz login no site https://programamaistop.com.br e tenta
baixar os arquivos necessários para atualização do relatório semanal.

Uso:
    MAIS_TOP_USUARIO=seu_usuario MAIS_TOP_SENHA=sua_senha python3 download_bases.py

Variáveis de ambiente:
    MAIS_TOP_USUARIO  - usuário de login no site
    MAIS_TOP_SENHA    - senha de login no site
    HEADLESS          - "false" para ver o navegador (padrão: "true")
    DOWNLOAD_DIR      - pasta de destino dos downloads (padrão: ./downloads)
"""

import os
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


URL_LOGIN = "https://programamaistop.com.br/#!/login"
DEFAULT_DOWNLOAD_DIR = Path(__file__).resolve().parent / "downloads"


def get_env_or_die(name):
    value = os.environ.get(name)
    if not value:
        print(f"Erro: variável de ambiente {name} não definida.")
        sys.exit(1)
    return value


def main():
    usuario = get_env_or_die("MAIS_TOP_USUARIO")
    senha = get_env_or_die("MAIS_TOP_SENHA")
    headless = os.environ.get("HEADLESS", "true").lower() != "false"
    download_dir = Path(os.environ.get("DOWNLOAD_DIR", DEFAULT_DOWNLOAD_DIR))
    download_dir.mkdir(parents=True, exist_ok=True)

    print(f"Iniciando automação...")
    print(f"URL: {URL_LOGIN}")
    print(f"Download: {download_dir}")
    print(f"Headless: {headless}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, args=["--no-sandbox"])
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = context.new_page()

        try:
            print("Abrindo página de login...")
            page.goto(URL_LOGIN, wait_until="networkidle", timeout=60000)
            page.screenshot(path=str(download_dir / "01_login_page.png"))

            print("Procurando campos de login...")
            # Estratégia 1: procura por campos comuns
            email_selectors = [
                'input[type="email"]',
                'input[name="email"]',
                'input[name="usuario"]',
                'input[name="user"]',
                'input[id*="email"]',
                'input[id*="usuario"]',
                'input[placeholder*="email" i]',
                'input[placeholder*="usuário" i]',
                'input[placeholder*="usuario" i]',
                'input[ng-model*="email"]',
                'input[ng-model*="usuario"]',
            ]
            senha_selectors = [
                'input[type="password"]',
                'input[name="senha"]',
                'input[name="password"]',
                'input[id*="senha"]',
                'input[id*="password"]',
                'input[placeholder*="senha" i]',
                'input[placeholder*="password" i]',
                'input[ng-model*="senha"]',
                'input[ng-model*="password"]',
            ]

            campo_usuario = None
            for selector in email_selectors:
                try:
                    campo_usuario = page.locator(selector).first
                    if campo_usuario.is_visible(timeout=2000):
                        print(f"Campo de usuário encontrado: {selector}")
                        break
                except Exception:
                    continue

            campo_senha = None
            for selector in senha_selectors:
                try:
                    campo_senha = page.locator(selector).first
                    if campo_senha.is_visible(timeout=2000):
                        print(f"Campo de senha encontrado: {selector}")
                        break
                except Exception:
                    continue

            if not campo_usuario or not campo_senha:
                print("Não foi possível encontrar os campos de login.")
                page.screenshot(path=str(download_dir / "02_login_not_found.png"))
                browser.close()
                sys.exit(1)

            print("Preenchendo credenciais...")
            campo_usuario.fill(usuario)
            campo_senha.fill(senha)
            page.screenshot(path=str(download_dir / "02_credentials_filled.png"))

            print("Clicando no botão de login...")
            # Procura botão de login
            botao_selectors = [
                'button[type="submit"]',
                'button:has-text("Entrar")',
                'button:has-text("Login")',
                'button:has-text("Acessar")',
                'input[type="submit"]',
                'a:has-text("Entrar")',
            ]
            botao_login = None
            for selector in botao_selectors:
                try:
                    botao_login = page.locator(selector).first
                    if botao_login.is_visible(timeout=2000):
                        print(f"Botão de login encontrado: {selector}")
                        break
                except Exception:
                    continue

            if not botao_login:
                print("Botão de login não encontrado. Tentando pressionar Enter.")
                campo_senha.press("Enter")
            else:
                botao_login.click()

            # Aguarda carregamento após login
            print("Aguardando carregamento após login...")
            time.sleep(5)
            page.screenshot(path=str(download_dir / "03_after_login.png"))

            current_url = page.url
            print(f"URL atual: {current_url}")

            # Procura por links de download comuns
            print("Procurando arquivos para download...")
            extensoes = [".xlsx", ".csv", ".xls", ".zip"]
            links = page.locator("a").all()
            downloads_encontrados = []

            for link in links:
                try:
                    href = link.get_attribute("href") or ""
                    text = link.inner_text().strip()
                    if any(ext in href.lower() for ext in extensoes) or any(
                        palavra in text.lower()
                        for palavra in ["download", "exportar", "baixar", "excel"]
                    ):
                        downloads_encontrados.append({"text": text, "href": href})
                        print(f"Encontrado: {text} -> {href}")
                except Exception:
                    continue

            print(f"Total de links de download encontrados: {len(downloads_encontrados)}")

            # Lista todos os botões e links visíveis para análise
            botoes = page.locator("button").all()
            print(f"\nBotoes visiveis na pagina ({len(botoes)}):")
            for i, botao in enumerate(botoes[:20]):
                try:
                    print(f"  {i+1}. {botao.inner_text().strip()}")
                except Exception:
                    pass

            print("\nAutomação de teste finalizada.")
            print(f"Screenshots salvos em: {download_dir}")

        except PlaywrightTimeout as e:
            print(f"Timeout: {e}")
            page.screenshot(path=str(download_dir / "erro_timeout.png"))
        except Exception as e:
            print(f"Erro: {e}")
            try:
                page.screenshot(path=str(download_dir / "erro.png"))
            except Exception:
                pass
        finally:
            browser.close()


if __name__ == "__main__":
    main()
