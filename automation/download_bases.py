#!/usr/bin/env python3
"""
Automação de download de bases do Programa Mais Top.

Fluxo:
    1. Acessa https://programamaistop.com.br
    2. Realiza login
    3. Navega até a seção de relatórios
    4. Faz download de "cadastro" e "acessos"
    5. Salva na pasta bases_recorrentes/DDMMYY/

Uso:
    # Modo headless (sem abrir navegador - recomendado para cron/agendador)
    python download_bases.py

    # Modo com navegador visível (para depuração)
    python download_bases.py --visible

    # Usar credenciais via variáveis de ambiente
    export MAISTOP_USER="seu_usuario"
    export MAISTOP_PASS="sua_senha"
    python download_bases.py
"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from playwright.sync_api import (
    Browser,
    BrowserContext,
    Download,
    Page,
    Playwright,
    sync_playwright,
    TimeoutError as PlaywrightTimeout,
)

import config

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("maistop_automation")


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
DOWNLOAD_WAIT_MS = 2_000  # Espera adicional após iniciar download


class MaisTopAutomation:
    """Orquestra o download automatizado das bases do Programa Mais Top."""

    def __init__(self, headless: bool = True) -> None:
        self.headless = headless
        self.playwright: Optional[Playwright] = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.download_dir: Optional[Path] = None

    # -----------------------------------------------------------------------
    # Ciclo de vida
    # -----------------------------------------------------------------------
    def start(self) -> None:
        """Inicia o navegador Playwright."""
        logger.info("Iniciando navegador...")
        self.playwright = sync_playwright().start()

        # Diretório temporário para downloads
        self.download_dir = Path(f"/tmp/maistop_downloads_{datetime.now():%Y%m%d_%H%M%S}")
        self.download_dir.mkdir(parents=True, exist_ok=True)
        logger.info("Diretório de download temporário: %s", self.download_dir)

        self.browser = self.playwright.chromium.launch(
            headless=self.headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )

        self.context = self.browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            accept_downloads=True,
        )

        self.page = self.context.new_page()
        self.page.set_default_timeout(config.DEFAULT_TIMEOUT * 1000)
        self.page.set_default_navigation_timeout(config.NAVIGATION_TIMEOUT * 1000)
        logger.info("Navegador iniciado com sucesso.")

    def stop(self) -> None:
        """Encerra o navegador e limpa arquivos temporários."""
        logger.info("Encerrando navegador...")
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        # Limpa diretório temporário
        if self.download_dir and self.download_dir.exists():
            shutil.rmtree(self.download_dir, ignore_errors=True)
            logger.info("Diretório temporário removido.")
        logger.info("Navegador encerrado.")

    def __enter__(self) -> MaisTopAutomation:
        self.start()
        return self

    def __exit__(self, *_) -> None:
        self.stop()

    # -----------------------------------------------------------------------
    # Ações do site
    # -----------------------------------------------------------------------
    def acessar_site(self) -> None:
        """Acessa a página inicial do site."""
        logger.info("Acessando %s ...", config.BASE_URL)
        try:
            response = self.page.goto(config.BASE_URL, wait_until="networkidle")
            if response and response.status >= 400:
                logger.warning("Status HTTP: %s", response.status)
        except PlaywrightTimeout:
            logger.warning("Timeout ao carregar página inicial. Tentando continuar...")

    def realizar_login(self) -> bool:
        """
        Realiza login no site.

        IMPORTANTE: Os seletores CSS abaixo são placeholders.
        Você deve atualizá-los inspecionando o site real (F12 → Elements).
        """
        logger.info("Tentando realizar login como '%s'...", config.USERNAME)

        # --- AJUSTE ESTES SELETORES CONFORME O SITE REAL ---
        # Exemplos comuns de seletores para formulários de login:
        SELECTORES_LOGIN = {
            "campo_usuario": [
                'input[name="username"]',
                'input[name="email"]',
                'input[name="login"]',
                'input[type="email"]',
                'input[id="username"]',
                'input[id="email"]',
                'input[placeholder*="usuário" i]',
                'input[placeholder*="email" i]',
            ],
            "campo_senha": [
                'input[name="password"]',
                'input[type="password"]',
                'input[id="password"]',
                'input[placeholder*="senha" i]',
            ],
            "botao_login": [
                'button[type="submit"]',
                'input[type="submit"]',
                'button:has-text("Entrar")',
                'button:has-text("Login")',
                'a:has-text("Entrar")',
            ],
        }

        # Localiza e preenche usuário
        usuario_encontrado = False
        for seletor in SELECTORES_LOGIN["campo_usuario"]:
            try:
                self.page.fill(seletor, config.USERNAME, timeout=3_000)
                logger.info("Campo de usuário preenchido (%s)", seletor)
                usuario_encontrado = True
                break
            except PlaywrightTimeout:
                continue

        if not usuario_encontrado:
            logger.error(
                "Não foi possível localizar o campo de usuário. "
                "Atualize os seletores no script inspecionando o site."
            )
            return False

        # Localiza e preenche senha
        senha_encontrada = False
        for seletor in SELECTORES_LOGIN["campo_senha"]:
            try:
                self.page.fill(seletor, config.PASSWORD, timeout=3_000)
                logger.info("Campo de senha preenchido (%s)", seletor)
                senha_encontrada = True
                break
            except PlaywrightTimeout:
                continue

        if not senha_encontrada:
            logger.error(
                "Não foi possível localizar o campo de senha. "
                "Atualize os seletores no script inspecionando o site."
            )
            return False

        # Clica no botão de login
        botao_encontrado = False
        for seletor in SELECTORES_LOGIN["botao_login"]:
            try:
                self.page.click(seletor, timeout=3_000)
                logger.info("Botão de login clicado (%s)", seletor)
                botao_encontrado = True
                break
            except PlaywrightTimeout:
                continue

        if not botao_encontrado:
            logger.error(
                "Não foi possível localizar o botão de login. "
                "Atualize os seletores no script inspecionando o site."
            )
            return False

        # Aguarda navegação pós-login
        try:
            self.page.wait_for_load_state("networkidle", timeout=10_000)
            logger.info("Login realizado com sucesso!")
            return True
        except PlaywrightTimeout:
            logger.warning(
                "Página não atingiu estado 'networkidle' após login. "
                "Verificando URL atual..."
            )
            logger.info("URL atual: %s", self.page.url)
            return True  # Pode ter funcionado mesmo com timeout

    def navegar_para_relatorios(self) -> bool:
        """Navega até a seção de relatórios/downloads."""
        logger.info("Navegando para relatórios...")

        # --- AJUSTE ESTES SELETORES CONFORME O SITE REAL ---
        SELECTORES_RELATORIOS = {
            "menu_relatorios": [
                'a:has-text("Relatórios")',
                'a:has-text("Relatorios")',
                'a:has-text("relatórios")',
                'a:has-text("Downloads")',
                'a:has-text("Base")',
                'a:has-text("Dados")',
                'nav a[href*="relatorio" i]',
                'nav a[href*="report" i]',
            ],
        }

        for seletor in SELECTORES_RELATORIOS["menu_relatorios"]:
            try:
                self.page.click(seletor, timeout=5_000)
                logger.info("Menu de relatórios clicado (%s)", seletor)
                self.page.wait_for_load_state("networkidle", timeout=10_000)
                logger.info("URL atual: %s", self.page.url)
                return True
            except PlaywrightTimeout:
                continue

        # Se não encontrou pelo seletor, tenta navegar direto pela URL
        try:
            self.page.goto(config.RELATORIOS_URL, wait_until="networkidle")
            logger.info("Navegação direta para %s", config.RELATORIOS_URL)
            return True
        except PlaywrightTimeout:
            logger.warning("Timeout na navegação direta para relatórios.")
            return False

    def _download_arquivo(self, nome_base: str, seletores_possiveis: list[str]) -> Optional[Path]:
        """
        Tenta fazer download de um arquivo clicando em um dos seletores fornecidos.

        Args:
            nome_base: Nome descritivo do arquivo (ex: "cadastro", "acessos")
            seletores_possiveis: Lista de seletores CSS para tentar clicar

        Returns:
            Path do arquivo baixado, ou None se falhou
        """
        logger.info("Buscando download de '%s'...", nome_base)

        for seletor in seletores_possiveis:
            try:
                # Configura handler de download
                download_path: Optional[Path] = None

                with self.page.expect_download(timeout=config.DOWNLOAD_TIMEOUT * 1000) as download_info:
                    self.page.click(seletor, timeout=5_000)
                    logger.info("Clique em '%s' realizado (%s)", nome_base, seletor)

                download: Download = download_info.value
                caminho = self.download_dir / download.suggested_filename
                download.save_as(caminho)
                logger.info(
                    "Download de '%s' concluído: %s (%.2f KB)",
                    nome_base,
                    caminho.name,
                    caminho.stat().st_size / 1024,
                )
                return caminho

            except PlaywrightTimeout:
                logger.debug("Seletor '%s' não funcionou para '%s'", seletor, nome_base)
                continue
            except Exception as exc:
                logger.warning("Erro ao tentar download com '%s': %s", seletor, exc)
                continue

        logger.error(
            "Não foi possível fazer download de '%s'. "
            "Verifique os seletores no script.",
            nome_base,
        )
        return None

    def baixar_cadastro(self) -> Optional[Path]:
        """Faz download da base de cadastro."""
        # --- AJUSTE ESTES SELETORES CONFORME O SITE REAL ---
        seletores = [
            'a:has-text("Cadastro")',
            'a:has-text("cadastro")',
            'button:has-text("Cadastro")',
            'a[href*="cadastro"]',
            'a[download*="cadastro"]',
            'button:has-text("Exportar Cadastro")',
            'a:has-text("Exportar Cadastro")',
        ]
        return self._download_arquivo("cadastro", seletores)

    def baixar_acessos(self) -> Optional[Path]:
        """Faz download da base de acessos."""
        # --- AJUSTE ESTES SELETORES CONFORME O SITE REAL ---
        seletores = [
            'a:has-text("Acessos")',
            'a:has-text("acessos")',
            'button:has-text("Acessos")',
            'a[href*="acesso"]',
            'a[download*="acesso"]',
            'button:has-text("Exportar Acessos")',
            'a:has-text("Exportar Acessos")',
        ]
        return self._download_arquivo("acessos", seletores)

    # -----------------------------------------------------------------------
    # Organização de arquivos
    # -----------------------------------------------------------------------
    def _mover_para_destino(self, arquivo: Path, nome_base: str) -> Path:
        """Move o arquivo baixado para a pasta de destino organizada por data."""
        hoje = datetime.now()
        pasta_data = config.BASES_DIR / f"{hoje:%d%m%y}"
        pasta_data.mkdir(parents=True, exist_ok=True)

        # Nome do arquivo: nome_base_YYYY-MM-DDTHH_MM_SS.csv
        extensao = arquivo.suffix or ".csv"
        nome_destino = f"{nome_base}_{hoje:%Y-%m-%dT%H_%M_%S}{extensao}"
        destino = pasta_data / nome_destino

        shutil.copy2(arquivo, destino)
        logger.info("Arquivo salvo em: %s", destino)
        return destino

    # -----------------------------------------------------------------------
    # Execução principal
    # -----------------------------------------------------------------------
    def executar(self) -> dict[str, Optional[Path]]:
        """
        Executa o fluxo completo de automação.

        Returns:
            Dict com os caminhos dos arquivos baixados.
            Valor None indica falha no download daquela base.
        """
        resultados: dict[str, Optional[Path]] = {
            "cadastro": None,
            "acessos": None,
        }

        # Valida configurações
        erros = config.validar_config()
        if erros:
            for erro in erros:
                logger.error(erro)
            return resultados

        # 1. Acessa site
        self.acessar_site()

        # 2. Login
        if not self.realizar_login():
            logger.error("Falha no login. Encerrando.")
            return resultados

        # 3. Navega para relatórios
        if not self.navegar_para_relatorios():
            logger.warning(
                "Não conseguiu navegar para relatórios. "
                "Tentando continuar na página atual..."
            )

        # 4. Downloads
        arquivo_cadastro = self.baixar_cadastro()
        if arquivo_cadastro:
            resultados["cadastro"] = self._mover_para_destino(arquivo_cadastro, "cadastro")

        # Pequena pausa entre downloads
        time.sleep(1)

        arquivo_acessos = self.baixar_acessos()
        if arquivo_acessos:
            resultados["acessos"] = self._mover_para_destino(arquivo_acessos, "acessos")

        return resultados


def salvar_screenshot(page: Page, nome: str, pasta: Path) -> Path:
    """Salva um screenshot da página atual (útil para debug)."""
    pasta.mkdir(parents=True, exist_ok=True)
    caminho = pasta / f"{nome}_{datetime.now():%Y%m%d_%H%M%S}.png"
    page.screenshot(path=caminho, full_page=True)
    logger.info("Screenshot salvo: %s", caminho)
    return caminho


def main() -> int:
    """Entry point do script."""
    parser = argparse.ArgumentParser(
        description="Download automatizado de bases do Programa Mais Top",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Executa com navegador visível (modo debug)",
    )
    parser.add_argument(
        "--screenshot",
        action="store_true",
        help="Salva screenshots de cada etapa (modo debug)",
    )
    args = parser.parse_args()

    logger.info("=" * 60)
    logger.info("Programa Mais Top - Automação de Download")
    logger.info("Data: %s", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    logger.info("=" * 60)

    try:
        with MaisTopAutomation(headless=not args.visible) as bot:
            # Salva screenshot inicial se em modo debug
            if args.screenshot:
                pasta_screenshots = config.PROJECT_ROOT / "automation" / "screenshots"
                salvar_screenshot(bot.page, "01_inicial", pasta_screenshots)

            resultados = bot.executar()

            # Salva screenshot final se em modo debug
            if args.screenshot:
                pasta_screenshots = config.PROJECT_ROOT / "automation" / "screenshots"
                salvar_screenshot(bot.page, "99_final", pasta_screenshots)

        # Resumo
        logger.info("=" * 60)
        logger.info("RESUMO DO DOWNLOAD")
        logger.info("=" * 60)
        sucesso = 0
        for nome, caminho in resultados.items():
            if caminho and caminho.exists():
                tamanho = caminho.stat().st_size / 1024
                logger.info("✅ %s: %s (%.2f KB)", nome, caminho, tamanho)
                sucesso += 1
            else:
                logger.error("❌ %s: FALHA no download", nome)

        logger.info("=" * 60)
        logger.info("Concluído: %d/2 downloads com sucesso", sucesso)
        logger.info("=" * 60)

        return 0 if sucesso == 2 else 1

    except KeyboardInterrupt:
        logger.info("Operação cancelada pelo usuário.")
        return 130
    except Exception as exc:
        logger.exception("Erro inesperado: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
