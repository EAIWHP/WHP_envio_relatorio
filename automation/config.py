"""
Configuracoes do script de automacao do Programa Mais Top.

Credenciais podem ser definidas via variaveis de ambiente (recomendado)
ou diretamente neste arquivo (nao recomendado para producao).

Variaveis de ambiente:
    MAISTOP_USER     - Usuario do site programamaistop.com.br
    MAISTOP_PASS     - Senha do site
    MAISTOP_BASE_DIR - Diretorio base do projeto (opcional)
"""

import os
from pathlib import Path

# ---------------------------------------------------------------------------
# Caminhos
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Diretorio onde as bases recorrentes serao salvas
BASES_DIR = Path(os.getenv("MAISTOP_BASE_DIR", PROJECT_ROOT / "bases_recorrentes"))

# ---------------------------------------------------------------------------
# Credenciais (prioridade: variaveis de ambiente > arquivo)
# ---------------------------------------------------------------------------
USERNAME = os.getenv("MAISTOP_USER", "")
PASSWORD = os.getenv("MAISTOP_PASS", "")

# ---------------------------------------------------------------------------
# URL do site
# ---------------------------------------------------------------------------
BASE_URL = "https://programamaistop.com.br"
LOGIN_URL = f"{BASE_URL}/login"
# Ajuste o path de relatorios conforme a navegacao real do site
RELATORIOS_URL = f"{BASE_URL}/relatorios"

# ---------------------------------------------------------------------------
# Timeouts (segundos)
# ---------------------------------------------------------------------------
DEFAULT_TIMEOUT = 30          # Timeout padrao para operacoes
NAVIGATION_TIMEOUT = 60       # Timeout para navegacao de pagina
DOWNLOAD_TIMEOUT = 120        # Timeout para downloads

# ---------------------------------------------------------------------------
# Nome dos arquivos esperados no download
# ---------------------------------------------------------------------------
ARQUIVOS_ESPERADOS = {
    "cadastro": "cadastro",
    "acessos": "acessos",
}

# ---------------------------------------------------------------------------
# Validacao
# ---------------------------------------------------------------------------
def validar_config() -> list[str]:
    """Retorna lista de erros de configuracao. Lista vazia = OK."""
    erros = []
    if not USERNAME:
        erros.append("USERNAME nao configurado. Defina a variavel MAISTOP_USER.")
    if not PASSWORD:
        erros.append("PASSWORD nao configurado. Defina a variavel MAISTOP_PASS.")
    if not BASES_DIR.exists():
        erros.append(f"Diretorio de bases nao encontrado: {BASES_DIR}")
    return erros
