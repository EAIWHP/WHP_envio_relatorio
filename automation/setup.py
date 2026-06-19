#!/usr/bin/env python3
"""
Script de setup/instalacao das dependencias do automacao.

Uso:
    python setup.py
"""

import subprocess
import sys
from pathlib import Path


def instalar_dependencias() -> None:
    """Instala as dependencias do projeto."""
    req_file = Path(__file__).resolve().parent / "requirements.txt"

    if not req_file.exists():
        print(f"Arquivo nao encontrado: {req_file}")
        sys.exit(1)

    print("Instalando dependencias...")
    subprocess.check_call([
        sys.executable, "-m", "pip", "install", "-r", str(req_file)
    ])

    print("Instalando navegadores do Playwright...")
    subprocess.check_call([
        sys.executable, "-m", "playwright", "install", "chromium"
    ])

    print("\n✅ Instalacao concluida!")
    print("\nProximos passos:")
    print("  1. Copie .env.example para .env")
    print("  2. Preencha suas credenciais no arquivo .env")
    print("  3. Execute: python download_bases.py --visible")


if __name__ == "__main__":
    instalar_dependencias()
