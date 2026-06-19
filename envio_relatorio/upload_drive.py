#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Upload automático dos relatórios gerados para Google Drive ou SharePoint via rclone.

Requisitos:
- Para Google Drive: rclone configurado com remote 'gdrive'
- Para SharePoint: rclone configurado com remote 'sharepoint'

Uso:
    python3 upload_drive.py
"""

import json
import subprocess
import sys
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "relatorios_gerados"
CONFIG_FILE = BASE_DIR / "config_email.json"


def is_sharepoint_link(link):
    return bool(link and "sharepoint.com" in link.lower())


def extrair_folder_id_drive(link):
    """Extrai o ID da pasta do Google Drive a partir do link compartilhado."""
    if not link or "folders/" not in link:
        return None
    partes = link.split("folders/")
    if len(partes) < 2:
        return None
    folder_id = partes[1].split("?")[0].split("/")[0]
    return folder_id if folder_id else None


def main():
    if not CONFIG_FILE.exists():
        print("ERRO: config_email.json não encontrado.")
        sys.exit(1)

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)

    link = config.get("link_drive", "")

    if not link or link == "SEU_LINK_DO_DRIVE_AQUI":
        print("ERRO: Nenhum link válido configurado em config_email.json.")
        sys.exit(1)

    if not OUTPUT_DIR.exists():
        print(f"ERRO: Pasta de relatórios não encontrada: {OUTPUT_DIR}")
        sys.exit(1)

    data_str = date.today().strftime("%Y%m%d")
    arquivos = sorted(OUTPUT_DIR.glob(f"relatorio_top_{data_str}*"))

    if not arquivos:
        print(f"Nenhum arquivo encontrado para upload em {OUTPUT_DIR} com data {data_str}.")
        sys.exit(0)

    if is_sharepoint_link(link):
        remote = "sharepoint"
        destino = f"{remote}:relatorio_semanal"
        print(f"Enviando {len(arquivos)} arquivo(s) para o SharePoint (pasta relatorio_semanal)...")
    else:
        folder_id = extrair_folder_id_drive(link)
        if not folder_id:
            print("ERRO: Não foi possível extrair o ID da pasta do Google Drive.")
            sys.exit(1)
        remote = "gdrive"
        destino = f"{remote}:{folder_id}"
        print(f"Enviando {len(arquivos)} arquivo(s) para o Drive (pasta {folder_id})...")

    for arquivo in arquivos:
        resultado = subprocess.run(
            ["rclone", "copy", str(arquivo), destino],
            capture_output=True, text=True
        )
        if resultado.returncode == 0:
            print(f"  ✓ {arquivo.name}")
        else:
            print(f"  ✗ Falha ao enviar {arquivo.name}:")
            print(resultado.stderr)
            sys.exit(1)

    print("Upload concluído com sucesso.")


if __name__ == "__main__":
    main()
