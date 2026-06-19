#!/bin/bash
# -*- coding: utf-8 -*-
# Configura o rclone para upload automático no Google Drive.
# Executar uma única vez: bash setup_rclone.sh

set -e

echo "==================================="
echo " Configuração do Google Drive"
echo " Relatório Semanal +TOP"
echo "==================================="
echo ""

# Verifica/instala o rclone
if ! command -v rclone &> /dev/null; then
    echo "rclone não encontrado. Instalando..."
    curl https://rclone.org/install.sh | sudo bash
else
    echo "rclone já está instalado."
fi

echo ""
echo "Vamos configurar o acesso ao Google Drive."
echo "Siga as instruções abaixo:"
echo ""
echo "1. Quando aparecer 'n/s/q>', digite 'n' (novo remote)"
echo "2. Name: gdrive"
echo "3. Storage: 18 (Google Drive)"
echo "4. Client Id: pressione Enter (deixe em branco)"
echo "5. Client Secret: pressione Enter (deixe em branco)"
echo "6. Scope: 1 (Full access)"
echo "7. Root Folder ID: pressione Enter"
echo "8. Service Account File: pressione Enter"
echo "9. Edit advanced config? n"
echo "10. Use auto config? y  (vai abrir o navegador para login)"
echo ""
read -p "Pressione Enter para começar a configuração..."

rclone config

echo ""
echo "Criando pasta 'relatorio_semanal' no Drive..."
rclone mkdir gdrive:relatorio_semanal || true

echo ""
echo "Obtendo link da pasta..."
FOLDER_ID=$(rclone lsf gdrive: --format pi 2>/dev/null | grep "relatorio_semanal" | awk '{print $1}' || true)

if [ -z "$FOLDER_ID" ]; then
    echo "Não foi possível obter o ID da pasta automaticamente."
    echo "No Google Drive, clique com o botão direito na pasta 'relatorio_semanal'"
    echo "e copie o link. Cole esse link no config_email.json no campo 'link_drive'."
else
    LINK="https://drive.google.com/drive/folders/$FOLDER_ID"
    echo ""
    echo "==================================="
    echo " Link da pasta no Drive:"
    echo " $LINK"
    echo "==================================="
    echo ""
    echo "Atualize o arquivo config_email.json com:"
    echo "  \"link_drive\": \"$LINK\""
fi

echo ""
echo "Pronto! A partir de agora, o relatório semanal fará upload automático dos arquivos."
echo "Para testar manualmente: python3 upload_drive.py"
