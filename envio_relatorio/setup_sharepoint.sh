#!/bin/bash
# -*- coding: utf-8 -*-
# Configura o rclone para upload automático no SharePoint via backend OneDrive.
# Executar uma única vez: bash setup_sharepoint.sh

set -e

echo "==================================="
echo " Configuração do SharePoint"
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
echo "Vamos configurar o acesso ao SharePoint/OneDrive for Business."
echo "Siga as instruções abaixo:"
echo ""
echo "1. Quando aparecer 'n/s/q>', digite 'n' (novo remote)"
echo "2. Name: sharepoint"
echo "3. Storage: 31 (OneDrive) - o número pode variar, procure por 'OneDrive'"
echo "4. Region: 0 (Microsoft Cloud Global)"
echo "5. tenant: pressione Enter (deixe em branco)"
echo "6. Edit advanced config? n"
echo "7. Use auto config? y (vai abrir o navegador para login)"
echo ""
echo "Após o login, o rclone vai listar os sites/drives disponíveis."
echo "Escolha o drive correspondente ao site:"
echo "  https://eaimkt1.sharepoint.com/sites/DadosPessoais-Sensveis"
echo ""
read -p "Pressione Enter para começar a configuração..."

rclone config

echo ""
echo "Criando pasta 'relatorio_semanal' no SharePoint..."
rclone mkdir sharepoint:relatorio_semanal || true

echo ""
echo "Pronto! A partir de agora, o relatório semanal fará upload automático para o SharePoint."
echo "Para testar manualmente: python3 upload_drive.py"
echo ""
echo "IMPORTANTE: Atualize o config_email.json com o link do SharePoint que você deseja mostrar no e-mail."
