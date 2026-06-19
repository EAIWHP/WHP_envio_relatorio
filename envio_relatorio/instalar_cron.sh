#!/usr/bin/env bash
# Instala o agendamento semanal do relatório +TOP no cron do usuário
# Uso: ./instalar_cron.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_PATH="${SCRIPT_DIR}/.venv/bin/python3"

# Se o ambiente virtual não existir, tenta python3 do sistema
if [ ! -f "${PYTHON_PATH}" ]; then
    PYTHON_PATH="python3"
fi

CRON_LINE="0 11 * * 1 cd ${SCRIPT_DIR} && ${PYTHON_PATH} ${SCRIPT_DIR}/gerar_relatorio_semanal.py >> ${SCRIPT_DIR}/logs/cron.log 2>&1"

echo "=============================================="
echo "Instalador de cron - Relatório Semanal +TOP"
echo "=============================================="
echo ""
echo "Será adicionada a seguinte linha no crontab:"
echo ""
echo "${CRON_LINE}"
echo ""
echo "Isso executa o relatório toda SEGUNDA-FEIRA às 11:00."
echo "Para mudar o dia/horário, edite o crontab com: crontab -e"
echo ""
read -p "Deseja continuar? (s/N): " confirm

if [[ "${confirm,,}" != "s" ]]; then
    echo "Instalação cancelada."
    exit 0
fi

# Remove linha antiga se existir
(crontab -l 2>/dev/null | grep -v "gerar_relatorio_semanal.py") | crontab -

# Adiciona nova linha
(crontab -l 2>/dev/null; echo "${CRON_LINE}") | crontab -

echo ""
echo "Cron instalado com sucesso!"
echo "Para visualizar: crontab -l"
echo "Para testar manualmente: cd ${SCRIPT_DIR} && ${PYTHON_PATH} gerar_relatorio_semanal.py --teste"
