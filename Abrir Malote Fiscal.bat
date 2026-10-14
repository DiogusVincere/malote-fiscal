@echo off
rem Abre o Malote Fiscal com o Outlook desta máquina. Precisa do Python 3.12+ com as bibliotecas do requirements.txt.
cd /d "%~dp0"
start "" pythonw malote.pyw %*
