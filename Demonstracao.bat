@echo off
rem Abre o Malote Fiscal em modo demonstração: Outlook simulado e clientes fictícios, sem precisar de Outlook.
cd /d "%~dp0"
start "" pythonw malote.pyw --demo
