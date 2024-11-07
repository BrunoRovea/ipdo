@echo off
REM Ativar o ambiente virtual
call .\.venv\Scripts\activate.bat

REM Executar o script Python
python ipdo.py

REM Desativar o ambiente virtual ao finalizar o script
deactivate
