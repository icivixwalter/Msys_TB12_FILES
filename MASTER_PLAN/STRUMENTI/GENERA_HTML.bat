@echo off
setlocal
title GENERA HTML - documenti del Master Plan
REM ==================================================================================================
REM  FILE   : GENERA_HTML.bat
REM  SCOPO  : rigenera i file .html della cartella superiore partendo dai file .md con lo stesso nome
REM  USO    : modifica un .md, poi doppio click su questo file. Serve Python 3 (nessuna libreria esterna).
REM  INDICE : #01# controllo Python   #02# conversione   #03# esito
REM ==================================================================================================

REM //#01# controllo Python
where python >nul 2>&1
if errorlevel 1 (
    echo ERRORE #01#: Python non trovato. Installalo da https://www.python.org oppure dal Microsoft Store.
    pause
    exit /b 1
)

REM //#02# conversione di tutti i .md della cartella superiore
python "%~dp0md2html.py" "%~dp0.."
set "RC=%ERRORLEVEL%"

REM //#03# esito
echo.
if "%RC%"=="0" (echo HTML aggiornati.) else (echo ERRORE #02#: conversione non riuscita, codice %RC%.)
pause
exit /b %RC%
