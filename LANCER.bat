@echo off
cd /d "%~dp0"

:: Ajouter les emplacements possibles de uv au PATH
:: (le PATH permanent mis a jour par l'installateur n'est pas toujours visible
::  dans un nouveau terminal — on le force ici pour etre sur)
set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%LOCALAPPDATA%\Programs\uv;%PATH%"

uv run python app.py

if %errorlevel% neq 0 (
    echo.
    echo -------------------------------------------------------
    echo  ERREUR : L'application n'a pas pu demarrer.
    echo.
    echo  Solutions :
    echo    1. Avez-vous execute INSTALLER.bat une premiere fois ?
    echo    2. Fermez cette fenetre et relancez INSTALLER.bat.
    echo -------------------------------------------------------
    echo.
    pause
)