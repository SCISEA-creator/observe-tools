@echo off
cd /d "%~dp0"
echo.
echo =====================================================
echo    Observe-Tools - Mise a jour
echo    IRD OB7 - SCISEA
echo =====================================================
echo.
echo Une connexion internet est requise.
echo.

:: Ajouter les emplacements possibles de uv au PATH
set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%LOCALAPPDATA%\Programs\uv;%PATH%"

echo [1/2] Recuperation des derniers correctifs (git pull)...
git pull
if %errorlevel% neq 0 (
    echo.
    echo AVERTISSEMENT : Impossible de recuperer les mises a jour.
    echo Verifiez votre connexion internet.
    echo L'application reste utilisable avec la version actuelle.
    echo.
    pause
    exit /b 1
)

echo.
echo [2/2] Mise a jour des librairies Python...
uv sync
if %errorlevel% neq 0 (
    echo.
    echo ERREUR lors de la mise a jour des librairies.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo    Mise a jour terminee !
echo    Lancez LANCER.bat pour demarrer l'application.
echo =====================================================
echo.
pause