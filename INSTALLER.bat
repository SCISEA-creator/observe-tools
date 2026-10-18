@echo off
cd /d "%~dp0"
echo.
echo =====================================================
echo    Observe-Tools - Installation
echo    IRD OB7 - SCISEA
echo =====================================================
echo.
echo Ce script installe Python et toutes les librairies
echo necessaires. Une connexion internet est requise.
echo (Une seule fois - en mer, utiliser LANCER.bat)
echo.

:: Etape 1 : Verifier si uv est deja installe
where uv > nul 2>&1
if %errorlevel% == 0 (
    echo [OK] uv est deja installe.
    goto installer_dependances
)

echo [1/3] Telechargement de uv (gestionnaire Python leger, ~12 Mo)...
echo.
powershell -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
if %errorlevel% neq 0 (
    echo.
    echo ERREUR : Le telechargement a echoue.
    echo Verifiez votre connexion internet et reessayez.
    pause
    exit /b 1
)

:: Rafraichir le PATH pour trouver uv dans ce meme script
:: Couvre les deux emplacements possibles selon la version de uv
set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%LOCALAPPDATA%\Programs\uv;%PATH%"

:: Verifier que uv est maintenant accessible
where uv > nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo uv a ete installe mais n'est pas encore dans le PATH.
    echo Fermez cette fenetre, rouvrez-en une nouvelle,
    echo et relancez INSTALLER.bat.
    pause
    exit /b 1
)

:installer_dependances
echo.
echo [2/3] Verification de Python 3.11 (telechargement si absent)...
echo.
uv python install 3.11
if %errorlevel% neq 0 (
    echo.
    echo ERREUR : Impossible d'installer Python 3.11.
    echo Verifiez votre connexion internet et reessayez.
    pause
    exit /b 1
)

echo.
echo [3/3] Installation des librairies Python (pandas, matplotlib, etc.)...
echo.
uv sync
if %errorlevel% neq 0 (
    echo.
    echo ERREUR lors de l'installation des librairies.
    echo Verifiez votre connexion internet et reessayez.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo    Installation terminee avec succes !
echo.
echo    Vous pouvez maintenant utiliser LANCER.bat
echo    pour demarrer l'application (meme hors ligne).
echo =====================================================
echo.
pause