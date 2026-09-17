@echo off
REM ---------------------------------------------------------------------------
REM  Compilation d'INJXL en executable autonome (INJXL.exe)
REM  Prerequis : Python 3.9+ installe et accessible dans le PATH
REM ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

echo.
echo === Installation des dependances ===
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller
REM  xlrd n'est utile que pour LIRE d'anciens fichiers .xls
python -m pip install xlrd

echo.
echo === Compilation (interface graphique, sans console) ===
python -m PyInstaller --onefile --windowed --clean --noconfirm ^
    --name INJXL ^
    --hidden-import openpyxl.cell._writer ^
    --hidden-import injxl.gui ^
    --hidden-import injxl.engine ^
    --hidden-import injxl.tables ^
    --hidden-import injxl.model ^
    INJXL.py
if errorlevel 1 goto :echec

echo.
echo === Signature de l'executable ===
REM  PyInstaller reecrit l'exe a chaque compilation : la signature precedente
REM  disparait, il faut donc resigner ici et pas ailleurs.
powershell -ExecutionPolicy Bypass -File "%~dp0signer.ps1"
if errorlevel 1 (
    echo.
    echo *** L'exe est compile mais NON SIGNE. Ne le publiez pas en l'etat. ***
    goto :fin
)

echo.
echo === Copie des profils d'exemple a cote de l'executable ===
if not exist "dist\profils" mkdir "dist\profils"
copy /Y "profils\*.json" "dist\profils\" >nul

echo.
echo ===========================================================
echo  Termine : dist\INJXL.exe
echo.
echo  Double-clic          -^> interface graphique
echo  INJXL.exe mon.json   -^> interface prechargee avec un profil
echo  INJXL.exe --profil mon.json --oui  -^> execution silencieuse
echo     (l'exe est compile sans console : le resultat et le
echo      rapport .txt sont ecrits sur le disque)
echo ===========================================================
goto :fin

:echec
echo.
echo *** La compilation a echoue. Lisez les messages ci-dessus. ***

:fin
pause
