@ECHO OFF

REM Temporarily redefining environment variables Pythonpath and Pythonhome, in case there is a previous installation of Python on the machine
SET Pythonpath=
SET Pythonhome=

REM Defining paths to miniconda and to the app
SET "MinicondaPath=%LOCALAPPDATA%\DocCleaner\Miniconda3\"
SET "InstallDir=%LOCALAPPDATA%\DocCleaner\"

REM Launching the script with Miniconda3's pythonw.exe
START "" "%MinicondaPath%\pythonw.exe" "%InstallDir%\wordaddin.py"
