@ECHO OFF
REM Defining paths to miniconda and to the app
SET "MinicondaPath=%LOCALAPPDATA%\DocCleaner\Miniconda3\"
SET "InstallDir=%LOCALAPPDATA%\DocCleaner\"
SET CurrentDir=%~dp0

REM Temporarily redefining environment variables Pythonpath and Pythonhome, in case there is a previous installation of Python on the machine
SET Pythonpath=
SET Pythonhome=

REM Lauching the install of miniconda
START /WAIT "" "%CurrentDir%\Miniconda3-latest-Windows-x86.exe" /S /D=%MinicondaPath%

REM Installing pip in Miniconda
START /WAIT "" "%MinicondaPath%\Scripts\conda.exe" install pip -y

REM installing dependencies LXML and pywin32
"%MinicondaPath%\Scripts\pip.exe" install "%CurrentDir%\Prerequisites\lxml-3.4.1-cp34-none-win32.whl" 
"%MinicondaPath%\Scripts\pip.exe" install "%CurrentDir%\Prerequisites\pywin32-219-cp34-none-win32.whl" 

REM If the user is admin, we need to launch the post install script for pywin32, otherwise it won't work
IF NOT %ERRORLEVEL%==0 GOTO skipPostInstall
"%MinicondaPath%\Python.exe" "%MinicondaPath%\Scripts\pywin32_postinstall.py" -install
:skipPostInstall

REM Installating dependencies simplejson and Pillow
"%MinicondaPath%\Scripts\pip.exe" install "%CurrentDir%\Prerequisites\simplejson-3.6.5-cp34-none-win32.whl" 
"%MinicondaPath%\Scripts\pip.exe" install "%CurrentDir%\Prerequisites\Pillow-2.7.0-cp34-none-win32.whl"

