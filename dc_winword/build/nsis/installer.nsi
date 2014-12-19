!define PRODUCT_ICON "glossyorb.ico"
!define PRODUCT_NAME "DocCleaner for Winword"
!define ARCH_TAG ""
!define PY_MAJOR_VERSION "3.4"
!define PRODUCT_VERSION "0.2"
!define PY_QUALIFIER "3.4-32"
!define PY_VERSION "3.4.1"
!define INSTALLER_NAME "DocCleaner_for_Winword_0.2.exe"
!define BITNESS "32"

; Definitions will be added above
 
SetCompressor lzma

RequestExecutionLevel admin ;

; Modern UI installer stuff 
!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"

; UI pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_LANGUAGE "English"
; MUI end ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "${INSTALLER_NAME}"
InstallDir "$PROGRAMFILES${BITNESS}\${PRODUCT_NAME}"
ShowInstDetails show

Section -SETTINGS
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
SectionEnd

Section "Python ${PY_VERSION}" sec_py
  File "python-${PY_VERSION}${ARCH_TAG}.msi"
  DetailPrint "Installing Python ${PY_MAJOR_VERSION}, ${BITNESS} bit"
  ExecWait 'msiexec /i "$INSTDIR\python-${PY_VERSION}${ARCH_TAG}.msi" \
            /qb ALLUSERS=1 TARGETDIR="$COMMONFILES${BITNESS}\Python\${PY_MAJOR_VERSION}"'
  Delete $INSTDIR\python-${PY_VERSION}${ARCH_TAG}.msi
SectionEnd

SectionGroup "Python modules" sec_pymod
	Section "LXML" sec_lxml
	SetOutPath $INSTDIR\Prerequisites
	File ".\Prerequisites\lxml-3.4.0.win32-py3.4.exe"
		ExecWait "$INSTDIR\Prerequisites\lxml-3.4.0.win32-py3.4.exe"
	SectionEnd

	Section "PyWin32" sec_pywin
	File ".\Prerequisites\pywin32-219.win32-py3.4.exe"
	ExecWait "$INSTDIR\Prerequisites\pywin32-219.win32-py3.4.exe"
	SectionEnd

	Section "SimpleJSON" sec_simplejson
	File ".\Prerequisites\simplejson-3.6.5.win32-py3.4.exe"
	ExecWait "$INSTDIR\Prerequisites\simplejson-3.6.5.win32-py3.4.exe"
	SectionEnd
SectionGroupEnd

;PYLAUNCHER_INSTALL
;------------------

Section "!${PRODUCT_NAME}" sec_app
  SectionIn RO
  SetShellVarContext all
  File ${PRODUCT_ICON}
  SetOutPath "$INSTDIR\pkgs"
  File /r "pkgs\*.*"
  SetOutPath "$INSTDIR"
  ;INSTALL_FILES
  SetOutPath "$INSTDIR"
  File "DocCleaner_for_Winword.launch.pyw"
  File "glossyorb.ico"
  File "winword_addin.json"
  File "winword_addin.xml"
  File "wordaddin.py"
  SetOutPath "$INSTDIR"
  ;INSTALL_DIRECTORIES
  SetOutPath "$INSTDIR"
  ;INSTALL_SHORTCUTS
  SetOutPath "%HOMEDRIVE%\%HOMEPATH%"
  CreateShortCut "$SMPROGRAMS\DocCleaner for Winword.lnk" "pyw" '"$INSTDIR\DocCleaner_for_Winword.launch.pyw"' \
      "$INSTDIR\glossyorb.ico"
  SetOutPath "$INSTDIR"
  ; Byte-compile Python files.
  DetailPrint "Byte-compiling Python modules..."
  nsExec::ExecToLog 'py -${PY_QUALIFIER} -m compileall -q "$INSTDIR\pkgs"'
  WriteUninstaller $INSTDIR\uninstall.exe
  ; Add ourselves to Add/remove programs
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "DisplayName" "${PRODUCT_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "DisplayIcon" "$INSTDIR\${PRODUCT_ICON}"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "NoRepair" 1
  
  ;Launch the pyw script, for registering the MS Office COM addin
  Call LaunchScript
SectionEnd

Section "Uninstall"
  SetShellVarContext all
  Delete $INSTDIR\uninstall.exe
  Delete "$INSTDIR\${PRODUCT_ICON}"
  RMDir /r "$INSTDIR\pkgs"
  RMDir /r "$INSTDIR\Prerequisites"
  ;UNINSTALL_FILES
  Delete "$INSTDIR\DocCleaner_for_Winword.launch.pyw"
  Delete "$INSTDIR\glossyorb.ico"
  Delete "$INSTDIR\winword_addin.json"
  Delete "$INSTDIR\winword_addin.xml"
  Delete "$INSTDIR\wordaddin.py"
  ;UNINSTALL_DIRECTORIES
  ;UNINSTALL_SHORTCUTS
  Delete "$SMPROGRAMS\DocCleaner for Winword.lnk"
  RMDir $INSTDIR
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
SectionEnd

; Functions

Function .onMouseOverSection
    ; Find which section the mouse is over, and set the corresponding description.
    FindWindow $R0 "#32770" "" $HWNDPARENT
    GetDlgItem $R0 $R0 1043 ; description item (must be added to the UI)

    StrCmp $0 ${sec_py} 0 +2
      SendMessage $R0 ${WM_SETTEXT} 0 "STR:The Python interpreter. \
            This is required for ${PRODUCT_NAME} to run."
			
			
	StrCmp $0 ${sec_pymod} 0 +2
		SendMessage $R0 ${WM_SETTEXT} 0 "STR:Python modules. \
				They are required for ${PRODUCT_NAME} to run."
    ;
    ;PYLAUNCHER_HELP
    ;------------------

    StrCmp $0 ${sec_app} "" +2
      SendMessage $R0 ${WM_SETTEXT} 0 "STR:${PRODUCT_NAME}"
FunctionEnd

Function LaunchScript  
  ;Launch the script as user, to pop up UAC prompt under Windows Vista (won't work otherwise). Need to install the ShellExecAsUser NSIS plugin: http://nsis.sourceforge.net/ShellExecAsUser_plug-in
  ShellExecAsUser::ShellExecAsUser "" "$INSTDIR\DocCleaner_for_Winword.launch.pyw"  
FunctionEnd
