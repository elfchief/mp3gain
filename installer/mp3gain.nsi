;do a /DFULL on the command line for FULL installation
;/DDEUTSCH will make a German installer
;/DSPANISH will make a Spanish installer
;/DFRENCH will make a French installer

!define VER_MAJOR 1
!define VER_MINOR .2
!define VER_REL .3

Name "MP3Gain"
Caption "MP3Gain ${VER_MAJOR}${VER_MINOR}${VER_REL}"
Icon misc\mp3gain16.ico

UninstallText "This will uninstall MP3Gain ${VER_MAJOR}${VER_MINOR}${VER_REL} from your system. Would you like to go ahead?" "Uninstalling: "
UninstallCaption "Uninstalling"
ShowUninstDetails hide
;UninstallButtonText Uninstall
UninstallIcon misc/uninstall.ico

!ifdef FULL
  !ifdef DEUTSCH
    OutFile MP3GainFULLDeutschInstall_${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !else ifdef SPANISH
    OutFile MP3GainFULLEspañolInstall_${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !else ifdef FRENCH
    OutFile MP3GainFULLFrenchInstall_${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !else
    OutFile mp3gain-win-full-${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !endif
!else
  !ifdef DEUTSCH
    OutFile MP3GainNORMALDeutschInstall_${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !else ifdef SPANISH
    OutFile MP3GainNORMALEspañolInstall_${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !else ifdef FRENCH
    OutFile MP3GainNORMALFrenchInstall_${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !else
    OutFile mp3gain-win-${VER_MAJOR}${VER_MINOR}${VER_REL}.exe
  !endif
!endif

WindowIcon off

CRCCheck on
;LogSet off
SetCompress auto
SetDatablockOptimize on
SetDateSave on
InstProgressFlags smooth

FileErrorText `Can't overwrite $0. Please close any programs which may be using that particular file. Or just click "Ignore" and see whether MP3Gain will still run.`

DirText "Installation Directory"

!ifdef NSIS_CONFIG_COMPONENTPAGE
ComponentText "This will install MP3Gain ${VER_MAJOR}${VER_MINOR}${VER_REL} on your computer:"
!ifdef FULL
InstType "Full"
!endif
InstType "Normal"
!endif
InstType "Lite"

AutoCloseWindow true
ShowInstDetails hide
ShowUninstDetails show
SetOverwrite on
SetDateSave on

InstallDir "$PROGRAMFILES\MP3Gain\"
InstallDirRegKey HKCU "Software\VB and VBA Program Settings\MP3GainAnalysis" "InstallDir"

;This macro puts a "0" on top of the stack if the specified DLL is not up-to-date,
;or a "1" if the specified DLL _is_ up-to-date.
!macro CheckDLL WHICH_PATH DLL_NAME
  Push $0
  Push $1
  Push $2
  Push $3
  Push $4
  
  ClearErrors
  GetDLLVersionLocal vb\${DLL_NAME} $1 $2
  GetDLLVersion "${WHICH_PATH}\${DLL_NAME}" $3 $4
  IfErrors CheckDLLFail_${DLL_NAME}
  IntCmpU $1 $3 "" CheckDLLOK_${DLL_NAME} CheckDLLFail_${DLL_NAME}
  IntCmpU $2 $4 CheckDLLOK_${DLL_NAME} CheckDLLOK_${DLL_NAME}
CheckDLLFail_${DLL_NAME}:
  StrCpy $0 "0"
  Goto CheckDLLDONE_${DLL_NAME}
CheckDLLOK_${DLL_NAME}:
  StrCpy $0 "1"
CheckDLLDONE_${DLL_NAME}:
  Pop $4
  Pop $3
  Pop $2
  Pop $1
  Exch $0
!macroend

!ifdef FULL
!macro UpgradeDLL WHICH_PATH DLL_NAME REG_YN
;REG_YN is "Y" if the DLL needs to be un/registered, "N" if no registration necessary
  Push $0
  Push $1
  StrCpy $0 "${WHICH_PATH}\${DLL_NAME}"
  !insertmacro CheckDLL ${WHICH_PATH} ${DLL_NAME}
  Pop $1
  StrCmp $1 "1" noupgrade_${DLL_NAME}
;  upgrade_${DLL_NAME}:
    StrCmp ${REG_YN} "N" NoUnReg_${DLL_NAME}
    UnRegDLL $0
  NoUnReg_${DLL_NAME}:
    File /oname=$0 vb\${DLL_NAME}
    StrCmp ${REG_YN} "N" done_${DLL_NAME}
    RegDLL $0
    Goto done_${DLL_NAME}
  noupgrade_${DLL_NAME}:
    StrCmp ${REG_YN} "Y" ReRegisterDLL_${DLL_NAME} $0
    DetailPrint "${DLL_NAME} already current"
    Goto done_${DLL_NAME}
  ReRegisterDLL_${DLL_NAME}:
    RegDLL $0
  done_${DLL_NAME}:
  Pop $1
  Pop $0
!macroend
!endif

!ifndef FULL
Function .onInit
  Push $0
!insertmacro CheckDLL $SYSDIR MSCOMCTL.OCX
  Pop $0
  StrCmp $0 "0" DLLCheckFailed
  RegDLL $SYSDIR\MSCOMCTL.OCX
!insertmacro CheckDLL $SYSDIR itircl.dll
  Pop $0
  StrCmp $0 "0" DLLCheckFailed
  RegDLL $SYSDIR\itircl.dll
!insertmacro CheckDLL $SYSDIR itss.dll
  Pop $0
  StrCmp $0 "0" DLLCheckFailed
  RegDLL $SYSDIR\itss.dll
!insertmacro CheckDLL $SYSDIR msvbvm60.dll
  Pop $0
  StrCmp $0 "0" DLLCheckFailed
  RegDLL $SYSDIR\msvbvm60.dll
!insertmacro CheckDLL $SYSDIR hhctrl.ocx
  Pop $0
  StrCmp $0 "0" DLLCheckFailed
  RegDLL $SYSDIR\hhctrl.ocx
!insertmacro CheckDLL $WINDIR hh.exe
  Pop $0
  StrCmp $0 "0" DLLCheckFailed

  Goto DoneCheck

DLLCheckFailed:
  MessageBox MB_YESNO|MB_ICONEXCLAMATION 'Some Microsoft files used by MP3Gain are missing or out-of-date on your system.$\r$\nYou will probably need to go back to the MP3Gain web site and download the "Microsoft VB Run-Time files" before MP3Gain will work.$\r$\n$\r$\nWould you like to install MP3Gain right now?' IDNO InstLater
  Goto DoneCheck
InstLater:
  Abort
DoneCheck:
  Pop $0
FunctionEnd
!endif

Section "Main program files (required)"
SectionIn RO

SetOutPath $INSTDIR
SetOverwrite on
File mp3gain\MP3GainGUI.exe
File mp3gain\mp3gain.exe
!ifndef FULL
File mp3gain\README.txt
!endif
File mp3gain\MP3Gain.chm
StrCpy $9 "0"
SectionEnd


!ifdef FULL
Section "Visual Basic Runtime"
SectionIn 1

!insertmacro UpgradeDLL $SYSDIR MSCOMCTL.OCX Y

!insertmacro UpgradeDLL $SYSDIR itircl.dll Y

!insertmacro UpgradeDLL $SYSDIR itss.dll Y

!insertmacro UpgradeDLL $SYSDIR msvbvm60.dll Y

!insertmacro UpgradeDLL $SYSDIR hhctrl.ocx Y

!insertmacro UpgradeDLL $WINDIR hh.exe N

SectionEnd
!endif

Section "Links in Start menu"
!ifdef FULL
  SectionIn 1 2
!else
  SectionIn 1
!endif

CreateDirectory $SMPROGRAMS\MP3Gain
SetOutPath $INSTDIR
CreateShortCut "$SMPROGRAMS\MP3Gain\MP3Gain.lnk" "$INSTDIR\MP3GainGUI.exe"
CreateShortCut "$SMPROGRAMS\MP3Gain\MP3Gain Help.lnk" "$INSTDIR\MP3Gain.chm"
StrCpy $9 "1" ;Save adding the Uninstall shortcut until the end
SectionEnd

!ifdef DEUTSCH
  Section "Deutsche Sprachdatei"
  !ifdef FULL
    SectionIn 1 2
  !else
    SectionIn 1
  !endif
  SetOutPath $INSTDIR
  File "otherlang\Deutsch.mp3gain.ini"
  WriteRegStr HKCU "Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp" "LanguageFile" "DEUTSCH.MP3GAIN.INI"
  SectionEnd

  Section "Deutsche Hilfe"
    !ifdef FULL
      SectionIn 1 2
    !else
      SectionIn 1
    !endif
    File /r "otherlang\Help Deutsch\MP3GainDeutsch.chm"

    IntCmp $9 0 NoGermanShortCut ; Global variable set to 1 IF Links in StartMenu added
    CreateShortCut "$SMPROGRAMS\MP3Gain\MP3Gain Anleitung.lnk" "$INSTDIR\MP3GainDeutsch.chm"

    NoGermanShortCut:
  SectionEnd

  Subsection "Other Language files"

!else ifdef SPANISH
  Section "Traducción Español"
    !ifdef FULL
      SectionIn 1 2
    !else
      SectionIn 1
    !endif
    SetOutPath $INSTDIR
    File "otherlang\Español.mp3gain.ini"
    WriteRegStr HKCU "Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp" "LanguageFile" "ESPAÑOL.MP3GAIN.INI"
  SectionEnd

  Section "Ayuda (en Español)"
    !ifdef FULL
      SectionIn 1 2
    !else
      SectionIn 1
    !endif
    File /r "otherlang\Help Español\MP3GainEspañol.chm"

    IntCmp $9 0 NoEspañolShortCut ; Global variable set to 1 IF Links in StartMenu added
    CreateShortCut "$SMPROGRAMS\MP3Gain\Ayuda de MP3Gain.lnk" "$INSTDIR\MP3GainEspañol.chm"

    NoEspañolShortCut:
  SectionEnd

  Subsection "Other Language files"


!else ifdef FRENCH
  Section "Traduction Française"
    !ifdef FULL
      SectionIn 1 2
    !else
      SectionIn 1
    !endif
    SetOutPath $INSTDIR
    File "otherlang\French.mp3gain.ini"
    WriteRegStr HKCU "Software\VB and VBA Program Settings\MP3GainAnalysis\StartUp" "LanguageFile" "FRENCH.MP3GAIN.INI"
  SectionEnd

  Section "Aide (Française)"
    !ifdef FULL
      SectionIn 1 2
    !else
      SectionIn 1
    !endif
    File /r "otherlang\Help French\MP3GainFrench.chm"

    IntCmp $9 0 NoFrenchShortCut ; Global variable set to 1 IF Links in StartMenu added
    CreateShortCut "$SMPROGRAMS\MP3Gain\MP3Gain Aide.lnk" "$INSTDIR\MP3GainFrench.chm"

    NoFrenchShortCut:
  SectionEnd

  Subsection "Other Language files"


!else
  ;Not DEUTSCH or SPANISH or FRENCH
  Subsection "Language files"

!endif

Section "Bulgarian"
SetOutPath $INSTDIR
File "otherlang\Bulgarian.mp3gain.ini"
SectionEnd

Section "Català"
SetOutPath $INSTDIR
File "otherlang\Català.mp3gain.ini"
SectionEnd

Section "Chinese(Traditional)"
SetOutPath $INSTDIR
File "otherlang\Chinese(Traditional).mp3gain.ini"
SectionEnd

Section "Croatian"
SetOutPath $INSTDIR
File "otherlang\Hrvatski.mp3gain.ini"
SectionEnd

Section "Czech"
SetOutPath $INSTDIR
File "otherlang\Cesky.mp3gain.ini"
SectionEnd

Section "Dutch"
SetOutPath $INSTDIR
File "otherlang\Dutch.mp3gain.ini"
SectionEnd

Section "Finnish"
SetOutPath $INSTDIR
File "otherlang\Finnish.mp3gain.ini"
SectionEnd

!ifndef FRENCH
Section "French"
SetOutPath $INSTDIR
File "otherlang\French.mp3gain.ini"
SectionEnd
!endif

!ifndef DEUTSCH
Section "German"
SetOutPath $INSTDIR
File "otherlang\Deutsch.mp3gain.ini"
;Check if Deutsch Help file is already installed. If not, notify user in Deutsch

IfFileExists $INSTDIR\MP3GainDeutsch.chm HaveGermanHelp
MessageBox MB_ICONINFORMATION|MB_OK "Für MP3Gain ist eine deutsche Hilfedatei verfügbar, die Sie unter http://mp3gain.sourceforge.net/translation.php downloaden können"

HaveGermanHelp:

SectionEnd
!endif

Section "Hungarian"
SetOutPath $INSTDIR
File "otherlang\Magyar.mp3gain.ini"
SectionEnd

Section "Italian"
SetOutPath $INSTDIR
File "otherlang\Italian.mp3gain.ini"
SectionEnd

Section "Japanese"
SetOutPath $INSTDIR
File "otherlang\Japanese.mp3gain.ini"
SectionEnd

Section "Nederlands"
SetOutPath $INSTDIR
File "otherlang\Nederlands.mp3gain.ini"
SectionEnd

Section "Polish"
SetOutPath $INSTDIR
File "otherlang\Polish.mp3gain.ini"
SectionEnd

Section "Portuguese"
SetOutPath $INSTDIR
File "otherlang\Portugues.mp3gain.ini"
SectionEnd

Section "Romanian"
SetOutPath $INSTDIR
File "otherlang\Romanian.mp3gain.ini"
SectionEnd

Section "Russian"
SetOutPath $INSTDIR
File "otherlang\Russian.mp3gain.ini"
SectionEnd

!ifndef SPANISH
Section "Spanish"
SetOutPath $INSTDIR
File "otherlang\Español.mp3gain.ini"
SectionEnd
!endif

Section "Slovak"
SetOutPath $INSTDIR
File "otherlang\Slovensky.mp3gain.ini"
SectionEnd

Section "Turkish"
SetOutPath $INSTDIR
File "otherlang\Turkish.mp3gain.ini"
SectionEnd

Section "Uzbek"
SetOutPath $INSTDIR
File "otherlang\Uzbek.mp3gain.ini"
SectionEnd

SubsectionEnd 

Section -post

  SetOutPath $INSTDIR

; since the installer is now created last (in 1.2+), this makes sure 
; that any old installer that is readonly is overwritten.
  Delete $INSTDIR\uninst-mp3gain.exe 
  WriteUninstaller $INSTDIR\uninst-mp3gain.exe

  IntCmp $9 0 NoUninstallShortCut ; Global set to 1 IF Links in StartMenu selected
  CreateShortCut "$SMPROGRAMS\MP3Gain\Uninstall MP3Gain.lnk" "$INSTDIR\uninst-mp3gain.exe"

  NoUninstallShortCut:

SectionEnd

  
Function .onInstSuccess
  WriteRegStr HKCU "Software\VB and VBA Program Settings\MP3GainAnalysis" "InstallDir" $INSTDIR
  WriteRegStr HKCR ".m3g" "" "MP3GainAnalysisResults"
  WriteRegStr HKCR "MP3GainAnalysisResults" "" "MP3Gain Analysis Results"
  WriteRegStr HKCR "MP3GainAnalysisResults\shell" "" "open"
  WriteRegStr HKCR "MP3GainAnalysisResults\shell\open\command" "" `"$INSTDIR\MP3GainGUI.exe" "%1"`
  MessageBox MB_OK \
             'MP3Gain is now installed in $INSTDIR'
FunctionEnd

Section Uninstall
  
  IfFileExists "$INSTDIR\MP3GainGUI.exe" skip_confirmation
    MessageBox MB_YESNO `MP3Gain ${VER_MAJOR}${VER_MINOR}${VER_REL} might not be installed properly in "$INSTDIR"$\r$\nWould you still like to try to uninstall (not recommended)?` IDYES skip_confirmation
    Abort "Cancelled Uninstall"
  skip_confirmation:
  
  SetShellVarContext current
  
  Delete $SMPROGRAMS\MP3Gain\*.*
  RMDir $SMPROGRAMS\MP3Gain
  
;  RMDir /r $INSTDIR\MP3Gain
  Delete $INSTDIR\MP3GainGUI.exe
  Delete $INSTDIR\MP3Gain.chm
  Delete $INSTDIR\mp3gain.exe
  Delete $INSTDIR\README.txt
  Delete $INSTDIR\uninst-mp3gain.exe
  Delete $INSTDIR\*.mp3gain.ini
!ifdef DEUTSCH
  Delete "$INSTDIR\MP3GainDeutsch.chm"
!endif
!ifdef SPANISH
  Delete "$INSTDIR\MP3GainEspañol.chm"
!endif
!ifdef FRENCH
  Delete "$INSTDIR\MP3GainFrench.chm"
!endif
  ; if $INSTDIR was removed, skip these next ones
  IfFileExists $INSTDIR 0 Removed 
;    MessageBox MB_YESNO|MB_ICONQUESTION \
;      `Would you like to delete any remaining files from "$INSTDIR"? (Click "No" if want to keep additional files you've installed in that folder)` IDNO Removed
;    Delete $INSTDIR\*.* ; this would be skipped if the user hits no
    SetOutPath $TEMP
;    RMDir /r $INSTDIR
    RMDir $INSTDIR
    IfFileExists $INSTDIR 0 Removed 
      MessageBox MB_OK|MB_ICONEXCLAMATION \
                 "Warning: $INSTDIR could not be removed. There are extra files in this folder. You will need to delete them yourself."
  Removed:
  
  MessageBox MB_YESNO|MB_ICONQUESTION \
    `Would you like to delete the MP3Gain Registry keys? (Click "No" if you plan on re-installing MP3Gain and you want to keep your preferences)` IDNO RegRemoved
  DeleteRegKey HKCU "Software\VB and VBA Program Settings\MP3GainAnalysis"
  ReadRegStr $R1 HKCR ".m3g" ""
  StrCmp $R1 "MP3GainAnalysisResults" 0 DoNotDeleteM3GKey ;don't delete .m3g file association if it's pointing to something else

  DeleteRegKey HKCR ".m3g"

DoNotDeleteM3GKey:
  DeleteRegKey HKCR "MP3GainAnalysisResults"
  ClearErrors
  RegRemoved:

SectionEnd

