; DacTime_setup.nsi — Installateur NSIS pour DacTime
; Prérequis : NSIS 3.x installé (https://nsis.sourceforge.io)
; Compiler  : makensis DacTime_setup.nsi
; Sortie    : dist\DacTime_Setup_1.2.exe

!include "MUI2.nsh"

; ================================================================ INFOS
Name              "DacTime"
OutFile           "dist\DacTime_Setup_1.2.exe"
InstallDir        "$PROGRAMFILES64\DacTime"
InstallDirRegKey  HKCU "Software\DacTime" "InstallDir"
RequestExecutionLevel admin
SetCompressor     /SOLID lzma
Unicode           True

; ================================================================ INTERFACE
!define MUI_ICON              "dactime.ico"
!define MUI_UNICON             "dactime.ico"
!define MUI_WELCOMEFINISHPAGE_BITMAP_NOSTRETCH
!define MUI_ABORTWARNING

!define MUI_WELCOMEPAGE_TITLE    "Installation de DacTime 1.2"
!define MUI_WELCOMEPAGE_TEXT     "DacTime est une application de rédaction$\net d'impression de courriers postaux.$\n$\nConçu par Clément LATTAR."

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "French"

; ================================================================ INSTALLATION
Section "DacTime" SecMain
    SectionIn RO

    ; Desinstaller version precedente si elle existe
    ReadRegStr $0 HKCU "Software\DacTime" "InstallDir"
    StrCmp $0 "" install_direct
        IfFileExists "$0\Uninstall.exe" do_uninstall install_direct
        do_uninstall:
            DetailPrint "Desinstallation version precedente..."
            ExecWait '"$0\Uninstall.exe" /S'
            Sleep 1000
    install_direct:

    SetOutPath "$INSTDIR"

    ; Fichier principal
    File "dist\DacTime\DacTime.exe"

    ; Dossier interne PyInstaller
    File /r "dist\DacTime\_internal"

    ; Icône
    File "dactime.ico"

    ; Désinstallateur
    WriteUninstaller "$INSTDIR\Uninstall.exe"

    ; Raccourci Menu Démarrer
    CreateDirectory "$SMPROGRAMS\DacTime"
    CreateShortcut "$SMPROGRAMS\DacTime\DacTime.lnk" \
        "$INSTDIR\DacTime.exe" "" "$INSTDIR\dactime.ico" 0
    CreateShortcut "$SMPROGRAMS\DacTime\Désinstaller DacTime.lnk" \
        "$INSTDIR\Uninstall.exe"

    ; Raccourci Bureau
    CreateShortcut "$DESKTOP\DacTime.lnk" \
        "$INSTDIR\DacTime.exe" "" "$INSTDIR\dactime.ico" 0

    ; Registre — infos programme (Ajout/Suppression de programmes)
    WriteRegStr HKCU "Software\DacTime" "InstallDir" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "DisplayName" "DacTime"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "DisplayVersion" "1.2"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "Publisher" "Clément LATTAR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "DisplayIcon" "$INSTDIR\dactime.ico"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime" \
        "NoRepair" 1

    ; Association extension .dactime
    WriteRegStr HKCU "Software\Classes\.dactime" "" "DacTime.Document"
    WriteRegStr HKCU "Software\Classes\DacTime.Document" "" "Courrier DacTime"
    WriteRegStr HKCU "Software\Classes\DacTime.Document\DefaultIcon" \
        "" "$INSTDIR\dactime.ico,0"
    WriteRegStr HKCU "Software\Classes\DacTime.Document\shell\open\command" \
        "" '"$INSTDIR\DacTime.exe" "%1"'

SectionEnd

; ================================================================ DÉSINSTALLATION
Section "Uninstall"

    ; Fichiers
    Delete "$INSTDIR\DacTime.exe"
    Delete "$INSTDIR\dactime.ico"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir  /r "$INSTDIR\_internal"
    RMDir  "$INSTDIR"

    ; Raccourcis
    Delete "$DESKTOP\DacTime.lnk"
    Delete "$SMPROGRAMS\DacTime\DacTime.lnk"
    Delete "$SMPROGRAMS\DacTime\Désinstaller DacTime.lnk"
    RMDir  "$SMPROGRAMS\DacTime"

    ; Registre
    DeleteRegKey HKCU "Software\DacTime"
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DacTime"
    DeleteRegKey HKCU "Software\Classes\.dactime"
    DeleteRegKey HKCU "Software\Classes\DacTime.Document"

SectionEnd
