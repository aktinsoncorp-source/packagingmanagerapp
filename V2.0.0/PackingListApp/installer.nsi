Unicode true
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "FileFunc.nsh"
!include "x64.nsh"
!ifndef APPDIR
  !error "APPDIR must point to the verified packaged application."
!endif
!ifndef OUTFILE
  !error "OUTFILE must name the installer to build."
!endif
!ifndef UNINSTALL_FILES
  !error "UNINSTALL_FILES must name the generated exact-file removal list."
!endif

Name "Packing List Manager"
OutFile "${OUTFILE}"
InstallDir "$PROGRAMFILES64\Packing List Manager"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
SetCompressorDictSize 32
ManifestDPIAware true
VIProductVersion "2.0.0.0"
VIAddVersionKey "ProductName" "Packing List Manager"
VIAddVersionKey "ProductVersion" "2.0.0"
VIAddVersionKey "FileVersion" "2.0.0"
VIAddVersionKey "FileDescription" "Packing List Manager Setup"
VIAddVersionKey "LegalCopyright" "Packing List Manager"
ShowInstDetails show
ShowUninstDetails show
BrandingText "Packing List Manager 2.0.0"
Var NoIntegration
Var Parameters

!define MUI_ICON "${APPDIR}\Packing List Manager.ico"
!define MUI_UNICON "${APPDIR}\Packing List Manager.ico"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "Packing List Manager 2.0.0"
!define MUI_WELCOMEPAGE_TEXT "This will install or reinstall Packing List Manager in Program Files for all users.$\r$\n$\r$\nPython is included. Microsoft Word desktop is required for packing PDFs. Microsoft Excel desktop is needed for live Excel editing.$\r$\n$\r$\nSaved names and settings are kept in each user's Local AppData folder. Existing data and packing documents will be preserved. Close the app before reinstalling."
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TEXT "Packing List Manager is ready.$\r$\n$\r$\nOpen Packing List Manager V2 from your desktop or Start menu, then choose your usual packing folder and date."
!insertmacro MUI_PAGE_FINISH
!define MUI_UNCONFIRMPAGE_TEXT_TOP "Remove Packing List Manager program files? Your saved product definitions, SKU catalog, source PDFs, outputs and history will be kept."
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH
!insertmacro MUI_LANGUAGE "English"

Function .onInit
  SetShellVarContext all
  SetRegView 64
  ${IfNot} ${RunningX64}
    MessageBox MB_OK|MB_ICONSTOP "Packing List Manager requires 64-bit Windows." /SD IDOK
    SetErrorLevel 2
    Abort
  ${EndIf}
  StrCpy $NoIntegration 0
  ${GetParameters} $Parameters
  ClearErrors
  ${GetOptions} $Parameters "/NOINTEGRATION" $0
  ${IfNot} ${Errors}
    StrCpy $NoIntegration 1
  ${EndIf}
  ${If} $NoIntegration == 0
    System::Call 'kernel32::OpenMutexW(i 0x100000, i 0, w "Local\PackingListManagerV2Python") p.r0'
    ${If} $0 != 0
      System::Call 'kernel32::CloseHandle(p r0)'
      MessageBox MB_OK|MB_ICONEXCLAMATION "Close Packing List Manager V2 before installing this update." /SD IDOK
      SetErrorLevel 2
      Abort
    ${EndIf}
  ${EndIf}
FunctionEnd

Section "Packing List Manager" MainSection
  SectionIn RO
  SetOverwrite on
  SetOutPath "$INSTDIR\App Files"
  File /r "${APPDIR}\App Files\*"
  SetOutPath "$INSTDIR"
  File "${APPDIR}\Packing List Manager.exe"
  File "${APPDIR}\Packing List Manager.ico"
  File "${APPDIR}\Employee Guide.txt"
  File "${APPDIR}\Release Notes.txt"
  File "${APPDIR}\Package Manifest.json"
  ; Normal installations seed per-user writable data on first launch. Existing
  ; catalogs beside an earlier executable are retained as migration sources.
  ${If} $NoIntegration == 1
    SetOverwrite off
    File "${APPDIR}\ProductNames.json"
    File "${APPDIR}\SkuCatalog.json"
    SetOverwrite on
    WriteINIStr "$INSTDIR\App Files\InstallMode.ini" "Setup" "DataMode" "portable"
  ${Else}
    WriteINIStr "$INSTDIR\App Files\InstallMode.ini" "Setup" "DataMode" "user"
  ${EndIf}
  WriteINIStr "$INSTDIR\App Files\InstallMode.ini" "Setup" "NoIntegration" "$NoIntegration"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  ${If} $NoIntegration == 0
    WriteRegStr HKLM "Software\PackingListManagerV2" "InstallDir" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "DisplayName" "Packing List Manager"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "DisplayVersion" "2.0.0"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "DisplayIcon" "$INSTDIR\Packing List Manager.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "UninstallString" '$\"$INSTDIR\Uninstall.exe$\"'
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "QuietUninstallString" '$\"$INSTDIR\Uninstall.exe$\" /S'
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2" "NoRepair" 1
    CreateDirectory "$SMPROGRAMS\Packing List Manager V2"
    CreateShortcut "$SMPROGRAMS\Packing List Manager V2\Packing List Manager.lnk" "$INSTDIR\Packing List Manager.exe"
    CreateShortcut "$SMPROGRAMS\Packing List Manager V2\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
    CreateShortcut "$DESKTOP\Packing List Manager V2.lnk" "$INSTDIR\Packing List Manager.exe"
  ${EndIf}
SectionEnd

Function un.onInit
  SetShellVarContext all
  SetRegView 64
  IfFileExists "$INSTDIR\App Files\Runtime Manifest.json" +4
    MessageBox MB_OK|MB_ICONSTOP "The application folder could not be verified. No files were removed." /SD IDOK
    SetErrorLevel 2
    Abort
  ReadINIStr $NoIntegration "$INSTDIR\App Files\InstallMode.ini" "Setup" "NoIntegration"
  ${If} $NoIntegration != 1
    System::Call 'kernel32::OpenMutexW(i 0x100000, i 0, w "Local\PackingListManagerV2Python") p.r0'
    ${If} $0 != 0
      System::Call 'kernel32::CloseHandle(p r0)'
      MessageBox MB_OK|MB_ICONEXCLAMATION "Close Packing List Manager V2 before uninstalling." /SD IDOK
      SetErrorLevel 2
      Abort
    ${EndIf}
  ${EndIf}
FunctionEnd

Section "Uninstall"
  ; Exact installed program files only: no recursive deletion or user-data deletion.
  !include "${UNINSTALL_FILES}"
  Delete "$INSTDIR\App Files\InstallMode.ini"
  RMDir "$INSTDIR\App Files"
  Delete "$INSTDIR\Uninstall.exe"
  ${If} $NoIntegration != 1
    ReadRegStr $0 HKLM "Software\PackingListManagerV2" "InstallDir"
    ${If} $0 == $INSTDIR
      Delete "$DESKTOP\Packing List Manager V2.lnk"
      Delete "$SMPROGRAMS\Packing List Manager V2\Packing List Manager.lnk"
      Delete "$SMPROGRAMS\Packing List Manager V2\Uninstall.lnk"
      RMDir "$SMPROGRAMS\Packing List Manager V2"
      DeleteRegKey HKLM "Software\PackingListManagerV2"
      DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PackingListManagerV2"
    ${EndIf}
  ${EndIf}
  RMDir "$INSTDIR"
SectionEnd
