!addplugindir "plugins"
!include "LogicLib.nsh"
!include "Sections.nsh"
!include "nsDialogs.nsh"
!include "MUI2.nsh" 
!include "MUI.nsh"
!include "WinVer.nsh"
!include "x64.nsh"
 
 VIProductVersion "1.1.0.0"
  VIAddVersionKey /LANG=1033 "ProductName" "SALMA for MS Word"
  VIAddVersionKey /LANG=1033 "Comments" "SALMA for MS Word"
  VIAddVersionKey /LANG=1033 "CompanyName" "SoftServe Inc."
  VIAddVersionKey /LANG=1033 "LegalTrademarks" "SALMA is registered trademark of SoftServe Ltd."
  VIAddVersionKey /LANG=1033 "LegalCopyright" "SoftServe Ltd."
  VIAddVersionKey /LANG=1033 "FileDescription" "SALMA Installer"
  VIAddVersionKey /LANG=1033 "FileVersion" "1.1"
  VIAddVersionKey /LANG=1033 "ProductVersion" "1.1"

; Compressor Setup
SetCompressor /SOLID lzma
SetCompressorDictSize 12

  !define PRODUCT_PUBLISHER "SoftServe Ltd."

; MUI settings
	
  !define MUI_COMPONENTSPAGE_TEXT_COMPLIST "$(COMPONENTS_TEXT)"
  !define MUI_HEADERIMAGE
  !define MUI_HEADERIMAGE_RIGHT
  !define MUI_HEADERIMAGE_BITMAP "BannerBmp.bmp"
  !define MUI_HEADERIMAGE_UNBITMAP "BannerBmp.bmp"
  !define MUI_ICON "icon.ico"
  !define MUI_UNICON  "icon.ico"
; Welcome bitmaps
!define MUI_WELCOMEFINISHPAGE_BITMAP "DialogBmp.bmp"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP "DialogBmp.bmp";

;Variables
Var IfVstoRuntimeChecked
Var SALMAStatus

;Branding text
BrandingText "SoftServe Ltd."

!define MUI_COMPONENTSPAGE_SMALLDESC

;UI wizard 
;Interface Settings

  !define MUI_ABORTWARNING 

;Install Pages
Page custom ApplicatuionExistsDialog ApplicatuionExistsDialogLeave
!define MUI_WELCOMEPAGE_TITLE_3LINES 
!insertmacro MUI_PAGE_WELCOME
;!insertmacro MUI_PAGE_LICENSE "License.txt"

!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TITLE_3LINES
!insertmacro MUI_PAGE_FINISH

; Uninstall Pages
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Language Files
  !insertmacro MUI_LANGUAGE "English"
  !insertmacro MUI_LANGUAGE "Russian"  
  !insertmacro MUI_LANGUAGE "Ukrainian"  
  !insertmacro MUI_RESERVEFILE_LANGDLL

LangString COMPONENTS_TEXT ${LANG_ENGLISH} "List of necessary components:"
LangString COMPONENTS_TEXT ${LANG_RUSSIAN} "Список необходимых компонентов:"
LangString COMPONENTS_TEXT ${LANG_UKRAINIAN} "Список необхідних компонентів:"
  
LangString PRODUCT_NAME_STRING ${LANG_ENGLISH} "SALMA for MS Word"
LangString PRODUCT_NAME_STRING ${LANG_RUSSIAN} "SALMA для MS Word"
LangString PRODUCT_NAME_STRING ${LANG_UKRAINIAN} "SALMA для MS Word"
LangString PRODUCT_NAME_SHORT ${LANG_ENGLISH} ""
LangString PRODUCT_NAME_SHORT ${LANG_RUSSIAN} ""
LangString PRODUCT_NAME_SHORT ${LANG_UKRAINIAN} ""

LangString INSTALLING_SALMA ${LANG_ENGLISH} "Installing SALMA..."
LangString INSTALLED_SALMA ${LANG_ENGLISH} "SALMA was successfully installed."
LangString REMOVING_TEMPORARIES ${LANG_ENGLISH} "Removing temporary files..."
LangString INSTALLING_TFS_OBJECT_MODEL ${LANG_ENGLISH} "Installing TFS Object Model temporary files..."
LangString INSTALLED_TFS_OBJECT_MODEL ${LANG_ENGLISH} "TFS Object Model was successfully installed."
LangString PREREQUISITES ${LANG_ENGLISH} "Prerequisites"

LangString INSTALLING_SALMA ${LANG_RUSSIAN} "Идет установка SALMA..."
LangString INSTALLED_SALMA ${LANG_RUSSIAN} "SALMA успешно установлен."
LangString REMOVING_TEMPORARIES ${LANG_RUSSIAN} "Удаление временных файлов..."
LangString INSTALLING_TFS_OBJECT_MODEL ${LANG_RUSSIAN} "Установка TFS Object Model..."
LangString INSTALLED_TFS_OBJECT_MODEL ${LANG_RUSSIAN} "TFS Object Model успешно установлен."
LangString PREREQUISITES ${LANG_RUSSIAN} "Пререквизиты"

LangString INSTALLING_SALMA ${LANG_UKRAINIAN} "Встановлення SALMA..."
LangString INSTALLED_SALMA ${LANG_UKRAINIAN} "SALMA успішно встановлено."
LangString REMOVING_TEMPORARIES ${LANG_UKRAINIAN} "Видалення тимчасових файлів..."
LangString INSTALLING_TFS_OBJECT_MODEL ${LANG_UKRAINIAN} "Встановлення TFS Object Model..."
LangString INSTALLED_TFS_OBJECT_MODEL ${LANG_UKRAINIAN} "TFS Object Model успішно встановлено."
LangString PREREQUISITES ${LANG_UKRAINIAN} "Prerequisites"


LangString DESC_REMAINING ${LANG_ENGLISH} " (%d %s%s remaining)"
LangString DESC_PROGRESS ${LANG_ENGLISH} "%d.%01dkB/s" ;"%dkB (%d%%) of %dkB @ %d.%01dkB/s"
LangString DESC_PLURAL ${LANG_ENGLISH} "s"
LangString DESC_HOUR ${LANG_ENGLISH} "hour"
LangString DESC_MINUTE ${LANG_ENGLISH} "minute"
LangString DESC_SECOND ${LANG_ENGLISH} "second"
LangString DESC_CONNECTING ${LANG_ENGLISH} "Connecting..."
LangString DESC_DOWNLOADING ${LANG_ENGLISH} "Downloading %s"
LangString DESC_SHORTDOTNET ${LANG_ENGLISH} "Microsoft VSTO Runtime"
LangString DESC_LONGDOTNET ${LANG_ENGLISH} "Microsoft VSTO Runtime"
LangString DESC_DOTNET_DECISION ${LANG_ENGLISH} "$(DESC_SHORTDOTNET) is required.$\nIt is strongly advised that you install $(DESC_SHORTDOTNET) before continuing.$\nIf you choose to continue, you will need to connect to the internet before proceeding.$\nWould you like to continue with the installation?"
LangString SEC_DOTNET ${LANG_ENGLISH} "$(DESC_SHORTDOTNET) "
LangString DESC_INSTALLING ${LANG_ENGLISH} "Installing..."
LangString DESC_DOWNLOADING1 ${LANG_ENGLISH} "Downloading..."
LangString DESC_DOWNLOADFAILED ${LANG_ENGLISH} "Download Failed:"
LangString ERROR_DOTNET_DUPLICATE_INSTANCE ${LANG_ENGLISH} "The $(DESC_SHORTDOTNET) Installer is already running."
LangString ERROR_NOT_ADMINISTRATOR ${LANG_ENGLISH} "Not Administrator"
LangString ERROR_INVALID_PLATFORM ${LANG_ENGLISH} "Invalid Platform"
LangString DESC_DOTNET_TIMEOUT ${LANG_ENGLISH} "The installation of the $(DESC_SHORTDOTNET) has timed out."
LangString ERROR_DOTNET_INVALID_PATH ${LANG_ENGLISH} "The $(DESC_SHORTDOTNET) Installation$\nwas not found in the following location:$\n"
LangString ERROR_DOTNET_FATAL ${LANG_ENGLISH} "A fatal error occurred during the installation$\nof the $(DESC_SHORTDOTNET)."
LangString FAILED_DOTNET_INSTALL ${LANG_ENGLISH} "The installation of $(PRODUCT_NAME_STRING) will$\n\  continue. However, it may not function properly$\nuntil $(DESC_SHORTDOTNET)$\nis installed."

LangString DESC_REMAINING ${LANG_RUSSIAN} " (%d %s%s осталось)"
LangString DESC_PROGRESS ${LANG_RUSSIAN} "%d.%01dkB/s" ;"%dkB (%d%%) of %dkB @ %d.%01dkB/s"
LangString DESC_PLURAL ${LANG_RUSSIAN} "s"
LangString DESC_HOUR ${LANG_RUSSIAN} "часов"
LangString DESC_MINUTE ${LANG_RUSSIAN} "минут"
LangString DESC_SECOND ${LANG_RUSSIAN} "секунд"
LangString DESC_CONNECTING ${LANG_RUSSIAN} "Соединение..."
LangString DESC_DOWNLOADING ${LANG_RUSSIAN} "Загзузка %s"
LangString DESC_SHORTDOTNET ${LANG_RUSSIAN} "Microsoft VSTO Runtime"
LangString DESC_LONGDOTNET ${LANG_RUSSIAN} "Microsoft VSTO Runtime"
LangString DESC_DOTNET_DECISION ${LANG_RUSSIAN} "$(DESC_SHORTDOTNET) требуетса.$\nНастоятельно реклмендуем $(DESC_SHORTDOTNET) прежде чем$\nпродолжить установку SALMA.$\nЕсли Вы продолжите будет соединение с сетью интернет перед продолжением установки.$\nХотите продолжить установку?"
LangString SEC_DOTNET ${LANG_RUSSIAN} "$(DESC_SHORTDOTNET) "
LangString DESC_INSTALLING ${LANG_RUSSIAN} "Установка..."
LangString DESC_DOWNLOADING1 ${LANG_RUSSIAN} "Загрузка..."
LangString DESC_DOWNLOADFAILED ${LANG_RUSSIAN} "Загзузка не удалась:"
LangString ERROR_DOTNET_DUPLICATE_INSTANCE ${LANG_RUSSIAN} "$(DESC_SHORTDOTNET) установщик уже запушен."
LangString ERROR_NOT_ADMINISTRATOR ${LANG_RUSSIAN} "Нет пправ администратора"
LangString ERROR_INVALID_PLATFORM ${LANG_RUSSIAN} "Не поддерживаемая платформа"
LangString DESC_DOTNET_TIMEOUT ${LANG_RUSSIAN} "Установка $(DESC_SHORTDOTNET) превысила лимит времени."
LangString ERROR_DOTNET_INVALID_PATH ${LANG_RUSSIAN} "$(DESC_SHORTDOTNET) установщик$\n\  не был найден:$\n"
LangString ERROR_DOTNET_FATAL ${LANG_RUSSIAN} "Критическая ошибка во время установки$\n\  $(DESC_SHORTDOTNET)."
LangString FAILED_DOTNET_INSTALL ${LANG_RUSSIAN} "Установка$(PRODUCT_NAME_STRING) будет$\n проболжена. Однако, он может функционировать не праывильно пока не будет установлен $(DESC_SHORTDOTNET)."

LangString DESC_REMAINING ${LANG_UKRAINIAN} " (%d %s%s залишилось)"
LangString DESC_PROGRESS ${LANG_UKRAINIAN} "%d.%01dkB/s" ;"%dkB (%d%%) of %dkB @ %d.%01dkB/s"
LangString DESC_PLURAL ${LANG_UKRAINIAN} "s"
LangString DESC_HOUR ${LANG_UKRAINIAN} "годин"
LangString DESC_MINUTE ${LANG_UKRAINIAN} "хвилин"
LangString DESC_SECOND ${LANG_UKRAINIAN} "секунд"
LangString DESC_CONNECTING ${LANG_UKRAINIAN} "Підключення..."
LangString DESC_DOWNLOADING ${LANG_UKRAINIAN} "Завантаження %s"
LangString DESC_SHORTDOTNET ${LANG_UKRAINIAN} "Microsoft VSTO Runtime"
LangString DESC_LONGDOTNET ${LANG_UKRAINIAN} "Microsoft VSTO Runtime"
LangString DESC_DOTNET_DECISION ${LANG_UKRAINIAN} "$(DESC_SHORTDOTNET) необхідний.$\nРекомендуємо встановити $(DESC_SHORTDOTNET) перед встановленням SALMA.$\nЯкщо Ви продовжите, буде встанролене зєднання з мередею інтернет.$\nБажаєте продовжити встановлення?"
LangString SEC_DOTNET ${LANG_UKRAINIAN} "$(DESC_SHORTDOTNET) "
LangString DESC_INSTALLING ${LANG_UKRAINIAN} "Встановлення..."
LangString DESC_DOWNLOADING1 ${LANG_UKRAINIAN} "Завантаження..."
LangString DESC_DOWNLOADFAILED ${LANG_UKRAINIAN} "Завантаження не вдалось:"
LangString ERROR_DOTNET_DUPLICATE_INSTANCE ${LANG_UKRAINIAN} "Встановлення $(DESC_SHORTDOTNET) вже почалось."
LangString ERROR_NOT_ADMINISTRATOR ${LANG_UKRAINIAN} "Немає прав адміністратора"
LangString ERROR_INVALID_PLATFORM ${LANG_UKRAINIAN} "Поатформа не підтримується"
LangString DESC_DOTNET_TIMEOUT ${LANG_UKRAINIAN} "Встановлення $(DESC_SHORTDOTNET) перевищило час очікування."
LangString ERROR_DOTNET_INVALID_PATH ${LANG_UKRAINIAN} "Інсталятор $(DESC_SHORTDOTNET) $\nне був знайдений:$\n"
LangString ERROR_DOTNET_FATAL ${LANG_UKRAINIAN} "Критична помилка під час встановлення$\n$(DESC_SHORTDOTNET)."
LangString FAILED_DOTNET_INSTALL ${LANG_UKRAINIAN} "Встановлення $(PRODUCT_NAME_STRING) продовжиться.$\nОднак коректне функціонування не гарантується допоки не буде встановлено $(DESC_SHORTDOTNET)."

  
LangString DESC_SEC01 ${LANG_ENGLISH} "SALMA Word Add-in"
LangString DESC_SEC02 ${LANG_ENGLISH} "Microsoft VSTO Runtime Package"
LangString DESC_SECGR01 ${LANG_ENGLISH} "Additional required Components for SALMA"

LangString DESC_SEC01 ${LANG_RUSSIAN} "SALMA Word Add-in"
LangString DESC_SEC02 ${LANG_RUSSIAN} "Microsoft VSTO Runtime Package"
LangString DESC_SECGR01 ${LANG_RUSSIAN} "Допольнительные необходимые компоненты для SALMA"

LangString DESC_SEC01 ${LANG_UKRAINIAN} "SALMA Word Add-in"
LangString DESC_SEC02 ${LANG_UKRAINIAN} "Microsoft VSTO Runtime Package"
LangString DESC_SECGR01 ${LANG_UKRAINIAN} "Додаткові необхідні компоненти для SALMA"

LangString DEINSTALLATION_MESSAGE_BEFORE ${LANG_ENGLISH} "Removing files"
LangString DEINSTALLATION_MESSAGE_AFTER ${LANG_ENGLISH} "Uninstallation is complete"

LangString DEINSTALLATION_MESSAGE_BEFORE ${LANG_RUSSIAN} "Удаление начато"
LangString DEINSTALLATION_MESSAGE_AFTER ${LANG_RUSSIAN} "Завершение удаления"

LangString DEINSTALLATION_MESSAGE_BEFORE ${LANG_UKRAINIAN} "Видалення файлів"
LangString DEINSTALLATION_MESSAGE_AFTER ${LANG_UKRAINIAN} "Видалення завершено"

LangString ABORTING_INSTALLATION ${LANG_ENGLISH} "Aborting Installation , critical errors occurred. Please try again."
LangString ABORTING_INSTALLATION ${LANG_RUSSIAN} "Завершение установки в связи с ошибкои во время инсталяции. пожалуйста повторите попытку"
LangString ABORTING_INSTALLATION ${LANG_UKRAINIAN} "Завершення встановлення в зв'язку з помилкою.Будь ласка повторіть спробу."

LangString SALMA_NAME ${LANG_ENGLISH} "SALMA Word Add-in"
LangString SALMA_NAME ${LANG_RUSSIAN} "SALMA Word Add-in"
LangString SALMA_NAME ${LANG_UKRAINIAN} "SALMA Word Add-in"

LangString VSTOR_NAME ${LANG_ENGLISH} "VSTO Runtime"
LangString VSTOR_NAME ${LANG_RUSSIAN} "VSTO Runtime"
LangString VSTOR_NAME ${LANG_UKRAINIAN} "VSTO Runtime"

LangString NOT_INSTALLED ${LANG_ENGLISH} "(Not Installed)"
LangString NOT_INSTALLED ${LANG_RUSSIAN} "(Не Установлено)"	
LangString NOT_INSTALLED ${LANG_UKRAINIAN} "(Не Встановлено)"	

LangString INSTALLED ${LANG_ENGLISH} "(Installed)"
LangString INSTALLED ${LANG_RUSSIAN} "(Установлено)"	
LangString INSTALLED ${LANG_UKRAINIAN} "(Встановлено)"	

!define SETUP_NAME "Setup.exe"
!define PRODUCT_NAME "$(PRODUCT_NAME_STRING)"
!define VERSION "1.0.0"

!define REG_UNINSTALL "SOFTWARE\Microsoft\Office\Word\AddIns\SoftServe.SALMA"
########### Installer Configurations ####################
Name "$(PRODUCT_NAME_STRING)"
OutFile ${SETUP_NAME}
InstallDir "$PROGRAMFILES\SoftServe\SALMA for Word"

RequestExecutionLevel Admin

ShowInstDetails show
ShowUnInstDetails show 

# create a default section.
Section
	# define output path
	SetOutPath $INSTDIR
	#define uninstaller name
	WriteUninstaller $INSTDIR\uninstaller.exe
SectionEnd

!define BASE_URL http://download.microsoft.com/download
!define URL_VSTO_COMMON "${BASE_URL}/9/4/9/949B0B7C-6385-4664-8EA8-3F6038172322/vstor_redist.exe"

Var "URL_VSTO"
Var "DOTNET_RETURN_CODE"	
	Section "$(VSTOR_NAME) $(NOT_INSTALLED)" SEC02		
	${If} $IfVstoRuntimeChecked == 1	
		; SectionIn RO		
		IfSilent lbl_IsSilent
		!define DOTNETFILESDIR "Common\Files\MSNET"
		StrCpy $DOTNET_RETURN_CODE "0"		
	!ifdef DOTNET_ONCD_1033
		StrCmp "$OSLANGUAGE" "1033" 0 lbl_Not1033
		SetOutPath "$PLUGINSDIR"
		File /r "${DOTNETFILESDIR}\vsto_redist.exe"
		DetailPrint "$(DESC_INSTALLING) $(DESC_SHORTDOTNET)..."
		Banner::show /NOUNLOAD "$(DESC_INSTALLING) $(DESC_SHORTDOTNET)..."
		ClearErrors
		nsExec::ExecToStack '"$PLUGINSDIR\vsto_redist.exe" /q /c:"install.exe /norestart /noaspupgrade /q"'
		ClearErrors
		pop $DOTNET_RETURN_CODE
		Banner::destroy		
		Goto lbl_NoDownloadRequired
		lbl_Not1033:
!endif		
    Goto lbl_DownloadRequired
    lbl_DownloadRequired:
    DetailPrint "$(DESC_DOWNLOADING1) $(DESC_SHORTDOTNET)..."
    ; MessageBox MB_ICONEXCLAMATION|MB_YESNO|MB_DEFBUTTON2 "$(DESC_DOTNET_DECISION)" /SD IDNO \
      ; IDYES +2 IDNO 0
    ; Abort
    ; ; "Downloading Microsoft .Net Framework"
    AddSize 153600
    nsisdl::download /TRANSLATE "$(DESC_DOWNLOADING)" "$(DESC_CONNECTING)" \
       "$(DESC_SECOND)" "$(DESC_MINUTE)" "$(DESC_HOUR)" "$(DESC_PLURAL)" \
       "$(DESC_PROGRESS)" "$(DESC_REMAINING)" \
       /TIMEOUT=30000 "$URL_VSTO" "$PLUGINSDIR\vsto_redist.exe"
    Pop $0
    StrCmp "$0" "success" lbl_continue
    DetailPrint "$(DESC_DOWNLOADFAILED) $0"
    Abort
 
    lbl_continue:
      DetailPrint "$(DESC_INSTALLING) $(DESC_SHORTDOTNET)..."
      Banner::show /NOUNLOAD "$(DESC_INSTALLING) $(DESC_SHORTDOTNET)..."
	  ClearErrors
      nsExec::ExecToStack '"$PLUGINSDIR\vsto_redist.exe" /norestart /q /c:"install.exe /noaspupgrade /q"'
      ClearErrors
	  pop $DOTNET_RETURN_CODE
      Banner::destroy
      SetRebootFlag false
      ; silence the compiler
      Goto lbl_NoDownloadRequired
      lbl_NoDownloadRequired:
  
      StrCmp "$DOTNET_RETURN_CODE" "" lbl_NoError
      StrCmp "$DOTNET_RETURN_CODE" "0" lbl_NoError
      StrCmp "$DOTNET_RETURN_CODE" "3010" lbl_NoError
      StrCmp "$DOTNET_RETURN_CODE" "8192" lbl_NoError
      StrCmp "$DOTNET_RETURN_CODE" "error" lbl_Error
      StrCmp "$DOTNET_RETURN_CODE" "timeout" lbl_TimeOut
      StrCmp "$DOTNET_RETURN_CODE" "4101" lbl_Error_DuplicateInstance
      StrCmp "$DOTNET_RETURN_CODE" "4097" lbl_Error_NotAdministrator
      StrCmp "$DOTNET_RETURN_CODE" "1633" lbl_Error_InvalidPlatform lbl_FatalError
      ; all others are fatal
 
    lbl_Error_DuplicateInstance:
    DetailPrint "$(ERROR_DOTNET_DUPLICATE_INSTANCE)"
    GoTo lbl_Done
 
    lbl_Error_NotAdministrator:
    DetailPrint "$(ERROR_NOT_ADMINISTRATOR)"
    GoTo lbl_Done
 
    lbl_Error_InvalidPlatform:
    DetailPrint "$(ERROR_INVALID_PLATFORM)"
    GoTo lbl_Done
 
    lbl_TimeOut:
    DetailPrint "$(DESC_DOTNET_TIMEOUT)"
	IfErrors 0 +2
				Call RollBackAndAbort
    GoTo lbl_Done
 
    lbl_Error:
    DetailPrint "$(ERROR_DOTNET_INVALID_PATH)"
    GoTo lbl_Done
 
    lbl_FatalError:
    DetailPrint "$(ERROR_DOTNET_FATAL)[$DOTNET_RETURN_CODE]"
	IfErrors 0 +2
				Call RollBackAndAbort
    GoTo lbl_Done
 
    lbl_Done:
    DetailPrint "$(FAILED_DOTNET_INSTALL)"
    lbl_NoError:
    lbl_IsSilent:
	${EndIf}		
	SectionEnd	

Section "$(SALMA_NAME) $(NOT_INSTALLED)" SEC01
	File SALMA.msi
	DetailPrint "$(INSTALLING_SALMA)"
	ClearErrors
	nsExec::ExecToStack "msiexec.exe /i Salma.msi /quiet /norestart" 
	ClearErrors
	IfErrors 0 +2
				Call RollBackAndAbort
	DetailPrint "$(INSTALLED_SALMA)"
	DetailPrint "$(REMOVING_TEMPORARIES)" 
SectionEnd

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC01} $(DESC_SEC01)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC02} $(DESC_SEC02)
  ;!insertmacro MUI_DESCRIPTION_TEXT ${SECGR01} $(DESC_SECGR01)
!insertmacro MUI_FUNCTION_DESCRIPTION_END 

########### UnInstaller Section ###########

Section Uninstall
${If} ${FileExists} "$INSTDIR\SALMA.msi"
	DetailPrint $(DEINSTALLATION_MESSAGE_BEFORE)
    nsExec::ExecToStack '"msiexec" /x "$INSTDIR\SALMA.msi" /quiet'
    Delete "$INSTDIR\SALMA.msi"
	Delete "$INSTDIR\uninstaller.exe"
	RMDir "$INSTDIR"
	RMDir /r "$PROGRAMFILES\SoftServe"
	DetailPrint $(DEINSTALLATION_MESSAGE_AFTER)
${Else}
	DetailPrint $(DEINSTALLATION_MESSAGE_BEFORE)
	Delete "$INSTDIR\en-US\*.*"
	RMDir "$INSTDIR\en-US"
	Delete "$INSTDIR\Help\*.*"
	RMDir "$INSTDIR\Help"
	Delete "$INSTDIR\ru-RU\*.*"
	RMDir "$INSTDIR\ru-RU"
	Delete $INSTDIR\*.*
	Delete "$INSTDIR\uninstaller.exe"
	RMDir "$INSTDIR"
	RMDir /r "$PROGRAMFILES\SoftServe"
	DeleteRegKey HKCU "SOFTWARE\Microsoft\Office\Word\AddIns\SoftServe.SALMA"
	DetailPrint $(DEINSTALLATION_MESSAGE_AFTER)
${EndIf}
SectionEnd

Function DeleteFiles
Delete "$INSTDIR\en-US\*.*"
RMDir "$INSTDIR\en-US"
Delete "$INSTDIR\Help\*.*"
RMDir "$INSTDIR\Help"
Delete "$INSTDIR\ru-RU\*.*"
RMDir "$INSTDIR\ru-RU"
Delete $INSTDIR\*.*
Delete "$INSTDIR\uninstaller.exe"
RMDir "$INSTDIR"
RMDir /r "$PROGRAMFILES\SoftServe"
DeleteRegKey HKCU "SOFTWARE\Microsoft\Office\Word\AddIns\SoftServe.SALMA"

FunctionEnd

############### Functions #########################

Function un.onInit
System::Call "kernel32::GetUserDefaultLangID()i.a"
FunctionEnd

Function .onInit 
	InitPluginsDir
	SetOutPath "$PLUGINSDIR"
	File /r "${NSISDIR}\Plugins\*.*"	
	System::Call "kernel32::GetUserDefaultLangID()i.a"
	Call CheckComponentSectionGroup		
	ReadRegStr $0 HKCU "SOFTWARE\Microsoft\Office\Word\AddIns\SoftServe.SALMA" "Description"
	${If} $0 == "SoftServe SALMA Add-In"
		StrCpy $SALMAStatus "RpairRemove"
	${EndIf}
	StrCpy $URL_VSTO "${URL_VSTO_COMMON}"		
FunctionEnd

Function CheckComponentSectionGroup
	Call CheckSalmaInstalled		
	Call CheckVSTORInstalled
	Call CheckSectionGroupIfHasComponents
FunctionEnd

Function CheckSalmaInstalled
ReadRegStr $0 HKCU "SOFTWARE\Microsoft\Office\Word\AddIns\SoftServe.SALMA" "Description"
	${If} $0 == "SoftServe SALMA Add-In"
		;SectionSetText ${SEC01} ""
		SectionSetFlags ${SEC01} 17
		!insertmacro UnselectSection ${SEC01}		
	${EndIf}
FunctionEnd

Function CheckVSTORInstalled	
	StrCpy $IfVstoRuntimeChecked 1	
	ReadRegStr $0 HKLM "SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4" "Install"
	${If} $0 == "1"
		Call DisableVSTOSection
	${Else}
		ReadRegStr $0 HKLM "SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4M" "VSTOMFeature_CLR35"
		${If} $0 == "1"
			Call DisableVSTOSection
		${Else}
			ReadRegStr $0 HKLM "SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R" "ProductCode"
			${If} $0 == "{CB0FD760-C6C6-3AF6-AD18-FE3B3B78727D}"
				Call DisableVSTOSection
			${Else}
				ReadRegStr $0 HKLM "SOFTWARE\Microsoft\VSTO Runtime Setup\v4" "Install"
				${If} $0 == "1"
					Call DisableVSTOSection
				${Else}
					ReadRegStr $0 HKLM "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" "ProductCode"
					${If} $0 == "{A0FE0292-D3BE-3447-80F2-72E032A54875}"
						Call DisableVSTOSection
					${EndIf}
				${EndIf}
			${EndIf}
		${EndIf}
	${EndIf}

FunctionEnd

Function DisableVSTOSection
	SectionSetFlags ${SEC02} 17
	SectionSetSize ${SEC02} 0
	StrCpy $IfVstoRuntimeChecked 0
	SectionSetText ${SEC02} "$(VSTOR_NAME) $(INSTALLED)"
FunctionEnd

Function CheckSectionGroupIfHasComponents
	${If} $IfVstoRuntimeChecked == 0
		;SectionSetText ${SECGR01} ""
	${EndIf}
FunctionEnd

######################## Custom Page #########################

Function ApplicatuionExistsDialog 

	Call CheckComponentSectionGroup		
  ${If} $SALMAStatus == "RpairRemove"
  ${OrIf} $SALMAStatus == "Update"
	WriteUninstaller $INSTDIR\uninstaller.exe
	File SALMA.msi
    Exec $INSTDIR\uninstaller.exe
    Quit
  ${EndIf}  
FunctionEnd 

Function ApplicatuionExistsDialogLeave
FunctionEnd
#################################################################

########################### Roll Back ###########################
Function RollBackAndAbort
 Call DeleteFiles
 DetailPrint "$(ABORTING_INSTALLATION)"
 Abort
FunctionEnd

Function .onSelChange
!insertmacro SelectSection ${SEC01}
FunctionEnd