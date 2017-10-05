/*
  Version 1.0.1.1 by Emin Onad
  CommonInitial Revision

  For more information please refer to the following Local MSDN Library pages:

  Platform SDK: SDK Tools - VERSIONINFO Resource
  ms-help://MS.MSDNQTR.2003FEB.1033/tools/tools/versioninfo_resource.htm

  And the documentation for the following functions:
  GetFileVersionInfoSize, GetFileVersionInfo, and VerQueryValue

  Thanks to Peter Windridge and Bernhard mayer to get me started...

  CommonInitially in Delphi 7.0.
  
  Version 1.0.1.2 - C version with NSIS & Unicode NSIS builds

*/

// For the DLL version info and icon is added, this way there is less chance the
// NSIS DLL is recognized a false positive with the crappy MS Antispyware.
// If you recompile, change the moreinfoversion.rc file and run make_res.bat first.

//#define __DEBUG


#define STRICT
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_DEPRECATE
#include <windows.h>
#include "nsis/pluginapi.h"


/*{-----}*/

void GetFileInfo(TCHAR const *Filename, TCHAR const *InfoProperty)
{
  void *Info;
  void *InfoData;
  DWORD InfoSize;
  DWORD Len;
  struct
  {
    WORD wLanguage;
    WORD wCodePage;
  } *LangPtr;
  void *Result = NULL;

  InfoSize = GetFileVersionInfoSize(Filename, &Len);
  if ( (InfoSize > 0) && (Info = GlobalAlloc(GMEM_FIXED, InfoSize)) )
  {
	if (GetFileVersionInfo(Filename, 0, InfoSize, Info))
	{
      if (VerQueryValue(Info, _T("\\VarFileInfo\\Translation"), &LangPtr, &Len))
	  {
	    TCHAR InfoType[256];  /* just needs to be long enough for fixed length + InfoProperty */
		wsprintf(InfoType, _T("\\StringFileInfo\\%04X%04X\\%s"), LangPtr->wLanguage, LangPtr->wCodePage, InfoProperty);
        if (VerQueryValue(Info, InfoType, &InfoData, &Len))
          Result=InfoData;
	  }
	}
  }

  /* copy results onto NSIS stack prior to deallocating memory it points to */
  if (Result != NULL)
    pushstring(Result);
  else
    pushstring(_T(""));

  /* free memory we allocated above */
  if (Info != NULL)
	GlobalFree(Info);
}


/*{-----}*/

void CommonInit(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, TCHAR const *InfoProperty)
{
  TCHAR *TempFileName;
  TCHAR *TempInfoProperty = NULL;

  EXDLL_INIT();
  if ((TempFileName = GlobalAlloc(GMEM_FIXED, string_size)) != NULL)
  {
    popstring(TempFileName);

	/* if InfoProperty not provided then pop from NSIS stack */
	if (InfoProperty == NULL)
	{
	  if ((TempInfoProperty = GlobalAlloc(GMEM_FIXED, string_size)) != NULL)
	  {
	    popstring(TempInfoProperty);
	    InfoProperty = TempInfoProperty;
	  }
	}

	/* retrieve the resource information */
    GetFileInfo(TempFileName, InfoProperty);

	/* free memory allocated */
	if (TempInfoProperty != NULL)
	  GlobalFree(TempInfoProperty);
    if (TempFileName != NULL)
      GlobalFree(TempFileName);
  }
  else
  {
    pushstring(_T(""));
  }
}


/*{-----}*/

void __declspec(dllexport) GetComments(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("Comments"));
}


/*{-----}*/

void __declspec(dllexport) GetPrivateBuild(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("PrivateBuild"));
}


/*{-----}*/

void __declspec(dllexport) GetSpecialBuild(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("SpecialBuild"));
}


/*{-----}*/

void __declspec(dllexport) GetProductName(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("ProductName"));
}


/*{-----}*/

void __declspec(dllexport) GetProductVersion(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("ProductVersion"));
}


/*{-----}*/

void __declspec(dllexport) GetCompanyName(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("CompanyName"));
}


/*{-----}*/

void __declspec(dllexport) GetFileVersion(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("FileVersion"));
}


/*{-----}*/

void __declspec(dllexport) GetFileDescription(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("FileDescription"));
}


/*{-----}*/

void __declspec(dllexport) GetInternalName(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("InternalName"));
}


/*{-----}*/

void __declspec(dllexport) GetLegalCopyright(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("LegalCopyright"));
}


/*{-----}*/

void __declspec(dllexport) GetLegalTrademarks(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("LegalTrademarks"));
}


/*{-----}*/

void __declspec(dllexport) GetOriginalFilename(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, _T("OriginalFilename"));
}


/*{-----}*/

void __declspec(dllexport) GetUserDefined(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  CommonInit(hwndParent, string_size, variables, stacktop, NULL);
}


/*{-----}*/

TCHAR *GetWindowSystemFolder()
{
  UINT Required = GetSystemDirectory(NULL, 0); /* determine size of buffer required */
  if (Required)
  {
    TCHAR *buffer = GlobalAlloc(GMEM_FIXED, Required + g_stringsize); /* allocate memory to store system directory path + some */
	if (buffer != NULL)
	{
	  *buffer = _T('\0');
	  GetSystemDirectory(buffer, Required);
	  return buffer;                                   /* caller must free memory */
	}
  }
  return NULL;
}


/*{-----}*/

// Get the real OS interface language, NOT the locale settings.
//
// IMHO:
// If one can read the OS language one for sure can use this 
// Language in a NSIS installer

void __declspec(dllexport) GetOSUserinterfaceLanguage(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
  TCHAR const * CheckFName = _T("USER.EXE"); //This OS file is available in Windows 95,NT4,98,98SE,ME,W2K,XP
  DWORD LangID, BSize, Dummy;
  TCHAR *FName;
  static TCHAR sLCID[16];
  void *InfoBuffer;
  struct
  {
    WORD wLanguage;
    WORD wCodePage;
  } *LangPtr;

  EXDLL_INIT();

  // The LCID, or "locale identifer," is a 32-bit data type into which are packed
  // several different values that help to identify a particular geographical
  // region. One of these internal values is the "primary language ID" which
  // identifies the basic language of the region or locale, such as English,
  // Spanish, or Turkish.

  LangID=0;//1033;  //English, to make sure that we have a default language

  FName = GetWindowSystemFolder();

  // FName does not end in \ unless SystemFolder is root, so add
  if (FName != NULL)
  {
    if (FName[lstrlen(FName)-1] != _T('\\'))
      lstrcat(FName, _T("\\"));

	lstrcat(FName, CheckFName); //e.g. in W2K this is in C:\WINNT\SYSTEM32\USER.EXE
    BSize = GetFileVersionInfoSize(FName, &Dummy);

    if ( BSize && ((InfoBuffer = GlobalAlloc(GMEM_FIXED, BSize)) != NULL) )
    {
      if ( GetFileVersionInfo(FName, 0, BSize, InfoBuffer) && 
		   VerQueryValue(InfoBuffer, _T("\\VarFileInfo\\Translation"), (void *)&LangPtr, &Dummy) 
		 )
	  {
        LangID=LangPtr->wLanguage;
	  }
	  GlobalFree(InfoBuffer);
    }
    GlobalFree(FName);
  }

  wsprintf(sLCID, _T("%d"), LangID); //  Str(LangID, sLCID);
  setuservariable(INST_LANG, sLCID);
  //pushstring(sLCID);
}


/*{-----}*/

// DLL entry point
BOOL WINAPI _DllMainCRTStartup(HANDLE _hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
  return TRUE;
}
