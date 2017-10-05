{
  Version 1.0.1.1 by Emin Onad
  Initial Revision

  For more information please refer to the following Local MSDN Library pages:

  Platform SDK: SDK Tools - VERSIONINFO Resource
  ms-help://MS.MSDNQTR.2003FEB.1033/tools/tools/versioninfo_resource.htm

  And the documentation for the following functions:
  GetFileVersionInfoSize, GetFileVersionInfo, and VerQueryValue

  Thanks to Peter Windridge and Bernhard mayer to get me started...

  Tested in Delphi 7.0, No not with other but will no doubt work in Delphi 6 and
  Delpi 2005 (v9) If you use older Delphi versions consider upgrading, it it
  worth it, really ;)

  Why Delphi ?
  Well I like to finish my project in time and without buffer overflows ;)
  and have readable code.

 Can the solution be improved?
   no doubt, feel free, after all you have the source...

 In return, please do something against software patent proposal in Europe.

}

{..DEFINE __DEBUG}

library MoreInfo;

uses
  windows;

type
  VarConstants = (
    INST_0,       // $0
    INST_1,       // $1
    INST_2,       // $2
    INST_3,       // $3
    INST_4,       // $4
    INST_5,       // $5
    INST_6,       // $6
    INST_7,       // $7
    INST_8,       // $8
    INST_9,       // $9
    INST_R0,      // $R0
    INST_R1,      // $R1
    INST_R2,      // $R2
    INST_R3,      // $R3
    INST_R4,      // $R4
    INST_R5,      // $R5
    INST_R6,      // $R6
    INST_R7,      // $R7
    INST_R8,      // $R8
    INST_R9,      // $R9
    INST_CMDLINE, // $CMDLINE
    INST_INSTDIR, // $INSTDIR
    INST_OUTDIR,  // $OUTDIR
    INST_EXEDIR,  // $EXEDIR
    INST_LANG,    // $LANGUAGE
    __INST_LAST
    );
  TVariableList = INST_0..__INST_LAST;
  pstack_t = ^stack_t;
  stack_t = record
    next: pstack_t;
    text: PChar;
  end;

var
  g_stringsize: integer;
  g_stacktop: ^pstack_t;
  g_variables: PChar;
  g_hwndParent: HWND;

// For the DLL version info and icon is added, this way there is less chance the
// NSIS DLL is recognized a false positive with the crappy MS Antispyware.
// If you recompile, change the moreinfoversion.rc file and run make_res.bat first.

{$R moreinfoversion.res}
{$R moreinfoextra.res}

{-----}

procedure Init(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer);
begin
  g_stringsize := string_size;
  g_hwndParent := hwndParent;
  g_stacktop   := stacktop;
  g_variables  := variables;
end;

{-----}

function PopString(): string;
var
  th: pstack_t;
begin
  if integer(g_stacktop^) <> 0 then
  begin
    th := g_stacktop^;
    Result := PChar(@th.text);
    g_stacktop^ := th.next;
    GlobalFree(HGLOBAL(th));
  end;
end;

{-----}

procedure PushString(const str: string='');
var
  th: pstack_t;
begin
  if integer(g_stacktop) <> 0 then
  begin
    th := pstack_t(GlobalAlloc(GPTR, SizeOf(stack_t) + g_stringsize));
    lstrcpyn(@th.text, PChar(str), g_stringsize);
    th.next := g_stacktop^;
    g_stacktop^ := th;
  end;
end;

{-----}

function GetUserVariable(const varnum: TVariableList): string;
begin
  if (integer(varnum) >= 0) and (integer(varnum) < integer(__INST_LAST)) then
  begin
    Result := g_variables + integer(varnum) * g_stringsize;
  end
  else
  begin
    Result := '';
  end;  
end;

{-----}

procedure SetUserVariable(const varnum: TVariableList; const value: string);
begin
  if (value <> '') and (integer(varnum) >= 0) and (integer(varnum) < integer(__INST_LAST)) then
  begin
    lstrcpy(g_variables + integer(varnum) * g_stringsize, PChar(value));
  end;  
end;

{-----}

function IntToHex(Value: Int64; Digits: Word): String;
const
  HexChars: array [0..15] of Char = '0123456789ABCDEF';
  DestLen: Integer = SizeOf(Value) * 2;
var
  OutBuf: array [0..SizeOf(Value) * 2] of Char;
  i: Integer;
  PValue: Pointer;
  POut: PChar;
  B: Byte;
begin
  PValue := Pointer(Integer(@Value) + SizeOf(Value));
  POut := OutBuf;
  try
    for i := 1 to SizeOf(Value) do
    begin
      Dec(PChar(PValue));
      B := Byte(PValue^);
      POut^ := HexChars[B shr 4];
      Inc(POut);
      POut^ := HexChars[B and $0F];
      Inc(POut);
    end;
    POut^ := Char(0);

    if Digits = 0 then
    begin
      Inc(Digits, 2);
    end
    else
    if Digits > DestLen then
       Digits := DestLen;
       
    POut := OutBuf;
    for i := 1 to DestLen do
    begin
      if ((DestLen - i) < Digits) or (POut^ <> '0') then break;
      Inc(POut);
    end;

    Result := String(POut);
  except
    Result := IntToHex(Value, Integer(Digits));
  end;
end;

{-----}

function GetFileInfo(Filename, InfoProperty: PChar):Pointer;
var
  Info: Pointer;
  InfoData: Pointer;
  InfoSize: LongInt;
  Len: DWORD;
  Infotype: string;
  LangPtr: Pointer;
begin
  Result:=nil;
  Len := MAX_PATH + 1;
  InfoType := InfoProperty;

  InfoSize := GetFileVersionInfoSize(Filename, Len);  //TODO: Filename shortpath on 95, 98, 98SE, ME  ?
  if (InfoSize > 0) then
  begin
    GetMem(Info, InfoSize);
    try
      if GetFileVersionInfo(Filename, Len, InfoSize, Info) then
      begin
        Len := 255;
        if VerQueryValue(Info, '\VarFileInfo\Translation', LangPtr, Len) then
        begin
          InfoType := '\StringFileInfo\'+
                      IntToHex(LoWord(LongInt(LangPtr^)),4)+
                      IntToHex(HiWord(LongInt(LangPtr^)),4)+
                      '\'+InfoType+#0;
        end;
        
        if VerQueryValue(Info, Pchar(InfoType), InfoData, len) then
        begin
          Result:=InfoData;
        end;  
      end;
    finally
      FreeMem(Info, InfoSize);
    end;
  end;
end;

{-----}

procedure GetComments(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'Comments')));
end;

{-----}

procedure GetPrivateBuild(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'PrivateBuild')));
end;

{-----}

procedure GetSpecialBuild(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'SpecialBuild')));
end;

{-----}

procedure GetProductName(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'ProductName')));
end;

{-----}

procedure GetProductVersion(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'ProductVersion')));
end;

{-----}

procedure GetCompanyName(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'CompanyName')));
end;

{-----}

procedure GetFileVersion(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'FileVersion')));
end;

{-----}

procedure GetFileDescription(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'FileDescription')));
end;

{-----}

procedure GetInternalName(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'InternalName')));
end;

{-----}

procedure GetLegalCopyright(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'LegalCopyright')));
end;

{-----}

procedure GetLegalTrademarks(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'LegalTrademarks')));
end;

{-----}

procedure GetOriginalFilename(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  PushString(PChar(GetFileInfo(PChar(PopString), 'OriginalFilename')));
end;

{-----}

procedure GetUserDefined(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
var
  TempFileName, TempInfoProperty: String;
Begin
  Init(hwndParent, string_size, variables, stacktop);
  TempFileName := PopString;
  TempInfoProperty := PopString;
  PushString(PChar(GetFileInfo(PChar(TempFileName), PChar(TempInfoProperty))));
end;

{-----}
function GetWindowsSystemFolder: string;
var
  Required: Cardinal;
begin
  Result   := '';
  Required := GetSystemDirectory(nil, 0);
  if Required <> 0 then
  begin
    SetLength(Result, Required);
    GetSystemDirectory(PChar(Result), Required);
    SetLength(Result, Length(PChar(Result)));  //TODO: StrLen ?
  end;
end;

{-----}

// Get the real OS interface language, NOT the locale settings.
//
// IMHO:
// If one can read the OS language one for sure can use this 
// Language in a NSIS installer

procedure GetOSUserinterfaceLanguage(const hwndParent: HWND; const string_size: integer; const variables: PChar; const stacktop: pointer); cdecl;
type
  TBuf = array [1..4] of smallint;
  PBuf = ^TBuf;
const
  CheckFName = 'USER.EXE'; //This OS file is available in Windows 95,NT4,98,98SE,ME,W2K,XP
var
  DefaultLangID, LangID, BSize, DummySize, Handle: LongWord; // For Delphi 3, change Longword to DWORD
  FName, sLCID: string;
  //FFolder : string;
  Buffer:     PBuf;
  InfoBuffer: Pointer;
begin
  Init(hwndParent, string_size, variables, stacktop);

  // The LCID, or "locale identifer," is a 32-bit data type into which are packed
  // several different values that help to identify a particular geographical
  // region. One of these internal values is the "primary language ID" which
  // identifies the basic language of the region or locale, such as English,
  // Spanish, or Turkish.

  DefaultLangID:=0;//1033;  //English, to make sure that we have a default language

  FName := GetWindowsSystemFolder;

  //FFolder := '\TEMP';  //for testing only: Store USER.EXE of different windows versions  in a temp folder below root of drive

  if FName[Length(FName)] <> '\' then
  begin
    FName := FName + '\';
  end;

  FName := FName + CheckFName; //e.g. in W2K this is in C:\WINNT\SYSTEM32\USER.EXE
  BSize := GetFileVersionInfoSize(PChar(FName), Handle);

  if BSize = 0 then
  Begin
    LangID:=DefaultLangID; // Unknown. Maybe the file was not there...  TODO?: use OS Locale value
  end
  else
  begin
    GetMem(InfoBuffer, BSize);
    try
      if GetFileVersionInfo(PChar(FName), Handle, BSize, InfoBuffer) then
      begin
        VerQueryValue(InfoBuffer, PChar('\VarFileInfo\Translation'), Pointer(Buffer), DummySize);
        LangID:=Buffer^[1];
      end
      else
        LangID:=DefaultLangID; // Unknown. Could not determine language information... Improve? then take locale value and use this value
    finally
      FreeMem(InfoBuffer);
    end;
  end;

  {$IFNDEF __DEBUG}
  Str(LangID, sLCID);
  SetUserVariable(INST_LANG, sLCID);
  {$ELSE}
  SetUserVariable(INST_LANG, '1055');  // Turkish for debugging purpose
  {$ENDIF __DEBUG}

end;

{-----}

exports GetComments;
exports GetPrivateBuild;
exports GetSpecialBuild;
exports GetProductName;
exports GetProductVersion;
exports GetCompanyName;
exports GetFileVersion;
exports GetFileDescription;
exports GetInternalName;
exports GetLegalCopyright;
exports GetLegalTrademarks;
exports GetOriginalFilename;
exports GetUserDefined;
exports GetOSUserinterfaceLanguage;

begin
end.


