Global Const $GUI_EVENT_CLOSE = + 4294967293
Global Const $GUI_CHECKED = 1
Global Const $WS_BORDER = 8388608
Global Const $WS_EX_DLGMODALFRAME = 1
Global Const $STR_NOCASESENSEBASIC = 2
Global Const $STR_STRIPLEADING = 1
Global Const $STR_STRIPTRAILING = 2
Global Const $STR_ENTIRESPLIT = 1
Global Const $STR_NOCOUNT = 2
Global Const $SB_UTF8 = 4
Global Const $FORMAT_MESSAGE_IGNORE_INSERTS = 512
Global Const $FORMAT_MESSAGE_FROM_SYSTEM = 4096
Func _WINAPI_FORMATMESSAGE ( $IFLAGS , $PSOURCE , $IMESSAGEID , $ILANGUAGEID , ByRef $PBUFFER , $ISIZE , $VARGUMENTS )
	Local $SBUFFERTYPE = "struct*"
	If IsString ( $PBUFFER ) Then $SBUFFERTYPE = "wstr"
	Local $ACALL = DllCall ( "kernel32.dll" , "dword" , "FormatMessageW" , "dword" , $IFLAGS , "struct*" , $PSOURCE , "dword" , $IMESSAGEID , "dword" , $ILANGUAGEID , $SBUFFERTYPE , $PBUFFER , "dword" , $ISIZE , "ptr" , $VARGUMENTS )
	If @error Then Return SetError ( @error , @extended , 0 )
	If Not $ACALL [ 0 ] Then Return SetError ( 10 , _WINAPI_GETLASTERROR ( ) , 0 )
	If $SBUFFERTYPE = "wstr" Then $PBUFFER = $ACALL [ 5 ]
	Return $ACALL [ 0 ]
EndFunc
Func _WINAPI_GETERRORMESSAGE ( $ICODE , $ILANGUAGE = 0 , Const $_ICALLERERROR = @error , Const $_ICALLEREXTENDED = @extended )
	Local $ACALL = DllCall ( "kernel32.dll" , "dword" , "FormatMessageW" , "dword" , BitOR ( $FORMAT_MESSAGE_FROM_SYSTEM , $FORMAT_MESSAGE_IGNORE_INSERTS ) , "ptr" , 0 , "dword" , $ICODE , "dword" , $ILANGUAGE , "wstr" , "" , "dword" , 4096 , "ptr" , 0 )
	If @error Or Not $ACALL [ 0 ] Then Return SetError ( @error , @extended , "" )
	Return SetError ( $_ICALLERERROR , $_ICALLEREXTENDED , StringRegExpReplace ( $ACALL [ 5 ] , "[" & @LF & "," & @CR & "]*\Z" , "" ) )
EndFunc
Func _WINAPI_GETLASTERROR ( Const $_ICALLERERROR = @error , Const $_ICALLEREXTENDED = @extended )
	Local $ACALL = DllCall ( "kernel32.dll" , "dword" , "GetLastError" )
	Return SetError ( $_ICALLERERROR , $_ICALLEREXTENDED , $ACALL [ 0 ] )
EndFunc
Func _SECURITY__GETLENGTHSID ( $PSID )
	If Not _SECURITY__ISVALIDSID ( $PSID ) Then Return SetError ( @error + 10 , @extended , 0 )
	Local $ACALL = DllCall ( "advapi32.dll" , "dword" , "GetLengthSid" , "struct*" , $PSID )
	If @error Then Return SetError ( @error , @extended , 0 )
	Return $ACALL [ 0 ]
EndFunc
Func _SECURITY__ISVALIDSID ( $PSID )
	Local $ACALL = DllCall ( "advapi32.dll" , "bool" , "IsValidSid" , "struct*" , $PSID )
	If @error Then Return SetError ( @error , @extended , False )
	Return Not ( $ACALL [ 0 ] = 0 )
EndFunc
Func _SECURITY__LOOKUPACCOUNTSID ( $VSID , $SSYSTEM = "" )
	Local $PSID , $AACCT [ 3 ]
	If IsString ( $VSID ) Then
		$PSID = _SECURITY__STRINGSIDTOSID ( $VSID )
	Else
		$PSID = $VSID
	EndIf
	If Not _SECURITY__ISVALIDSID ( $PSID ) Then Return SetError ( @error + 20 , @extended , 0 )
	If $SSYSTEM = "" Then $SSYSTEM = Null
	Local $ACALL = DllCall ( "advapi32.dll" , "bool" , "LookupAccountSidW" , "wstr" , $SSYSTEM , "struct*" , $PSID , "wstr" , "" , "dword*" , 65536 , "wstr" , "" , "dword*" , 65536 , "int*" , 0 )
	If @error Or Not $ACALL [ 0 ] Then Return SetError ( @error + 10 , @extended , 0 )
	Local $AACCT [ 3 ]
	$AACCT [ 0 ] = $ACALL [ 3 ]
	$AACCT [ 1 ] = $ACALL [ 5 ]
	$AACCT [ 2 ] = $ACALL [ 7 ]
	Return $AACCT
EndFunc
Func _SECURITY__STRINGSIDTOSID ( $SSID )
	Local $ACALL = DllCall ( "advapi32.dll" , "bool" , "ConvertStringSidToSidW" , "wstr" , $SSID , "ptr*" , 0 )
	If @error Or Not $ACALL [ 0 ] Then Return SetError ( @error + 10 , @extended , 0 )
	Local $PSID = $ACALL [ 2 ]
	Local $TBUFFER = DllStructCreate ( "byte Data[" & _SECURITY__GETLENGTHSID ( $PSID ) & "]" , $PSID )
	Local $TSID = DllStructCreate ( "byte Data[" & DllStructGetSize ( $TBUFFER ) & "]" )
	DllStructSetData ( $TSID , "Data" , DllStructGetData ( $TBUFFER , "Data" ) )
	DllCall ( "kernel32.dll" , "handle" , "LocalFree" , "handle" , $PSID )
	Return $TSID
EndFunc
Global Const $TAGFILETIME = "struct;dword Lo;dword Hi;endstruct"
Global Const $TAGSYSTEMTIME = "struct;word Year;word Month;word Dow;word Day;word Hour;word Minute;word Second;word MSeconds;endstruct"
Global Const $TAGEVENTLOGRECORD = "dword Length;dword Reserved;dword RecordNumber;dword TimeGenerated;dword TimeWritten;dword EventID;" & "word EventType;word NumStrings;word EventCategory;word ReservedFlags;dword ClosingRecordNumber;dword StringOffset;" & "dword UserSidLength;dword UserSidOffset;dword DataLength;dword DataOffset"
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $NUMBER_DOUBLE = 3
Global Const $FO_OVERWRITE = 2
Global Const $FLTA_FILESFOLDERS = 0
Global Const $FLTAR_FILESFOLDERS = 0
Global Const $FLTAR_NOHIDDEN = 4
Global Const $FLTAR_NOSYSTEM = 8
Global Const $FLTAR_NOLINK = 16
Global Const $FLTAR_NORECUR = 0
Global Const $FLTAR_NOSORT = 0
Global Const $FLTAR_RELPATH = 1
Global Const $TAGOSVERSIONINFO = "struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct"
Func _WINAPI_FREELIBRARY ( $HMODULE )
	Local $ACALL = DllCall ( "kernel32.dll" , "bool" , "FreeLibrary" , "handle" , $HMODULE )
	If @error Then Return SetError ( @error , @extended , False )
	Return $ACALL [ 0 ]
EndFunc
Func _WINAPI_GETVERSION ( )
	Local $TOSVI = DllStructCreate ( $TAGOSVERSIONINFO )
	DllStructSetData ( $TOSVI , 1 , DllStructGetSize ( $TOSVI ) )
	Local $ACALL = DllCall ( "kernel32.dll" , "bool" , "GetVersionExW" , "struct*" , $TOSVI )
	If @error Or Not $ACALL [ 0 ] Then Return SetError ( @error , @extended , 0 )
	Return Number ( DllStructGetData ( $TOSVI , 2 ) & "." & DllStructGetData ( $TOSVI , 3 ) , $NUMBER_DOUBLE )
EndFunc
Func _DATE_TIME_FILETIMETOLOCALFILETIME ( $TFILETIME )
	Local $TLOCAL = DllStructCreate ( $TAGFILETIME )
	Local $ACALL = DllCall ( "kernel32.dll" , "bool" , "FileTimeToLocalFileTime" , "struct*" , $TFILETIME , "struct*" , $TLOCAL )
	If @error Then Return SetError ( @error , @extended , 0 )
	Return SetExtended ( $ACALL [ 0 ] , $TLOCAL )
EndFunc
Func _DATE_TIME_FILETIMETOSYSTEMTIME ( $TFILETIME )
	Local $TSYSTTIME = DllStructCreate ( $TAGSYSTEMTIME )
	Local $ACALL = DllCall ( "kernel32.dll" , "bool" , "FileTimeToSystemTime" , "struct*" , $TFILETIME , "struct*" , $TSYSTTIME )
	If @error Then Return SetError ( @error , @extended , 0 )
	Return SetExtended ( $ACALL [ 0 ] , $TSYSTTIME )
EndFunc
Global Const $HGDI_ERROR = Ptr ( + 4294967295 )
Global Const $INVALID_HANDLE_VALUE = Ptr ( + 4294967295 )
Global Const $KF_EXTENDED = 256
Global Const $KF_ALTDOWN = 8192
Global Const $KF_UP = 32768
Global Const $LLKHF_EXTENDED = BitShift ( $KF_EXTENDED , 8 )
Global Const $LLKHF_ALTDOWN = BitShift ( $KF_ALTDOWN , 8 )
Global Const $LLKHF_UP = BitShift ( $KF_UP , 8 )
Func _WINAPI_LOADINDIRECTSTRING ( $SSTRIN )
	Local $ACALL = DllCall ( "shlwapi.dll" , "uint" , "SHLoadIndirectString" , "wstr" , $SSTRIN , "wstr" , "" , "uint" , 4096 , "ptr*" , 0 )
	If @error Then Return SetError ( @error , @extended , "" )
	If $ACALL [ 0 ] Then Return SetError ( 10 , $ACALL [ 0 ] , "" )
	Return $ACALL [ 2 ]
EndFunc
Func _WINAPI_LOADLIBRARYEX ( $SFILENAME , $IFLAGS = 0 )
	Local $ACALL = DllCall ( "kernel32.dll" , "handle" , "LoadLibraryExW" , "wstr" , $SFILENAME , "ptr" , 0 , "dword" , $IFLAGS )
	If @error Then Return SetError ( @error , @extended , 0 )
	Return $ACALL [ 0 ]
EndFunc
Func _WINAPI_EXPANDENVIRONMENTSTRINGS ( $SSTRING )
	Local $ACALL = DllCall ( "kernel32.dll" , "dword" , "ExpandEnvironmentStringsW" , "wstr" , $SSTRING , "wstr" , "" , "dword" , 4096 )
	If @error Or Not $ACALL [ 0 ] Then Return SetError ( @error + 10 , @extended , "" )
	Return $ACALL [ 2 ]
EndFunc
Global $__G_SSOURCENAME_EVENT
Global Const $EVENTLOG_SUCCESS = 0
Global Const $EVENTLOG_ERROR_TYPE = 1
Global Const $EVENTLOG_WARNING_TYPE = 2
Global Const $EVENTLOG_INFORMATION_TYPE = 4
Global Const $EVENTLOG_AUDIT_SUCCESS = 8
Global Const $EVENTLOG_AUDIT_FAILURE = 16
Global Const $EVENTLOG_SEQUENTIAL_READ = 1
Global Const $EVENTLOG_SEEK_READ = 2
Global Const $EVENTLOG_FORWARDS_READ = 4
Global Const $EVENTLOG_BACKWARDS_READ = 8
Global Const $__EVENTLOG_LOAD_LIBRARY_AS_DATAFILE = 2
Global Const $__EVENTLOG_FORMAT_MESSAGE_FROM_HMODULE = 2048
Global Const $__EVENTLOG_FORMAT_MESSAGE_IGNORE_INSERTS = 512
Func _EVENTLOG__CLOSE ( $HEVENTLOG )
	Local $ARESULT = DllCall ( "advapi32.dll" , "bool" , "CloseEventLog" , "handle" , $HEVENTLOG )
	If @error Then Return SetError ( @error , @extended , False )
	Return $ARESULT [ 0 ] <> 0
EndFunc
Func __EVENTLOG_DECODECATEGORY ( $TEVENTLOG )
	Return DllStructGetData ( $TEVENTLOG , "EventCategory" )
EndFunc
Func __EVENTLOG_DECODECOMPUTER ( $TEVENTLOG )
	Local $PEVENTLOG = DllStructGetPtr ( $TEVENTLOG )
	Local $ILENGTH = DllStructGetData ( $TEVENTLOG , "UserSidOffset" ) + 4294967295
	Local $IOFFSET = DllStructGetSize ( $TEVENTLOG )
	$IOFFSET += 2 * ( StringLen ( __EVENTLOG_DECODESOURCE ( $TEVENTLOG ) ) + 1 )
	$ILENGTH -= $IOFFSET
	Local $TBUFFER = DllStructCreate ( "wchar Text[" & $ILENGTH & "]" , $PEVENTLOG + $IOFFSET )
	Return DllStructGetData ( $TBUFFER , "Text" )
EndFunc
Func __EVENTLOG_DECODEDATA ( $TEVENTLOG )
	Local $PEVENTLOG = DllStructGetPtr ( $TEVENTLOG )
	Local $IOFFSET = DllStructGetData ( $TEVENTLOG , "DataOffset" )
	Local $ILENGTH = DllStructGetData ( $TEVENTLOG , "DataLength" )
	Local $TBUFFER = DllStructCreate ( "byte[" & $ILENGTH & "]" , $PEVENTLOG + $IOFFSET )
	Local $ADATA [ $ILENGTH + 1 ]
	$ADATA [ 0 ] = $ILENGTH
	For $II = 1 To $ILENGTH
		$ADATA [ $II ] = DllStructGetData ( $TBUFFER , 1 , $II )
	Next
	Return $ADATA
EndFunc
Func __EVENTLOG_DECODEDATE ( $IEVENTTIME )
	Local $TINT64 = DllStructCreate ( "int64" )
	Local $PINT64 = DllStructGetPtr ( $TINT64 )
	Local $TFILETIME = DllStructCreate ( $TAGFILETIME , $PINT64 )
	DllStructSetData ( $TINT64 , 1 , ( $IEVENTTIME * 10000000 ) + 116444736000000000 )
	Local $TLOCALTIME = _DATE_TIME_FILETIMETOLOCALFILETIME ( $TFILETIME )
	Local $TSYSTTIME = _DATE_TIME_FILETIMETOSYSTEMTIME ( $TLOCALTIME )
	Local $IMONTH = DllStructGetData ( $TSYSTTIME , "Month" )
	Local $IDAY = DllStructGetData ( $TSYSTTIME , "Day" )
	Local $IYEAR = DllStructGetData ( $TSYSTTIME , "Year" )
	Return StringFormat ( "%02d/%02d/%04d" , $IMONTH , $IDAY , $IYEAR )
EndFunc
Func __EVENTLOG_DECODEDESC ( $TEVENTLOG )
	$ASTRINGS = __EVENTLOG_DECODESTRINGS ( $TEVENTLOG )
	$SSOURCE = __EVENTLOG_DECODESOURCE ( $TEVENTLOG )
	$IEVENTID = DllStructGetData ( $TEVENTLOG , "EventID" )
	$SKEY = "HKLM\SYSTEM\CurrentControlSet\Services\Eventlog\" & $__G_SSOURCENAME_EVENT & "\" & $SSOURCE
	$PROVIDERGUID = RegRead ( $SKEY , "providerGuid" )
	$MESSAGEDLL = RegRead ( $SKEY , "EventMessageFile" )
	$PARAMETERSDLL = RegRead ( $SKEY , "ParameterMessageFile" )
	If $MESSAGEDLL = "" Then
		$MESSAGEDLL = RegRead ( "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Publishers\" & $PROVIDERGUID , "MessageFileName" )
		If $MESSAGEDLL = "" Then
			$MESSAGEDLL = RegRead ( "HKEY_CLASSES_ROOT\CLSID\" & $PROVIDERGUID , "MessageFileName" )
		EndIf
	EndIf
	$AMSGDLL = StringSplit ( _WINAPI_EXPANDENVIRONMENTSTRINGS ( $MESSAGEDLL ) , ";" )
	If $PARAMETERSDLL = "" Then
		$PARAMETERSDLL = RegRead ( "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Publishers\" & $PROVIDERGUID , "ParameterFileName" )
		If $PARAMETERSDLL = "" Then
			$PARAMETERSDLL = RegRead ( "HKEY_CLASSES_ROOT\CLSID\" & $PROVIDERGUID , "ParameterFileName" )
		EndIf
	EndIf
	$APARAMDLL = StringSplit ( _WINAPI_EXPANDENVIRONMENTSTRINGS ( $PARAMETERSDLL ) , ";" )
	$ERRORNUMBER = 0
	$STRTEMP = ""
	$IFLAGS = BitOR ( $__EVENTLOG_FORMAT_MESSAGE_FROM_HMODULE , $__EVENTLOG_FORMAT_MESSAGE_IGNORE_INSERTS )
	$SDESC = ""
	If $MESSAGEDLL = "" Then
		Return "Event-ID " & String ( BitAND ( DllStructGetData ( $TEVENTLOG , "EventID" ) , 32767 ) )
	Else
		For $II = 1 To $AMSGDLL [ 0 ]
			$HDLL = _WINAPI_LOADLIBRARYEX ( $AMSGDLL [ $II ] , $__EVENTLOG_LOAD_LIBRARY_AS_DATAFILE )
			If $HDLL = 0 Then ContinueLoop
			$TBUFFER = DllStructCreate ( "wchar Text[4096]" )
			_WINAPI_FORMATMESSAGE ( $IFLAGS , $HDLL , $IEVENTID , 0 , $TBUFFER , 4096 , 0 )
			_WINAPI_FREELIBRARY ( $HDLL )
			$SDESC &= DllStructGetData ( $TBUFFER , "Text" )
		Next
		If $SDESC = "" Then
			For $II = 1 To $ASTRINGS [ 0 ]
				$SDESC &= $ASTRINGS [ $II ]
			Next
		Else
			For $II = $ASTRINGS [ 0 ] To 1 Step + 4294967295
				If StringLeft ( $ASTRINGS [ $II ] , 2 ) = "%%" Then
					$STRTEMP = StringMid ( $ASTRINGS [ $II ] , 3 )
					$ERRORNUMBER = Number ( $STRTEMP )
					For $IIF = 1 To $APARAMDLL [ 0 ]
						$HDLLP = _WINAPI_LOADLIBRARYEX ( $APARAMDLL [ $IIF ] , $__EVENTLOG_LOAD_LIBRARY_AS_DATAFILE )
						If $HDLLP = 0 Then ContinueLoop
						$TFBUFFER = DllStructCreate ( "wchar Text[4096]" )
						If _WINAPI_FORMATMESSAGE ( $IFLAGS , $HDLLP , $ERRORNUMBER , 0 , $TFBUFFER , 4096 , 0 ) Then
							$ASTRINGS [ $II ] = DllStructGetData ( $TFBUFFER , "Text" )
							$IIF = $APARAMDLL [ 0 ]
						EndIf
						_WINAPI_FREELIBRARY ( $HDLL )
						$TFBUFFER = 0
					Next
					$ASTRINGS [ $II ] = StringReplace ( $ASTRINGS [ $II ] , "%%" , "_____" )
				EndIf
				$SDESC = StringReplace ( $SDESC , "%" & $II , $ASTRINGS [ $II ] )
			Next
			$SDESC = StringReplace ( $SDESC , "_____" , "%%" )
		EndIf
		Return StringStripWS ( $SDESC , $STR_STRIPLEADING + $STR_STRIPTRAILING )
	EndIf
EndFunc
Func __EVENTLOG_DECODEEVENTID ( $TEVENTLOG )
	Return BitAND ( DllStructGetData ( $TEVENTLOG , "EventID" ) , 32767 )
EndFunc
Func __EVENTLOG_DECODESOURCE ( $TEVENTLOG )
	Local $PEVENTLOG = DllStructGetPtr ( $TEVENTLOG )
	Local $ILENGTH = DllStructGetData ( $TEVENTLOG , "UserSidOffset" ) + 4294967295
	Local $IOFFSET = DllStructGetSize ( $TEVENTLOG )
	$ILENGTH -= $IOFFSET
	Local $TBUFFER = DllStructCreate ( "wchar Text[" & $ILENGTH & "]" , $PEVENTLOG + $IOFFSET )
	Return DllStructGetData ( $TBUFFER , "Text" )
EndFunc
Func __EVENTLOG_DECODESTRINGS ( $TEVENTLOG )
	Local $PEVENTLOG = DllStructGetPtr ( $TEVENTLOG )
	Local $INUMSTRS = DllStructGetData ( $TEVENTLOG , "NumStrings" )
	Local $IOFFSET = DllStructGetData ( $TEVENTLOG , "StringOffset" )
	Local $IDATAOFFSET = DllStructGetData ( $TEVENTLOG , "DataOffset" )
	Local $TBUFFER = DllStructCreate ( "wchar Text[" & $IDATAOFFSET - $IOFFSET & "]" , $PEVENTLOG + $IOFFSET )
	Local $ASTRINGS [ $INUMSTRS + 1 ]
	$ASTRINGS [ 0 ] = $INUMSTRS
	For $II = 1 To $INUMSTRS
		$ASTRINGS [ $II ] = DllStructGetData ( $TBUFFER , "Text" )
		$IOFFSET += 2 * ( StringLen ( $ASTRINGS [ $II ] ) + 1 )
		$TBUFFER = DllStructCreate ( "wchar Text[" & $IDATAOFFSET - $IOFFSET & "]" , $PEVENTLOG + $IOFFSET )
	Next
	Return $ASTRINGS
EndFunc
Func __EVENTLOG_DECODETIME ( $IEVENTTIME )
	Local $TINT64 = DllStructCreate ( "int64" )
	Local $PINT64 = DllStructGetPtr ( $TINT64 )
	Local $TFILETIME = DllStructCreate ( $TAGFILETIME , $PINT64 )
	DllStructSetData ( $TINT64 , 1 , ( $IEVENTTIME * 10000000 ) + 116444736000000000 )
	Local $TLOCALTIME = _DATE_TIME_FILETIMETOLOCALFILETIME ( $TFILETIME )
	Local $TSYSTTIME = _DATE_TIME_FILETIMETOSYSTEMTIME ( $TLOCALTIME )
	Local $IHOURS = DllStructGetData ( $TSYSTTIME , "Hour" )
	Local $IMINUTES = DllStructGetData ( $TSYSTTIME , "Minute" )
	Local $ISECONDS = DllStructGetData ( $TSYSTTIME , "Second" )
	Local $SAMPM = "AM"
	If $IHOURS < 12 Then
		If $IHOURS = 0 Then
			$IHOURS = 12
		EndIf
	Else
		$SAMPM = "PM"
		If $IHOURS > 12 Then
			$IHOURS -= 12
		EndIf
	EndIf
	Return StringFormat ( "%02d:%02d:%02d %s" , $IHOURS , $IMINUTES , $ISECONDS , $SAMPM )
EndFunc
Func __EVENTLOG_DECODETYPESTR ( $IEVENTTYPE )
	Select
	Case $IEVENTTYPE = $EVENTLOG_SUCCESS
		Return "Success"
	Case $IEVENTTYPE = $EVENTLOG_ERROR_TYPE
		Return "Error"
	Case $IEVENTTYPE = $EVENTLOG_WARNING_TYPE
		Return "Warning"
	Case $IEVENTTYPE = $EVENTLOG_INFORMATION_TYPE
		Return "Information"
	Case $IEVENTTYPE = $EVENTLOG_AUDIT_SUCCESS
		Return "Success audit"
	Case $IEVENTTYPE = $EVENTLOG_AUDIT_FAILURE
		Return "Failure audit"
Case Else
		Return $IEVENTTYPE
	EndSelect
EndFunc
Func __EVENTLOG_DECODEUSERNAME ( $TEVENTLOG )
	Local $PEVENTLOG = DllStructGetPtr ( $TEVENTLOG )
	If DllStructGetData ( $TEVENTLOG , "UserSidLength" ) = 0 Then Return ""
	Local $PACCTSID = $PEVENTLOG + DllStructGetData ( $TEVENTLOG , "UserSidOffset" )
	Local $AACCTINFO = _SECURITY__LOOKUPACCOUNTSID ( $PACCTSID )
	If IsArray ( $AACCTINFO ) Then Return $AACCTINFO [ 1 ]
	Return ""
EndFunc
Func _EVENTLOG__OPEN ( $SSERVERNAME , $SSOURCENAME )
	$__G_SSOURCENAME_EVENT = $SSOURCENAME
	Local $ARESULT = DllCall ( "advapi32.dll" , "handle" , "OpenEventLogW" , "wstr" , $SSERVERNAME , "wstr" , $SSOURCENAME )
	If @error Then Return SetError ( @error , @extended , 0 )
	Return $ARESULT [ 0 ]
EndFunc
Func _EVENTLOG__READ ( $HEVENTLOG , $BREAD = True , $BFORWARD = True , $IOFFSET = 0 )
	Local $IREADFLAGS , $AEVENT [ 15 ]
	$AEVENT [ 0 ] = False
	If $BREAD Then
		$IREADFLAGS = $EVENTLOG_SEQUENTIAL_READ
	Else
		$IREADFLAGS = $EVENTLOG_SEEK_READ
	EndIf
	If $BFORWARD Then
		$IREADFLAGS = BitOR ( $IREADFLAGS , $EVENTLOG_FORWARDS_READ )
	Else
		$IREADFLAGS = BitOR ( $IREADFLAGS , $EVENTLOG_BACKWARDS_READ )
	EndIf
	Local $TBUFFER = DllStructCreate ( "wchar[1]" )
	Local $ARESULT = DllCall ( "advapi32.dll" , "bool" , "ReadEventLogW" , "handle" , $HEVENTLOG , "dword" , $IREADFLAGS , "dword" , $IOFFSET , "struct*" , $TBUFFER , "dword" , 0 , "dword*" , 0 , "dword*" , 0 )
	If @error Then Return SetError ( @error , @extended , $AEVENT )
	Local $IBYTESMIN = $ARESULT [ 7 ]
	$TBUFFER = DllStructCreate ( "wchar[" & $IBYTESMIN + 1 & "]" )
	$ARESULT = DllCall ( "advapi32.dll" , "bool" , "ReadEventLogW" , "handle" , $HEVENTLOG , "dword" , $IREADFLAGS , "dword" , $IOFFSET , "struct*" , $TBUFFER , "dword" , $IBYTESMIN , "dword*" , 0 , "dword*" , 0 )
	If @error Or Not $ARESULT [ 0 ] Then Return SetError ( @error , @extended , $AEVENT )
	Local $TEVENTLOG = DllStructCreate ( $TAGEVENTLOGRECORD , DllStructGetPtr ( $TBUFFER ) )
	$AEVENT [ 0 ] = True
	$AEVENT [ 1 ] = DllStructGetData ( $TEVENTLOG , "RecordNumber" )
	$AEVENT [ 2 ] = __EVENTLOG_DECODEDATE ( DllStructGetData ( $TEVENTLOG , "TimeGenerated" ) )
	$AEVENT [ 3 ] = __EVENTLOG_DECODETIME ( DllStructGetData ( $TEVENTLOG , "TimeGenerated" ) )
	$AEVENT [ 4 ] = __EVENTLOG_DECODEDATE ( DllStructGetData ( $TEVENTLOG , "TimeWritten" ) )
	$AEVENT [ 5 ] = __EVENTLOG_DECODETIME ( DllStructGetData ( $TEVENTLOG , "TimeWritten" ) )
	$AEVENT [ 6 ] = __EVENTLOG_DECODEEVENTID ( $TEVENTLOG )
	$AEVENT [ 7 ] = DllStructGetData ( $TEVENTLOG , "EventType" )
	$AEVENT [ 8 ] = __EVENTLOG_DECODETYPESTR ( DllStructGetData ( $TEVENTLOG , "EventType" ) )
	$AEVENT [ 9 ] = __EVENTLOG_DECODECATEGORY ( $TEVENTLOG )
	$AEVENT [ 10 ] = __EVENTLOG_DECODESOURCE ( $TEVENTLOG )
	$AEVENT [ 11 ] = __EVENTLOG_DECODECOMPUTER ( $TEVENTLOG )
	$AEVENT [ 12 ] = __EVENTLOG_DECODEUSERNAME ( $TEVENTLOG )
	$AEVENT [ 13 ] = __EVENTLOG_DECODEDESC ( $TEVENTLOG )
	$AEVENT [ 14 ] = __EVENTLOG_DECODEDATA ( $TEVENTLOG )
	Return $AEVENT
EndFunc
Global Enum $ARRAYFILL_FORCE_DEFAULT , $ARRAYFILL_FORCE_SINGLEITEM , $ARRAYFILL_FORCE_INT , $ARRAYFILL_FORCE_NUMBER , $ARRAYFILL_FORCE_PTR , $ARRAYFILL_FORCE_HWND , $ARRAYFILL_FORCE_STRING , $ARRAYFILL_FORCE_BOOLEAN
Global Enum $ARRAYUNIQUE_NOCOUNT , $ARRAYUNIQUE_COUNT
Global Enum $ARRAYUNIQUE_AUTO , $ARRAYUNIQUE_FORCE32 , $ARRAYUNIQUE_FORCE64 , $ARRAYUNIQUE_MATCH , $ARRAYUNIQUE_DISTINCT
Func _ARRAYADD ( ByRef $AARRAY , $VVALUE , $ISTART = 0 , $SDELIM_ITEM = "|" , $SDELIM_ROW = @CRLF , $IFORCE = $ARRAYFILL_FORCE_DEFAULT )
	If $ISTART = Default Then $ISTART = 0
	If $SDELIM_ITEM = Default Then $SDELIM_ITEM = "|"
	If $SDELIM_ROW = Default Then $SDELIM_ROW = @CRLF
	If $IFORCE = Default Then $IFORCE = $ARRAYFILL_FORCE_DEFAULT
	If Not IsArray ( $AARRAY ) Then Return SetError ( 1 , 0 , + 4294967295 )
	Local $IDIM_1 = UBound ( $AARRAY , $UBOUND_ROWS )
	Local $HDATATYPE = 0
	Switch $IFORCE
	Case $ARRAYFILL_FORCE_INT
		$HDATATYPE = Int
	Case $ARRAYFILL_FORCE_NUMBER
		$HDATATYPE = Number
	Case $ARRAYFILL_FORCE_PTR
		$HDATATYPE = Ptr
	Case $ARRAYFILL_FORCE_HWND
		$HDATATYPE = HWnd
	Case $ARRAYFILL_FORCE_STRING
		$HDATATYPE = String
	Case $ARRAYFILL_FORCE_BOOLEAN
		$HDATATYPE = "Boolean"
	EndSwitch
	Switch UBound ( $AARRAY , $UBOUND_DIMENSIONS )
	Case 1
		If $IFORCE = $ARRAYFILL_FORCE_SINGLEITEM Then
			ReDim $AARRAY [ $IDIM_1 + 1 ]
			$AARRAY [ $IDIM_1 ] = $VVALUE
			Return $IDIM_1
		EndIf
		If IsArray ( $VVALUE ) Then
			If UBound ( $VVALUE , $UBOUND_DIMENSIONS ) <> 1 Then Return SetError ( 5 , 0 , + 4294967295 )
			$HDATATYPE = 0
		Else
			Local $ATMP = StringSplit ( $VVALUE , $SDELIM_ITEM , $STR_NOCOUNT + $STR_ENTIRESPLIT )
			If UBound ( $ATMP , $UBOUND_ROWS ) = 1 Then
				$ATMP [ 0 ] = $VVALUE
			EndIf
			$VVALUE = $ATMP
		EndIf
		Local $IADD = UBound ( $VVALUE , $UBOUND_ROWS )
		ReDim $AARRAY [ $IDIM_1 + $IADD ]
		For $I = 0 To $IADD + 4294967295
			If String ( $HDATATYPE ) = "Boolean" Then
				Switch $VVALUE [ $I ]
				Case "True" , "1"
					$AARRAY [ $IDIM_1 + $I ] = True
				Case "False" , "0" , ""
					$AARRAY [ $IDIM_1 + $I ] = False
				EndSwitch
			ElseIf IsFunc ( $HDATATYPE ) Then
				$AARRAY [ $IDIM_1 + $I ] = $HDATATYPE ( $VVALUE [ $I ] )
			Else
				$AARRAY [ $IDIM_1 + $I ] = $VVALUE [ $I ]
			EndIf
		Next
		Return $IDIM_1 + $IADD + 4294967295
	Case 2
		Local $IDIM_2 = UBound ( $AARRAY , $UBOUND_COLUMNS )
		If $ISTART < 0 Or $ISTART > $IDIM_2 + 4294967295 Then Return SetError ( 4 , 0 , + 4294967295 )
		Local $IVALDIM_1 , $IVALDIM_2 = 0 , $ICOLCOUNT
		If IsArray ( $VVALUE ) Then
			If UBound ( $VVALUE , $UBOUND_DIMENSIONS ) <> 2 Then Return SetError ( 5 , 0 , + 4294967295 )
			$IVALDIM_1 = UBound ( $VVALUE , $UBOUND_ROWS )
			$IVALDIM_2 = UBound ( $VVALUE , $UBOUND_COLUMNS )
			$HDATATYPE = 0
		Else
			Local $ASPLIT_1 = StringSplit ( $VVALUE , $SDELIM_ROW , $STR_NOCOUNT + $STR_ENTIRESPLIT )
			$IVALDIM_1 = UBound ( $ASPLIT_1 , $UBOUND_ROWS )
			Local $ATMP [ $IVALDIM_1 ] [ 0 ] , $ASPLIT_2
			For $I = 0 To $IVALDIM_1 + 4294967295
				$ASPLIT_2 = StringSplit ( $ASPLIT_1 [ $I ] , $SDELIM_ITEM , $STR_NOCOUNT + $STR_ENTIRESPLIT )
				$ICOLCOUNT = UBound ( $ASPLIT_2 )
				If $ICOLCOUNT > $IVALDIM_2 Then
					$IVALDIM_2 = $ICOLCOUNT
					ReDim $ATMP [ $IVALDIM_1 ] [ $IVALDIM_2 ]
				EndIf
				For $J = 0 To $ICOLCOUNT + 4294967295
					$ATMP [ $I ] [ $J ] = $ASPLIT_2 [ $J ]
				Next
			Next
			$VVALUE = $ATMP
		EndIf
		If UBound ( $VVALUE , $UBOUND_COLUMNS ) + $ISTART > UBound ( $AARRAY , $UBOUND_COLUMNS ) Then Return SetError ( 3 , 0 , + 4294967295 )
		ReDim $AARRAY [ $IDIM_1 + $IVALDIM_1 ] [ $IDIM_2 ]
		For $IWRITETO_INDEX = 0 To $IVALDIM_1 + 4294967295
			For $J = 0 To $IDIM_2 + 4294967295
				If $J < $ISTART Then
					$AARRAY [ $IWRITETO_INDEX + $IDIM_1 ] [ $J ] = ""
				ElseIf $J - $ISTART > $IVALDIM_2 + 4294967295 Then
					$AARRAY [ $IWRITETO_INDEX + $IDIM_1 ] [ $J ] = ""
				Else
					If String ( $HDATATYPE ) = "Boolean" Then
						Switch $VVALUE [ $IWRITETO_INDEX ] [ $J - $ISTART ]
						Case "True" , "1"
							$AARRAY [ $IWRITETO_INDEX + $IDIM_1 ] [ $J ] = True
						Case "False" , "0" , ""
							$AARRAY [ $IWRITETO_INDEX + $IDIM_1 ] [ $J ] = False
						EndSwitch
					ElseIf IsFunc ( $HDATATYPE ) Then
						$AARRAY [ $IWRITETO_INDEX + $IDIM_1 ] [ $J ] = $HDATATYPE ( $VVALUE [ $IWRITETO_INDEX ] [ $J - $ISTART ] )
					Else
						$AARRAY [ $IWRITETO_INDEX + $IDIM_1 ] [ $J ] = $VVALUE [ $IWRITETO_INDEX ] [ $J - $ISTART ]
					EndIf
				EndIf
			Next
		Next
Case Else
		Return SetError ( 2 , 0 , + 4294967295 )
	EndSwitch
	Return UBound ( $AARRAY , $UBOUND_ROWS ) + 4294967295
EndFunc
Func _ARRAYCONCATENATE ( ByRef $AARRAYTARGET , Const ByRef $AARRAYSOURCE , $ISTART = 0 )
	If $ISTART = Default Then $ISTART = 0
	If Not IsArray ( $AARRAYTARGET ) Then Return SetError ( 1 , 0 , + 4294967295 )
	If Not IsArray ( $AARRAYSOURCE ) Then Return SetError ( 2 , 0 , + 4294967295 )
	Local $IDIM_TOTAL_TGT = UBound ( $AARRAYTARGET , $UBOUND_DIMENSIONS )
	Local $IDIM_TOTAL_SRC = UBound ( $AARRAYSOURCE , $UBOUND_DIMENSIONS )
	Local $IDIM_1_TGT = UBound ( $AARRAYTARGET , $UBOUND_ROWS )
	Local $IDIM_1_SRC = UBound ( $AARRAYSOURCE , $UBOUND_ROWS )
	If $ISTART < 0 Or $ISTART > $IDIM_1_SRC + 4294967295 Then Return SetError ( 6 , 0 , + 4294967295 )
	Switch $IDIM_TOTAL_TGT
	Case 1
		If $IDIM_TOTAL_SRC <> 1 Then Return SetError ( 4 , 0 , + 4294967295 )
		ReDim $AARRAYTARGET [ $IDIM_1_TGT + $IDIM_1_SRC - $ISTART ]
		For $I = $ISTART To $IDIM_1_SRC + 4294967295
			$AARRAYTARGET [ $IDIM_1_TGT + $I - $ISTART ] = $AARRAYSOURCE [ $I ]
		Next
	Case 2
		If $IDIM_TOTAL_SRC <> 2 Then Return SetError ( 4 , 0 , + 4294967295 )
		Local $IDIM_2_TGT = UBound ( $AARRAYTARGET , $UBOUND_COLUMNS )
		If UBound ( $AARRAYSOURCE , $UBOUND_COLUMNS ) <> $IDIM_2_TGT Then Return SetError ( 5 , 0 , + 4294967295 )
		ReDim $AARRAYTARGET [ $IDIM_1_TGT + $IDIM_1_SRC - $ISTART ] [ $IDIM_2_TGT ]
		For $I = $ISTART To $IDIM_1_SRC + 4294967295
			For $J = 0 To $IDIM_2_TGT + 4294967295
				$AARRAYTARGET [ $IDIM_1_TGT + $I - $ISTART ] [ $J ] = $AARRAYSOURCE [ $I ] [ $J ]
			Next
		Next
Case Else
		Return SetError ( 3 , 0 , + 4294967295 )
	EndSwitch
	Return UBound ( $AARRAYTARGET , $UBOUND_ROWS )
EndFunc
Func _ARRAYDELETE ( ByRef $AARRAY , $VRANGE )
	If Not IsArray ( $AARRAY ) Then Return SetError ( 1 , 0 , + 4294967295 )
	Local $IDIM_1 = UBound ( $AARRAY , $UBOUND_ROWS ) + 4294967295
	If IsArray ( $VRANGE ) Then
		If UBound ( $VRANGE , $UBOUND_DIMENSIONS ) <> 1 Or UBound ( $VRANGE , $UBOUND_ROWS ) < 2 Then Return SetError ( 4 , 0 , + 4294967295 )
	Else
		Local $INUMBER , $ASPLIT_1 , $ASPLIT_2
		$VRANGE = StringStripWS ( $VRANGE , 8 )
		$ASPLIT_1 = StringSplit ( $VRANGE , ";" )
		$VRANGE = ""
		For $I = 1 To $ASPLIT_1 [ 0 ]
			If Not StringRegExp ( $ASPLIT_1 [ $I ] , "^\d+(-\d+)?$" ) Then Return SetError ( 3 , 0 , + 4294967295 )
			$ASPLIT_2 = StringSplit ( $ASPLIT_1 [ $I ] , "-" )
			Switch $ASPLIT_2 [ 0 ]
			Case 1
				$VRANGE &= $ASPLIT_2 [ 1 ] & ";"
			Case 2
				If Number ( $ASPLIT_2 [ 2 ] ) >= Number ( $ASPLIT_2 [ 1 ] ) Then
					$INUMBER = $ASPLIT_2 [ 1 ] + 4294967295
					Do
						$INUMBER += 1
						$VRANGE &= $INUMBER & ";"
					Until $INUMBER = $ASPLIT_2 [ 2 ]
				EndIf
			EndSwitch
		Next
		$VRANGE = StringSplit ( StringTrimRight ( $VRANGE , 1 ) , ";" )
	EndIf
	For $I = 1 To $VRANGE [ 0 ]
		$VRANGE [ $I ] = Number ( $VRANGE [ $I ] )
	Next
	If $VRANGE [ 1 ] < 0 Or $VRANGE [ $VRANGE [ 0 ] ] > $IDIM_1 Then Return SetError ( 5 , 0 , + 4294967295 )
	Local $ICOPYTO_INDEX = 0
	Switch UBound ( $AARRAY , $UBOUND_DIMENSIONS )
	Case 1
		For $I = 1 To $VRANGE [ 0 ]
			$AARRAY [ $VRANGE [ $I ] ] = ChrW ( 64177 )
		Next
		For $IREADFROM_INDEX = 0 To $IDIM_1
			If $AARRAY [ $IREADFROM_INDEX ] == ChrW ( 64177 ) Then
				ContinueLoop
			Else
				If $IREADFROM_INDEX <> $ICOPYTO_INDEX Then
					$AARRAY [ $ICOPYTO_INDEX ] = $AARRAY [ $IREADFROM_INDEX ]
				EndIf
				$ICOPYTO_INDEX += 1
			EndIf
		Next
		ReDim $AARRAY [ $IDIM_1 - $VRANGE [ 0 ] + 1 ]
	Case 2
		Local $IDIM_2 = UBound ( $AARRAY , $UBOUND_COLUMNS ) + 4294967295
		For $I = 1 To $VRANGE [ 0 ]
			$AARRAY [ $VRANGE [ $I ] ] [ 0 ] = ChrW ( 64177 )
		Next
		For $IREADFROM_INDEX = 0 To $IDIM_1
			If $AARRAY [ $IREADFROM_INDEX ] [ 0 ] == ChrW ( 64177 ) Then
				ContinueLoop
			Else
				If $IREADFROM_INDEX <> $ICOPYTO_INDEX Then
					For $J = 0 To $IDIM_2
						$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = $AARRAY [ $IREADFROM_INDEX ] [ $J ]
					Next
				EndIf
				$ICOPYTO_INDEX += 1
			EndIf
		Next
		ReDim $AARRAY [ $IDIM_1 - $VRANGE [ 0 ] + 1 ] [ $IDIM_2 + 1 ]
Case Else
		Return SetError ( 2 , 0 , False )
	EndSwitch
	Return UBound ( $AARRAY , $UBOUND_ROWS )
EndFunc
Func _ARRAYINSERT ( ByRef $AARRAY , $VRANGE , $VVALUE = "" , $ISTART = 0 , $SDELIM_ITEM = "|" , $SDELIM_ROW = @CRLF , $IFORCE = $ARRAYFILL_FORCE_DEFAULT )
	If $VVALUE = Default Then $VVALUE = ""
	If $ISTART = Default Then $ISTART = 0
	If $SDELIM_ITEM = Default Then $SDELIM_ITEM = "|"
	If $SDELIM_ROW = Default Then $SDELIM_ROW = @CRLF
	If $IFORCE = Default Then $IFORCE = $ARRAYFILL_FORCE_DEFAULT
	If Not IsArray ( $AARRAY ) Then Return SetError ( 1 , 0 , + 4294967295 )
	Local $IDIM_1 = UBound ( $AARRAY , $UBOUND_ROWS ) + 4294967295
	Local $HDATATYPE = 0
	Switch $IFORCE
	Case $ARRAYFILL_FORCE_INT
		$HDATATYPE = Int
	Case $ARRAYFILL_FORCE_NUMBER
		$HDATATYPE = Number
	Case $ARRAYFILL_FORCE_PTR
		$HDATATYPE = Ptr
	Case $ARRAYFILL_FORCE_HWND
		$HDATATYPE = HWnd
	Case $ARRAYFILL_FORCE_STRING
		$HDATATYPE = String
	EndSwitch
	Local $ASPLIT_1 , $ASPLIT_2
	If IsArray ( $VRANGE ) Then
		If UBound ( $VRANGE , $UBOUND_DIMENSIONS ) <> 1 Or UBound ( $VRANGE , $UBOUND_ROWS ) < 2 Then Return SetError ( 4 , 0 , + 4294967295 )
	Else
		Local $INUMBER
		$VRANGE = StringStripWS ( $VRANGE , 8 )
		$ASPLIT_1 = StringSplit ( $VRANGE , ";" )
		$VRANGE = ""
		For $I = 1 To $ASPLIT_1 [ 0 ]
			If Not StringRegExp ( $ASPLIT_1 [ $I ] , "^\d+(-\d+)?$" ) Then Return SetError ( 3 , 0 , + 4294967295 )
			$ASPLIT_2 = StringSplit ( $ASPLIT_1 [ $I ] , "-" )
			Switch $ASPLIT_2 [ 0 ]
			Case 1
				$VRANGE &= $ASPLIT_2 [ 1 ] & ";"
			Case 2
				If Number ( $ASPLIT_2 [ 2 ] ) >= Number ( $ASPLIT_2 [ 1 ] ) Then
					$INUMBER = $ASPLIT_2 [ 1 ] + 4294967295
					Do
						$INUMBER += 1
						$VRANGE &= $INUMBER & ";"
					Until $INUMBER = $ASPLIT_2 [ 2 ]
				EndIf
			EndSwitch
		Next
		$VRANGE = StringSplit ( StringTrimRight ( $VRANGE , 1 ) , ";" )
	EndIf
	For $I = 1 To $VRANGE [ 0 ]
		$VRANGE [ $I ] = Number ( $VRANGE [ $I ] )
	Next
	If $VRANGE [ 1 ] < 0 Or $VRANGE [ $VRANGE [ 0 ] ] > $IDIM_1 Then Return SetError ( 5 , 0 , + 4294967295 )
	For $I = 2 To $VRANGE [ 0 ]
		If $VRANGE [ $I ] < $VRANGE [ $I + 4294967295 ] Then Return SetError ( 3 , 0 , + 4294967295 )
	Next
	Local $ICOPYTO_INDEX = $IDIM_1 + $VRANGE [ 0 ]
	Local $IINSERTPOINT_INDEX = $VRANGE [ 0 ]
	Local $IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
	Switch UBound ( $AARRAY , $UBOUND_DIMENSIONS )
	Case 1
		If $IFORCE = $ARRAYFILL_FORCE_SINGLEITEM Then
			ReDim $AARRAY [ $IDIM_1 + $VRANGE [ 0 ] + 1 ]
			For $IREADFROMINDEX = $IDIM_1 To 0 Step + 4294967295
				$AARRAY [ $ICOPYTO_INDEX ] = $AARRAY [ $IREADFROMINDEX ]
				$ICOPYTO_INDEX -= 1
				$IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
				While $IREADFROMINDEX = $IINSERT_INDEX
					$AARRAY [ $ICOPYTO_INDEX ] = $VVALUE
					$ICOPYTO_INDEX -= 1
					$IINSERTPOINT_INDEX -= 1
					If $IINSERTPOINT_INDEX < 1 Then ExitLoop 2
					$IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
				WEnd
			Next
			Return $IDIM_1 + $VRANGE [ 0 ] + 1
		EndIf
		ReDim $AARRAY [ $IDIM_1 + $VRANGE [ 0 ] + 1 ]
		If IsArray ( $VVALUE ) Then
			If UBound ( $VVALUE , $UBOUND_DIMENSIONS ) <> 1 Then Return SetError ( 5 , 0 , + 4294967295 )
			$HDATATYPE = 0
		Else
			Local $ATMP = StringSplit ( $VVALUE , $SDELIM_ITEM , $STR_NOCOUNT + $STR_ENTIRESPLIT )
			If UBound ( $ATMP , $UBOUND_ROWS ) = 1 Then
				$ATMP [ 0 ] = $VVALUE
				$HDATATYPE = 0
			EndIf
			$VVALUE = $ATMP
		EndIf
		For $IREADFROMINDEX = $IDIM_1 To 0 Step + 4294967295
			$AARRAY [ $ICOPYTO_INDEX ] = $AARRAY [ $IREADFROMINDEX ]
			$ICOPYTO_INDEX -= 1
			$IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
			While $IREADFROMINDEX = $IINSERT_INDEX
				If $IINSERTPOINT_INDEX <= UBound ( $VVALUE , $UBOUND_ROWS ) Then
					If IsFunc ( $HDATATYPE ) Then
						$AARRAY [ $ICOPYTO_INDEX ] = $HDATATYPE ( $VVALUE [ $IINSERTPOINT_INDEX + 4294967295 ] )
					Else
						$AARRAY [ $ICOPYTO_INDEX ] = $VVALUE [ $IINSERTPOINT_INDEX + 4294967295 ]
					EndIf
				Else
					$AARRAY [ $ICOPYTO_INDEX ] = ""
				EndIf
				$ICOPYTO_INDEX -= 1
				$IINSERTPOINT_INDEX -= 1
				If $IINSERTPOINT_INDEX = 0 Then ExitLoop 2
				$IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
			WEnd
		Next
	Case 2
		Local $IDIM_2 = UBound ( $AARRAY , $UBOUND_COLUMNS )
		If $ISTART < 0 Or $ISTART > $IDIM_2 + 4294967295 Then Return SetError ( 6 , 0 , + 4294967295 )
		Local $IVALDIM_1 , $IVALDIM_2
		If IsArray ( $VVALUE ) Then
			If UBound ( $VVALUE , $UBOUND_DIMENSIONS ) <> 2 Then Return SetError ( 7 , 0 , + 4294967295 )
			$IVALDIM_1 = UBound ( $VVALUE , $UBOUND_ROWS )
			$IVALDIM_2 = UBound ( $VVALUE , $UBOUND_COLUMNS )
			$HDATATYPE = 0
		Else
			$ASPLIT_1 = StringSplit ( $VVALUE , $SDELIM_ROW , $STR_NOCOUNT + $STR_ENTIRESPLIT )
			$IVALDIM_1 = UBound ( $ASPLIT_1 , $UBOUND_ROWS )
			StringReplace ( $ASPLIT_1 [ 0 ] , $SDELIM_ITEM , "" )
			$IVALDIM_2 = @extended + 1
			Local $ATMP [ $IVALDIM_1 ] [ $IVALDIM_2 ]
			For $I = 0 To $IVALDIM_1 + 4294967295
				$ASPLIT_2 = StringSplit ( $ASPLIT_1 [ $I ] , $SDELIM_ITEM , $STR_NOCOUNT + $STR_ENTIRESPLIT )
				For $J = 0 To $IVALDIM_2 + 4294967295
					$ATMP [ $I ] [ $J ] = $ASPLIT_2 [ $J ]
				Next
			Next
			$VVALUE = $ATMP
		EndIf
		If UBound ( $VVALUE , $UBOUND_COLUMNS ) + $ISTART > UBound ( $AARRAY , $UBOUND_COLUMNS ) Then Return SetError ( 8 , 0 , + 4294967295 )
		ReDim $AARRAY [ $IDIM_1 + $VRANGE [ 0 ] + 1 ] [ $IDIM_2 ]
		For $IREADFROMINDEX = $IDIM_1 To 0 Step + 4294967295
			For $J = 0 To $IDIM_2 + 4294967295
				$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = $AARRAY [ $IREADFROMINDEX ] [ $J ]
			Next
			$ICOPYTO_INDEX -= 1
			$IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
			While $IREADFROMINDEX = $IINSERT_INDEX
				For $J = 0 To $IDIM_2 + 4294967295
					If $J < $ISTART Then
						$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = ""
					ElseIf $J - $ISTART > $IVALDIM_2 + 4294967295 Then
						$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = ""
					Else
						If $IINSERTPOINT_INDEX + 4294967295 < $IVALDIM_1 Then
							If IsFunc ( $HDATATYPE ) Then
								$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = $HDATATYPE ( $VVALUE [ $IINSERTPOINT_INDEX + 4294967295 ] [ $J - $ISTART ] )
							Else
								$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = $VVALUE [ $IINSERTPOINT_INDEX + 4294967295 ] [ $J - $ISTART ]
							EndIf
						Else
							$AARRAY [ $ICOPYTO_INDEX ] [ $J ] = ""
						EndIf
					EndIf
				Next
				$ICOPYTO_INDEX -= 1
				$IINSERTPOINT_INDEX -= 1
				If $IINSERTPOINT_INDEX = 0 Then ExitLoop 2
				$IINSERT_INDEX = $VRANGE [ $IINSERTPOINT_INDEX ]
			WEnd
		Next
Case Else
		Return SetError ( 2 , 0 , + 4294967295 )
	EndSwitch
	Return UBound ( $AARRAY , $UBOUND_ROWS )
EndFunc
Func _ARRAYREVERSE ( ByRef $AARRAY , $ISTART = 0 , $IEND = 0 )
	If $ISTART = Default Then $ISTART = 0
	If $IEND = Default Then $IEND = 0
	If Not IsArray ( $AARRAY ) Then Return SetError ( 1 , 0 , 0 )
	If UBound ( $AARRAY , $UBOUND_DIMENSIONS ) <> 1 Then Return SetError ( 3 , 0 , 0 )
	If Not UBound ( $AARRAY ) Then Return SetError ( 4 , 0 , 0 )
	Local $VTMP , $IUBOUND = UBound ( $AARRAY ) + 4294967295
	If $IEND < 1 Or $IEND > $IUBOUND Then $IEND = $IUBOUND
	If $ISTART < 0 Then $ISTART = 0
	If $ISTART > $IEND Then Return SetError ( 2 , 0 , 0 )
	For $I = $ISTART To Int ( ( $ISTART + $IEND + 4294967295 ) / 2 )
		$VTMP = $AARRAY [ $I ]
		$AARRAY [ $I ] = $AARRAY [ $IEND ]
		$AARRAY [ $IEND ] = $VTMP
		$IEND -= 1
	Next
	Return 1
EndFunc
Func _ARRAYSORT ( ByRef $AARRAY , $IDESCENDING = 0 , $ISTART = 0 , $IEND = 0 , $ISUBITEM = 0 , $IPIVOT = 0 )
	If $IDESCENDING = Default Then $IDESCENDING = 0
	If $ISTART = Default Then $ISTART = 0
	If $IEND = Default Then $IEND = 0
	If $ISUBITEM = Default Then $ISUBITEM = 0
	If $IPIVOT = Default Then $IPIVOT = 0
	If Not IsArray ( $AARRAY ) Then Return SetError ( 1 , 0 , 0 )
	Local $IUBOUND = UBound ( $AARRAY ) + 4294967295
	If $IUBOUND = + 4294967295 Then Return SetError ( 5 , 0 , 0 )
	If $IEND = Default Then $IEND = 0
	If $IEND < 1 Or $IEND > $IUBOUND Or $IEND = Default Then $IEND = $IUBOUND
	If $ISTART < 0 Or $ISTART = Default Then $ISTART = 0
	If $ISTART > $IEND Then Return SetError ( 2 , 0 , 0 )
	Switch UBound ( $AARRAY , $UBOUND_DIMENSIONS )
	Case 1
		If $IPIVOT Then
			__ARRAYDUALPIVOTSORT ( $AARRAY , $ISTART , $IEND )
		Else
			__ARRAYQUICKSORT1D ( $AARRAY , $ISTART , $IEND )
		EndIf
		If $IDESCENDING Then _ARRAYREVERSE ( $AARRAY , $ISTART , $IEND )
	Case 2
		If $IPIVOT Then Return SetError ( 6 , 0 , 0 )
		Local $ISUBMAX = UBound ( $AARRAY , $UBOUND_COLUMNS ) + 4294967295
		If $ISUBITEM > $ISUBMAX Then Return SetError ( 3 , 0 , 0 )
		If $IDESCENDING Then
			$IDESCENDING = + 4294967295
		Else
			$IDESCENDING = 1
		EndIf
		__ARRAYQUICKSORT2D ( $AARRAY , $IDESCENDING , $ISTART , $IEND , $ISUBITEM , $ISUBMAX )
Case Else
		Return SetError ( 4 , 0 , 0 )
	EndSwitch
	Return 1
EndFunc
Func __ARRAYQUICKSORT1D ( ByRef $AARRAY , Const ByRef $ISTART , Const ByRef $IEND )
	If $IEND <= $ISTART Then Return
	Local $VTMP
	If ( $IEND - $ISTART ) < 15 Then
		Local $VCUR
		For $I = $ISTART + 1 To $IEND
			$VTMP = $AARRAY [ $I ]
			If IsNumber ( $VTMP ) Then
				For $J = $I + 4294967295 To $ISTART Step + 4294967295
					$VCUR = $AARRAY [ $J ]
					If ( $VTMP >= $VCUR And IsNumber ( $VCUR ) ) Or ( Not IsNumber ( $VCUR ) And StringCompare ( $VTMP , $VCUR ) >= 0 ) Then ExitLoop
					$AARRAY [ $J + 1 ] = $VCUR
				Next
			Else
				For $J = $I + 4294967295 To $ISTART Step + 4294967295
					If ( StringCompare ( $VTMP , $AARRAY [ $J ] ) >= 0 ) Then ExitLoop
					$AARRAY [ $J + 1 ] = $AARRAY [ $J ]
				Next
			EndIf
			$AARRAY [ $J + 1 ] = $VTMP
		Next
		Return
	EndIf
	Local $L = $ISTART , $R = $IEND , $VPIVOT = $AARRAY [ Int ( ( $ISTART + $IEND ) / 2 ) ] , $BNUM = IsNumber ( $VPIVOT )
	Do
		If $BNUM Then
			While ( $AARRAY [ $L ] < $VPIVOT And IsNumber ( $AARRAY [ $L ] ) ) Or ( Not IsNumber ( $AARRAY [ $L ] ) And StringCompare ( $AARRAY [ $L ] , $VPIVOT ) < 0 )
				$L += 1
			WEnd
			While ( $AARRAY [ $R ] > $VPIVOT And IsNumber ( $AARRAY [ $R ] ) ) Or ( Not IsNumber ( $AARRAY [ $R ] ) And StringCompare ( $AARRAY [ $R ] , $VPIVOT ) > 0 )
				$R -= 1
			WEnd
		Else
			While ( StringCompare ( $AARRAY [ $L ] , $VPIVOT ) < 0 )
				$L += 1
			WEnd
			While ( StringCompare ( $AARRAY [ $R ] , $VPIVOT ) > 0 )
				$R -= 1
			WEnd
		EndIf
		If $L <= $R Then
			$VTMP = $AARRAY [ $L ]
			$AARRAY [ $L ] = $AARRAY [ $R ]
			$AARRAY [ $R ] = $VTMP
			$L += 1
			$R -= 1
		EndIf
	Until $L > $R
	__ARRAYQUICKSORT1D ( $AARRAY , $ISTART , $R )
	__ARRAYQUICKSORT1D ( $AARRAY , $L , $IEND )
EndFunc
Func __ARRAYQUICKSORT2D ( ByRef $AARRAY , Const ByRef $ISTEP , Const ByRef $ISTART , Const ByRef $IEND , Const ByRef $ISUBITEM , Const ByRef $ISUBMAX )
	If $IEND <= $ISTART Then Return
	Local $VTMP , $L = $ISTART , $R = $IEND , $VPIVOT = $AARRAY [ Int ( ( $ISTART + $IEND ) / 2 ) ] [ $ISUBITEM ] , $BNUM = IsNumber ( $VPIVOT )
	Do
		If $BNUM Then
			While ( $ISTEP * ( $AARRAY [ $L ] [ $ISUBITEM ] - $VPIVOT ) < 0 And IsNumber ( $AARRAY [ $L ] [ $ISUBITEM ] ) ) Or ( Not IsNumber ( $AARRAY [ $L ] [ $ISUBITEM ] ) And $ISTEP * StringCompare ( $AARRAY [ $L ] [ $ISUBITEM ] , $VPIVOT ) < 0 )
				$L += 1
			WEnd
			While ( $ISTEP * ( $AARRAY [ $R ] [ $ISUBITEM ] - $VPIVOT ) > 0 And IsNumber ( $AARRAY [ $R ] [ $ISUBITEM ] ) ) Or ( Not IsNumber ( $AARRAY [ $R ] [ $ISUBITEM ] ) And $ISTEP * StringCompare ( $AARRAY [ $R ] [ $ISUBITEM ] , $VPIVOT ) > 0 )
				$R -= 1
			WEnd
		Else
			While ( $ISTEP * StringCompare ( $AARRAY [ $L ] [ $ISUBITEM ] , $VPIVOT ) < 0 )
				$L += 1
			WEnd
			While ( $ISTEP * StringCompare ( $AARRAY [ $R ] [ $ISUBITEM ] , $VPIVOT ) > 0 )
				$R -= 1
			WEnd
		EndIf
		If $L <= $R Then
			For $I = 0 To $ISUBMAX
				$VTMP = $AARRAY [ $L ] [ $I ]
				$AARRAY [ $L ] [ $I ] = $AARRAY [ $R ] [ $I ]
				$AARRAY [ $R ] [ $I ] = $VTMP
			Next
			$L += 1
			$R -= 1
		EndIf
	Until $L > $R
	__ARRAYQUICKSORT2D ( $AARRAY , $ISTEP , $ISTART , $R , $ISUBITEM , $ISUBMAX )
	__ARRAYQUICKSORT2D ( $AARRAY , $ISTEP , $L , $IEND , $ISUBITEM , $ISUBMAX )
EndFunc
Func __ARRAYDUALPIVOTSORT ( ByRef $AARRAY , $IPIVOT_LEFT , $IPIVOT_RIGHT , $BLEFTMOST = True )
	If $IPIVOT_LEFT > $IPIVOT_RIGHT Then Return
	Local $ILENGTH = $IPIVOT_RIGHT - $IPIVOT_LEFT + 1
	Local $I , $J , $K , $IAI , $IAK , $IA1 , $IA2 , $ILAST
	If $ILENGTH < 45 Then
		If $BLEFTMOST Then
			$I = $IPIVOT_LEFT
			While $I < $IPIVOT_RIGHT
				$J = $I
				$IAI = $AARRAY [ $I + 1 ]
				While $IAI < $AARRAY [ $J ]
					$AARRAY [ $J + 1 ] = $AARRAY [ $J ]
					$J -= 1
					If $J + 1 = $IPIVOT_LEFT Then ExitLoop
				WEnd
				$AARRAY [ $J + 1 ] = $IAI
				$I += 1
			WEnd
		Else
			While 1
				If $IPIVOT_LEFT >= $IPIVOT_RIGHT Then Return 1
				$IPIVOT_LEFT += 1
				If $AARRAY [ $IPIVOT_LEFT ] < $AARRAY [ $IPIVOT_LEFT + 4294967295 ] Then ExitLoop
			WEnd
			While 1
				$K = $IPIVOT_LEFT
				$IPIVOT_LEFT += 1
				If $IPIVOT_LEFT > $IPIVOT_RIGHT Then ExitLoop
				$IA1 = $AARRAY [ $K ]
				$IA2 = $AARRAY [ $IPIVOT_LEFT ]
				If $IA1 < $IA2 Then
					$IA2 = $IA1
					$IA1 = $AARRAY [ $IPIVOT_LEFT ]
				EndIf
				$K -= 1
				While $IA1 < $AARRAY [ $K ]
					$AARRAY [ $K + 2 ] = $AARRAY [ $K ]
					$K -= 1
				WEnd
				$AARRAY [ $K + 2 ] = $IA1
				While $IA2 < $AARRAY [ $K ]
					$AARRAY [ $K + 1 ] = $AARRAY [ $K ]
					$K -= 1
				WEnd
				$AARRAY [ $K + 1 ] = $IA2
				$IPIVOT_LEFT += 1
			WEnd
			$ILAST = $AARRAY [ $IPIVOT_RIGHT ]
			$IPIVOT_RIGHT -= 1
			While $ILAST < $AARRAY [ $IPIVOT_RIGHT ]
				$AARRAY [ $IPIVOT_RIGHT + 1 ] = $AARRAY [ $IPIVOT_RIGHT ]
				$IPIVOT_RIGHT -= 1
			WEnd
			$AARRAY [ $IPIVOT_RIGHT + 1 ] = $ILAST
		EndIf
		Return 1
	EndIf
	Local $ISEVENTH = BitShift ( $ILENGTH , 3 ) + BitShift ( $ILENGTH , 6 ) + 1
	Local $IE1 , $IE2 , $IE3 , $IE4 , $IE5 , $T
	$IE3 = Ceiling ( ( $IPIVOT_LEFT + $IPIVOT_RIGHT ) / 2 )
	$IE2 = $IE3 - $ISEVENTH
	$IE1 = $IE2 - $ISEVENTH
	$IE4 = $IE3 + $ISEVENTH
	$IE5 = $IE4 + $ISEVENTH
	If $AARRAY [ $IE2 ] < $AARRAY [ $IE1 ] Then
		$T = $AARRAY [ $IE2 ]
		$AARRAY [ $IE2 ] = $AARRAY [ $IE1 ]
		$AARRAY [ $IE1 ] = $T
	EndIf
	If $AARRAY [ $IE3 ] < $AARRAY [ $IE2 ] Then
		$T = $AARRAY [ $IE3 ]
		$AARRAY [ $IE3 ] = $AARRAY [ $IE2 ]
		$AARRAY [ $IE2 ] = $T
		If $T < $AARRAY [ $IE1 ] Then
			$AARRAY [ $IE2 ] = $AARRAY [ $IE1 ]
			$AARRAY [ $IE1 ] = $T
		EndIf
	EndIf
	If $AARRAY [ $IE4 ] < $AARRAY [ $IE3 ] Then
		$T = $AARRAY [ $IE4 ]
		$AARRAY [ $IE4 ] = $AARRAY [ $IE3 ]
		$AARRAY [ $IE3 ] = $T
		If $T < $AARRAY [ $IE2 ] Then
			$AARRAY [ $IE3 ] = $AARRAY [ $IE2 ]
			$AARRAY [ $IE2 ] = $T
			If $T < $AARRAY [ $IE1 ] Then
				$AARRAY [ $IE2 ] = $AARRAY [ $IE1 ]
				$AARRAY [ $IE1 ] = $T
			EndIf
		EndIf
	EndIf
	If $AARRAY [ $IE5 ] < $AARRAY [ $IE4 ] Then
		$T = $AARRAY [ $IE5 ]
		$AARRAY [ $IE5 ] = $AARRAY [ $IE4 ]
		$AARRAY [ $IE4 ] = $T
		If $T < $AARRAY [ $IE3 ] Then
			$AARRAY [ $IE4 ] = $AARRAY [ $IE3 ]
			$AARRAY [ $IE3 ] = $T
			If $T < $AARRAY [ $IE2 ] Then
				$AARRAY [ $IE3 ] = $AARRAY [ $IE2 ]
				$AARRAY [ $IE2 ] = $T
				If $T < $AARRAY [ $IE1 ] Then
					$AARRAY [ $IE2 ] = $AARRAY [ $IE1 ]
					$AARRAY [ $IE1 ] = $T
				EndIf
			EndIf
		EndIf
	EndIf
	Local $ILESS = $IPIVOT_LEFT
	Local $IGREATER = $IPIVOT_RIGHT
	If ( ( $AARRAY [ $IE1 ] <> $AARRAY [ $IE2 ] ) And ( $AARRAY [ $IE2 ] <> $AARRAY [ $IE3 ] ) And ( $AARRAY [ $IE3 ] <> $AARRAY [ $IE4 ] ) And ( $AARRAY [ $IE4 ] <> $AARRAY [ $IE5 ] ) ) Then
		Local $IPIVOT_1 = $AARRAY [ $IE2 ]
		Local $IPIVOT_2 = $AARRAY [ $IE4 ]
		$AARRAY [ $IE2 ] = $AARRAY [ $IPIVOT_LEFT ]
		$AARRAY [ $IE4 ] = $AARRAY [ $IPIVOT_RIGHT ]
		Do
			$ILESS += 1
		Until $AARRAY [ $ILESS ] >= $IPIVOT_1
		Do
			$IGREATER -= 1
		Until $AARRAY [ $IGREATER ] <= $IPIVOT_2
		$K = $ILESS
		While $K <= $IGREATER
			$IAK = $AARRAY [ $K ]
			If $IAK < $IPIVOT_1 Then
				$AARRAY [ $K ] = $AARRAY [ $ILESS ]
				$AARRAY [ $ILESS ] = $IAK
				$ILESS += 1
			ElseIf $IAK > $IPIVOT_2 Then
				While $AARRAY [ $IGREATER ] > $IPIVOT_2
					$IGREATER -= 1
					If $IGREATER + 1 = $K Then ExitLoop 2
				WEnd
				If $AARRAY [ $IGREATER ] < $IPIVOT_1 Then
					$AARRAY [ $K ] = $AARRAY [ $ILESS ]
					$AARRAY [ $ILESS ] = $AARRAY [ $IGREATER ]
					$ILESS += 1
				Else
					$AARRAY [ $K ] = $AARRAY [ $IGREATER ]
				EndIf
				$AARRAY [ $IGREATER ] = $IAK
				$IGREATER -= 1
			EndIf
			$K += 1
		WEnd
		$AARRAY [ $IPIVOT_LEFT ] = $AARRAY [ $ILESS + 4294967295 ]
		$AARRAY [ $ILESS + 4294967295 ] = $IPIVOT_1
		$AARRAY [ $IPIVOT_RIGHT ] = $AARRAY [ $IGREATER + 1 ]
		$AARRAY [ $IGREATER + 1 ] = $IPIVOT_2
		__ARRAYDUALPIVOTSORT ( $AARRAY , $IPIVOT_LEFT , $ILESS + 4294967294 , True )
		__ARRAYDUALPIVOTSORT ( $AARRAY , $IGREATER + 2 , $IPIVOT_RIGHT , False )
		If ( $ILESS < $IE1 ) And ( $IE5 < $IGREATER ) Then
			While $AARRAY [ $ILESS ] = $IPIVOT_1
				$ILESS += 1
			WEnd
			While $AARRAY [ $IGREATER ] = $IPIVOT_2
				$IGREATER -= 1
			WEnd
			$K = $ILESS
			While $K <= $IGREATER
				$IAK = $AARRAY [ $K ]
				If $IAK = $IPIVOT_1 Then
					$AARRAY [ $K ] = $AARRAY [ $ILESS ]
					$AARRAY [ $ILESS ] = $IAK
					$ILESS += 1
				ElseIf $IAK = $IPIVOT_2 Then
					While $AARRAY [ $IGREATER ] = $IPIVOT_2
						$IGREATER -= 1
						If $IGREATER + 1 = $K Then ExitLoop 2
					WEnd
					If $AARRAY [ $IGREATER ] = $IPIVOT_1 Then
						$AARRAY [ $K ] = $AARRAY [ $ILESS ]
						$AARRAY [ $ILESS ] = $IPIVOT_1
						$ILESS += 1
					Else
						$AARRAY [ $K ] = $AARRAY [ $IGREATER ]
					EndIf
					$AARRAY [ $IGREATER ] = $IAK
					$IGREATER -= 1
				EndIf
				$K += 1
			WEnd
		EndIf
		__ARRAYDUALPIVOTSORT ( $AARRAY , $ILESS , $IGREATER , False )
	Else
		Local $IPIVOT = $AARRAY [ $IE3 ]
		$K = $ILESS
		While $K <= $IGREATER
			If $AARRAY [ $K ] = $IPIVOT Then
				$K += 1
				ContinueLoop
			EndIf
			$IAK = $AARRAY [ $K ]
			If $IAK < $IPIVOT Then
				$AARRAY [ $K ] = $AARRAY [ $ILESS ]
				$AARRAY [ $ILESS ] = $IAK
				$ILESS += 1
			Else
				While $AARRAY [ $IGREATER ] > $IPIVOT
					$IGREATER -= 1
				WEnd
				If $AARRAY [ $IGREATER ] < $IPIVOT Then
					$AARRAY [ $K ] = $AARRAY [ $ILESS ]
					$AARRAY [ $ILESS ] = $AARRAY [ $IGREATER ]
					$ILESS += 1
				Else
					$AARRAY [ $K ] = $IPIVOT
				EndIf
				$AARRAY [ $IGREATER ] = $IAK
				$IGREATER -= 1
			EndIf
			$K += 1
		WEnd
		__ARRAYDUALPIVOTSORT ( $AARRAY , $IPIVOT_LEFT , $ILESS + 4294967295 , True )
		__ARRAYDUALPIVOTSORT ( $AARRAY , $IGREATER + 1 , $IPIVOT_RIGHT , False )
	EndIf
EndFunc
Func _ARRAYUNIQUE ( Const ByRef $AARRAY , $ICOLUMN = 0 , $IBASE = 0 , $ICASE = 0 , $ICOUNT = $ARRAYUNIQUE_COUNT , $IINTTYPE = $ARRAYUNIQUE_AUTO )
	If $ICOLUMN = Default Then $ICOLUMN = 0
	If $IBASE = Default Then $IBASE = 0
	If $ICASE = Default Then $ICASE = 0
	If $ICOUNT = Default Then $ICOUNT = $ARRAYUNIQUE_COUNT
	If $IINTTYPE = Default Then $IINTTYPE = $ARRAYUNIQUE_AUTO
	If UBound ( $AARRAY , $UBOUND_ROWS ) = 0 Then Return SetError ( 1 , 0 , 0 )
	Local $IDIMS = UBound ( $AARRAY , $UBOUND_DIMENSIONS ) , $INUMCOLUMNS = UBound ( $AARRAY , $UBOUND_COLUMNS )
	If $IDIMS > 2 Then Return SetError ( 2 , 0 , 0 )
	If $IBASE < 0 Or $IBASE > 1 Or ( Not IsInt ( $IBASE ) ) Then Return SetError ( 3 , 0 , 0 )
	If $ICASE < 0 Or $ICASE > 1 Or ( Not IsInt ( $ICASE ) ) Then Return SetError ( 3 , 0 , 0 )
	If $ICOUNT < 0 Or $ICOUNT > 1 Or ( Not IsInt ( $ICOUNT ) ) Then Return SetError ( 4 , 0 , 0 )
	If $IINTTYPE < 0 Or $IINTTYPE > 4 Or ( Not IsInt ( $IINTTYPE ) ) Then Return SetError ( 5 , 0 , 0 )
	If $ICOLUMN < 0 Or ( $INUMCOLUMNS = 0 And $ICOLUMN > 0 ) Or ( $INUMCOLUMNS > 0 And $ICOLUMN >= $INUMCOLUMNS ) Then Return SetError ( 6 , 0 , 0 )
	If $IINTTYPE = $ARRAYUNIQUE_AUTO Then
		Local $BINT , $SVARTYPE
		If $IDIMS = 1 Then
			$BINT = IsInt ( $AARRAY [ $IBASE ] )
			$SVARTYPE = VarGetType ( $AARRAY [ $IBASE ] )
		Else
			$BINT = IsInt ( $AARRAY [ $IBASE ] [ $ICOLUMN ] )
			$SVARTYPE = VarGetType ( $AARRAY [ $IBASE ] [ $ICOLUMN ] )
		EndIf
		If $BINT And $SVARTYPE = "Int64" Then
			$IINTTYPE = $ARRAYUNIQUE_FORCE64
		Else
			$IINTTYPE = $ARRAYUNIQUE_FORCE32
		EndIf
	EndIf
	ObjEvent ( "AutoIt.Error" , __ARRAYUNIQUE_AUTOERRFUNC )
	Local $ODICTIONARY = ObjCreate ( "Scripting.Dictionary" )
	$ODICTIONARY .CompareMode = Number ( Not $ICASE )
	Local $VELEM , $STYPE , $VKEY , $BCOMERROR = False
	For $I = $IBASE To UBound ( $AARRAY ) + 4294967295
		If $IDIMS = 1 Then
			$VELEM = $AARRAY [ $I ]
		Else
			$VELEM = $AARRAY [ $I ] [ $ICOLUMN ]
		EndIf
		Switch $IINTTYPE
		Case $ARRAYUNIQUE_FORCE32
			$ODICTIONARY .Item ( $VELEM )
			If @error Then
				$BCOMERROR = True
				ExitLoop
			EndIf
		Case $ARRAYUNIQUE_FORCE64
			$STYPE = VarGetType ( $VELEM )
			If $STYPE = "Int32" Then
				$BCOMERROR = True
				ExitLoop
			EndIf
			$VKEY = "#" & $STYPE & "#" & String ( $VELEM )
			If Not $ODICTIONARY .Item ( $VKEY ) Then
				$ODICTIONARY ( $VKEY ) = $VELEM
			EndIf
		Case $ARRAYUNIQUE_MATCH
			$STYPE = VarGetType ( $VELEM )
			If StringLeft ( $STYPE , 3 ) = "Int" Then
				$VKEY = "#Int#" & String ( $VELEM )
			Else
				$VKEY = "#" & $STYPE & "#" & String ( $VELEM )
			EndIf
			If Not $ODICTIONARY .Item ( $VKEY ) Then
				$ODICTIONARY ( $VKEY ) = $VELEM
			EndIf
		Case $ARRAYUNIQUE_DISTINCT
			$VKEY = "#" & VarGetType ( $VELEM ) & "#" & String ( $VELEM )
			If Not $ODICTIONARY .Item ( $VKEY ) Then
				$ODICTIONARY ( $VKEY ) = $VELEM
			EndIf
		EndSwitch
	Next
	Local $AVALUES , $J = 0
	If $BCOMERROR Then
		Return SetError ( 7 , 0 , 0 )
	ElseIf $IINTTYPE <> $ARRAYUNIQUE_FORCE32 Then
		Local $AVALUES [ $ODICTIONARY .Count ]
		For $VKEY In $ODICTIONARY .Keys ( )
			$AVALUES [ $J ] = $ODICTIONARY ( $VKEY )
			If StringLeft ( $VKEY , 5 ) = "#Ptr#" Then
				$AVALUES [ $J ] = Ptr ( $AVALUES [ $J ] )
			EndIf
			$J += 1
		Next
	Else
		$AVALUES = $ODICTIONARY .Keys ( )
	EndIf
	If $ICOUNT Then
		_ARRAYINSERT ( $AVALUES , 0 , $ODICTIONARY .Count )
	EndIf
	Return $AVALUES
EndFunc
Func __ARRAYUNIQUE_AUTOERRFUNC ( )
EndFunc
Func _FILECOUNTLINES ( $SFILEPATH )
	FileReadToArray ( $SFILEPATH )
	If @error Then Return SetError ( @error , @extended , 0 )
	Return @extended
EndFunc
Func _FILELISTTOARRAY ( $SFILEPATH , $SFILTER = "*" , $IFLAG = $FLTA_FILESFOLDERS , $BRETURNPATH = False )
	Local $SDELIMITER = "|" , $SFILELIST = "" , $SFILENAME = "" , $SFULLPATH = ""
	$SFILEPATH = StringRegExpReplace ( $SFILEPATH , "[\\/]+$" , "" ) & "\"
	If $IFLAG = Default Then $IFLAG = $FLTA_FILESFOLDERS
	If $BRETURNPATH Then $SFULLPATH = $SFILEPATH
	If $SFILTER = Default Then $SFILTER = "*"
	If Not FileExists ( $SFILEPATH ) Then Return SetError ( 1 , 0 , 0 )
	If StringRegExp ( $SFILTER , "[\\/:><\|]|(?s)^\s*$" ) Then Return SetError ( 2 , 0 , 0 )
	If Not ( $IFLAG = 0 Or $IFLAG = 1 Or $IFLAG = 2 ) Then Return SetError ( 3 , 0 , 0 )
	Local $HSEARCH = FileFindFirstFile ( $SFILEPATH & $SFILTER )
	If @error Then Return SetError ( 4 , 0 , 0 )
	While 1
		$SFILENAME = FileFindNextFile ( $HSEARCH )
		If @error Then ExitLoop
		If ( $IFLAG + @extended = 2 ) Then ContinueLoop
		$SFILELIST &= $SDELIMITER & $SFULLPATH & $SFILENAME
	WEnd
	FileClose ( $HSEARCH )
	If $SFILELIST = "" Then Return SetError ( 4 , 0 , 0 )
	Return StringSplit ( StringTrimLeft ( $SFILELIST , 1 ) , $SDELIMITER )
EndFunc
Func _FILELISTTOARRAYREC ( $SFILEPATH , $SMASK = "*" , $IRETURN = $FLTAR_FILESFOLDERS , $IRECUR = $FLTAR_NORECUR , $ISORT = $FLTAR_NOSORT , $IRETURNPATH = $FLTAR_RELPATH )
	If Not FileExists ( $SFILEPATH ) Then Return SetError ( 1 , 1 , "" )
	If $SMASK = Default Then $SMASK = "*"
	If $IRETURN = Default Then $IRETURN = $FLTAR_FILESFOLDERS
	If $IRECUR = Default Then $IRECUR = $FLTAR_NORECUR
	If $ISORT = Default Then $ISORT = $FLTAR_NOSORT
	If $IRETURNPATH = Default Then $IRETURNPATH = $FLTAR_RELPATH
	If $IRECUR > 1 Or Not IsInt ( $IRECUR ) Then Return SetError ( 1 , 6 , "" )
	Local $BLONGPATH = False
	If StringLeft ( $SFILEPATH , 4 ) == "\\?\" Then
		$BLONGPATH = True
	EndIf
	Local $SFOLDERSLASH = ""
	If StringRight ( $SFILEPATH , 1 ) = "\" Then
		$SFOLDERSLASH = "\"
	Else
		$SFILEPATH = $SFILEPATH & "\"
	EndIf
	Local $ASFOLDERSEARCHLIST [ 100 ] = [ 1 ]
	$ASFOLDERSEARCHLIST [ 1 ] = $SFILEPATH
	Local $IHIDE_HS = 0 , $SHIDE_HS = ""
	If BitAND ( $IRETURN , $FLTAR_NOHIDDEN ) Then
		$IHIDE_HS += 2
		$SHIDE_HS &= "H"
		$IRETURN -= $FLTAR_NOHIDDEN
	EndIf
	If BitAND ( $IRETURN , $FLTAR_NOSYSTEM ) Then
		$IHIDE_HS += 4
		$SHIDE_HS &= "S"
		$IRETURN -= $FLTAR_NOSYSTEM
	EndIf
	Local $IHIDE_LINK = 0
	If BitAND ( $IRETURN , $FLTAR_NOLINK ) Then
		$IHIDE_LINK = 1024
		$IRETURN -= $FLTAR_NOLINK
	EndIf
	Local $IMAXLEVEL = 0
	If $IRECUR < 0 Then
		StringReplace ( $SFILEPATH , "\" , "" , 0 , $STR_NOCASESENSEBASIC )
		$IMAXLEVEL = @extended - $IRECUR
	EndIf
	Local $SEXCLUDE_LIST = "" , $SEXCLUDE_LIST_FOLDER = "" , $SINCLUDE_LIST = "*"
	Local $AMASKSPLIT = StringSplit ( $SMASK , "|" )
	Switch $AMASKSPLIT [ 0 ]
	Case 3
		$SEXCLUDE_LIST_FOLDER = $AMASKSPLIT [ 3 ]
		ContinueCase
	Case 2
		$SEXCLUDE_LIST = $AMASKSPLIT [ 2 ]
		ContinueCase
	Case 1
		$SINCLUDE_LIST = $AMASKSPLIT [ 1 ]
	EndSwitch
	Local $SINCLUDE_FILE_MASK = ".+"
	If $SINCLUDE_LIST <> "*" Then
		If Not __FLTAR_LISTTOMASK ( $SINCLUDE_FILE_MASK , $SINCLUDE_LIST ) Then Return SetError ( 1 , 2 , "" )
	EndIf
	Local $SINCLUDE_FOLDER_MASK = ".+"
	Switch $IRETURN
	Case 0
		Switch $IRECUR
		Case 0
			$SINCLUDE_FOLDER_MASK = $SINCLUDE_FILE_MASK
		EndSwitch
	Case 2
		$SINCLUDE_FOLDER_MASK = $SINCLUDE_FILE_MASK
	EndSwitch
	Local $SEXCLUDE_FILE_MASK = ":"
	If $SEXCLUDE_LIST <> "" Then
		If Not __FLTAR_LISTTOMASK ( $SEXCLUDE_FILE_MASK , $SEXCLUDE_LIST ) Then Return SetError ( 1 , 3 , "" )
	EndIf
	Local $SEXCLUDE_FOLDER_MASK = ":"
	If $IRECUR Then
		If $SEXCLUDE_LIST_FOLDER Then
			If Not __FLTAR_LISTTOMASK ( $SEXCLUDE_FOLDER_MASK , $SEXCLUDE_LIST_FOLDER ) Then Return SetError ( 1 , 4 , "" )
		EndIf
		If $IRETURN = 2 Then
			$SEXCLUDE_FOLDER_MASK = $SEXCLUDE_FILE_MASK
		EndIf
	Else
		$SEXCLUDE_FOLDER_MASK = $SEXCLUDE_FILE_MASK
	EndIf
	If Not ( $IRETURN = 0 Or $IRETURN = 1 Or $IRETURN = 2 ) Then Return SetError ( 1 , 5 , "" )
	If Not ( $ISORT = 0 Or $ISORT = 1 Or $ISORT = 2 ) Then Return SetError ( 1 , 7 , "" )
	If Not ( $IRETURNPATH = 0 Or $IRETURNPATH = 1 Or $IRETURNPATH = 2 ) Then Return SetError ( 1 , 8 , "" )
	If $IHIDE_LINK Then
		Local $TFILE_DATA = DllStructCreate ( "struct;align 4;dword FileAttributes;uint64 CreationTime;uint64 LastAccessTime;uint64 LastWriteTime;" & "dword FileSizeHigh;dword FileSizeLow;dword Reserved0;dword Reserved1;wchar FileName[260];wchar AlternateFileName[14];endstruct" )
		Local $HDLL = DllOpen ( "kernel32.dll" ) , $ADLL_RET
	EndIf
	Local $ASRETURNLIST [ 100 ] = [ 0 ]
	Local $ASFILEMATCHLIST = $ASRETURNLIST , $ASROOTFILEMATCHLIST = $ASRETURNLIST , $ASFOLDERMATCHLIST = $ASRETURNLIST
	Local $BFOLDER = False , $HSEARCH = 0 , $SCURRENTPATH = "" , $SNAME = "" , $SRETPATH = ""
	Local $IATTRIBS = 0 , $SATTRIBS = ""
	Local $ASFOLDERFILESECTIONLIST [ 100 ] [ 2 ] = [ [ 0 , 0 ] ]
	While $ASFOLDERSEARCHLIST [ 0 ] > 0
		$SCURRENTPATH = $ASFOLDERSEARCHLIST [ $ASFOLDERSEARCHLIST [ 0 ] ]
		$ASFOLDERSEARCHLIST [ 0 ] -= 1
		Switch $IRETURNPATH
		Case 1
			$SRETPATH = StringReplace ( $SCURRENTPATH , $SFILEPATH , "" )
		Case 2
			If $BLONGPATH Then
				$SRETPATH = StringTrimLeft ( $SCURRENTPATH , 4 )
			Else
				$SRETPATH = $SCURRENTPATH
			EndIf
		EndSwitch
		If $IHIDE_LINK Then
			$ADLL_RET = DllCall ( $HDLL , "handle" , "FindFirstFileW" , "wstr" , $SCURRENTPATH & "*" , "struct*" , $TFILE_DATA )
			If @error Or Not $ADLL_RET [ 0 ] Then
				ContinueLoop
			EndIf
			$HSEARCH = $ADLL_RET [ 0 ]
		Else
			$HSEARCH = FileFindFirstFile ( $SCURRENTPATH & "*" )
			If $HSEARCH = + 4294967295 Then
				ContinueLoop
			EndIf
		EndIf
		If $IRETURN = 0 And $ISORT And $IRETURNPATH Then
			__FLTAR_ADDTOLIST ( $ASFOLDERFILESECTIONLIST , $SRETPATH , $ASFILEMATCHLIST [ 0 ] + 1 )
		EndIf
		$SATTRIBS = ""
		While 1
			If $IHIDE_LINK Then
				$ADLL_RET = DllCall ( $HDLL , "int" , "FindNextFileW" , "handle" , $HSEARCH , "struct*" , $TFILE_DATA )
				If @error Or Not $ADLL_RET [ 0 ] Then
					ExitLoop
				EndIf
				$SNAME = DllStructGetData ( $TFILE_DATA , "FileName" )
				If $SNAME = ".." Or $SNAME = "." Then
					ContinueLoop
				EndIf
				$IATTRIBS = DllStructGetData ( $TFILE_DATA , "FileAttributes" )
				If $IHIDE_HS And BitAND ( $IATTRIBS , $IHIDE_HS ) Then
					ContinueLoop
				EndIf
				If BitAND ( $IATTRIBS , $IHIDE_LINK ) Then
					ContinueLoop
				EndIf
				$BFOLDER = False
				If BitAND ( $IATTRIBS , 16 ) Then
					$BFOLDER = True
				EndIf
			Else
				$BFOLDER = False
				$SNAME = FileFindNextFile ( $HSEARCH , 1 )
				If @error Then
					ExitLoop
				EndIf
				If $SNAME = ".." Or $SNAME = "." Then
					ContinueLoop
				EndIf
				$SATTRIBS = @extended
				If StringInStr ( $SATTRIBS , "D" ) Then
					$BFOLDER = True
				EndIf
				If StringRegExp ( $SATTRIBS , "[" & $SHIDE_HS & "]" ) Then
					ContinueLoop
				EndIf
			EndIf
			If $BFOLDER Then
				Select
				Case $IRECUR < 0
					StringReplace ( $SCURRENTPATH , "\" , "" , 0 , $STR_NOCASESENSEBASIC )
					If @extended < $IMAXLEVEL Then
						ContinueCase
					EndIf
				Case $IRECUR = 1
					If Not StringRegExp ( $SNAME , $SEXCLUDE_FOLDER_MASK ) Then
						__FLTAR_ADDTOLIST ( $ASFOLDERSEARCHLIST , $SCURRENTPATH & $SNAME & "\" )
					EndIf
				EndSelect
			EndIf
			If $ISORT Then
				If $BFOLDER Then
					If StringRegExp ( $SNAME , $SINCLUDE_FOLDER_MASK ) And Not StringRegExp ( $SNAME , $SEXCLUDE_FOLDER_MASK ) Then
						__FLTAR_ADDTOLIST ( $ASFOLDERMATCHLIST , $SRETPATH & $SNAME & $SFOLDERSLASH )
					EndIf
				Else
					If StringRegExp ( $SNAME , $SINCLUDE_FILE_MASK ) And Not StringRegExp ( $SNAME , $SEXCLUDE_FILE_MASK ) Then
						If $SCURRENTPATH = $SFILEPATH Then
							__FLTAR_ADDTOLIST ( $ASROOTFILEMATCHLIST , $SRETPATH & $SNAME )
						Else
							__FLTAR_ADDTOLIST ( $ASFILEMATCHLIST , $SRETPATH & $SNAME )
						EndIf
					EndIf
				EndIf
			Else
				If $BFOLDER Then
					If $IRETURN <> 1 And StringRegExp ( $SNAME , $SINCLUDE_FOLDER_MASK ) And Not StringRegExp ( $SNAME , $SEXCLUDE_FOLDER_MASK ) Then
						__FLTAR_ADDTOLIST ( $ASRETURNLIST , $SRETPATH & $SNAME & $SFOLDERSLASH )
					EndIf
				Else
					If $IRETURN <> 2 And StringRegExp ( $SNAME , $SINCLUDE_FILE_MASK ) And Not StringRegExp ( $SNAME , $SEXCLUDE_FILE_MASK ) Then
						__FLTAR_ADDTOLIST ( $ASRETURNLIST , $SRETPATH & $SNAME )
					EndIf
				EndIf
			EndIf
		WEnd
		If $IHIDE_LINK Then
			DllCall ( $HDLL , "int" , "FindClose" , "ptr" , $HSEARCH )
		Else
			FileClose ( $HSEARCH )
		EndIf
	WEnd
	If $IHIDE_LINK Then
		DllClose ( $HDLL )
	EndIf
	If $ISORT Then
		Switch $IRETURN
		Case 2
			If $ASFOLDERMATCHLIST [ 0 ] = 0 Then Return SetError ( 1 , 9 , "" )
			ReDim $ASFOLDERMATCHLIST [ $ASFOLDERMATCHLIST [ 0 ] + 1 ]
			$ASRETURNLIST = $ASFOLDERMATCHLIST
			__ARRAYDUALPIVOTSORT ( $ASRETURNLIST , 1 , $ASRETURNLIST [ 0 ] )
		Case 1
			If $ASROOTFILEMATCHLIST [ 0 ] = 0 And $ASFILEMATCHLIST [ 0 ] = 0 Then Return SetError ( 1 , 9 , "" )
			If $IRETURNPATH = 0 Then
				__FLTAR_ADDFILELISTS ( $ASRETURNLIST , $ASROOTFILEMATCHLIST , $ASFILEMATCHLIST )
				__ARRAYDUALPIVOTSORT ( $ASRETURNLIST , 1 , $ASRETURNLIST [ 0 ] )
			Else
				__FLTAR_ADDFILELISTS ( $ASRETURNLIST , $ASROOTFILEMATCHLIST , $ASFILEMATCHLIST , 1 )
			EndIf
		Case 0
			If $ASROOTFILEMATCHLIST [ 0 ] = 0 And $ASFOLDERMATCHLIST [ 0 ] = 0 Then Return SetError ( 1 , 9 , "" )
			If $IRETURNPATH = 0 Then
				__FLTAR_ADDFILELISTS ( $ASRETURNLIST , $ASROOTFILEMATCHLIST , $ASFILEMATCHLIST )
				$ASRETURNLIST [ 0 ] += $ASFOLDERMATCHLIST [ 0 ]
				ReDim $ASFOLDERMATCHLIST [ $ASFOLDERMATCHLIST [ 0 ] + 1 ]
				_ARRAYCONCATENATE ( $ASRETURNLIST , $ASFOLDERMATCHLIST , 1 )
				__ARRAYDUALPIVOTSORT ( $ASRETURNLIST , 1 , $ASRETURNLIST [ 0 ] )
			Else
				Local $ASRETURNLIST [ $ASFILEMATCHLIST [ 0 ] + $ASROOTFILEMATCHLIST [ 0 ] + $ASFOLDERMATCHLIST [ 0 ] + 1 ]
				$ASRETURNLIST [ 0 ] = $ASFILEMATCHLIST [ 0 ] + $ASROOTFILEMATCHLIST [ 0 ] + $ASFOLDERMATCHLIST [ 0 ]
				__ARRAYDUALPIVOTSORT ( $ASROOTFILEMATCHLIST , 1 , $ASROOTFILEMATCHLIST [ 0 ] )
				For $I = 1 To $ASROOTFILEMATCHLIST [ 0 ]
					$ASRETURNLIST [ $I ] = $ASROOTFILEMATCHLIST [ $I ]
				Next
				Local $INEXTINSERTIONINDEX = $ASROOTFILEMATCHLIST [ 0 ] + 1
				__ARRAYDUALPIVOTSORT ( $ASFOLDERMATCHLIST , 1 , $ASFOLDERMATCHLIST [ 0 ] )
				Local $SFOLDERTOFIND = ""
				For $I = 1 To $ASFOLDERMATCHLIST [ 0 ]
					$ASRETURNLIST [ $INEXTINSERTIONINDEX ] = $ASFOLDERMATCHLIST [ $I ]
					$INEXTINSERTIONINDEX += 1
					If $SFOLDERSLASH Then
						$SFOLDERTOFIND = $ASFOLDERMATCHLIST [ $I ]
					Else
						$SFOLDERTOFIND = $ASFOLDERMATCHLIST [ $I ] & "\"
					EndIf
					Local $IFILESECTIONENDINDEX = 0 , $IFILESECTIONSTARTINDEX = 0
					For $J = 1 To $ASFOLDERFILESECTIONLIST [ 0 ] [ 0 ]
						If $SFOLDERTOFIND = $ASFOLDERFILESECTIONLIST [ $J ] [ 0 ] Then
							$IFILESECTIONSTARTINDEX = $ASFOLDERFILESECTIONLIST [ $J ] [ 1 ]
							If $J = $ASFOLDERFILESECTIONLIST [ 0 ] [ 0 ] Then
								$IFILESECTIONENDINDEX = $ASFILEMATCHLIST [ 0 ]
							Else
								$IFILESECTIONENDINDEX = $ASFOLDERFILESECTIONLIST [ $J + 1 ] [ 1 ] + 4294967295
							EndIf
							If $ISORT = 1 Then
								__ARRAYDUALPIVOTSORT ( $ASFILEMATCHLIST , $IFILESECTIONSTARTINDEX , $IFILESECTIONENDINDEX )
							EndIf
							For $K = $IFILESECTIONSTARTINDEX To $IFILESECTIONENDINDEX
								$ASRETURNLIST [ $INEXTINSERTIONINDEX ] = $ASFILEMATCHLIST [ $K ]
								$INEXTINSERTIONINDEX += 1
							Next
							ExitLoop
						EndIf
					Next
				Next
			EndIf
		EndSwitch
	Else
		If $ASRETURNLIST [ 0 ] = 0 Then Return SetError ( 1 , 9 , "" )
		ReDim $ASRETURNLIST [ $ASRETURNLIST [ 0 ] + 1 ]
	EndIf
	Return $ASRETURNLIST
EndFunc
Func __FLTAR_ADDFILELISTS ( ByRef $ASTARGET , $ASSOURCE_1 , $ASSOURCE_2 , $ISORT = 0 )
	ReDim $ASSOURCE_1 [ $ASSOURCE_1 [ 0 ] + 1 ]
	If $ISORT = 1 Then __ARRAYDUALPIVOTSORT ( $ASSOURCE_1 , 1 , $ASSOURCE_1 [ 0 ] )
	$ASTARGET = $ASSOURCE_1
	$ASTARGET [ 0 ] += $ASSOURCE_2 [ 0 ]
	ReDim $ASSOURCE_2 [ $ASSOURCE_2 [ 0 ] + 1 ]
	If $ISORT = 1 Then __ARRAYDUALPIVOTSORT ( $ASSOURCE_2 , 1 , $ASSOURCE_2 [ 0 ] )
	_ARRAYCONCATENATE ( $ASTARGET , $ASSOURCE_2 , 1 )
EndFunc
Func __FLTAR_ADDTOLIST ( ByRef $ALIST , $VVALUE_0 , $VVALUE_1 = + 4294967295 )
	If $VVALUE_1 = + 4294967295 Then
		$ALIST [ 0 ] += 1
		If UBound ( $ALIST ) <= $ALIST [ 0 ] Then ReDim $ALIST [ UBound ( $ALIST ) * 2 ]
		$ALIST [ $ALIST [ 0 ] ] = $VVALUE_0
	Else
		$ALIST [ 0 ] [ 0 ] += 1
		If UBound ( $ALIST ) <= $ALIST [ 0 ] [ 0 ] Then ReDim $ALIST [ UBound ( $ALIST ) * 2 ] [ 2 ]
		$ALIST [ $ALIST [ 0 ] [ 0 ] ] [ 0 ] = $VVALUE_0
		$ALIST [ $ALIST [ 0 ] [ 0 ] ] [ 1 ] = $VVALUE_1
	EndIf
EndFunc
Func __FLTAR_LISTTOMASK ( ByRef $SMASK , $SLIST )
	If StringRegExp ( $SLIST , "\\|/|:|\<|\>|\|" ) Then Return 0
	$SLIST = StringReplace ( StringStripWS ( StringRegExpReplace ( $SLIST , "\s*;\s*" , ";" ) , BitOR ( $STR_STRIPLEADING , $STR_STRIPTRAILING ) ) , ";" , "|" )
	$SLIST = StringReplace ( StringReplace ( StringRegExpReplace ( $SLIST , "[][$^.{}()+\-]" , "\\$0" ) , "?" , "." ) , "*" , ".*?" )
	$SMASK = "(?i)^(" & $SLIST & ")\z"
	Return 1
EndFunc
Func _FILEWRITEFROMARRAY ( $SFILEPATH , Const ByRef $AARRAY , $IBASE = Default , $IUBOUND = Default , $SDELIMITER = "|" )
	Local $IRETURN = 0
	If Not IsArray ( $AARRAY ) Then Return SetError ( 2 , 0 , $IRETURN )
	Local $IDIMS = UBound ( $AARRAY , $UBOUND_DIMENSIONS )
	If $IDIMS > 2 Then Return SetError ( 4 , 0 , 0 )
	Local $ILAST = UBound ( $AARRAY ) + 4294967295
	If $IUBOUND = Default Or $IUBOUND > $ILAST Then $IUBOUND = $ILAST
	If $IBASE < 0 Or $IBASE = Default Then $IBASE = 0
	If $IBASE > $IUBOUND Then Return SetError ( 5 , 0 , $IRETURN )
	If $SDELIMITER = Default Then $SDELIMITER = "|"
	Local $HFILEOPEN = $SFILEPATH
	If IsString ( $SFILEPATH ) Then
		$HFILEOPEN = FileOpen ( $SFILEPATH , $FO_OVERWRITE )
		If $HFILEOPEN = + 4294967295 Then Return SetError ( 1 , 0 , $IRETURN )
	EndIf
	Local $IERROR = 0
	$IRETURN = 1
	Switch $IDIMS
	Case 1
		For $I = $IBASE To $IUBOUND
			If Not FileWrite ( $HFILEOPEN , $AARRAY [ $I ] & @CRLF ) Then
				$IERROR = 3
				$IRETURN = 0
				ExitLoop
			EndIf
		Next
	Case 2
		Local $STEMP = ""
		For $I = $IBASE To $IUBOUND
			$STEMP = $AARRAY [ $I ] [ 0 ]
			For $J = 1 To UBound ( $AARRAY , $UBOUND_COLUMNS ) + 4294967295
				$STEMP &= $SDELIMITER & $AARRAY [ $I ] [ $J ]
			Next
			If Not FileWrite ( $HFILEOPEN , $STEMP & @CRLF ) Then
				$IERROR = 3
				$IRETURN = 0
				ExitLoop
			EndIf
		Next
	EndSwitch
	If IsString ( $SFILEPATH ) Then FileClose ( $HFILEOPEN )
	Return SetError ( $IERROR , 0 , $IRETURN )
EndFunc
Func _HEXTOSTRING ( $SHEX )
	If Not ( StringLeft ( $SHEX , 2 ) == "0x" ) Then $SHEX = "0x" & $SHEX
	Return BinaryToString ( $SHEX , $SB_UTF8 )
EndFunc
Global Const $SERVICES_ACTIVE_DATABASE = "ServicesActive"
Global Const $SC_MANAGER_CONNECT = 1
Global Const $SERVICE_QUERY_STATUS = 4
Global Const $SERVICE_START = 16
Global Const $SC_STATUS_PROCESS_INFO = 0
Func _SERVICE_QUERYSTATUS ( $SSERVICENAME , $SCOMPUTERNAME = "" )
	Local $HSC , $HSERVICE , $TSERVICE_STATUS_PROCESS , $AVQSSE , $IQSSE , $AISTATUS [ 9 ]
	$HSC = OPENSCMANAGER ( $SCOMPUTERNAME , $SC_MANAGER_CONNECT )
	$HSERVICE = OPENSERVICE ( $HSC , $SSERVICENAME , $SERVICE_QUERY_STATUS )
	$TSERVICE_STATUS_PROCESS = DllStructCreate ( "dword[9]" )
	Local $AVQSSE = DllCall ( "advapi32.dll" , "int" , "QueryServiceStatusEx" , "ptr" , $HSERVICE , "dword" , $SC_STATUS_PROCESS_INFO , "ptr" , DllStructGetPtr ( $TSERVICE_STATUS_PROCESS ) , "dword" , DllStructGetSize ( $TSERVICE_STATUS_PROCESS ) , "dword*" , 0 )
	If $AVQSSE [ 0 ] = 0 Then $IQSSE = _WINAPI_GETLASTERROR ( )
	CLOSESERVICEHANDLE ( $HSERVICE )
	CLOSESERVICEHANDLE ( $HSC )
	For $I = 0 To 8
		$AISTATUS [ $I ] = DllStructGetData ( $TSERVICE_STATUS_PROCESS , 1 , $I + 1 )
	Next
	Return SetError ( $IQSSE , 0 , $AISTATUS )
EndFunc
Func _SERVICE_START ( $SSERVICENAME , $SCOMPUTERNAME = "" )
	Local $HSC , $HSERVICE , $AVSS , $ISS
	$HSC = OPENSCMANAGER ( $SCOMPUTERNAME , $SC_MANAGER_CONNECT )
	$HSERVICE = OPENSERVICE ( $HSC , $SSERVICENAME , $SERVICE_START )
	$AVSS = DllCall ( "advapi32.dll" , "int" , "StartServiceW" , "ptr" , $HSERVICE , "dword" , 0 , "ptr" , 0 )
	If $AVSS [ 0 ] = 0 Then $ISS = _WINAPI_GETLASTERROR ( )
	CLOSESERVICEHANDLE ( $HSERVICE )
	CLOSESERVICEHANDLE ( $HSC )
	Return SetError ( $ISS , 0 , $AVSS [ 0 ] )
EndFunc
Func CLOSESERVICEHANDLE ( $HSCOBJECT )
	Local $AVCSH = DllCall ( "advapi32.dll" , "int" , "CloseServiceHandle" , "ptr" , $HSCOBJECT )
	If @error Then Return SetError ( @error , 0 , 0 )
	Return $AVCSH [ 0 ]
EndFunc
Func OPENSCMANAGER ( $SCOMPUTERNAME , $IACCESS )
	Local $AVOSCM = DllCall ( "advapi32.dll" , "ptr" , "OpenSCManagerW" , "wstr" , $SCOMPUTERNAME , "wstr" , $SERVICES_ACTIVE_DATABASE , "dword" , $IACCESS )
	If @error Then Return SetError ( @error , 0 , 0 )
	Return $AVOSCM [ 0 ]
EndFunc
Func OPENSERVICE ( $HSC , $SSERVICENAME , $IACCESS )
	Local $AVOS = DllCall ( "advapi32.dll" , "ptr" , "OpenServiceW" , "ptr" , $HSC , "wstr" , $SSERVICENAME , "dword" , $IACCESS )
	If @error Then Return SetError ( @error , 0 , 0 )
	Return $AVOS [ 0 ]
EndFunc
Global $VERSION = " Version: 13-05-2022"
Global $PROGRESS , $SVERSION , $SYSTEMDRIVE , $HMTB , $DEFAULTPATH , $C = EnvGet ( "systemdrive" )
$FORM1 = GUICreate ( "MiniToolBox by Farbar" & $VERSION , 420 , 420 , 192 , 124 )
$CHECKBOXA = GUICtrlCreateCheckbox ( "" , 24 , 20 , 15 , 17 )
$LABEL1 = GUICtrlCreateLabel ( "Select All" , 40 , 22 , 97 , 17 )
GUICtrlSetColor ( $LABEL1 , 16127 )
$CHECKBOX1 = GUICtrlCreateCheckbox ( "Flush DNS" , 24 , 50 , 97 , 17 )
$CHECKBOX2 = GUICtrlCreateCheckbox ( "Report IE Proxy Settings" , 24 , 70 , 200 , 17 )
$CHECKBOX3 = GUICtrlCreateCheckbox ( "Reset IE Proxy Settings" , 24 , 90 , 200 , 17 )
$CHECKBOX4 = GUICtrlCreateCheckbox ( "Report FF Proxy Settings" , 24 , 110 , 200 , 17 )
$CHECKBOX5 = GUICtrlCreateCheckbox ( "Reset FF Proxy Settings" , 24 , 130 , 200 , 17 )
$CHECKBOX6 = GUICtrlCreateCheckbox ( "List content of Hosts" , 24 , 150 , 200 , 17 )
$CHECKBOX7 = GUICtrlCreateCheckbox ( "List IP Configuration" , 24 , 170 , 200 , 17 )
$CHECKBOX8 = GUICtrlCreateCheckbox ( "List Winsock Entries" , 24 , 190 , 200 , 17 )
$CHECKBOX9 = GUICtrlCreateCheckbox ( "List last 10 Event Viewer Errors" , 24 , 210 , 200 , 17 )
$CHECKBOX10 = GUICtrlCreateCheckbox ( "List Installed Programs" , 24 , 230 , 200 , 17 )
$CHECKBOX11 = GUICtrlCreateCheckbox ( "List Devices" , 24 , 250 , 95 , 17 )
$RADIO1 = GUICtrlCreateRadio ( "Only Problems" , 125 , 253 , 108 , 15 )
GUICtrlSetState ( + 4294967295 , $GUI_CHECKED )
$RADIO2 = GUICtrlCreateRadio ( "No Driver" , 243 , 253 , 76 , 15 )
$RADIO3 = GUICtrlCreateRadio ( "All" , 325 , 253 , 50 , 15 )
$CHECKBOX12 = GUICtrlCreateCheckbox ( "List Users, Partitions and Memory size" , 24 , 270 , 250 , 17 )
$CHECKBOX13 = GUICtrlCreateCheckbox ( "List Minidump Files" , 24 , 290 , 250 , 17 )
$CHECKBOX14 = GUICtrlCreateCheckbox ( "List Restore Points" , 24 , 310 , 250 , 17 )
$BUTTON1 = GUICtrlCreateButton ( "GO" , 110 , 340 , 99 , 25 , $WS_BORDER , $WS_EX_DLGMODALFRAME )
GUICtrlSetBkColor ( $BUTTON1 , 12178414 )
GUICtrlSetColor ( $BUTTON1 , 2240876 )
$LABEL = GUICtrlCreateLabel ( "" , 24 , 370 , 310 , 20 )
GUISetState ( @SW_SHOW )
If Not FileExists ( @ScriptDir & "\MTB.txt" ) Then
	$YN = MsgBox ( 4 + 64 , "MiniToolBox by Farbar" , "This software is not permitted for commercial purposes." & @CRLF & @CRLF & "Are you sure you want to continue?" & @CRLF & @CRLF & "Click Yes to continue. Click No to exit." )
	If $YN = 7 Then Exit
EndIf
$SYSTEMDRIVE = StringRegExpReplace ( @WindowsDir , "(?i)([A-Z]:).+" , "$1" )
While 1
	If GUICtrlRead ( $CHECKBOXA ) = $GUI_CHECKED Then
		GUICtrlSetState ( $CHECKBOX1 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX2 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX3 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX4 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX5 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX6 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX7 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX8 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX9 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX10 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX11 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX12 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX13 , $GUI_CHECKED )
		GUICtrlSetState ( $CHECKBOX14 , $GUI_CHECKED )
	EndIf
	$GUIMSG = GUIGetMsg ( )
	Select
	Case $GUIMSG = $GUI_EVENT_CLOSE
		Exit
	Case $GUIMSG = $BUTTON1
		If GUICtrlRead ( $CHECKBOX1 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX2 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX3 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX4 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX5 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX6 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX7 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX8 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX9 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX10 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX11 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX12 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX13 ) = $GUI_CHECKED Or GUICtrlRead ( $CHECKBOX14 ) = $GUI_CHECKED Then
			$PROGRESS = GUICtrlCreateProgress ( 18 , 392 , 280 , 18 , 8 , 8192 )
			GUICtrlSendMsg ( $PROGRESS , 1034 , 1 , 50 )
			OS ( )
			If $SVERSION = "" Then
				$SVERSION = RegRead ( "HKLM64\software\Microsoft\Windows NT\CurrentVersion" , "ProductName" )
				If $SVERSION = "" Then
					$SVERSION = @OSVersion
					If $SVERSION = "WIN_XP" Then $SVERSION = "Windows XP"
					If $SVERSION = "WIN_VISTA" Then $SVERSION = "Windows Vista"
					If $SVERSION = "WIN_7" Then $SVERSION = "Windows 7"
					If $SVERSION = "WIN_8" Then $SVERSION = "Windows 8"
					If $SVERSION = "WIN_2008" Then $SVERSION = "Windows 2008"
					If $SVERSION = "WIN_10" Then $SVERSION = "Windows 10"
					If $SVERSION = "WIN_11" Then $SVERSION = "Windows 11"
				EndIf
			EndIf
			Local $BOOT , $VAL , $CDATE , $ADMIN
			$VAL = RegRead ( "HKLM64\system\currentcontrolset\control\safeboot\option" , "OptionValue" )
			Select
			Case @error = 1
				$BOOT = "Normal"
			Case @error = 0
				If $VAL = 1 Then $BOOT = "Minimal"
				If $VAL = 2 Then $BOOT = "Network"
			EndSelect
			$CDATE = @MDAY & "-" & @MON & "-" & @YEAR & " at " & @HOUR & ":" & @MIN & ":" & @SEC
			If IsAdmin ( ) Then
				$ADMIN = " (administrator)"
			Else
				$ADMIN = " (ATTENTION: The logged in user is not administrator)"
			EndIf
			$MANUFACTURER = ""
			$MODEL = ""
			MAKEUPMODEL ( $MANUFACTURER , $MODEL )
			$MANUFACTURER = StringRegExpReplace ( $MANUFACTURER , "\s+$" , "" )
			$MODEL = StringRegExpReplace ( $MODEL , "\s+$" , "" )
			$HMTB = FileOpen ( "MTB.txt" , 256 + 2 )
			FileWrite ( $HMTB , "MiniToolBox by Farbar " & $VERSION & @CRLF & "Ran by " & @UserName & $ADMIN & " on " & $CDATE & @CRLF & "Running from """ & @ScriptDir & """" & @CRLF & $SVERSION & " " & @OSServicePack & " (" & @OSArch & ")" & @CRLF & "Model: " & $MODEL & " Manufacturer: " & $MANUFACTURER & @CRLF & "Boot Mode: " & $BOOT & @CRLF & "***************************************************************************" & @CRLF )
			If GUICtrlRead ( $CHECKBOX1 ) = $GUI_CHECKED Then FLUSHDNS ( )
			If GUICtrlRead ( $CHECKBOX2 ) = $GUI_CHECKED Then SHOWPROXY ( )
			If GUICtrlRead ( $CHECKBOX3 ) = $GUI_CHECKED Then DELPOXY ( )
			If GUICtrlRead ( $CHECKBOX4 ) = $GUI_CHECKED Then PROXYLIST ( )
			If GUICtrlRead ( $CHECKBOX5 ) = $GUI_CHECKED Then PROXYRESETFF ( )
			If GUICtrlRead ( $CHECKBOX6 ) = $GUI_CHECKED Then HOSTS ( )
			If GUICtrlRead ( $CHECKBOX7 ) = $GUI_CHECKED Then IPCONFIG ( )
			If GUICtrlRead ( $CHECKBOX8 ) = $GUI_CHECKED Then WINSOCK ( )
			If GUICtrlRead ( $CHECKBOX9 ) = $GUI_CHECKED Then EVENTS1 ( )
			If GUICtrlRead ( $CHECKBOX10 ) = $GUI_CHECKED Then PROGRAMS ( )
			If GUICtrlRead ( $CHECKBOX11 ) = $GUI_CHECKED Then
				If _SRVSTAT ( "winmgmt" ) = "R" Then
					DEVICES ( )
				Else
					FileWrite ( $HMTB , @CRLF & "=========================" & @CRLF & "Windows Management Instrumentation service is not running. Could not scan devices" & @CRLF & "=========================" & @CRLF & @CRLF )
				EndIf
			EndIf
			If GUICtrlRead ( $CHECKBOX12 ) = $GUI_CHECKED Then PART ( )
			If GUICtrlRead ( $CHECKBOX13 ) = $GUI_CHECKED Then MINIDUMP ( )
			If GUICtrlRead ( $CHECKBOX14 ) = $GUI_CHECKED And _SRVSTAT ( "winmgmt" ) = "R" Then RESTOREPOINTS ( )
			FileWrite ( $HMTB , @CRLF & "**** End of log ****" & @CRLF )
			FileClose ( $HMTB )
			Run ( "notepad.exe MTB.txt" )
			GUICtrlDelete ( $PROGRESS )
			Exit
		Else
			MsgBox ( 262144 , "MiniToolBox" , "No option is selected." & @CRLF & @CRLF & "Please check the option(s) then press ""Go"" button" )
		EndIf
	EndSelect
WEnd
Func _AAAAP1 ( )
	Local $ARRPACK [ 1 ]
	$KEY = "HKCR\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages"
	$E = 0
	While 1
		$E += 1
		$SUB = RegEnumKey ( $KEY , $E )
		If @error Then ExitLoop
		$PRF = RegRead ( $KEY & "\" & $SUB , "PackageRootFolder" )
		If Not FileExists ( $PRF ) Then ContinueLoop
		$READ = FileRead ( $PRF & "\" & "AppxManifest.xml" )
		$DEP = ""
		If StringRegExp ( $READ , "(?i)PackageDependency Name=""Microsoft.Advertising" ) Or FileExists ( $PRF & "\Microsoft.Advertising" ) Or FileExists ( $PRF & "\MSAdvertisingJS" ) Then $DEP = " [MS Ad]"
		If StringRegExp ( $READ , "(?i)(desktop|uap5):StartupTask" ) Then
			If StringRegExp ( $SUB , "(?i)^Microsoft\.?(Teams|WindowsTerminal|YourPhone|549981C3F5F10|GamingApp)_(\d+\.)+\d+(_x\d+)*(_neutral)*_+8wekyb3d8bbwe" ) Then ContinueLoop
			If StringRegExp ( $SUB , "(?i)^Microsoft.SkypeApp_(\d+\.)+\d+(_x\d+)*__kzf8qxf38zg5c" ) Then ContinueLoop
			$DEP &= " [Startup Task]"
		EndIf
		If Not $DEP Then
			If StringRegExp ( $SUB , "(?i)^((WhatsNew|WebAuthBridgeIntranetSso|WebAuthBridgeInternetSso|WebAuthBridgeInternet|RoomAdjustment|Passthrough|MixedRealityLearning|EnvironmentsApp|DesktopLearning|HoloItemPlayerApp|HoloCamera|DesktopView|HoloShell|CortanaListenUIApp|winstore|SonicWALL.MobileConnect|JuniperNetworks.JunosPulseVpn|FileManager|f5.vpn.client|CheckPoint.VPN|InputApp|1527c705-839a-4832-9118-54d4Bd6a0c89|c5e2524a-ea46-4f67-841f-6a9465d9d515|E2A4F912-2574-4A75-9BB0-0D023378592B|F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE|Microsoft\.?(Windows.Client.(CBS|WebExperience)|Windows.UndockedDevKit|Windows.Search|Windows.DevicesFlowHost|EdgeDevtoolsPlugin|XboxIdentityProvider|WindowsFeedback|MoCamera|PPIProjection|Win32WebViewHost|Windows.(StartMenuExperienceHost|CallingShellApp|WindowPicker|ModalSharePickerHost|HolographicFirstRun|SecondaryTileExperience|SecureAssessmentBrowser|ContentDeliveryManager|PeopleExperienceHost|XGpuEjectDialog|ParentalControls|ShellExperienceHost|AssignedAccessLockApp|CloudExperienceHost|Apprep.ChxApp|CapturePicker|Cortana)|AccountsControl|BioEnrollment|CredDialogHost|LockApp|Windows.OOBENetwork(CaptivePortal|ConnectionFlow)|AAD.BrokerPlugin|XboxGameCallableUI|Windows.PinningConfirmationDialog|Windows.SecHealthUI)|Windows.(PurchaseDialog|devicesflow|CBSPreview|immersivecontrolpanel|PrintDialog|MiracastView|ContactSupport))_(\d+\.)+\d+(_x\d+)*(_neutral)*_+cw5n1h2txyewy)$" ) Then ContinueLoop
			If StringRegExp ( $SUB , "(?i)^(NcsiUwpApp|Microsoft.(Todos|SecHealthUI|PowerAutomateDesktop|OneDriveSync|GamingApp|WindowsNotepad|UI.Xaml.CBS|Paint|YourPhone|ZuneVideo|DirectXRuntime|GamingServices|MicrosoftEdge.Stable|BingNews|WindowsMixedRealityRuntimeApp|WindowsMixedReality.Runtime|360Viewer|MicrosoftSolitaireCollection|BingWeather|BingSports|BingFinance|XboxCompanion|TranslatorforMicrosoftEdge|XboxLIVEGames|Reader|Media.PlayReadyClient(.\d)*|HelpAndTips|BingFoodAndDrink|BingHealthAndFitness|BingTravel|BingMaps|LanguageExperiencePack\w+-\w+|ZuneMusic|XboxSpeechToTextOverlay|XboxIdentityProvider|XboxGamingOverlay|XboxGameOverlay|XboxApp|Xbox.TCUI|3DBuilder|MicrosoftEdgeDevToolsClient|Windows.NarratorQuickStart|Appconnector|AsyncTextService|HEIFImageExtension|HEVCVideoExtension|Messaging|MicrosoftEdge|MicrosoftOfficeHub|MicrosoftStickyNotes|ConnectivityStore|DesktopAppInstaller|ECApp|GetHelp|Getstarted|Microsoft3DViewer|MixedReality.Portal|MSPaint|NET.Native.Framework.\d.\d|NET.Native.Runtime.\d.\d|Office.OneNote|Office.Sway|OneConnect|People|Print3D|ScreenSketch|Services.Store.Engagement|StorePurchaseApp|UI.Xaml.\d.\d|VCLibs(.\d+)+.(UWPDesktop|Preview|Universal)*|WinJS.\d.\d|VP9VideoExtensions|Wallet|WebMediaExtensions|WebpImageExtension|Windows(communicationsapps|ReadingList|Scan|.Photos|.SecHealthUI|Alarms|Calculator|Camera|FeedbackHub|Maps|SoundRecorder|Store)))_(\d+\.)+\d+(_x\d+)*(_neutral)*_+8wekyb3d8bbwe$" ) Then ContinueLoop
		EndIf
		$DATECR = FILETIMECM ( $PRF )
		If Not $DATECR Then $DATECR = FILETIMECM ( $PRF & "\AppxManifest.xml" )
		$DN = RegRead ( $KEY & "\" & $SUB , "DisplayName" )
		$NAME = $DN
		$PUB = ""
		If $DN And StringRegExp ( $DN , "^@" ) Then
			$RETNAME = PACKAGES1 ( $DN )
			If $RETNAME Then $NAME = $RETNAME
			If StringRegExp ( $DN , "(?i)AppDisplayName|AppStoreName|StoreAppName|PackageDisplayName|PkgDisplayName|DisplayName|ConnectorStubTitle|ApplicationTitleWithBranding|AppName|ProductName|DisplayTitle|appDescription|StoreTitle|GameBar|IDS_MANIFEST_(MUSIC|VIDEO)_APP_NAME" ) Then
				$RETPUB = PACKAGES1 ( $DN , 1 )
				If $RETPUB Then $PUB = " (" & $RETPUB & ")"
			EndIf
		EndIf
		If Not $PUB Then $PUB = " (" & PACKAGES2 ( $READ ) & ")"
		_ARRAYADD ( $ARRPACK , $NAME & " -> " & $PRF & " [" & $DATECR & "]" & $PUB & $DEP , 0 , "|||" )
	WEnd
	$ARRPACK = _ARRAYUNIQUE ( $ARRPACK , 0 , 0 , 0 , 0 , 1 )
	_ARRAYSORT ( $ARRPACK )
	If UBound ( $ARRPACK ) > 1 Then
		FileWrite ( $HMTB , @CRLF & "Packages:" & @CRLF & "=========" & @CRLF )
		_FILEWRITEFROMARRAY ( $HMTB , $ARRPACK , 1 )
	EndIf
EndFunc
Func _RUN ( $COM )
	$PID = Run ( @ComSpec & " /u /c " & $COM , "" , @SW_HIDE , 2 + 4 )
	ProcessWaitClose ( $PID )
	$READ1 = StdoutRead ( $PID , False , True )
	$READ2 = StderrRead ( $PID )
	$PATH1 = @TempDir & "\cmd1" & Random ( 1000 , 9999 , 1 ) & ".txt"
	$PATH2 = @TempDir & "\cmd2" & Random ( 1000 , 9999 , 1 ) & ".txt"
	$HLOG1 = FileOpen ( $PATH1 , 256 + 2 )
	$HLOG2 = FileOpen ( $PATH2 , 256 + 2 )
	FileWrite ( $HLOG1 , $READ1 )
	FileWrite ( $HLOG2 , $READ2 )
	FileClose ( $HLOG1 )
	FileClose ( $HLOG2 )
	$READ1 = FileRead ( $PATH1 )
	$READ2 = FileRead ( $PATH2 )
	FileDelete ( $PATH1 )
	FileDelete ( $PATH2 )
	Return $READ1 & @CRLF & $READ2
EndFunc
Func _SRVSTAT ( $SNAME )
	$ST = _SERVICE_QUERYSTATUS ( $SNAME )
	If IsArray ( $ST ) Then
		Switch $ST [ 1 ]
		Case 1
			Return "S"
		Case 4
			Return "R"
	Case Else
			Return "U"
		EndSwitch
	EndIf
EndFunc
Func CODEINTEGRITY ( )
	$EVENT = _RUN ( "wevtutil qe ""Microsoft-Windows-CodeIntegrity/Operational"" ""/q:*[System [(Level=2)]]"" /c:12 /rd:true /uni:true /f:text""" )
	If Not StringInStr ( $EVENT , ":" ) Then Return
	$EVENT = StringRegExpReplace ( $EVENT , "(?m)^\s*" , "" )
	$EVENT = StringRegExpReplace ( $EVENT , "(?m)^(Event\[\d\]|Log Name|Event ID|Task|Level|Opcode|Keyword|Source|User|User Name|Computer):?.*\v{2}" , "" )
	$EVENT = StringRegExpReplace ( $EVENT , "(?m)^(Date:.+\d)T(\d.+\v{2})" , @CRLF & "$1 $2" )
	$EVENT = StringRegExpReplace ( $EVENT , "(?m)^(Date:.+?)\.\d+Z" , "$1" )
	$EVENT = StringRegExpReplace ( $EVENT , "(?s)Date:[\s\d:-]+\RDescription:\s*\RN/A\R" , "" )
	$ARR1 = StringRegExp ( $EVENT , "(?s)Date:[\s\d:-]+\RDescription:\s*.+?\R" , 3 )
	If UBound ( $ARR1 ) < 1 Then Return
	$HCODE = FileOpen ( @TempDir & "\codeint2" , 256 + 2 )
	For $I = 0 To UBound ( $ARR1 ) + 4294967295
		$HCODE = FileOpen ( @TempDir & "\codeint2" , 256 )
		$READ = FileRead ( $HCODE )
		FileClose ( $HCODE )
		$RET = StringRegExpReplace ( $ARR1 [ $I ] , "(?s).+Description:\s*(.+)" , "$1" )
		If Not StringInStr ( $READ , $RET ) Then
			$HCODE = FileOpen ( @TempDir & "\codeint2" , 256 + 1 )
			FileWrite ( $HCODE , $ARR1 [ $I ] & @CRLF )
			FileClose ( $HCODE )
		EndIf
	Next
	$HCODE = FileOpen ( @TempDir & "\codeint2" , 256 )
	$READ = FileRead ( $HCODE )
	FileWrite ( $HMTB , @CRLF & "CodeIntegrity Errors:" & @CRLF & "====================" & @CRLF & $READ )
EndFunc
Func DELPOXY ( )
	Local $R
	GUICtrlSetData ( $LABEL , "Resetting IE Proxy setting" )
	RegDelete ( "HKCU64\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyServer" )
	RegDelete ( "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyServer" )
	RegDelete ( "HKCU64\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyEnable" )
	RegDelete ( "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyEnable" )
	DllCall ( "WININET.DLL" , "long" , "InternetSetOption" , "int" , 0 , "long" , 39 , "str" , 0 , "long" , 0 )
	RunWait ( @ComSpec & " /c proxycfg -d" , "" , @SW_HIDE )
	$R = RegDelete ( "HKCU64\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyEnable" )
	If Not @error Then FileWrite ( $HMTB , @CRLF & """Reset IE Proxy Settings"": IE Proxy Settings were reset." & @CRLF )
EndFunc
Func DEVICES ( )
	Local $OBJWMISERVICE , $DEVCOLITEMS
	GUICtrlSetData ( $LABEL , "Getting Devices..." )
	FileWrite ( $HMTB , @CRLF & "========================= Devices: ================================" & @CRLF & @CRLF )
	$OMYERROR = ""
	$OMYERROR = ObjEvent ( "AutoIt.Error" , "MyErrFunc" )
	$OBJWMISERVICE = ObjGet ( "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2" )
	If Not @error = 0 Then
		FileWrite ( $HMTB , "Could not list devices." & @CRLF )
	Else
		$DEVCOLITEMS = $OBJWMISERVICE .ExecQuery ( "Select * from Win32_PnPEntity" )
		For $OBJECT In $DEVCOLITEMS
			Local $NAME = $OBJECT .Name
			Local $DESCRIP = $OBJECT .description
			Local $CLASSGUID = $OBJECT .classguid
			Local $COMPANY = $OBJECT .Manufacturer
			Local $SERVICE = $OBJECT .service
			Local $DEVICEID = $OBJECT .DeviceID
			Local $CODE , $CODEMSG
			If $OBJECT .ConfigManagerErrorCode = 0 Then
				$CODE = ""
				$CODEMSG = ""
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 1 Then
				$CODE = ": This device is not configured correctly. (Code1)"
				$CODEMSG = "You may be prompted to provide the path of the driver. Windows may have the driver built-in, or may still have the driver files installed from the last time that you set up the device. If you are asked for the driver and you do not have it, you can try to download the latest driver from the hardware vendor’s Web site." & @CRLF & "In the device properties dialog box, click the ""Driver"" tab, and then click ""Update Driver"" to start the ""Hardware Update Wizard"". Follow the instructions to update the driver. If updating the driver does not work, see your hardware documentation for more information."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 3 Then
				$CODE = ": The driver for this device might be corrupted, or your system may be running low on memory or other resources. (Code3)"
				$CODEMSG = "If the driver is corrupted, uninstall the driver and scan for new hardware to install the driver again. To scan for new hardware, click on the ""Action"" menu in Device Manager, and then select ""Scan for hardware changes""." & @CRLF & "If your computer does not have enough memory to run the device, you can close some applications to make memory available. To check memory and system resources, right-click ""My Computer"", click ""Properties"", click the ""Advanced"" tab, and then click ""Settings"" under ""Performance""." & @CRLF & "You may need to install additional random access memory (RAM)." & @CRLF & "On the ""General Properties"" tab of the device, click ""Troubleshoot"" to start the troubleshooting wizard."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 10 Then
				$CODE = ": This device cannot start. (Code10)"
				$CODEMSG = "Device failed to start. Click ""Update Driver"" to update the drivers for this device." & @CRLF & "On the ""General Properties"" tab of the device, click ""Troubleshoot"" to start the troubleshooting wizard."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 12 Then
				$CODE = ": This device cannot find enough free resources that it can use. If you want to use this device, you will need to disable one of the other devices on this system. (Code12)"
				$CODEMSG = "Two devices have been assigned the same input/output (I/O) ports, the same interrupt, or the same Direct Memory Access channel (either by the BIOS, the operating system, or a combination of the two). This error message can also appear if the BIOS did not allocate enough resources to the device (for example, if a universal serial bus (USB) controller does not get an interrupt from the BIOS because of a corrupt Multiprocessor System (MPS) table)." & @CRLF & "You can use Device Manager to determine where the conflict is and disable the conflicting device." & @CRLF & "On the ""General Properties"" tab of the device, click ""Troubleshoot"" to start the troubleshooting wizard."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 14 Then
				$CODE = ": This device cannot work properly until you restart your computer. (Code14)"
				$CODEMSG = "Restart your computer."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 16 Then
				$CODE = ": Windows cannot identify all the resources this device uses. (Code 16)"
				$CODEMSG = "The device is only partially configured." & @CRLF & "To specify additional resources for this device, click the ""Resources"" tab in Device Manager. If there is a resource with a question mark next to it in the list of resources assigned to the device, select that resource to assign it to the device. If the resource cannot be changed, click ""Change Settings"". If ""Change Settings"" is unavailable, try clearing the ""Use automatic settings"" check box to make it available. If this is not a Plug and Play device, check the hardware documentation for more information." & @CRLF & "On the ""General Properties"" tab of the device, click ""Troubleshoot"" to start the troubleshooting wizard."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 18 Then
				$CODE = ": Reinstall the drivers for this device. (Code 18)"
				$CODEMSG = "The drivers for this device must be reinstalled." & @CRLF & " Click ""Update Driver"", which starts the Hardware Update wizard." & @CRLF & "Alternately, uninstall the driver, and then click ""Scan for hardware changes"" to reload the drivers."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 19 Then
				$CODE = ": Windows cannot start this hardware device because its configuration information (in the registry) is incomplete or damaged. (Code 19)"
				$CODEMSG = "A registry problem was detected." & @CRLF & " This can occur when more than one service is defined for a device, if there is a failure opening the service subkey, or if the driver name cannot be obtained from the service subkey. Try these options:" & @CRLF & "On the ""General Properties"" tab of the device, click ""Troubleshoot"" to start the troubleshooting wizard." & @CRLF & "Click ""Uninstall"", and then click ""Scan for hardware changes"" to load a usable driver."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 21 Then
				$CODE = ": Windows is removing this device. (Code 21)"
				$CODEMSG = "Wait several seconds, and then press the F5 key to update the Device Manager view." & @CRLF & "If that does not resolve the problem, restart your computer. "
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 22 Then
				$CODE = ": This device is disabled. (Code 22)"
				$CODEMSG = "In Device Manager, click ""Action"", and then click ""Enable Device"". This starts the Enable Device wizard. Follow the instructions."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 24 Then
				$CODE = ": This device is not present, is not working properly, or does not have all its drivers installed. (Code 24)"
				$CODEMSG = "The device is installed incorrectly. The problem could be a hardware failure, or a new driver might be needed." & @CRLF & "Devices stay in this state if they have been prepared for removal." & @CRLF & "After you remove the device, this error disappears.Remove the device, and this error should be resolved."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 28 Then
				$CODE = ": The drivers for this device are not installed. (Code 28)"
				$CODEMSG = "To install the drivers for this device, click ""Update Driver"", which starts the Hardware Update wizard."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 29 Then
				$CODE = ": This device is disabled because the firmware of the device did not give it the required resources. (Code 29)"
				$CODEMSG = "Enable the device in the BIOS of the device."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 31 Then
				$CODE = ": This device is not working properly because Windows cannot load the drivers required for this device. (Code 31)"
				$CODEMSG = "Update the driver"
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 32 Then
				$CODE = ": A driver (service) for this device has been disabled. An alternate driver may be providing this functionality (Code 32)"
				$CODEMSG = "The start type for this driver is set to disabled in the registry." & @CRLF & "Uninstall the driver from Device Manager, and then scan for new hardware to install the driver again. If this does not work, you might have to change the device start type parameter in the registry."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 33 Then
				$CODE = ": Windows cannot determine which resources are required for this device. (Code 33)"
				$CODEMSG = "The translator that determines the kinds of resources that are required by the device has failed." & @CRLF & "Configure the hardware. If configuring the hardware does not work, you might have to replace it."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 34 Then
				$CODE = ": Windows cannot determine the settings for this device. Consult the documentation that came with this device and use the Resource tab to set the configuration. (Code 34)"
				$CODEMSG = "The device requires manual configuration. See the hardware documentation or contact the hardware vendor for instructions on manually configuring the device. After you configure the device itself, you can use the ""Resources"" tab in Device Manager to configure the resource settings in Windows."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 35 Then
				$CODE = ": Your computer's system firmware does not include enough information to properly configure and use this device. To use this device, contact your computer manufacturer to obtain a firmware or BIOS update. (Code 35)"
				$CODEMSG = "The Multiprocessor System (MPS) table, which stores the resource assignments for the BIOS, is missing an entry for your device and needs to be updated." & @CRLF & "Obtain a new BIOS from the system vendor."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 36 Then
				$CODE = ": This device is requesting a PCI interrupt but is configured for an ISA interrupt (or vice versa). Please use the computer's system setup program to reconfigure the interrupt for this device. (Code 36)"
				$CODEMSG = "The interrupt request (IRQ) translation failed. Change the settings for the IRQ reservations." & @CRLF & "This content is designed for an advanced computer user." & @CRLF & "Change the settings for IRQ reservations." & @CRLF & "For more information about how to change BIOS settings, see the hardware documentation." & @CRLF & "You can also try to use the BIOS setup tool to change the settings for IRQ reservations (if such options exist). The BIOS might have options to reserve certain IRQs for peripheral component interconnect (PCI) or ISA devices. "
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 37 Then
				$CODE = ": Windows cannot initialize the device driver for this hardware. (Code 37)"
				$CODEMSG = "The driver returned failure from its DriverEntry routine. Uninstall the driver, and then click ""Scan for hardware changes"" to reinstall or upgrade the driver."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 38 Then
				$CODE = ": Windows cannot load the device driver for this hardware because a previous instance of the device driver is still in memory. (Code 38)"
				$CODEMSG = "The driver could not be loaded because a previous instance is still loaded." & @CRLF & "Restart the computer."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 39 Then
				$CODE = ": Windows cannot load the device driver for this hardware. The driver may be corrupted or missing. (Code 39)"
				$CODEMSG = "Reasons for this error include a driver that is not present; a binary file that is corrupt; a file I/O problem, or a driver that references an entry point in another binary file that could not be loaded." & @CRLF & "Uninstall the driver, and then click ""Scan for hardware changes"" to reinstall or upgrade the driver."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 40 Then
				$CODE = ": Windows cannot access this hardware because its service key information in the registry is missing or recorded incorrectly. (Code 40)"
				$CODEMSG = "Information in the registry's service subkey for the driver is invalid. Uninstall the driver, and then click Scan for hardware changes to load the driver again. "
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 41 Then
				$CODE = ": Windows successfully loaded the device driver for this hardware but cannot find the hardware device. (Code 41)"
				$CODEMSG = "A driver was loaded but Windows cannot find the device. This happens when Windows does not detect a non-Plug and Play device." & @CRLF & "If the device was removed, uninstall the driver, install the device, and then click ""Scan for hardware changes"" to reinstall the driver. If the hardware was not removed, obtain a new or updated driver for the device." & @CRLF & "If the device is a non-Plug and Play device, a newer version of the driver might be needed. To install non-Plug and Play devices, use the Add Hardware wizard." & @CRLF & "Click ""Performance and Maintenance"" on ""Control Panel"", click ""System"", and on the ""Hardware"" tab, click ""Add Hardware Wizard""."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 42 Then
				$CODE = ": Bus driver has created two devices with the same names. (Code 42)"
				$CODEMSG = "Restart the computer."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 43 Then
				$CODE = ": Windows has stopped this device because it has reported problems. (Code 43)"
				$CODEMSG = "One of the drivers controlling the device notified the operating system that the device failed in some manner. For more information about how to diagnose the problem, see the hardware documentation. "
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 44 Then
				$CODE = ": An application or service has shut down this hardware device. (Code 44)"
				$CODEMSG = "Restart the computer."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 45 Then
				$CODE = ": Currently, this hardware device is not connected to the computer. (Code 45)."
				$CODEMSG = "The device is not present or was previously attached to the computer." & @CRLF & "To fix this problem, reconnect this hardware device to the computer." & @CRLF & "If Device Manager is started with the environment variable DEVMGR_SHOW_NONPRESENT_DEVICES set to 1 (which means show these devices), then any previously attached (NONPRESENT) devices are displayed in the device list and assigned this error code."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 46 Then
				$CODE = ": Windows cannot gain access to this hardware device because the operating system is in the process of shutting down (Code 46)."
				$CODEMSG = "The device is not available because the system is shutting down. When computer restarts the device should function as normal."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 47 Then
				$CODE = ": Windows cannot use this hardware device because it has been prepared for safe removal, but it has not been removed from the computer. (Code 47)"
				$CODEMSG = "Unplug the device, and then plug it in again. Alternately, restart the computer to make the device available."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 48 Then
				$CODE = ": The software for this device has been blocked from starting because it is known to have problems with Windows. Contact the hardware vendor for a new driver. (Code 48)"
				$CODEMSG = "Download the latest drivers from the manufacturer, uninstall the current driver, and then install the latest drivers."
			EndIf
			If $OBJECT .ConfigManagerErrorCode = 49 Then
				$CODE = ": Windows cannot start new hardware devices because the system hive is too large (exceeds the Registry Size Limit). (Code 49) "
				$CODEMSG = "Uninstall any un-used hardware devices."
			EndIf
			If GUICtrlRead ( $RADIO1 ) = $GUI_CHECKED And $CODE <> "" Then FileWrite ( $HMTB , "Name: " & $NAME & @CRLF & "Description: " & $DESCRIP & @CRLF & "Class Guid: " & $CLASSGUID & @CRLF & "Manufacturer: " & $COMPANY & @CRLF & "Service: " & $SERVICE & @CRLF & "Device ID: " & $DEVICEID & @CRLF & "Problem: " & $CODE & @CRLF & "Resolution: " & $CODEMSG & @CRLF & @CRLF )
			If GUICtrlRead ( $RADIO2 ) = $GUI_CHECKED And $COMPANY <> "" Then
				If $CODE = "" Then FileWrite ( $HMTB , "Name: " & $NAME & @CRLF & "Description: " & $DESCRIP & @CRLF & "Class Guid: " & $CLASSGUID & @CRLF & "Manufacturer: " & $COMPANY & @CRLF & "Service: " & $SERVICE & @CRLF & "Device ID: " & $DEVICEID & @CRLF & @CRLF )
				If $CODE <> "" Then FileWrite ( $HMTB , "Name: " & $NAME & @CRLF & "Description: " & $DESCRIP & @CRLF & "Class Guid: " & $CLASSGUID & @CRLF & "Manufacturer: " & $COMPANY & @CRLF & "Service: " & $SERVICE & @CRLF & "Device ID: " & $DEVICEID & @CRLF & "Problem: " & $CODE & @CRLF & "Resolution: " & $CODEMSG & @CRLF & @CRLF )
			EndIf
			If GUICtrlRead ( $RADIO3 ) = $GUI_CHECKED Then
				If $CODE = "" Then FileWrite ( $HMTB , "Name: " & $NAME & @CRLF & "Description: " & $DESCRIP & @CRLF & "Class Guid: " & $CLASSGUID & @CRLF & "Manufacturer: " & $COMPANY & @CRLF & "Service: " & $SERVICE & @CRLF & "Device ID: " & $DEVICEID & @CRLF & @CRLF )
				If $CODE <> "" Then FileWrite ( $HMTB , "Name: " & $NAME & @CRLF & "Description: " & $DESCRIP & @CRLF & "Class Guid: " & $CLASSGUID & @CRLF & "Manufacturer: " & $COMPANY & @CRLF & "Service: " & $SERVICE & @CRLF & "Device ID: " & $DEVICEID & @CRLF & "Problem: " & $CODE & @CRLF & "Resolution: " & $CODEMSG & @CRLF & @CRLF )
			EndIf
		Next
	EndIf
EndFunc
Func EVENTS ( )
	GUICtrlSetData ( $LABEL , "Getting events, please wait..." )
	FileDelete ( @TempDir & "\event" )
	MAINAPP ( )
	MAINSYS ( )
	WINDEF ( )
	If Not StringInStr ( $SVERSION , "XP" ) Then CODEINTEGRITY ( )
	GUICtrlSetData ( $LABEL , "" )
EndFunc
Func EVENTS1 ( )
	If _SRVSTAT ( "eventlog" ) = "R" Then
		EVENTS ( )
	Else
		GUICtrlSetData ( $LABEL , "Attempting to start eventlog, please wait..." )
		RunWait ( @ComSpec & " /c " & "sc config eventlog start= auto" , "" , @SW_HIDE )
		_SERVICE_START ( "eventlog" )
		Sleep ( 5000 )
		If _SRVSTAT ( "eventlog" ) = "R" Then
			EVENTS ( )
		Else
			FileWrite ( $HMTB , @CRLF & "========================= Event log errors: ================================" & @CRLF & @CRLF & "Could" & " not start eventlog service, could not read events." & @CRLF & @CRLF )
			RunWait ( @ComSpec & " /c " & "net start eventlog >> MTB.txt 2>&1" , "" , @SW_HIDE )
		EndIf
	EndIf
EndFunc
Func FILETIMECM ( $PATH1 , $1 = 0 )
	If Not FileExists ( $PATH1 ) Then Return ""
	If $1 Then
		$T = FileGetTime ( $PATH1 , 1 )
	Else
		$T = FileGetTime ( $PATH1 )
	EndIf
	If IsArray ( $T ) Then
		$DATEMO = $T [ 0 ] & "-" & $T [ 1 ] & "-" & $T [ 2 ]
	Else
		$DATEMO = "0000-00-00"
	EndIf
	Return $DATEMO
EndFunc
Func FLUSHDNS ( )
	GUICtrlSetData ( $LABEL , "Flushing DNS" )
	FileWrite ( $HMTB , @CRLF & "========================= Flush DNS: ===================================" & @CRLF )
	$READ = _RUN ( "ipconfig /flushdns" )
	FileWrite ( $HMTB , $READ )
EndFunc
Func GETFILEPATH ( $FOLDERNAME )
	Local $FILEARRAY , $I
	$FILEARRAY = _FILELISTTOARRAY ( $FOLDERNAME )
	If @error = 0 Then
		For $I = 1 To $FILEARRAY [ 0 ]
			$DEFAULTPATH = $FOLDERNAME & "\" & $FILEARRAY [ $I ]
			If StringInStr ( $DEFAULTPATH , ".default" ) Then ExitLoop
		Next
	EndIf
EndFunc
Func HKUUSERS ( )
	Local $HKUUSERS [ 1 ]
	$I = 0
	While 1
		$USER = RegEnumKey ( "HKEY_USERS" , $I )
		If @error Then ExitLoop
		If Not StringRegExp ( $USER , "(?i)(^S-1-5-18|_Classes)" ) Then _ARRAYADD ( $HKUUSERS , $USER , 0 , "||||" )
		$I += 1
	WEnd
	Return $HKUUSERS
EndFunc
Func HOSTS ( )
	Local $HOSTS , $REGEX , $READ , $REST , $Y
	GUICtrlSetData ( $LABEL , "Getting Hosts Content: " )
	If FileExists ( @TempDir & "\temp0" ) Then FileDelete ( @TempDir & "\temp0" )
	If FileExists ( @TempDir & "\temp00" ) Then FileDelete ( @TempDir & "\temp00" )
	If Not FileExists ( @WindowsDir & "\system32\drivers\etc\hosts" ) Then FileWrite ( $HMTB , "Hosts file not detected in the default directory" & @CRLF )
	If FileExists ( @WindowsDir & "\system32\drivers\etc\hosts" ) Then
		FileWrite ( $HMTB , "========================= Hosts content: =================================" & @CRLF )
		$HOSTS = FileRead ( @WindowsDir & "\system32\drivers\etc\hosts" )
		$HOSTS = StringRegExpReplace ( $HOSTS , "\n\v{1}" , "" )
		$HOSTS = StringRegExpReplace ( $HOSTS , "(?m)(^|\R)#.*?\R" , "" )
		$HANTEMP0 = FileOpen ( @TempDir & "\temp0" , 2 + 128 )
		FileWrite ( $HANTEMP0 , $HOSTS )
		FileClose ( $HANTEMP0 )
		$I = 1
		$Z = 1
		While 1
			$READ = FileReadLine ( @TempDir & "\temp0" , $I )
			If @error Or $Z = 31 Then ExitLoop
			If StringRegExp ( $READ , "\A\d.*" ) Then
				FileWrite ( $HMTB , $READ & @CRLF )
				$Z += 1
			EndIf
			$I += 1
		WEnd
		$NR = _FILECOUNTLINES ( @TempDir & "\temp0" )
		If $Z = 31 Then
			$Y = $NR - $Z
			If $Y > 0 Then FileWrite ( $HMTB , @CRLF & "There are " & $Y & " entries." & @CRLF & @CRLF )
		EndIf
		FileDelete ( @TempDir & "\temp0" )
	Else
		FileWrite ( $HMTB , @CRLF & "Hosts file not detected in the default directory" & @CRLF )
	EndIf
	GUICtrlSetData ( $LABEL , "" )
EndFunc
Func IPCONFIG ( )
	GUICtrlSetData ( $LABEL , "Getting ipconfig..." )
	FileWrite ( $HMTB , "========================= IP Configuration: ================================" & @CRLF & @CRLF )
	NIC ( )
	RunWait ( @ComSpec & " /c netsh int ip dump > """ & @TempDir & "\IP0402.txt"" & (ipconfig /all & nslookup google.com & ping -n 2 google.com & nslookup yahoo.com & ping -n 2 yahoo.com & ping -n 2 127.0.0.1 & route print)  >> """ & @TempDir & "\IP0402.txt""" , "" , @SW_HIDE )
	FileWrite ( $HMTB , FileRead ( @TempDir & "\IP0402.txt" ) )
	FileDelete ( @TempDir & "\IP0402.txt" )
	GUICtrlSetData ( $LABEL , "" )
EndFunc
Func MAINAPP ( )
	Local $CO = 0
	$HEVENTLOG = _EVENTLOG__OPEN ( "" , "application" )
	$AEVENT = _EVENTLOG__READ ( $HEVENTLOG , True , False )
	FileWrite ( $HMTB , @CRLF & "========================= Event log errors: ===============================" & @CRLF & @CRLF & "Application errors:" & @CRLF & "==================" & @CRLF )
	While 1
		$AEVENT = _EVENTLOG__READ ( $HEVENTLOG , False , False , $AEVENT [ 1 ] )
		If $AEVENT [ 1 ] = 0 Then ExitLoop
		GUICtrlSetData ( $LABEL , "Getting Application errors: " & $AEVENT [ 1 ] )
		If $AEVENT [ 7 ] = 1 Then
			FileWrite ( $HMTB , $AEVENT [ 8 ] & ": (" & $AEVENT [ 4 ] & " " & $AEVENT [ 5 ] & ") (" & "Source: " & $AEVENT [ 10 ] & ") (EventID: " & $AEVENT [ 6 ] & ") (User: " & $AEVENT [ 12 ] & ")" & @CRLF & "Description: " & $AEVENT [ 13 ] & @CRLF & @CRLF )
			$CO = $CO + 1
			If $CO = 10 Then ExitLoop
		EndIf
		$AEVENT [ 1 ] -= 1
	WEnd
	_EVENTLOG__CLOSE ( $HEVENTLOG )
EndFunc
Func MAINSYS ( )
	$CO = 0
	$HEVENTLOG = _EVENTLOG__OPEN ( "" , "system" )
	$AEVENT = _EVENTLOG__READ ( $HEVENTLOG , True , False )
	FileWrite ( $HMTB , @CRLF & "System errors:" & @CRLF & "=============" & @CRLF )
	While 1
		$AEVENT = _EVENTLOG__READ ( $HEVENTLOG , False , False , $AEVENT [ 1 ] )
		If $AEVENT [ 1 ] = 0 Then ExitLoop
		GUICtrlSetData ( $LABEL , "Getting system errors: " & $AEVENT [ 1 ] )
		If $AEVENT [ 7 ] = 1 Then
			$DESC = $AEVENT [ 13 ]
			If StringRegExp ( $DESC , "%%-?\d" ) Then
				$RET = StringRegExp ( $DESC , "%%-?\d+" , 1 )
				If IsArray ( $RET ) Then
					$RET = StringRegExpReplace ( $RET [ 0 ] , "%%" , "" )
					$RET = _WINAPI_GETERRORMESSAGE ( $RET )
					If $RET Then $DESC = StringRegExpReplace ( $DESC , "(%%-?\d+)" , "\1 = " & $RET )
				EndIf
			EndIf
			FileWrite ( $HMTB , $AEVENT [ 8 ] & ": (" & $AEVENT [ 4 ] & " " & $AEVENT [ 5 ] & ") (" & "Source: " & $AEVENT [ 10 ] & ") (EventID: " & $AEVENT [ 6 ] & ") (User: " & $AEVENT [ 12 ] & ")" & @CRLF & "Description: " & $DESC & @CRLF & @CRLF )
			$CO = $CO + 1
			If $CO = 10 Then ExitLoop
		EndIf
		$AEVENT [ 1 ] -= 1
	WEnd
	_EVENTLOG__CLOSE ( $HEVENTLOG )
EndFunc
Func MAKEUPMODEL ( ByRef $MANUFACTURER , ByRef $MODEL )
	$WBEMFLAGRETURNIMMEDIATELY = 16
	$WBEMFLAGFORWARDONLY = 32
	$COLITEMS = ""
	$STRCOMPUTER = "localhost"
	$OMYERROR = ""
	$OMYERROR = ObjEvent ( "AutoIt.Error" , "MyErrFunc" )
	$OBJWMISERVICE = ObjGet ( "winmgmts:\\" & $STRCOMPUTER & "\root\CIMV2" )
	$COLITEMS = $OBJWMISERVICE .ExecQuery ( "SELECT * FROM Win32_ComputerSystem" , "WQL" , $WBEMFLAGRETURNIMMEDIATELY + $WBEMFLAGFORWARDONLY )
	If IsObj ( $COLITEMS ) Then
		For $OBJITEM In $COLITEMS
			$MANUFACTURER = $OBJITEM .Manufacturer
			$MODEL = $OBJITEM .Model
			If $MANUFACTURER <> "" And $MODEL <> "" Then ExitLoop
		Next
	EndIf
EndFunc
Func MINIDUMP ( )
	FileWrite ( $HMTB , "========================= Minidump Files ==================================" & @CRLF & @CRLF )
	$ARR = _FILELISTTOARRAYREC ( @WindowsDir & "\Minidump" , "*.dmp" , 1 , 0 , 0 , 2 )
	For $I = 1 To UBound ( $ARR ) + 4294967295
		FileWrite ( $HMTB , $ARR [ $I ] & @CRLF )
	Next
	If UBound ( $ARR ) < 1 Then FileWrite ( $HMTB , "No minidump file found" & @CRLF & @CRLF )
EndFunc
Func MYERRFUNC ( )
EndFunc
Func NIC ( )
	Local $OBJWMISERVICE , $COLNIC
	$OMYERROR = ""
	$OMYERROR = ObjEvent ( "AutoIt.Error" , "MyErrFunc" )
	$OBJWMISERVICE = ObjGet ( "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2" )
	If @error = 0 Then
		Dim $PCINFO , $STATUS
		Local $I = 0
		While 1
			If $I = 13 Then ExitLoop
			$COLNIC = $OBJWMISERVICE .ExecQuery ( "Select * from Win32_NetworkAdapter WHERE Netconnectionstatus = """ & $I & """" )
			For $OBJECT In $COLNIC
				$PCINFO = $OBJECT .description & " = " & $OBJECT .NetConnectionID
				If $PCINFO <> "" Then
					$STATUS = ""
					If $I = 0 Then $STATUS = "(Disconnected)"
					If $I = 1 Then $STATUS = "(Connecting)"
					If $I = 2 Then $STATUS = "(Connected)"
					If $I = 3 Then $STATUS = "(Disconnecting)"
					If $I = 4 Then $STATUS = "(Hardware not present)"
					If $I = 5 Then $STATUS = "(Hardware disabled)"
					If $I = 6 Then $STATUS = "(Hardware malfunction)"
					If $I = 7 Then $STATUS = "(Media disconnected)"
					If $I = 8 Then $STATUS = "(Authenticating)"
					If $I = 9 Then $STATUS = "(Authentication succeeded)"
					If $I = 10 Then $STATUS = "(Authentication failed)"
					If $I = 11 Then $STATUS = "(Invalid address)"
					If $I = 12 Then $STATUS = "(Credentials required)"
					FileWrite ( $HMTB , $PCINFO & " " & $STATUS & @CRLF )
				EndIf
			Next
			$I = $I + 1
		WEnd
	EndIf
EndFunc
Func OS ( )
	Local $OBJWMISERVICE , $DEVCOLITEMS
	$SVERSION = ""
	If _SRVSTAT ( "winmgmt" ) = "R" Then
		$OMYERROR = ""
		$OMYERROR = ObjEvent ( "AutoIt.Error" , "MyErrFunc" )
		$OBJWMISERVICE = ObjGet ( "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2" )
		If Not @error Then
			$DEVCOLITEMS = $OBJWMISERVICE .ExecQuery ( "Select * from Win32_OperatingSystem" )
			If Not @error Then
				For $OBJECT In $DEVCOLITEMS
					If ( $OBJECT .caption ) Then $SVERSION = $OBJECT .caption
				Next
			EndIf
		EndIf
	EndIf
EndFunc
Func PACKAGES1 ( $STR , $PUB = 0 )
	If $PUB = 1 Then
		$STR = StringRegExpReplace ( $STR , "(?i)AppDisplayName|AppStoreName|StoreAppName|PackageDisplayName|PkgDisplayName|DisplayName|ConnectorStubTitle|ApplicationTitleWithBranding|AppName|ProductName|DisplayTitle|appDescription|StoreTitle|GameBar|IDS_MANIFEST_(MUSIC|VIDEO)_APP_NAME" , "PublisherDisplayName" )
	EndIf
	$STR = _WINAPI_LOADINDIRECTSTRING ( $STR )
	If Not @error Then Return $STR
EndFunc
Func PACKAGES2 ( $READ )
	If StringRegExp ( $READ , "(?i)<PublisherDisplayName>[^<]*<" ) Then
		Return StringRegExpReplace ( $READ , "(?is).+<PublisherDisplayName>([^<]*)</PublisherDisplayName.+" , "$1" )
	EndIf
EndFunc
Func PART ( )
	Local $MEM , $VAR , $DT , $DLABEL , $DSPT , $DSPACE , $DFS , $REGEX
	$MEM = MemGetStats ( )
	$MEM [ 1 ] = $MEM [ 1 ] / 1024
	$MEM [ 1 ] = Round ( $MEM [ 1 ] , 2 )
	$MEM [ 2 ] = $MEM [ 2 ] / 1024
	$MEM [ 2 ] = Round ( $MEM [ 2 ] , 2 )
	$MEM [ 3 ] = $MEM [ 3 ] / 1024
	$MEM [ 3 ] = Round ( $MEM [ 3 ] , 2 )
	$MEM [ 4 ] = $MEM [ 4 ] / 1024
	$MEM [ 4 ] = Round ( $MEM [ 4 ] , 2 )
	FileWrite ( $HMTB , @CRLF & "========================= Memory info: ===================================" & @CRLF & @CRLF & "Percentage of memory in use: " & $MEM [ 0 ] & "%" & @CRLF & "Total physical RAM: " & $MEM [ 1 ] & " " & "MB" & @CRLF & "Available physical RAM: " & $MEM [ 2 ] & " MB" & @CRLF & "Total Virtual: " & $MEM [ 3 ] & " MB" & @CRLF & "Available " & "Virtual: " & $MEM [ 4 ] & " MB" & @CRLF & @CRLF & "========================= Partitions: =====================================" & @CRLF & @CRLF )
	$VAR = DriveGetDrive ( "all" )
	If Not @error Then
		For $I = 1 To $VAR [ 0 ]
			$DT = DriveGetType ( $VAR [ $I ] & "\" )
			$DLABEL = DriveGetLabel ( $VAR [ $I ] & "\" )
			$DSPT = DriveSpaceTotal ( $VAR [ $I ] & "\" )
			$DSPT = $DSPT / 1024
			$DSPT = Round ( $DSPT , 2 )
			$DSPACE = DriveSpaceFree ( $VAR [ $I ] & "\" )
			$DSPACE = $DSPACE / 1024
			$DSPACE = Round ( $DSPACE , 2 )
			$DFS = DriveGetFileSystem ( $VAR [ $I ] & "\" )
			If $DSPT > 0 Then FileWrite ( $HMTB , $I & " Drive " & $VAR [ $I ] & " (" & $DLABEL & ") (" & $DT & ") (Total:" & $DSPT & " GB) (Free:" & $DSPACE & " GB) " & $DFS & @CRLF )
		Next
	EndIf
	If FileExists ( @TempDir & "\temp1" ) Then FileDelete ( @TempDir & "\temp1" )
	FileWrite ( $HMTB , @CRLF & "========================= Users: ========================================" & @CRLF )
	RunWait ( @ComSpec & " /u/c " & "net users >> """ & @TempDir & "\temp1""" , "" , @SW_HIDE )
	$REGEX = FileRead ( @TempDir & "\temp1" )
	$REGEX = StringRegExpReplace ( $REGEX , "The command[^\.]+\.\v{2}" , "" )
	$REGEX = StringRegExpReplace ( $REGEX , "----" , "" )
	$REGEX = StringRegExpReplace ( $REGEX , "---\v{2}" , "" )
	FileWrite ( $HMTB , $REGEX )
	FileDelete ( @TempDir & "\temp1" )
EndFunc
Func PROGRAMS ( )
	Local $VAR1 , $VAR2 , $VERS0 , $VERS1 , $VAR , $LINE , $AARRAY
	Local $HPROG [ 1 ]
	GUICtrlSetData ( $LABEL , "Listing Installed Programs..." )
	$UNINSTALL = "hklm64\Software\Microsoft\Windows\CurrentVersion\Uninstall"
	$I = 0
	While 1
		$I = $I + 1
		$VAR1 = RegEnumKey ( $UNINSTALL , $I )
		If @error <> 0 Then ExitLoop
		If Not RegRead ( $UNINSTALL & "\" & $VAR1 , "UninstallString" ) Then ContinueLoop
		$VAR2 = RegRead ( $UNINSTALL & "\" & $VAR1 , "Displayname" )
		$VAR2 = StringRegExpReplace ( $VAR2 , "^\s+" , "" )
		If @error = 0 And $VAR2 <> "" And Not StringRegExp ( $VAR2 , "(?i)(Security Update for|hotfix for|Update for Microsoft)" ) Then
			$PUB = RegRead ( $UNINSTALL & "\" & $VAR1 , "Publisher" )
			$VERS0 = RegRead ( $UNINSTALL & "\" & $VAR1 , "Displayversion" )
			$VARH = RegRead ( $UNINSTALL & "\" & $VAR1 , "SystemComponent" )
			If $VARH = 1 Then
				If Not StringRegExp ( $VAR2 , "(?i)^(Service Pack|Microsoft|NVIDIA|AMD Fuel|ccc-utility|Zune|Visual|Windows|amd |Logitech|Intel|Native|TOSHIBA|HP |MSVCRT)" ) Then _ARRAYADD ( $HPROG , $VAR2 & " (HKLM\...\" & $VAR1 & ") (Version: " & $VERS0 & " - " & $PUB & ") Hidden" , 0 , "||||" )
			Else
				_ARRAYADD ( $HPROG , $VAR2 & " (HKLM\...\" & $VAR1 & ") (Version: " & $VERS0 & " - " & $PUB & ")" , 0 , "||||" )
			EndIf
		EndIf
	WEnd
	$USERREG = HKUUSERS ( )
	For $U = 1 To UBound ( $USERREG ) + 4294967295
		$UNINSTALL = "hku\" & $USERREG [ $U ] & "\Software\Microsoft\Windows\CurrentVersion\Uninstall"
		$I = 0
		While 1
			$I += 1
			$VAR1 = RegEnumKey ( $UNINSTALL , $I )
			If @error <> 0 Then ExitLoop
			If Not RegRead ( $UNINSTALL & "\" & $VAR1 , "UninstallString" ) Then ContinueLoop
			$VAR2 = RegRead ( $UNINSTALL & "\" & $VAR1 , "Displayname" )
			$VAR2 = StringRegExpReplace ( $VAR2 , "^\s+" , "" )
			If @error = 0 And $VAR2 <> "" And Not StringRegExp ( $VAR2 , "(?i)(Security Update for|hotfix for|Update for Microsoft)" ) Then
				$PUB = RegRead ( $UNINSTALL & "\" & $VAR1 , "Publisher" )
				$VERS0 = RegRead ( $UNINSTALL & "\" & $VAR1 , "Displayversion" )
				$VARH = RegRead ( $UNINSTALL & "\" & $VAR1 , "SystemComponent" )
				$VARHH = ""
				If $VARH = 1 Then $VARHH = " Hidden"
				_ARRAYADD ( $HPROG , $VAR2 & " (HKU\" & $USERREG [ $U ] & "\...\" & $VAR1 & ") (Version: " & $VERS0 & " - " & $PUB & ")" & $VARHH , 0 , "||||" )
			EndIf
		WEnd
	Next
	$UNINSTALL = "hklm64\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	$I = 0
	While 1
		$I += 1
		$VAR1 = RegEnumKey ( $UNINSTALL , $I )
		If @error <> 0 Then ExitLoop
		If Not RegRead ( $UNINSTALL & "\" & $VAR1 , "UninstallString" ) Then ContinueLoop
		$VAR2 = RegRead ( $UNINSTALL & "\" & $VAR1 , "Displayname" )
		$VAR2 = StringRegExpReplace ( $VAR2 , "^\s+" , "" )
		If @error = 0 And $VAR2 <> "" And Not StringRegExp ( $VAR2 , "(?i)(Security Update for|hotfix for|Update for Microsoft)" ) Then
			$PUB = RegRead ( $UNINSTALL & "\" & $VAR1 , "Publisher" )
			$VERS0 = RegRead ( $UNINSTALL & "\" & $VAR1 , "Displayversion" )
			$VARH = RegRead ( $UNINSTALL & "\" & $VAR1 , "SystemComponent" )
			$VARHH = ""
			If $VARH = 1 Then $VARHH = " Hidden"
			If Not StringRegExp ( $VAR2 , "(?i)^(tools|Microsoft|NVIDIA|Nero|Turbo|Visual|CCC Help|Catalyst|Service Pack|CyberLink|Windows|Adobe|LWS|Citrix|Roxio|amd |Logitech|MSVCRT|Java|TOSHIBA|HP |Movie Maker |Photo )" ) Then _ARRAYADD ( $HPROG , $VAR2 & " (HKLM-x32\...\" & $VAR1 & ") (Version: " & $VERS0 & " - " & $PUB & ")" & $VARHH , 0 , "||||" )
		EndIf
	WEnd
	_ARRAYDELETE ( $HPROG , 0 )
	$HPROG = _ARRAYUNIQUE ( $HPROG , 0 , 0 , 0 , 0 , 1 )
	_ARRAYSORT ( $HPROG , 0 )
	FileWrite ( $HMTB , @CRLF & "=========================== Installed Programs ============================" & @CRLF & @CRLF )
	_FILEWRITEFROMARRAY ( $HMTB , $HPROG )
	_AAAAP1 ( )
EndFunc
Func PROXYLIST ( )
	Local $FOLDERNAME
	GUICtrlSetData ( $LABEL , "Checking FF Proxy setting" )
	$FOLDERNAME = @AppDataDir & "\Mozilla\Firefox\Profiles"
	If FileExists ( $FOLDERNAME ) Then
		GETFILEPATH ( $FOLDERNAME )
		READPROXY ( )
	EndIf
EndFunc
Func PROXYRESETFF ( )
	Local $FOLDERNAME , $REGEXPR
	GUICtrlSetData ( $LABEL , "Resetting FF Proxy setting" )
	$FOLDERNAME = @AppDataDir & "\Mozilla\Firefox\Profiles"
	If FileExists ( $FOLDERNAME ) Then
		GETFILEPATH ( $FOLDERNAME )
		RunWait ( @ComSpec & " /c " & "tskill firefox" , "" , @SW_HIDE )
		$REGEXPR = FileRead ( $DEFAULTPATH & "\prefs.js" )
		$REGEXPR = StringRegExpReplace ( $REGEXPR , "user_pref\(""network.proxy[^;]*;\v{2}" , "" )
		FileWrite ( @TempDir & "\temp0.js" , $REGEXPR )
		FileMove ( @TempDir & "\temp0.js" , $DEFAULTPATH & "\prefs.js" , 1 )
		FileWrite ( $HMTB , @CRLF & """Reset FF Proxy Settings"": Firefox Proxy settings were reset." & @CRLF & @CRLF )
	EndIf
EndFunc
Func READPROXY ( )
	Local $LINE , $REGEXPR
	FileWrite ( $HMTB , @CRLF & "========================= FF Proxy Settings: ============================== " & @CRLF & @CRLF )
	Local $I = 1
	While 1
		$LINE = FileReadLine ( $DEFAULTPATH & "\prefs.js" , $I )
		If @error <> 0 Then ExitLoop
		If StringInStr ( $LINE , "network.proxy." ) Then
			$REGEXPR = StringRegExpReplace ( $LINE , "user_pref\(" , "" )
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "\);" , "" )
			FileWrite ( $HMTB , $REGEXPR & @CRLF )
		EndIf
		$I = $I + 1
	WEnd
EndFunc
Func RESTOREPOINTS ( )
	Local $NAME , $DATE , $OBJWMISERVICE , $DEVCOLITEMS
	GUICtrlSetData ( $LABEL , "Getting Restore Points" )
	FileWrite ( $HMTB , "========================= Restore Points ==================================" & @CRLF & @CRLF )
	$OMYERROR = ""
	$OMYERROR = ObjEvent ( "AutoIt.Error" , "MyErrFunc" )
	$OBJWMISERVICE = ObjGet ( "winmgmts:{impersonationLevel=impersonate}!\\.\root\default" )
	If Not @error = 0 Then
		FileWrite ( $HMTB , "Could not list Restore Points." & @CRLF )
	Else
		$DEVCOLITEMS = $OBJWMISERVICE .ExecQuery ( "Select * from SystemRestore" )
		For $OBJECT In $DEVCOLITEMS
			$NAME = $OBJECT .description
			$DATE = $OBJECT .CreationTime
			If $NAME <> "" Then
				$DATE = StringRegExpReplace ( $DATE , "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2}).+" , "$3-$2-$1 $4:$5:$6 " )
				FileWrite ( $HMTB , $DATE & $NAME & @CRLF )
			EndIf
		Next
	EndIf
EndFunc
Func SHOWPROXY ( )
	Local $VAL1 , $RES1 , $VAL2 , $RES2
	GUICtrlSetData ( $LABEL , "Checking IE Proxy setting" )
	$VAL1 = RegRead ( "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyEnable" )
	If Not $VAL1 = "1" Then $RES1 = "Proxy is not enabled."
	If $VAL1 = "1" Then $RES1 = "Proxy is enabled."
	$VAL2 = RegRead ( "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" , "ProxyServer" )
	If $VAL2 = "" Then $RES2 = "No Proxy Server is set."
	If Not $VAL2 = "" Then $RES2 = "ProxyServer: " & $VAL2
	FileWrite ( $HMTB , @CRLF & "========================= IE Proxy Settings: ============================== " & @CRLF & @CRLF & $RES1 & @CRLF & $RES2 & @CRLF )
EndFunc
Func WINDEF ( )
	$PATH = @TempDir & "\codeint" & Random ( 1000 , 9999 , 1 )
	RunWait ( @ComSpec & " /c " & "wevtutil qe ""Microsoft-Windows-Windows Defender/Operational"" ""/q:*[System [(Level=3)]]"" /c:5 /rd:true /uni:true /f:text >> """ & $PATH & """" , "" , @SW_HIDE )
	RunWait ( @ComSpec & " /c " & "wevtutil qe ""Microsoft-Windows-Windows Defender/Operational"" ""/q:*[System [(Level=2)]]"" /c:5 /rd:true /uni:true /f:text >> """ & $PATH & """" , "" , @SW_HIDE )
	$HCODE = FileOpen ( $PATH , 256 )
	$EVENTS = FileRead ( $HCODE )
	FileClose ( $HCODE )
	If Not StringInStr ( $EVENTS , ":" ) Then Return
	$EVENTS = StringRegExpReplace ( $EVENTS , "(?m)^\s*" , "" )
	$EVENTS = StringRegExpReplace ( $EVENTS , "(?m)^Event\[\d\]:?\R" , "" )
	$EVENTS = StringRegExpReplace ( $EVENTS , "(?m)^(Log Name|(Scan|Event) ID|Task|Level|Opcode|Keyword|Source|User|User Name|Computer|ID):.*\R" , "" )
	$EVENTS = StringRegExpReplace ( $EVENTS , "(?m)^(Date:.+\d)T(\d.+\v{2})" , @CRLF & "$1 $2" )
	$EVENTS = StringRegExpReplace ( $EVENTS , "(?m)^(Date:.+?)\.\d+Z" , "$1" )
	$EVENTS = StringRegExpReplace ( $EVENTS , "(?s)\R{2,}" , @CRLF & @CRLF )
	FileWrite ( $HMTB , @CRLF & "Windows Defender:" & @CRLF & "================" & $EVENTS )
	FileDelete ( $PATH )
EndFunc
Func WINSOCK ( )
	FileWrite ( $HMTB , "========================= Winsock entries =====================================" & @CRLF & @CRLF )
	$OSNUM = _WINAPI_GETVERSION ( )
	$KEY = "HKLM64\SYSTEM\CurrentControlSet\Services\WinSock2\Parameters\"
	Local $I = 1
	While 1
		$WSCHECK = ""
		$VAR = RegEnumKey ( $KEY & "NameSpace_Catalog5\Catalog_Entries" , $I )
		If @error <> 0 Then ExitLoop
		$VAL = RegRead ( $KEY & "NameSpace_Catalog5\Catalog_Entries\" & $VAR , "LibraryPath" )
		If @error <> 0 Then ExitLoop
		$REGEXPR = $VAL
		If StringRegExp ( $REGEXPR , "(?i)(%SystemRoot%|%windir%)" ) Then
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)%SystemRoot%|%windir%" , @WindowsDir )
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)([a-z]:)" , "$1\\" )
		EndIf
		If @OSArch = "x64" Then
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)\\System32\\" , "\\SysWOW64\\" )
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)%Programfiles%" , $C & "\\Program Files \(x86\)" )
			$SYSDIR = "SysWOW64"
		Else
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)%Programfiles%" , $C & "\\Program Files" )
			$SYSDIR = "System32"
		EndIf
		$VAL2 = RegRead ( $KEY & "NameSpace_Catalog5\Catalog_Entries\" & $VAR , "ProviderId" )
		Select
		Case $VAL2 == "0x409D05229E7ECF11AE5A00AA00A7112B"
			If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\" & $SYSDIR & "\mswsock.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\System32\mswsock.dll""" & @CRLF
		Case $VAL2 == "0xA2CB4A96BCB2EB408C6AA6DB40161CAE"
			If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\" & $SYSDIR & "\napinsp.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\napinsp.dll""" & @CRLF
		Case $VAL2 == "0xCE89FE036D767649B9C1BB9BC42C7B4D" Or $VAL2 == "0xCD89FE036D767649B9C1BB9BC42C7B4D"
			If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\" & $SYSDIR & "\pnrpnsp.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\pnrpnsp.dll""" & @CRLF
		Case $VAL2 == "0x3A244266A83BA64ABAA52E0BD71FDD83"
			If $OSNUM < 6 Then
				If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\System32\mswsock.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\mswsock.dll"""
			Else
				If Not StringRegExp ( $REGEXPR , "(?i)^" & StringRegExpReplace ( EnvGet ( "SystemRoot" ) , "\\" , "\\\\" ) & "\\" & $SYSDIR & "\\nla(nsp_c|api).dll" ) Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\NLAapi.dll"""
			EndIf
		Case $VAL2 == "0xEE37263B80E5CF11A55500C04FD8D4AC"
			If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\" & $SYSDIR & "\winrnr.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\System32\winrnr.dll""" & @CRLF
		EndSelect
		If StringRegExp ( $VAL , "(?i)(%SystemRoot%|%windir%)" ) Then
			$VAL = StringRegExpReplace ( $VAL , "(?i)(%SystemRoot%|%windir%)\\System32" , @SystemDir )
			$VAL = StringRegExpReplace ( $VAL , "(?i)([a-z]:)" , "$1\\" )
			If @OSArch = "X86" Then
				$VAL = StringRegExpReplace ( $VAL , "(?i)(System32)" , "\\$1" )
			Else
				$VAL = StringRegExpReplace ( $VAL , "(?i)(SysWOW64)" , "\\$1" )
			EndIf
		EndIf
		If StringRegExp ( $VAL , "(?i)%Programfiles%" ) Then
			$VAL = StringRegExpReplace ( $VAL , "(?i)%Programfiles%" , @ProgramFilesDir )
			$VAL = StringRegExpReplace ( $VAL , "(?i)([a-z]:)" , "$1\\" )
		EndIf
		$VER = FileGetVersion ( $VAL , "CompanyName" )
		$SIZE = FileGetSize ( $VAL )
		If Not FileExists ( $VAL ) Then
			$SIZE = "File Not found"
			$SIZE = ""
		EndIf
		$VAR = StringRegExpReplace ( $VAR , "0000000000" , "" )
		FileWrite ( $HMTB , "Catalog5 " & $VAR & " " & $VAL & " [" & $SIZE & "]" & " (" & $VER & ")" & $WSCHECK & @CRLF )
		$I = $I + 1
	WEnd
	$NUM = RegRead ( $KEY & "NameSpace_Catalog5" , "Num_Catalog_Entries" )
	If $NUM > $I + 4294967295 Then FileWrite ( $HMTB , "Winsock: Missing Catalog5 entry, broken internet access. " & "<===== ATTENTION." & @CRLF )
	$I = 1
	While 1
		$VAL = ""
		$VAR = ""
		$VAR = RegEnumKey ( $KEY & "Protocol_Catalog9\Catalog_Entries" , $I )
		If @error <> 0 Then ExitLoop
		$VAL = RegRead ( $KEY & "Protocol_Catalog9\Catalog_Entries\" & $VAR , "PackedCatalogItem" )
		$REGEXPR = ""
		$REGEXPR = StringRegExpReplace ( $VAL , "00.+" , "" )
		$REGEXPR = _HEXTOSTRING ( $REGEXPR )
		If StringRegExp ( $REGEXPR , "(?i)(%SystemRoot%|%windir%)" ) Then
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)(%SystemRoot%|%windir%)\\System32" , @SystemDir )
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)([a-z]:)" , "$1\\" )
			If @OSArch = "X86" Then
				$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)(System32)" , "\\$1" )
			Else
				$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)(SysWOW64)" , "\\$1" )
			EndIf
		EndIf
		$VER = ""
		$VER = FileGetVersion ( $REGEXPR , "CompanyName" )
		$SIZE = FileGetSize ( $REGEXPR )
		If Not FileExists ( $REGEXPR ) Then
			$SIZE = "File not found"
		EndIf
		$VAR = StringRegExpReplace ( $VAR , "0000000000" , "" )
		FileWrite ( $HMTB , "Catalog9 " & $VAR & " " & $REGEXPR & " [" & $SIZE & "]" & " (" & $VER & ")" & @CRLF )
		$I = $I + 1
	WEnd
	$NUM = RegRead ( $KEY & "Protocol_Catalog9" , "Num_Catalog_Entries" )
	If $NUM > $I + 4294967295 Then FileWrite ( $HMTB , "Winsock: Missing Catalog9 entry, broken internet access. " & "<===== ATTENTION." & @CRLF )
	If @OSArch = "X64" Then WINSOCK64 ( )
EndFunc
Func WINSOCK64 ( )
	Local $REGEXPR
	$KEY = "HKLM64\SYSTEM\CurrentControlSet\Services\WinSock2\Parameters\"
	Local $I = 1
	While 1
		$WSCHECK = ""
		$VAR = RegEnumKey ( $KEY & "NameSpace_Catalog5\Catalog_Entries64" , $I )
		If @error Then ExitLoop
		$VAL = RegRead ( $KEY & "NameSpace_Catalog5\Catalog_Entries64\" & $VAR , "LibraryPath" )
		If @error = 0 Then
			$REGEXPR = $VAL
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)%SystemRoot%" , $C & "\\Windows" )
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)%windir%" , $C & "\\Windows" )
			$REGEXPR = StringRegExpReplace ( $REGEXPR , "(?i)%Programfiles%" , $C & "\\Program Files" )
			$VAL2 = RegRead ( $KEY & "NameSpace_Catalog5\Catalog_Entries64\" & $VAR , "ProviderId" )
			Select
			Case $VAL2 == "0x409D05229E7ECF11AE5A00AA00A7112B"
				If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\System32\mswsock.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\System32\mswsock.dll""" & @CRLF
			Case $VAL2 == "0xA2CB4A96BCB2EB408C6AA6DB40161CAE"
				If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\System32\napinsp.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\napinsp.dll""" & @CRLF
			Case $VAL2 = "0xCE89FE036D767649B9C1BB9BC42C7B4D" Or $VAL2 = "0xCD89FE036D767649B9C1BB9BC42C7B4D"
				If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\System32\pnrpnsp.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\pnrpnsp.dll""" & @CRLF
			Case $VAL2 == "0x3A244266A83BA64ABAA52E0BD71FDD83"
				If Not StringRegExp ( $REGEXPR , "(?i)^" & StringRegExpReplace ( EnvGet ( "SystemRoot" ) , "\\" , "\\\\" ) & "\\System32\\nla(nsp_c|api).dll" ) Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\system32\NLAapi.dll""" & @CRLF
			Case $VAL2 == "0xEE37263B80E5CF11A55500C04FD8D4AC"
				If $REGEXPR <> EnvGet ( "SystemRoot" ) & "\System32\winrnr.dll" Then $WSCHECK = @CRLF & "ATTENTION: The LibraryPath should be ""%SystemRoot%\System32\winrnr.dll""" & @CRLF
			EndSelect
			$REGEXP = ""
			$REGEXP = StringRegExpReplace ( $VAL , "(?i)%SystemRoot%" , @HomeDrive & "\\Windows" )
			$REGEXP = StringRegExpReplace ( $REGEXP , "(?i)%windir%" , @HomeDrive & "\\Windows" )
			$REGEXP = StringRegExpReplace ( $REGEXP , "(?i)System32" , "Sysnative" )
			$VER = ""
			$VER = FileGetVersion ( $REGEXP , "CompanyName" )
			$SIZE = FileGetSize ( $REGEXP )
			If Not FileExists ( $REGEXP ) Then
				$SIZE = "File Not found"
			EndIf
			$VAR = StringRegExpReplace ( $VAR , "0000000000" , "" )
			$REGEXP = StringRegExpReplace ( $REGEXP , "(?i)Sysnative" , "System32" )
			FileWrite ( $HMTB , "x64-Catalog5 " & $VAR & " " & $REGEXP & " [" & $SIZE & "]" & " (" & $VER & ")" & $WSCHECK & @CRLF )
		EndIf
		$I = $I + 1
	WEnd
	$NUM = RegRead ( $KEY & "NameSpace_Catalog5" , "Num_Catalog_Entries64" )
	If $NUM > $I + 4294967295 Then FileWrite ( $HMTB , "Winsock: Missing Catalog5-x64 entry, broken internet access. " & "<===== ATTENTION." & @CRLF )
	$I = 1
	While 1
		$VAR = RegEnumKey ( $KEY & "Protocol_Catalog9\Catalog_Entries64" , $I )
		If @error <> 0 Then ExitLoop
		$VAL = RegRead ( $KEY & "Protocol_Catalog9\Catalog_Entries64\" & $VAR , "PackedCatalogItem" )
		If @error = 0 Then
			$REGEXPR = StringRegExpReplace ( $VAL , "00.+" , "" )
			$STRING = _HEXTOSTRING ( $REGEXPR )
			$REGEXP = StringRegExpReplace ( $STRING , "(?i)%SystemRoot%" , @HomeDrive & "\\Windows" )
			$REGEXP = StringRegExpReplace ( $REGEXP , "(?i)%windir%" , @HomeDrive & "\\Windows" )
			$REGEXP = StringRegExpReplace ( $REGEXP , "(?i)System32" , "Sysnative" )
			$VER = FileGetVersion ( $REGEXP , "CompanyName" )
			$SIZE = FileGetSize ( $REGEXP )
			If Not FileExists ( $REGEXP ) Then
				$SIZE = "File Not found"
			EndIf
			$VAR = StringRegExpReplace ( $VAR , "0000000000" , "" )
			$REGEXP = StringRegExpReplace ( $REGEXP , "(?i)Sysnative" , "System32" )
			FileWrite ( $HMTB , "x64-Catalog9 " & $VAR & " " & $REGEXP & " [" & $SIZE & "]" & " (" & $VER & ")" & @CRLF )
		EndIf
		$I = $I + 1
	WEnd
	$NUM = RegRead ( $KEY & "Protocol_Catalog9" , "Num_Catalog_Entries64" )
	If $NUM > $I + 4294967295 Then FileWrite ( $HMTB , "Winsock: Missing Catalog9-x64 entry, broken internet access. " & "<===== ATTENTION." & @CRLF )
EndFunc
