#SingleInstance Off
#Warn Unreachable, Off

; __________________________________________________________________________________________________
; Accept the main JS file from the command-line
params := getArgs()
main := ""
for key, val in params {
	if (RegExMatch(val, "i)\.js$")) {
		main := val
		break
	}
}
if (!main) {
	main := "demo.js"
}
global mainPath := ""
if (SubStr(main, 2, 1) != ":") ; Not absolute path
	mainPath := rel2AbsPath(main)
if !FileExist(mainPath)
	end("The main file could not be located!`n" . main)



; __________________________________________________________________________________________________
; Constants
; TODO: Check whats actually needed (eg IsCompiled)
; and whats number type
ENUM_VARIABLES := "
(
A_AhkPath A_AhkVersion A_AllowMainWindow A_AppData A_AppDataCommon A_Args A_Clipboard
A_ComSpec A_ComputerName A_CoordModeCaret A_CoordModeMenu A_CoordModeHouse A_CoordModePixel
A_CoordModeTooltip A_Cursor A_DD A_DDD A_DDDD A_Desktop A_DesktopCommon A_DetectHiddenText
A_DetectHiddenWindows A_ENDCHAR A_EVENTINFO A_FileEncoding A_HotIf A_HotkeyInterval A_HotkeyModifierTimeout
A_Hour A_IconFile A_IconNumber A_IconTip A_Index A_InitialWorkingDir A_IsCompiled
A_KeyDuration A_KeyDurationPlay A_KeybdHookInstalled A_Language A_LastError A_LineFile
A_LineNumber A_ListLines A_LoopField A_LoopFileName A_LoopReadLine A_LoopRegKey A_LoopRegName A_MM
A_MMM A_MMMM A_MSEC A_MaxHotkeysPerInterval A_MenuMaskKey A_Min
A_MouseHookInstalled A_MyDocuments A_Now A_NowUTC A_OSVersion A_PriorHotkey A_PriorKey
A_ProgramFiles A_Programs A_ProgramsCommon A_RegView A_ScriptDir A_ScriptFullPath A_ScriptHwnd
A_ScriptName A_Sec A_SendLevel A_SendMode A_Space A_StartMenu A_StartMenuCommon
A_Startup A_StartupCommon A_StoreCapsLockMode A_Tab A_Temp A_ThisFunc A_ThisHotkey
A_TimeIdleKeyboard A_TimeIdleMouse A_TitleMatchMode A_TitleMatchModeSpeed A_TrayMeny A_UserName
A_WDay A_WinDir A_WorkingDir A_YDay A_YWeek A_YYYY
)"

;FOR GOTO IsSet
;TODO Loop SysGetIPAddresses
ENUM_FUNCTIONS := "
(
Abs ASin ACos ATan BlockInput Buffer CallbackCreate CallbackFree CaretGetPos 
Ceil Chr Click ClipboardAll ClipWait ComCall ComObjActive ComObjArray ComObjConnect ComObject 
ComObjFlags ComObjFromPtr ComObjGet ComObjQuery ComObjType ComObjValue ComValue ControlAddItem ControlChooseIndex 
ControlChooseString ControlClick ControlDeleteItem ControlFindItem ControlFocus ControlGetChecked ControlGetChoice ControlGetClassNN ControlGetEnabled ControlGetFocus 
ControlGetHwnd ControlGetIndex ControlGetItems ControlGetPos ControlGetStyle ControlGetExStyle ControlGetText ControlGetVisible ControlHide ControlHideDropDown 
ControlMove ControlSend ControlSendText ControlSetChecked ControlSetEnabled ControlSetStyle ControlSetExStyle ControlSetText ControlShow ControlShowDropDown 
CoordMode Cos Critical DateAdd DateDiff DetectHiddenText DetectHiddenWindows DirCopy DirCreate DirDelete 
DirExist DirMove DirSelect DllCall Download DriveEject DriveGetCapacity DriveGetFileSystem DriveGetLabel DriveGetList 
DriveGetSerial DriveGetSpaceFree DriveGetStatus DriveGetStatusCD DriveGetType DriveLock DriveRetract DriveSetLabel DriveUnlock Edit 
EditGetCurrentCol EditGetCurrentLine EditGetLine EditGetLineCount EditGetSelectedText EditPaste EnvGet EnvSet Exit 
ExitApp Exp FileAppend FileCopy FileCreateShortcut FileDelete FileEncoding FileExist FileInstall FileGetAttrib 
FileGetShortcut FileGetSize FileGetTime FileGetVersion FileMove FileOpen FileRead FileRecycle FileRecycleEmpty FileSelect 
FileSetAttrib FileSetTime Float Floor Format FormatTime GetKeyName GetKeyVK 
GetKeySC GetKeyState GetMethod GroupActivate GroupAdd GroupClose GroupDeactivate Gui GuiCtrlFromHwnd 
GuiFromHwnd HasBase HasMethod HasProp Hotkey Hotstring IL_Create IL_Add IL_Destroy 
ImageSearch IniDelete IniRead IniWrite InputBox InputHook InstallKeybdHook InstallMouseHook InStr Integer 
IsLabel IsObject IsSetRef KeyHistory KeyWait ListHotkeys ListLines ListVars ListViewGetContent 
LoadPicture Log Ln Map Max MenuBar Menu MenuFromHandle MenuSelect 
Min Mod MonitorGet MonitorGetCount MonitorGetName MonitorGetPrimary MonitorGetWorkArea MouseClick MouseClickDrag MouseGetPos 
MouseMove MsgBox Number NumGet NumPut ObjAddRef ObjRelease ObjBindMethod ObjHasOwnProp ObjOwnProps 
ObjGetBase ObjGetCapacity ObjGetDataPtr ObjGetDataSize ObjOwnPropCount ObjSetBase ObjSetCapacity ObjSetDataPtr OnClipboardChange OnError 
OnExit OnMessage Ord OutputDebug Pause Persistent PixelGetColor PixelSearch PostMessage ProcessClose 
ProcessExist ProcessGetName ProcessGetParent ProcessGetPath ProcessSetPriority ProcessWait ProcessWaitClose Random RegExMatch RegExReplace 
RegCreateKey RegDelete RegDeleteKey RegRead RegWrite Reload Round Run RunAs 
RunWait Send SendText SendInput SendPlay SendEvent SendLevel SendMessage SendMode SetCapsLockState 
SetControlDelay SetDefaultMouseSpeed SetKeyDelay SetMouseDelay SetNumLockState SetScrollLockState SetRegView SetStoreCapsLockMode SetTimer SetTitleMatchMode 
SetWinDelay SetWorkingDir Shutdown Sin Sleep Sort SoundBeep SoundGetInterface SoundGetMute SoundGetName 
SoundGetVolume SoundPlay SoundSetMute SoundSetVolume SplitPath Sqrt StatusBarGetText StatusBarWait StrCompare StrGet 
String StrLen StrLower StrPtr StrPut StrReplace StrSplit StrTitle StrUpper SubStr 
Suspend SysGet SysGetIPAddresses Tan Thread Throw ToolTip TraySetIcon TrayTip 
Trim LTrim RTrim Type VarSetStrCapacity VerCompare WinActivate WinActivateBottom 
WinActive WinClose WinExist WinGetAlwaysOnTop WinGetClass WinGetClientPos WinGetControls WinGetControlsHwnd WinGetCount WinGetEnabled 
WinGetID WinGetIDLast WinGetList WinGetMinMax WinGetPID WinGetPos WinGetProcessName WinGetProcessPath WinGetStyle WinGetExStyle 
WinGetText WinGetTitle WinGetTransColor WinGetTransparent WinHide WinKill WinMaximize WinMinimize WinMinimizeAll WinMinimizeAllUndo 
WinMove WinMoveBottom WinMoveTop WinRedraw WinRestore WinSetAlwaysOnTop WinSetEnabled WinSetRegion WinSetStyle WinSetExStyle 
WinSetTitle WinSetTransColor WinSetTransparent WinShow WinWait WinWaitActive WinWaitNotActive WinWaitClose Require
)"

ENUM_NUMBER_TYPES := "
(
A_ControlDelay,A_DefaultMouseSpeed,A_IconHidden,A_Is64bitOS,A_IsAdmin,A_IsCritical,A_IsPaused,
A_IsSuspended,A_KeyDelay,A_KeyDelayPlay,A_MouseDelay,A_MouseDelayPlay,A_PtrSize,A_ScreenDPI,A_ScreenHeight,A_ScreenWidth,A_TickCount,
A_TimeIdle,A_TimeIdlePhysical,A_TimeSincePriorHotkey,A_TimeSinceThisHotkey,A_WinDelay
)"

ENUM_REG_NAMES := "
(
HKEY_LOCAL_MACHINE,HKLM
HKEY_USERS,HKU
HKEY_CURRENT_USER,HKCU
HKEY_CLASSES_ROOT,HKCR
HKEY_CURRENT_CONFIG,HKCC
)"

WIDTH := 800
HEIGHT := 600

; __________________________________________________________________________________________________
; Global variables
global enumVariables := enum(ENUM_VARIABLES)
global enumFunctions := enum(ENUM_FUNCTIONS)
global enumNumberTypes := enum(ENUM_NUMBER_TYPES)
global enumRegNames := enum(ENUM_REG_NAMES)
global closures := {}	; a map that links a string identifier to a js native closure. Used by "trigger()"
global JS ; a JavaScript helper object. Used extensively by the API functions.
global window ; a shortcut to wb.document.parentWindow. Used by "_Require()".
global mainDir


; __________________________________________________________________________________________________
; Set working dir
SplitPath(mainPath, &mainFilename, &mainDir, &mainExt, &mainSeed)
SetWorkingDir(mainDir)  ; From now on, all relative paths are based on the JS location
mainDirURI := "file:///" . RegExReplace(mainDir, "\\", "/") . "/"
OnExit(LabelOnExit)


; __________________________________________________________________________________________________
; Open a virtual HTML file.
; We want the latest IE documentMode, so we use "X-UA-Compatible".
; If we were to hook into the page while the page is loading, we would encounter race conditions.
; To avoid this, we need to carefully manage the document with following flow:
;					open > write > exec hook > exec main > close
; When opening the document in IE8, the documentMode is lost and so we have to write it again (thus,
; it might also work if we used "about:blank", but let's err on the safe side).
; Further reading:
; 	● http://ahkscript.org/boards/viewtopic.php?t=5714&p=33477#p33477
; 	● http://ahkscript.org/boards/viewtopic.php?f=14&t=5778
ExoGui := Gui()

ExoGui.OnEvent('Close', GuiClose)

ogcActiveXwb := ExoGui.Add("ActiveX", "w" . WIDTH . " h" . HEIGHT . " x0 y0 vwb", "Shell.Explorer")
wb := ogcActiveXwb.Value
wb.Navigate("about:<!DOCTYPE html><meta http-equiv='X-UA-Compatible' content='IE=edge'>")
while wb.readyState < 4
	Sleep(10)
wb.document.open() ; important
document := wb.document ; shortcut
window := document.parentWindow ; shortcut
html := 
(
"<!DOCTYPE html>
<html>
	<head>
		<meta http-equiv='X-UA-Compatible' content='IE=edge'>
		<meta charset='utf-8'>
		<base href='" mainDirURI "'>
	</head>
	<body>
	</body>
</html>"
)
document.write(html)



; __________________________________________________________________________________________________
; Create a virtual helper
if (false) { ; TODO: implement the "#UseExoScope" directive
	scope := "Exo"
} else {
	scope := "window"
}
exoHelper .= 
(
"	var __ExoHelper = {
		Object: function() {
			var obj = {};
			for (var i = 0; i < arguments.length; i += 2) {
				obj[arguments[i]] = arguments[i+1];
			}
			return obj;
		},
		Array: function() {
			return Array.prototype.slice.call(arguments);
		},
		
		
		registerFunction: function(name, func) {
			" scope "[name] = func;
		},
		registerGetter: function(name, func){
			Object.defineProperty(" scope ", name, {
				get: function() {
					alert(func.Name + '`\n' + name);
					return func(name);
				}
			});
		},
		registerAccessor: function(name, func){
			Object.defineProperty(" scope ", name, {
				get: function(){
					return func();
				},
				set: function(value){
					func(value);
				}
			});
		}
	};"
)
; "eval" doesn't work correctly in IE8 at this point.
; If it did, we would have used an anonymous helper object and avoided polluting the global scope.
window.execScript(exoHelper)
JS := window.__ExoHelper


; __________________________________________________________________________________________________
; Register the API
getBuiltInVarReference := getBuiltInVar
for key in StrSplit(RegExReplace(ENUM_VARIABLES, "\n", " "), " ") {
	apiFunction := "_" . key
	if IsSet(%apiFunction%)
		apiFunction := %apiFunction%
	else  ; found no custom API function
		apiFunction := getBuiltInVarReference ; this must be a normal built-in variable
	JS.registerGetter(key, apiFunction)
}
for key in StrSplit(RegExReplace(ENUM_FUNCTIONS, "\n", " "), " ") {
	apiFunction := "_" . key
	if IsSet(%apiFunction%)
		apiFunction := %apiFunction%
	else ; found no custom API function
		apiFunction := %key% ; this must be a normal built-in function
	JS.registerFunction(key, apiFunction)
}
JS.registerAccessor("Clipboard", _A_Clipboard) ; exception: Clipboard is read-write

; __________________________________________________________________________________________________
; Evaluate the main JS.
; Note that there are 4 ways to add the main JS script:
;		1) <script src='...'></script>
;			Works ok, but errors in JS are reported as havine Line Number:0 (most of the times...)
;		2) <script>...</script>
;			Works ok, but errors in JS are reporting with a Char Number offseted by the "<script>" tag
;		3) eval(...)
;			Works ok and doesn't dirty the page source, but you have no script URL (mostly harmless).
;		4) execScript(...)
;			Same as eval, plus a bit safer for IE8. This is the chosen method.
mainContent := FileRead(mainPath)
try {
	window.execScript(mainContent)
} catch {
	end("There was an error running the script")
}
document.close()	; allows the triggering of window.onload


; __________________________________________________________________________________________________
; Finish the Auto-execute Section
trigger("OnClipboardChange")
OnMessage(0x100, WB_onKey, 2) ; support for key down
OnMessage(0x101, WB_onKey, 2) ; support for key up
ExoGui.Show("w" WIDTH " h" HEIGHT)
return


;###################################################################################################
;######################################   F U N C T I O N S   ######################################
;###################################################################################################

;ESC::
;ExitApp


/**
 *
 */
getBuiltInVar(name){
	if (enumNumberTypes.HasOwnProp(name)) {
		return %name%+0
	} else {
		return %name%
	}
}


/**
 * Wrapper for SKAN's function (see below)
 */
getArgs(){
	CmdLine := DllCall( "GetCommandLine", "Str" )
	CmdLine := RegExReplace(CmdLine, " /ErrorStdOut", "")
	Skip    := ( A_IsCompiled ? 1 : 2 )
	argv    := Args( CmdLine, Skip )
	return argv
}


/**
 * By SKAN,  http://goo.gl/JfMNpN,  CD:23/Aug/2014 | MD:24/Aug/2014
 */
Args( CmdLine := "", Skip := 0 ) {    
	Local pArgs := 0, nArgs := 0, A := []
	pArgs := DllCall("Shell32\CommandLineToArgvW", "WStr", CmdLine, "PtrP", &nArgs, "Ptr")
	Loop ( nArgs )
		If ( A_Index > Skip )
			A.Push(StrGet( NumGet(( A_Index - 1 ) * A_PtrSize + pArgs, "UPtr"), "UTF-16")) 
	A.InsertAt(1, nArgs - Skip)
	Return A
}


/**
 * 
 */
trigger(key, args*){
	try closure := closures[key]
	if IsSet(closure) {
		return closure.call(0,args*)
	}
}


/**
 *
 */
enum(blob){
	blob := RegExReplace(blob, "\s+", ",")
	parts := StrSplit(blob, ",")
	output := Map()
	for key, val in parts {
		output[val] := 1
	}
	return output
}

/**
 * Converts relative path to absolute
 * TODO: Support when ahk file is not in terminal dir
 */
rel2AbsPath(lPath) {
    Path := A_WorkingDir "\" lPath
    c := 1
    While (c != 0)
        Path := RegExReplace(Path, "\w+\\\.\.\\", "", &c)
    Path := RegExReplace(Path, "\\\.\\", "\")
	return Path
}


/**
 * Used only for debugging. It's a lightweight JSON.stringify, but with no quotes and no escapes.
 */
trace(s){
	s := serialize(s,0)
	MsgBox(s,, 262144) ; Always on top
}
serialize(obj,indent){
	if (IsObject(obj)) {
		prefix := ""
		Loop indent
		{
			prefix .= "`t"
		}
		out := "`n"
		out .= prefix . "{`n"
		for key, val in obj {
			out .= prefix . "`t" . key . ":" . serialize(val, indent+1) . "`n"
		}
		out .= prefix . "}"
		return out
	} else {
		return obj
	}
}


/**
 *
 */
end(message){
	message .= "`nThe application will now exit."
	MsgBox message
	ExitApp
}


;###################################################################################################
;#########################################   L A B E L S   #########################################
;###################################################################################################


/**
 * Built-in label.
 */
GuiClose(*) {
	if (trigger("GuiClose") == 0) {
		return
	}
	ExitApp
}


/**
 * "OnExit()" redirects here.
 */
LabelOnExit(*) {
	if (trigger("OnExit") == 0) {
		return
	}
	ExitApp
}


/**
 * "Hotkey()" redirects here.
 */
LabelHotkey() {
	trigger("Hotkey" . A_ThisHotkey)
}


/**
 * "OnMessage()" redirects here. Not actually a label, but close enough.
 */
OnMessageClosure(wParam, lParam, msg, hwnd){
	trigger("OnMessage" . msg, wParam, lParam, msg, hwnd)
}


;###################################################################################################
;############################################   A P I   ############################################
;###################################################################################################
#Include <FileObject>
#Include <WB_onKey>
#Include <API>