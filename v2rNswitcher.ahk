; 作者 -Z-
; 日期 2022.9.24  趋近完善
; 地址：https://github.com/921j/ICO/blob/main/v2rNswitcher.ahk
; 引用、转载、修改时请勿去掉源地址，谢谢！
; 获取权限
Loop, %0% ; For each parameter:
  {
    param := %A_Index%  ; Fetch the contents of the variable whose name is contained in A_Index.
    params .= A_Space . param
  }
ShellExecute := A_IsUnicode ? "shell32\ShellExecute":"shell32\ShellExecuteA"
if not A_IsAdmin
{
    If A_IsCompiled
       DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_ScriptFullPath, str, params , str, A_WorkingDir, int, 1)
    Else
       DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_AhkPath, str, """" . A_ScriptFullPath . """" . A_Space . params, str, A_WorkingDir, int, 1)
    ExitApp
}
#NoTrayIcon
#InstallKeybdHook
#InstallMouseHook
#Persistent
#MaxMem 4
#NoEnv
#SingleInstance Force
#MaxHotkeysPerInterval 1000
#KeyHistory 0
;#WinActivateForce
;Process Priority,,High
ListLines, Off
SetWorkingDir, %A_ScriptDir%
SetCapsLockState, AlwaysOff
SetBatchLines, -1
SetKeyDelay, -1, -1
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
SetWinDelay, 0
SetControlDelay, 0
CoordMode, Menu, Window
SendMode, Input
DetectHiddenWindows, on
SetTitleMatchMode, 2
EnvGet, COMMANDER_PATH, COMMANDER_PATH


; 隐藏指定程序托盘图标
HideTrayApps := ["Quicker.exe"]  ; 这里加入想隐藏托盘图标的进程名称
for Index, tIcon in HideTrayApps
  {
    hideIcon := TrayIcon_GetInfo(tIcon)
    Loop, % hideIcon.MaxIndex()
      TrayIcon_Remove(hideIcon[A_Index].hWnd,hideIcon[A_Index].uID)
  }



;任务栏鼠标滚轮调节音量大小
#If MouseIsOver("ahk_class Shell_TrayWnd")
WheelUp::
  SoundSet +5, MASTER
  SoundGet, i
  Tip("当前音量：" Round(i))
  Return
WheelDown::
  SoundSet -5, MASTER
  SoundGet, i
  Tip("当前音量：" Round(i))
  Return
#If

MouseIsOver(WinTitle) {
    MouseGetPos,,, Win
    return WinExist(WinTitle . " ahk_id " . Win)
}

Tip(s:="") {
  SetTimer, %A_ThisFunc%, % s="" ? "Off" : -2000
  ToolTip, %s%
}


;v2rayN 自动化操作
^F1::
	Enabled := ComObjError(0)
	Process, Exist, v2rayN.exe
	v2rayNPid = %ErrorLevel%
	IfWinNotExist, ahk_pid %v2rayNPid%
	{
		Run "D:\zNet\v2rayN-Core\v2rayN.exe",,Hide  ; 这里替换 v2rayN 程序路径
		sleep 3000
		MsgBox, 48, , 看起来 v2rayN 并未运行 已为您运行 v2rayN.`n若想继续测速请再按一次 Ctrl+F1., 5
    return
	}
	Else
	{
		IfWinActive, ahk_pid %v2rayNPid%
		{
			WinHide, ahk_pid %v2rayNPid%
		}
		Else
		{
			v2rayNCls := "ahk_class WindowsForms10.Window.8.app.0.12ab327_r6_ad1"
      Winshow, % v2rayNCls
      Winactivate, % v2rayNCls
      WinWaitActive, % v2rayNCls
			wingetpos, x, y, w, h, % v2rayNCls  ; 首次启动会自动调整 v2rayN 的窗口位置与尺寸
			;msgbox, %x%`n%y%`n%w%`n%h%
			if (%h% != 1088) {
				WinMove, % v2rayNCls,,842,28,1088,  ; 这里修改成自己想要的位置和尺寸
			}
			WinGet, hWnd, ID, % v2rayNCls
			oAcc := Acc_Get("Object", "4.3.4.1.4.1.4.1.4.1.4.1.4", 0, "ahk_id " hWnd)
			oAcc.AccDoDefaultAction(1)
			send, ^a
			send, ^t
			sleep 1000
      Loop
      {
        if (oAcc.accName(A_Index) = "")
				{
          LastIdx := A_Index - 1
          ;MsgBox 节点数  %LastIdx%
          goto WaitRes
				}
			}
      WaitRes:
      Loop
      {
        result := ""
        result .= oAcc.accDescription(LastIdx)
        if ((Sift_Regex(result, "测试结果","OC") != "") && (Sift_Regex(result, "∞|测速中","REGEX") = ""))
        {
					sleep 2000
          goto GetVal
        }
      }
      GetVal:
      Loop %LastIdx%
      {
        result := ""
        result .= oAcc.accDescription(A_Index)
        if (Sift_Regex(result, "\d+\.\d\s+M\/s$","REGEX") != "")
				{
					result1 := Sift_Regex(result, "\d+\.\d\s+M\/s$","REGEX")
					RegExMatch(result1, "\d+\.\d(?=\s+M\/s$)", vnumb)
					if (vnumb > maxv){
						maxv := vnumb
						maxi := A_Index
					}
          ;MsgBox  当前编号:%A_Index%  节点速度:%vnumb% `n选中编号:%maxi%  节点速度:%maxv%
				}
      }
			sleep 200
			Acc_Get("DoAction", "4.3.4.1.4.1.4.1.4.1.4.1.4", maxi, "ahk_id " hWnd)
			send {enter}
			oAcc := oRect := ""
		  Return
		}
	}
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;托盘图标操作函数;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

TrayIcon_GetInfo(sExeName := "")
{
    DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
    oTrayIcon_GetInfo := {}
    For key, sTray in ["Shell_TrayWnd", "NotifyIconOverflowWindow"]
    {
        idxTB := TrayIcon_GetTrayBar(sTray)
        WinGet, pidTaskbar, PID, ahk_class %sTray%

        hProc := DllCall("OpenProcess", UInt, 0x38, Int, 0, UInt, pidTaskbar)
        pRB   := DllCall("VirtualAllocEx", Ptr, hProc, Ptr, 0, UPtr, 20, UInt, 0x1000, UInt, 0x4)

        SendMessage, 0x418, 0, 0, ToolbarWindow32%idxTB%, ahk_class %sTray%   ; TB_BUTTONCOUNT

        szBtn := VarSetCapacity(btn, (A_Is64bitOS ? 32 : 20), 0)
        szNfo := VarSetCapacity(nfo, (A_Is64bitOS ? 32 : 24), 0)
        szTip := VarSetCapacity(tip, 128 * 2, 0)

        Loop, %ErrorLevel%
        {
            SendMessage, 0x417, A_Index - 1, pRB, ToolbarWindow32%idxTB%, ahk_class %sTray%   ; TB_GETBUTTON
            DllCall("ReadProcessMemory", Ptr, hProc, Ptr, pRB, Ptr, &btn, UPtr, szBtn, UPtr, 0)

            iBitmap := NumGet(btn, 0, "Int")
            IDcmd   := NumGet(btn, 4, "Int")
            statyle := NumGet(btn, 8)
            dwData  := NumGet(btn, (A_Is64bitOS ? 16 : 12))
            iString := NumGet(btn, (A_Is64bitOS ? 24 : 16), "Ptr")

            DllCall("ReadProcessMemory", Ptr, hProc, Ptr, dwData, Ptr, &nfo, UPtr, szNfo, UPtr, 0)

            hWnd  := NumGet(nfo, 0, "Ptr")
            uID   := NumGet(nfo, (A_Is64bitOS ? 8 : 4), "UInt")
            msgID := NumGet(nfo, (A_Is64bitOS ? 12 : 8))
            hIcon := NumGet(nfo, (A_Is64bitOS ? 24 : 20), "Ptr")

            WinGet, pID, PID, ahk_id %hWnd%
            WinGet, sProcess, ProcessName, ahk_id %hWnd%
            WinGetClass, sClass, ahk_id %hWnd%

            If !sExeName || (sExeName = sProcess) || (sExeName = pID)
            {
                DllCall("ReadProcessMemory", Ptr, hProc, Ptr, iString, Ptr, &tip, UPtr, szTip, UPtr, 0)
                Index := (oTrayIcon_GetInfo.MaxIndex()>0 ? oTrayIcon_GetInfo.MaxIndex()+1 : 1)
                oTrayIcon_GetInfo[Index,"idx"]     := A_Index - 1
                oTrayIcon_GetInfo[Index,"IDcmd"]   := IDcmd
                oTrayIcon_GetInfo[Index,"pID"]     := pID
                oTrayIcon_GetInfo[Index,"uID"]     := uID
                oTrayIcon_GetInfo[Index,"msgID"]   := msgID
                oTrayIcon_GetInfo[Index,"hIcon"]   := hIcon
                oTrayIcon_GetInfo[Index,"hWnd"]    := hWnd
                oTrayIcon_GetInfo[Index,"Class"]   := sClass
                oTrayIcon_GetInfo[Index,"Process"] := sProcess
                oTrayIcon_GetInfo[Index,"Tooltip"] := StrGet(&tip, "UTF-16")
                oTrayIcon_GetInfo[Index,"Tray"]    := sTray
            }
        }
        DllCall("VirtualFreeEx", Ptr, hProc, Ptr, pRB, UPtr, 0, Uint, 0x8000)
        DllCall("CloseHandle", Ptr, hProc)
    }
    DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
    Return oTrayIcon_GetInfo
}

TrayIcon_Remove(hWnd, uID)
{
        NumPut(VarSetCapacity(NID,(A_IsUnicode ? 2 : 1) * 384 + A_PtrSize * 5 + 40,0), NID)
        NumPut(hWnd , NID, (A_PtrSize == 4 ? 4 : 8 ))
        NumPut(uID  , NID, (A_PtrSize == 4 ? 8  : 16 ))
        Return DllCall("shell32\Shell_NotifyIcon", "Uint", 0x2, "Uint", &NID)
}

TrayIcon_Set(hWnd, uId, hIcon, hIconSmall:=0, hIconBig:=0)
{
    d := A_DetectHiddenWindows
    DetectHiddenWindows, On
    ; WM_SETICON = 0x0080
    If ( hIconSmall )
        SendMessage, 0x0080, 0, hIconSmall,, ahk_id %hWnd%
    If ( hIconBig )
        SendMessage, 0x0080, 1, hIconBig,, ahk_id %hWnd%
    DetectHiddenWindows, %d%

    VarSetCapacity(NID, szNID := ((A_IsUnicode ? 2 : 1) * 384 + A_PtrSize*5 + 40),0)
    NumPut( szNID, NID, 0                           )
    NumPut( hWnd,  NID, (A_PtrSize == 4) ? 4   : 8  )
    NumPut( uId,   NID, (A_PtrSize == 4) ? 8   : 16 )
    NumPut( 2,     NID, (A_PtrSize == 4) ? 12  : 20 )
    NumPut( hIcon, NID, (A_PtrSize == 4) ? 20  : 32 )

    ; NIM_MODIFY := 0x1
    Return DllCall("Shell32.dll\Shell_NotifyIcon", UInt,0x1, Ptr,&NID)
}

TrayIcon_GetTrayBar(Tray:="Shell_TrayWnd")
{
    DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
    WinGet, ControlList, ControlList, ahk_class %Tray%
    RegExMatch(ControlList, "(?<=ToolbarWindow32)\d+(?!.*ToolbarWindow32)", nTB)
    Loop, %nTB%
    {
        ControlGet, hWnd, hWnd,, ToolbarWindow32%A_Index%, ahk_class %Tray%
        hParent := DllCall( "GetParent", Ptr, hWnd )
        WinGetClass, sClass, ahk_id %hParent%
        If !(sClass = "SysPager" or sClass = "NotifyIconOverflowWindow" )
            Continue
        idxTB := A_Index
        Break
    }
    DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
    Return  idxTB
}

TrayIcon_Button(sExeName, sButton := "L", bDouble := false, index := 1)
{
    DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
    WM_MOUSEMOVE      = 0x0200
    WM_LBUTTONDOWN    = 0x0201
    WM_LBUTTONUP      = 0x0202
    WM_LBUTTONDBLCLK  = 0x0203
    WM_RBUTTONDOWN    = 0x0204
    WM_RBUTTONUP      = 0x0205
    WM_RBUTTONDBLCLK  = 0x0206
    WM_MBUTTONDOWN    = 0x0207
    WM_MBUTTONUP      = 0x0208
    WM_MBUTTONDBLCLK  = 0x0209
    sButton := "WM_" sButton "BUTTON"
    oIcons := {}
    oIcons := TrayIcon_GetInfo(sExeName)
    msgID  := oIcons[index].msgID
    uID    := oIcons[index].uID
    hWnd   := oIcons[index].hWnd
    if bDouble
        PostMessage, msgID, uID, %sButton%DBLCLK, , ahk_id %hWnd%
    else
    {
        PostMessage, msgID, uID, %sButton%DOWN, , ahk_id %hWnd%
        PostMessage, msgID, uID, %sButton%UP, , ahk_id %hWnd%
    }
    DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
    return
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; ACC1 操作函数;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Version: 2022.07.01.1
; https://gist.github.com/58d2b141be2608a2f7d03a982e552a71

; Private
Acc_Init(Function) {
    static hModule := DllCall("Kernel32\LoadLibrary", "Str","oleacc.dll", "Ptr")
    return DllCall("Kernel32\GetProcAddress", "Ptr",hModule, "AStr",Function, "Ptr")
}

Acc_ObjectFromEvent(ByRef ChildIdOut, hWnd, ObjectId, ChildId) {
    static address := Acc_Init("AccessibleObjectFromEvent")
    VarSetCapacity(varChild, A_PtrSize * 2 + 8, 0)
    hResult := DllCall(address, "Ptr",hWnd, "UInt",ObjectId, "UInt",ChildId
        , "Ptr*",pAcc:=0, "Ptr",&varChild)
    if (!hResult) {
        ChildIdOut := NumGet(varChild, 8, "UInt")
        return ComObj(9, pAcc, 1)
    }
}

Acc_ObjectFromPoint(ByRef ChildIdOut := "", x := 0, y := 0) {
    static address := Acc_Init("AccessibleObjectFromPoint")
    try
        point := x & 0xFFFFFFFF | y << 32
    catch
        DllCall("User32\GetCursorPos", "Int64*",point:=0)
    VarSetCapacity(varChild, A_PtrSize * 2 + 8, 0)
    hResult := DllCall(address, "Int64",point, "Ptr*",pAcc:=0, "Ptr",&varChild)
    if (!hResult) {
        ChildIdOut := NumGet(varChild, 8, "UInt")
        return ComObj(9, pAcc, 1)
    }
}

/* ObjectId
    0xFFFFFFFF = OBJID_SYSMENU
    0xFFFFFFFE = OBJID_TITLEBAR
    0xFFFFFFFD = OBJID_MENU
    0xFFFFFFFC = OBJID_CLIENT*
    0xFFFFFFFB = OBJID_VSCROLL
    0xFFFFFFFA = OBJID_HSCROLL
    0xFFFFFFF9 = OBJID_SIZEGRIP
    0xFFFFFFF8 = OBJID_CARET
    0xFFFFFFF7 = OBJID_CURSOR
    0xFFFFFFF6 = OBJID_ALERT
    0xFFFFFFF5 = OBJID_SOUND
    0xFFFFFFF4 = OBJID_QUERYCLASSNAMEIDX
    0xFFFFFFF0 = OBJID_NATIVEOM
    0x00000000 = OBJID_WINDOW
*/
Acc_ObjectFromWindow(hWnd, ObjectId := -4) {
    static address := Acc_Init("AccessibleObjectFromWindow")
    ObjectId &= 0xFFFFFFFF
    VarSetCapacity(IID, 16, 0)
    addr := ObjectId = 0xFFFFFFF0 ? 0x0000000000020400 : 0x11CF3C3D618736E0
    rIID := NumPut(addr, IID, "Int64")
    addr := ObjectId = 0xFFFFFFF0 ? 0x46000000000000C0 : 0x719B3800AA000C81
    rIID := NumPut(addr, rIID + 0, "Int64") - 16
    hResult := DllCall(address, "Ptr",hWnd, "UInt",ObjectId, "Ptr",rIID, "Ptr*",pAcc:=0)
    if (!hResult)
        return ComObj(9, pAcc, 1)
}

Acc_WindowFromObject(pAcc) {
    static address := Acc_Init("WindowFromAccessibleObject")
    if (IsObject(pAcc))
        pAcc := ComObjValue(pAcc)
    hResult := DllCall(address, "Ptr",pAcc, "Ptr*",hWnd:=0)
    if (!hResult)
        return hWnd
}

Acc_GetRoleText(nRole) {
    static address := Acc_Init("GetRoleTextW")
    nSize := DllCall(address, "UInt",nRole, "Ptr",0, "UInt",0)
    VarSetCapacity(sRole, ++nSize, 0)
    DllCall(address, "UInt",nRole, "Ptr",&sRole, "UInt",nSize)
    return StrGet(&sRole)
}

Acc_GetStateText(nState) {
    static address := Acc_Init("GetStateTextW")
    nSize := DllCall(address, "UInt",nState, "Ptr",0, "UInt",0)
    VarSetCapacity(sState, ++nSize, 0)
    DllCall(address, "UInt",nState, "Ptr",&sState, "UInt",nSize)
    return StrGet(&sState)
}

Acc_SetWinEventHook(EventMin, EventMax, Callback) {
    return DllCall("User32\SetWinEventHook", "UInt",EventMin, "UInt",EventMax
        , "Ptr",0, "Ptr",Callback, "UInt",0, "UInt",0, "UInt",0)
}

Acc_UnhookWinEvent(hHook) {
    return DllCall("User32\UnhookWinEvent", "Ptr",hHook)
}

/* Win Events
Callback := RegisterCallback("WinEventProc")
WinEventProc(hHook, Event, hWnd, ObjectId, ChildId, EventThread, EventTime) {
    Critical
    oAcc := Acc_ObjectFromEvent(ChildIdOut, hWnd, ObjectId, ChildId)
    ; Code Here
}
*/

Acc_Role(oAcc, ChildId := 0) {
    try {
        role := oAcc.accRole(ChildId + 0)
        return Acc_GetRoleText(role)
    }
    return "invalid object"
}

Acc_State(oAcc, ChildId := 0) {
    try {
        state := oAcc.accState(ChildId + 0)
        return Acc_GetStateText(state)
    }
    return "invalid object"
}

Acc_Location(oAcc, ChildId := 0, ByRef Position := "") {
    x := 0, w := 0
    y := 0, h := 0
    ; 0x4003 = VT_BYREF | VT_I4
    x := ComObject(0x4003, &x), w := ComObject(0x4003, &w)
    y := ComObject(0x4003, &y), h := ComObject(0x4003, &h)
    try {
        oAcc.accLocation(x, y, w, h, ChildId + 0)
        x := NumGet(x, 0, "Int"), w := NumGet(w, 0, "Int")
        y := NumGet(y, 0, "Int"), h := NumGet(h, 0, "Int")
        Position := "x" x " y" y " w" w " h" h
        return {"x":x, "y":y, "w":w, "h":h, "pos":Position}
    }
}

Acc_Parent(oAcc) {
    try {
        if (oAcc.accParent)
            return Acc_Query(oAcc.accParent)
    }
}

Acc_Child(oAcc, ChildId := 0) {
    try {
        Child := oAcc.AccChild(ChildId + 0)
        if (Child)
            return Acc_Query(Child)
    }
}

; Private
Acc_Query(oAcc) {
    try {
        query := ComObjQuery(oAcc, "{618736E0-3C3D-11CF-810C-00AA00389B71}")
        return ComObj(9, query, 1)
    }
}

; Private, deprecated
Acc_Error(Previous := "") {
    static setting := 0
    if (StrLen(Previous))
        setting := Previous
    return setting
}

Acc_Children(oAcc) {
    static address := Acc_Init("AccessibleChildren")
    if (ComObjType(oAcc, "Name") != "IAccessible")
        throw Exception("Invalid IAccessible Object", -1, oAcc)
    pAcc := ComObjValue(oAcc)
    size := A_PtrSize * 2 + 8
    VarSetCapacity(varChildren, oAcc.accChildCount * size, 0)
    hResult := DllCall(address, "Ptr",pAcc, "Int",0, "Int",oAcc.accChildCount
        , "Ptr",&varChildren, "Int*",obtained:=0)
    if (hResult)
        throw Exception("AccessibleChildren DllCall Failed", -1)
    children := []
    loop % obtained {
        i := (A_Index - 1) * size
        child := NumGet(varChildren, i + 8, "Int64")
        if (NumGet(varChildren, i, "Int64") = 9) {
            accChild := Acc_Query(child)
            children.Push(accChild)
            ObjRelease(child)
        } else {
            children.Push(child)
        }
    }
    return children
}

Acc_ChildrenByRole(oAcc, RoleText) {
    children := []
    for _,child in Acc_Children(oAcc) {
        if (Acc_Role(child) = RoleText)
            children.Push(child)
    }
    return children
}

/* Commands
    - Aliases:
    Action → DefaultAction
    DoAction → DoDefaultAction
    Keyboard → KeyboardShortcut
    - Properties:
    Child
    ChildCount
    DefaultAction
    Description
    Focus
    Help
    HelpTopic
    KeyboardShortcut
    Name
    Parent
    Role
    Selection
    State
    Value
    - Methods:
    DoDefaultAction
    Location
    - Other:
    Object
*/
Acc_Get(Command, ChildPath, ChildId := 0, Target*) {
    if (Command ~= "i)^(?:HitTest|Navigate|Select)$")
        throw Exception("Command not implemented", -1, Command)
    ChildPath := StrReplace(ChildPath, "_", " ")
    if (IsObject(Target[1])) {
        oAcc := Target[1]
    } else {
        hWnd := WinExist(Target*)
        oAcc := Acc_ObjectFromWindow(hWnd, 0)
    }
    if (ComObjType(oAcc, "Name") != "IAccessible")
        throw Exception("Cannot access an IAccessible Object", -1, oAcc)
    ChildPath := StrSplit(ChildPath, ".")
    for level,item in ChildPath {
        RegExMatch(item, "OS)(?<Role>\D+)(?<Index>\d*)", match)
        if (match) {
            item := match.Index ? match.Index : 1
            children := Acc_ChildrenByRole(oAcc, match.Role)
        } else {
            children := Acc_Children(oAcc)
        }
        if (children.HasKey(item)) {
            oAcc := children[item]
            continue
        }
        extra := match.Role
            ? "Role: " match.Role ", Index: " item
            : "Item: " item ", Level: " level
        throw Exception("Cannot access ChildPath Item", -1, extra)
    }
    switch (Command) {
        case "Action": Command := "DefaultAction"
        case "DoAction": Command := "DoDefaultAction"
        case "Keyboard": Command := "KeyboardShortcut"
        case "Object":
            return oAcc
    }
    switch (Command) {
        case "Location":
            out := Acc_Location(oAcc, ChildId).pos
        case "Parent":
            out := Acc_Parent(oAcc)
        case "Role", "State":
            out := Acc_%Command%(oAcc, ChildId)
        case "ChildCount", "Focus", "Selection":
            out := oAcc["acc" Command]
        default:
            out := oAcc["acc" Command](ChildId + 0)
    }
    return out
}



Sift_Regex(ByRef Haystack, ByRef Needle, Options := "IN", Delimit := "`n")
{
	Sifted := {}
	if (Options = "IN")		
		Needle_Temp := "\Q" Needle "\E"
	else if (Options = "LEFT")
		Needle_Temp := "^\Q" Needle "\E"
	else if (Options = "RIGHT")
		Needle_Temp := "\Q" Needle "\E$"
	else if (Options = "EXACT")		
		Needle_Temp := "^\Q" Needle "\E$"
	else if (Options = "REGEX")
		Needle_Temp := Needle
	else if (Options = "OC")
		Needle_Temp := RegExReplace(Needle,"(.)","\Q$1\E.*")
	else if (Options = "OW")
		Needle_Temp := RegExReplace(Needle,"( )","\Q$1\E.*")
	else if (Options = "UW")
		Loop, Parse, Needle, " "
			Needle_Temp .= "(?=.*\Q" A_LoopField "\E)"
	else if (Options = "UC")
		Loop, Parse, Needle
			Needle_Temp .= "(?=.*\Q" A_LoopField "\E)"

	if Options is lower
		Needle_Temp := "i)" Needle_Temp
	
	if IsObject(Haystack)
	{
		for key, Hay in Haystack
			if RegExMatch(Hay, Needle_Temp)
				Sifted.Insert(Hay)
	}
	else
	{
		Loop, Parse, Haystack, %Delimit%
			if RegExMatch(A_LoopField, Needle_Temp)
				Sifted .= A_LoopField Delimit
		Sifted := SubStr(Sifted,1,-1)
	}
	return Sifted
}
