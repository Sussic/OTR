; =============================================================
; EVE Multibox Tiler — AutoHotkey v1 (Unicode)
; -------------------------------------------------------------
; LOG FILE: %TEMP%\EVE_Tiler.log
; Hotkeys:
;   Ctrl+Win+R  → pick region (click–click) → choose clients → tile
;   Ctrl+Win+L  → re-tile last selection
;   Ctrl+Win+Q  → close all clones
;   Ctrl+Win+Y  → DEBUG: clone full client (no region)
;   Ctrl+Win+G  → open Groups Manager
;   Ctrl+Win+T  → tile currently selected group
;   Ctrl+Win+N  → quick new group
; =============================================================

#NoEnv
#SingleInstance Force
#Warn All, Off
SetBatchLines, -1
ListLines, Off

; -------------------- CONFIG --------------------
OTR_PATH := "C:\Program Files (x86)\OnTopReplica\OnTopReplica\OnTopReplica.exe"
COMMON_FLAGS := "--chromeOff"
TARGET_MONITOR := 1
GAP_X := 6
GAP_Y := 6
MIN_W := 160
MIN_H := 90
IsLaunching := 0

; --- clone sizing ---
USE_REGION_SCALE := true
REGION_SCALE     := 1.0

; --- origin quirk (OTR expects TOP-LEFT XYWH; leave false) ---
REGION_FLIP_Y    := false

; --- selection safety ---
SEL_MIN_W := 80
SEL_MIN_H := 60

; --- debugging ---
DEBUG := false
LOG_PATH := A_Temp "\EVE_Tiler.log"

OnExit, _Cleanup   ; do NOT return here!

; -------------------- STATE (original) ---------------------
LAST_REGION := ""            ; "x,y,w,h" (client-relative)
LAST_SELECTION := []
LAST_HWND := 0
launchedPIDs := []
Picker_Clients := []
Picker_hLV := 0
LABEL_COLOR   := "FFFF00"  ; default yellow
LABEL_OFF_X   := 8
LABEL_OFF_Y   := 6
SetTitleMatchMode, 3   ; exact window titles only
DEBUG_VERBOSE := 0

; ============================================================
;                      NEW: GROUPS/PERSIST
; ============================================================

; Groups model: map groupName -> object
; g := { members:[], region:"x,y,w,h", layout:{mode:"auto"/"grid", cols:0, rows:0}, monitor:int, gapX:int, gapY:int, scale:float }
Groups := {}
SelectedGroup := ""
Manager_hLV_Groups := 0
Manager_hLV_Members := 0
Manager_hLV_Clients := 0
MGR_Layout := "auto"
MGR_Cols   := 0
MGR_Rows   := 0
MGR_Mon    := TARGET_MONITOR
MGR_GapX   := GAP_X
MGR_GapY   := GAP_Y
MGR_Scale  := REGION_SCALE
GroupPIDs := {}            ; map name -> [pid,...] for launched clones
MGR_Accumulate := 1        ; checkbox: keep existing clones when tiling
HIDE_FROM_ALTTAB := 0      ; 1 = hide clones from Alt-Tab by using WS_EX_TOOLWINDOW
AUTO_RETILE := 1           ; default ON (persisted in INI)
__WATCH_INTERVAL := 500    ; ms, poll rate for client window changes
__LAST_EVE_KEY := ""       ; fingerprint of current EVE windows
__RETIPLE_AT := 0          ; debounce timestamp (ms)

; Defaults when creating a new group
DEF_GROUP := { members: []
            , region: ""
            , layout: { mode: "auto", cols: 0, rows: 0 }
            , monitor: TARGET_MONITOR
            , gapX: GAP_X
            , gapY: GAP_Y
            , scale: REGION_SCALE }

; -------------------- SANITY --------------------
if !FileExist(OTR_PATH) {
    MsgBox, 16, OnTopReplica not found, Not found:`n%OTR_PATH%
    ExitApp
}

; ============================================================
; NEW: GROUPS/PERSIST — resolve CFG_FILE robustly (v1-safe)
; ============================================================
; Look for config.ini beside the script first, then in Documents and AppData.
CFG_FILE := ""
possiblePaths := []                               ; build list without line-continuations
possiblePaths.Push(A_ScriptDir "\config.ini")
possiblePaths.Push(A_MyDocuments "\AutoHotkey\EVE_Tiler\config.ini")
possiblePaths.Push(A_AppData "\EVE_Tiler\config.ini")

for idx, p in possiblePaths {
    if FileExist(p) {
        CFG_FILE := p
        break
    }
}
; Fallback to the script folder if still blank
if (CFG_FILE = "")
    CFG_FILE := A_ScriptDir "\config.ini"

; Make sure the directory for CFG_FILE exists (for SaveConfig)
SplitPath, CFG_FILE,, __cfgDir
if !FileExist(__cfgDir)
    FileCreateDir, %__cfgDir%

; One-time diagnostic so you can see EXACTLY where it’s looking.
; For production, keep this behind DEBUG so it doesn't nag every launch.
if (DEBUG) {
    MsgBox, 64, EVE Tiler (debug), % "CFG_FILE:`n" CFG_FILE "`nExists: " (FileExist(CFG_FILE) ? "YES" : "NO")
}


StartAutoRetile()   ; now runs properly

return               ; <--- THIS is where the auto-execute section should end

_Cleanup:
    CloseLaunchedClones()
    ExitApp
return


; ============================================================
;                        HOTKEYS
; ============================================================
; force Win-combos through the keyboard hook
#UseHook On
#InstallKeybdHook

^#F6::DBG_ShowSummary("Last run")   ; Ctrl+Win+F6 → re-show last summary

^#F5::                              ; Ctrl+Win+F5 → dump current OTR windows
    WinGet, L, List, ahk_exe OnTopReplica.exe
    rep := "OTR windows: " L "`n"
    Loop, %L% {
        h := L%A_Index%
        WinGetTitle, t, ahk_id %h%
        WinGet, pid, PID, ahk_id %h%
        WinGetPos, x,y,w,h2, ahk_id %h%
        rep .= "PID " pid "  hwnd " h "  " t "  [" x "," y "  " w "x" h2 "]`n"
    }
    LogBlock("OTR-LIST", { text: rep })
    MsgBox, 64, OTR list, %rep%
return


^#F9::
    dbg := []
    live := BuildLiveClientMap()
    for name, g in Groups {
        for _, saved in g.members {
            clean := ExtractCharName(saved), lc := clean
            StringLower, lc, lc
            ok := live.HasKey(lc)
            dbg.Push( (ok ? "[OK]   " : "[MISS] ") . saved . "  -> clean=" . clean )
        }
    }
    txt := ""
    for _, line in dbg
        txt .= line . "`n"
    ; write to log and show a compact box
    LogBlock("MEMBER-CHECK", { report: txt })
    MsgBox, 64, Member check, % "Wrote details to log.`nFirst few:`n`n" . SubStr(txt,1,1000)
return

^#F12::  ; Ctrl+Win+F12 — show EXACT paths, open folder, and select the file
txt := "Search order (exists?):`n"
paths := [ A_ScriptDir "\config.ini"
        , A_MyDocuments "\AutoHotkey\EVE_Tiler\config.ini"
        , A_AppData "\EVE_Tiler\config.ini" ]
for i, p in paths
    txt .= i ". " p "  -> " (FileExist(p) ? "YES" : "NO") "`n"

txt .= "`nResolved CFG_FILE:`n" CFG_FILE
txt .= "`nExists now: " (FileExist(CFG_FILE) ? "YES" : "NO")

; show it and copy to clipboard for easy paste
Clipboard := CFG_FILE
MsgBox, 64, EVE Tiler — CFG probe, %txt%`n`n(Full path copied to clipboard.)

; open the folder, and select the file if present
SplitPath, %CFG_FILE%, , _dir
if (FileExist(CFG_FILE))
    Run, explorer.exe /select,"%CFG_FILE%"
else
    Run, explorer.exe "%_dir%"
return


^#F11:: ; list groups currently in memory
names := ""
for name, g in Groups
    names .= (names="" ? "" : "`n") . name
if (names = "")
    names := "(none)"
MsgBox, 64, Groups in memory, %names%
return

^#r::
    __coords := PickRegion_WithGuard()
    if (__coords = "")
        return
    LAST_REGION := __coords
    LogBlock("PICK-FINAL", { region_xywh: __coords })

    ; Ad-hoc flow: show the client picker AFTER the pick
    PickClientsAndLaunch(LAST_REGION)
return


^#p::Pause  ; Ctrl+Win+P toggles Pause
return

^#s::       ; Ctrl+Win+S toggles Suspend (disables hotkeys)
    Suspend, Toggle
    msg := "Hotkeys " . (A_IsSuspended ? "SUSPENDED" : "ACTIVE")
    TrayTip, EVE Tiler, %msg%, 1200, 1
return

^#x::       ; Ctrl+Win+X — close clones (tracked + untracked) and exit
    CloseLaunchedClones()
    CloseAllOTRWindows()
    ExitApp
return


^#l::
    if (LAST_REGION = "") {
        MsgBox, 48, Pick a region first, Press Ctrl+Win+R to select a region.
        return
    }
    ; If a group is selected, use its scale/monitor/gaps
    if (SelectedGroup != "" && Groups.HasKey(SelectedGroup)) {
        LaunchGroup(SelectedGroup)
        return
    }
    ; fallback: last manual selection
    if (LAST_SELECTION.MaxIndex() = "" || LAST_SELECTION.MaxIndex() <= 0) {
        PickClientsAndLaunch(LAST_REGION)
    } else {
        __tiles := []
        for __i, __t in LAST_SELECTION
            __tiles.Insert({ title: __t, region: LAST_REGION })
        LaunchAll(__tiles) ; uses global REGION_SCALE
    }
return

^#F10::
    CloseAllClones()
return

^#y::  ; DEBUG: full clone of active window (no region)
    WinGet, __activeH, ID, A
    if (!__activeH) {
        MsgBox, 48, No active window, Activate the EVE window first.
        return
    }
    WinGetTitle, __tNow, ahk_id %__activeH%
    Run, % Chr(34) . OTR_PATH . Chr(34) . " --windowTitle=" . Chr(34) . __tNow . Chr(34) . " " . COMMON_FLAGS,, Hide
return

; --- NEW: Groups Manager hotkeys ---
^#g::OpenGroupManager() ; open manager
return

^#t:: ; tile currently selected group (from manager)
    if (SelectedGroup = "") {
        TrayTip, EVE Tiler, Open Manager (Ctrl+Win+G) and select a group., 1500, 1
        return
    }
    LaunchGroup(SelectedGroup)
return

^#n:: ; quick new group
    InputBox, __gn, New Group, Enter a new group name:
    if (ErrorLevel || __gn = "")
        return
    if (Groups.HasKey(__gn)) {
        MsgBox, 48, Exists, A group with that name already exists.
        return
    }
    Groups[__gn] := CloneDefaultGroup()
    SelectedGroup := __gn
    SaveConfig()
    OpenGroupManager(true)
return

; ============================================================
;                Detect EVE Clients + Picker GUI
; ============================================================

DetectEVEClients() {
    __clients := []
    WinGet, __idList, List, ahk_exe exefile.exe
    Loop, % __idList {
        __h := __idList%A_Index%
        WinGetTitle, __t, ahk_id %__h%
        if (__t != "")
            __clients.Insert({ hwnd: __h, title: __t })
    }
    if (__clients.MaxIndex() = "") {
        WinGet, __all, List
        Loop, % __all {
            __h := __all%A_Index%
            WinGetTitle, __t, ahk_id %__h%
            if InStr(__t, "EVE")
                __clients.Insert({ hwnd: __h, title: __t })
        }
    }
    return __clients
}

PickClientsAndLaunch(region) {
    global Picker_Clients, Picker_hLV, LAST_SELECTION, LAST_HWND
    __cl := DetectEVEClients()
    __c := __cl.MaxIndex()
    if ((__c = "" || __c = 0) && LAST_HWND) {
        WinGetTitle, __tf, ahk_id %LAST_HWND%
        if (__tf != "")
            __cl.Insert({ hwnd: LAST_HWND, title: __tf })
        __c := __cl.MaxIndex()
    }
    if (__c = "" || __c = 0) {
        MsgBox, 48, No EVE clients, None found. Start an EVE client and try again.
        return
    }
    Picker_Clients := __cl

    Gui, Picker:New, +AlwaysOnTop +Hwnd__hPick, Select EVE Clients
    Gui, Picker:Margin, 10, 10
    Gui, Picker:Add, Text,, Region: %region%
    Gui, Picker:Add, Button, gPicker_SelAll w90, Select All
    Gui, Picker:Add, Button, gPicker_SelNone x+5 w90, None
    Gui, Picker:Add, Button, gPicker_InvSel x+5 w90, Invert
    Gui, Picker:Add, ListView, xm w520 r12 Checked hwnd__hLV, Title
    Picker_hLV := __hLV

    Gui, Picker:Default
    Gui, ListView, %Picker_hLV%
    Loop, % __c {
        __o := __cl[A_Index]
        LV_Add("Check", __o.title)
    }
    LV_ModifyCol(1, 500)
    Gui, Picker:Add, Button, xm w100 Default gPicker_DoOK, OK
    Gui, Picker:Add, Button, x+5 w100 gPicker_DoCancel, Cancel
    Gui, Picker:Show, Center
    WinActivate, ahk_id %__hPick%
    return
}

Picker_SelAll:
Gui, Picker:Default
Gui, ListView, %Picker_hLV%
LV_Modify(0, "+Check")
return

Picker_SelNone:
Gui, Picker:Default
Gui, ListView, %Picker_hLV%
LV_Modify(0, "-Check")
return

Picker_InvSel:
Gui, Picker:Default
Gui, ListView, %Picker_hLV%
__cnt := LV_GetCount()
Loop, %__cnt% {
    if (LV_GetNext(A_Index-1, "C") = A_Index)
        LV_Modify(A_Index, "-Check")
    else
        LV_Modify(A_Index, "+Check")
}
return

Picker_DoCancel:
    Gui, Picker:Destroy
return

Picker_DoOK:
global LAST_SELECTION, LAST_REGION, Picker_Clients, Picker_hLV, IsLaunching
Gui, Picker:Default
Gui, ListView, %Picker_hLV%
__sel := []
__idx := 0
Loop {
    __idx := LV_GetNext(__idx, "C")
    if (!__idx)
        break
    LV_GetText(__t, __idx)
    __sel.Insert(__t)
}
if (__sel.MaxIndex() = "") {
    MsgBox, 48, Nothing selected, Select at least one client.
    return
}
LAST_SELECTION := __sel
Gui, Picker:Destroy

__tiles := []
for __k, __t in __sel
    __tiles.Insert({ title: __t, region: LAST_REGION })

; ---- pause watcher only for the actual launch
if (!IsLaunching) {
    IsLaunching := 1
    StopAutoRetile()
}
LaunchAll(__tiles)  ; auto grid
StartAutoRetile()
IsLaunching := 0
return

; ============================================================
;                           LaunchAll
;                 (now supports cols/rows overrides)
; ============================================================
LaunchAll(__tiles, __colsOverride := 0, __rowsOverride := 0) {
    global TARGET_MONITOR, GAP_X, GAP_Y, MIN_W, MIN_H
    global OTR_PATH, COMMON_FLAGS, launchedPIDs
    global USE_REGION_SCALE, REGION_SCALE, REGION_FLIP_Y, HIDE_FROM_ALTTAB
    global LABEL_COLOR, LABEL_OFF_X, LABEL_OFF_Y

    CloseLaunchedClones()

    __n := __tiles.MaxIndex()
    if (__n = "" || __n = 0)
        return

    _GetMonitorWA(TARGET_MONITOR, __mL, __mT, __mR, __mB)

    __mW := __mR - __mL, __mH := __mB - __mT

    if (__colsOverride > 0 && __rowsOverride > 0)
        __cols := __colsOverride, __rows := __rowsOverride
    else
        __cols := Ceil(Sqrt(__n)), __rows := Ceil((__n + __cols - 1) / __cols)

    __tileW := Floor((__mW - GAP_X * (__cols - 1)) / __cols)
    __tileH := Floor((__mH - GAP_Y * (__rows - 1)) / __rows)
    if (__tileW < MIN_W) __tileW := MIN_W
    if (__tileH < MIN_H) __tileH := MIN_H

    ; -------- Pass 1: fire off ALL processes quickly ----------
    pinfo := []   ; {pid, title}
    __r := 0, __c := 0
    Loop, % __n {
        __i := A_Index
        __title  := __tiles[__i].title
        __region := __tiles[__i].region

        StringSplit, __R, __region, `,
        __rX := __R1, __rY := __R2, __rW := __R3, __rH := __R4

        if (USE_REGION_SCALE)
            __outW := Round(__rW * REGION_SCALE), __outH := Round(__rH * REGION_SCALE)
        else
            __outW := __tileW, __outH := __tileH

        __rY_use := __rY
        if (REGION_FLIP_Y) {
            WinGet, __hSrc, ID, %__title%
            if (__hSrc) {
                VarSetCapacity(__rc2, 16)
                DllCall("GetClientRect","ptr",__hSrc,"ptr",&__rc2)
                __clientH := NumGet(__rc2,12,"Int")
                if (__clientH)
                    __rY_use := Max(0, __clientH - __rY - __rH)
            }
        }

        __x := __mL + __c * (__tileW + GAP_X)
        __y := __mT + __r * (__tileH + GAP_Y)

        __args := "--windowTitle=" . Chr(34) . __title . Chr(34)
        __args .= " " . COMMON_FLAGS
        __args .= " --region=" . __rX "," . __rY_use "," . __rW "," . __rH
        __args .= " --size="   . __outW "," . __outH
        __args .= " --position=" . __x "," . __y

        Run, % Chr(34) . OTR_PATH . Chr(34) . " " . __args,, Hide, __pid
        if (__pid) {
            launchedPIDs.Push(__pid)
            pinfo.Push({pid:__pid, title:__title})
            if (HIDE_FROM_ALTTAB)
                SetToolWindowStyle_ForPid(__pid)
        }
        ; tiny delay just to avoid stampeding the shell
        Sleep, 150

        __c++
        if (__c >= __cols)
            __c := 0, __r++
    }

    ; -------- Pass 2: after they exist, rename + add overlays ----------
    for _, o in pinfo {
        pid := o.pid  ; AHK v1: can't do %o.pid%, use a temp var
        WinWait, ahk_pid %pid%, , 400
        if (ErrorLevel)
            continue
        WinGet, __list, List, ahk_pid %pid%
        if (__list >= 1) {
            __main := __list1
            __clean := ExtractCharName(o.title)
            Loop, %__list% {
                __h := __list%A_Index%
                WinSetTitle, ahk_id %__h%,, %__clean%
            }
            Overlay_Attach(__main, __clean, LABEL_COLOR, LABEL_OFF_X, LABEL_OFF_Y)
        }
    }
}



CloseLaunchedClones() {
    global launchedPIDs
    __closed := (launchedPIDs.MaxIndex() = "" ? 0 : launchedPIDs.MaxIndex())
    for __k, __pid in launchedPIDs
        ClosePidWindows(__pid)
    launchedPIDs := []
    __dbgC := {}
    __dbgC.closed := __closed
    LogBlock("CLOSE-ALL", __dbgC)
    Overlay_DestroyAll()
}

; ---- compat helper: lowercase a string in AHK v1 ----
Lower(s) {
    StringLower, s, s
    return s
}

; ================================================================
; PickRegion_TwoClick — hardened against border/overlay/preview issues
; ================================================================
PickRegion_TwoClick() {
    global LAST_HWND, SEL_MIN_W, SEL_MIN_H

    __cx1 := 0, __cy1 := 0, __cx2 := 0, __cy2 := 0
    __rx := 0, __ry := 0, __rw := 0, __rh := 0

    ; --- Get real target window under cursor ---
    ToolTip, Click first corner...
    KeyWait, LButton, D
    Sleep, 10

    VarSetCapacity(pt, 8, 0)
    DllCall("GetCursorPos", "ptr", &pt)
    __sx1 := NumGet(pt, 0, "Int")
    __sy1 := NumGet(pt, 4, "Int")

    __hTarget := DllCall("user32\WindowFromPoint", "int64", (__sy1 << 32) | (__sx1 & 0xFFFFFFFF), "ptr")
    if (__hTarget)
        __hTarget := DllCall("user32\GetAncestor", "ptr", __hTarget, "uint", 2, "ptr")  ; GA_ROOT

    if (!__hTarget) {
        MsgBox, 48, No Window, Couldn't detect the window under the cursor.
        ToolTip
        return ""
    }

    ; --- Detect and skip EVE-O Preview ---
    WinGet, procName, ProcessName, ahk_id %__hTarget%
    procLower := procName
    StringLower, procLower, procLower
    if (procLower = "eve-o-preview.exe") {
        ; find real EVE client under preview
        VarSetCapacity(pt2, 8, 0)
        DllCall("GetCursorPos", "ptr", &pt2)
        sx := NumGet(pt2, 0, "Int")
        sy := NumGet(pt2, 4, "Int")
        hwndAtPt := DllCall("user32\WindowFromPoint", "int64", (sy << 32) | (sx & 0xFFFFFFFF), "ptr")
        if (hwndAtPt)
            __hTarget := DllCall("user32\GetAncestor", "ptr", hwndAtPt, "uint", 2, "ptr") ; GA_ROOT

        WinGet, procName2, ProcessName, ahk_id %__hTarget%
        StringLower, procName2, procName2
        if (procName2 != "exefile.exe") {
            MsgBox, 48, Wrong Window, You clicked on %procName2%. Please click an EVE client instead.
            ToolTip
            return ""
        }
    }

    LAST_HWND := __hTarget
    ToolTip, Click opposite corner...
    KeyWait, LButton
    ToolTip

    ; --- Wait for second click ---
    KeyWait, LButton, D
    Sleep, 10
    VarSetCapacity(pt3, 8, 0)
    DllCall("GetCursorPos", "ptr", &pt3)
    __sx2 := NumGet(pt3, 0, "Int")
    __sy2 := NumGet(pt3, 4, "Int")
    KeyWait, LButton
    ToolTip

    ; --- Get client rect + origin ---
    VarSetCapacity(__rc, 16, 0)
    DllCall("GetClientRect", "ptr", __hTarget, "ptr", &__rc)
    __cW := NumGet(__rc, 8, "Int")
    __cH := NumGet(__rc, 12, "Int")

    VarSetCapacity(__pt0, 8, 0)
    NumPut(0, __pt0, 0, "Int"), NumPut(0, __pt0, 4, "Int")
    DllCall("ClientToScreen", "ptr", __hTarget, "ptr", &__pt0)
    __cxs := NumGet(__pt0, 0, "Int")
    __cys := NumGet(__pt0, 4, "Int")

    ; --- Convert screen coords to client coords ---
    __cx1 := __sx1 - __cxs
    __cy1 := __sy1 - __cys
    __cx2 := __sx2 - __cxs
    __cy2 := __sy2 - __cys

    ; --- Normalize + clamp ---
    __rx := (__cx1 < __cx2 ? __cx1 : __cx2)
    __ry := (__cy1 < __cy2 ? __cy1 : __cy2)
    __rw := Abs(__cx2 - __cx1)
    __rh := Abs(__cy2 - __cy1)

    if (__rx < 0)
        __rx := 0
    if (__ry < 0)
        __ry := 0
    if (__rx + __rw > __cW)
        __rw := __cW - __rx
    if (__ry + __rh > __cH)
        __rh := __cH - __ry

    ; --- Enforce minimum size ---
    if (__rw < SEL_MIN_W || __rh < SEL_MIN_H) {
        __rw := Max(__rw, SEL_MIN_W)
        __rh := Max(__rh, SEL_MIN_H)
        if (__rx + __rw > __cW)
            __rx := Max(0, __cW - __rw)
        if (__ry + __rh > __cH)
            __ry := Max(0, __cH - __rh)
    }

    if (__rw <= 0 || __rh <= 0)
        return ""

    ; --- Log debug data ---
    __dbgP := {}
    __dbgP.client_size := __cW "x" __cH
    __dbgP.client_origin_screen := __cxs "," __cys
    __dbgP.screen_click1 := __sx1 "," __sy1
    __dbgP.screen_click2 := __sx2 "," __sy2
    __dbgP.client_click1 := __cx1 "," __cy1
    __dbgP.client_click2 := __cx2 "," __cy2
    __dbgP.xywh := __rx "," __ry "," __rw "," __rh
    LogBlock("PICK-FINAL", __dbgP)

    TrayTip, Region Picked, % "x=" __rx "  y=" __ry "  w=" __rw "  h=" __rh, 1000, 1
    return __rx "," __ry "," __rw "," __rh
}


ChooseClientX(__a, __b, __maxW) {
    if (__a >= 0 && __a < __maxW && __a != 0)
        return __a
    if (__b >= 0 && __b < __maxW && __b != 0)
        return __b
    if (__a >= 0 && __a < __maxW)
        return __a
    if (__b >= 0 && __b < __maxW)
        return __b
    return 0
}
ChooseClientY(__a, __b, __maxH) {
    if (__a >= 0 && __a < __maxH && __a != 0)
        return __a
    if (__b >= 0 && __b < __maxH && __b != 0)
        return __b
    if (__a >= 0 && __a < __maxH)
        return __a
    if (__b >= 0 && __b < __maxH)
        return __b
    return 0
}

GetScreenCursor_Safe(ByRef sx, ByRef sy) {
    VarSetCapacity(__pt, 8, 0)
    DllCall("GetCursorPos", "ptr", &__pt)
    __sx1 := NumGet(__pt, 0, "Int")
    __sy1 := NumGet(__pt, 4, "Int")

    CoordMode, Mouse, Screen
    MouseGetPos, __sx2, __sy2

    __msg := DllCall("GetMessagePos", "UInt")
    __sx3 := __msg & 0xFFFF
    __sy3 := (__msg >> 16) & 0xFFFF
    if (__sx3 >= 0x8000)
        __sx3 -= 0x10000
    if (__sy3 >= 0x8000)
        __sy3 -= 0x10000

    if (__sx1 != 0) {
        sx := __sx1, sy := __sy1
    } else if (__sx2 != 0) {
        sx := __sx2, sy := __sy2
    } else {
        sx := __sx3, sy := __sy3
    }

    __dbgS := {}
    __dbgS.GetCursorPos := __sx1 "," __sy1
    __dbgS.MouseGetPos  := __sx2 "," __sy2
    __dbgS.GetMessagePos:= __sx3 "," __sy3
    LogBlock("RAW-CURSOR", __dbgS)
}

; ============================================================
;                          Helpers
; ============================================================
; v1-safe lowercase helper
ToLower(str) {
    out := str
    StringLower, out, out
    return out
}

; ========= UI Guard for clean region picking (AHK v1-safe) =========
BeginPickUIGuard() {
    global Overlays
    Overlays_HideAll()
    st := { hidMGR:0, hidPicker:0, overlay:[], otr:[] }

    ; Hide Manager (if open)
    Gui, MGR:+LastFound
    if (WinExist()) {
        st.hidMGR := 1
        Gui, MGR:Hide
    }

    ; Hide Picker (if open)
    Gui, Picker:+LastFound
    if (WinExist()) {
        st.hidPicker := 1
        Gui, Picker:Hide
    }

    ; Hide all overlay labels
    if (IsObject(Overlays)) {
        for owner, o in Overlays {
            if (o.lbl && DllCall("IsWindow","ptr",o.lbl)) {
                lbl := o.lbl            ; <-- temp var so WinHide can use %lbl%
                WinHide, ahk_id %lbl%
                st.overlay.Push(lbl)
            }
        }
    }

    ; Hide all existing OTR clone windows (helps unobstruct the source)
    WinGet, L, List, ahk_exe OnTopReplica.exe
    Loop, %L% {
        h := L%A_Index%
        WinHide, ahk_id %h%
        st.otr.Push(h)
    }
    return st
}

EndPickUIGuard(st) {
    ; Restore OTR windows
    for _, h in st.otr {
        if (h && DllCall("IsWindow","ptr",h)) {
            WinShow, ahk_id %h%
        }
    }

    ; Restore overlays
    for _, lbl in st.overlay {
        if (lbl && DllCall("IsWindow","ptr",lbl)) {
            WinShow, ahk_id %lbl%
        }
    }

    ; Restore GUIs
    if (st.hidPicker)
        Gui, Picker:Show
    if (st.hidMGR)
        Gui, MGR:Show
    Overlays_ShowAll()
}

; Wrapper that uses the guard while picking
PickRegion_WithGuard() {
    st := BeginPickUIGuard()
    Sleep, 100  ; let the hides settle
    xywh := PickRegion_TwoClick()
    EndPickUIGuard(st)
    return xywh
}



Clamp(n, lo, hi) {
    return (n < lo) ? lo : (n > hi) ? hi : n
}

CloseAllOTRWindows() {
    WinGet, L, List, ahk_exe OnTopReplica.exe
    Loop, %L% {
        h := L%A_Index%
        Overlay_Destroy(h)
        PostMessage, 0x0010, 0, 0,, ahk_id %h%
    }
    Sleep, 200
    ; Nuke any leftovers
    Process, Close, OnTopReplica.exe
}

ResolveConfigPath() {
    local pList := []
    pList.Push(A_ScriptDir "\config.ini")
    pList.Push(A_MyDocuments "\AutoHotkey\EVE_Tiler\config.ini")
    pList.Push(A_AppData "\EVE_Tiler\config.ini")
    for _, p in pList
        if FileExist(p)
            return p
    ; if none exist, prefer Documents path but ensure folder exists
    if !FileExist(A_MyDocuments "\AutoHotkey\EVE_Tiler")
        FileCreateDir, % A_MyDocuments "\AutoHotkey\EVE_Tiler"
    return A_MyDocuments "\AutoHotkey\EVE_Tiler\config.ini"
}

_GetMonitorWA(monIndex, ByRef L, ByRef T, ByRef R, ByRef B) {
    L := "", T := "", R := "", B := ""
    SysGet, MonWA, MonitorWorkArea, %monIndex%
    L := MonWALeft, T := MonWATop, R := MonWARight, B := MonWABottom
    ; Fallbacks if SysGet failed or returned blanks (hotplug, sleep/wake, etc.)
    if (L = "" || T = "" || R = "" || B = "") {
        L := 0, T := 0
        ; try primary monitor as a second chance
        SysGet, v, 78  ; SM_CXVIRTUALSCREEN
        SysGet, h, 79  ; SM_CYVIRTUALSCREEN
        if (v != "" && h != "" && v > 0 && h > 0) {
            R := v, B := h
        } else {
            R := A_ScreenWidth, B := A_ScreenHeight
        }
    }
}

DBG_StartSession(label, total) {
    global DBG_LastLaunches
    DBG_LastLaunches := []
    LogBlock("TILE-SESSION-BEGIN", { label: label, total: total })
}

DBG_ShowSummary(label := "") {
    global DBG_LastLaunches
    ok := 0, fail := 0, body := ""
    for _, o in DBG_LastLaunches {
        if (o.ok)
            ok++
        else
            fail++

        body .= (o.ok ? "[OK]   " : "[FAIL] ")
            . o.title
            . "  attempts=" . o.attempts
            . "  hwnds=" . o.hwnds
            . "  " . o.elapsed . "ms`n"
    }
    if (body = "")
        body := "(no launches recorded)"
    msg := (label!=""?label:"Tile summary")
        . "`nOK=" ok "  Fail=" fail
        . "`n-------------------------`n" body
    MsgBox, 64, Tile summary, %msg%

    LogBlock("TILE-SESSION-END", { label: label, ok: ok, fail: fail })
}


; Global store for the most recent launch session:
DBG_LastLaunches := []    ; array of {title, ok, pid, attempts, elapsed, hwnds}

_robustLaunchOTR(args, wait_ms := 4000, tries := 4, title := "") {
    global OTR_PATH, DEBUG_VERBOSE, DBG_LastLaunches
    t0 := A_TickCount
    attempt := 0
    ok := 0, pid := 0, hwnds := 0

    Loop, %tries% {
        attempt++
        pid := 0
        Run, % Chr(34) . OTR_PATH . Chr(34) . " " . args,, Hide, pid
        if (pid) {
            WinWait, ahk_pid %pid%,, %wait_ms%
            if (!ErrorLevel) {
                WinGet, cnt, List, ahk_pid %pid%
                hwnds := cnt+0
                if (cnt >= 1) {
                    ok := 1
                    break
                }
            }
            ; no window appeared -> kill and retry
            Process, Close, %pid%
        }
        Sleep, 150 + (attempt-1)*150
    }

    elapsed := A_TickCount - t0
    if (DEBUG_VERBOSE) {
        LogBlock("OTR-SPAWN", { title: (title!=""?title:"(n/a)")
                            , ok: ok, pid: pid, attempts: attempt
                            , hwnds: hwnds, waited_ms: elapsed })
    }
    ; keep a short in-memory summary for on-screen report
    DBG_LastLaunches.Push({ title: (title!=""?title:"(n/a)"), ok: ok
                        , pid: pid, attempts: attempt
                        , elapsed: elapsed, hwnds: hwnds })
    return pid
}

EnsureSourceReadyExact(title) {
    WinGet, h, ID, %title%
    if (!h) {
        LogBlock("SRC-READY", { title: title, found: 0 })
        return 0
    }
    WinGet, mm, MinMax, ahk_id %h%   ; 0=normal 1=min 2=max
    if (mm = 1)
        WinRestore, ahk_id %h%
    WinShow, ahk_id %h%

    ; DWM cloaked?
    VarSetCapacity(cloaked, 4, 0)
    isC := ""
    if DllCall("dwmapi\DwmGetWindowAttribute","ptr",h,"int",14,"ptr",&cloaked,"int",4)=0
        isC := NumGet(cloaked,0,"UInt")

    VarSetCapacity(rc,16,0)
    DllCall("GetClientRect","ptr",h,"ptr",&rc)
    cw := NumGet(rc,8,"Int"), ch := NumGet(rc,12,"Int")

    LogBlock("SRC-READY", { title: title, found: 1, minimized: mm, cloaked: isC
                        , client: cw "x" ch })
    return h
}


; Map: cleaned character name -> live full title
BuildLiveClientMap() {
    m := {}
    WinGet, idList, List, ahk_exe exefile.exe
    Loop, % idList {
        h := idList%A_Index%
        WinGetTitle, t, ahk_id %h%
        if (t != "") {
            ct := ExtractCharName(t)
            StringLower, ct, ct
            m[ct] := t
        }
    }
    return m
}

CleanupLaunch(ByRef oldMon:="", ByRef oldGX:="", ByRef oldGY:="", ByRef oldScale:="") {
    global TARGET_MONITOR, GAP_X, GAP_Y, REGION_SCALE, IsLaunching
    if (oldMon != "")
        TARGET_MONITOR := oldMon, GAP_X := oldGX, GAP_Y := oldGY, REGION_SCALE := oldScale
    StartAutoRetile()
    IsLaunching := 0
}

; ================= Overlay Labels (simple timer, click-through, topmost) =================
; One timer drives all label movement. Very robust with OTR.

Overlays := {}            ; map ownerHWND -> { lbl:hLbl, offX:int, offY:int }
OverlaysHidden := false   ; global hide flag (used by region picker)
OVERLAY_TIMER_RUNNING := 0

Overlay_StartTimer() {
    global OVERLAY_TIMER_RUNNING
    if (!OVERLAY_TIMER_RUNNING) {
        SetTimer, Overlay_FollowAll, 50
        OVERLAY_TIMER_RUNNING := 1
    }
}

Overlay_StopTimerIfNone() {
    global Overlays, OVERLAY_TIMER_RUNNING
    if (!IsObject(Overlays) || Overlays.MaxIndex() = "") {
        SetTimer, Overlay_FollowAll, Off
        OVERLAY_TIMER_RUNNING := 0
    }
}

; PUBLIC: attach a label to an owner window (OTR clone)
Overlay_Attach(ownerHwnd, text, colorHex:="FFFF00", offX:=8, offY:=6) {
    global Overlays, OverlaysHidden

    if (!ownerHwnd || !DllCall("IsWindow","ptr",ownerHwnd))
        return 0

    if (!IsObject(Overlays))
        Overlays := {}

    ; if already attached, just update offsets and reposition
    if (Overlays.HasKey(ownerHwnd)) {
        o := Overlays[ownerHwnd]
        o.offX := offX
        o.offY := offY
        Overlays[ownerHwnd] := o
        Overlay_UpdatePosition(ownerHwnd)
        return o.lbl
    }

    ; ---- create a fresh overlay GUI ----
    Gui, New, +AlwaysOnTop -Caption +ToolWindow +E0x20 +HwndhLbl
    Gui, Margin, 0, 0
    Gui, Color, 000000
    Gui, Font, s10 Bold c%colorHex%, Segoe UI
    Gui, Add, Text, BackgroundTrans, %text%

    ; show somewhere; Overlay_UpdatePosition will place it correctly
    Gui, Show, NoActivate x0 y0

    ; extended styles: click-through, layered, toolwindow, no-activate
    ex := DllCall("GetWindowLong" (A_PtrSize=8 ? "Ptr" : "")
                , "ptr", hLbl, "int", -20, "ptr")
    ex |= 0x00000020    ; WS_EX_TRANSPARENT (click-through)
    ex |= 0x00080000    ; WS_EX_LAYERED
    ex |= 0x00000080    ; WS_EX_TOOLWINDOW
    ex |= 0x08000000    ; WS_EX_NOACTIVATE
    DllCall("SetWindowLong" (A_PtrSize=8 ? "Ptr" : "")
        , "ptr", hLbl, "int", -20, "ptr", ex)

    Overlays[ownerHwnd] := { lbl:hLbl, offX:offX, offY:offY }

    if (OverlaysHidden) {
        WinHide, ahk_id %hLbl%
    } else {
        Overlay_UpdatePosition(ownerHwnd)
    }

    Overlay_StartTimer()
    return hLbl
}

; Position a single overlay relative to its owner
Overlay_UpdatePosition(ownerHwnd) {
    global Overlays, OverlaysHidden
    if (!IsObject(Overlays))
        return
    if (!Overlays.HasKey(ownerHwnd))
        return

    o := Overlays[ownerHwnd]
    hLbl := o.lbl
    if (!hLbl || !DllCall("IsWindow","ptr",hLbl))
        return
    if (!DllCall("IsWindow","ptr",ownerHwnd)) {
        ; owner died – clean up
        this_id := hLbl
        WinClose, ahk_id %this_id%
        Overlays.Delete(ownerHwnd)
        Overlay_StopTimerIfNone()
        return
    }

    offX := o.offX
    offY := o.offY

    ; respect global hide flag
    if (OverlaysHidden) {
        WinHide, ahk_id %hLbl%
        return
    }

    ; hide when owner is minimized or invisible
    WinGet, mm, MinMax, ahk_id %ownerHwnd%
    if (mm = 1) {
        WinHide, ahk_id %hLbl%
        return
    }
    WinGet, style, Style, ahk_id %ownerHwnd%
    if !(style & 0x10000000) {   ; WS_VISIBLE
        WinHide, ahk_id %hLbl%
        return
    }

    VarSetCapacity(rc, 16, 0)
    if !DllCall("GetWindowRect","ptr",ownerHwnd,"ptr",&rc)
        return
    x := NumGet(rc,0,"Int")
    y := NumGet(rc,4,"Int")

    ; keep overlay on top of everything, no focus steal
    WinShow, ahk_id %hLbl%
    DllCall("SetWindowPos","ptr",hLbl,"ptr",-1
        ,"int",x+offX,"int",y+offY,"int",0,"int",0,"uint",0x0011) ; NOSIZE|NOACTIVATE
}

; Timer: follow all overlays
Overlay_FollowAll:
    global Overlays
    if (!IsObject(Overlays) || Overlays.MaxIndex() = "")
        return
    for ownerHwnd, o in Overlays
        Overlay_UpdatePosition(ownerHwnd)
return

; Destroy overlay for a single owner
Overlay_Destroy(ownerHwnd) {
    global Overlays
    if (!IsObject(Overlays))
        return
    if (!Overlays.HasKey(ownerHwnd))
        return
    o := Overlays[ownerHwnd]
    if (o.lbl && DllCall("IsWindow","ptr",o.lbl)) {
        this_id := o.lbl
        WinClose, ahk_id %this_id%
    }
    Overlays.Delete(ownerHwnd)
    Overlay_StopTimerIfNone()
}

; Destroy all overlays
Overlay_DestroyAll() {
    global Overlays
    if (!IsObject(Overlays))
        return
    for ownerHwnd, o in Overlays {
        if (o.lbl && DllCall("IsWindow","ptr",o.lbl)) {
            this_id := o.lbl
            WinClose, ahk_id %this_id%
        }
    }
    Overlays := {}
    Overlay_StopTimerIfNone()
}

; Auto-hide/show helpers used by PickRegion_WithGuard()
Overlays_HideAll() {
    global Overlays, OverlaysHidden
    OverlaysHidden := true
    if (!IsObject(Overlays))
        return
    for ownerHwnd, o in Overlays {
        if (DllCall("IsWindow","ptr",o.lbl)) {
            this_id := o.lbl
            WinHide, ahk_id %this_id%
        }
    }
}

Overlays_ShowAll() {
    global Overlays, OverlaysHidden
    OverlaysHidden := false
    if (!IsObject(Overlays))
        return
    for ownerHwnd, o in Overlays {
        if (DllCall("IsWindow","ptr",o.lbl)) {
            this_id := o.lbl
            WinShow, ahk_id %this_id%
        }
    }
}
; ======================================================================


; Return hwnd of an EVE client whose title matches exactly, else 0
FindClientHwndByTitleExact(title) {
    ; try exact first
    WinGet, idList, List, ahk_exe exefile.exe
    Loop, % idList {
        h := idList%A_Index%
        WinGetTitle, t, ahk_id %h%
        if (t = title)
            return h
    }

    ; fallback: match by cleaned character name (case-insensitive)
    cleanWanted := ExtractCharName(title)
    StringLower, cleanWanted, cleanWanted

    WinGet, idList2, List, ahk_exe exefile.exe
    Loop, % idList2 {
        h := idList2%A_Index%
        WinGetTitle, t2, ahk_id %h%
        if (t2 != "") {
            ct := ExtractCharName(t2)
            StringLower, ct, ct
            if (ct = cleanWanted)
                return h
        }
    }
    return 0
}


; True if an EVE client with this exact title is present
ClientWindowExists(titleOrName) {
    return FindClientHwndByTitleExact(titleOrName) != 0
}

ExtractCharName(fullTitle) {
    ; tries to get "Character" from things like "EVE - Character"
    name := fullTitle
    if InStr(name, " - ")
        name := SubStr(name, InStr(name, " - ",, 0) + 3)
    ; trim common prefixes
    name := RegExReplace(name, "i)^(EVE\s*-\s*)", "")
    name := RegExReplace(name, "i)^(OnTopReplica\s*-\s*)", "")
    name := Trim(name)
    if (name = "")
        name := fullTitle
    return name
}

LaunchOneClone(title, rX, rY, rW, rH, x, y, outW, outH) {
    ;global OTR_PATH, COMMON_FLAGS, REGION_FLIP_Y
    global
    ; Don't launch OTR if the source client isn't present
    if (!ClientWindowExists(title))
        return 0

    rY_use := rY
    if (REGION_FLIP_Y) {
        hSrc := FindClientHwndByTitleExact(title)
        if (hSrc) {
            VarSetCapacity(rc2, 16)
            DllCall("GetClientRect", "ptr", hSrc, "ptr", &rc2)
            clientH := NumGet(rc2, 12, "Int")
            if (clientH) {
                rY_use := clientH - rY - rH
                if (rY_use < 0)
                    rY_use := 0
            }
        }
    }

    args := "--windowTitle=" . Chr(34) . title . Chr(34)
    args .= " " . COMMON_FLAGS
    args .= " --region=" . rX "," . rY_use "," . rW "," . rH
    args .= " --size="   . outW "," . outH
    args .= " --position=" . x "," . y

    Run, % Chr(34) . OTR_PATH . Chr(34) . " " . args,, Hide, newPID
    if (!newPID)
        return 0

    ; Rename the clone window to the character name (clean)
    clean := ExtractCharName(title)
    WinWait, ahk_pid %newPID%,, 1500
    if (ErrorLevel=0) {
        ; pick the visible main window for this PID
        WinGet, list, List, ahk_pid %newPID%
        if (list >= 1) {
            mainHwnd := list1
            ; rename all windows to clean title (as you already do)
            Loop, %list% {
                hwnd := list%A_Index%
                WinSetTitle, ahk_id %hwnd%,, %clean%
            }
            ; attach overlay label to the chosen owner window
            Overlay_Attach(mainHwnd, clean, LABEL_COLOR, LABEL_OFF_X, LABEL_OFF_Y)
        }
    }
    return newPID
}


; Hide a process' windows from Alt-Tab by toggling extended styles
SetToolWindowStyle_ForPid(pid) {
    if (!pid)
        return
    ; enumerate all windows for this PID
    WinGet, list, List, ahk_pid %pid%
    Loop, %list% {
        hwnd := list%A_Index%
        ; get exstyle
        ex := DllCall("GetWindowLong" (A_PtrSize=8 ? "Ptr" : ""), "ptr", hwnd, "int", -20, "ptr")
        ; add TOOLWINDOW, remove APPWINDOW
        ex |= 0x80
        ex &= ~0x40000
        DllCall("SetWindowLong" (A_PtrSize=8 ? "Ptr" : ""), "ptr", hwnd, "int", -20, "ptr", ex)
        ; force style refresh
        DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", 0, "int", 0
            , "uint", 0x27) ; NOSIZE|NOMOVE|NOZORDER|FRAMECHANGED
    }
}


Log(__msg) {
    global DEBUG, LOG_PATH
    if (!DEBUG)
        return
    __time := A_YYYY "-" A_MM "-" A_DD " " A_Hour ":" A_Min ":" A_Sec
    FileAppend, % __time "  " __msg "`r`n", %LOG_PATH%
}

LogBlock(__tag, __obj) {
    global DEBUG, LOG_PATH
    if (!DEBUG)
        return
    __time := A_YYYY "-" A_MM "-" A_DD " " A_Hour ":" A_Min ":" A_Sec
    FileAppend, % "----- " __tag " @ " __time " -----`r`n", %LOG_PATH%
    for __k, __v in __obj
        FileAppend, % __k ": " __v "`r`n", %LOG_PATH%
    FileAppend, % "--------------------------------`r`n", %LOG_PATH%
}

Sqrt(__n) {
    return DllCall("msvcrt\sqrt", "double", __n, "double")
}

; ============================================================
; EVE Multibox Tiler — Manager (clean layout)
; ============================================================
OpenGroupManager(reopen:=false) {
    global Manager_hLV_Groups, Manager_hLV_Members, Manager_hLV_Clients
    global Groups, SelectedGroup, TARGET_MONITOR, GAP_X, GAP_Y, REGION_SCALE
    global MGR_Layout, MGR_Cols, MGR_Rows, MGR_Mon, MGR_GapX, MGR_GapY, MGR_Scale
    global MGR_Accumulate, AUTO_RETILE

    if (!reopen)
        LoadConfig()

    DetectClients := DetectEVEClients()

    ; ----- main window -----
    Gui, MGR:New, +AlwaysOnTop +Resize +MinSize820x540, EVE Multibox Tiler — Manager
    Gui, MGR:+DPIScale
    Gui, MGR:Color, F4F4F4
    Gui, MGR:Font, s9, Segoe UI
    Gui, MGR:Margin, 12, 10

    ; ------------------------------------------------------------------
    ; Title + subtitle
    ; ------------------------------------------------------------------
    Gui, MGR:Font, s11 Bold, Segoe UI
    Gui, MGR:Add, Text, xm ym, EVE Multibox Tiler
    Gui, MGR:Font, s9 c555555, Segoe UI
    Gui, MGR:Add, Text, x+8 yp+4, Window groups and tiled clones for EVE
    Gui, MGR:Font, s9 c000000, Segoe UI

    ; small vertical gap
    Gui, MGR:Add, Text, xm y+4, 

    ; ------------------------------------------------------------------
    ; Top action bar
    ; ------------------------------------------------------------------
    Gui, MGR:Add, Button, gMGR_NewGroup       w80  h24 xm         , New
    Gui, MGR:Add, Button, gMGR_RenameGroup   w80  h24 x+6        , Rename
    Gui, MGR:Add, Button, gMGR_DeleteGroup   w80  h24 x+6        , Delete
    Gui, MGR:Add, Button, gMGR_PickRegion    w90  h24 x+10       , Pick Region
    Gui, MGR:Add, Button, gMGR_TileChecked   w100 h24 x+8 Default, Tile Checked
    Gui, MGR:Add, Button, gMGR_CloseChecked  w100 h24 x+6        , Close Checked
    Gui, MGR:Add, Button, gMGR_Save          w80  h24 x+10       , Save
    Gui, MGR:Add, Button, gMGR_Load          w80  h24 x+6        , Load
    Gui, MGR:Add, Button, gMGR_SaveLayout    w120 h24 x+10       , Save Layout
    Gui, MGR:Add, Button, gMGR_ClearLayout   w110 h24 x+6        , Clear Layout

    ; ------------------------------------------------------------------
    ; Options row
    ; ------------------------------------------------------------------
    Gui, MGR:Add, Checkbox, xm   y+10 vMGR_Accumulate Checked%MGR_Accumulate%
        , Keep existing clones (accumulate)
    Gui, MGR:Add, Checkbox, x+30 yp vAUTO_RETILE gMGR_OnAutoRetile Checked%AUTO_RETILE%
        , Auto retile on window changes

    ; ------------------------------------------------------------------
    ; Layout & placement group
    ; ------------------------------------------------------------------
    Gui, MGR:Add, GroupBox, xm y+10 w790 h90, Layout & placement

    ; First row inside groupbox
    Gui, MGR:Add, Text,  xp+16 yp+26, Layout:
    Gui, MGR:Add, DropDownList, vMGR_Layout w90 yp-3 x+4 gMGR_OnLayoutChoice, auto|grid

    Gui, MGR:Add, Text,  x+24 yp+3, Cols:
    Gui, MGR:Add, Edit,  vMGR_Cols w40 yp-3 x+4 Number, 0

    Gui, MGR:Add, Text,  x+20 yp+3, Rows:
    Gui, MGR:Add, Edit,  vMGR_Rows w40 yp-3 x+4 Number, 0

    Gui, MGR:Add, Text,  x+24 yp+3, Mon:
    Gui, MGR:Add, Edit,  vMGR_Mon w36 yp-3 x+4 Number, %TARGET_MONITOR%

    Gui, MGR:Add, Text,  x+24 yp+3, GapX:
    Gui, MGR:Add, Edit,  vMGR_GapX w40 yp-3 x+4 Number, %GAP_X%

    Gui, MGR:Add, Text,  x+20 yp+3, GapY:
    Gui, MGR:Add, Edit,  vMGR_GapY w40 yp-3 x+4 Number, %GAP_Y%

    Gui, MGR:Add, Text,  x+24 yp+3, Scale:
    Gui, MGR:Add, Edit,  vMGR_Scale w60 yp-3 x+4, %REGION_SCALE%

    ; ------------------------------------------------------------------
    ; Groups / Members / Clients area
    ; ------------------------------------------------------------------
    Gui, MGR:Add, GroupBox, xm y+14 w790 h270, Groups & clients

    ; column headers
    Gui, MGR:Add, Text, xp+16 yp+24, Groups (check to auto-retile)
    Gui, MGR:Add, Text, x+260 yp, Members in group
    Gui, MGR:Add, Text, x+260 yp, Available clients

    ; lists (three columns)
    Gui, MGR:Add, ListView
        , hwndManager_hLV_Groups xm+16 y+4 w230 r12 Checked gMGR_OnGroupClick
        , Group
    Gui, MGR:Add, ListView
        , hwndManager_hLV_Members x+10 yp w260 r12 Multi gMGR_MembersLV
        , Character title
    Gui, MGR:Add, ListView
        , hwndManager_hLV_Clients x+10 yp w260 r12 Multi gMGR_ClientsLV
        , Window title

    ; Add / Remove buttons centered under first two lists
    Gui, MGR:Add, Button, xm+16    y+8 w120 gMGR_AddMember  , << Add
    Gui, MGR:Add, Button, x+10     w120 gMGR_RemoveMember   , Remove >>

    ; Status bar
    Gui, MGR:Add, StatusBar
    SB_SetText("Ready")

    ; ------------------------------------------------------------------
    ; Populate data
    ; ------------------------------------------------------------------
    Gui, MGR:Default

    ; groups list
    Gui, ListView, %Manager_hLV_Groups%
    LV_Delete()
    for name, g in Groups {
        row := LV_Add((g.enabled=1 ? "Check" : ""), name)
        if (name = SelectedGroup)
            LV_Modify(row, "Select Vis")
    }
    LV_ModifyCol(1, "AutoHdr")

    ; clients list
    Gui, ListView, %Manager_hLV_Clients%
    LV_Delete()
    for idx, o in DetectClients
        LV_Add("", o.title)
    LV_ModifyCol(1, "AutoHdr")

    ; members list for currently selected group
    RefreshMembersLV()

    ; restore layout controls from group
    if (SelectedGroup != "")
        ApplyGroupSettingsToControls(SelectedGroup)

    Gui, MGR:Show, Center w820 h560
}


CloneDefaultGroup() {
    global DEF_GROUP
    g := {}
    g.members := []
    g.region := DEF_GROUP.region
    g.layout := { mode: DEF_GROUP.layout.mode, cols: DEF_GROUP.layout.cols, rows: DEF_GROUP.layout.rows }
    g.monitor := DEF_GROUP.monitor
    g.gapX := DEF_GROUP.gapX
    g.gapY := DEF_GROUP.gapY
    g.scale := DEF_GROUP.scale
    return g
}

SelectGroupInLV(name) {
    global Manager_hLV_Groups, Groups, SelectedGroup
    Gui, ListView, %Manager_hLV_Groups%
    __cnt := LV_GetCount()
    Loop, %__cnt% {
        LV_GetText(__t, A_Index)
        LV_Modify(A_Index, ( __t = name ? "Select Vis" : "-Select" ))
    }
    SelectedGroup := name
    RefreshMembersLV()
    ShowGroupStatus()
}

RefreshMembersLV() {
    global Manager_hLV_Members, Groups, SelectedGroup
    Gui, ListView, %Manager_hLV_Members%
    LV_Delete()
    if (SelectedGroup = "" || !Groups.HasKey(SelectedGroup))
        return
    for _, t in Groups[SelectedGroup].members
        LV_Add("", t)
    LV_ModifyCol(1, 230)
}

ShowGroupStatus() {
    global Groups, SelectedGroup
    if (SelectedGroup = "" || !Groups.HasKey(SelectedGroup)) {
        SB_SetText("No group selected")
        return
    }
    g := Groups[SelectedGroup]
    text := "Group: " . SelectedGroup
        . " | Region: " . (g.region != "" ? g.region : "(none)")
        . " | Layout: " . g.layout.mode . " " . g.layout.cols . "x" . g.layout.rows
        . " | Mon: " . g.monitor
        . " | Gap: " . g.gapX . "," . g.gapY
        . " | Scale: " . g.scale
    SB_SetText(text)
}


ApplyGroupSettingsToControls(name) {
    global Groups
    if (name = "" || !Groups.HasKey(name))
        return
    g := Groups[name]

    ; Select the current layout in the DDL
    GuiControl, ChooseString, MGR_Layout, % g.layout.mode

    ; Fill the edit boxes
    GuiControl,, MGR_Cols,  % g.layout.cols
    GuiControl,, MGR_Rows,  % g.layout.rows
    GuiControl,, MGR_Mon,   % g.monitor
    GuiControl,, MGR_GapX,  % g.gapX
    GuiControl,, MGR_GapY,  % g.gapY
    GuiControl,, MGR_Scale, % g.scale

    ShowGroupStatus()
}

StartAutoRetile() {
    global AUTO_RETILE, __WATCH_INTERVAL
    if (!AUTO_RETILE) {
        SetTimer, _WatchEVE, Off
        return
    }
    SetTimer, _WatchEVE, % __WATCH_INTERVAL
}

StopAutoRetile() {
    SetTimer, _WatchEVE, Off
}

_WatchEVE:
    global __LAST_EVE_KEY, __RETIPLE_AT
    ; build a stable fingerprint of all EVE client titles
    WinGet, idList, List, ahk_exe exefile.exe
    names := []
    Loop, % idList {
        h := idList%A_Index%
        WinGetTitle, t, ahk_id %h%
        if (t != "")
            names.Push(ExtractCharName(t))
    }
    ; sort + join
    Sort, names, CL
    key := ""
    for i, v in names
        key .= (i=1 ? "" : "|") v

    if (key != __LAST_EVE_KEY) {
        __LAST_EVE_KEY := key
        ; debounce ~1s
        __RETIPLE_AT := A_TickCount + 1000
    }

    if (__RETIPLE_AT && A_TickCount >= __RETIPLE_AT) {
        __RETIPLE_AT := 0
        Gosub, _DoRetile
    }
return

_DoRetile:
    ; re-tile only groups that are checked in the Manager list
    global Manager_hLV_Groups, Groups
    if (!IsObject(Groups))
        return
    ; if manager is open, respect checks; otherwise, tile all enabled groups
    tiled := 0
    if (Manager_hLV_Groups) {
        Gui, ListView, %Manager_hLV_Groups%
        idx := 0
        while (idx := LV_GetNext(idx, "C")) {
            LV_GetText(name, idx)
            if (Groups.HasKey(name))
                LaunchGroup(name), tiled++
        }
    } else {
        for name, g in Groups
            if (g.enabled=1)
                LaunchGroup(name), tiled++
    }
return


PersistControlsIntoGroup() {
    global Groups, SelectedGroup
    global MGR_Layout, MGR_Cols, MGR_Rows, MGR_Mon, MGR_GapX, MGR_GapY, MGR_Scale

    if (SelectedGroup = "" || !Groups.HasKey(SelectedGroup))
        return

    Gui, MGR:Submit, NoHide

    g := Groups[SelectedGroup]
    g.layout.mode := MGR_Layout
    g.layout.cols := MGR_Cols+0
    g.layout.rows := MGR_Rows+0
    g.monitor     := MGR_Mon+0
    g.gapX        := MGR_GapX+0
    g.gapY        := MGR_GapY+0
    g.scale       := MGR_Scale+0.0
    Groups[SelectedGroup] := g
    ShowGroupStatus()
}


; ---- Manager events ----
MGR_OnAutoRetile:
    global AUTO_RETILE
    Gui, MGR:Submit, NoHide
    if (AUTO_RETILE)
        StartAutoRetile()
    else
        StopAutoRetile()
    SaveConfig()  ; persist immediately
return

MGR_TileChecked:
    ; persist settings first
    PersistControlsIntoGroup()

    global Manager_hLV_Groups, Groups
    Gui, ListView, %Manager_hLV_Groups%

    ; collect all CHECKED names
    idx := 0
    names := []
    while (idx := LV_GetNext(idx, "C")) {
        LV_GetText(nm, idx)
        if (nm != "")
            names.Push(nm)
    }

    ; fallback: if none checked, use the SELECTED row
    if (names.Length() = 0) {
        sel := LV_GetNext()
        if (sel) {
            LV_GetText(nm, sel)
            if (nm != "")
                names.Push(nm)
        }
    }

    if (names.Length() = 0) {
        MsgBox, 48, Nothing to tile, Check a group or select one first.
        return
    }

    ; tile each requested group
    for _, nm in names
        if (Groups.HasKey(nm))
            LaunchGroup(nm)
return


MGR_CloseChecked:
    CloseCheckedGroups()
return

MGR_OnGroupClick:
    global Manager_hLV_Groups, Groups, SelectedGroup
    Gui, ListView, %Manager_hLV_Groups%
    row := LV_GetNext()
    if (!row)
        return
    LV_GetText(name, row)
    ; mirror check state into Groups[name].enabled
    checked := (LV_GetNext(row-1, "C") = row) ? 1 : 0
    if (Groups.HasKey(name))
        Groups[name].enabled := checked
    SelectGroupInLV(name)
    ApplyGroupSettingsToControls(name)
return

MGR_OnLayoutChoice:
    PersistControlsIntoGroup()
return

MGR_NewGroup:
    InputBox, __gn, New Group, Enter a new group name:
    if (ErrorLevel || __gn = "")
        return
    if (Groups.HasKey(__gn)) {
        MsgBox, 48, Exists, A group with that name already exists.
        return
    }
    Groups[__gn] := CloneDefaultGroup()
    SelectedGroup := __gn
    OpenGroupManager(true)
return

MGR_RenameGroup:
    global Groups, SelectedGroup
    if (SelectedGroup = "") {
        MsgBox, 48, Rename, Select a group first.
        return
    }
    InputBox, __new, Rename Group, New name for "%SelectedGroup%":
    if (ErrorLevel || __new = "")
        return
    if (Groups.HasKey(__new)) {
        MsgBox, 48, Exists, Another group already has that name.
        return
    }
    Groups[__new] := Groups[SelectedGroup]
    Groups.Delete(SelectedGroup)
    SelectedGroup := __new
    OpenGroupManager(true)
return

MGR_DeleteGroup:
    global Groups, SelectedGroup
    if (SelectedGroup = "")
        return
    MsgBox, 49, Delete Group, Delete "%SelectedGroup%"?
    IfMsgBox, OK
    {
        Groups.Delete(SelectedGroup)
        SelectedGroup := ""
        OpenGroupManager(true)
    }
return

MGR_AddMember:
    global Manager_hLV_Clients, Groups, SelectedGroup
    if (SelectedGroup = "")
        return
    Gui, ListView, %Manager_hLV_Clients%
    idx := 0, added := 0
    g := Groups[SelectedGroup]
    while (idx := LV_GetNext(idx)) {
        LV_GetText(title, idx)
        exists := false
        for _, t in g.members
            if (t = title) {
                exists := true
                break
            }
        if (!exists) {
            g.members.Push(title), added++
        }
    }
    Groups[SelectedGroup] := g
    if (added=0) {
        ; fall back: if nothing selected, treat current row as one item
        idx := LV_GetNext()
        if (idx) {
            LV_GetText(title, idx)
            exists := false
            for _, t in g.members
                if (t = title) {
                    exists := true
                    break
                }
            if (!exists) {
                g.members.Push(title)
                Groups[SelectedGroup] := g
            }
        }
    }
    RefreshMembersLV()
    ShowGroupStatus()
return

MGR_RemoveMember:
    global Manager_hLV_Members, Groups, SelectedGroup
    if (SelectedGroup = "")
        return
    Gui, ListView, %Manager_hLV_Members%

    sel := {}
    idx := 0, cnt := 0
    while (idx := LV_GetNext(idx)) {
        LV_GetText(title, idx)
        sel[title] := 1, cnt++
    }
    if (cnt = 0) {           ; fallback: single current row
        idx := LV_GetNext()
        if (!idx)
            return
        LV_GetText(title, idx)
        sel := { (title): 1 }
    }

    g := Groups[SelectedGroup]
    newArr := []
    for _, t in g.members
        if (!sel.HasKey(t))
            newArr.Push(t)
    g.members := newArr
    Groups[SelectedGroup] := g
    RefreshMembersLV()
    ShowGroupStatus()
return

MGR_ClientsLV:
    if (A_GuiEvent = "DoubleClick")
        Gosub, MGR_AddMember
return

MGR_MembersLV:
    if (A_GuiEvent = "DoubleClick")
        Gosub, MGR_RemoveMember
return

MGR_SaveLayout:
    global SelectedGroup, Groups, GroupPIDs
    if (SelectedGroup = "" || !Groups.HasKey(SelectedGroup)) {
        MsgBox, 48, No group, Select a group first.
        return
    }
    if (!GroupPIDs.HasKey(SelectedGroup) || GroupPIDs[SelectedGroup].Length()=0) {
        MsgBox, 48, No clones, Tile the group first so I can read their positions.
        return
    }
    fixed := {}
    saved := 0
    for _, obj in GroupPIDs[SelectedGroup] {
        pid := obj.pid
        ttl := obj.title
        WinGet, list, List, ahk_pid %pid%
        if (list < 1)
            continue
        hwnd := list1
        WinGetPos, x, y, w, h, ahk_id %hwnd%
        if (w != "" && h != "")
            fixed[ ExtractCharName(ttl) ] := x "," y "," w "," h
            saved++
    }
    if (saved=0) {
        MsgBox, 48, No windows, Couldn't read clone windows for this group.
        return
    }
    g := Groups[SelectedGroup]
    g.fixed := fixed
    Groups[SelectedGroup] := g
    SaveConfig()
    SB_SetText("Saved layout for " SelectedGroup)
return

MGR_ClearLayout:
    global SelectedGroup, Groups
    if (SelectedGroup = "" || !Groups.HasKey(SelectedGroup))
        return
    g := Groups[SelectedGroup]
    if (g.HasKey("fixed"))
        g.Delete("fixed")
    Groups[SelectedGroup] := g
    SaveConfig()
    SB_SetText("Cleared saved layout for " SelectedGroup)
return


MGR_PickRegion:
    global Groups, SelectedGroup
    if (SelectedGroup = "")
        return
    __xywh := PickRegion_WithGuard()
    if (__xywh = "")
        return
    g := Groups[SelectedGroup]
    g.region := __xywh
    Groups[SelectedGroup] := g
    ShowGroupStatus()
return


MGR_TileGroup:
    global SelectedGroup
    if (SelectedGroup != "")
        LaunchGroup(SelectedGroup)
return

MGR_Save:
    PersistControlsIntoGroup()
    SaveConfig()
    SB_SetText("Saved.")
return

MGR_Load:
    LoadConfig()
    ; Make a guess for SelectedGroup if empty
    if (!SelectedGroup) {
        for name, g in Groups {
            SelectedGroup := name
            break
        }
    }
    ; Re-open manager using current Groups in memory
    OpenGroupManager(true)
return

; ---- Persistence (INI) ----
SaveConfig() {
    global CFG_FILE, Groups, AUTO_RETILE

    if (CFG_FILE = "")
        CFG_FILE := ResolveConfigPath()

    tmp := CFG_FILE ".tmp"
    FileDelete, %tmp%

    ; General settings
    IniWrite, % (AUTO_RETILE ? 1 : 0), %tmp%, General, AutoRetile

    ; Groups
    for name, g in Groups {
        section := "Group:" name
        IniWrite, % (g.enabled ? 1 : 0), %tmp%, %section%, Enabled
        IniWrite, % JoinPipe(g.members), %tmp%, %section%, Members
        IniWrite, % (g.region="") ? "" : g.region, %tmp%, %section%, Region
        IniWrite, % g.layout.mode, %tmp%, %section%, Layout
        IniWrite, % g.layout.cols, %tmp%, %section%, Cols
        IniWrite, % g.layout.rows, %tmp%, %section%, Rows
        IniWrite, % g.monitor, %tmp%, %section%, Monitor
        IniWrite, % g.gapX, %tmp%, %section%, GapX
        IniWrite, % g.gapY, %tmp%, %section%, GapY
        IniWrite, % g.scale, %tmp%, %section%, Scale

        if (g.HasKey("fixed")) {
            idx := 0
            for t, xywh in g.fixed {
                idx++
                IniWrite, % t "|||" xywh, %tmp%, %section%, Fixed_%idx%
            }
        }
    }

    ; swap in atomically
    FileDelete, %CFG_FILE%
    FileMove, %tmp%, %CFG_FILE%, 1
}

; ============================================================
; LaunchTiles_NoClose — flexible grid overrides + clean math
; ============================================================
LaunchTiles_NoClose(__tiles, __colsOverride := 0, __rowsOverride := 0) {
    global TARGET_MONITOR, GAP_X, GAP_Y, MIN_W, MIN_H
    global OTR_PATH, COMMON_FLAGS, USE_REGION_SCALE, REGION_SCALE, REGION_FLIP_Y, DEBUG, HIDE_FROM_ALTTAB

    __pids := []
    __n := __tiles.MaxIndex()
    if (__n = "" || __n = 0)
        return __pids

    _GetMonitorWA(TARGET_MONITOR, __mL, __mT, __mR, __mB)
    __mW := __mR - __mL
    __mH := __mB - __mT

    ; ---- Decide rows/cols (accept cols-only or rows-only) ----
    if (__colsOverride > 0 || __rowsOverride > 0) {
        if (__colsOverride > 0 && __rowsOverride > 0) {
            __cols := __colsOverride, __rows := __rowsOverride
        } else if (__colsOverride > 0) {
            __cols := __colsOverride
            __rows := Ceil((__n + __cols - 1) / __cols)
        } else {
            __rows := __rowsOverride
            __cols := Ceil((__n + __rows - 1) / __rows)
        }
    } else {
        __cols := Ceil(Sqrt(__n))
        __rows := Ceil((__n + __cols - 1) / __cols)
    }

    ; ---- Per-cell size (compute once) ----
    __tileW := Floor((__mW - GAP_X * (__cols - 1)) / __cols)
    __tileH := Floor((__mH - GAP_Y * (__rows - 1)) / __rows)
    if (__tileW < MIN_W) __tileW := MIN_W
    if (__tileH < MIN_H) __tileH := MIN_H

    __r := 0, __c := 0
    Loop, % __n {
        __title  := __tiles[A_Index].title
        __region := __tiles[A_Index].region

        StringSplit, __R, __region, `,
        __rX := __R1, __rY := __R2, __rW := __R3, __rH := __R4

        if (USE_REGION_SCALE) {
            __useW := Round(__rW * REGION_SCALE)
            __useH := Round(__rH * REGION_SCALE)
        } else {
            __useW := __tileW
            __useH := __tileH
        }

        __rY_use := __rY
        if (REGION_FLIP_Y) {
            WinGet, __hSrc, ID, %__title%
            if (__hSrc) {
                VarSetCapacity(__rc2, 16)
                DllCall("GetClientRect","ptr",__hSrc,"ptr",&__rc2)
                __clientH := NumGet(__rc2, 12, "Int")
                if (__clientH) {
                    __rY_use := __clientH - __rY - __rH
                    if (__rY_use < 0)
                        __rY_use := 0
                }
            }
        }

        __strideW := __useW + GAP_X
        __strideH := __useH + GAP_Y

        __x := __mL + __c * __strideW
        __y := __mT + __r * __strideH

        ; Ensure the source is visible/restored before cloning
        EnsureSourceReadyExact(__title)

        __args := "--windowTitle=" . Chr(34) . __title . Chr(34)
        __args .= " " . COMMON_FLAGS
        __args .= " --region="   . __rX "," . __rY_use "," . __rW "," . __rH
        __args .= " --size="     . __useW "," . __useH      ; <— use __useW/__useH
        __args .= " --position=" . __x "," . __y            ; <— uses the new __x/__y

        __newPID := _robustLaunchOTR(__args, 4000, 4, __title)
        if (__newPID) {
            ; snap any 0,0 spawns back to their grid slot (no focus steal)
            WinWait, ahk_pid %__newPID%,, 4000
            if (!ErrorLevel) {
                WinGet, _list, List, ahk_pid %__newPID%
                if (_list >= 1) {
                    _main := _list1
                    DllCall("SetWindowPos","ptr",_main,"ptr",0
                        ,"int",__x,"int",__y,"int",0,"int",0,"uint",0x0015)  ; NOSIZE|NOZORDER|NOACTIVATE
                }
            }

            __pids.Push(__newPID)
            if (HIDE_FROM_ALTTAB)
                SetToolWindowStyle_ForPid(__newPID)
        } else {
            Log("OTR launch FAILED after retries for title: " . __title)
        }

        ; pacing to avoid DWM pileups
        Sleep, 220
        if (Mod(A_Index, 4) = 0)
            Sleep, 600

        __c++
        if (__c >= __cols) {
            __c := 0
            __r++
        }
    }
    return __pids
}


CloseAllClones() {
    global launchedPIDs, GroupPIDs
    ; close “simple” launch clones
    for _, pid in launchedPIDs
        ClosePidWindows(pid)
    launchedPIDs := []

    ; close per-group clones
    for name, arr in GroupPIDs {
        for _, obj in arr
            ClosePidWindows(obj.pid)
    }
    GroupPIDs := {}
    Overlay_DestroyAll()
}

ClosePidWindows(pid) {
    ; destroy overlays tied to this PID (handles multiple windows per pid)
    WinGet, list, List, ahk_pid %pid%
    Loop, %list% {
        hwnd := list%A_Index%
        Overlay_Destroy(hwnd)                  ; remove overlay bound to this clone window
        PostMessage, 0x0010, 0, 0,, ahk_id %hwnd%  ; WM_CLOSE (graceful)
    }
    Sleep, 200                                 ; let OTR exit cleanly
    Process, Exist, %pid%
    if (ErrorLevel)
        Process, Close, %pid%                  ; hard kill only if still alive
}

CloseGroupClones(name) {
    global GroupPIDs
    if (!GroupPIDs.HasKey(name))
        return
    for _, obj in GroupPIDs[name]
        ClosePidWindows(obj.pid)   ; this already destroys overlays per window
    GroupPIDs.Delete(name)
}

CloseCheckedGroups() {
    global Manager_hLV_Groups, GroupPIDs
    Gui, ListView, %Manager_hLV_Groups%
    idx := 0
    did := 0
    while (idx := LV_GetNext(idx, "C")) {
        LV_GetText(name, idx)
        CloseGroupClones(name), did := 1
    }
    ; fallback: if none checked, close the selected row
    if (!did) {
        row := LV_GetNext()
        if (row) {
            LV_GetText(name, row)
            CloseGroupClones(name)
        }
    }
}

; ============================================================
; BurstLaunch_NoClose — fast like LaunchAll, but does NOT close
; ============================================================
BurstLaunch_NoClose(__tiles, __colsOverride := 0, __rowsOverride := 0) {
    global TARGET_MONITOR, GAP_X, GAP_Y, MIN_W, MIN_H
    global OTR_PATH, COMMON_FLAGS, USE_REGION_SCALE, REGION_SCALE, REGION_FLIP_Y, HIDE_FROM_ALTTAB
    global LABEL_COLOR, LABEL_OFF_X, LABEL_OFF_Y

    __pids := []
    __n := __tiles.MaxIndex()
    if (__n = "" || __n = 0)
        return __pids

    _GetMonitorWA(TARGET_MONITOR, __mL, __mT, __mR, __mB)
    __mW := __mR - __mL
    __mH := __mB - __mT

    ; rows/cols (accept overrides)
    if (__colsOverride > 0 && __rowsOverride > 0) {
        __cols := __colsOverride, __rows := __rowsOverride
    } else {
        __cols := Ceil(Sqrt(__n)), __rows := Ceil((__n + __cols - 1) / __cols)
    }

    __tileW := Floor((__mW - GAP_X * (__cols - 1)) / __cols)
    __tileH := Floor((__mH - GAP_Y * (__rows - 1)) / __rows)
    if (__tileW < MIN_W) __tileW := MIN_W
    if (__tileH < MIN_H) __tileH := MIN_H

    pinfo := []   ; {pid, title, x, y}
    __r := 0, __c := 0

    ; -------- Pass 1: fire them all quickly ----------
    Loop, % __n {
        __title  := __tiles[A_Index].title
        __region := __tiles[A_Index].region

        StringSplit, __R, __region, `,
        __rX := __R1, __rY := __R2, __rW := __R3, __rH := __R4

        if (USE_REGION_SCALE) {
            __outW := Round(__rW * REGION_SCALE), __outH := Round(__rH * REGION_SCALE)
        } else {
            __outW := __tileW, __outH := __tileH
        }

        __rY_use := __rY
        if (REGION_FLIP_Y) {
            WinGet, __hSrc, ID, %__title%
            if (__hSrc) {
                VarSetCapacity(__rc2, 16)
                DllCall("GetClientRect","ptr",__hSrc,"ptr",&__rc2)
                __clientH := NumGet(__rc2,12,"Int")
                if (__clientH)
                    __rY_use := Max(0, __clientH - __rY - __rH)
            }
        }

        ; cell origin using actual clone size (so large clones stride properly)
        __strideW := __outW + GAP_X
        __strideH := __outH + GAP_Y
        __x := __mL + __c * __strideW
        __y := __mT + __r * __strideH

        ; make sure source is visible
        EnsureSourceReadyExact(__title)

        __args := "--windowTitle=" . Chr(34) . __title . Chr(34)
        __args .= " " . COMMON_FLAGS
        __args .= " --region="   . __rX "," . __rY_use "," . __rW "," . __rH
        __args .= " --size="     . __outW "," . __outH
        __args .= " --position=" . __x "," . __y

        Run, % Chr(34) . OTR_PATH . Chr(34) . " " . __args,, Hide, __pid
        if (__pid) {
            __pids.Push(__pid)
            pinfo.Push({pid:__pid, title:__title, x:__x, y:__y})
            if (HIDE_FROM_ALTTAB)
                SetToolWindowStyle_ForPid(__pid)
        }
        ; tiny nudge so the shell doesn’t choke, but still “burst”
        Sleep, 120

        __c++
        if (__c >= __cols)
            __c := 0, __r++
    }

    ; -------- Pass 2: rename + attach overlay + snap (no focus steal) ----------
    for _, o in pinfo {
        pid := o.pid
        WinWait, ahk_pid %pid%, , 600
        if (ErrorLevel)
            continue
        WinGet, __list, List, ahk_pid %pid%
        if (__list >= 1) {
            __main := __list1
            __clean := ExtractCharName(o.title)
            Loop, %__list% {
                __h := __list%A_Index%
                WinSetTitle, ahk_id %__h%,, %__clean%
            }
            ; snap in case OTR ignored --position
            DllCall("SetWindowPos","ptr",__main,"ptr",0
                ,"int",o.x,"int",o.y,"int",0,"int",0,"uint",0x0015) ; NOSIZE|NOZORDER|NOACTIVATE

            Overlay_Attach(__main, __clean, LABEL_COLOR, LABEL_OFF_X, LABEL_OFF_Y)
        }
    }
    return __pids
}


LoadConfig() {
    global CFG_FILE, Groups, DEF_GROUP
    global TARGET_MONITOR, GAP_X, GAP_Y, REGION_SCALE, AUTO_RETILE

    if (CFG_FILE = "")
        CFG_FILE := ResolveConfigPath()

    Groups := {}

    ; ---- General settings ----
    if FileExist(CFG_FILE) {
        IniRead, ar, %CFG_FILE%, General, AutoRetile, 1
        AUTO_RETILE := (ar+0) ? 1 : 0
    } else {
        AUTO_RETILE := 1
        ; Seed a default group if no config exists
        Groups["Main"] := CloneDefaultGroup()
        Log("LoadConfig: NO config file at " . CFG_FILE . " -> seeded default 'Main'")
        return
    }

    ; ---- Parse groups ----
    FileRead, __cfg, %CFG_FILE%
    sect := ""
    count := 0

    Loop, Parse, __cfg, `n, `r
    {
        line := Trim(A_LoopField)
        if (line = "")
            continue

        if (SubStr(line,1,1)="[" && InStr(line, "Group:")) {
            sect := SubStr(line, 2, StrLen(line)-2) ; e.g. [Group:Capacitors]
            name := SubStr(sect, 7)                 ; strip "Group:"
            if (name = "")
                continue

            IniRead, en,   %CFG_FILE%, %sect%, Enabled, 0
            IniRead, mem,  %CFG_FILE%, %sect%, Members,
            IniRead, reg,  %CFG_FILE%, %sect%, Region,
            IniRead, lay,  %CFG_FILE%, %sect%, Layout, auto
            IniRead, cols, %CFG_FILE%, %sect%, Cols, 0
            IniRead, rows, %CFG_FILE%, %sect%, Rows, 0
            IniRead, mon,  %CFG_FILE%, %sect%, Monitor, %TARGET_MONITOR%
            IniRead, gx,   %CFG_FILE%, %sect%, GapX, %GAP_X%
            IniRead, gy,   %CFG_FILE%, %sect%, GapY, %GAP_Y%
            IniRead, sc,   %CFG_FILE%, %sect%, Scale, %REGION_SCALE%

            g := {}
            g.enabled := en+0
            g.members := SplitPipe(mem)
            g.region  := reg
            g.layout  := { mode: lay, cols: cols+0, rows: rows+0 }
            g.monitor := mon+0
            g.gapX    := gx+0
            g.gapY    := gy+0
            g.scale   := sc+0
            g.fixed   := {}

            ; ---- Normalize Fixed_* keys to cleaned, lowercase names ----
            ; Accept legacy entries saved as full titles (e.g. "EVE - Name")
            ; and convert to the same key LaunchGroup() uses (clean+lower).
            Loop, 200 {
                IniRead, fx, %CFG_FILE%, %sect%, Fixed_%A_Index%,
                if (fx = "" || fx = "ERROR")
                    continue
                parts := StrSplit(fx, "|||")
                raw := Trim(parts[1])     ; could be "EVE - Cassablanca" or just "Cassablanca"
                xywh := Trim(parts[2])    ; "x,y,w,h"
                if (xywh = "")
                    continue

                key := ExtractCharName(raw)
                StringLower, key, key
                if (key != "")
                    g.fixed[key] := xywh
            }

            Groups[name] := g
            count++
        }
    }

    Log("LoadConfig: loaded " . count . " group(s) from " . CFG_FILE)
}


JoinPipe(arr) {
    out := ""
    for i, v in arr {
        out .= (i=1 ? "" : "|") v
    }
    return out
}
SplitPipe(s) {
    arr := []
    if (s = "" || s = "ERROR")
        return arr
    Loop, Parse, s, |
        arr.Push(A_LoopField)
    return arr
}

; ============================================================
; LaunchGroup — uses group scale, flexible grid, saved layouts
; ============================================================
LaunchGroup(name) {
    global IsLaunching
    global Groups, TARGET_MONITOR, GAP_X, GAP_Y, USE_REGION_SCALE, REGION_SCALE
    global GroupPIDs, MGR_Accumulate, HIDE_FROM_ALTTAB

    if (IsLaunching)
        return
    IsLaunching := 1
    StopAutoRetile()

    if (!Groups.HasKey(name)) {
        MsgBox, 48, Missing group, Group "%name%" not found.
        StartAutoRetile(), IsLaunching := 0
        return
    }
    g := Groups[name]
    if (g.members.Length() = 0) {
        MsgBox, 48, Empty group, Add members to the group first.
        StartAutoRetile(), IsLaunching := 0
        return
    }

    region := g.region
    if (region = "") {
        region := PickRegion_WithGuard()
        if (region = "") {
            StartAutoRetile(), IsLaunching := 0
            return
        }
        g.region := region
        Groups[name] := g
    }

    ; ---- Per-group overrides (restore later) ----
    oldMon   := TARGET_MONITOR
    oldGX    := GAP_X
    oldGY    := GAP_Y
    oldScale := REGION_SCALE

    TARGET_MONITOR   := g.monitor
    GAP_X            := g.gapX
    GAP_Y            := g.gapY
    USE_REGION_SCALE := true
    REGION_SCALE     := g.scale  ; accepts floats like 0.5

    ; ---- Build tiles only from PRESENT clients ----
    tiles := []
    missing := []
    live := BuildLiveClientMap()
    for _, saved in g.members {
        clean := ExtractCharName(saved)
        StringLower, clean, clean
        if (live.HasKey(clean)) {
            tiles.Push({ title: live[clean], region: region })
        } else {
            missing.Push(saved)
        }
    }

    if (tiles.MaxIndex() = "" || tiles.MaxIndex() = 0) {
        TARGET_MONITOR := oldMon
        GAP_X          := oldGX
        GAP_Y          := oldGY
        REGION_SCALE   := oldScale
        if (missing.MaxIndex() != "")
            TrayTip, EVE Tiler, No clients found to tile for "%name%"., 1200, 1
        StartAutoRetile(), IsLaunching := 0
        return
    }

    ; ---- Prefer saved fixed layout if present ----
    useFixed := false
    if (g.HasKey("fixed")) {
        for _k, _v in g.fixed {
            if (_v != "") {
                useFixed := true
                break
            }
        }
    }

    if (useFixed) {
        if (!MGR_Accumulate)
            CloseGroupClones(name)

        out := []
        StringSplit, R, region, `,
        rX := R1, rY := R2, rW := R3, rH := R4

        for _, tile in tiles {
            t := tile.title
            cleanT := ExtractCharName(t)
            if (g.fixed.HasKey(cleanT)) {
                val := g.fixed[cleanT]
                StringSplit, P, val, `,
                x := P1, y := P2, w := P3, h := P4
                pid := LaunchOneClone(t, rX, rY, rW, rH, x, y, w, h)
            } else {
                pid := LaunchOneClone(t, rX, rY, rW, rH
                    , 0, 0
                    , Round(rW*REGION_SCALE), Round(rH*REGION_SCALE))
            }
            if (HIDE_FROM_ALTTAB && pid)
                SetToolWindowStyle_ForPid(pid)
            if (pid)
                out.Push({ pid: pid, title: t })
            Sleep, 120
        }
        GroupPIDs[name] := out

        ; restore + resume
        TARGET_MONITOR := oldMon
        GAP_X          := oldGX
        GAP_Y          := oldGY
        REGION_SCALE   := oldScale
        StartAutoRetile()
        IsLaunching := 0

        if (missing.MaxIndex() != "")
            TrayTip, EVE Tiler, % "Skipped " missing.MaxIndex() " offline"
                . (missing.MaxIndex() ? ": " missing[1] : ""), 1200, 1
        return
    }

    ; ---- Normal tiling (auto or grid with partial overrides) ----
    if (!MGR_Accumulate)
        CloseGroupClones(name)

    cols := 0, rows := 0
    if (g.layout.mode = "grid") {
        cols := g.layout.cols+0
        rows := g.layout.rows+0
    }
    pids := BurstLaunch_NoClose(tiles, cols, rows)

    out := []
    i := 0
    for _, tile in tiles {
        i++
        if (i <= pids.Length())
            out.Push({ pid: pids[i], title: tile.title })
    }
    GroupPIDs[name] := out

    ; restore + resume
    TARGET_MONITOR := oldMon
    GAP_X          := oldGX
    GAP_Y          := oldGY
    REGION_SCALE   := oldScale

    if (missing.MaxIndex() != "")
        TrayTip, EVE Tiler, % "Skipped " missing.MaxIndex() " offline"
            . (missing.MaxIndex() ? ": " missing[1] : ""), 1200, 1

    StartAutoRetile()
    IsLaunching := 0
}
