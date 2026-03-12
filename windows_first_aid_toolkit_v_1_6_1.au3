#RequireAdmin
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <Date.au3>
#include <GuiToolTip.au3>
#include <GuiRichEdit.au3>
#include <ColorConstants.au3>
#include <WinAPISysWin.au3>

Opt("GUIOnEventMode", 0)

Global $g_hGUI = 0
Global $g_hConsoleEdit = 0
Global $g_hToolTip = 0
Global $g_sLogFile = @ScriptDir & "\Logs\Windows_First_Aid_" & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & @MIN & @SEC & ".log"

Global $g_iLastW = 0, $g_iLastH = 0
Global Const $MIN_GUI_W = 760, $MIN_GUI_H = 620
Global $g_iTotalTasks = 0, $g_iDoneTasks = 0

Global $lblTitle = 0, $lblSub = 0
Global $grpPrep = 0, $grpNetwork = 0, $grpCleanup = 0, $grpRepairs = 0, $grpServices = 0, $grpBottom = 0
Global $lblPrepNote = 0, $lblCleanupNote = 0, $lblRepairNote = 0

Global $chkRestorePoint = 0
Global $chkFlushDNS = 0, $chkResetWinsock = 0, $chkRenewIP = 0
Global $chkUserTemp = 0, $chkSystemTemp = 0, $chkBrowserCache = 0
Global $chkSFC = 0, $chkDISM = 0
Global $chkWU = 0, $chkBITS = 0, $chkSpooler = 0

Global $mnuPresets = 0, $mnuPresetRecommended = 0, $mnuPresetNetwork = 0, $mnuPresetCleanup = 0, $mnuPresetRepairs = 0, $mnuPresetServices = 0
Global $mnuActions = 0, $mnuRunSelected = 0, $mnuCheckAll = 0, $mnuClearAll = 0, $mnuSaveLog = 0, $mnuOpenLogs = 0, $mnuClearConsole = 0, $mnuExit = 0
Global $mnuHelp = 0, $mnuHelpTxt = 0, $mnuHelpHtml = 0, $mnuAbout = 0

Global Enum _
    $CONSOLE_NORMAL = 0, _
    $CONSOLE_COMMAND, _
    $CONSOLE_WARN, _
    $CONSOLE_ERROR, _
    $CONSOLE_SUCCESS, _
    $CONSOLE_HEADER

_Main()

Func _Main()
    DirCreate(@ScriptDir & "\Logs")

    $g_hGUI = GUICreate("Windows First Aid Toolkit", 980, 760, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF4F7FA)
    GUIRegisterMsg($WM_GETMINMAXINFO, "WM_GETMINMAXINFO")

    ; Menus
    $mnuPresets = GUICtrlCreateMenu("Presets")
    $mnuPresetRecommended = GUICtrlCreateMenuItem("Recommended", $mnuPresets)
    GUICtrlCreateMenuItem("", $mnuPresets)
    $mnuPresetNetwork = GUICtrlCreateMenuItem("Network repair only", $mnuPresets)
    $mnuPresetCleanup = GUICtrlCreateMenuItem("Cleanup only", $mnuPresets)
    $mnuPresetRepairs = GUICtrlCreateMenuItem("System repairs only", $mnuPresets)
    $mnuPresetServices = GUICtrlCreateMenuItem("Service restarts only", $mnuPresets)

    $mnuActions = GUICtrlCreateMenu("Actions")
    $mnuRunSelected = GUICtrlCreateMenuItem("Run Selected", $mnuActions)
    GUICtrlCreateMenuItem("", $mnuActions)
    $mnuCheckAll = GUICtrlCreateMenuItem("Check All", $mnuActions)
    $mnuClearAll = GUICtrlCreateMenuItem("Clear All", $mnuActions)
    GUICtrlCreateMenuItem("", $mnuActions)
    $mnuSaveLog = GUICtrlCreateMenuItem("Save Log Copy", $mnuActions)
    $mnuOpenLogs = GUICtrlCreateMenuItem("Open Logs Folder", $mnuActions)
    $mnuClearConsole = GUICtrlCreateMenuItem("Clear Console Output", $mnuActions)
    GUICtrlCreateMenuItem("", $mnuActions)
    $mnuExit = GUICtrlCreateMenuItem("Exit", $mnuActions)

    $mnuHelp = GUICtrlCreateMenu("Help")
    $mnuHelpTxt = GUICtrlCreateMenuItem("Open Help TXT", $mnuHelp)
    $mnuHelpHtml = GUICtrlCreateMenuItem("Open Help HTML", $mnuHelp)
    GUICtrlCreateMenuItem("", $mnuHelp)
    $mnuAbout = GUICtrlCreateMenuItem("About", $mnuHelp)

    ; Header
    $lblTitle = GUICtrlCreateLabel("Windows First Aid Toolkit", 20, 16, 330, 24)
    GUICtrlSetFont($lblTitle, 13, 700, 0, "Segoe UI")

    $lblSub = GUICtrlCreateLabel("Safe first-step repair and cleanup actions for common Windows problems", 20, 44, 620, 20)
    GUICtrlSetColor($lblSub, 0x4A5568)

    ; Left / right sections
    $grpPrep = GUICtrlCreateGroup("Preparation", 20, 76, 450, 72)
    $chkRestorePoint = GUICtrlCreateCheckbox("Create restore point first (recommended)", 35, 106, 280, 20)
    GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)
    $lblPrepNote = GUICtrlCreateLabel("System Protection must be enabled on the PC for this to work.", 35, 126, 360, 16)
    GUICtrlSetColor($lblPrepNote, 0x8A1C1C)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    $grpCleanup = GUICtrlCreateGroup("Cleanup", 490, 76, 470, 206)
    $chkUserTemp = GUICtrlCreateCheckbox("Clear user temp folder", 505, 104, 230, 20)
    $chkSystemTemp = GUICtrlCreateCheckbox("Clear system temp folder", 505, 130, 230, 20)
    $chkBrowserCache = GUICtrlCreateCheckbox("Clear browser cache (Edge / Chrome / Firefox)", 505, 156, 330, 20)
    $lblCleanupNote = GUICtrlCreateLabel("Best results when browsers are closed. Locked files will be skipped.", 505, 184, 390, 18)
    GUICtrlSetColor($lblCleanupNote, 0x8A1C1C)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    $grpNetwork = GUICtrlCreateGroup("Network repairs", 20, 160, 450, 122)
    $chkFlushDNS = GUICtrlCreateCheckbox("Flush DNS", 35, 188, 180, 20)
    $chkResetWinsock = GUICtrlCreateCheckbox("Reset Winsock", 35, 214, 180, 20)
    $chkRenewIP = GUICtrlCreateCheckbox("Release and renew IP", 35, 240, 180, 20)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    $grpRepairs = GUICtrlCreateGroup("Windows repairs", 20, 294, 450, 104)
    $chkSFC = GUICtrlCreateCheckbox("Run SFC /scannow", 35, 322, 220, 20)
    $chkDISM = GUICtrlCreateCheckbox("Run DISM RestoreHealth", 35, 348, 220, 20)
    $lblRepairNote = GUICtrlCreateLabel("These checks can take a long time.", 35, 372, 260, 18)
    GUICtrlSetColor($lblRepairNote, 0x8A1C1C)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    $grpServices = GUICtrlCreateGroup("Service restarts", 490, 294, 470, 104)
    $chkWU = GUICtrlCreateCheckbox("Restart Windows Update (wuauserv)", 505, 322, 320, 20)
    $chkBITS = GUICtrlCreateCheckbox("Restart BITS", 505, 348, 180, 20)
    $chkSpooler = GUICtrlCreateCheckbox("Restart Print Spooler", 505, 374, 220, 20)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; Console section
    $grpBottom = GUICtrlCreateGroup("Console Output", 20, 420, 940, 320)
    $g_hConsoleEdit = _GUICtrlRichEdit_Create($g_hGUI, "", 35, 448, 910, 252, BitOR($ES_MULTILINE, $WS_VSCROLL, $WS_HSCROLL, $ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY))
    _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 10, "Consolas")
    _GUICtrlRichEdit_SetBkColor($g_hConsoleEdit, 0x111111)
    _GUICtrlRichEdit_SetReadOnly($g_hConsoleEdit, True)
    _GUICtrlRichEdit_SetEventMask($g_hConsoleEdit, $ENM_SCROLL)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    _InitToolTips()

    GUISetState(@SW_SHOWMAXIMIZED, $g_hGUI)
    _CacheWindowSize()
    _ApplyLayout($g_iLastW, $g_iLastH)

    ; Default state: only restore point checked
    _SetChecks($GUI_UNCHECKED)
    GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)

    _ResetProgress()
    _ConsoleWrite("Windows First Aid Toolkit console initialized.", $CONSOLE_HEADER)
    _ConsoleWrite("Session log: " & $g_sLogFile, $CONSOLE_NORMAL)

    While 1
        _TrackResize()

        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $mnuExit
                ExitLoop

            Case $mnuPresetRecommended
                _ApplyRecommended()

            Case $mnuPresetNetwork
                _ApplyPresetNetwork()

            Case $mnuPresetCleanup
                _ApplyPresetCleanup()

            Case $mnuPresetRepairs
                _ApplyPresetRepairs()

            Case $mnuPresetServices
                _ApplyPresetServices()

            Case $mnuRunSelected
                _RunSelected()

            Case $mnuCheckAll
                _SetChecks($GUI_CHECKED)

            Case $mnuClearAll
                _SetChecks($GUI_UNCHECKED)
                GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)

            Case $mnuSaveLog
                _SaveLogCopy()

            Case $mnuOpenLogs
                ShellExecute(@ScriptDir & "\Logs")

            Case $mnuClearConsole
                _ClearConsole()

            Case $mnuHelpTxt
                ShellExecute(@ScriptDir & "\help.txt")

            Case $mnuHelpHtml
                ShellExecute(@ScriptDir & "\help.html")

            Case $mnuAbout
                _ShowAbout()
        EndSwitch
    WEnd
EndFunc

Func WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
    Local $tMinMaxInfo = DllStructCreate( _
            "int ReservedX;" & _
            "int ReservedY;" & _
            "int MaxSizeX;" & _
            "int MaxSizeY;" & _
            "int MaxPositionX;" & _
            "int MaxPositionY;" & _
            "int MinTrackSizeX;" & _
            "int MinTrackSizeY;" & _
            "int MaxTrackSizeX;" & _
            "int MaxTrackSizeY", _
            $lParam)

    DllStructSetData($tMinMaxInfo, "MinTrackSizeX", $MIN_GUI_W)
    DllStructSetData($tMinMaxInfo, "MinTrackSizeY", $MIN_GUI_H)

    Return 0
EndFunc

Func _CacheWindowSize()
    Local $aClient = WinGetClientSize($g_hGUI)
    If IsArray($aClient) Then
        $g_iLastW = $aClient[0]
        $g_iLastH = $aClient[1]
    Else
        $g_iLastW = 980
        $g_iLastH = 760
    EndIf
EndFunc

Func _TrackResize()
    Local $aClient = WinGetClientSize($g_hGUI)
    If Not IsArray($aClient) Then Return

    Local $iW = $aClient[0]
    Local $iH = $aClient[1]

    If $iW <> $g_iLastW Or $iH <> $g_iLastH Then
        $g_iLastW = $iW
        $g_iLastH = $iH
        _ApplyLayout($g_iLastW, $g_iLastH)
    EndIf
EndFunc

Func _ApplyLayout($iW, $iH)
    Local $margin = 20
    Local $gap = 20
    Local $consoleTop = 420
    Local $bottomMargin = 20
    Local $minBottomPanelH = 160
    Local $bottomPanelH = $iH - $consoleTop - $bottomMargin
    If $bottomPanelH < $minBottomPanelH Then $bottomPanelH = $minBottomPanelH

    Local $topSectionY1 = 76
    Local $topSectionY2 = 160
    Local $topSectionY3 = 294

    Local $colW = Int(($iW - ($margin * 2) - $gap) / 2)
    If $colW < 350 Then $colW = 350

    Local $leftX = $margin
    Local $rightX = $margin + $colW + $gap

    GUICtrlSetPos($lblTitle, 20, 16, $iW - 40, 24)
    GUICtrlSetPos($lblSub, 20, 44, $iW - 40, 20)

    GUICtrlSetPos($grpPrep, $leftX, $topSectionY1, $colW, 72)
    GUICtrlSetPos($chkRestorePoint, $leftX + 15, 106, $colW - 30, 20)
    GUICtrlSetPos($lblPrepNote, $leftX + 15, 126, $colW - 30, 16)

    GUICtrlSetPos($grpCleanup, $rightX, $topSectionY1, $colW, 206)
    GUICtrlSetPos($chkUserTemp, $rightX + 15, 104, $colW - 30, 20)
    GUICtrlSetPos($chkSystemTemp, $rightX + 15, 130, $colW - 30, 20)
    GUICtrlSetPos($chkBrowserCache, $rightX + 15, 156, $colW - 30, 20)
    GUICtrlSetPos($lblCleanupNote, $rightX + 15, 184, $colW - 30, 18)

    GUICtrlSetPos($grpNetwork, $leftX, $topSectionY2, $colW, 122)
    GUICtrlSetPos($chkFlushDNS, $leftX + 15, 188, $colW - 30, 20)
    GUICtrlSetPos($chkResetWinsock, $leftX + 15, 214, $colW - 30, 20)
    GUICtrlSetPos($chkRenewIP, $leftX + 15, 240, $colW - 30, 20)

    GUICtrlSetPos($grpRepairs, $leftX, $topSectionY3, $colW, 104)
    GUICtrlSetPos($chkSFC, $leftX + 15, 322, $colW - 30, 20)
    GUICtrlSetPos($chkDISM, $leftX + 15, 348, $colW - 30, 20)
    GUICtrlSetPos($lblRepairNote, $leftX + 15, 372, $colW - 30, 18)

    GUICtrlSetPos($grpServices, $rightX, $topSectionY3, $colW, 104)
    GUICtrlSetPos($chkWU, $rightX + 15, 322, $colW - 30, 20)
    GUICtrlSetPos($chkBITS, $rightX + 15, 348, $colW - 30, 20)
    GUICtrlSetPos($chkSpooler, $rightX + 15, 374, $colW - 30, 20)

    GUICtrlSetPos($grpBottom, 20, $consoleTop, $iW - 40, $bottomPanelH)
    _WinAPI_MoveWindow($g_hConsoleEdit, 35, $consoleTop + 28, $iW - 70, $bottomPanelH - 45, True)
EndFunc

Func _ApplyPresetNetwork()
    _SetChecks($GUI_UNCHECKED)
    GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)
    GUICtrlSetState($chkFlushDNS, $GUI_CHECKED)
    GUICtrlSetState($chkResetWinsock, $GUI_CHECKED)
    GUICtrlSetState($chkRenewIP, $GUI_CHECKED)
    _ConsoleWrite("Preset applied: Network repair only.", $CONSOLE_SUCCESS)
EndFunc

Func _ApplyPresetCleanup()
    _SetChecks($GUI_UNCHECKED)
    GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)
    GUICtrlSetState($chkUserTemp, $GUI_CHECKED)
    GUICtrlSetState($chkSystemTemp, $GUI_CHECKED)
    GUICtrlSetState($chkBrowserCache, $GUI_CHECKED)
    _ConsoleWrite("Preset applied: Cleanup only.", $CONSOLE_SUCCESS)
EndFunc

Func _ApplyPresetRepairs()
    _SetChecks($GUI_UNCHECKED)
    GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)
    GUICtrlSetState($chkSFC, $GUI_CHECKED)
    GUICtrlSetState($chkDISM, $GUI_CHECKED)
    _ConsoleWrite("Preset applied: System repairs only.", $CONSOLE_SUCCESS)
EndFunc

Func _ApplyPresetServices()
    _SetChecks($GUI_UNCHECKED)
    GUICtrlSetState($chkRestorePoint, $GUI_UNCHECKED)
    GUICtrlSetState($chkWU, $GUI_CHECKED)
    GUICtrlSetState($chkBITS, $GUI_CHECKED)
    GUICtrlSetState($chkSpooler, $GUI_CHECKED)
    _ConsoleWrite("Preset applied: Service restarts only.", $CONSOLE_SUCCESS)
EndFunc

Func _ApplyRecommended()
    _SetChecks($GUI_UNCHECKED)
    GUICtrlSetState($chkRestorePoint, $GUI_CHECKED)
    GUICtrlSetState($chkFlushDNS, $GUI_CHECKED)
    GUICtrlSetState($chkResetWinsock, $GUI_CHECKED)
    GUICtrlSetState($chkRenewIP, $GUI_CHECKED)
    GUICtrlSetState($chkUserTemp, $GUI_CHECKED)
    GUICtrlSetState($chkSystemTemp, $GUI_CHECKED)
    GUICtrlSetState($chkBrowserCache, $GUI_CHECKED)
    GUICtrlSetState($chkWU, $GUI_CHECKED)
    GUICtrlSetState($chkBITS, $GUI_CHECKED)
    GUICtrlSetState($chkSpooler, $GUI_CHECKED)
    _ConsoleWrite("Preset applied: Recommended.", $CONSOLE_SUCCESS)
EndFunc

Func _SetChecks($iState)
    GUICtrlSetState($chkFlushDNS, $iState)
    GUICtrlSetState($chkResetWinsock, $iState)
    GUICtrlSetState($chkRenewIP, $iState)
    GUICtrlSetState($chkUserTemp, $iState)
    GUICtrlSetState($chkSystemTemp, $iState)
    GUICtrlSetState($chkBrowserCache, $iState)
    GUICtrlSetState($chkSFC, $iState)
    GUICtrlSetState($chkDISM, $iState)
    GUICtrlSetState($chkWU, $iState)
    GUICtrlSetState($chkBITS, $iState)
    GUICtrlSetState($chkSpooler, $iState)
EndFunc

Func _RunSelected()
    Local $iCount = _CountTasks()
    If $iCount = 0 And Not _IsChecked($chkRestorePoint) Then
        MsgBox($MB_ICONINFORMATION, "Nothing selected", "Select at least one action.")
        Return
    EndIf

    Local $sWarn = "The selected repair tasks are about to run." & @CRLF & @CRLF
    If _IsChecked($chkSFC) Or _IsChecked($chkDISM) Then $sWarn &= "- SFC and DISM can take a long time." & @CRLF
    If _IsChecked($chkBrowserCache) Then $sWarn &= "- Browser cleanup is best done with browsers closed." & @CRLF
    $sWarn &= "- Some locked files may be skipped." & @CRLF & @CRLF & "Continue?"
    If MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO), "Confirm repair actions", $sWarn) <> $IDYES Then Return

    $g_iDoneTasks = 0
    $g_iTotalTasks = $iCount
    If _IsChecked($chkRestorePoint) Then $g_iTotalTasks += 1
    _ResetProgress()

    GUISetState(@SW_DISABLE, $g_hGUI)
    _ConsoleWrite("==================================================", $CONSOLE_HEADER)
    _ConsoleWrite("Run started by " & @UserName & " on " & @ComputerName, $CONSOLE_HEADER)

    If _IsChecked($chkRestorePoint) Then _CreateRestorePoint()
    If _IsChecked($chkFlushDNS) Then _RunTask("Flush DNS", 'ipconfig /flushdns')
    If _IsChecked($chkResetWinsock) Then _RunTask("Reset Winsock", 'netsh winsock reset')
    If _IsChecked($chkRenewIP) Then
        _RunTask("Release IP", 'ipconfig /release')
        _RunTask("Renew IP", 'ipconfig /renew')
    EndIf

    If _IsChecked($chkUserTemp) Then _ClearFolder(@TempDir, "User temp")
    If _IsChecked($chkSystemTemp) Then _ClearFolder(@WindowsDir & "\Temp", "System temp")
    If _IsChecked($chkBrowserCache) Then _ClearBrowserCaches()

    If _IsChecked($chkSFC) Then _RunTask("System File Checker", 'sfc /scannow')
    If _IsChecked($chkDISM) Then _RunTask("DISM RestoreHealth", 'DISM /Online /Cleanup-Image /RestoreHealth')

    If _IsChecked($chkWU) Then _RestartService("wuauserv", "Windows Update")
    If _IsChecked($chkBITS) Then _RestartService("BITS", "BITS")
    If _IsChecked($chkSpooler) Then _RestartService("Spooler", "Print Spooler")

    _ConsoleWrite("Run completed.", $CONSOLE_SUCCESS)
    _ConsoleWrite("==================================================", $CONSOLE_HEADER)
    GUISetState(@SW_ENABLE, $g_hGUI)
    WinActivate($g_hGUI)
    MsgBox($MB_ICONINFORMATION, "Done", "Selected actions have completed." & @CRLF & @CRLF & "Log file:" & @CRLF & $g_sLogFile)
EndFunc

Func _CountTasks()
    Local $i = 0
    If _IsChecked($chkFlushDNS) Then $i += 1
    If _IsChecked($chkResetWinsock) Then $i += 1
    If _IsChecked($chkRenewIP) Then $i += 2
    If _IsChecked($chkUserTemp) Then $i += 1
    If _IsChecked($chkSystemTemp) Then $i += 1
    If _IsChecked($chkBrowserCache) Then $i += 1
    If _IsChecked($chkSFC) Then $i += 1
    If _IsChecked($chkDISM) Then $i += 1
    If _IsChecked($chkWU) Then $i += 1
    If _IsChecked($chkBITS) Then $i += 1
    If _IsChecked($chkSpooler) Then $i += 1
    Return $i
EndFunc

Func _IsChecked($idCtrl)
    Return BitAND(GUICtrlRead($idCtrl), $GUI_CHECKED) = $GUI_CHECKED
EndFunc

Func _RunTask($sName, $sCommand)
    _ConsoleWrite("> " & $sCommand, $CONSOLE_COMMAND)
    _ConsoleWrite("[START] " & $sName, $CONSOLE_HEADER)

    Local $sOutput = _RunCapture($sCommand)
    If StringStripWS($sOutput, 8) = "" Then
        _ConsoleWrite("[INFO] No command output returned.", $CONSOLE_WARN)
    Else
        _ConsoleWriteBlock(StringStripCR($sOutput), $CONSOLE_NORMAL)
    EndIf

    _ConsoleWrite("[END] " & $sName, $CONSOLE_HEADER)
    _TaskDone()
EndFunc

Func _RunCapture($sCommand)
    Local $iPID = Run(@ComSpec & " /c " & $sCommand, @SystemDir, @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
    If @error Or $iPID = 0 Then Return "Failed to start command: " & $sCommand

    Local $sData = ""
    While 1
        $sData &= StdoutRead($iPID)
        $sData &= StderrRead($iPID)
        If Not ProcessExists($iPID) Then ExitLoop
        Sleep(150)
    WEnd
    $sData &= StdoutRead($iPID)
    $sData &= StderrRead($iPID)
    Return $sData
EndFunc

Func _ClearFolder($sFolder, $sLabel)
    _ConsoleWrite("[START] " & $sLabel & " cleanup", $CONSOLE_HEADER)

    If Not FileExists($sFolder) Then
        _ConsoleWrite("[WARN] Folder does not exist: " & $sFolder, $CONSOLE_WARN)
        _ConsoleWrite("[END] " & $sLabel & " cleanup", $CONSOLE_HEADER)
        _TaskDone()
        Return
    EndIf

    Local $iDeleted = 0, $iSkipped = 0
    _DeleteContentsRecursive($sFolder, $iDeleted, $iSkipped)

    _ConsoleWrite("Target: " & $sFolder, $CONSOLE_NORMAL)
    _ConsoleWrite("Deleted: " & $iDeleted & " | Skipped: " & $iSkipped, $CONSOLE_NORMAL)
    _ConsoleWrite("[END] " & $sLabel & " cleanup", $CONSOLE_HEADER)
    _TaskDone()
EndFunc

Func _DeleteContentsRecursive($sFolder, ByRef $iDeleted, ByRef $iSkipped)
    Local $hSearch = FileFindFirstFile($sFolder & "\*")
    If $hSearch = -1 Then Return

    While 1
        Local $sItem = FileFindNextFile($hSearch)
        If @error Then ExitLoop
        If $sItem = "." Or $sItem = ".." Then ContinueLoop

        Local $sFull = $sFolder & "\" & $sItem
        Local $sAttr = FileGetAttrib($sFull)
        If @error Then
            $iSkipped += 1
            ContinueLoop
        EndIf

        If StringInStr($sAttr, "D") Then
            _DeleteContentsRecursive($sFull, $iDeleted, $iSkipped)
            DirRemove($sFull, 1)
            If @error Then
                $iSkipped += 1
            Else
                $iDeleted += 1
            EndIf
        Else
            FileDelete($sFull)
            If @error Then
                $iSkipped += 1
            Else
                $iDeleted += 1
            EndIf
        EndIf
    WEnd

    FileClose($hSearch)
EndFunc

Func _ClearBrowserCaches()
    _ConsoleWrite("[START] Browser cache cleanup", $CONSOLE_HEADER)

    Local $aBrowsers[3] = ["msedge.exe", "chrome.exe", "firefox.exe"]
    For $i = 0 To UBound($aBrowsers) - 1
        If ProcessExists($aBrowsers[$i]) Then _
            _ConsoleWrite("[WARN] Browser appears open: " & $aBrowsers[$i], $CONSOLE_WARN)
    Next

    Local $aPaths[9] = [ _
        @LocalAppDataDir & "\Microsoft\Edge\User Data\Default\Cache\Cache_Data", _
        @LocalAppDataDir & "\Microsoft\Edge\User Data\Default\Code Cache", _
        @LocalAppDataDir & "\Microsoft\Edge\User Data\Default\Service Worker\CacheStorage", _
        @LocalAppDataDir & "\Google\Chrome\User Data\Default\Cache\Cache_Data", _
        @LocalAppDataDir & "\Google\Chrome\User Data\Default\Code Cache", _
        @LocalAppDataDir & "\Google\Chrome\User Data\Default\Service Worker\CacheStorage", _
        @AppDataDir & "\Mozilla\Firefox\Profiles", _
        @LocalAppDataDir & "\Microsoft\Windows\INetCache", _
        @LocalAppDataDir & "\Microsoft\Windows\WebCache" ]

    For $i = 0 To UBound($aPaths) - 1
        If StringRight($aPaths[$i], 8) = "Profiles" Then
            _ClearFirefoxProfiles($aPaths[$i])
        Else
            If FileExists($aPaths[$i]) Then
                Local $d = 0, $s = 0
                _DeleteContentsRecursive($aPaths[$i], $d, $s)
                _ConsoleWrite("Cleaned: " & $aPaths[$i] & " | Deleted: " & $d & " | Skipped: " & $s, $CONSOLE_NORMAL)
            Else
                _ConsoleWrite("Skipped missing path: " & $aPaths[$i], $CONSOLE_WARN)
            EndIf
        EndIf
    Next

    RunWait(@ComSpec & ' /c RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8', @SystemDir, @SW_HIDE)
    _ConsoleWrite("Requested Internet temporary files cleanup.", $CONSOLE_SUCCESS)
    _ConsoleWrite("[END] Browser cache cleanup", $CONSOLE_HEADER)
    _TaskDone()
EndFunc

Func _ClearFirefoxProfiles($sProfilesRoot)
    If Not FileExists($sProfilesRoot) Then
        _ConsoleWrite("Skipped missing Firefox profile root: " & $sProfilesRoot, $CONSOLE_WARN)
        Return
    EndIf

    Local $hSearch = FileFindFirstFile($sProfilesRoot & "\*")
    If $hSearch = -1 Then Return

    While 1
        Local $sProfile = FileFindNextFile($hSearch)
        If @error Then ExitLoop

        Local $sProfileDir = $sProfilesRoot & "\" & $sProfile
        If Not StringInStr(FileGetAttrib($sProfileDir), "D") Then ContinueLoop

        Local $aFirefoxTargets[3] = [ _
            $sProfileDir & "\cache2", _
            $sProfileDir & "\startupCache", _
            $sProfileDir & "\storage\default" ]

        For $i = 0 To UBound($aFirefoxTargets) - 1
            If FileExists($aFirefoxTargets[$i]) Then
                Local $d = 0, $s = 0
                _DeleteContentsRecursive($aFirefoxTargets[$i], $d, $s)
                _ConsoleWrite("Firefox cleaned: " & $aFirefoxTargets[$i] & " | Deleted: " & $d & " | Skipped: " & $s, $CONSOLE_NORMAL)
            EndIf
        Next
    WEnd

    FileClose($hSearch)
EndFunc

Func _RestartService($sServiceName, $sLabel)
    _ConsoleWrite("> net stop " & $sServiceName & " /y", $CONSOLE_COMMAND)
    _ConsoleWrite("> net start " & $sServiceName, $CONSOLE_COMMAND)
    _ConsoleWrite("[START] Restart service: " & $sLabel, $CONSOLE_HEADER)

    _RunCapture('net stop "' & $sServiceName & '" /y')
    Sleep(800)
    _RunCapture('net start "' & $sServiceName & '"')

    _ConsoleWrite("[END] Restart service: " & $sLabel, $CONSOLE_HEADER)
    _TaskDone()
EndFunc

Func _CreateRestorePoint()
    _ConsoleWrite("> powershell Checkpoint-Computer ...", $CONSOLE_COMMAND)
    _ConsoleWrite("[START] Create restore point", $CONSOLE_HEADER)

    Local $sPS = 'powershell -ExecutionPolicy Bypass -Command "try { Checkpoint-Computer -Description ''Windows First Aid Toolkit'' -RestorePointType ''MODIFY_SETTINGS''; Write-Output ''Restore point created successfully.'' } catch { Write-Output $_.Exception.Message; exit 1 }"'
    Local $sOutput = _RunCapture($sPS)
    If StringStripWS($sOutput, 8) = "" Then $sOutput = "No output returned."

    _ConsoleWriteBlock(StringStripCR($sOutput), $CONSOLE_NORMAL)
    _ConsoleWrite("[END] Create restore point", $CONSOLE_HEADER)
    _TaskDone()
EndFunc

Func _TaskDone()
    $g_iDoneTasks += 1
EndFunc

Func _ResetProgress()
    $g_iDoneTasks = 0
    $g_iTotalTasks = 0
EndFunc

Func _SaveLogCopy()
    Local $sTarget = FileSaveDialog("Save log copy", @ScriptDir, "Log files (*.log)|Text files (*.txt)", 2, "Windows_First_Aid_Log.txt")
    If @error Then Return

    FileCopy($g_sLogFile, $sTarget, 9)
    If @error Then
        MsgBox($MB_ICONERROR, "Save log copy", "Could not save a copy of the log.")
    Else
        MsgBox($MB_ICONINFORMATION, "Save log copy", "Log copy saved to:" & @CRLF & $sTarget)
    EndIf
EndFunc

Func _ConsoleWrite($sText, $iStyle = $CONSOLE_NORMAL)
    Local $sLine = "[" & @HOUR & ":" & @MIN & ":" & @SEC & "] " & $sText & @CRLF

    _GUICtrlRichEdit_SetSel($g_hConsoleEdit, -1, -1)
    _SetConsoleStyle($iStyle)
    _GUICtrlRichEdit_AppendText($g_hConsoleEdit, $sLine)
    _GUICtrlRichEdit_Deselect($g_hConsoleEdit)
    _GUICtrlRichEdit_ScrollToCaret($g_hConsoleEdit)

    Local $hFile = FileOpen($g_sLogFile, 1 + 8)
    If $hFile <> -1 Then
        FileWrite($hFile, $sLine)
        FileClose($hFile)
    EndIf
EndFunc

Func _ConsoleWriteBlock($sText, $iStyle = $CONSOLE_NORMAL)
    Local $aLines = StringSplit(StringReplace($sText, @CRLF, @LF), @LF, 1)
    If @error Then
        _ConsoleWrite($sText, $iStyle)
        Return
    EndIf

    For $i = 1 To $aLines[0]
        If $aLines[$i] <> "" Then _ConsoleWrite($aLines[$i], $iStyle)
    Next
EndFunc

Func _SetConsoleStyle($iStyle)
    Switch $iStyle
        Case $CONSOLE_COMMAND
            _GUICtrlRichEdit_SetCharColor($g_hConsoleEdit, 0x7CFC00)
            _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 10, "Consolas")

        Case $CONSOLE_WARN
            _GUICtrlRichEdit_SetCharColor($g_hConsoleEdit, 0xFFB347)
            _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 10, "Consolas")

        Case $CONSOLE_ERROR
            _GUICtrlRichEdit_SetCharColor($g_hConsoleEdit, 0xFF6B6B)
            _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 10, "Consolas")

        Case $CONSOLE_SUCCESS
            _GUICtrlRichEdit_SetCharColor($g_hConsoleEdit, 0x66D9EF)
            _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 10, "Consolas")

        Case $CONSOLE_HEADER
            _GUICtrlRichEdit_SetCharColor($g_hConsoleEdit, 0xFFFFFF)
            _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 11, "Consolas")

        Case Else
            _GUICtrlRichEdit_SetCharColor($g_hConsoleEdit, 0xEAEAEA)
            _GUICtrlRichEdit_SetFont($g_hConsoleEdit, 10, "Consolas")
    EndSwitch
EndFunc

Func _ClearConsole()
    _GUICtrlRichEdit_SetText($g_hConsoleEdit, "")
    _ConsoleWrite("Console output cleared.", $CONSOLE_SUCCESS)
EndFunc

Func _ShowAbout()
    Local $hAbout = GUICreate("About - Windows First Aid Toolkit", 520, 275, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU), -1, $g_hGUI)
    GUISetBkColor(0xF7FAFC)

    Local $lblApp = GUICtrlCreateLabel("Windows First Aid Toolkit", 20, 18, 320, 24)
    GUICtrlSetFont($lblApp, 12, 700, 0, "Segoe UI")

    Local $lblDesc = GUICtrlCreateLabel("A practical one-click first-step repair toolkit for Windows technicians and power users.", 20, 50, 470, 36)
    GUICtrlSetColor($lblDesc, 0x334155)

    GUICtrlCreateLabel("Author: Giorgos Xanthopoulos", 20, 98, 260, 18)
    GUICtrlCreateLabel("Icon credits: Hristos Kalaitzis", 20, 122, 260, 18)
    GUICtrlCreateLabel("Open-source utility", 20, 146, 160, 18)

    Local $lblRepo = GUICtrlCreateLabel("GitHub repository: https://github.com/Gexos/Windows-First-Aid-Toolkit", 20, 182, 470, 20)
    GUICtrlSetColor($lblRepo, 0x0A58CA)
    GUICtrlSetCursor($lblRepo, 0)

    Local $lblBlog = GUICtrlCreateLabel("Blog: https://gexos.org", 20, 208, 240, 20)
    GUICtrlSetColor($lblBlog, 0x0A58CA)
    GUICtrlSetCursor($lblBlog, 0)

    Local $btnClose = GUICtrlCreateButton("Close", 400, 228, 90, 28)

    GUISetState(@SW_SHOW, $hAbout)

    While 1
        Local $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE, $btnClose
                ExitLoop
            Case $lblRepo
                ShellExecute("https://github.com/Gexos/Windows-First-Aid-Toolkit")
            Case $lblBlog
                ShellExecute("https://gexos.org")
        EndSwitch
    WEnd

    GUIDelete($hAbout)
EndFunc

Func _InitToolTips()
    $g_hToolTip = _GUIToolTip_Create($g_hGUI)
    _GUIToolTip_SetMaxTipWidth($g_hToolTip, 340)
    _GUIToolTip_SetDelayTime($g_hToolTip, $TTDT_INITIAL, 300)
    _GUIToolTip_SetDelayTime($g_hToolTip, $TTDT_AUTOPOP, 12000)
    _GUIToolTip_SetTipBkColor($g_hToolTip, 0xFFFFF0)
    _GUIToolTip_SetTipTextColor($g_hToolTip, 0x202020)

    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Creates a Windows restore point before repairs. Useful as a safety rollback if System Protection is enabled.", GUICtrlGetHandle($chkRestorePoint))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Clears the DNS resolver cache. Useful after DNS changes or name resolution issues.", GUICtrlGetHandle($chkFlushDNS))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Resets the Windows Winsock catalog. Useful for broken network stack problems.", GUICtrlGetHandle($chkResetWinsock))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Releases the current IP lease and requests a fresh one from DHCP.", GUICtrlGetHandle($chkRenewIP))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Deletes files from the current user's temp folder.", GUICtrlGetHandle($chkUserTemp))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Deletes files from the Windows temp folder. Locked files may be skipped.", GUICtrlGetHandle($chkSystemTemp))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Clears common Edge, Chrome, and Firefox cache locations. Best used when browsers are closed.", GUICtrlGetHandle($chkBrowserCache))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Runs System File Checker to scan and repair protected Windows system files.", GUICtrlGetHandle($chkSFC))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Runs DISM RestoreHealth to repair the Windows component store.", GUICtrlGetHandle($chkDISM))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Restarts the Windows Update service.", GUICtrlGetHandle($chkWU))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Restarts the Background Intelligent Transfer Service.", GUICtrlGetHandle($chkBITS))
    _GUIToolTip_AddTool($g_hToolTip, $g_hGUI, "Restarts the Print Spooler service.", GUICtrlGetHandle($chkSpooler))
EndFunc