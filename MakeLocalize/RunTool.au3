#RequireAdmin

#include <WinAPIProc.au3>

#include <WinAPIProc.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <file.au3>


Global $TITLE_VER = "ShortcutCopyTool - v0.1"
Global $Author    = "YanBin"

; $CmdLine[0] ;参数的数量
; $CmdLine[1] ;第一个参数 (脚本名称后面)
; $CmdLine[2] ;第二个参数等等

$PATH =$CmdLine[1]
$FUN  =$CmdLine[2]

$sLogPath = @ScriptDir & "\my.log"

Local $hWnd

If $FUN == 0 Then
    _RunSCCopyTool()
Else
    $RUNF  =$CmdLine[3]
    _FileWriteLog($sLogPath, "开始:     " & $RUNF)
    _RunECTTool()
    _FileWriteLog($sLogPath, "结束:     " & $RUNF)
EndIF


Func _RunSCCopyTool()
    ; If Not WinExists("Shortcut Copy Tool") Then
    ;     MsgBox(0, $TITLE_VER, "No [SCCopy_v0.4.1.exe] is running.")
    ;     Exit
    ; EndIf
    If Not WinExists("Shortcut Copy Tool") Then
        $WS = 1
        While $WS
            $hWnd = _WinGetHandleByPnmAndCls("SCCopy_v0.4.1.exe")
            If $hWnd Then
               $WS = 0
            EndIf
            WinActivate($hWnd)
        WEnd
    Else
        WinActivate("Shortcut Copy Tool")
    EndIf

    ControlSetText("Shortcut Copy Tool", "", "Edit1", $PATH)
    ControlSend("Shortcut Copy Tool", "", "Button4", "{SPACE}")

    $WS = 1
    While $WS
        If WinExists("SCCopy_v0.4.1") Then
            $WS = 0
        EndIf
    WEnd

    $WinText = WinGetText("SCCopy_v0.4.1")

    $WS = 1
    While $WS
        If WinExists("Shortcut Copy Tool") Then
            $WS = 0
        EndIf
    WEnd

    If ControlSend("SCCopy_v0.4.1", "", "Button1", "{SPACE}") Then
        WinClose("Shortcut Copy Tool")
    EndIF
EndFunc

Func _RunECTTool()
    If Not WinExists("ECT") Then
        $WS = 1
        While $WS
            $hWnd = _WinGetHandleByPnmAndCls("ECT.exe")
            If $hWnd Then
               $WS = 0
            EndIf
            WinActivate($hWnd)
        WEnd
    Else
        WinActivate("ECT")
    EndIf

    Send("!fo")
    ; WinWaitActive("[CLASS:#32770]")

    ; $WS = 1
    ; While $WS
    ;     If WinExists("[CLASS:#32770]") Then
    ;         $WS = 0
    ;         WinWaitActive("[CLASS:#32770]")
    ;     EndIf
    ; WEnd

    $WS = 1
    While $WS
        If WinExists("ファイルを開く") Then
            WinWaitActive("ファイルを開く")
            _FileWriteLog($sLogPath, "ECT 选择:     -开始")
            $WS = 0
         EndIf
    WEnd

    $WS = 1
    ; _FileWriteLog($sLogPath, "ECT 路径:" & ControlGetText("ファイルを開く", "", "Edit1"))
    While $WS
        If ControlGetText("ファイルを開く", "", "Edit1") == "" Then
            ControlSetText("ファイルを開く", "", "Edit1", $PATH)
        Else
            _FileWriteLog($sLogPath, "ECT 路径:" & ControlGetText("ファイルを開く", "", "Edit1"))
            $WS = 0
        EndIf
    WEnd

    Send("!o")
    _FileWriteLog($sLogPath, "ECT 选择:     -结束")

    Local $hWnd = _WinGetHandleByPnmAndCls("ECT.exe")
    If Not $hWnd Then
       MsgBox($TITLE_VER, "", "窗口没找到")
       Exit
    EndIf
    WinActivate($hWnd)

    Send("!c")
    Send("!u")

    $WS = 1
    While $WS
        If WinExists("Exec Specified Function") Then
            $WS = 0
        ElseIf WinExists("ECT", "No driver will") Then
            ControlSend("ECT", "No driver will", "Button1", "{SPACE}")
            Send("!c")
            Send("!u")
        Else
            If WinExists("ファイルを開く") Then
                WinWaitActive("ファイルを開く")

                _FileWriteLog($sLogPath, "ECT 选择:     -开始2")

                ControlSetText("ファイルを開く", "", "Edit1", $PATH)
                _FileWriteLog($sLogPath, "ECT 路径:" & ControlGetText("ファイルを開く", "", "Edit1"))
                _FileWriteLog($sLogPath, "ECT 选择:     -结束2")
                
                Send("!o")
                Send("!c")
                Send("!u")
            EndIf
        EndIf
    WEnd

    WinActivate("Exec Specified Function")
    ControlSend("Exec Specified Function", "", "Button7", "{SPACE}")

    $WS = 1
    While $WS
        If WinExists("ECT", "Is it OK") Then
            $WS = 0
        EndIf
    WEnd
    WinActivate("ECT", "Is it OK")
    ControlSend("ECT", "Is it OK", "Button1", "{SPACE}")

    $WS = 1
    While $WS
        If WinExists("ECT", "Result") Then
            $WS = 0
        EndIf
    WEnd
    WinActivate("ECT", "Result")
    If ControlSend("ECT", "Result", "Button2", "{SPACE}") Then
        $hWnd = _WinGetHandleByPnmAndCls("ECT.exe")
        $WS = 1
        While $WS
            If WinExists($hWnd) Then
                
                WinWaitActive($hWnd)
                Send("!fx")

                $WE = 1
                While $WE
                    $hWnd = _WinGetHandleByPnmAndCls("ECT.exe")
                    If WinExists($hWnd) Then
                        If WinExists("ECT", "Current project") Then
                            Send("!y")
                        EndIf
                    Else
                        $WE = 0
                    EndIf
                WEnd
                $WS = 0
            Else
                $hWnd = _WinGetHandleByPnmAndCls("ECT.exe")
            EndIf
        WEnd
        ; Send("!fx")

        ; $WS = 1
        ; While $WS
        ;     $hWnd = _WinGetHandleByPnmAndCls("ECT.exe")
        ;     If WinExists($hWnd) Then
        ;         If WinExists("ECT", "Current project") Then
        ;             Send("!y")
        ;         EndIf
        ;     Else
        ;         $WS = 0
        ;     EndIf
        ; WEnd
    EndIf
EndFunc

; 根据pname和class获取窗口句柄，找不到则返回0
Func _WinGetHandleByPnmAndCls($pname)
   ; 根据进程名查找进程id
   Local $pid = ProcessExists($pname)
   ; MsgBox($MB_SYSTEMMODAL, "", $pname)
   ; 如果进程存在，继续
   If $pid Then
      return _WinGetHandleByPidAndCls($pid)
   Else
      Return 0
   EndIf
EndFunc

; 根据pid和class获取窗口句柄，找不到则返回0
Func _WinGetHandleByPidAndCls($pid)
   ; 这里使用枚举所有顶层窗口方法，WinList方法会返回大量隐藏窗口
   Local $winArr = _WinAPI_EnumWindowsTop()
   ; 遍历所有窗口,进程id与指定进程id比较
   For $i=1 To $winArr[0][0]
      If $pid=WinGetProcess($winArr[$i][0]) Then
         ; 一个进程会有多个窗口，所以要用class来筛选
         return $winArr[$i][0]
      EndIf
   Next
   Return 0
EndFunc