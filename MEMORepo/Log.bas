Attribute VB_Name = "Log"
Option Explicit

'******************************************************************************
' 処理概要：LOG出力ユーティリティ
' 利用方法：
' 　①プロジェクト内で最初に１回、LogOpen関数をコールする
' 　　（引数が空の場合、実行フォルダに作成）
' 　②必要に応じてLogDebug、LogInfo、LogWarn、LogErrorをコールする
' 　③プロジェクト内で最後に１回、LogClose関数をコールする
' 補足情報：
' 　本ユーティリティを利用する場合はプロジェクトの参照設定で
' 　[Microsoft Scripting Runtime]のライブラリを追加すること
'******************************************************************************

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public Declare Sub GetLocalTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME)

Dim oFSO As FileSystemObject
Dim oLog As TextStream

Public Sub LogOpen(Optional ByVal psFilePath As String = "")
    Dim sFilePath As String
    sFilePath = psFilePath
    
    If oFSO Is Nothing Then
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        If sFilePath = "" Then
            sFilePath = App.Path & "\" & App.EXEName & ".log"
        End If
        On Error Resume Next
        Err.Clear
        Set oLog = oFSO.OpenTextFile(sFilePath, ForAppending, True)
        If Err.Number <> 0 Then
            sFilePath = App.Path & "\" & App.EXEName & Format(Now, "yyyymmddhhmmss") & ".log"
            Set oLog = oFSO.OpenTextFile(sFilePath, ForAppending, True)
            Call LogWarn("LogOpen", "ファイルがオープンできなかったため、別ファイルにログを出力します。")
        End If
        Call oLog.WriteLine("********** Logging Start **********")
        
    End If
    
End Sub

Public Sub LogDebug(ByVal psModuleName As String, ByVal psLog As String)
    
    Call writeLog(psLog, "Debug", psModuleName)

End Sub

Public Sub LogInfo(ByVal psModuleName As String, ByVal psLog As String)

    Call writeLog(psLog, "Info", psModuleName)
    
End Sub

Public Sub LogWarn(ByVal psModuleName As String, ByVal psLog As String)

    Call writeLog(psLog, "Warn", psModuleName)
    
End Sub

Public Sub LogError(ByVal psModuleName As String, ByVal psLog As String)

    Call writeLog(psLog, "Error", psModuleName)
    
End Sub

Private Sub writeLog(ByVal psLog As String, ByVal psLevel As String, ByVal psModuleName As String)

    Dim sDate As String
    sDate = Format(Now, "yyyy/mm/dd") & " " & getNowTime
    
    If oLog Is Nothing Then Call LogOpen
    Call oLog.WriteLine(sDate & " [" & psLevel & "][" & psModuleName & "]" & psLog)
    
End Sub

Private Function getNowTime() As String

    Dim t As SYSTEMTIME
    
    On Error GoTo ERR_RTN
    
    GetLocalTime t

    getNowTime = Format$(t.wHour, "00") & ":" & _
                 Format$(t.wMinute, "00") & ":" & _
                 Format$(t.wSecond, "00") & "." & _
                 Format$(t.wMilliseconds, "000")
    Exit Function
    
ERR_RTN:
    getNowTime = Format$(Now, "hh:mm:ss")
    
End Function

Public Sub LogClose()

    Call oLog.WriteLine("********** Logging Finish **********")
    oLog.Close
    Set oLog = Nothing
    Set oFSO = Nothing
    
End Sub
