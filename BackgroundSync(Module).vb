' ==========================================================
' 功能：背景定時「強制」同步 (解決同一台電腦多帳戶切換的閃退 Bug)
' 位置：請將此程式碼放在「一般模組 (Module)」中
' ==========================================================

Public RunTime As Double

' 啟動定時器 (由 ThisWorkbook 呼叫)
Sub StartSyncTimer()
    ScheduleNextCheck
End Sub

' 排程下一次檢查 (每 10 秒)
Sub ScheduleNextCheck()
    RunTime = Now + TimeValue("00:00:10")
    Application.OnTime RunTime, "CheckForUpdates"
End Sub

' 停止定時器 (由 ThisWorkbook 呼叫)
Sub StopSyncTimer()
    On Error Resume Next
    Application.OnTime RunTime, "CheckForUpdates", , False
End Sub

' 強制執行同步
Sub CheckForUpdates()
    On Error GoTo ErrHandler
    
    ' 【重要防護 1】：如果目前這個 Excel 視窗處於最小化 (或者被切換到背景)，直接跳過本次存檔，避免衝突
    If Application.WindowState = xlMinimized Then GoTo ErrHandler
    
    ' 確保檔案在共用模式下才執行
    If ThisWorkbook.MultiUserEditing Then
        
        ' 【重要防護 2】：拿掉 ScreenUpdating = False，避免切換使用者時畫面崩潰
        Application.EnableEvents = False
        
        ' 強制存檔以拉取遠端最新數據
        ThisWorkbook.Save
        
        ' 恢復事件
        Application.EnableEvents = True
        
        ' 在狀態列給予微小提示
        Application.StatusBar = "背景同步完成 (" & Format(Now, "HH:mm:ss") & ")"
    End If

ErrHandler:
    ' 確保不論成功或失敗，都會啟動下一次 10 秒的倒數
    ScheduleNextCheck
End Sub
