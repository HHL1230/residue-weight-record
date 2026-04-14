' ==========================================================
' 功能：背景定時「強制」同步 (解決同一台電腦多帳戶切換的閃退 Bug)
' 新增：智慧防干擾，操作其他活頁簿時自動暫停同步
' 位置：請將此程式碼放在「一般模組 (Module)」中
' ==========================================================

Public RunTime As Double
Public LastCheckTime As Date

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

' 主邏輯：檢查檔案更新時間
Sub CheckForUpdates()
    Dim currentFileTime As Date
    
    ' 1. 【防干擾核心】：如果使用者正在操作其他活頁簿，跳過這次同步
    ' 避免在處理其他報表時，被這個背景連線動作干擾輸入
    If Not Application.ActiveWorkbook Is ThisWorkbook Then
        GoTo Reschedule
    End If
    
    ' 2. 獲取遠端檔案 (共用活頁簿) 的最後修改時間
    On Error Resume Next
    currentFileTime = FileDateTime(ThisWorkbook.FullName)
    
    ' 3. 如果檔案時間大於上次檢查時間，代表有新數據
    ' 使用強制方法連線：這裡採用 Save 觸發 Excel 的背景合併機制
    If currentFileTime > LastCheckTime Then
        ' 暫時關閉事件與更新，減少閃爍
        Application.EnableEvents = False
        
        ' 【重要】：這裡使用 Save 但不顯示彈窗。
        ' 在「共用活頁簿」模式下，Save 會自動抓取別人的更新並合併進來。
        ThisWorkbook.Save
        
        ' 更新記錄點
        LastCheckTime = currentFileTime
        Application.StatusBar = "背景同步完成 (" & Format(Now, "HH:mm:ss") & ")"
        Application.EnableEvents = True
    End If

Reschedule:
    ' 繼續循環
    ScheduleNextCheck
End Sub