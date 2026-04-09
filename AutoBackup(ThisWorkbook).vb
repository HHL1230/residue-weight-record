' ==========================================================
' 功能：1. 開啟時啟動背景同步 2. 關閉時停止同步並自動備份
' 位置：請務必將此程式碼貼在「ThisWorkbook (本活頁簿)」的程式碼視窗中
' ==========================================================

Private Sub Workbook_Open()
    ' 檔案開啟時，啟動背景同步定時器 (呼叫一般模組中的程序)
    Call StartSyncTimer
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim backupPath As String
    Dim backupFileName As String
    Dim currentName As String
    Dim ext As String
    Dim baseName As String
    Dim dotPos As Integer
    
    ' --- 設定區 ---
    ' 備份資料夾路徑 (請確保最後有反斜線 \)
    backupPath = "\\10.213.74.205\InstrRawData\FCM\backup\"
    ' --------------
    
    ' 停止背景同步定時器，避免檔案關閉後仍在背景執行導致報錯
    Call StopSyncTimer
    
    ' 1. 關閉前先儲存目前的活頁簿，確保備份出去的是最新狀態
    ThisWorkbook.Save
    
    ' 2. 解析原檔名，準備加上時間戳記
    currentName = ThisWorkbook.Name
    dotPos = InStrRev(currentName, ".")
    
    If dotPos > 0 Then
        baseName = Left(currentName, dotPos - 1)
        ext = Mid(currentName, dotPos) ' 包含點，例如 .xlsm 或 .xlsx
    Else
        baseName = currentName
        ext = ".xlsm"
    End If
    
    ' 組成備份檔名，格式例如：坩堝秤重紀錄_20260408_173022.xlsm
    backupFileName = baseName & "_" & Format(Now, "yyyyMMdd_HHmmss") & ext
    
    ' 3. 執行網路備份
    On Error GoTo ErrorHandler ' 開啟錯誤捕捉，避免網路不通導致卡死無法關閉 Excel
    
    ' 使用 SaveCopyAs 可以在不改變當前檔案開啟路徑的情況下，存一份拷貝到目標路徑
    ThisWorkbook.SaveCopyAs backupPath & backupFileName
    
    Application.StatusBar = False ' 清除狀態列
    Exit Sub
    
ErrorHandler:
    ' 當網路路徑無法連線、權限不足或路徑不存在時，觸發警告但仍允許關閉
    MsgBox "自動備份至網路磁碟機失敗！" & vbCrLf & vbCrLf & _
           "目標路徑：" & backupPath & vbCrLf & _
           "系統錯誤訊息：" & Err.Description & vbCrLf & vbCrLf & _
           "請確認您的網路連線 (10.213.74.205) 是否正常。" & vbCrLf & _
           "注意：原檔案已於本機存檔，按下確定後將繼續關閉 Excel。", _
           vbExclamation + vbOKOnly, "自動備份防護通知"
           
    ' 清除錯誤，確保能順利關閉
    Err.Clear
End Sub



