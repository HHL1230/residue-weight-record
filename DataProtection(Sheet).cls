' ==========================================================
' 功能：共用活頁簿專用版 - 「VBA 虛擬鎖」與自動存檔
' 特色：完全移除 Native Protect/Unprotect，徹底解決 1004 錯誤
'       利用 SelectionChange 與 Undo 模擬強大的防覆寫鎖定
' ==========================================================

Private prevSelectedCell As Range
Private prevHasData As Boolean

' --- 事件 A：選取儲格時 (記憶舊資料狀態，取代原生鎖定檢查) ---
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim watchRange As Range
    Set watchRange = Range("C3:E34, M3:O34")
    
    If Target.Cells.Count = 1 Then
        If Not Application.Intersect(Target, watchRange) Is Nothing Then
            Set prevSelectedCell = Target
            ' 紀錄該格在被修改前「是否已經有資料」
            prevHasData = (Not IsEmpty(Target.Value) And Target.Value <> "")
        Else
            Set prevSelectedCell = Nothing
            prevHasData = False
        End If
    Else
        Set prevSelectedCell = Nothing
        prevHasData = False
    End If
End Sub

' --- 事件 B：當內容改變時 (執行覆蓋確認與存檔) ---
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watchRange As Range
    Dim isIntersect As Range
    Dim cell As Range
    Dim isOverwrite As Boolean
    
    Set watchRange = Range("C3:E34, M3:O34")
    Set isIntersect = Application.Intersect(Target, watchRange)
    
    If isIntersect Is Nothing Then Exit Sub
        
    ' 【步驟 1：判斷是否為覆寫舊資料】
    isOverwrite = False
    
    If Not prevSelectedCell Is Nothing Then
        ' 情境 1 (手動鍵盤輸入): 剛才選取的格子跟現在修改的格子是同一個
        If Target.Cells.Count = 1 And Target.Address = prevSelectedCell.Address Then
            isOverwrite = prevHasData
        Else
            ' 防呆：如果位址不符，使用 Undo 檢查過去狀態
            isOverwrite = CheckIfOverwriteByUndo(isIntersect)
        End If
    Else
        ' 情境 2 (COM/天平直接寫入 或 多格貼上): 因為沒經過選取動作，直接檢查
        isOverwrite = CheckIfOverwriteByUndo(isIntersect)
    End If
    
    ' 【步驟 2：執行覆寫確認與撤銷 (虛擬鎖核心)】
    If isOverwrite Then
        ' 第一關確認
        If MsgBox("確認修改數據 ?", vbYesNo + vbQuestion, "數據修改確認 (1/2)") = vbNo Then
            GoTo CancelEdit
        End If
        
        ' 第二關確認 (防止誤觸加嚴)
        If MsgBox("現有數據是否已備份，並再次確認要修改 ?", vbYesNo + vbExclamation, "最終覆寫確認 (2/2)") = vbNo Then
            GoTo CancelEdit
        End If
        
        ' 只有兩關都過才繼續執行存檔
    End If
    
    ' 【步驟 3：合法輸入，執行存檔與同步標記】
    Application.EnableEvents = False
    ' 【修正】：移除 ScreenUpdating = False，避免同一台電腦切換帳戶時導致畫面崩潰
    On Error GoTo CleanUp
    
    ' 執行存檔 (將新數據寫入共用活頁簿)
    ThisWorkbook.Save
    Application.StatusBar = isIntersect.Address(False, False) & " 數據已更新並完成存檔 (" & Format(Now, "HH:mm:ss") & ")"
    
    ' 【重要防衝突】：更新背景同步模組的基準時間
    ' 讓 background_sync_force 知道這是自己存的檔，不要立刻又去抓一次
    On Error Resume Next
    LastCheckTime = FileDateTime(ThisWorkbook.FullName)
    On Error GoTo CleanUp

CleanUp:
    Application.EnableEvents = True
    ' 【修正】：移除 ScreenUpdating = True
    Exit Sub

CancelEdit:
    ' 使用者選「否」：瞬間撤銷，恢復舊資料
    On Error Resume Next
    Application.EnableEvents = False
    Application.Undo
    Application.StatusBar = "已當作誤觸處理，維持原數據。"
    Application.EnableEvents = True
    Exit Sub
End Sub

' --- 輔助函數：透過 Undo 探查修改前的狀態 (專門對付 COM 寫入) ---
Private Function CheckIfOverwriteByUndo(Target As Range) As Boolean
    Dim hadData As Boolean
    Dim cell As Range
    hadData = False
    
    On Error GoTo UndoFail
    Application.EnableEvents = False
    
    ' 1. 退回上一步 (恢復成修改前的狀態)
    Application.Undo
    
    ' 2. 檢查原本是否有資料
    For Each cell In Target
        If Not IsEmpty(cell.Value) And cell.Value <> "" Then
            hadData = True
            Exit For
        End If
    Next cell
    
    ' 3. 再次執行 Undo (取消剛剛的退回，恢復成含有新資料的狀態，以便後續存檔)
    Application.Undo
    
UndoFail:
    Application.EnableEvents = True
    CheckIfOverwriteByUndo = hadData
End Function