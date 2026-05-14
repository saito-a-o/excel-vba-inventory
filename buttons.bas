Attribute VB_Name = "buttons"
Option Explicit

Sub btn_edit_Click()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("発行者情報を編集しますか？", vbYesNo)
    If result = vbYes Then
        Dim frm As editForm
        Set frm = New editForm
        frm.Show
    Else
        MsgBox "処理を中止しました"
        Exit Sub
    End If
    
End Sub

Sub btn_pst_Click()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("入力内容を転記しますか？", vbYesNo)
    If result = vbYes Then
        Call OpenTransfer
    Else
        MsgBox "処理を中止しました"
        Exit Sub
    End If

End Sub

Sub btn_pdf_Click()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("PDFファイルに出力しますか？", vbYesNo)
    If result = vbYes Then
        Application.EnableEvents = False
        Dim frm As frmPDFSelect
        Set frm = New frmPDFSelect
        frm.Show
        Application.EnableEvents = True
    Else
        MsgBox "処理を中止しました"
        Exit Sub
    End If

End Sub

Sub btn_vld_Click()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("データのチェックをしますか？", vbYesNo)
    If result = vbYes Then
        Call procedure.ProcessCheck_Main
    Else
        MsgBox "処理を中止しました"
        Exit Sub
    End If

End Sub

Sub btn_ClientSelect_Click()
    Dim frm As frmClientSelect
    Set frm = New frmClientSelect
    frm.Show
End Sub

Sub btn_prodMaster_Click()
    Dim frm As frmProductSelect
    Set frm = New frmProductSelect
    frm.Show
End Sub

Sub btn_prodSelect_Click()
    Dim frm As frmItemSelect
    Set frm = New frmItemSelect
    frm.Show
End Sub

Sub btn_sync_Click()
    Call SyncStock
End Sub

Sub btn_report_Click()
    Dim frm As frmReport
    Set frm = New frmReport
    frm.Show
End Sub
