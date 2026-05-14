Attribute VB_Name = "procedure"

'税率
Public Const TAX_RATE1 As Double = 0.1
Public Const TAX_RATE2 As Double = 0.08

Dim myErrs(1 To 8, 1 To 2) As String
Dim outErrs(1 To 6, 1 To 2) As String
Dim dtlErrs(1 To 4, 0 To 2) As String

'取引先データ貼り付け形式
Private Const EXP As Integer = 1 '郵便番号 & 住所1 & vbCrLf & 住所2
Private Const ORD As Integer = 0 '郵便番号 & vbCrLf & 住所1 & vbCrLf & 住所2

Dim srcdata_detail As Variant
Dim src_sheet As Worksheet

Public Sub ProcessWithMerge()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("入力")
    '--- 保護を解除 ---
    Dim isProtected As Boolean
    isProtected = ws.ProtectContents
    If isProtected Then ws.Unprotect
    
    ' 結合セル情報の定義（アドレス, IMEモード）
    ' IME_MODE: 3=ひらがな, 8=半角英数字
    Dim mergeInfo(14, 1) As String
    mergeInfo(0, 0) = "$B$14:$O$14":  mergeInfo(0, 1) = "hiragana"
    mergeInfo(1, 0) = "$B$16:$O$16":  mergeInfo(1, 1) = "hiragana"
    mergeInfo(2, 0) = "$B$17:$O$17":  mergeInfo(2, 1) = "hiragana"
    mergeInfo(3, 0) = "$B$18:$O$18":  mergeInfo(3, 1) = "hankaku"
    mergeInfo(4, 0) = "$B$19:$O$19":  mergeInfo(4, 1) = "hiragana"
    mergeInfo(5, 0) = "$A$23:$K$23":  mergeInfo(5, 1) = "hiragana"
    mergeInfo(6, 0) = "$A$24:$K$24":  mergeInfo(6, 1) = "hiragana"
    mergeInfo(7, 0) = "$A$25:$K$25":  mergeInfo(7, 1) = "hiragana"
    mergeInfo(8, 0) = "$A$26:$K$26":  mergeInfo(8, 1) = "hiragana"
    mergeInfo(9, 0) = "$A$27:$K$27":  mergeInfo(9, 1) = "hiragana"
    mergeInfo(10, 0) = "$A$28:$K$28": mergeInfo(10, 1) = "hiragana"
    mergeInfo(11, 0) = "$A$29:$K$29": mergeInfo(11, 1) = "hiragana"
    mergeInfo(12, 0) = "$A$30:$K$30": mergeInfo(12, 1) = "hiragana"
    mergeInfo(13, 0) = "$A$31:$K$31": mergeInfo(13, 1) = "hiragana"
    mergeInfo(14, 0) = "$A$32:$K$32": mergeInfo(14, 1) = "hiragana"
    ' 先頭セルの値を退避
    Dim mergeValues(14) As Variant
    Dim i As Integer
    For i = 0 To 14
        mergeValues(i) = ws.Range(mergeInfo(i, 0)).Cells(1, 1).Value
    Next i

    '--- Step1: 結合を解除 ---
    For i = 0 To 14
        ws.Range(mergeInfo(i, 0)).UnMerge
    Next i

    '--- Step2: IMEモードを設定 ---
    For i = 0 To 14
        Dim targetRange As Range
        Set targetRange = ws.Range(mergeInfo(i, 0))

        If mergeInfo(i, 1) = "hiragana" Then
            targetRange.Validation.Delete
            targetRange.Validation.Add Type:=xlValidateInputOnly
            targetRange.Validation.IMEMode = xlIMEModeHiragana
        Else
            targetRange.Validation.Delete
            targetRange.Validation.Add Type:=xlValidateInputOnly
            targetRange.Validation.IMEMode = xlIMEModeDisable
        End If
    Next i

    '--- Step3: 結合を元に戻す ---
    For i = 0 To 14
        With ws.Range(mergeInfo(i, 0))
            .Merge
            .Cells(1, 1).Value = mergeValues(i)
        End With
    Next i
    
    With ws.Range("C15:J15,L23:M32")
        .Validation.Delete
        .Validation.Add Type:=xlValidateInputOnly
        .Validation.IMEMode = xlIMEModeDisable
    End With

    '--- 保護を元に戻す ---
'    If isProtected Then ws.Protect
    ws.Protect

End Sub

Public Sub ProcessCheck_Main()
    Set src_sheet = ThisWorkbook.Worksheets("入力")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    If src_sheet.ProtectContents = True Then src_sheet.Unprotect
    If ProcessCheck_Fnc <> False Then GoTo step_end
    Application.ScreenUpdating = True
    MsgBox "チェックＯＫ", buttons:=vbInformation, Title:="検証の結果"
step_end:
    src_sheet.Protect
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
Private Function ProcessCheck_Fnc() As Boolean
    ProcessCheck_Fnc = False
    Dim rng As Range, rng_area As Range
    Set rng_area = src_sheet.Range("B4:O19, A23:N32")
    If src_sheet.ProtectContents Then src_sheet.Unprotect
    For Each rng In rng_area
        If rng.Interior.Color = 13421823 Then rng.Interior.ColorIndex = 0
    Next rng
    '発行者情報のチェック
    If Check_myInfo <> False Then
       ProcessCheck_Fnc = True
    End If
    '取引先情報のチェック
    If Check_custInfo <> False Then ProcessCheck_Fnc = True
    '明細のチェック
    If Check_detailInfo <> False Then ProcessCheck_Fnc = True
      
End Function
Private Function Check_myInfo() As Boolean
    Check_myInfo = False
    Dim err_id As Long: err_id = 1
    Erase myErrs
    '会社名
    If Check_Valid_text(SRC_RANGE_MYCOMP, err_id, myErrs()) Then
        myErrs(err_id, 1) = " 会社名"
        err_id = err_id + 1
    End If
    '郵便番号
    If Check_Valid_postNum(SRC_RANGE_MYPOST3, err_id, myErrs()) Then
        myErrs(err_id, 1) = " 郵便番号"
        err_id = err_id + 1
    End If
    '住所①
    If Check_Valid_text(SRC_RANGE_MYADDR1, err_id, myErrs()) Then
        myErrs(err_id, 1) = " 住所①"
        err_id = err_id + 1
    End If

    '住所②は空欄チェックなし。※①におさまるケース


    '電話番号
    If Check_Valid_number(SRC_RANGE_MYTELNUMB, err_id, myErrs()) Then
        myErrs(err_id, 1) = " 電話番号"
        err_id = err_id + 1
    End If

    'FAX番号
    If Check_Valid_Faxnumber(SRC_RANGE_MYFAXNUMB, err_id, myErrs()) Then
        myErrs(err_id, 1) = " FAX番号"
        err_id = err_id + 1
    End If

    'インボイス登録番号
    If Check_Valid_Tnumber(SRC_RANGE_MYTNUMB, err_id, myErrs()) Then
        myErrs(err_id, 1) = " インボイス登録番号"
        err_id = err_id + 1
    End If

    '担当者
    If Check_Valid_text(SRC_RANGE_MYPERSON, err_id, myErrs()) Then
        myErrs(err_id, 1) = " 担当者"
        err_id = err_id + 1
    End If

    'メッセージ出力
    If myErrs(1, 1) <> "" Then
        Check_myInfo = True
        Call outMsg(myErrs(), "【発行者情報】", err_id)
    End If
    
End Function
Private Function Check_custInfo() As Boolean
    Check_custInfo = False
    Dim err_id As Long: err_id = 1
    Erase outErrs
    '取引先名
    If Check_Valid_text(SRC_RANGE_OUTCOMP, err_id, outErrs()) Then
        outErrs(err_id, 1) = " 会社名"
        err_id = err_id + 1
    End If
    '郵便番号
    If Check_Valid_postNum(SRC_RANGE_OUTPOST3, err_id, outErrs()) Then
        outErrs(err_id, 1) = " 郵便番号"
        err_id = err_id + 1
    End If
    '住所①
    If Check_Valid_text(SRC_RANGE_OUTADDR1, err_id, outErrs()) Then
        outErrs(err_id, 1) = " 住所①"
        err_id = err_id + 1
    End If

    '住所②は空欄チェックなし。※①におさまるケース


    '注文番号
    If Check_Valid_text(SRC_RANGE_OUTORDRNUM, err_id, outErrs()) Then
        outErrs(err_id, 1) = " 注文番号"
        err_id = err_id + 1
    End If

    '日付
    If Check_Valid_Date(SRC_RANGE_OUTPRCDATE, err_id, outErrs()) Then
        outErrs(err_id, 1) = " 日付"
        err_id = err_id + 1
    End If

    'メッセージ出力
    If outErrs(1, 1) <> "" Then
        Check_custInfo = True
        Call outMsg(outErrs(), "【取引先情報】", err_id)
    End If
    
End Function

Private Sub outMsg(errs() As String, Lbl As String, err_id)
    If errs(1, 1) <> "" Then
        Dim i As Long
        Dim msg As String
        msg = Lbl & "に不備があります" & vbCrLf & vbCrLf
        For i = 1 To err_id - 1
            If i <> (err_id - 1) Then
                msg = msg & errs(i, 1) & ": " & errs(i, 2) & vbCrLf
            Else
                msg = msg & errs(i, 1) & ": " & errs(i, 2)
            End If
        Next i
        Application.ScreenUpdating = True
        MsgBox msg, vbExclamation
        Application.ScreenUpdating = False
    End If

End Sub
Private Function Check_detailInfo() As Boolean
    Check_detailInfo = False
    Dim err_id As Long: err_id = 1
    Dim j As Long
    Dim startcell As Range
    Dim errflg As Integer
    Erase dtlErrs
    '品目 ※1行目のデータ有無
    If Check_Valid_text(SRC_RANGE_DETAILITEM1, err_id, dtlErrs()) Then
        dtlErrs(err_id, 0) = " 品目"
        dtlErrs(err_id, 1) = src_sheet.Range(SRC_RANGE_DETAILITEM1).row & "行目"
        err_id = err_id + 1
    End If
 
    '数量
    Set startcell = src_sheet.Range(SRC_RANGE_DETAILNUM1)
    errflg = 0
    For j = 0 To 9
        If Len(Trim(startcell.Offset(j, -1).MergeArea.Cells(1, 1).Value)) > 0 Then '品目データが入っていたら数量をチェック
            If Check_Valid_ItemNum(startcell.Offset(j, 0).Address, err_id, dtlErrs()) Then
                If errflg = 0 Then
                    dtlErrs(err_id, 0) = " 数量"
                    dtlErrs(err_id, 1) = startcell.Offset(j, 0).row & "行目"
                    errflg = 1
                End If
            End If
        End If
    Next j
    If errflg > 0 Then err_id = err_id + 1

    '単価
    Set startcell = src_sheet.Range(SRC_RANGE_DETAILPRCE1)
    errflg = 0
    For j = 0 To 9
        If Len(Trim(startcell.Offset(j, -2).MergeArea.Cells(1, 1).Value)) > 0 Then '品目データが入っていたら単価をチェック
            If Check_Valid_ItemNum(startcell.Offset(j, 0).Address, err_id, dtlErrs()) Then
                If errflg = 0 Then
                    dtlErrs(err_id, 0) = " 単価"
                    dtlErrs(err_id, 1) = startcell.Offset(j, 0).row & "行目"
                    errflg = 1
                End If
            End If
        End If
    Next j
    If errflg > 0 Then err_id = err_id + 1

    '税率
    Set startcell = src_sheet.Range(SRC_RANGE_DETAILRATE1)
    errflg = 0
    For j = 0 To 9
        If Len(Trim(startcell.Offset(j, -3).MergeArea.Cells(1, 1).Value)) > 0 Then '品目データが入っていたら税率をチェック
            If Check_Valid_Rate(startcell.Offset(j, 0).Address, err_id, dtlErrs()) Then
                If errflg = 0 Then
                    dtlErrs(err_id, 0) = " 税率"
                    dtlErrs(err_id, 1) = startcell.Offset(j, 0).row & "行目"
                    errflg = 1
                End If
            End If
        End If
    Next j
    If errflg > 0 Then err_id = err_id + 1

    'メッセージ出力
    If dtlErrs(1, 0) <> "" Then
        Check_detailInfo = True
        Dim i As Long
        Dim msg As String
        msg = "【明細】に不備があります" & vbCrLf & vbCrLf
        For i = 1 To err_id - 1
            If i <> (err_id - 1) Then
                msg = msg & dtlErrs(i, 0) & " " & dtlErrs(i, 1) & ": " & dtlErrs(i, 2) & vbCrLf
            Else
                msg = msg & dtlErrs(i, 0) & " " & dtlErrs(i, 1) & ": " & dtlErrs(i, 2)
            End If
        Next i
        Application.ScreenUpdating = True
        MsgBox msg, vbExclamation
        Application.ScreenUpdating = False
    End If
    
End Function
Private Function Check_Valid_text(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_text = False
    If Len(Trim(src_sheet.Range(rng).Value)) < 1 Then
        Check_Valid_text = True
        arr(err_id, 2) = "未入力です"
        src_sheet.Range(rng).Interior.Color = 13421823
    End If
End Function
Private Function Check_Valid_postNum(rng As String, err_id As Long, arr() As String) As Boolean
    Dim i As Long
    Check_Valid_postNum = False
    For i = 0 To 7
        If i <> 3 Then
            If Len(Trim(src_sheet.Range(rng).Offset(0, i).Value)) < 1 Then
                Check_Valid_postNum = True
                arr(err_id, 2) = "空欄があります"
                src_sheet.Range(rng).Offset(0, i).Interior.Color = 13421823
            ElseIf IsNumeric(src_sheet.Range(rng).Offset(0, i).Value) = False Then
                Check_Valid_postNum = True
                arr(err_id, 2) = "数字以外の文字が含まれています"
                src_sheet.Range(rng).Offset(0, i).Interior.Color = 13421823
            End If
        End If
    Next i

End Function
Private Function Check_Valid_number(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_number = False
    If Len(Trim(src_sheet.Range(rng).Value)) < 1 Then
        Check_Valid_number = True
        arr(err_id, 2) = "未入力です"
        src_sheet.Range(rng).Interior.Color = 13421823
    Else
        src_sheet.Range(rng).Value = Trim(StrConv(src_sheet.Range(rng).Value, vbNarrow))
        If InStr(src_sheet.Range(rng).Value, "-") = 0 Then
            Check_Valid_number = True
            arr(err_id, 2) = "ハイフン（-）がありません"
            src_sheet.Range(rng).Interior.Color = 13421823
            Exit Function
        End If
        If IsNumeric(Replace(src_sheet.Range(rng).Value, "-", "")) = False Then
            Check_Valid_number = True
            arr(err_id, 2) = "数字以外の文字が含まれています"
            src_sheet.Range(rng).Interior.Color = 13421823
        End If
    End If
End Function
Private Function Check_Valid_Faxnumber(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_Faxnumber = False
    'FAXが無い場合チェック不要
    If Len(Trim(src_sheet.Range(rng).Value)) > 0 Then
        src_sheet.Range(rng).Value = Trim(StrConv(src_sheet.Range(rng).Value, vbNarrow))

        If InStr(src_sheet.Range(rng).Value, "-") = 0 Then
            Check_Valid_Faxnumber = True
            arr(err_id, 2) = "ハイフン（-）がありません"
            src_sheet.Range(rng).Interior.Color = 13421823
            Exit Function
        End If
        If IsNumeric(Replace(src_sheet.Range(rng).Value, "-", "")) = False Then
            Check_Valid_Faxnumber = True
            arr(err_id, 2) = "数字以外の文字が含まれています"
            src_sheet.Range(rng).Interior.Color = 13421823
        End If
    End If
End Function
Private Function Check_Valid_Tnumber(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_Tnumber = False
    If Len(Trim(src_sheet.Range(rng).Value)) < 1 Then
        Check_Valid_Tnumber = True
        arr(err_id, 2) = "未入力です"
        src_sheet.Range(rng).Interior.Color = 13421823
    Else
        src_sheet.Range(rng).Value = Trim(StrConv(src_sheet.Range(rng).Value, vbNarrow))
        If Left(src_sheet.Range(rng).Value, 1) <> "T" Then
            Check_Valid_Tnumber = True
            arr(err_id, 2) = "1桁目が「Ｔ」ではありません"
            src_sheet.Range(rng).Interior.Color = 13421823
            Exit Function
        End If
        If IsNumeric(Replace(src_sheet.Range(rng).Value, "T", "")) = False Then
            Check_Valid_Tnumber = True
            arr(err_id, 2) = "「Ｔ」と数字以外の文字が含まれています"
            src_sheet.Range(rng).Interior.Color = 13421823
            Exit Function
        End If
        If Len(src_sheet.Range(rng).Value) <> 14 Then
            Check_Valid_Tnumber = True
            arr(err_id, 2) = "桁数が不正です"
            src_sheet.Range(rng).Interior.Color = 13421823
        End If

    End If
End Function

Private Function Check_Valid_Date(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_Date = False
    If Len(Trim(src_sheet.Range(rng).Value)) < 1 Then
        Check_Valid_Date = True
        arr(err_id, 2) = "未入力です"
        src_sheet.Range(rng).Interior.Color = 13421823
        Exit Function
    End If
    If IsDate(src_sheet.Range(rng).Value) = False Then
        Check_Valid_Date = True
        arr(err_id, 2) = "日付ではありません"
        src_sheet.Range(rng).Interior.Color = 13421823
    End If

End Function

Private Function Check_Valid_ItemNum(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_ItemNum = False
    If Len(Trim(src_sheet.Range(rng).Value)) < 1 Then
        Check_Valid_ItemNum = True
        arr(err_id, 2) = "未入力です"
        src_sheet.Range(rng).Interior.Color = 13421823
    Else
        src_sheet.Range(rng).Value = Trim(StrConv(src_sheet.Range(rng).Value, vbNarrow))
        If IsNumeric(src_sheet.Range(rng).Value) = False Then
            Check_Valid_ItemNum = True
            arr(err_id, 2) = "数字以外の文字が含まれています"
            src_sheet.Range(rng).Interior.Color = 13421823
        End If

    End If
End Function

Private Function Check_Valid_Rate(rng As String, err_id As Long, arr() As String) As Boolean
    Check_Valid_Rate = False
    If Len(Trim(src_sheet.Range(rng).Value)) < 1 Then
        Check_Valid_Rate = True
        arr(err_id, 2) = "未入力です"
        src_sheet.Range(rng).Interior.Color = 13421823
    Else
        src_sheet.Range(rng).Value = Trim(StrConv(src_sheet.Range(rng).Value, vbNarrow))
        If (src_sheet.Range(rng).Value <> TAX_RATE1) And _
        (src_sheet.Range(rng).Value <> TAX_RATE2) Then
            Check_Valid_Rate = True
            arr(err_id, 2) = "不正な値です"
            src_sheet.Range(rng).Interior.Color = 13421823
        End If

    End If
End Function

'── 転記フォーム起動 btn_pst_Click()から呼び出し──
Sub OpenTransfer()
    Dim frm As frmTransfer
    Set frm = New frmTransfer
    frm.Show
End Sub

'── 選択シート用の転記処理 ──
Public Sub CopyPaste_To_Single(ws As Worksheet, srcWs As Worksheet)
    Dim dataTo As Data_To
    Dim pnumstr As String: pnumstr = ""
    Dim i As Integer
    dataTo.CompTo = srcWs.Range(SRC_RANGE_OUTCOMP).Value
    For i = 0 To 7
        If i = 3 Then
            pnumstr = pnumstr & "-"
        Else
            pnumstr = pnumstr & srcWs.Range(SRC_RANGE_OUTPOST3).Offset(0, i).Value
        End If
    Next i
    dataTo.PostnumTo = pnumstr
    dataTo.Addr1To = srcWs.Range(SRC_RANGE_OUTADDR1).Value
    dataTo.Addr2To = srcWs.Range(SRC_RANGE_OUTADDR2).Value
    dataTo.OrdernumTo = srcWs.Range(SRC_RANGE_OUTORDRNUM).Value
    dataTo.DateTo = srcWs.Range(SRC_RANGE_OUTPRCDATE).Value

    Select Case ws.Name
        Case SHEET_MITSUMORI, SHEET_SEIKYUU
            Call PasteData_To(ws, dataTo, EXP)
        Case SHEET_NOUHINN, SHEET_RYOUSYUU
            Call PasteData_To(ws, dataTo, ORD)
    End Select
End Sub

Public Sub CopyPaste_From_Single(ws As Worksheet, srcWs As Worksheet, sheetName As String)
    Dim dataFrom As Data_From
    Dim pnumstr As String: pnumstr = ""
    Dim i As Integer
    dataFrom.CompFrm = srcWs.Range(SRC_RANGE_MYCOMP).Value
    For i = 0 To 7
        If i = 3 Then
            pnumstr = pnumstr & "-"
        Else
            pnumstr = pnumstr & srcWs.Range(SRC_RANGE_MYPOST3).Offset(0, i).Value
        End If
    Next i
    dataFrom.PostnumFrm = pnumstr
    dataFrom.Addr1Frm = srcWs.Range(SRC_RANGE_MYADDR1).Value
    dataFrom.Addr2Frm = srcWs.Range(SRC_RANGE_MYADDR2).Value
    dataFrom.TnumFrm = srcWs.Range(SRC_RANGE_MYTNUMB).Value
    dataFrom.TelFrm = srcWs.Range(SRC_RANGE_MYTELNUMB).Value
    dataFrom.FaxFrm = srcWs.Range(SRC_RANGE_MYFAXNUMB).Value
    dataFrom.PrsnFrm = srcWs.Range(SRC_RANGE_MYPERSON).Value

    Select Case sheetName
        Case SHEET_MITSUMORI
            Call PasteData_From(ws, dataFrom, EXP)
        Case Else
            Call PasteData_From(ws, dataFrom, ORD)
    End Select
End Sub

Public Sub CopyPaste_Detail_Single(ws As Worksheet, srcWs As Worksheet, sheetName As String)
    Dim srcdata As Variant
    ReDim srcdata(1 To 10, dtail_item To dtail_rate)
    srcdata = srcWs.Range(SRC_RANGE_DETAILITEM1).Resize(10, dtail_rate)

    Select Case sheetName
        Case SHEET_MITSUMORI
            Call PasteData_Detail(ws, srcdata, EXP)
        Case Else
            Call PasteData_Detail(ws, srcdata, ORD)
    End Select
End Sub

Private Sub PasteData_To(ws As Worksheet, srcdata As Data_To, pastetype As Integer)
    ws.Range(TGT_RANGE_OUTCOMP) = srcdata.CompTo & " 御中"
    If pastetype = EXP Then
        ws.Range(TGT_RANGE_OUTPOST) = "〒" & srcdata.PostnumTo & " " & srcdata.Addr1To
        ws.Range(TGT_RANGE_OUTADDR1) = "  " & srcdata.Addr2To
    Else
        ws.Range(TGT_RANGE_OUTPOST) = "〒" & srcdata.PostnumTo
        ws.Range(TGT_RANGE_OUTADDR1) = srcdata.Addr1To
        ws.Range(TGT_RANGE_OUTADDR2) = srcdata.Addr2To
    End If
    ws.Range(TGT_RANGE_OUTCOMP & "," & TGT_RANGE_OUTADDR1 & "," & TGT_RANGE_OUTADDR2).ShrinkToFit = True
End Sub

Private Sub PasteData_From(ws As Worksheet, srcdata As Data_From, pastetype As Integer)
    Dim tp As Integer
    If pastetype = 1 Then
        tp = 1
    Else
        tp = 0
    End If

    ws.Range(TGT_RANGE_MYCOMP).Offset(tp, 0) = srcdata.CompFrm
    ws.Range(TGT_RANGE_MYPOST).Offset(tp, 0) = "〒" & srcdata.PostnumFrm
    ws.Range(TGT_RANGE_MYADDR).Offset(tp, 0) = srcdata.Addr1Frm & " " & srcdata.Addr2Frm
    ws.Range(TGT_RANGE_MYTNUM).Offset(tp, 0) = "登録番号：" & srcdata.TnumFrm
    ws.Range(TGT_RANGE_MYEL).Offset(tp, 0) = "TEL：" & srcdata.TelFrm
    ws.Range(TGT_RANGE_MYFAX).Offset(tp, 0) = "FAX：" & srcdata.FaxFrm
    ws.Range(TGT_RANGE_MYPSN).Offset(tp, 0) = "担当：" & srcdata.PrsnFrm
    ws.Range(TGT_RANGE_MYCOMP).Offset(tp, 0).ShrinkToFit = True
    ws.Range(TGT_RANGE_MYADDR).Offset(tp, 0).ShrinkToFit = True
End Sub

Private Sub PasteData_Detail(ws As Worksheet, srcdata As Variant, pastetype As Integer)
    Dim tp As Integer, i As Integer
    If pastetype = 1 Then
        tp = EXP
    Else
        tp = ORD
    End If
    For i = LBound(srcdata) To UBound(srcdata)
        ws.Range(TGT_RANGE_ITEM).Offset(tp + i - 1, 0) = srcdata(i, dtail_item)
        ws.Range(TGT_RANGE_LGT).Offset(tp + i - 1, 0) = IIf(srcdata(i, dtail_rate) = TAX_RATE2, "※", "")
        ws.Range(TGT_RANGE_NUM).Offset(tp + i - 1, 0) = srcdata(i, dtail_num)
        ws.Range(TGT_RANGE_NLBL).Offset(tp + i - 1, 0) = ""
        ws.Range(TGT_RANGE_PRICE).Offset(tp + i - 1, 0) = srcdata(i, dtail_price)
        ws.Range(TGT_RANGE_RATE).Offset(tp + i - 1, 0) = srcdata(i, dtail_rate)
    Next i
End Sub

Public Sub CopyPaste_Order_Single(ews As Worksheet, srcWs As Worksheet, sheetName As String, frm As Object)

    '── 発行日 ──
    ews.Range(TGT_RANGE_COMM_DOCD).Value = Date

    '── 振込先をDBから取得 ──
    Dim cn As Object
    Set cn = GetConnection()
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT bank_name, account_type, account_number, account_name " & _
            "FROM my_company WHERE company_id = 1", cn
    Dim bankStr As String
    If Not rs.EOF Then
        bankStr = rs("bank_name") & "　" & _
                  rs("account_type") & "　" & _
                  rs("account_number") & "　" & _
                  rs("account_name") & ""
    End If
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    '── 書類ごとに転記 ──
    Select Case sheetName
        Case SHEET_MITSUMORI
            ews.Range(TGT_RANGE_MIT_SUBJ).Value = frm.Controls("txtSubject").Text
            ews.Range(TGT_RANGE_MIT_TERM).Value = frm.Controls("txtPayTerms").Text
            ews.Range(TGT_RANGE_MIT_DLDT).Value = CDate(frm.Controls("txtDelivery").Text)
            ews.Range(TGT_RANGE_MIT_EXPD).Value = CDate(frm.Controls("txtExpiry").Text)
            With ews.Range(TGT_RANGE_MIT_DLDT, TGT_RANGE_MIT_EXPD)
                .NumberFormat = "yyyy/mm/dd"
            End With
        Case SHEET_NOUHINN
            ews.Range(TGT_RANGE_NOU_SUBJ).Value = frm.Controls("txtSubject").Text

        Case SHEET_SEIKYUU
            ews.Range(TGT_RANGE_SEI_SUBJ).Value = frm.Controls("txtSubject").Text
            ews.Range(TGT_RANGE_SEI_DUED).Value = CDate(frm.Controls("txtPayDue").Text)
            With ews.Range(TGT_RANGE_SEI_DUED)
                .NumberFormat = "yyyy/mm/dd"
            End With
            ews.Range(TGT_RANGE_SEI_BANK).Value = bankStr

        Case SHEET_RYOUSYUU
            ews.Range(TGT_RANGE_RYO_SUBJ).Value = frm.Controls("txtSubject").Text
    End Select

End Sub
