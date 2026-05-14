VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBtnHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public act As String
Public frm As Object

Private Sub btn_Click()
    Dim ws As Worksheet
    Dim c As Integer
    Set ws = ThisWorkbook.Sheets("入力")
    Application.EnableEvents = False
    Select Case act

        ' ==============================
        '  editForm 用
        ' ==============================
        Case "CANCEL"
            If MsgBox("編集を中断しますか？" & vbCrLf & _
                      "変更内容は保存されません。", _
                      vbYesNo + vbQuestion, "確認") = vbYes Then
                Unload frm
            End If

        Case "OK"
            If MsgBox("変更を確定しますか？", _
                      vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

            ws.Unprotect
            ' 郵便番号
            Dim z3 As String: z3 = frm.Controls("txtZip3").Text
            Dim z4 As String: z4 = frm.Controls("txtZip4").Text
            
            For c = 1 To 3
                ws.Cells(5, c + 2).Value = IIf(c <= Len(z3), Mid(z3, c, 1), "")
            Next c
            For c = 1 To 4
                ws.Cells(5, c + 6).Value = IIf(c <= Len(z4), Mid(z4, c, 1), "")
            Next c

            ' 各フィールド
            Dim fields As Variant
            fields = Array( _
                Array("txtComp", 4), Array("txtAdr1", 6), _
                Array("txtAdr2", 7), Array("txtTel", 8), _
                Array("txtFax", 9), Array("txtInv", 10), _
                Array("txtPrs", 11))
            Dim i As Integer
            For i = 0 To UBound(fields)
                ws.Cells(fields(i)(1), 2).Value = _
                    frm.Controls(fields(i)(0)).Text
            Next i
            ws.Protect
            '── DBへの保存処理 ──
            Dim cn As Object
            Dim rs As Object
            Dim sql As String
            Set cn = GetConnection()
            Set rs = CreateObject("ADODB.Recordset")
            
            '── レコードの存在確認 ──
            rs.Open "SELECT COUNT(*) FROM my_company;", cn
            Dim recCount As Long
            recCount = rs(0)
            
            If recCount = 0 Then '── 初回：INSERT ──
                sql = "INSERT INTO my_company " & _
                  "(company_name, postal_code, address1, address2, " & _
                  "tel, fax, invoice_number, staff_name, " & _
                  "bank_name, account_type, account_number, account_name) " & _
                  "VALUES ('" & _
                  frm.Controls("txtComp").Text & "','" & _
                  frm.Controls("txtZip3").Text & "-" & frm.Controls("txtZip4").Text & "','" & _
                  frm.Controls("txtAdr1").Text & "','" & _
                  frm.Controls("txtAdr2").Text & "','" & _
                  frm.Controls("txtTel").Text & "','" & _
                  frm.Controls("txtFax").Text & "','" & _
                  frm.Controls("txtInv").Text & "','" & _
                  frm.Controls("txtPrs").Text & "','" & _
                  frm.Controls("txtBank").Text & "','" & _
                  frm.Controls("txtAccType").Text & "','" & _
                  frm.Controls("txtAccNum").Text & "','" & _
                  frm.Controls("txtAccName").Text & "')"
            Else '── 2回目以降：UPDATE ──
                sql = "UPDATE my_company SET " & _
                  "company_name = '" & frm.Controls("txtComp").Text & "'," & _
                  "postal_code = '" & frm.Controls("txtZip3").Text & "-" & frm.Controls("txtZip4").Text & "'," & _
                  "address1 = '" & frm.Controls("txtAdr1").Text & "'," & _
                  "address2 = '" & frm.Controls("txtAdr2").Text & "'," & _
                  "tel = '" & frm.Controls("txtTel").Text & "'," & _
                  "fax = '" & frm.Controls("txtFax").Text & "'," & _
                  "invoice_number = '" & frm.Controls("txtInv").Text & "'," & _
                  "staff_name = '" & frm.Controls("txtPrs").Text & "'," & _
                  "bank_name = '" & frm.Controls("txtBank").Text & "'," & _
                  "account_type = '" & frm.Controls("txtAccType").Text & "'," & _
                  "account_number = '" & frm.Controls("txtAccNum").Text & "'," & _
                  "account_name = '" & frm.Controls("txtAccName").Text & "' " & _
                  "WHERE company_id = 1"
            End If
            
            cn.Execute sql
            cn.Close
            Set cn = Nothing
            
            MsgBox "変更が完了しました。", vbInformation, "完了"
            Unload frm

        ' ==============================
        '  frmPDFSelect 用
        ' ==============================
        Case "BROWSE"
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "出力先フォルダを選択してください"
                .AllowMultiSelect = False
                If .Show = -1 Then
                    frm.Controls("txtFolder").Text = .SelectedItems(1)
                End If
            End With

        Case "CANCEL_PDF"
            Unload frm

        Case "OUTPUT"
            ' フォルダ未選択チェック
            Dim folderPath As String
            folderPath = frm.Controls("txtFolder").Text
            If folderPath = "" Then
                MsgBox "出力先フォルダを選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If

            ' シート未選択チェック
            Dim ctrl As MSForms.Control
            Dim hasChecked As Boolean
            For Each ctrl In frm.Controls
                If Left(ctrl.Name, 4) = "chk_" Then
                    If ctrl.Value Then hasChecked = True: Exit For
                End If
            Next ctrl
            If Not hasChecked Then
                MsgBox "出力するシートを1つ以上選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If

            ' PDF出力
            Dim errList As String
            For Each ctrl In frm.Controls
                If Left(ctrl.Name, 4) = "chk_" Then
                    If ctrl.Value Then
                        Dim pdfPath As String
                        pdfPath = folderPath & "\" & ctrl.Caption & ".pdf"
                        On Error Resume Next
'                        Call ThisWorkbook.SaveAsPDF_WithTempClear(ctrl.Caption, pdfPath)
                        ThisWorkbook.Sheets(ctrl.Caption).ExportAsFixedFormat _
                            Type:=xlTypePDF, _
                            fileName:=pdfPath, _
                            Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, _
                            IgnorePrintAreas:=False, _
                            OpenAfterPublish:=False
                        If err.Number <> 0 Then
                            errList = errList & "・" & ctrl.Caption & vbCrLf
                            err.Clear
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next ctrl

            If errList = "" Then
                MsgBox "PDF出力が完了しました。" & vbCrLf & vbCrLf & _
                       "保存先：" & folderPath, vbInformation, "完了"
            Else
                MsgBox "一部のシートで出力に失敗しました：" & vbCrLf & _
                       errList, vbExclamation, "エラー"
            End If
            Unload frm
        ' ==============================
        '  frmClientSelect 用
        ' ==============================
        Case "SEARCH_CLIENT"
            Call frm.LoadClientList(frm, frm.Controls("txtSearch").Text)

        Case "NEW_CLIENT"
            Dim frmNew As frmClientEdit
            Set frmNew = New frmClientEdit
            frmNew.m_mode = "NEW"
            Set frmNew.m_cn = frm.m_cn
            Set frmNew.m_parentFrm = frm
            frmNew.Caption = "取引先の新規登録"
            frmNew.Controls("lblTitle").Caption = "取引先の新規登録"
            frmNew.Show

        Case "EDIT_CLIENT"
            Dim lst As MSForms.ListBox
            Set lst = frm.Controls("lstClients")
            If lst.ListIndex = -1 Then
                MsgBox "取引先を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            Dim frmEdit As frmClientEdit
            Set frmEdit = New frmClientEdit
            frmEdit.m_mode = "EDIT"
            Set frmEdit.m_cn = frm.m_cn
            frmEdit.m_clientId = CLng(lst.List(lst.ListIndex, 0))
            Set frmEdit.m_parentFrm = frm
            frmEdit.Caption = "取引先の編集"
            frmEdit.Controls("lblTitle").Caption = "取引先の編集"
            frmEdit.LoadData
            frmEdit.Show

        Case "DELETE_CLIENT"
            Dim lstDel As MSForms.ListBox
            Set lstDel = frm.Controls("lstClients")
            If lstDel.ListIndex = -1 Then
                MsgBox "取引先を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If MsgBox("選択した取引先を削除しますか？", _
                      vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
            Dim delId As Long
            delId = CLng(lstDel.List(lstDel.ListIndex, 0))
            frm.m_cn.Execute "UPDATE clients SET is_deleted = TRUE WHERE client_id = " & delId
            MsgBox "削除しました。", vbInformation, "完了"
            Call frm.LoadClientList(frm, "")

        Case "SELECT_CLIENT"
            Dim lstSel As MSForms.ListBox
            Set lstSel = frm.Controls("lstClients")
            If lstSel.ListIndex = -1 Then
                MsgBox "取引先を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            '── 選択した取引先の詳細をDBから取得 ──
            Dim selId As Long
            selId = CLng(lstSel.List(lstSel.ListIndex, 0))
            Dim rsSel As Object
            Set rsSel = CreateObject("ADODB.Recordset")
            rsSel.Open "SELECT * FROM clients WHERE client_id = " & selId, frm.m_cn
            If Not rsSel.EOF Then
                ws.Unprotect
                '── 郵便番号の分割セット ──
                Dim zip As String
                zip = rsSel("postal_code") & ""
                Dim zipArr As Variant
                zipArr = Split(zip, "-")
                If UBound(zipArr) >= 1 Then
                    For c = 1 To 3
                        ws.Cells(15, c + 2).Value = IIf(c <= Len(zipArr(0)), Mid(zipArr(0), c, 1), "")
                    Next c
                    For c = 1 To 4
                        ws.Cells(15, c + 6).Value = IIf(c <= Len(zipArr(1)), Mid(zipArr(1), c, 1), "")
                    Next c
                End If
                ws.Cells(14, 2).Value = rsSel("client_name") & ""
                ws.Cells(16, 2).Value = rsSel("address1") & ""
                ws.Cells(17, 2).Value = rsSel("address2") & ""
                ws.Protect
            End If
            rsSel.Close
            Set rsSel = Nothing
            Unload frm

        Case "CLOSE_CLIENT"
            Unload frm
        ' ==============================
        '  frmClientEdit 用
        ' ==============================
        Case "CANCEL_CLIENT_EDIT"
            Unload frm

        Case "OK_CLIENT_EDIT"
            '── 入力チェック ──
            If frm.Controls("txtName").Text = "" Then
                MsgBox "取引先名を入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If

            '── SQL組み立て ──
            Dim sqlCE As String
            Dim zipCE As String
            zipCE = frm.Controls("txtZip3").Text & "-" & frm.Controls("txtZip4").Text

            If frm.m_mode = "NEW" Then
                sqlCE = "INSERT INTO clients " & _
                        "(client_name, postal_code, address1, address2, " & _
                        "tel, email, contact_name) VALUES ('" & _
                        frm.Controls("txtName").Text & "','" & _
                        zipCE & "','" & _
                        frm.Controls("txtAdr1").Text & "','" & _
                        frm.Controls("txtAdr2").Text & "','" & _
                        frm.Controls("txtTel").Text & "','" & _
                        frm.Controls("txtEmail").Text & "','" & _
                        frm.Controls("txtContact").Text & "')"
            Else
                sqlCE = "UPDATE clients SET " & _
                        "client_name = '" & frm.Controls("txtName").Text & "'," & _
                        "postal_code = '" & zipCE & "'," & _
                        "address1 = '" & frm.Controls("txtAdr1").Text & "'," & _
                        "address2 = '" & frm.Controls("txtAdr2").Text & "'," & _
                        "tel = '" & frm.Controls("txtTel").Text & "'," & _
                        "email = '" & frm.Controls("txtEmail").Text & "'," & _
                        "contact_name = '" & frm.Controls("txtContact").Text & "' " & _
                        "WHERE client_id = " & frm.m_clientId
            End If

            frm.m_cn.Execute sqlCE

            '── 親フォームの一覧を更新 ──
            Call frm.m_parentFrm.LoadClientList( _
                frm.m_parentFrm, "")

            MsgBox "保存しました。", vbInformation, "完了"
            Unload frm
            
        ' ==============================
        '  frmProductSelect 用
        ' ==============================
        Case "SEARCH_PRODUCT"
            Call frm.LoadProductList(frm, frm.Controls("txtSearch").Text)

        Case "NEW_PRODUCT"
            Dim frmProdNew As frmProductEdit
            Set frmProdNew = New frmProductEdit
            frmProdNew.m_mode = "NEW"
            Set frmProdNew.m_cn = frm.m_cn
            Set frmProdNew.m_parentFrm = frm
            frmProdNew.Caption = "商品の新規登録"
            frmProdNew.Controls("lblTitle").Caption = "商品の新規登録"
            frmProdNew.Show

        Case "EDIT_PRODUCT"
            Dim lstP As MSForms.ListBox
            Set lstP = frm.Controls("lstProducts")
            If lstP.ListIndex = -1 Then
                MsgBox "商品を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            Dim frmProdEdit As frmProductEdit
            Set frmProdEdit = New frmProductEdit
            frmProdEdit.m_mode = "EDIT"
            Set frmProdEdit.m_cn = frm.m_cn
            frmProdEdit.m_prodId = CLng(lstP.List(lstP.ListIndex, 0))
            Set frmProdEdit.m_parentFrm = frm
            frmProdEdit.Caption = "商品の編集"
            frmProdEdit.Controls("lblTitle").Caption = "商品の編集"
            frmProdEdit.LoadData
            frmProdEdit.Show

        Case "DELETE_PRODUCT"
            Dim lstPDel As MSForms.ListBox
            Set lstPDel = frm.Controls("lstProducts")
            If lstPDel.ListIndex = -1 Then
                MsgBox "商品を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If MsgBox("選択した商品を削除しますか？" & vbCrLf & _
                      "※この操作は元に戻せません。", _
                      vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
            Dim delProdId As Long
            delProdId = CLng(lstPDel.List(lstPDel.ListIndex, 0))
            frm.m_cn.Execute "DELETE FROM products WHERE prod_id = " & delProdId
            MsgBox "削除しました。", vbInformation, "完了"
            Call frm.LoadProductList(frm, "")

        Case "CLOSE_PRODUCT"
            Unload frm

        ' ==============================
        '  frmProductEdit 用
        ' ==============================
        Case "CANCEL_PRODUCT_EDIT"
            Unload frm

        Case "OK_PRODUCT_EDIT"
            '── 入力チェック ──
            If frm.Controls("txtName").Text = "" Then
                MsgBox "商品名を入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If frm.Controls("txtPrice").Text = "" Then
                MsgBox "単価を入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If Not IsNumeric(frm.Controls("txtPrice").Text) Then
                MsgBox "単価は数値で入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If Not IsNumeric(frm.Controls("txtStock").Text) Then
                MsgBox "在庫数は数値で入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If Not IsNumeric(frm.Controls("txtReorder").Text) Then
                MsgBox "発注点は数値で入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If

            '── SQL組み立て ──
            Dim sqlProd As String
            Dim taxRateVal As String
            taxRateVal = frm.Controls("cmbTax").Text

            If frm.m_mode = "NEW" Then
                sqlProd = "INSERT INTO products " & _
                          "(product_name, product_category, product_subcategory, " & _
                          "unit, price, tax_rate, stock, reorder_point) VALUES ('" & _
                          frm.Controls("txtName").Text & "','" & _
                          frm.Controls("txtCat").Text & "','" & _
                          frm.Controls("txtSubCat").Text & "','" & _
                          frm.Controls("txtUnit").Text & "'," & _
                          CLng(frm.Controls("txtPrice").Text) & ",'" & _
                          taxRateVal & "'," & _
                          CLng(frm.Controls("txtStock").Text) & "," & _
                          CLng(frm.Controls("txtReorder").Text) & ")"
            Else
                sqlProd = "UPDATE products SET " & _
                          "product_name = '" & frm.Controls("txtName").Text & "'," & _
                          "product_category = '" & frm.Controls("txtCat").Text & "'," & _
                          "product_subcategory = '" & frm.Controls("txtSubCat").Text & "'," & _
                          "unit = '" & frm.Controls("txtUnit").Text & "'," & _
                          "price = " & CLng(frm.Controls("txtPrice").Text) & "," & _
                          "tax_rate = '" & taxRateVal & "'," & _
                          "stock = " & CLng(frm.Controls("txtStock").Text) & "," & _
                          "reorder_point = " & CLng(frm.Controls("txtReorder").Text) & " " & _
                          "WHERE prod_id = " & frm.m_prodId
            End If

            frm.m_cn.Execute sqlProd

            '── 親フォームの一覧を更新 ──
            Call frm.m_parentFrm.LoadProductList( _
                frm.m_parentFrm, "")

            MsgBox "保存しました。", vbInformation, "完了"
            Unload frm
        ' ==============================
        '  frmTransfer 用
        ' ==============================
        Case "CANCEL_TRANS"
            Unload frm

        Case "EXEC_TRANS"
            '── 書類未選択チェック ──
            If Not frm.Controls("chkMit").Value And _
               Not frm.Controls("chkNou").Value And _
               Not frm.Controls("chkSei").Value And _
               Not frm.Controls("chkRyo").Value Then
                MsgBox "書類を1つ以上選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If

            If MsgBox("転記を実行しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

            Application.ScreenUpdating = False
            Application.EnableEvents = False

            Dim ews As Worksheet
'            Set srcWs = ThisWorkbook.Sheets("入力")

            '── シート保護解除 ──
            For Each ews In ThisWorkbook.Worksheets
                If ews.ProtectContents Then ews.Unprotect
            Next ews

            ws.Calculate

            '── 選択シートのみ転記 ──
            Dim sheetMap(3) As String
            sheetMap(0) = IIf(frm.Controls("chkMit").Value, SHEET_MITSUMORI, "")
            sheetMap(1) = IIf(frm.Controls("chkNou").Value, SHEET_NOUHINN, "")
            sheetMap(2) = IIf(frm.Controls("chkSei").Value, SHEET_SEIKYUU, "")
            sheetMap(3) = IIf(frm.Controls("chkRyo").Value, SHEET_RYOUSYUU, "")

            Dim si As Integer
            For si = 0 To 3
                If sheetMap(si) <> "" Then
                    Set ews = ThisWorkbook.Sheets(sheetMap(si))
                    Call CopyPaste_To_Single(ews, ws)
                    Call CopyPaste_From_Single(ews, ws, sheetMap(si))
                    Call CopyPaste_Detail_Single(ews, ws, sheetMap(si))
                    Call CopyPaste_Order_Single(ews, ws, sheetMap(si), frm)
                End If
            Next si
            
            '── 全シート保護 ──
            For Each ews In ThisWorkbook.Worksheets
                ews.Calculate
                ews.Protect
            Next ews
            ws.Activate

            '── DB保存 ──
            Call SaveOrderToDB(frm, frm.m_cn)

            Application.ScreenUpdating = True
            Application.EnableEvents = True

            MsgBox "転記が完了しました。" & vbCrLf & "書類番号(J2セル)を入力してください", vbInformation, "完了"
            Unload frm
        ' ==============================
        '  frmItemSelect 用
        ' ==============================
        Case "ADD_ITEM"
            Dim frmSearch As frmItemSearch
            Set frmSearch = New frmItemSearch
            Set frmSearch.m_cn = frm.m_cn
            Set frmSearch.m_parentFrm = frm
            frmSearch.Caption = "商品検索"
            frmSearch.LoadData
            frmSearch.Show

        Case "DEL_ITEM"
            
            If frm.m_selectedRow = 0 Then
                MsgBox "削除する行を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            Dim delIdx As Integer
            delIdx = frm.m_selectedRow
'            itmlstDel.RemoveItem itmlstDel.ListIndex
            '── TextBox/ComboBoxを1行ずつ上にシフト ──
            Dim csi As Integer
            For csi = delIdx To frm.m_rowCount - 1
                frm.Controls("txtName_" & csi).Text = frm.Controls("txtName_" & csi + 1).Text
                frm.Controls("txtPrice_" & csi).Text = frm.Controls("txtPrice_" & csi + 1).Text
                frm.Controls("txtQty_" & csi).Text = frm.Controls("txtQty_" & csi + 1).Text
                frm.Controls("cmbTax_" & csi).ListIndex = frm.Controls("cmbTax_" & csi + 1).ListIndex
            Next csi
            '── 最終行をクリア ──
            frm.Controls("txtName_" & frm.m_rowCount).Text = ""
            frm.Controls("txtPrice_" & frm.m_rowCount).Text = ""
            frm.Controls("txtQty_" & frm.m_rowCount).Text = ""
            frm.Controls("cmbTax_" & frm.m_rowCount).ListIndex = 0
            '── 選択状態をリセット ──
            frm.m_rowCount = frm.m_rowCount - 1
            frm.SelectRow 0
        Case "CANCEL_ITEM"
            Unload frm

        Case "OK_ITEM"
            '── 入力チェック ──
            If frm.m_rowCount = 0 Then
                MsgBox "商品を1つ以上追加してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            Dim oi As Integer
            For oi = 1 To frm.m_rowCount
                If frm.Controls("txtQty_" & oi).Text = "" Or _
                   Not IsNumeric(frm.Controls("txtQty_" & oi).Text) Then
                    MsgBox oi & "行目の数量を正しく入力してください。", vbExclamation, "エラー"
                    Exit Sub
                End If
                If frm.Controls("txtPrice_" & oi).Text = "" Or _
                   Not IsNumeric(frm.Controls("txtPrice_" & oi).Text) Then
                    MsgBox oi & "行目の単価を正しく入力してください。", vbExclamation, "エラー"
                    Exit Sub
                End If
            Next oi

            '── 入力シートに反映 ──
'            Dim wsOI As Worksheet
'            Set wsOI = ThisWorkbook.Sheets("入力")
            ws.Unprotect

            '── 既存明細をクリア ──
            Dim ci As Integer
            For ci = 0 To 9
                ws.Range(SRC_RANGE_DETAILITEM1).Offset(ci, 0).MergeArea.Cells(1, 1).Value = ""
                ws.Range(SRC_RANGE_DETAILNUM1).Offset(ci, 0).Value = ""
                ws.Range(SRC_RANGE_DETAILPRCE1).Offset(ci, 0).Value = ""
                ws.Range(SRC_RANGE_DETAILRATE1).Offset(ci, 0).Value = ""
            Next ci

            '── 新しい明細をセット ──
            For oi = 1 To frm.m_rowCount
                ws.Range(SRC_RANGE_DETAILITEM1).Offset(oi - 1, 0).MergeArea.Cells(1, 1).Value = frm.Controls("txtName_" & oi).Text
                ws.Range(SRC_RANGE_DETAILNUM1).Offset(oi - 1, 0).Value = CLng(frm.Controls("txtQty_" & oi).Text)
                ws.Range(SRC_RANGE_DETAILPRCE1).Offset(oi - 1, 0).Value = CLng(frm.Controls("txtPrice_" & oi).Text)
                ws.Range(SRC_RANGE_DETAILRATE1).Offset(oi - 1, 0).Value = frm.Controls("cmbTax_" & oi).Text
            Next oi

            ws.Protect
            MsgBox "明細を反映しました。", vbInformation, "完了"
            Unload frm

        ' ==============================
        '  frmItemSearch 用
        ' ==============================
        Case "SEARCH_ITEM"
            frm.LoadItemList frm.Controls("txtKeyword").Text

        Case "CANCEL_ITEM_SEARCH"
            Unload frm

        Case "SELECT_ITEM"
            Dim lstSrc As MSForms.ListBox
            Set lstSrc = frm.Controls("lstSearch")
            '── 選択件数チェック ──
            Dim selCount As Integer
            Dim li As Integer
            For li = 0 To lstSrc.ListCount - 1
                If lstSrc.Selected(li) Then selCount = selCount + 1
            Next li
            If selCount = 0 Then
                MsgBox "商品を選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            
            Dim parentFrm As frmItemSelect
            Set parentFrm = frm.m_parentFrm
            
            If parentFrm.m_rowCount + selCount > 10 Then
                MsgBox "明細は10行までです。" & vbCrLf & _
                       "現在" & parentFrm.m_rowCount & "行、追加予定" & selCount & "行", _
                       vbExclamation, "エラー"
                Exit Sub
            End If
            '── 親フォームに追加 ──
            For li = 0 To lstSrc.ListCount - 1
                If lstSrc.Selected(li) Then
                    parentFrm.AddItemRow _
                        lstSrc.List(li, 1), _
                        CLng(lstSrc.List(li, 2)), _
                        lstSrc.List(li, 3)
                End If
            Next li
            Unload frm
        ' ==============================
        '  frmReport 用
        ' ==============================
        Case "BROWSE_REPORT"
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "出力先フォルダを選択してください"
                .AllowMultiSelect = False
                If .Show = -1 Then
                    frm.Controls("txtFolder").Text = .SelectedItems(1)
                End If
            End With
        
        Case "CANCEL_REPORT"
            Unload frm
        
        Case "OUTPUT_REPORT"
            If frm.Controls("txtFolder").Text = "" Then
                MsgBox "出力先フォルダを選択してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If Not IsDate(frm.Controls("txtStart").Text) Then
                MsgBox "開始日を正しく入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            If Not IsDate(frm.Controls("txtEnd").Text) Then
                MsgBox "終了日を正しく入力してください。", vbExclamation, "エラー"
                Exit Sub
            End If
            Call OutputReport(frm)
            Unload frm

    End Select
    Application.EnableEvents = True

End Sub

