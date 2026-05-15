Attribute VB_Name = "odbc"
Option Explicit
'── DB接続 ──
Public Function GetConnection() As Object
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "DSN=inventory_dsn;UID=postgres;PWD=＜パスワード＞;"'＜パスワード＞をPostgreSQLのPasswordに置き換える
    Set GetConnection = cn
End Function

'── DB保存処理 ──
Public Sub SaveOrderToDB(frm As Object, cn As Object)

    Dim srcWs As Worksheet
    Set srcWs = ThisWorkbook.Sheets("入力")

    '── 取引先名からclient_idを取得 ──
    Dim clientName As String
    clientName = srcWs.Range(SRC_RANGE_OUTCOMP).Value
    Dim rsClient As Object
    Set rsClient = CreateObject("ADODB.Recordset")
    rsClient.Open "SELECT client_id FROM clients WHERE client_name = '" & clientName & "' AND is_deleted = FALSE LIMIT 1", cn
    If rsClient.EOF Then
        rsClient.Close
        cn.Close
        MsgBox "取引先「" & clientName & "」がDBに登録されていません。" & vbCrLf & _
               "取引先マスタに登録してから転記してください。", vbExclamation, "エラー"
        Exit Sub
    End If
    Dim clientId As Long
    clientId = CLng(rsClient("client_id"))
    rsClient.Close
    Set rsClient = Nothing

    '── 注文番号で既存ordersを検索 ──
    Dim orderNum As String
    orderNum = srcWs.Range(SRC_RANGE_OUTORDRNUM).Value
    Dim rsOrder As Object
    Set rsOrder = CreateObject("ADODB.Recordset")
    rsOrder.Open "SELECT order_id FROM orders WHERE order_number = '" & orderNum & "' LIMIT 1", cn

    '── フォームから値を取得 ──
    Dim subj     As String: subj = frm.Controls("txtSubject").Text
    Dim deliv    As String: deliv = frm.Controls("txtDelivery").Text
    Dim expiry   As String: expiry = frm.Controls("txtExpiry").Text
    Dim payTerms As String: payTerms = frm.Controls("txtPayTerms").Text
    Dim payDue   As String: payDue = frm.Controls("txtPayDue").Text

    '── NULLセーフな日付変換 ──
    Dim delivSQL   As String: delivSQL = IIf(deliv = "", "NULL", "'" & deliv & "'")
    Dim expirySQL  As String: expirySQL = IIf(expiry = "", "NULL", "'" & expiry & "'")
    Dim payDueSQL  As String: payDueSQL = IIf(payDue = "", "NULL", "'" & payDue & "'")
    Dim orderDate  As String: orderDate = Format(srcWs.Range(SRC_RANGE_OUTPRCDATE).Value, "yyyy/mm/dd")

    '── is_****フラグ ──
    Dim isQuoted    As String: isQuoted = IIf(frm.Controls("chkMit").Value, "TRUE", "FALSE")
    Dim isDelivered As String: isDelivered = IIf(frm.Controls("chkNou").Value, "TRUE", "FALSE")
    Dim isInvoiced  As String: isInvoiced = IIf(frm.Controls("chkSei").Value, "TRUE", "FALSE")
    Dim isReceipted As String: isReceipted = IIf(frm.Controls("chkRyo").Value, "TRUE", "FALSE")

    Dim orderId As Long
    Dim sqlOrder As String

    If rsOrder.EOF Then
        '── INSERT ──
        sqlOrder = "INSERT INTO orders " & _
                   "(client_id, order_number, order_date, subject, " & _
                   "delivery_date, expiry_date, payment_terms, payment_due_date, " & _
                   "is_quoted, is_delivered, is_invoiced, is_receipted) VALUES (" & _
                   clientId & ",'" & orderNum & "','" & orderDate & "','" & subj & "'," & _
                   delivSQL & "," & expirySQL & ",'" & payTerms & "'," & payDueSQL & "," & _
                   isQuoted & "," & isDelivered & "," & isInvoiced & "," & isReceipted & ")"
        cn.Execute sqlOrder

        '── 発行されたorder_idを取得 ──
        Dim rsNewId As Object
        Set rsNewId = CreateObject("ADODB.Recordset")
        rsNewId.Open "SELECT order_id FROM orders WHERE order_number = '" & orderNum & "' LIMIT 1", cn
        orderId = CLng(rsNewId("order_id"))
        rsNewId.Close
        Set rsNewId = Nothing
    Else
        '── UPDATE ──
        orderId = CLng(rsOrder("order_id"))
        sqlOrder = "UPDATE orders SET " & _
                   "client_id = " & clientId & "," & _
                   "order_date = '" & orderDate & "'," & _
                   "subject = '" & subj & "'," & _
                   "delivery_date = " & delivSQL & "," & _
                   "expiry_date = " & expirySQL & "," & _
                   "payment_terms = '" & payTerms & "'," & _
                   "payment_due_date = " & payDueSQL & "," & _
                   "is_quoted = " & isQuoted & "," & _
                   "is_delivered = " & isDelivered & "," & _
                   "is_invoiced = " & isInvoiced & "," & _
                   "is_receipted = " & isReceipted & " " & _
                   "WHERE order_id = " & orderId
        cn.Execute sqlOrder

        '── 既存明細を削除 ──
        cn.Execute "DELETE FROM order_details WHERE order_id = " & orderId
    End If
    rsOrder.Close
    Set rsOrder = Nothing

    '── order_detailsにINSERT ──
    Dim i As Integer
    Dim itemCell As Range
    Set itemCell = srcWs.Range(SRC_RANGE_DETAILITEM1)

    For i = 0 To 9
        Dim itemName As String
        itemName = Trim(itemCell.Offset(i, 0).MergeArea.Cells(1, 1).Value)
        If itemName = "" Then GoTo NextRow

        Dim qty      As Long:   qty = CLng(srcWs.Range(SRC_RANGE_DETAILNUM1).Offset(i, 0).Value)
        Dim price    As Long:   price = CLng(srcWs.Range(SRC_RANGE_DETAILPRCE1).Offset(i, 0).Value)
        Dim taxRate  As String: taxRate = srcWs.Range(SRC_RANGE_DETAILRATE1).Offset(i, 0).Value
        taxRate = Replace(Replace(taxRate, TAX_RATE1, "10%"), TAX_RATE2, "8%")
        '── prod_idを商品名で検索 ──
        Dim rsProd As Object
        Set rsProd = CreateObject("ADODB.Recordset")
        rsProd.Open "SELECT prod_id FROM products WHERE product_name = '" & itemName & "' LIMIT 1", cn
        Dim prodId As String, freeItemName As String
'        prodId = IIf(rsProd.EOF, "NULL", CLng(rsProd("prod_id")) & "")
        If rsProd.EOF Then
            prodId = "NULL"
            freeItemName = "'" & itemName & "'"
        Else
            prodId = CLng(rsProd("prod_id")) & ""
            freeItemName = "NULL"
        End If
        rsProd.Close
        Set rsProd = Nothing

        Dim sqlDtl As String
        sqlDtl = "INSERT INTO order_details " & _
                 "(order_id, prod_id, quantity, unit_price, tax_rate, free_item_name) VALUES (" & _
                 orderId & "," & prodId & "," & qty & "," & price & ",'" & taxRate & "', " & _
                 freeItemName & ")"
        cn.Execute sqlDtl
NextRow:
    Next i
    
End Sub

Public Sub SyncStock()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("入力")

    '── 注文番号を取得 ──
    Dim orderNum As String
    orderNum = ws.Range(SRC_RANGE_OUTORDRNUM).Value
    If orderNum = "" Then
        MsgBox "注文番号が入力されていません。", vbExclamation, "エラー"
        Exit Sub
    End If

    Dim cn As Object
    Set cn = GetConnection()

    '── order_idとis_stock_syncedを取得 ──
    Dim rsOrder As Object
    Set rsOrder = CreateObject("ADODB.Recordset")
    rsOrder.Open "SELECT order_id, is_stock_synced FROM orders " & _
                 "WHERE order_number = '" & orderNum & "' LIMIT 1", cn

    If rsOrder.EOF Then
        rsOrder.Close
        cn.Close
        MsgBox "注文番号「" & orderNum & "」がDBに登録されていません。" & vbCrLf & _
               "先に転記を実行してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    '── 二重減算チェック ──
    If CBool(rsOrder("is_stock_synced")) Then
        rsOrder.Close
        cn.Close
        MsgBox "この注文はすでに在庫と同期済みです。" & vbCrLf & _
               "二重減算を防ぐため処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    Dim orderId As Long
    orderId = CLng(rsOrder("order_id"))
    rsOrder.Close
    Set rsOrder = Nothing

    '── order_detailsを取得 ──
    Dim rsDetail As Object
    Set rsDetail = CreateObject("ADODB.Recordset")
    rsDetail.Open "SELECT od.prod_id, od.quantity, " & _
                  "p.product_name, p.stock, p.reorder_point " & _
                  "FROM order_details od " & _
                  "INNER JOIN products p ON od.prod_id = p.prod_id " & _
                  "WHERE od.order_id = " & orderId, cn

    If rsDetail.EOF Then
        rsDetail.Close
        cn.Close
        MsgBox "明細データが見つかりません。", vbExclamation, "エラー"
        Exit Sub
    End If

    '── 各商品をチェックしながら減算 ──
    Dim shortageMsg  As String: shortageMsg = ""   ' 不足リスト
    Dim reorderMsg   As String: reorderMsg = ""    ' 発注点リスト

    Do While Not rsDetail.EOF
        Dim prodId    As Long:   prodId = CLng(rsDetail("prod_id"))
        Dim prodName  As String: prodName = rsDetail("product_name") & ""
        Dim qty       As Long:   qty = CLng(rsDetail("quantity"))
        Dim stock     As Long:   stock = CLng(rsDetail("stock"))
        Dim reorder   As Long:   reorder = CLng(rsDetail("reorder_point"))
        Dim newStock  As Long:   newStock = stock - qty

        '── 不足チェック（減算前のstockで判定）──
        If newStock < 0 Then
            shortageMsg = shortageMsg & "　・" & prodName & _
                          "（" & Abs(newStock) & "個不足）" & vbCrLf
        '── 発注点チェック（不足の場合は除外）──
        ElseIf newStock <= reorder Then
            reorderMsg = reorderMsg & "　・" & prodName & _
                         "（残" & newStock & "個）" & vbCrLf
        End If

        '── stock減算 ──
        cn.Execute "UPDATE products SET stock = " & newStock & _
                   " WHERE prod_id = " & prodId

        rsDetail.MoveNext
    Loop

    rsDetail.Close
    Set rsDetail = Nothing

    '── is_stock_synced を TRUE に更新 ──
    cn.Execute "UPDATE orders SET is_stock_synced = TRUE " & _
               "WHERE order_id = " & orderId

    cn.Close
    Set cn = Nothing

    '── 完了メッセージ ──
    MsgBox "在庫の同期が完了しました。", vbInformation, "完了"

    '── 不足メッセージ（深刻） ──
    If shortageMsg <> "" Then
        MsgBox "【緊急】在庫が不足しています！" & vbCrLf & vbCrLf & _
               "以下の商品は在庫が足りません。早急に発注してください。" & vbCrLf & vbCrLf & _
               shortageMsg, vbCritical, "在庫不足"
    End If

    '── 発注点メッセージ（警告） ──
    If reorderMsg <> "" Then
        MsgBox "【発注点到達】在庫の補充をご検討ください" & vbCrLf & vbCrLf & _
               "以下の商品が発注点以下になりました。" & vbCrLf & vbCrLf & _
               reorderMsg, vbExclamation, "発注点到達"
    End If

End Sub

Public Sub OutputReport(frm As Object)

    Dim startDate As String: startDate = frm.Controls("txtStart").Text
    Dim endDate   As String: endDate = frm.Controls("txtEnd").Text
    Dim folderPath As String: folderPath = frm.Controls("txtFolder").Text

    '── 新しいBookを作成 ──
    Dim wb As Workbook
    Set wb = Workbooks.Add

    Dim cn As Object
    Set cn = GetConnection()

    '── 4シートを作成・データ出力 ──
    Call CreateMonthlySheet(wb, cn, startDate, endDate)
    Call CreateClientSheet(wb, cn, startDate, endDate)
    Call CreateProductSheet(wb, cn, startDate, endDate)
    Call CreateStockAlertSheet(wb, cn)

    '── デフォルトの空シートを削除 ──
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name <> "月別売上" And ws.Name <> "取引先別売上" And _
           ws.Name <> "商品別売上" And ws.Name <> "在庫アラート" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True

    cn.Close
    Set cn = Nothing

    '── Book保存 ──
    Dim fileName As String
    fileName = folderPath & "\売上レポート_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    wb.SaveAs fileName, xlOpenXMLWorkbook
    wb.Close

    MsgBox "レポートを出力しました。" & vbCrLf & vbCrLf & _
           "保存先：" & fileName, vbInformation, "完了"

End Sub

'── 月別売上シート ──
Private Sub CreateMonthlySheet(wb As Workbook, cn As Object, startDate As String, endDate As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = "月別売上"

    '── ヘッダー ──
    ws.Range("A1").Value = "月別売上レポート"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "集計期間：" & startDate & " ～ " & endDate
    ws.Range("A2").Font.Color = RGB(100, 100, 100)

    ws.Range("A4").Value = "年月"
    ws.Range("B4").Value = "売上合計"
    ws.Range("C4").Value = "注文件数"

    With ws.Range("A4:C4")
        .Font.Bold = True
        .Interior.Color = RGB(47, 84, 150)
        .Font.Color = RGB(255, 255, 255)
    End With

    '── データ取得 ──
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT TO_CHAR(o.order_date, 'YYYY/MM') AS month, " & _
            "SUM(od.quantity * od.unit_price) AS total_sales, " & _
            "COUNT(DISTINCT o.order_id) AS order_count " & _
            "FROM orders o " & _
            "INNER JOIN order_details od ON o.order_id = od.order_id " & _
            "WHERE o.order_date BETWEEN '" & startDate & "' AND '" & endDate & "' " & _
            "GROUP BY TO_CHAR(o.order_date, 'YYYY/MM') " & _
            "ORDER BY month", cn

    Dim row As Integer: row = 5
    Do While Not rs.EOF
        ws.Cells(row, 1).Value = rs("month") & ""
        ws.Cells(row, 2).Value = CLng(rs("total_sales"))
        ws.Cells(row, 3).Value = CLng(rs("order_count"))
        ws.Cells(row, 2).NumberFormat = "#,##0"
        row = row + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    '── 合計行 ──
    If row > 5 Then
        ws.Cells(row, 1).Value = "合計"
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 2).Formula = "=SUM(B5:B" & row - 1 & ")"
        ws.Cells(row, 2).Font.Bold = True
        ws.Cells(row, 2).NumberFormat = "#,##0"
        ws.Cells(row, 3).Formula = "=SUM(C5:C" & row - 1 & ")"
        ws.Cells(row, 3).Font.Bold = True
        With ws.Range("A" & row & ":C" & row)
            .Interior.Color = RGB(220, 230, 241)
        End With
    End If

    ws.Columns("A:C").AutoFit
End Sub

'── 取引先別売上シート ──
Private Sub CreateClientSheet(wb As Workbook, cn As Object, startDate As String, endDate As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = "取引先別売上"

    '── ヘッダー ──
    ws.Range("A1").Value = "取引先別売上レポート"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "集計期間：" & startDate & " ～ " & endDate
    ws.Range("A2").Font.Color = RGB(100, 100, 100)

    ws.Range("A4").Value = "取引先名"
    ws.Range("B4").Value = "売上合計"
    ws.Range("C4").Value = "注文件数"

    With ws.Range("A4:C4")
        .Font.Bold = True
        .Interior.Color = RGB(47, 84, 150)
        .Font.Color = RGB(255, 255, 255)
    End With

    '── データ取得 ──
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT c.client_name, " & _
            "SUM(od.quantity * od.unit_price) AS total_sales, " & _
            "COUNT(DISTINCT o.order_id) AS order_count " & _
            "FROM orders o " & _
            "INNER JOIN clients c ON o.client_id = c.client_id " & _
            "INNER JOIN order_details od ON o.order_id = od.order_id " & _
            "WHERE o.order_date BETWEEN '" & startDate & "' AND '" & endDate & "' " & _
            "GROUP BY c.client_name " & _
            "ORDER BY total_sales DESC", cn

    Dim row As Integer: row = 5
    Do While Not rs.EOF
        ws.Cells(row, 1).Value = rs("client_name") & ""
        ws.Cells(row, 2).Value = CLng(rs("total_sales"))
        ws.Cells(row, 3).Value = CLng(rs("order_count"))
        ws.Cells(row, 2).NumberFormat = "#,##0"
        row = row + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    '── 合計行 ──
    If row > 5 Then
        ws.Cells(row, 1).Value = "合計"
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 2).Formula = "=SUM(B5:B" & row - 1 & ")"
        ws.Cells(row, 2).Font.Bold = True
        ws.Cells(row, 2).NumberFormat = "#,##0"
        ws.Cells(row, 3).Formula = "=SUM(C5:C" & row - 1 & ")"
        ws.Cells(row, 3).Font.Bold = True
        With ws.Range("A" & row & ":C" & row)
            .Interior.Color = RGB(220, 230, 241)
        End With
    End If

    ws.Columns("A:C").AutoFit
End Sub

'── 商品別売上シート ──
Private Sub CreateProductSheet(wb As Workbook, cn As Object, startDate As String, endDate As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = "商品別売上"

    '── ヘッダー ──
    ws.Range("A1").Value = "商品別売上レポート"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "集計期間：" & startDate & " ～ " & endDate
    ws.Range("A2").Font.Color = RGB(100, 100, 100)

    ws.Range("A4").Value = "商品名"
    ws.Range("B4").Value = "売上合計"
    ws.Range("C4").Value = "販売数量"

    With ws.Range("A4:C4")
        .Font.Bold = True
        .Interior.Color = RGB(47, 84, 150)
        .Font.Color = RGB(255, 255, 255)
    End With

    '── データ取得 ──
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COALESCE(p.product_name, od.free_item_name) AS item_name, " & _
            "SUM(od.quantity * od.unit_price) AS total_sales, " & _
            "SUM(od.quantity) AS total_qty " & _
            "FROM order_details od " & _
            "LEFT JOIN products p ON od.prod_id = p.prod_id " & _
            "LEFT JOIN orders o ON od.order_id = o.order_id " & _
            "WHERE o.order_date BETWEEN '" & startDate & "' AND '" & endDate & "' " & _
            "GROUP BY COALESCE(p.product_name, od.free_item_name) " & _
            "ORDER BY total_sales DESC", cn

    Dim row As Integer: row = 5
    Do While Not rs.EOF
        ws.Cells(row, 1).Value = rs("item_name") & ""
        ws.Cells(row, 2).Value = CLng(rs("total_sales"))
        ws.Cells(row, 3).Value = CLng(rs("total_qty"))
        ws.Cells(row, 2).NumberFormat = "#,##0"
        row = row + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    '── 合計行 ──
    If row > 5 Then
        ws.Cells(row, 1).Value = "合計"
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 2).Formula = "=SUM(B5:B" & row - 1 & ")"
        ws.Cells(row, 2).Font.Bold = True
        ws.Cells(row, 2).NumberFormat = "#,##0"
        ws.Cells(row, 3).Formula = "=SUM(C5:C" & row - 1 & ")"
        ws.Cells(row, 3).Font.Bold = True
        With ws.Range("A" & row & ":C" & row)
            .Interior.Color = RGB(220, 230, 241)
        End With
    End If

    ws.Columns("A:C").AutoFit
End Sub

'── 在庫アラートシート ──
Private Sub CreateStockAlertSheet(wb As Workbook, cn As Object)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = "在庫アラート"

    '── ヘッダー ──
    ws.Range("A1").Value = "在庫アラートレポート"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "出力日：" & Format(Date, "yyyy/mm/dd")
    ws.Range("A2").Font.Color = RGB(100, 100, 100)

    ws.Range("A4").Value = "商品名"
    ws.Range("B4").Value = "在庫数"
    ws.Range("C4").Value = "発注点"
    ws.Range("D4").Value = "状態"

    With ws.Range("A4:D4")
        .Font.Bold = True
        .Interior.Color = RGB(47, 84, 150)
        .Font.Color = RGB(255, 255, 255)
    End With

    '── データ取得（在庫不足 + 発注点以下）──
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT product_name, stock, reorder_point " & _
            "FROM products " & _
            "WHERE stock <= reorder_point " & _
            "ORDER BY stock ASC", cn

    Dim row As Integer: row = 5
    Do While Not rs.EOF
        Dim stock   As Long: stock = CLng(rs("stock"))
        Dim reorder As Long: reorder = CLng(rs("reorder_point"))

        ws.Cells(row, 1).Value = rs("product_name") & ""
        ws.Cells(row, 2).Value = stock
        ws.Cells(row, 3).Value = reorder

        If stock < 0 Then
            ws.Cells(row, 4).Value = "在庫不足"
            ws.Cells(row, 4).Font.Color = RGB(180, 0, 0)
            ws.Cells(row, 4).Font.Bold = True
            With ws.Range("A" & row & ":D" & row)
                .Interior.Color = RGB(255, 220, 220)
            End With
        Else
            ws.Cells(row, 4).Value = "発注点以下"
            ws.Cells(row, 4).Font.Color = RGB(180, 100, 0)
            With ws.Range("A" & row & ":D" & row)
                .Interior.Color = RGB(255, 243, 220)
            End With
        End If

        row = row + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ws.Columns("A:D").AutoFit
End Sub

