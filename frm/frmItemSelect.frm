VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmItemSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmItemSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_handlers  As Collection
Private m_txtHandlers As Collection
Public m_cn         As Object
Public m_rowCount   As Integer  ' 現在の行数
Public m_selectedRow As Integer ' 選択中の行（0=未選択）

' 定数：通常時・選択時の色
Private Const CLR_NAME_BACK_NORMAL   As Long = &HF0F0F0  ' RGB(240,240,240)
Private Const CLR_NAME_FORE_NORMAL   As Long = &H646464  ' RGB(100,100,100)
Private Const CLR_NAME_BACK_SELECTED As Long = &HFFD2B4  ' RGB(180,210,255)
Private Const CLR_NAME_FORE_SELECTED As Long = &H0       ' RGB(0,0,0)

Private Sub UserForm_Initialize()
    Set m_handlers = New Collection
    Set m_txtHandlers = New Collection
    Set m_cn = GetConnection()
    m_rowCount = 0
    m_selectedRow = 0

    Me.Caption = "商品選択"
    Me.Width = 460
    Me.BackColor = RGB(240, 244, 251)

    Const LX     As Integer = 14
    Const SP     As Integer = 32
    Const ROW_H  As Integer = 20  ' 行の高さ
    Const TW_N   As Integer = 196 ' 商品名幅（旧ListBox幅と同じ）
    Const TW_P   As Integer = 76  ' 単価幅
    Const TW_Q   As Integer = 56  ' 数量幅
    Const CW_T   As Integer = 56  ' 税率幅

    Dim topY As Integer: topY = 14

    '── タイトル ──
    With Lbl("lblTitle", "明細の商品を選択してください", LX, topY, 400, 22)
        .Font.Size = 11: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 34

    '── 列ヘッダーラベル ──
    Dim hNames  As Variant: hNames = Array("商品名", "単価", "数量", "税率")
    Dim hLefts  As Variant: hLefts = Array(LX, LX + 196, LX + 280, LX + 344)
    Dim hWidths As Variant: hWidths = Array(196, 76, 56, 56)
    Dim hi As Integer
    For hi = 0 To 3
        With Lbl("lblH" & hi, CStr(hNames(hi)), CInt(hLefts(hi)), topY, CInt(hWidths(hi)), 16)
            .Font.Bold = True
            .ForeColor = RGB(31, 56, 100)
            .Font.Size = 9
        End With
    Next hi
    topY = topY + 20

    '── 商品名TextBox・単価TextBox・数量TextBox・税率ComboBox（10行分）──
    Dim i As Integer
    For i = 1 To 10
        Dim rowTop As Integer: rowTop = topY + (i - 1) * ROW_H

        '── 商品名（編集不可・薄グレー）──
        Dim tn As MSForms.TextBox
        Set tn = Me.Controls.Add("Forms.TextBox.1", "txtName_" & i, True)
        With tn
            .Left = LX: .top = rowTop
            .Width = TW_N: .Height = ROW_H
            .Font.Size = 9
            .Locked = True
            .BackColor = CLR_NAME_BACK_NORMAL
            .ForeColor = CLR_NAME_FORE_NORMAL
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleSingle
            .Text = ""
        End With

        '── 単価 ──
        Dim tp As MSForms.TextBox
        Set tp = Me.Controls.Add("Forms.TextBox.1", "txtPrice_" & i, True)
        With tp
            .Left = LX + 196: .top = rowTop
            .Width = TW_P: .Height = ROW_H
            .Font.Size = 9
            .IMEMode = xlIMEModeDisable
            .Text = ""
        End With

        '── 数量 ──
        Dim tq As MSForms.TextBox
        Set tq = Me.Controls.Add("Forms.TextBox.1", "txtQty_" & i, True)
        With tq
            .Left = LX + 280: .top = rowTop
            .Width = TW_Q: .Height = ROW_H
            .Font.Size = 9
            .IMEMode = xlIMEModeDisable
            .Text = ""
        End With

        '── 税率（ComboBox）──
        Dim ct As MSForms.ComboBox
        Set ct = Me.Controls.Add("Forms.ComboBox.1", "cmbTax_" & i, True)
        With ct
            .Left = LX + 344: .top = rowTop
            .Width = CW_T: .Height = ROW_H
            .Font.Size = 9
            .Style = fmStyleDropDownList
            .AddItem "10%"
            .AddItem "8%"
            .ListIndex = 0
        End With
        
        '── 商品名Textboxイベントハンドラ登録 ──
        Dim ht As clsTxtHandler
        Set ht = New clsTxtHandler
        Set ht.txt = tn: ht.rowIdx = i: Set ht.frm = Me
        m_txtHandlers.Add ht
    Next i

    topY = topY + 10 * ROW_H + 8

    '── ボタン群 ──
    Dim btnAdd As MSForms.CommandButton
    Set btnAdd = Me.Controls.Add("Forms.CommandButton.1", "btnAddItem", True)
    With btnAdd
        .Caption = "商品を追加"
        .Left = LX: .top = topY
        .Width = 90: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    Dim btnDel As MSForms.CommandButton
    Set btnDel = Me.Controls.Add("Forms.CommandButton.1", "btnDelItem", True)
    With btnDel
        .Caption = "行を削除"
        .Left = LX + 98: .top = topY
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancelItem", True)
    With btnCancel
        .Caption = "キャンセル"
        .Left = LX + 240: .top = topY
        .Width = 90: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    Dim btnOK As MSForms.CommandButton
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOKItem", True)
    With btnOK
        .Caption = "確定"
        .Left = LX + 338: .top = topY
        .Width = 90: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    topY = topY + 34

    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = btnAdd: h.act = "ADD_ITEM": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnDel: h.act = "DEL_ITEM": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnCancel: h.act = "CANCEL_ITEM": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnOK: h.act = "OK_ITEM": Set h.frm = Me
    m_handlers.Add h
   
    Me.Height = topY + 40

    '── 入力シートの既存データを読み込む ──
    Call LoadSheetData
End Sub

'── 行を選択状態にする（ハイライト切り替え）──
Public Sub SelectRow(rowIdx As Integer)
    ' 前の選択行を通常色に戻す
    If m_selectedRow > 0 Then
        Me.Controls("txtName_" & m_selectedRow).BackColor = CLR_NAME_BACK_NORMAL
        Me.Controls("txtName_" & m_selectedRow).ForeColor = CLR_NAME_FORE_NORMAL
    End If
    ' 新しい行をハイライト
    m_selectedRow = rowIdx
    If m_selectedRow > 0 Then
        Me.Controls("txtName_" & m_selectedRow).BackColor = CLR_NAME_BACK_SELECTED
        Me.Controls("txtName_" & m_selectedRow).ForeColor = CLR_NAME_FORE_SELECTED
    End If
End Sub

'── 入力シートの明細を初期値としてセット ──
Public Sub LoadSheetData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("入力")
    m_rowCount = 0
    m_selectedRow = 0

    ' 全行クリア
    Dim i As Integer
    For i = 1 To 10
        Me.Controls("txtName_" & i).Text = ""
        Me.Controls("txtPrice_" & i).Text = ""
        Me.Controls("txtQty_" & i).Text = ""
        Me.Controls("cmbTax_" & i).ListIndex = 0
    Next i

    For i = 0 To 9
        Dim itemName As String
        itemName = Trim(ws.Range(SRC_RANGE_DETAILITEM1).Offset(i, 0).MergeArea.Cells(1, 1).Value)
        If itemName = "" Then Exit For

        Dim price   As String: price = ws.Range(SRC_RANGE_DETAILPRCE1).Offset(i, 0).Value & ""
        Dim qty     As String: qty = ws.Range(SRC_RANGE_DETAILNUM1).Offset(i, 0).Value & ""
        Dim taxRate As String: taxRate = ws.Range(SRC_RANGE_DETAILRATE1).Offset(i, 0).Value & ""

        m_rowCount = m_rowCount + 1
        Me.Controls("txtName_" & m_rowCount).Text = itemName
        Me.Controls("txtPrice_" & m_rowCount).Text = price
        Me.Controls("txtQty_" & m_rowCount).Text = qty
        If Trim(StrConv(taxRate, vbNarrow)) = TAX_RATE2 Then
            Me.Controls("cmbTax_" & m_rowCount).ListIndex = 1
        ElseIf Trim(StrConv(taxRate, vbNarrow)) = TAX_RATE1 Then
            Me.Controls("cmbTax_" & m_rowCount).ListIndex = 0
        End If
    Next i
    '── フォーカスを単価1行目に設定 ──
    Me.Controls("txtPrice_1").SetFocus
End Sub

'── 子フォームから行を追加するメソッド ──
Public Sub AddItemRow(prodName As String, price As Long, taxRate As String)
    If m_rowCount >= 10 Then
        MsgBox "明細は10行までです。", vbExclamation, "エラー"
        Exit Sub
    End If
    m_rowCount = m_rowCount + 1
    Me.Controls("txtName_" & m_rowCount).Text = prodName
    Me.Controls("txtPrice_" & m_rowCount).Text = price & ""
    Me.Controls("txtQty_" & m_rowCount).Text = "1"
    If taxRate = "8%" Then
        Me.Controls("cmbTax_" & m_rowCount).ListIndex = 1
    Else
        Me.Controls("cmbTax_" & m_rowCount).ListIndex = 0
    End If
End Sub

Private Sub UserForm_Terminate()
    If Not m_cn Is Nothing Then
        If m_cn.State = 1 Then m_cn.Close
    End If
    Set m_cn = Nothing
End Sub

Private Function Lbl(nm$, cap$, l%, t%, w%, h%) As MSForms.Label
    Set Lbl = Me.Controls.Add("Forms.Label.1", nm, True)
    With Lbl
        .Caption = cap: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
    End With
End Function
