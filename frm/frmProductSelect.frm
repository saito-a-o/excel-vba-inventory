VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProductSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmProductSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProductSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_cn As Object
Private m_handlers As Collection

Private Sub UserForm_Initialize()
    Set m_cn = GetConnection()
    Set m_handlers = New Collection

    '── フォーム外観 ──
    Me.Caption = "商品マスタ"
    Me.Width = 580
    Me.Height = 420
    Me.BackColor = RGB(240, 244, 251)

    '── 検索ラベル ──
    Dim lblSearch As MSForms.Label
    Set lblSearch = Me.Controls.Add("Forms.Label.1", "lblSearch", True)
    With lblSearch
        .Caption = "商品名で検索："
        .Left = 14: .top = 14
        .Width = 110: .Height = 18
        .Font.Size = 10
    End With

    '── 検索テキストボックス ──
    Dim txtSearch As MSForms.TextBox
    Set txtSearch = Me.Controls.Add("Forms.TextBox.1", "txtSearch", True)
    With txtSearch
        .Left = 130: .top = 12
        .Width = 200: .Height = 22
        .Font.Size = 10
        .IMEMode = xlIMEModeHiragana
    End With

    '── 検索ボタン ──
    Dim btnSearch As MSForms.CommandButton
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    With btnSearch
        .Caption = "検索"
        .Left = 338: .top = 11
        .Width = 60: .Height = 24
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    '── 一覧リストボックス ──
    Dim lstProducts As MSForms.ListBox
    Set lstProducts = Me.Controls.Add("Forms.ListBox.1", "lstProducts", True)
    With lstProducts
        .Left = 14: .top = 62
        .Width = 540: .Height = 256
        .Font.Size = 10
        .ColumnCount = 6
        .ColumnWidths = "36;160;80;60;60;60"
    End With

    '── 列ヘッダーラベル ──
    Dim headers As Variant
    headers = Array("ID", "商品名", "サブカテゴリ", "単価", "在庫", "税率")
    Dim hLeft As Variant
    hLeft = Array(14, 50, 210, 290, 350, 410)
    Dim hWidth As Variant
    hWidth = Array(36, 160, 80, 60, 60, 60)
    Dim hi As Integer
    For hi = 0 To 5
        Dim lh As MSForms.Label
        Set lh = Me.Controls.Add("Forms.Label.1", "lblH" & hi, True)
        With lh
            .Caption = headers(hi)
            .Left = hLeft(hi): .top = 46
            .Width = hWidth(hi): .Height = 16
            .Font.Size = 9: .Font.Bold = True
            .ForeColor = RGB(31, 56, 100)
        End With
    Next hi

    '── ボタン群 ──
    Dim btnNew As MSForms.CommandButton
    Set btnNew = Me.Controls.Add("Forms.CommandButton.1", "btnNew", True)
    With btnNew
        .Caption = "新規登録"
        .Left = 14: .top = 316
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    Dim btnEdit As MSForms.CommandButton
    Set btnEdit = Me.Controls.Add("Forms.CommandButton.1", "btnEdit", True)
    With btnEdit
        .Caption = "編集"
        .Left = 102: .top = 316
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    Dim btnDelete As MSForms.CommandButton
    Set btnDelete = Me.Controls.Add("Forms.CommandButton.1", "btnDelete", True)
    With btnDelete
        .Caption = "削除"
        .Left = 190: .top = 316
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With btnClose
        .Caption = "閉じる"
        .Left = 478: .top = 316
        .Width = 76: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = btnSearch: h.act = "SEARCH_PRODUCT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnNew: h.act = "NEW_PRODUCT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnEdit: h.act = "EDIT_PRODUCT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnDelete: h.act = "DELETE_PRODUCT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnClose: h.act = "CLOSE_PRODUCT": Set h.frm = Me
    m_handlers.Add h

    '── 初期表示：全件取得 ──
    Call LoadProductList(Me, "")

End Sub

'── 商品一覧を読み込む ──
Public Sub LoadProductList(frm As Object, keyword As String)
    Dim rs As Object
    Dim sql As String
    Set rs = CreateObject("ADODB.Recordset")

    If keyword = "" Then
        sql = "SELECT prod_id, product_name, product_subcategory, " & _
              "price, stock, tax_rate FROM products ORDER BY prod_id"
    Else
        sql = "SELECT prod_id, product_name, product_subcategory, " & _
              "price, stock, tax_rate FROM products " & _
              "WHERE product_name LIKE '%" & keyword & "%' " & _
              "ORDER BY prod_id"
    End If

    rs.Open sql, frm.m_cn

    Dim lst As MSForms.ListBox
    Set lst = frm.Controls("lstProducts")
    lst.Clear

    Do While Not rs.EOF
        lst.AddItem rs("prod_id") & ""
        lst.List(lst.ListCount - 1, 1) = rs("product_name") & ""
        lst.List(lst.ListCount - 1, 2) = rs("product_subcategory") & ""
        lst.List(lst.ListCount - 1, 3) = rs("price") & ""
        lst.List(lst.ListCount - 1, 4) = rs("stock") & ""
        lst.List(lst.ListCount - 1, 5) = rs("tax_rate") & ""
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub UserForm_Terminate()
    If Not m_cn Is Nothing Then
        If m_cn.State = 1 Then m_cn.Close
    End If
    Set m_cn = Nothing
End Sub

