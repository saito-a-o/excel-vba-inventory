VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSearch 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmItemSearch.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_cn        As Object
Public m_parentFrm As Object
Private m_handlers As Collection
Private m_cmbHandlers As Collection

Private Sub UserForm_Initialize()
    Set m_handlers = New Collection

    Me.Caption = "商品検索"
    Me.Width = 480
    Me.Height = 420
    Me.BackColor = RGB(240, 244, 251)

    Const LX As Integer = 14
    Dim topY As Integer: topY = 14

    '── タイトル ──
    With Lbl("lblTitle", "追加する商品を選択してください", LX, topY, 420, 22)
        .Font.Size = 11: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 34

    '── カテゴリ ──
    Lbl "lblCat", "カテゴリ", LX, topY + 4, 80, 18
    Dim cmbCat As MSForms.ComboBox
    Set cmbCat = Me.Controls.Add("Forms.ComboBox.1", "cmbCat", True)
    With cmbCat
        .Left = LX + 86: .top = topY
        .Width = 120: .Height = 24
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .AddItem "（すべて）"
        .ListIndex = 0
    End With
    
    '── カテゴリの変更イベント登録 ──
    Set m_cmbHandlers = New Collection
    Dim hc As clsCmbHandler
    Set hc = New clsCmbHandler
    Set hc.cmb = cmbCat: hc.act = "CAT_CHANGED": Set hc.frm = Me
    m_cmbHandlers.Add hc
    
    '── サブカテゴリ ──
    Lbl "lblSubCat", "サブカテゴリ", LX + 216, topY + 4, 90, 18
    Dim cmbSub As MSForms.ComboBox
    Set cmbSub = Me.Controls.Add("Forms.ComboBox.1", "cmbSub", True)
    With cmbSub
        .Left = LX + 314: .top = topY
        .Width = 120: .Height = 24
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .AddItem "（すべて）"
        .ListIndex = 0
    End With
    topY = topY + 32

    '── キーワード検索 ──
    Lbl "lblKw", "キーワード", LX, topY + 4, 80, 18
    Dim txtKw As MSForms.TextBox
    Set txtKw = Me.Controls.Add("Forms.TextBox.1", "txtKeyword", True)
    With txtKw
        .Left = LX + 86: .top = topY
        .Width = 200: .Height = 24
        .Font.Size = 10
        .IMEMode = xlIMEModeHiragana
    End With

    Dim btnSearch As MSForms.CommandButton
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    With btnSearch
        .Caption = "検索"
        .Left = LX + 294: .top = topY
        .Width = 60: .Height = 24
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With
    topY = topY + 32
    
     '── 複数選択注記 ──
    With Lbl("lblMulti", "※Ctrlキーを押しながらクリックすると複数選択できます", LX, topY, 420, 16)
        .Font.Size = 9
        .ForeColor = RGB(100, 100, 100)
    End With
    topY = topY + 18

    '── リストボックス（複数選択可）──
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls.Add("Forms.ListBox.1", "lstSearch", True)
    With lst
        .Left = LX: .top = topY
        .Width = 440: .Height = 200
        .Font.Size = 10
        .ColumnCount = 4
        .ColumnWidths = "36;160;60;60"
        .MultiSelect = fmMultiSelectMulti
    End With
    topY = topY + 210

    '── ボタン ──
    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancelSearch", True)
    With btnCancel
        .Caption = "キャンセル"
        .Left = LX + 220: .top = topY
        .Width = 96: .Height = 28
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    Dim btnSelect As MSForms.CommandButton
    Set btnSelect = Me.Controls.Add("Forms.CommandButton.1", "btnSelectItem", True)
    With btnSelect
        .Caption = "選択した商品を追加"
        .Left = LX + 324: .top = topY
        .Width = 130: .Height = 28
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = btnSearch: h.act = "SEARCH_ITEM": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnCancel: h.act = "CANCEL_ITEM_SEARCH": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnSelect: h.act = "SELECT_ITEM": Set h.frm = Me
    m_handlers.Add h

End Sub

'── カテゴリ一覧をDBから取得 ──
Public Sub LoadCategories()
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT DISTINCT product_category FROM products " & _
            "WHERE product_category IS NOT NULL AND " & _
            "product_category <> '' " & _
            "ORDER BY product_category", m_cn
    Dim cmbCat As MSForms.ComboBox
    Set cmbCat = Me.Controls("cmbCat")
    Do While Not rs.EOF
        cmbCat.AddItem rs("product_category") & ""
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

'── frmItemSearch に追加 ──
Public Sub LoadData()
    Call LoadCategories
    Call LoadItemList("")
End Sub

'── 商品一覧を検索して表示 ──
Public Sub LoadItemList(keyword As String)
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    Dim sql As String
    Dim catVal As String: catVal = Me.Controls("cmbCat").Text
    Dim subVal As String: subVal = Me.Controls("cmbSub").Text

    sql = "SELECT prod_id, product_name, price, tax_rate FROM products WHERE 1=1"
    If catVal <> "（すべて）" And catVal <> "" Then
        sql = sql & " AND product_category = '" & catVal & "'"
    End If
    If subVal <> "（すべて）" And subVal <> "" Then
        sql = sql & " AND product_subcategory = '" & subVal & "'"
    End If
    If keyword <> "" Then
        sql = sql & " AND product_name LIKE '%" & keyword & "%'"
    End If
    sql = sql & " ORDER BY prod_id"

    rs.Open sql, m_cn

    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstSearch")
    lst.Clear

    Do While Not rs.EOF
        lst.AddItem rs("prod_id") & ""
        lst.List(lst.ListCount - 1, 1) = rs("product_name") & ""
        lst.List(lst.ListCount - 1, 2) = rs("price") & ""
        lst.List(lst.ListCount - 1, 3) = rs("tax_rate") & ""
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Function Lbl(nm$, cap$, l%, t%, w%, h%) As MSForms.Label
    Set Lbl = Me.Controls.Add("Forms.Label.1", nm, True)
    With Lbl
        .Caption = cap: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
    End With
End Function
