VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProductEdit 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmProductEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProductEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_mode      As String   ' "NEW" or "EDIT"
Public m_cn        As Object
Public m_prodId    As Long
Public m_parentFrm As Object

Private m_handlers As Collection

Private Sub UserForm_Initialize()
    Set m_handlers = New Collection

    Me.Width = 420
    Me.BackColor = RGB(240, 244, 251)

    Const LX As Integer = 14
    Const TX As Integer = 160
    Const LW As Integer = 140
    Const TW As Integer = 222
    Const TH As Integer = 24
    Const SP As Integer = 34

    Dim topY As Integer: topY = 14
    Dim hira As String: hira = "hiragana"
    Dim hank As String: hank = "hankaku"

    '── タイトルラベル ──
    With Lbl("lblTitle", "", LX, topY, 380, 28)
        .Font.Size = 13: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 38

    '── 商品名 ──
    Lbl "lblName", "商品名 ※必須", LX, topY + 4, LW, 18
    txt "txtName", "", hira, TX, topY, TW, TH
    topY = topY + SP

    '── カテゴリ ──
    Lbl "lblCat", "カテゴリ", LX, topY + 4, LW, 18
    txt "txtCat", "", hira, TX, topY, TW, TH
    topY = topY + SP

    '── サブカテゴリ ──
    Lbl "lblSubCat", "サブカテゴリ", LX, topY + 4, LW, 18
    txt "txtSubCat", "", hira, TX, topY, TW, TH
    topY = topY + SP

    '── 単位 ──
    Lbl "lblUnit", "単位", LX, topY + 4, LW, 18
    txt "txtUnit", "", hira, TX, topY, 80, TH
    topY = topY + SP

    '── 単価 ──
    Lbl "lblPrice", "単価（税抜）", LX, topY + 4, LW, 18
    txt "txtPrice", "", hank, TX, topY, 100, TH
    topY = topY + SP

    '── 税率 ──
    Lbl "lblTax", "税率", LX, topY + 4, LW, 18
    Dim cmbTax As MSForms.ComboBox
    Set cmbTax = Me.Controls.Add("Forms.ComboBox.1", "cmbTax", True)
    With cmbTax
        .Left = TX: .top = topY
        .Width = 80: .Height = TH
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .AddItem "10%"
        .AddItem "8%"
        .ListIndex = 0  ' デフォルト：10%
    End With
    topY = topY + SP

    '── 在庫数 ──
    Lbl "lblStock", "在庫数", LX, topY + 4, LW, 18
    txt "txtStock", "0", hank, TX, topY, 100, TH
    topY = topY + SP

    '── 発注点 ──
    Lbl "lblReorder", "発注点（在庫アラート）", LX, topY + 4, LW, 18
    txt "txtReorder", "0", hank, TX, topY, 100, TH
    topY = topY + SP + 8
    
    '── ボタン（キャンセル / 保存） ──
    Const BW As Integer = 96
    Const BH As Integer = 30
    Const BTop As Integer = 10

    Dim bCancel As MSForms.CommandButton
    Set bCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancelProd", True)
    With bCancel
        .Caption = "キャンセル"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 - BW - 16
        .Width = BW: .Height = BH
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
        .Font.Size = 11
    End With

    Dim bOK As MSForms.CommandButton
    Set bOK = Me.Controls.Add("Forms.CommandButton.1", "btnOKProd", True)
    With bOK
        .Caption = "保存"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 + 16
        .Width = BW: .Height = BH
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = bCancel: h.act = "CANCEL_PRODUCT_EDIT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = bOK: h.act = "OK_PRODUCT_EDIT": Set h.frm = Me
    m_handlers.Add h
    Me.Height = topY + BH + BTop + 52
End Sub

Private Function Lbl(nm$, cap$, l%, t%, w%, h%) As MSForms.Label
    Set Lbl = Me.Controls.Add("Forms.Label.1", nm, True)
    With Lbl
        .Caption = cap: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
    End With
End Function

Private Function txt(nm$, val$, ime$, l%, t%, w%, h%) As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", nm, True)
    With txt
        .Text = val: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
        If ime = "hiragana" Then
            .IMEMode = xlIMEModeHiragana
        Else
            .IMEMode = xlIMEModeDisable
        End If
    End With
End Function

Public Sub LoadData()
    If m_mode = "EDIT" Then
        Dim rs As Object
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM products WHERE prod_id = " & m_prodId, m_cn
        If Not rs.EOF Then
            Me.Controls("txtName").Text = rs("product_name") & ""
            Me.Controls("txtCat").Text = rs("product_category") & ""
            Me.Controls("txtSubCat").Text = rs("product_subcategory") & ""
            Me.Controls("txtUnit").Text = rs("unit") & ""
            Me.Controls("txtPrice").Text = rs("price") & ""
            Me.Controls("txtStock").Text = rs("stock") & ""
            Me.Controls("txtReorder").Text = rs("reorder_point") & ""
            Dim taxval As String
            taxval = rs("tax_rate") & ""
            If taxval = "8%" Then
                Me.Controls("cmbTax").ListIndex = 1
            Else
                Me.Controls("cmbTax").ListIndex = 0
            End If
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub
