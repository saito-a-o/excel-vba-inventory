VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmClientSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmClientSelect"
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
    Me.Caption = "取引先選択"
    Me.Width = 500
    Me.Height = 400
    Me.BackColor = RGB(240, 244, 251)

    '── 検索ラベル ──
    Dim lblSearch As MSForms.Label
    Set lblSearch = Me.Controls.Add("Forms.Label.1", "lblSearch", True)
    With lblSearch
        .Caption = "取引先名で検索："
        .Left = 14: .top = 14
        .Width = 120: .Height = 18
        .Font.Size = 10
    End With

    '── 検索テキストボックス ──
    Dim txtSearch As MSForms.TextBox
    Set txtSearch = Me.Controls.Add("Forms.TextBox.1", "txtSearch", True)
    With txtSearch
        .Left = 140: .top = 12
        .Width = 200: .Height = 22
        .Font.Size = 10
        .IMEMode = xlIMEModeHiragana
    End With

    '── 検索ボタン ──
    Dim btnSearch As MSForms.CommandButton
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    With btnSearch
        .Caption = "検索"
        .Left = 348: .top = 11
        .Width = 60: .Height = 24
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    '── 一覧リストボックス ──
    Dim lstClients As MSForms.ListBox
    Set lstClients = Me.Controls.Add("Forms.ListBox.1", "lstClients", True)
    With lstClients
        .Left = 14: .top = 46
        .Width = 460: .Height = 240
        .Font.Size = 10
        .ColumnCount = 3
        .ColumnWidths = "40;180;120"
    End With

    '── ボタン群 ──
    Dim btnNew As MSForms.CommandButton
    Set btnNew = Me.Controls.Add("Forms.CommandButton.1", "btnNew", True)
    With btnNew
        .Caption = "新規登録"
        .Left = 14: .top = 300
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    Dim btnEdit As MSForms.CommandButton
    Set btnEdit = Me.Controls.Add("Forms.CommandButton.1", "btnEdit", True)
    With btnEdit
        .Caption = "編集"
        .Left = 102: .top = 300
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
    End With

    Dim btnDelete As MSForms.CommandButton
    Set btnDelete = Me.Controls.Add("Forms.CommandButton.1", "btnDelete", True)
    With btnDelete
        .Caption = "削除"
        .Left = 190: .top = 300
        .Width = 80: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    Dim btnSelect As MSForms.CommandButton
    Set btnSelect = Me.Controls.Add("Forms.CommandButton.1", "btnSelect", True)
    With btnSelect
        .Caption = "選択して反映"
        .Left = 310: .top = 300
        .Width = 100: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With btnClose
        .Caption = "閉じる"
        .Left = 418: .top = 300
        .Width = 60: .Height = 26
        .Font.Size = 10
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
    End With

    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = btnSearch: h.act = "SEARCH_CLIENT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnNew: h.act = "NEW_CLIENT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnEdit: h.act = "EDIT_CLIENT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnDelete: h.act = "DELETE_CLIENT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnSelect: h.act = "SELECT_CLIENT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = btnClose: h.act = "CLOSE_CLIENT": Set h.frm = Me
    m_handlers.Add h

    '── 初期表示：全件取得 ──
    Call LoadClientList(Me, "")

End Sub

'── 取引先一覧を読み込む ──
Public Sub LoadClientList(frm As Object, keyword As String)
    Dim rs As Object
    Dim sql As String
    Set rs = CreateObject("ADODB.Recordset")

    If keyword = "" Then
        sql = "SELECT client_id, client_name, tel FROM clients " & _
              "WHERE is_deleted = FALSE ORDER BY client_id"
    Else
        sql = "SELECT client_id, client_name, tel FROM clients " & _
              "WHERE is_deleted = FALSE AND client_name LIKE '%" & keyword & "%' " & _
              "ORDER BY client_id"
    End If

    rs.Open sql, frm.m_cn

    Dim lst As MSForms.ListBox
    Set lst = frm.Controls("lstClients")
    lst.Clear

    Do While Not rs.EOF
        lst.AddItem rs("client_id") & ""
        lst.List(lst.ListCount - 1, 1) = rs("client_name") & ""
        lst.List(lst.ListCount - 1, 2) = rs("tel") & ""
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

