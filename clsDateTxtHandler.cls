VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDateTxtHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1

Private Sub txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    If txt.ForeColor = RGB(180, 180, 180) Then
        txt.Text = ""
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub txt_Enter()
Debug.Print "Enter fired: " & txt.Text
    If txt.ForeColor = RGB(180, 180, 180) Then
        txt.Text = ""
        txt.SelStart = 0
        txt.SelLength = 0
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub txt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txt.Text = "" Then
        txt.Text = Format(Date, "yyyy/mm/dd")
        txt.ForeColor = RGB(180, 180, 180)
    End If
End Sub
