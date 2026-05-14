VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCmbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmb As MSForms.ComboBox
Attribute cmb.VB_VarHelpID = -1
Public act As String
Public frm As Object

Private Sub cmb_Change()
    Select Case act
        Case "CAT_CHANGED"
            Dim rs As Object
            Set rs = CreateObject("ADODB.Recordset")
            Dim sql As String
            Dim catVal As String
            catVal = frm.Controls("cmbCat").Text
            
            If catVal = "（すべて）" Or catVal = "" Then
                sql = "SELECT DISTINCT product_subcategory FROM products " & _
                        "WHERE product_subcategory IS NOT NULL AND " & _
                        "product_subcategory <> '' " & _
                        "ORDER BY product_subcategory"
            Else
                sql = "SELECT DISTINCT product_subcategory FROM products " & _
                        "WHERE product_subcategory IS NOT NULL AND " & _
                        "product_subcategory <> '' AND " & _
                        "product_category = '" & catVal & "'" & _
                        "ORDER BY product_subcategory"
            End If
            rs.Open sql, frm.m_cn
            
            Dim cmbSub As MSForms.ComboBox
            Set cmbSub = frm.Controls("cmbSub")
            cmbSub.Clear
            cmbSub.AddItem "（すべて）"
            Do While Not rs.EOF
                cmbSub.AddItem rs("product_subcategory") & ""
                rs.MoveNext
            Loop
            
            cmbSub.ListIndex = 0
            
            rs.Close
            Set rs = Nothing
            
            
            
        
    End Select
End Sub
