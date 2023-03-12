VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 查询窗口 
   Caption         =   "输入要查询的供应商"
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   2470
   OleObjectBlob   =   "查询窗口.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "查询窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    Dim arr1, cx, i, c
    c = ActiveSheet.[AZ65535].End(xlUp).Row
    cx = TextBox1.Value
    arr1 = Range("AZ2:AZ" & c)
    For i = 1 To UBound(arr1)
        If InStr(arr1(i, 1), TextBox1.Value) > 0 Then
            n = i + 1
        End If
    Next
    Range("AZ" & n).Select
End Sub
