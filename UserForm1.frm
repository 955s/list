VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1870
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3170
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
       .ListWidth = 53
       .AddItem "包件列"
       .AddItem "合同列"
       .AddItem "付款列"
       .AddItem "发票列"
       .AddItem "合汇列"
       .AddItem "付款列"
    End With
    
End Sub

Private Sub CommandButton_Click()
    Application.ScreenUpdating = False   '禁刷新
    Call 隐藏1
    'Worksheets("透视表").Activate
    MsgTimeout "已隐藏"
    Application.ScreenUpdating = True   '恢刷新
End Sub

Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False   '禁刷新
    Call 取消隐藏
    'Worksheets("透视表").Activate
    MsgTimeout "已取消隐藏"
    Application.ScreenUpdating = True   '恢刷新
End Sub

Private Sub 包件_Click()
    Worksheets("上报告").Activate
    ActiveWindow.ScrollColumn = 8 '包件列
End Sub

Private Sub 合同_Click()
    Worksheets("上报告").Activate
    ActiveWindow.ScrollColumn = 20 '合同列
End Sub

Private Sub 付款_Click()
    Worksheets("上报告").Activate
    ActiveWindow.ScrollColumn = 26 '付款列
End Sub

Private Sub 发票_Click()
    Worksheets("上报告").Activate
    ActiveWindow.ScrollColumn = 38 '发票列
End Sub

Private Sub 合汇_Click()
    Worksheets("上报告").Activate
    ActiveWindow.ScrollColumn = 48 '合同执行情况汇总列
End Sub


Private Sub 清理_Click()
    On Error Resume Next '忽略错误继续执行
    Kill "E:\Download\物资付款单*.xlsx" '删除文件
    Kill "E:\Download\合同台账*.xls" '删除文件
    Kill "E:\Download\物资设备采购合同汇总查询*.xlsx" '删除文件
    MsgTimeout "已清理"
End Sub

Private Sub 跳转_Click()
    Worksheets("上报告").Activate
    cx = ComboBox1.Value
    If cx = "包件列" Then
            ActiveWindow.ScrollColumn = 8 '包件列
        ElseIf cx = "合同列" Then
            ActiveWindow.ScrollColumn = 20 '合同列
        ElseIf cx = "付款列" Then
            ActiveWindow.ScrollColumn = 26 '付款列
        ElseIf cx = "发票列" Then
            ActiveWindow.ScrollColumn = 38 '发票列
        ElseIf cx = "合汇列" Then
            ActiveWindow.ScrollColumn = 48 '合同执行情况汇总列
        Else
            ActiveWindow.ScrollColumn = 1 '包件列
    End If
End Sub
