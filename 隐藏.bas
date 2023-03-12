Attribute VB_Name = "隐藏"
Sub 隐藏供货单列()
    With Worksheets("供货单")
        .Columns(4).Hidden = True
        .Columns(6).Hidden = True
        .Columns("H:J").Hidden = True
        .Columns("M:N").Hidden = True
    End With
End Sub

Sub 取消隐藏隐藏供货单列()
    With Worksheets("供货单")
        .Columns(4).Hidden = False
        .Columns(6).Hidden = False
        .Columns("H:J").Hidden = False
        .Columns("M:N").Hidden = False
    End With
End Sub

Sub 隐藏()
    Application.ScreenUpdating = False '禁刷新
            For i = 1 To ThisWorkbook.Worksheets.Count '遍历文件的工作表数
                表名 = ThisWorkbook.Worksheets(i).Name
                If InStr(表名, "出") Then
                    Worksheets(表名).Select
                    ActiveWindow.SelectedSheets.Visible = False
                End If
            Next i
End Sub

Sub 隐藏1()
    Call 取消隐藏
    Worksheets(Array("B", "交底", "交底_透", "计价", _
    "转辙机", "长缆", "lb", "透视表1")).Select
    
    ActiveWindow.SelectedSheets.Visible = False
End Sub

Sub 取消隐藏()
    Dim sht As Worksheet
    For Each sht In Worksheets
    sht.Visible = xlSheetVisible
    Next
End Sub

Sub 隐藏2()

    Call 取消隐藏
        Worksheets("B").Select
    ActiveWindow.SelectedSheets.Visible = False
        Worksheets("交底").Select
    ActiveWindow.SelectedSheets.Visible = False
            Worksheets("交底_透").Select
    ActiveWindow.SelectedSheets.Visible = False
    
        Worksheets("J7").Select
    ActiveWindow.SelectedSheets.Visible = False
        Worksheets("计价").Select
    ActiveWindow.SelectedSheets.Visible = False
        Worksheets("lb").Select
    ActiveWindow.SelectedSheets.Visible = False
        Worksheets("转辙机").Select
    ActiveWindow.SelectedSheets.Visible = False
        Worksheets("长缆").Select
    ActiveWindow.SelectedSheets.Visible = False

End Sub



