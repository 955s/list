Attribute VB_Name = "创建超链接"
Sub 创建超链接()
    '不要用纯数字做表名字 否则报错
    Application.ScreenUpdating = False   '禁刷新

    If ActiveSheet.Name = "透视表" Then '判断只有激活了总索引表才更新目录

        Set shtIndex = ThisWorkbook.Sheets("透视表") '为方便之后调用，定义索引表名

        For i = 1 To ThisWorkbook.Worksheets.Count ''遍历所有工作表

            shtIndex.Cells(i, 17).Select '选中第二列的单元格

            With Selection:
                

                .Value = ThisWorkbook.Worksheets(i).Name '选中的单元格赋予工作表名称

                '在单元格中加上超链，链接到目标工作表的A1单元格

                .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=shtIndex.Cells(i, 17).Value & "!A1", TextToDisplay:=shtIndex.Cells(i, 17).Value

            End With
            
        Next '结束循环
        
    End If
    x = Sheets("透视表").Range("Q" & Rows.Count).End(xlUp).Row '查询当前表占用的行
    y = Sheets("透视表").Range("N" & Rows.Count).End(xlUp).Row '查询当前表占用的行
    Range("N14:N" & y).ClearContents
    
    Range("Q1:Q" & x).Select
    Selection.Cut
    Range("N14").Select
    ActiveSheet.Paste
    y = Sheets("透视表").Range("N" & Rows.Count).End(xlUp).Row '查询当前表占用的行
    Range("N14:N" & y).Font.Size = 10 '指定区域字号
    Range("N14:N" & y).Font.Name = "宋体" '指定区域字体
    Range("N14:N" & y).Font.FontStyle = "加粗" '指定区域字体
    Range("N14:N" & y).HorizontalAlignment = xlRight '右对齐
    
    
End Sub
