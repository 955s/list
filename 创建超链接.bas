Attribute VB_Name = "����������"
Sub ����������()
    '��Ҫ�ô������������� ���򱨴�
    Application.ScreenUpdating = False   '��ˢ��

    If ActiveSheet.Name = "͸�ӱ�" Then '�ж�ֻ�м�������������Ÿ���Ŀ¼

        Set shtIndex = ThisWorkbook.Sheets("͸�ӱ�") 'Ϊ����֮����ã�������������

        For i = 1 To ThisWorkbook.Worksheets.Count ''�������й�����

            shtIndex.Cells(i, 17).Select 'ѡ�еڶ��еĵ�Ԫ��

            With Selection:
                

                .Value = ThisWorkbook.Worksheets(i).Name 'ѡ�еĵ�Ԫ���蹤��������

                '�ڵ�Ԫ���м��ϳ��������ӵ�Ŀ�깤�����A1��Ԫ��

                .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=shtIndex.Cells(i, 17).Value & "!A1", TextToDisplay:=shtIndex.Cells(i, 17).Value

            End With
            
        Next '����ѭ��
        
    End If
    x = Sheets("͸�ӱ�").Range("Q" & Rows.Count).End(xlUp).Row '��ѯ��ǰ��ռ�õ���
    y = Sheets("͸�ӱ�").Range("N" & Rows.Count).End(xlUp).Row '��ѯ��ǰ��ռ�õ���
    Range("N14:N" & y).ClearContents
    
    Range("Q1:Q" & x).Select
    Selection.Cut
    Range("N14").Select
    ActiveSheet.Paste
    y = Sheets("͸�ӱ�").Range("N" & Rows.Count).End(xlUp).Row '��ѯ��ǰ��ռ�õ���
    Range("N14:N" & y).Font.Size = 10 'ָ�������ֺ�
    Range("N14:N" & y).Font.Name = "����" 'ָ����������
    Range("N14:N" & y).Font.FontStyle = "�Ӵ�" 'ָ����������
    Range("N14:N" & y).HorizontalAlignment = xlRight '�Ҷ���
    
    
End Sub
