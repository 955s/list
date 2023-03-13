Attribute VB_Name = "合同台账_修正"
Sub 修正合同台账包件()
    Application.ScreenUpdating = False   '禁刷新
    On Error Resume Next
    Dim arr1, arr, cx, i, c
    Dim 批次, 厂家, 表行
    With Worksheets("上报告")
    
    Worksheets("上报告").Activate
    表名 = ActiveSheet.Name
    表行 = Sheets("上报告").Range("W" & Rows.Count).End(xlUp).Row '查询当前表占用的行
    Worksheets("上报告").Activate
    .Range("U2:V200").Interior.Pattern = xlNone '清理填充
    .Range("U2:V200").ClearComments '清理批注
    arr1 = Range("W2:W" & 表行)

    '循环查找第一次出现的位置，修正ZH-03
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZqybd-2022-0037") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n & ":" & "V" & n).ClearContents
    .Range("V" & n).Value = "电缆附件"
    .Range("U" & n).Value = "ZH-03"
    .Range("V" & n).Interior.Color = 15773696
    .Range("U" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正BD-13
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZqybd-2022-0049") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "BD-13"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-13
    arr4 = Range("W2:W" & 表行)
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ--WZxh-2022-0019") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "XH-13"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正ZH-02
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZqt-2022-0041") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "ZH-02"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")

    
    '循环查找第一次出现的位置，修正XH-23
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZxh-2022-0053") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "XH-23"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正DL-10
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZdl-2022-0055") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "DL-10"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正JCW-05
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZjcw-2022-0020") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "JCW-05"
    .Range("V" & n).Value = "钢芯铝绞线"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正JCW-10
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZjcw-2022-0056") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n & ":" & "V" & n).ClearContents
    .Range("U" & n).Value = "JCW-10"
    .Range("V" & n).Value = "H型钢柱1"
    .Range("U" & n & ":" & "V" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正ZH-05
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZqt-2022-0019") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "ZH-05"
    .Range("V" & n).Value = "电缆附件1"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正JCW-13包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZjcw-2022-0035") Then
            n = i + 1
        Exit For
            End If
    Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "铜绞线"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正JCW-07包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZjcw-2022-0050") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "电缆卡具"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正TX-10包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZtx-2022-0046") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "综合配线柜"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-17包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ--WZxh-2022-0017") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "道岔缺口监测"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-27包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZxh-2022-0027") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "道岔缺口监测1"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-18包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ--WZxh-2022-0018") Then
            n = i + 1
        Exit For
            End If
        Next
    '.Range("V" & n & ":" & "X" & n).ClearContents
    .Range("V" & n).Value = "高压脉冲及补充"
    .Range("X" & n).Value = "239130.00"
    '.Range("W" & n).Value = "CRDH0004-CSXZ-WZxh-2022-0002"
    .Range("U" & n & ":" & "X" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-24包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZxh-2022-0028") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "防雷分线盘1"

    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-29包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZxh-2022-0061") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "线料、固线器1"

    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正JCW-16包件
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZjcw-2022-0058") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "电缆卡具1"

    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正ZH-07
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZqt-2022-0057") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "ZH-07"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正DL-17
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZdl-2022-0057") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("U" & n).ClearContents
    .Range("U" & n).Value = "DL-17"
    .Range("U" & n).Interior.Color = 15773696
    .Range("U" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-11
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZxh-2022-0064") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "道岔报警设备1"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正XH-26
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZxh-2022-0065") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "信号机1"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正JCW-17 金具1
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZ-wj-2022-0024") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "金具1"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正TX-14   配线单元
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZ-txcl-2022-0024") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "配线单元"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正DL-18   户外隔离开关柜1
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZdl-2022-0051（2）") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "户外隔离开关柜1"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    
    '循环查找第一次出现的位置，修正DL-20   低压电缆3
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "CRDH0004-CSXZ-WZ-dlcl-2022-0024") Then
            n = i + 1
        Exit For
            End If
        Next
    .Range("V" & n).ClearContents
    .Range("V" & n).Value = "低压电缆3"
    .Range("V" & n).Interior.Color = 15773696
    .Range("V" & n).AddCommentThreaded ("蓝色为包件合并的合同，手动修正状态")
    End With
End Sub




