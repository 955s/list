VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1870
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3170
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
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
       .AddItem "������"
       .AddItem "��ͬ��"
       .AddItem "������"
       .AddItem "��Ʊ��"
       .AddItem "�ϻ���"
       .AddItem "������"
    End With
    
End Sub

Private Sub CommandButton_Click()
    Application.ScreenUpdating = False   '��ˢ��
    Call ����1
    'Worksheets("͸�ӱ�").Activate
    MsgTimeout "������"
    Application.ScreenUpdating = True   '��ˢ��
End Sub

Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False   '��ˢ��
    Call ȡ������
    'Worksheets("͸�ӱ�").Activate
    MsgTimeout "��ȡ������"
    Application.ScreenUpdating = True   '��ˢ��
End Sub

Private Sub ����_Click()
    Worksheets("�ϱ���").Activate
    ActiveWindow.ScrollColumn = 8 '������
End Sub

Private Sub ��ͬ_Click()
    Worksheets("�ϱ���").Activate
    ActiveWindow.ScrollColumn = 20 '��ͬ��
End Sub

Private Sub ����_Click()
    Worksheets("�ϱ���").Activate
    ActiveWindow.ScrollColumn = 26 '������
End Sub

Private Sub ��Ʊ_Click()
    Worksheets("�ϱ���").Activate
    ActiveWindow.ScrollColumn = 38 '��Ʊ��
End Sub

Private Sub �ϻ�_Click()
    Worksheets("�ϱ���").Activate
    ActiveWindow.ScrollColumn = 48 '��ִͬ�����������
End Sub


Private Sub ����_Click()
    On Error Resume Next '���Դ������ִ��
    Kill "E:\Download\���ʸ��*.xlsx" 'ɾ���ļ�
    Kill "E:\Download\��̨ͬ��*.xls" 'ɾ���ļ�
    Kill "E:\Download\�����豸�ɹ���ͬ���ܲ�ѯ*.xlsx" 'ɾ���ļ�
    MsgTimeout "������"
End Sub

Private Sub ��ת_Click()
    Worksheets("�ϱ���").Activate
    cx = ComboBox1.Value
    If cx = "������" Then
            ActiveWindow.ScrollColumn = 8 '������
        ElseIf cx = "��ͬ��" Then
            ActiveWindow.ScrollColumn = 20 '��ͬ��
        ElseIf cx = "������" Then
            ActiveWindow.ScrollColumn = 26 '������
        ElseIf cx = "��Ʊ��" Then
            ActiveWindow.ScrollColumn = 38 '��Ʊ��
        ElseIf cx = "�ϻ���" Then
            ActiveWindow.ScrollColumn = 48 '��ִͬ�����������
        Else
            ActiveWindow.ScrollColumn = 1 '������
    End If
End Sub
