Attribute VB_Name = "Function"
Public t_option() As String '���ڴ洢ABCDѡ��
Public present_id As Integer '��ǰ���
Public total_id As Integer '����Ŀ����
Public user_answer() As String '���ڴ洢���û��Ĵ�
Public user_point As Integer '���ڴ洢���û��ܵ÷�
Public single_point As Single '���ڴ洢����ѡ���ͷ���
Public muti_point As Single '���ڴ洢����ѡ���ͷ���
Public tf_point As Single '���ڴ洢���ж����ͷ���
Public CountAnswer As Integer '���ڴ洢���Ѵ�����Ŀ

'�������
Sub getView(viewtime As Integer)
Test_form.Adodc1.RecordSource = "SELECT * From view_test"
'������⿪ʼ
Dim i As Integer
i = 1
While (i <= viewtime)
    i = i + 1
    Test_form.Adodc1.Refresh
Wend
'����
End Sub

'��ʼ���ø����ͷ���
Sub setTpoint()
Test_form.Adodc1.RecordSource = "SELECT * From ks_score"
Test_form.Adodc1.Refresh
single_point = Test_form.Adodc1.Recordset.Fields(1).Value
muti_point = Test_form.Adodc1.Recordset.Fields(2).Value
tf_point = Test_form.Adodc1.Recordset.Fields(3).Value
End Sub

Sub clearBox() '��տؼ�����
Test_form.Frame1.Caption = ""
Test_form.Label1.Caption = ""
For i = 0 To 4
    Test_form.Check1(i).Visible = False
    Test_form.Option1(i).Visible = False
    Test_form.Check1(i).Value = 0
    Test_form.Option1(i).Value = False
Next
End Sub

'��ȡ��Ŀ�����ͣ�ѡ�������
Sub getText()
'�����ǻ�ȡ������ֵ
Test_form.Label1.Caption = Test_form.DataGrid1.Columns(0).Value
Test_form.Frame1.Caption = Test_form.DataGrid1.Columns(2).Value
t_option() = Split(Test_form.DataGrid1.Columns(1).Value, "|")
Test_form.Score_txt.Text = user_point
'���������ж�ʹ�ò�ͬ���͵�ѡ���
Dim i As Integer
Select Case Test_form.Frame1.Caption
Case "��ѡ��"
    For i = 0 To UBound(t_option)
        Test_form.Check1(i).Caption = t_option(i)
        Test_form.Check1(i).Visible = True
        Test_form.Check1(i).Enabled = True
        Test_form.Option1(i).Visible = False
    Next
    If i < 4 Then
        For a = 4 To i Step -1
            Test_form.Check1(a).Visible = False
            Test_form.Option1(a).Visible = False
        Next
    End If
Case "�ж���", "��ѡ��"
    For i = 0 To UBound(t_option)
        Test_form.Option1(i).Caption = t_option(i)
        Test_form.Option1(i).Visible = True
        Test_form.Option1(i).Enabled = True
         Test_form.Check1(i).Visible = False
    Next
    If i < 4 Then
        For a = 4 To i Step -1
            Test_form.Option1(a).Visible = False
            Test_form.Check1(a).Visible = False
        Next
    End If
End Select
'����
'�����Ǹ�����ţ��Լ��Ѵ�����
Test_form.Text1.Text = "�� " & present_id & " ��"
Test_form.Label2.Caption = "��ǰ��������" & vbLf & CountAnswer & "/" & total_id
'����
End Sub

 '�����/��һ��ʱ�������û��𰸣�����ѡ�ѡ��
Sub setOptionValue()
Dim useranswer As String
useranswer = getUserAnswer(present_id - 1)
If useranswer <> "" Then '�������û��Ѿ����𣬿�ʼ����ѡ�����
    Test_form.Ensure_B.Enabled = False '����������Ŀ�������/��һ���л��غ�ȷ�ϼ�Ϊ����״̬����ʹ��������״̬Ҳ���˷������޸ı��������ý�Ϊ�����߼���������
    '���ABCD
    Dim split_answer(3) As String
    For j = 0 To UBound(split_answer)
        split_answer(j) = Mid(useranswer, j + 1, 1)
    Next j
    '����
    '�����û��𰸣�����ѡ�ѡ��
    Select Case Test_form.DataGrid1.Columns(2).Value '��ȡ��ǰ��Ŀ������
    Case "��ѡ��"
        If split_answer(0) = "A" Then
            Test_form.Check1(0).Value = 1
            Test_form.Check1(0).Enabled = False
        End If
    For i = 0 To UBound(split_answer)
        If split_answer(i) = "B" And i <= 1 Then
            Test_form.Check1(i).Value = 1
            Test_form.Check1(i).Enabled = False
        End If
        If split_answer(i) = "C" And i <= 2 Then
            Test_form.Check1(i).Value = 1
            Test_form.Check1(i).Enabled = False
        End If
        If split_answer(i) = "D" And i <= 3 Then
            Test_form.Check1(i).Value = 1
            Test_form.Check1(i).Enabled = False
        End If
        If split_answer(i) = "E" Then
            Test_form.Check1(i).Value = 1
            Test_form.Check1(i).Enabled = False
        End If
    Next
    Case "��ѡ��", "�ж���"
        Select Case useranswer
            Case "A", "��"
                Test_form.Option1(0).Value = True
                Test_form.Option1(0).Enabled = False
            Case "B", "��"
                Test_form.Option1(1).Value = True
                Test_form.Option1(1).Enabled = False
            Case "C"
                Test_form.Option1(2).Value = True
                Test_form.Option1(2).Enabled = False
            Case "D"
                Test_form.Option1(3).Value = True
                Test_form.Option1(3).Enabled = False
            Case "E"
                Test_form.Option1(4).Value = True
                Test_form.Option1(4).Enabled = False
        End Select
    End Select
Else '���û�û�������ж��Ƿ���Ҫ�ָ�ȷ�ϼ�����
    If Test_form.Timer1.Enabled Then
        Test_form.Ensure_B.Enabled = True
    End If
End If
End Sub

'��ȡ�̶�����û��Ĵ�
Function getUserAnswer(i As Integer) As String 'Ŀǰ�÷�����ʵ�Ƕ���ģ�Ϊ�Ժ�����չԤ��
getUserAnswer = user_answer(i)
End Function

'��ѡ��ת��ΪABCD��ʽ�Ĵ𰸣�����ֵ��user_answer()��¼��
Sub setUserAnswer()
'��ʼ
Dim setUserAnswer As String
Select Case Test_form.Frame1.Caption
Case "��ѡ��"
    If Test_form.Check1(0).Value = 1 Then
        setUserAnswer = "A"
        user_answer(present_id - 1) = user_answer(present_id - 1) + setUserAnswer
    End If
    If Test_form.Check1(1).Value = 1 Then
        setUserAnswer = "B"
        user_answer(present_id - 1) = user_answer(present_id - 1) + setUserAnswer
    End If
    If Test_form.Check1(2).Value = 1 Then
        setUserAnswer = "C"
        user_answer(present_id - 1) = user_answer(present_id - 1) + setUserAnswer
    End If
    If Test_form.Check1(3).Value = 1 Then
        setUserAnswer = "D"
        user_answer(present_id - 1) = user_answer(present_id - 1) + setUserAnswer
    End If
    If Test_form.Check1(4).Value = 1 Then
        setUserAnswer = "E"
        user_answer(present_id - 1) = user_answer(present_id - 1) + setUserAnswer
    End If
Case "��ѡ��"
    Select Case True
        Case Test_form.Option1(0).Value
            setUserAnswer = "A"
        Case Test_form.Option1(1).Value
            setUserAnswer = "B"
        Case Test_form.Option1(2).Value
            setUserAnswer = "C"
        Case Test_form.Option1(3).Value
            setUserAnswer = "D"
        Case Test_form.Option1(4).Value
            setUserAnswer = "E"
    End Select
    user_answer(present_id - 1) = setUserAnswer
Case "�ж���"
    Select Case True
        Case Test_form.Option1(0).Value
            setUserAnswer = "��"
        Case Test_form.Option1(1).Value
            setUserAnswer = "��"
    End Select
    user_answer(present_id - 1) = setUserAnswer
End Select
'����
'����ǰ��Ŀ�û��𰸲�Ϊ��ʱ���ۼ��Ѵ���Ŀ��
If user_answer(present_id - 1) <> "" Then
    CountAnswer = CountAnswer + 1
End If
End Sub

'�ж������ַ����а�����Ԫ����ȫ��ͬ�������ж��û�����Ƿ�ͱ�׼����ͬ
Function IsEquals(str1 As String, str2 As String) As Boolean
If Len(str1) <> Len(str2) Then
   IsEquals = False
   Exit Function
End If
Dim i As Integer
Dim Count As Integer
For i = 1 To Len(str1)
    For j = 1 To Len(str2)
        If Mid(str1, i, 1) = Mid(str2, j, 1) Then
             Count = Count + 1
        End If
    Next j
Next i
If Count = Len(str1) Then
    IsEquals = True
End If
End Function

'��ȡ��ǰ��Ŀ����ȷ��
Function getAnswer() As String
getAnswer = Test_form.DataGrid1.Columns(3).Value
End Function

'�÷ּ��㣬�Ʒֳɹ� ����TRUE��ʧ�ܷ���FLASE
Function AddPoint() As Boolean
Dim useranswer As String '��ǰ��Ŀ�û���ʷ��
Dim t_point As Single '��ǰ��Ŀ�÷�
Dim t_false As Single '����𰸷���
useranswer = getUserAnswer(present_id - 1)
If useranswer = "" Then 'Ϊ��˵��֮ǰû��¼����𰸣��������޸ı�����
    Call setUserAnswer  '��¼�û���
    '��ʼ�Ʒ�
    Dim sys_answer As String '��ǰ��Ŀ��ȷ��
    Select Case Test_form.Frame1.Caption
        Case "��ѡ��"
            t_point = single_point
            t_false = 0
        Case "��ѡ��"
            t_point = muti_point
            t_false = 0
        Case "�ж���"
            t_point = tf_point
            t_false = tf_point / 2
    End Select
    sys_answer = getAnswer() '��ȡ��ǰ��Ŀ��ȷ��
    useranswer = getUserAnswer(present_id - 1) '���¼�û��𰸺󣬵�ǰ��Ŀ�û��𰸷����仯����Ҫ���»�ȡ
    If IsEquals(sys_answer, useranswer) Then '����ȷ����Ʒ�
        user_point = user_point + t_point
        AddPoint = True
    Else '�𰸴��󣬲��Ʒ�
        If useranswer <> "" Then '�жϴ���𰸣��Ƿ�Ϊ�գ��մ�Ҳ����Ϊ����𰸣�
            Call showAnswer(1) '��MSGBOX��ʽ����ʾ��ȷ��
            user_point = user_point - t_false
            AddPoint = True
        Else '�û�û��ѡ���κ�ѡ�����Ϊ�յĴ����
            MsgBox "����ѡ���", vbOKOnly, "��ʾ " '�����û���ѡ���
            AddPoint = False
        End If
    End If
    '�Ʒֽ���
End If

End Function


'��ʾ��ȷ��
Sub showAnswer(i As Integer) '����iΪ��Ҫ��չ������ 0 ��FRAME����չ�֣� 1 ��MSGBOX��չ��
Select Case i
    Case 0
        Test_form.Label6.Caption = getAnswer()
    Case 1
        MsgBox "��ȷ�𰸣�" & getAnswer(), vbOKOnly, "�ش����"
End Select
End Sub

Sub stopTest()
Dim check As Integer
check = MsgBox("�Ƿ�ȷ�Ͻ���?" & vbLf & "������벻Ҫ�رմ��塣��࿼��Աʾ�Ⲣ�Ǽǳɼ���" & vbLf & "��ǰ��������" & CountAnswer & "/" & total_id & vbLf & "Ŀǰ�÷֣�" & user_point, vbYesNo, "��ʾ")
If check = vbYes Then
    Test_form.Timer1.Enabled = False
    Test_form.Ensure_B.Enabled = False
    Test_form.Submmit_B.Enabled = False
    Test_form.Frame3.Visible = True
    Call showAnswer(0)
    MsgBox "����ɹ������Խ�����" & vbLf & "��� ��/��һ�� �ɻع�������Ŀ�Ľ����", vbOKOnly, "��ʾ"
End If
End Sub

Sub AddButton()
Dim avg_wid As Integer
avg_wid = (2360 - 2 * 5) / 6 '��λ���
Test_form.Text2.Text = (2360 - 2 * 5) / 6
Test_form.Command1(0).Width = avg_wid
Test_form.Command1(i).Height = avg_wid
X = Test_form.Command1(0).Left + Test_form.Command1(0).Width
Y = Test_form.Command1(0).Top

For i = 1 To total_id - 1
    Load Test_form.Command1(i)
    Test_form.Command1(i).Width = avg_wid
    Test_form.Command1(i).Height = avg_wid
    Test_form.Command1(i).Move X, Y
    Test_form.Command1(i).Caption = i + 1
    Test_form.Command1(i).Visible = True
    If X <= Test_form.Command1(0).Left + 4 * Test_form.Command1(0).Width Then
        X = X + Test_form.Command1(i).Width
    Else
        X = Test_form.Command1(0).Left
        Y = Y + Test_form.Command1(i).Height
    End If
Next
End Sub
