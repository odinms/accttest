Attribute VB_Name = "Function"
Public t_option() As String '用于存储ABCD选项
Public present_id As Integer '当前题号
Public total_id As Integer '总题目数量
Public user_answer() As String '用于存储，用户的答案
Public user_point As Integer '用于存储，用户总得分
Public single_point As Single '用于存储，单选题型分数
Public muti_point As Single '用于存储，多选题型分数
Public tf_point As Single '用于存储，判断题型分数
Public CountAnswer As Integer '用于存储，已答题数目

'随机抽题
Sub getView(viewtime As Integer)
Test_form.Adodc1.RecordSource = "SELECT * From view_test"
'随机抽题开始
Dim i As Integer
i = 1
While (i <= viewtime)
    i = i + 1
    Test_form.Adodc1.Refresh
Wend
'结束
End Sub

'初始设置各题型分数
Sub setTpoint()
Test_form.Adodc1.RecordSource = "SELECT * From ks_score"
Test_form.Adodc1.Refresh
single_point = Test_form.Adodc1.Recordset.Fields(1).Value
muti_point = Test_form.Adodc1.Recordset.Fields(2).Value
tf_point = Test_form.Adodc1.Recordset.Fields(3).Value
End Sub

Sub clearBox() '清空控件内容
Test_form.Frame1.Caption = ""
Test_form.Label1.Caption = ""
For i = 0 To 4
    Test_form.Check1(i).Visible = False
    Test_form.Option1(i).Visible = False
    Test_form.Check1(i).Value = 0
    Test_form.Option1(i).Value = False
Next
End Sub

'获取题目，题型，选项的内容
Sub getText()
'以下是获取，并附值
Test_form.Label1.Caption = Test_form.DataGrid1.Columns(0).Value
Test_form.Frame1.Caption = Test_form.DataGrid1.Columns(2).Value
t_option() = Split(Test_form.DataGrid1.Columns(1).Value, "|")
Test_form.Score_txt.Text = user_point
'根据题型判断使用不同类型的选项框
Dim i As Integer
Select Case Test_form.Frame1.Caption
Case "多选题"
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
Case "判断题", "单选题"
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
'结束
'以下是更新题号，以及已答题数
Test_form.Text1.Text = "第 " & present_id & " 题"
Test_form.Label2.Caption = "当前答题数：" & vbLf & CountAnswer & "/" & total_id
'结束
End Sub

 '点击上/下一题时，根据用户答案，设置选项被选定
Sub setOptionValue()
Dim useranswer As String
useranswer = getUserAnswer(present_id - 1)
If useranswer <> "" Then '若该题用户已经作答，开始设置选项禁用
    Test_form.Ensure_B.Enabled = False '已做过的题目，点击上/下一题切换回后，确认键为禁用状态（即使，在启用状态也做了防二次修改保护。禁用仅为界面逻辑更完整）
    '拆分ABCD
    Dim split_answer(3) As String
    For j = 0 To UBound(split_answer)
        split_answer(j) = Mid(useranswer, j + 1, 1)
    Next j
    '结束
    '根据用户答案，设置选项被选定
    Select Case Test_form.DataGrid1.Columns(2).Value '获取当前题目的题型
    Case "多选题"
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
    Case "单选题", "判断题"
        Select Case useranswer
            Case "A", "对"
                Test_form.Option1(0).Value = True
                Test_form.Option1(0).Enabled = False
            Case "B", "错"
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
Else '若用户没有作答，判断是否需要恢复确认键禁用
    If Test_form.Timer1.Enabled Then
        Test_form.Ensure_B.Enabled = True
    End If
End If
End Sub

'获取固定题号用户的答案
Function getUserAnswer(i As Integer) As String '目前该方法其实是多余的，为以后功能扩展预留
getUserAnswer = user_answer(i)
End Function

'将选项转换为ABCD格式的答案，并附值给user_answer()记录答案
Sub setUserAnswer()
'开始
Dim setUserAnswer As String
Select Case Test_form.Frame1.Caption
Case "多选题"
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
Case "单选题"
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
Case "判断题"
    Select Case True
        Case Test_form.Option1(0).Value
            setUserAnswer = "对"
        Case Test_form.Option1(1).Value
            setUserAnswer = "错"
    End Select
    user_answer(present_id - 1) = setUserAnswer
End Select
'结束
'当当前题目用户答案不为空时，累计已答题目数
If user_answer(present_id - 1) <> "" Then
    CountAnswer = CountAnswer + 1
End If
End Sub

'判断两个字符串中包含的元素完全相同，用于判断用户结果是否和标准答案相同
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

'获取当前题目的正确答案
Function getAnswer() As String
getAnswer = Test_form.DataGrid1.Columns(3).Value
End Function

'得分计算，计分成功 返回TRUE，失败返回FLASE
Function AddPoint() As Boolean
Dim useranswer As String '当前题目用户历史答案
Dim t_point As Single '当前题目得分
Dim t_false As Single '错误答案分数
useranswer = getUserAnswer(present_id - 1)
If useranswer = "" Then '为空说明之前没有录入过答案（防二次修改保护）
    Call setUserAnswer  '记录用户答案
    '开始计分
    Dim sys_answer As String '当前题目正确答案
    Select Case Test_form.Frame1.Caption
        Case "单选题"
            t_point = single_point
            t_false = 0
        Case "多选题"
            t_point = muti_point
            t_false = 0
        Case "判断题"
            t_point = tf_point
            t_false = tf_point / 2
    End Select
    sys_answer = getAnswer() '获取当前题目正确答案
    useranswer = getUserAnswer(present_id - 1) '因记录用户答案后，当前题目用户答案发生变化，需要重新获取
    If IsEquals(sys_answer, useranswer) Then '答案正确，则计分
        user_point = user_point + t_point
        AddPoint = True
    Else '答案错误，不计分
        If useranswer <> "" Then '判断错误答案，是否为空（空答案也可能为错误答案）
            Call showAnswer(1) '以MSGBOX方式，显示正确答案
            user_point = user_point - t_false
            AddPoint = True
        Else '用户没有选择任何选项，产生为空的错误答案
            MsgBox "请先选择答案", vbOKOnly, "提示 " '提醒用户，选择答案
            AddPoint = False
        End If
    End If
    '计分结束
End If

End Function


'显示正确答案
Sub showAnswer(i As Integer) '变量i为需要的展现类型 0 在FRAME框中展现， 1 在MSGBOX中展现
Select Case i
    Case 0
        Test_form.Label6.Caption = getAnswer()
    Case 1
        MsgBox "正确答案：" & getAnswer(), vbOKOnly, "回答错误"
End Select
End Sub

Sub stopTest()
Dim check As Integer
check = MsgBox("是否确认交卷?" & vbLf & "交卷后请不要关闭窗体。向监考人员示意并登记成绩。" & vbLf & "当前答题数：" & CountAnswer & "/" & total_id & vbLf & "目前得分：" & user_point, vbYesNo, "提示")
If check = vbYes Then
    Test_form.Timer1.Enabled = False
    Test_form.Ensure_B.Enabled = False
    Test_form.Submmit_B.Enabled = False
    Test_form.Frame3.Visible = True
    Call showAnswer(0)
    MsgBox "交卷成功，考试结束。" & vbLf & "点击 上/下一题 可回顾所有题目的结果。", vbOKOnly, "提示"
End If
End Sub

Sub AddButton()
Dim avg_wid As Integer
avg_wid = (2360 - 2 * 5) / 6 '单位宽度
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
