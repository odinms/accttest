Attribute VB_Name = "datagrid"
Private Const OCXSIZE = 260920  '�����ɵĿؼ���С��198456Byte,����Ϊmsdatgrd.OCX

Sub Main()
  Dim Ocx() As Byte 'OCX�Ǹ�BTye���͵�����
  Dim Counter As Long
  Ocx = LoadResData(101, "CUSTOM") '���Զ�����Դ��101����Դ��������OCX
  'ע�⣬΢��İ����жԼ����Զ�����Դ��˵���д����Զ�����Դ��ʶΪ"CUSTOM"�����ǰ�����˵������10
  If Right(App.Path, 1) = "($%$43%^#ASD#2@$#f$%^)" Then '��ȡ��������·��,�ж��Ƿ�Ϊ��Ŀ¼���ֱ���
    '�����ڸ�Ŀ¼��
    If Dir(App.Path & "MSDATGRD.OCX") = "" Then '����·�������޿ؼ�,�������ɿؼ�
      '�Զ����Ʒ�ʽд�����ɣ��ؼ���CoolToolBar.ocx�������������ڵ�Ŀ¼
      Open App.Path & "MSDATGRD.OCX" For Binary As #1
      For Counter = 0 To OCXSIZE - 1 'ע����Ϊ��0 Byte��ʼ������ļ���С - 1Byte Ϊ��ֵ
        Put #1, , Ocx(Counter)
      Next Counter
      Close #1
    End If
  Else
    '�����ڸ�Ŀ¼��
    If Dir(App.Path & "\MSDATGRD.OCX") = "" Then '����·�������޿ؼ�,�������ɿؼ�
      '�Զ����Ʒ�ʽд�����ɣ��ؼ���CoolToolBar.ocx�������������ڵ�Ŀ¼
      Open App.Path & "\MSDATGRD.OCX" For Binary As #1
      For Counter = 0 To OCXSIZE - 1 'ע����Ϊ��0 Byte��ʼ������ļ���С - 1Byte Ϊ��ֵ
        Put #1, , Ocx(Counter)
      Next Counter
      Close #1
    End If
  End If
  frmLogin.Visible = True '���������ÿؼ��Ѿ����ɣ���ʾ�����壬����������
End Sub
