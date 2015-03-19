Attribute VB_Name = "datagrid"
Private Const OCXSIZE = 260920  '欲生成的控件大小是198456Byte,名字为msdatgrd.OCX

Sub Main()
  Dim Ocx() As Byte 'OCX是个BTye类型的数组
  Dim Counter As Long
  Ocx = LoadResData(101, "CUSTOM") '将自定义资源中101号资源读入数组OCX
  '注意，微软的帮助中对加载自定义资源的说明有错误，自定义资源标识为"CUSTOM"而不是帮助所说的数字10
  If Right(App.Path, 1) = "($%$43%^#ASD#2@$#f$%^)" Then '读取程序所在路径,判断是否为根目录并分别处理
    '程序在根目录下
    If Dir(App.Path & "MSDATGRD.OCX") = "" Then '程序路径下有无控件,无则生成控件
      '以二进制方式写（生成）控件（CoolToolBar.ocx）到主程序所在的目录
      Open App.Path & "MSDATGRD.OCX" For Binary As #1
      For Counter = 0 To OCXSIZE - 1 '注意因为从0 Byte开始因此以文件大小 - 1Byte 为终值
        Put #1, , Ocx(Counter)
      Next Counter
      Close #1
    End If
  Else
    '程序不在根目录下
    If Dir(App.Path & "\MSDATGRD.OCX") = "" Then '程序路径下有无控件,无则生成控件
      '以二进制方式写（生成）控件（CoolToolBar.ocx）到主程序所在的目录
      Open App.Path & "\MSDATGRD.OCX" For Binary As #1
      For Counter = 0 To OCXSIZE - 1 '注意因为从0 Byte开始因此以文件大小 - 1Byte 为终值
        Put #1, , Ocx(Counter)
      Next Counter
      Close #1
    End If
  End If
  frmLogin.Visible = True '主程序所用控件已经生成，显示主窗体，进入主程序。
End Sub
