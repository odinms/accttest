VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Test_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "西南财经大学天府学院蒲公英大赛考试系统"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17220
   DrawMode        =   9  'Not Mask Pen
   DrawStyle       =   5  'Transparent
   Icon            =   "Test_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   393.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   861
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Start_tips 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5775
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Test_form.frx":1CCA
      Top             =   480
      Width           =   11895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "自由跳题框"
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   3480
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "E:"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   3480
      TabIndex        =   41
      Top             =   4200
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9240
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "正确答案"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   11895
      Begin VB.Label Label6 
         Caption         =   "答案"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   11415
      End
   End
   Begin VB.CommandButton Ensure_B 
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      MaskColor       =   &H80000010&
      TabIndex        =   18
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   11280
      TabIndex        =   17
      Text            =   "用于测试时，观看当前题目正确答案"
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10560
      Top             =   6600
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3000
      TabIndex        =   14
      Top             =   6240
      Width           =   11895
      Begin MSForms.Label Label14 
         Height          =   375
         Left            =   480
         TabIndex        =   38
         Top             =   360
         Width           =   6735
         Caption         =   "注意：考试过程中，有任何问题，请示意监考人员！！！"
         Size            =   "11880;661"
         FontName        =   "微软雅黑"
         FontHeight      =   240
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "00小时00分钟00秒"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   9720
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "当前用时："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "题号"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton ButtonPrevious 
      Caption         =   "上一题"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15120
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton ButtonNext 
      Caption         =   "下一题"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15120
      TabIndex        =   7
      Top             =   5040
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Test_form.frx":1DF3
      Height          =   3135
      Left            =   3000
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   15120
      Top             =   7800
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   1
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   6735
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   2655
      Begin VB.TextBox Name_txt 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "她有一个长名字"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox No_txt 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "41XXXXXX"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton Submmit_B 
         Caption         =   "交卷"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   27
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   5400
         Width           =   630
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   4800
         Width           =   630
      End
      Begin MSForms.CommandButton Command1 
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   345
         ForeColor       =   16777215
         BackColor       =   16576
         Caption         =   "1"
         Size            =   "609;609"
         FontName        =   "宋体"
         FontHeight      =   180
         FontCharSet     =   134
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "题型"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   5775
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      Begin VB.OptionButton Option1 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   6480
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   6480
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   2880
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "A:"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "B:"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   6480
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "D:"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   6480
         TabIndex        =   12
         Top             =   2880
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "C:"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   480
         TabIndex        =   11
         Top             =   2880
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "题目"
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2580
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11640
      End
   End
   Begin MSForms.Label Label13 
      Height          =   345
      Left            =   840
      TabIndex        =   37
      Top             =   7350
      Width           =   3825
      VariousPropertyBits=   276824091
      Caption         =   "版本号：V2.1.1 ABCDE选项 单机版"
      Size            =   "6747;609"
      FontName        =   "微软雅黑"
      FontHeight      =   210
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   375
      Left            =   6600
      TabIndex        =   36
      Top             =   7350
      Width           =   5055
      Caption         =   "仅供西南财经大学天府学院会计技能大赛比赛使用"
      Size            =   "8916;661"
      FontName        =   "微软雅黑"
      FontHeight      =   210
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label11 
      Height          =   300
      Left            =   14520
      TabIndex        =   35
      Top             =   7350
      Width           =   1845
      VariousPropertyBits=   276824091
      Caption         =   "版权所有：Zhouyu"
      Size            =   "3254;529"
      FontName        =   "微软雅黑"
      FontHeight      =   210
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Copyright_label 
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   7275
      Width           =   16695
      Size            =   "29448;873"
      BorderStyle     =   1
      FontName        =   "微软雅黑"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Score_txt 
      Height          =   750
      Left            =   15360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   645
      Width           =   1185
      VariousPropertyBits=   1015040019
      ForeColor       =   -2147483634
      Size            =   "2090;1323"
      Value           =   "100"
      SpecialEffect   =   0
      FontName        =   "微软雅黑"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "考试限时：60分钟"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15120
      TabIndex        =   23
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目前得分："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   14640
      TabIndex        =   22
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "当前答题数："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15000
      TabIndex        =   21
      Top             =   6120
      Width           =   2055
   End
   Begin MSForms.Label Label10 
      Height          =   1095
      Left            =   15120
      TabIndex        =   29
      Top             =   480
      Width           =   1815
      BackColor       =   255
      Size            =   "3201;1931"
      BorderColor     =   0
      BorderStyle     =   1
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "Test_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lefttime As String
Dim mm As Integer, ms As Integer, mh As Integer


Private Sub Command1_Click(Index As Integer)
Test_form.Adodc1.Recordset.AbsolutePosition = Index + 1
present_id = Index + 1
Call clearBox
Call getText
Call setOptionValue
Call showAnswer(0)
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\tk;Jet OLEDB:Database Password=1111"
Call setTpoint
Randomize
Call getView(Int(Rnd * 99)) '获取随机视图
total_id = Test_form.Adodc1.Recordset.RecordCount '获取总共抽取的题数
'初始化用户答案
ReDim user_answer(total_id - 1)
For i = 0 To UBound(user_answer)
    user_answer(i) = ""
Next
'结束
present_id = 1 '初始化题号
Call clearBox
Call getText
AddButton
For i = 0 To 3 '取消所有选项焦点
    Test_form.Check1(i).Value = 0
    Test_form.Option1(i).Value = False
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
If MsgBox("确定要退出吗?", vbYesNo) = vbYes Then
    Cancel = 1
    CheckEnd.Show 1, Me
Else
    Cancel = 1
End If
End Sub

Private Sub Start_tips_Click()
'result = MsgBox("是否认真阅读温馨提示", vbYesNo, "提示")
'If result = vbYes Then
'    Start_tips.Visible = False
'    Start_tips.Enabled = False
'    Timer1.Enabled = True
'End If
CheckStart.Show 1, Me
End Sub

Private Sub ButtonNext_Click()
Test_form.Adodc1.Recordset.MoveNext
If Test_form.Adodc1.Recordset.EOF Then
    Test_form.Adodc1.Recordset.MoveFirst
    present_id = 1
Else
    present_id = present_id + 1
End If
Call clearBox
Call getText
Call setOptionValue
Call showAnswer(0)
'Text4.Text = getAnswer()
End Sub

Private Sub ButtonPrevious_Click()
Test_form.Adodc1.Recordset.MovePrevious
If Test_form.Adodc1.Recordset.BOF Then
    Test_form.Adodc1.Recordset.MoveLast
    present_id = total_id
Else
    present_id = present_id - 1
End If
Call clearBox
Call getText
Call setOptionValue
Call showAnswer(0)
End Sub

Private Sub Ensure_B_Click()
Dim check As Boolean
check = AddPoint()
If check = True Then '计算得分,计算得分时会自动记录用户答案
    Test_form.Adodc1.Recordset.MoveNext
    If Test_form.Adodc1.Recordset.EOF Then
        Test_form.Adodc1.Recordset.MoveFirst
        present_id = 1
    Else
        present_id = present_id + 1
    End If
    Call clearBox
    
    Call getText
    Call setOptionValue
    'Text4.Text = getAnswer
End If
End Sub

Private Sub Submmit_B_Click()
Call stopTest
End Sub

Private Sub Timer1_Timer()
ms = ms + 1
If ms = 60 Then ms = 0
If ms = 0 Then mm = mm + 1
If mm = 60 Then
    mm = 0
    mh = mh + 1
End If
If mh = 1 Then '时间到达一小时，停止计时。考试结束
    Timer1.Enabled = False
    '考试结束
    Call stopTest
    Call showAnswer(0)
    MsgBox "你的得分为：" & user_point, vbOKOnly, "成绩"
End If
lefttime = mh & "小时" & mm & "分钟" & ms & "秒"
Label4.Caption = lefttime
End Sub
