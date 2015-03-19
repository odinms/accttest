VERSION 5.00
Begin VB.Form CheckStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "温馨提示"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   Icon            =   "CheckStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "否"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "是(10)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否认真阅读温馨提示？"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "CheckStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Counter As Integer

Private Sub Command1_Click()
Test_form.Start_tips.Visible = False
Test_form.Start_tips.Enabled = False
Test_form.Timer1.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Counter = 10
End Sub

Private Sub Timer1_Timer()
If Counter = 0 Then
    Command1.Enabled = True
    Command1.Caption = "是"
    Timer1.Enabled = False
Else
    Counter = Counter - 1
    Command1.Caption = "是(" & Counter & ")"
End If
End Sub
