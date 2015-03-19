VERSION 5.00
Begin VB.Form CheckEnd 
   Caption         =   "退出确认"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "CheckEnd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "请联系监考人员，输入退出确认密码"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "CheckEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "endKJ" Then
    End
Else
    MsgBox "密码输入错误。", vbOKOnly, "提示"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
