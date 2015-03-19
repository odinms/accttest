VERSION 5.00
Begin VB.Form ShowResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "考试成绩"
   ClientHeight    =   6375
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "ShowResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Label1.Caption = "是否确认交卷。" & vbLf & "当前答题数：" & CountAnswer & "/" & total_id & vbLf & "目前得分：" & user_point
End Sub
