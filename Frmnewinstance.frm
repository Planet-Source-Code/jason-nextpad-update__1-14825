VERSION 5.00
Begin VB.Form Frmnewinstance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Instance"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "Frmnewinstance.frx":0000
   LinkTopic       =   "frmnewinstance"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Launch"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.OptionButton optmaximized 
      Caption         =   "M&aximized"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.OptionButton optminimized 
      Caption         =   "&Minimized"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.OptionButton optnormal 
      Caption         =   "&Normal"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select how you would like too launch the new Instance and press The ""Launch"" Button."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Frmnewinstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If optnormal.Value = True Then
ShellNewNextPad (vbNormalFocus)
Unload Me
End If

If optminimized.Value = True Then
ShellNewNextPad (vbMinimizedFocus)
Unload Me
End If

If optmaximized.Value = True Then
ShellNewNextPad (vbMaximizedFocus)
Unload Me
End If


End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If optminimized.Value = True Then
Command1.ToolTipText = "Launch " & Chr(34) & "Minimized" & Chr(34)
End If

If optmaximized.Value = True Then
Command1.ToolTipText = "Launch " & Chr(34) & "Maximized" & Chr(34)
End If

If optnormal.Value = True Then
Command1.ToolTipText = "Launch " & Chr(34) & "Normal" & Chr(34)
End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

