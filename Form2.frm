VERSION 5.00
Begin VB.Form Frmleavefullscreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Full "
   ClientHeight    =   495
   ClientLeft      =   8970
   ClientTop       =   435
   ClientWidth     =   510
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   510
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   0
      Picture         =   "Form2.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Frmleavefullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
Unload frmfullscreen
Form1.mnufullscreen.Checked = False
Set frmfullscreen = Nothing
Set Frmleavefullscreen = Nothing
Form1.WindowState = vbNormal

End Sub

Private Sub Form_Load()
'Me.Move (Screen.Width - Me.Width) \ 1, (Screen.Height - Me.Height) \ 11 + 10 - 2

End Sub


