VERSION 5.00
Begin VB.Form Frmbugz 
   Caption         =   "Version History "
   ClientHeight    =   4335
   ClientLeft      =   2235
   ClientTop       =   2985
   ClientWidth     =   7455
   Icon            =   "Frmbugz.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Frmbugz.frx":27A2
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Frmbugz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

On Error GoTo Resizeerr: ' if an error occurs Vb will Jump too that line

    Text1.Height = Frmbugz.ScaleHeight 'Set the Text1's Height too the height of form1's scale height property
    Text1.Width = Frmbugz.ScaleWidth ' set the Text1's Height too the width of  form1's scale width property

Resizeerr:
    Exit Sub ' Exit the sub immediately

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set Frmbugz = Nothing ' Release the memory that This form had Held


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
  ' Let the user Know that This control cannot Be edited
  Beep

End Sub
