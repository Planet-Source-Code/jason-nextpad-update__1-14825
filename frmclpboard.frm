VERSION 5.00
Begin VB.Form frmclpboard 
   Caption         =   "Clipboard text"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "frmclpboard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Clipboard Text: "
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmclpboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

On Error GoTo Clipboarderr: ' if an error occurs Vb Will jump too that line

   If Clipboard.GetText <> "" Then
    Text1.Text = Clipboard.GetText 'set the text too the text form the clipboard ( IF any )
   Else ' if there isnt then
    frmclpboard.Hide ' First Hide the form
    ' Display the message too the user
    MsgBox "There Is no Clipboard text Too display", vbExclamation, "Clipboard Text Error"
    ' Now we are done so unload the form
    Unload Me
   End If

Clipboarderr:
    If Err.Number <> 0 Then ' if an error is equal too anything other or above zero then ......
     ' Display the error message too the user
     MsgBox "An Unexpected Error has occured while accessing the Clipboard ", vbCritical, "NextPad"
     Exit Sub ' exit the sub immediately
    End If


End Sub

Private Sub Form_Resize()
     On Error GoTo Resizeerror: ' if an error occurs vb will jump too that line
       
       Text1.Height = Me.ScaleHeight - 241 ' Resize the Text Box too fit the forms Scale height Minus the Labels height
       Text1.Width = Me.ScaleWidth ' Resize the text Box too fit the forms scale width

Resizeerror:
    Exit Sub
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
' Beep too Grab the users attention
' that this TextBox CANNOT be Edited

      Beep
      MsgBox "You Cannot Edit The Clipboard Text Here. If you Want Too Edit the Clipboard Text Freely then, On the Edit Menu Click " & Chr$(34) & "Edit Clipboard Text" & Chr$(34), vbInformation, "NextPad"
End Sub
