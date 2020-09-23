VERSION 5.00
Begin VB.Form frmeditclipboard 
   Caption         =   "Edit ClipBoard Text"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "Frmeditclipboard.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   4
      Top             =   3330
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "&Update Window From Clipboard"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Updates Window From Clipboard"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   4575
         Begin VB.CommandButton Command4 
            Caption         =   "C&lose"
            Height          =   375
            Left            =   2880
            TabIndex        =   6
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Save Too Clipboard"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Writes Currently edited text too clipboard"
            Top             =   720
            Width           =   2655
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Clear Clipboard"
            Height          =   375
            Left            =   2880
            TabIndex        =   3
            ToolTipText     =   "Clears Clipboard Text"
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmeditclipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      On Error GoTo clipboarderror: ' if an error occurs vb will jump too that line
       
       Text1.Text = Clipboard.GetText ' this will update the Text Box named text1 on current Clipboard text


clipboarderror: ' Error Control Starts Here

   If Err.Number <> 0 Then ' IF an Error' Umber is anything Above or at zero then .. . ...
     ' display error message too user
     MsgBox "An error Has been encountered while accessing the Clipboard" _
     & Chr(13) & Err.Description, vbCritical, "Error , Clipboard"
    Exit Sub ' exit the sub immediately
   End If

End Sub


Private Sub Command2_Click()

    Clipboard.SetText ("") ' This Will clear the clipboard
    ' and Will set it's Text Too NOTHING ("")

End Sub

Private Sub Command3_Click()

 Clipboard.SetText Text1.Text() ' This Will Set the Clipboards text too
 ' The Text that is in text1

End Sub

Private Sub Command4_Click()

 Unload Me ' unload the form

End Sub

Private Sub Form_Load()

   On Error GoTo clipboarderror: ' if an error occurs vb will jump too that line
      
      frmeditclipboard.Text1.Text = Clipboard.GetText ' Set the Text1's Text Too The Text That We
      ' Get from the Clipboard
      
clipboarderror: ' Error Control starts here
      If Err.Number <> 0 Then ' If an errors Number is anything above or at zero then .........
        ' Display the Error Message too the user
        MsgBox "An error Has been encountered while accessing the Clipboard" _
        & Chr(13) & Err.Description, vbCritical, "Error , Clipboard"
         Exit Sub ' exit the sub immediately
        End If
End Sub

Private Sub Form_Resize()
  
   On Error GoTo Errresize: ' if an error occurs while resizing the form Jump too that line
      
      Text1.Width = Me.ScaleWidth - (Text1.Left * 2) 'Set Text1's Width too the FOrms Width Multiplying that by the Text1's Left Property
       Text1.Height = Me.Height - 1600 ' Set Text1's Height too the Forms Height Minus The Frames Hieght
        Frame1.Width = Me.Width - 150 ' Set the Frames Width too the form's width minus the left Opposite of the frame itself

Errresize:
       Exit Sub ' Exit the sub Immediately
End Sub

