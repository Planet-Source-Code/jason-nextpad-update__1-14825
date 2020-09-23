VERSION 5.00
Begin VB.Form Frmnoreg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NextPad Setup"
   ClientHeight    =   6435
   ClientLeft      =   2550
   ClientTop       =   3690
   ClientWidth     =   5805
   Icon            =   "frmnoreg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtHelp 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4920
      Width           =   5655
   End
   Begin VB.CommandButton cmdsetdefaultoptions 
      Caption         =   "&Set Default Options (Recommended)"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CheckBox ChckExternalEditor 
      Caption         =   "&Use External Editor Too open Files Too large For NextPad Too open."
      Height          =   255
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   $"frmnoreg.frx":000C
      Top             =   3600
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      Picture         =   "frmnoreg.frx":00B5
      ScaleHeight     =   675
      ScaleWidth      =   5745
      TabIndex        =   7
      Top             =   0
      Width           =   5805
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Save settings ,Close window."
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CheckBox Chckassociations 
      Caption         =   "NextPad should &check whether it is the default text viewer"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "This enables NextPad too notify you when NextPad is not the default text viewer."
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CheckBox Chckwordwrap 
      Caption         =   "Use &Word - Wrap"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "This Enables NextPad too wrap text too the text box."
      Top             =   2880
      Width           =   4455
   End
   Begin VB.CheckBox Chcktoolbar 
      Caption         =   "Always show &Toolbar"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "This enables NextPad too Show the toolbar every time you start NextPad."
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   5655
      Begin VB.CheckBox ChckAskIfToobig 
         Caption         =   "Ask Too &launch External Editor if File Is too Large too Open"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   5295
      End
      Begin VB.CheckBox chckallowassociation 
         Caption         =   "&Allow NextPad too be the default Text Viewer (*.TXT) on this Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   5415
      End
   End
   Begin VB.Label Label2 
      Caption         =   "When Youre Done Choosing Options Press  The ""Done"" Button.  ."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmnoreg.frx":77EF
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "Frmnoreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chckallowassociation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option Will Allow You too have NextPad be The Default Viewer For (*.TXT) Files ( also Highly Recommended as Well )"

End Sub

Private Sub ChckAskIfToobig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option When Enabled Tells NextPad too Launch The External Editor Without Notifying the user about it , This in turn May make Thigs a bit easier but can Confuse some. NOTE : In order for this option too have any affect you must have on " & Chr$(34) & " Execute External Editor When File Is too large too open." & Chr$(34)

End Sub


Private Sub Chckassociations_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option if Enabled would Allow NextPad too prompt the user if NextPad isnt The Default Viewer For (*.TXT) Files On the Computer it Is Running on , Not Recommended it is Highly Annoying !!!"

End Sub


Private Sub ChckExternalEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option is highly Recommended !! If allowed, if a File is too large for NextPad too open , It will prompt the user too open the file with a More Powerful External Editor."

End Sub



Private Sub Chcktoolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option If Enabled Always Shows The Toolbar ( You Can Toggle The Toolbar ON and OFF in the View Menu ) if Disabled The Toolbar Isnt Always Shown"

End Sub

Private Sub Chckwordwrap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "Enable This Setting Too Wrap the Text in NextPads Window too the window , Disable for it Not Too be Wrapped Too the window."

End Sub

'***********************************************************************************************
'* Note : This is the only place in this project where you will see                            *
'* IF [mycheck].value = [ BOOLEAN ] then                                                       *
'* Why ? Becauseit is already Ready for that In ModOptions Check it out !!                     *
'* The onlly reason why it hasnt been used here is because it may                              *
'* Generate errors because of Attempting too Make Something Visible when the form is INvisible *
'***********************************************************************************************
Private Sub Command1_Click()

If Chcktoolbar.Value = vbChecked Then
SaveRegistryString "Toolbar", "Visible", 1
' Toolbar Will be Visible At startup
Else
If Chcktoolbar.Value = vbUnchecked Then
SaveRegistryString "Toolbar", "Visible", 0
' Toolbar  Will Not Be visible At startup
End If
End If

If Chckwordwrap.Value = vbChecked Then
SaveRegistryString "Wordwrap", "Wordwrap", 1
' this will save the settings too the registry
Else
If Chckwordwrap.Value = vbUnchecked Then
SaveRegistryString "Wordwrap", "Wordwrap", 0
End If
End If

If Chckassociations.Value = vbChecked Then
SaveRegistryString "chckassociations", "show", 1
Else
If Chckassociations.Value = vbUnchecked Then
SaveRegistryString "chckassociations", "show", 0
End If
End If

If ChckExternalEditor.Value = vbChecked Then
DetectExternalEditor
      SaveRegistryString "UseExternalEditor", "use", 1
         SaveRegistryString "UseExternalEditor", "path", ExternalEditorPath

    Else
If ChckExternalEditor.Value = vbUnchecked Then
      SaveRegistryString "UseExternalEditor", "use", 0
End If
End If

  If ChckAskIfToobig.Value = vbChecked Then
      SaveRegistryString "Misc", "AskifTooBig", 1
  Else
      SaveRegistryString "Misc", "AskifTooBig", 0
  End If
If chckallowassociation.Value = vbChecked Then
 SaveSettingString HKEY_CLASSES_ROOT, _
 "Txtfile\shell\open\command", _
 "", App.Path & "\" & App.EXEName & ".EXE" & " %1"
SaveRegistryString "associations", "isassociated", "1"
Else
SaveRegistryString "associations", "isassociated", "0"
End If

SaveSetting_Prioritylevel 10

Dim msg, style, Response, title

msg = "Settings Have been successfully saved" & _
vbCrLf & vbCrLf & "Would you like too Start NextPad now ?"
 
Beep
Response = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "NextPad")
Beep

Select Case Response 'Begin select case clause
Case vbYes ' if user selects The yes button then
ShellNewNextPad (vbNormalFocus)
Unload Me ' unload this form frmnoreg
End ' Stop code so the hidden form1 can be terminated
' and the new one can be shelled with the new settings
Case vbNo
End
End Select
End Sub

Private Sub cmdsetdefaultoptions_Click()
Chcktoolbar.Value = vbChecked
ChckExternalEditor.Value = vbChecked
Chckwordwrap.Value = vbChecked
chckallowassociation.Value = vbChecked
ChckAskIfToobig = vbChecked
End Sub

Private Sub Form_Load()
Beep ' beep too grab the users attention

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
' if the user closes this window we dont
'want NextPad too continually stay open
'in the backround so well stop
'the code and exit immediately.
End Sub
