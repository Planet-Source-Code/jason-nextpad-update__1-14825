VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options - Toolbar "
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   1500
   ClientWidth     =   5940
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Help"
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   3360
      Width           =   5775
      Begin VB.TextBox TxtHelp 
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   3
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   5535
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox Check3 
         Caption         =   "&Allow Usage"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Advanced"
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   5295
         Begin ComctlLib.Slider Sldpriority 
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   327682
            Max             =   3
         End
         Begin VB.Label Lbl 
            Caption         =   "Process Priority Class "
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Lbl 
            Caption         =   "Real Time"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   28
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Lbl 
            Caption         =   "High"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   27
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Lbl 
            Caption         =   "Normal"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   26
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Lbl 
            Caption         =   "Idle "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Only Advanced Users Should Change These Settings."
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   120
         Width           =   4215
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":000C
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   2
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   5535
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "External Editor"
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   5295
         Begin VB.CheckBox Chckaskiftoobig 
            Caption         =   "Ask Too &launch External Editor if File Is too Large too Open"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   5055
         End
         Begin VB.Frame Frame3 
            Caption         =   "Current External Editor"
            Height          =   735
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   5055
            Begin VB.CommandButton cmdChooseexternaleditor 
               Caption         =   "Select &External Editor ......"
               Height          =   375
               Left            =   120
               TabIndex        =   21
               ToolTipText     =   "Allows You too Select an External Editor ......."
               Top             =   240
               Width           =   4815
            End
         End
         Begin VB.CheckBox ChckExternalEditor 
            Caption         =   "&Use External Editor Too open Files Too large For NextPad Too open."
            Height          =   375
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   $"frmOptions.frx":044E
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Options For Setting up and Configuring The External Editor."
         Height          =   375
         Left            =   720
         TabIndex        =   33
         Top             =   120
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmOptions.frx":04F7
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   0
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   1
      Left            =   240
      ScaleHeight     =   2700
      ScaleMode       =   0  'User
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame1 
         Caption         =   "File associations"
         Height          =   2010
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   5295
         Begin VB.CheckBox Chckassociations 
            Caption         =   "&Nextpad should check wether it is the default text viewer"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Enable this option if you would like for NextPad too check wether it is the default text viewer."
            Top             =   720
            Width           =   4335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Allow Nextpad too be associated with Text files (*.TXT)"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Enable this option too associate NextPad  With Text files , disable for It not too be associated with Text files. "
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Options For NextPad's File Association."
         Height          =   375
         Left            =   600
         TabIndex        =   34
         Top             =   120
         Width           =   4815
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   0
         Picture         =   "frmOptions.frx":2221
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   1560
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.PictureBox Picoptionsd 
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox nonusable 
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   13
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox ippy 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   19
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   12
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2580
      Index           =   0
      Left            =   240
      ScaleHeight     =   2580
      ScaleWidth      =   5535
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   5535
      Begin VB.Frame fraSample1 
         Caption         =   "Toolbar"
         Height          =   2130
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   5415
         Begin VB.CheckBox Check2 
            Caption         =   "&Always Show toolbar  (default)"
            Height          =   255
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Enable this option too always show the toolbar , Disable Too hide the toolbar"
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Option For Enabling and Disabling The Toolbar"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   120
         Width           =   4695
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   0
         Picture         =   "frmOptions.frx":2CDB
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "A&pply"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Saves any changes you have made without closing this dialog box."
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      ToolTipText     =   "Cancels any changes you have made and Closes this dialog box."
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "saves changes and closes This dialog box. "
      Top             =   4695
      Width           =   1095
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   3165
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5583
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Toolbar "
            Key             =   "Group1"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click for toolbar options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "File associations"
            Key             =   "Group2"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click for File associations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "External Editor    "
            Key             =   "Group3"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click For External Editor Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced      "
            Key             =   "Group4"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click For Advanced Settings "
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ChckAskIfToobig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option When Enabled Tells NextPad too Launch The External Editor Without Notifying the user about it , This in turn May make Thigs a bit easier but can Confuse some. NOTE : In order for this option too have any affect you must have on " & Chr$(34) & " Execute External Editor When File Is too large too open." & Chr$(34)

End Sub


Private Sub Chckassociations_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option if Enabled would Allow NextPad too prompt the user if NextPad isnt The Default Viewer For (*.TXT) Files On the Computer it Is Running on , Not Recommended it is Highly Annoying !!!"
End Sub

Private Sub ChckExternalEditor_Click()

    Select Case ChckExternalEditor.Value
     Case 0
      Frame3.Enabled = False
      cmdChooseexternaleditor.Enabled = False
      ChckAskIfToobig.Enabled = False
     Case 1
      Frame3.Enabled = True
      cmdChooseexternaleditor.Enabled = True
      ChckAskIfToobig.Enabled = True
    End Select


End Sub





Private Sub ChckExternalEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option is highly Recommended !! If allowed, if a File is too large for NextPad too open , It will prompt the user too open the file with a More Powerful External Editor."
End Sub

Private Sub Chckwordwrap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "Enable This Setting Too Wrap the Text in NextPads Window too the window , Disable for it Not Too be Wrapped Too the window."

End Sub


Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option Will Allow You too have NextPad be The Default Viewer For (*.TXT) Files ( also Highly Recommended as Well )"
End Sub


Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option If Enabled Always Shows The Toolbar ( You Can Toggle The Toolbar ON and OFF in the View Menu ) if Disabled The Toolbar Isnt Always Shown"
End Sub

Private Sub Check3_Click()
   Dim i As Integer
   Select Case CBool(Check3.Value)
    Case False
      Sldpriority.Enabled = False
      Frame4.Enabled = False
      Check3.Enabled = True
    Case True
      Sldpriority.Enabled = True
      Frame4.Enabled = True
      Check3.Enabled = True
   End Select
    
End Sub

Private Sub cmdApply_Click()
  
 Call SaveMainSettings
  RetrieveALLSettings
Resizenotewithtoolbar

End Sub


Private Sub cmdCancel_Click()
    Unload Me
    Resizenotewithtoolbar
End Sub

Private Sub cmdChooseexternaleditor_Click()

   On Error GoTo CdlcCancelErr:
      CDialog.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
      CDialog.Filter = "Executable Files(*.EXE) |*.EXE"
      CDialog.DialogTitle = "Choose An External Viewer IE ;  Wordpad.EXE "
      CDialog.Cancelerror = True
    CDialog.ShowOpen
 
 If CDialog.Filename <> "" Then
   SaveRegistryString "Useexternaleditor", "Path", CDialog.Filename
    cmdChooseexternaleditor.caption = CDialog.Filename
 End If
 
CdlcCancelErr:
     If Err.Number = 32755 Then
      Exit Sub
     End If
End Sub

Private Sub cmdChooseexternaleditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "Allows you Too choose an External editor other than the default that is already being used."
End Sub

Private Sub cmdOK_Click()

 
 Call SaveMainSettings
  Unload Me
RetrieveALLSettings
Resizenotewithtoolbar


End Sub

Private Sub Cmdclose_Click()
Unload Me
End Sub
Private Sub SaveMainSettings()

   Dim retval As String ' Declare String variable
    
    retval = GetSettingString(HKEY_CLASSES_ROOT, _
    "Txtfile\shell\open\command", _
    "", App.Path & "\" & App.EXEName & ".EXE" & " %1")

  If Check1.Value = vbChecked Then
      SaveSettingString HKEY_CLASSES_ROOT, _
      "Txtfile\shell\open\command", _
      "", App.Path & "\" & App.EXEName & ".EXE" & " %1"
     SaveRegistryString "associations", "isassociated", "1"
  End If

   If Check1.Value = vbUnchecked Then

On Error GoTo cdialogerr:
  If retval = App.Path & "\" & App.EXEName & ".EXE" & " %1" Then ' 2

   CDialog.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
   CDialog.Filter = "Executables (*.EXE) |*.EXE"
   CDialog.Cancelerror = True
   CDialog.DialogTitle = "Please select notepad (Usually Located In your windows directory) and press open"
   CDialog.ShowOpen

   MsgBox CDialog.Filename & _
   vbCrLf & " will now be used too view text files on this computer", vbInformation, "NextPad"

     If CDialog.Filename <> "" Then ' 3
       SaveSettingString HKEY_CLASSES_ROOT, _
       "Txtfile\shell\open\command", _
       "", CDialog.Filename & " %1"
       SaveRegistryString "associations", "isassociated", "0"
     End If
' 2 TODO : Else Statement
  End If
' 3 TODO : ELSE STATEMENT
End If


'*****************************************************************
' the Subs Here are in ModOptions , they work depending on the
' Checks value then save the setting requested associated with the sub.
      
      SaveSetting_Toolbar CBool(Check2.Value)

      SaveSetting_chckassociations CBool(Chckassociations.Value)

      saveSetting_UseExternalEditor CBool(ChckExternalEditor.Value)

        If Sldpriority.Enabled = False Then
          SaveSetting_Prioritylevel 10 'set it too a value where prority Settings Will be ignored
        Else
          SaveSetting_Prioritylevel CInt(Sldpriority.Value)
        End If
      
     ' SaveSetting_Wordwrap CBool(Chckwordwrap.Value)

      SaveSetting_AskIfToobig CBool(ChckAskIfToobig.Value)

'******************************************************************

cdialogerr:
   Exit Sub

End Sub


Private Sub Form_Load()
 Dim retval As String
 Dim R As String, Response
   On Error GoTo AccessSettingsError
  
  R = GetSettingString(HKEY_CLASSES_ROOT, _
      "txtfile\Shell\open\command", _
      "", "")

    If R = App.Path & "\" & App.EXEName & ".EXE" & " %1" Then
       Check1.Value = vbChecked
     Else
       Check1.Value = vbUnchecked
    End If

   If GetSetting("NextPad", "chckassociations", "Show", 1) = 1 Then
       Chckassociations.Value = vbChecked
    Else
       Chckassociations.Value = vbUnchecked
   End If
' above for NextPad associations

'Below For toolbar reg options

  If GetSetting("NextPad", "Toolbar", "Visible", 1) = 1 Then
     Check2.Value = vbChecked
   Else
     Check2.Value = vbUnchecked
  End If

' Below For External Editor Options

 If GetSetting("NextPad", "UseExternaleditor", "use", 1) = 1 Then
     ChckExternalEditor.Value = vbChecked
   Else
     ChckExternalEditor.Value = vbUnchecked
 End If

  
  'If GetSetting("NextPad", "Wordwrap", "Wordwrap", 1) Then
  '    Chckwordwrap.Value = vbChecked
   ' Else
  '    Chckwordwrap.Value = vbUnchecked
 ' End If
  
  If GetSetting("NextPad", "Misc", "AskifTooBig", 1) Then
      ChckAskIfToobig.Value = vbChecked
    Else
      ChckAskIfToobig.Value = vbUnchecked
  End If

   retval = GetSetting("NextPad", "UseExternalEditor", "Path")
      If retval = vbNullString Then
        DetectExternalEditor
         cmdChooseexternaleditor.caption = ExternalEditorPath
        Else
      If retval <> "" Then
         cmdChooseexternaleditor.caption = retval
     End If
     End If

     Select Case UseExternalEditor.use
      Case 0
        Frame3.Enabled = False
        cmdChooseexternaleditor.Enabled = False
        ChckAskIfToobig.Enabled = False
      Case 1
        Frame3.Enabled = True
        cmdChooseexternaleditor.Enabled = True
        ChckAskIfToobig.Enabled = True
     End Select
      
      If CInt(GetSetting("NextPad", "Priority", "Level")) > 3 Then
         Sldpriority.Enabled = False
         Frame4.Enabled = False
         Check3.Value = 0
      Else
         Frame4.Enabled = True
         Sldpriority.Enabled = True
         Sldpriority.Value = CInt(GetSetting("NextPad", "Priority", "Level"))
         Check3.Value = 1
      End If
      
AccessSettingsError:
     If Err.Number <> 0 Then
       Response = MsgBox("An Error Has Occured While Accessing One or More Registry Entrys , Settings May Be Missing Or Corrupt." _
       & vbNewLine & vbNewLine & "Repair Settings ?" & vbNewLine & vbNewLine & "Click Yes Too Repair Settings", vbYesNo + vbExclamation, "Error,Options")
     Select Case Response
       Case vbYes
         SaveSetting_Toolbar True
         SaveSetting_chckassociations False
         saveSetting_UseExternalEditor True
         SaveSetting_Prioritylevel 10
         SaveSetting_Wordwrap True
         SaveSetting_AskIfToobig True
         SaveRegistryString "associations", "isassociated", "0"
         MsgBox "All Settings Were Saved Successfully." & vbNewLine & "Please Reload the Options Dialog for the Settings Too Take Effect.", vbInformation, "Successful"
         Unload Me
         Exit Sub
       Case vbNo
         MsgBox "Settings Were Chosen Not To be Repaired This Message Will OCntinue Too apear Until you do so.", vbCritical, "Settings Not Saved"
         Unload Me
       Exit Sub
     End Select
     End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        ' visual basic has created the bottom code
    'for the tab's key constants
    '\\\\\\\\\\\\\\\\\\\///////////////////////////
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmOptions = Nothing
End Sub

Private Sub Sldpriority_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtHelp.Text = "This Option Sets The process Priority Level For NextPad , Moving the Slider Higher Can Slow Down Windows or Crash it. Also NOTE : it Is HIGHLY Recommended you Leave This option at Its Default Level " _
& Chr(34) & "Idle" & Chr(34) & ", Furthermore it Has Not been Tested On Windows 2000,NT, or ME so I suggest You Leave This Option Alone."
End Sub

Private Sub tbsOptions_Click()
    Select Case tbsOptions.SelectedItem.Index
    Case 1
        frmOptions.caption = "Options - Toolbar "
    Case 2
        frmOptions.caption = "Options - File Associations"
    Case 3
        frmOptions.caption = "Options - External Editor"
    Case 4
           frmOptions.caption = "Options - Advanced"
     Case 5
              frmOptions.caption = "Options - Misc."
    End Select
    ' ABOVE ^^^^^^^ Use Select Case Statement Instead
    '       |||||||
    ' Of Barberic IF THEN Statement Too Set the forms Caption
    ' Depending on the options selected Through The Tbsoptions
    ' .selecteditem.index Property
    
    Dim i As Integer
    ' visual basic has created the bottom code
    'for the tab's
    '\\\\\\\\\\\\\\\\\\\///////////////////////////
    'show and enable the selected tab's controls
    'and hide and disable all others
     TxtHelp.Text = ""
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            Picoptions(i).Left = 210
            Picoptions(i).visible = True
            Picoptions(i).Enabled = True
        Else
            Picoptions(i).Left = -20000
            Picoptions(i).Enabled = False
        End If
    Next
    
End Sub



Private Sub TxtHelp_KeyPress(KeyAscii As Integer)
Beep
End Sub
