VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Untitled - NextPad                                                         "
   ClientHeight    =   5805
   ClientLeft      =   1035
   ClientTop       =   2565
   ClientWidth     =   8700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   767
      ButtonWidth     =   714
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "new"
            Description     =   ""
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "open"
            Description     =   ""
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "save"
            Description     =   ""
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "copy"
            Description     =   ""
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "cut"
            Description     =   ""
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "paste"
            Description     =   ""
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "find"
            Description     =   ""
            Object.ToolTipText     =   "Find..."
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "close"
            Description     =   ""
            Object.ToolTipText     =   "Close file"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "properties"
            Object.ToolTipText     =   "Properties  (File)"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "time&date"
            Object.ToolTipText     =   "Insert Time and Date"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "options"
            Object.ToolTipText     =   "Options"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "About"
            Object.ToolTipText     =   "About Text pad"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.TextBox txt 
      Height          =   4215
      Index           =   2
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog CfontDialog 
      Left            =   2160
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer mainTimer 
      Interval        =   1
      Left            =   120
      Top             =   4920
   End
   Begin VB.TextBox txt 
      Height          =   4250
      Index           =   1
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label lblfilename 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   5535
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":27A2
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":28B4
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3096
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":31A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":32BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":34DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnufileitem 
      Caption         =   " &File  "
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnunewfileitem 
         Caption         =   "&New "
      End
      Begin VB.Menu mnuopenfileitem 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuclosefileitem 
         Caption         =   "&Close "
      End
      Begin VB.Menu mnusaveitem 
         Caption         =   "S&ave"
      End
      Begin VB.Menu mnusaveasfile 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproperties 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnurecentfile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuedititem 
      Caption         =   " &Edit  "
      Begin VB.Menu mnuundo 
         Caption         =   "Und&o"
         Shortcut        =   ^Z
      End
      Begin VB.Menu line14 
         Caption         =   "-"
      End
      Begin VB.Menu mnucopyitem 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucutitem 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnupasteitem 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnudeleteitem 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselectallitem 
         Caption         =   "&Select All"
         Shortcut        =   ^S
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinserttimeitem 
         Caption         =   "Insert &Time"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuinsertdateitem 
         Caption         =   "Insert &Date"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuinserttimeanddateitem 
         Caption         =   "I&nsert Time\Date"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuclearclipboardtextitem 
         Caption         =   "Cl&ear Clipboard Text "
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuviewclipboardtextitem 
         Caption         =   "&View Clipboard Text..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnueditclipboard 
         Caption         =   "Edit Clipboard Te&xt... "
         Shortcut        =   ^E
      End
      Begin VB.Menu line13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuwordwrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnusetfont 
         Caption         =   "Set &Font"
      End
   End
   Begin VB.Menu mnusearchitem 
      Caption         =   " &Search  "
      Begin VB.Menu mnufinditem 
         Caption         =   "Fin&d..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnufindnextitem 
         Caption         =   "&Find Next"
      End
   End
   Begin VB.Menu mnuoptionsitem 
      Caption         =   " &View  "
      Begin VB.Menu mnulaunchnewinstanceitem 
         Caption         =   "&Launch new instance"
         Begin VB.Menu mnunormal 
            Caption         =   "&Normal "
         End
         Begin VB.Menu mnumaximized 
            Caption         =   "&Maximized "
         End
         Begin VB.Menu mnuminimized 
            Caption         =   "Minimi&zed"
         End
         Begin VB.Menu line2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuchoose 
            Caption         =   "&Choose..."
         End
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhidetoolbaritem 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnudeleteatextfileitem 
         Caption         =   "Text &File Manager..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufullscreen 
         Caption         =   "Full &Screen..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptionsdialogitem 
         Caption         =   "&Options... "
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu MNUHELPITEM 
      Caption         =   "  &Help    "
      Begin VB.Menu mnubugfixes 
         Caption         =   "&Version History..."
      End
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnueditFileinfoitem 
      Caption         =   "&mnueditFileinfoitem"
      Visible         =   0   'False
      Begin VB.Menu mnucopyitem1 
         Caption         =   "&Copy"
      End
   End
   Begin VB.Menu Mnueditfileinfoitem2 
      Caption         =   "&Mnueditfileinfoitem2"
      Visible         =   0   'False
      Begin VB.Menu Mnucopyitem2 
         Caption         =   "&Copy"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
' Since Everything ( Loading Wise ) Is in Sub Main() In Modmain
' We Havent Really Found Any thing Too use This Event Procedure For ....
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim msg, Response ' declare variables
   ' if the filestate is dirty then go to unload:
   If fstate.dirty = True Then GoTo Unload:
   ' else if its not END
   If fstate.dirty = False Then End
   


Unload: ' Vb will jump here if the File Sate is dirty

    If Form1.lblfilename.caption <> "" Then ' if the filename or anything in that label is equal to anything except "" then
    msg = "The Text in " _
   & Form1.lblfilename.caption & " File has changed" _
   & Chr(13) & Chr(13) & "Do you wish too save The changes ?"
Else ' * If there isnt then Display The One Below
   msg = "The Text in the untitled file has changed" _
   & Chr(13) & Chr(13) & "Do you wish too save the changes?"
End If

' show the MsgBox with the msg , Buttons ; Style , title
  Response = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton3, "NextPad ")

   Select Case Response ' select a response
    Case vbYes     ' User chose the Yes button .
     Filequicksave ' Call Procedure in modmain
     End ' Stop and quit
    Case vbCancel ' User chose the Cancel Button
     Cancel = True ' Escape the Unload mode
      Exit Sub ' Exit the sub before running any more code in this sub
    Case vbNo ' the user chose the No Button
     End ' Stop and quit
   End Select
End Sub
Private Sub Form_Resize()
   On Error GoTo Nxterr: ' if An error occurs go to that line
  ' A form CANNOT be Resized while it is minimized so if it is
  ' Exit This Sub Immediatley
  If Form1.WindowState = vbMinimized Then Exit Sub
   Call Resizenotewithtoolbar ' call resize procedure

Nxterr:
 If Err.Number <> 0 Then ' if the errors number equlas anything other than 0 then
   Exit Sub ' exit this sub immediately
 End If
End Sub
Private Sub Mnuabout_Click()
   Beep ' beep to grab the users attention
   Load frmAbout ' load the form
   frmAbout.Show (vbModal) ' show the form in Vbmodal mode
End Sub
Private Sub mnubugfixes_Click()
   Load Frmbugz ' load the form
   Frmbugz.Show (vbModal) ' show the form
End Sub

Private Sub mnuchoose_Click()
   Load Frmnewinstance ' load the form
   Frmnewinstance.Show (vbModal) ' show the form in Vbmodal Mode
End Sub

Private Sub Mnuclearclipboardtextitem_Click()
    ' clear the clipboard of any Text it may have
    Clipboard.SetText ("")
End Sub

Private Sub mnuclosefileitem_Click()
   ' ** since we are closing the file prompt the user too save
   ' ** any changes that have been made
   Call mnunewfileitem_Click
   
   On Error GoTo ErrFile: ' if an error occurs Go too that line
   CommonDialog1.Filename = ("") ' set the commondialog's filename property too  nothing ("")
   Form1.ActiveControl.Text = "" ' set the forms active control's text too nothing ("")
   Form1.caption = "Untitled - NextPad" ' set the forms caption
   lblfilename.caption = ("") ' set the Lblfilename's caption property too nothing
   fstate.dirty = False ' Tell NextPad That This is False now
   Close #1 ' close the file that is currently open ( If any )
ErrFile: ' if an error occurs vb will jump here
 If Err.Number <> 0 Then ' if the errors number is anything close or above zero then
  Exit Sub ' exit the sub immediately
 End If
End Sub
Private Sub mnucopyitem_Click()
  
  Clipboard.SetText Form1.ActiveControl.SelText 'set the clipboard text to
  ' the forms active control's Selected Text
End Sub
Private Sub mnucutitem_Click()
   ' set the clipboards text too the forms active control's text to
   ' selected text
   Clipboard.SetText Form1.ActiveControl.SelText
   Form1.ActiveControl.SelText = "" ' since we are cutting NOT copying
   ' set the forms active control's text too Nothing ("")
End Sub
Private Sub mnudeleteatextfileitem_Click()
    Load Frmtxtfilemanager ' load the form
    Frmtxtfilemanager.Show  ' show the form in Vbmodal Mode
End Sub
Private Sub mnudeleteitem_Click()
    ' delete the selected text in the form's active control's text
    ' property
    Form1.ActiveControl.SelText = ""
End Sub

Private Sub mnueditclipboard_Click()
    Load frmeditclipboard ' load the form
    frmeditclipboard.Show (vbModal) ' show the form in Vbmodal Mode
End Sub

Private Sub mnuexit_Click()
   Call Form_QueryUnload(1, 0)
' instead of Repitiusly leaving this exit code here
' Too save Space and make NextPad faster and smaller
' well just reuse Form_queryunloads code instead


End Sub
Private Sub mnufileinformationitem_Click()
    Load frmfileinfo ' load the form
    frmfileinfo.Show (vbModal) ' show the form in Vbmodal mode

End Sub
Private Sub mnufinditem_Click()
    Load frmfind ' load the form
    frmfind.Show (vbModeless), Me ' show the form Telling VB
    ' that the owner form is Me (Form1)
End Sub

Private Sub mnufindnextitem_Click()
    findnexttext ' Just call the same exact code from Modmain
' otherwise if we had rewritten it , it would
' be just another waste of code
End Sub
Private Sub mnufullscreen_Click()
 ' ** here we Select a Case of BOOLEAN
 ' ** Frmfullscreen.visble = [BOOLEAN]

  Select Case frmfullscreen.visible ' select a case of true or false
   Case False ' Frmfullscreen IS NOT visible
      mnufullscreen.Checked = True 'check the menu
      Load frmfullscreen ' load the form
      frmfullscreen.Show ' show the form
      Me.WindowState = vbMinimized ' Minimize This form
   Case True ' frmfullscreen IS visible
      mnufullscreen.Checked = False ' Uncheck the Menu
      Unload frmfullscreen ' Unload the Form
      Unload Frmleavefullscreen ' Unload the form
  End Select

End Sub

Private Sub mnuhidetoolbaritem_Click()
' ** Here we alter the visibility of the toolbar ,
' ** Update Registry Settings For the toolbar ,
' ** and Set and Unset Variables
    Toolbar.visible = Not Toolbar.visible
    '*************************************************
    ' ** Heres How it works, ;
    ' ** The Not keyword works like this ;
    ' ** Not uses a Logical return so if, Me.visible = true ,
    ' ** Not will Return the opossite which is False (0)
    '*************************************************
    ' Change the check to match the current state
    mnuhidetoolbaritem.Checked = Toolbar.visible
    ' Call the resize procedure
    Resizenotewithtoolbar ' resize the form
    
    ' ** here we set ; In the Registry ,
    ' ** if the toolbar is going too be visible the next time
    ' ** NextPad startup
    
    ' ** OK Here is where it May get tricky let me explain
    ' ** Code Chart  for expression
    ' Abs([Expression]) Returns the absolute value of a number
    ' Cint([Expression]) Converts an expression too an integer
    ' Cbool([Expression]) Converts an expression too a Boolean value of TRUE (-1) or FALSE (0)
    
    ' What we do here is simple We first
    ' Receive the absolute value of the number returned from converting
    ' the expression to an integer from converting it too BOOLEAN
    '; Since NextPad , At startup Doesnt check the toolbar menu
    ' because if the value Converted too BOOLEAN the integer Was true Would Return
    ' (-1) which would cause the problem , But if we Converted that -1 to 1 using the
    ' ABS Function we would be ok
    
    SaveRegistryString "Toolbar", "Visible", Abs(CInt(CBool(Toolbar.visible)))
      
      
End Sub

Private Sub mnuinsertdateitem_Click()
   On Error GoTo memoryerror: ' If an error Occurs Jump too that line
    
    Form1.ActiveControl.SelText = Date ' Set the Form's Active Control's
    ' selection point ( where the system Caret is )
    'Text to the systems current Date


memoryerror:
   If Err.Number <> 0 Then ' if an error is anything above or at zero then
   ' display the error message too the user
    MsgBox "NextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
    & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
    Exit Sub ' Exit the Sub Immediately
   End If

End Sub

Private Sub mnuinserttimeanddateitem_Click()
    On Error GoTo memoryerror:
     
     Form1.ActiveControl.SelText = Now ' Set the Form's Active Control's
    ' selection point ( where the system Caret is )
    'Text to the systems current Date & Time



memoryerror:
    If Err.Number <> 0 Then
     ' display the error Message too the user
     MsgBox "NextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
     & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
     
     Exit Sub ' Exit the sub Immediateley
    
    End If

End Sub

Private Sub mnuinserttimeitem_Click()
   On Error GoTo outofmemory:
     Form1.ActiveControl.SelText = Time ' Set the Form's Active Control's
    ' selection point ( where the system Caret is )
    'Text to the systems current Time


outofmemory:
   If Err.Number <> 0 Then
   ' display the error Message too the user
    MsgBox "NextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
    & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
    
    Exit Sub     ' Exit the sub Immediateley

   End If

End Sub


Private Sub mnumaximized_Click()
    ' Call the ShellnewNextPad Function in Modmain
    ShellNewNextPad (vbMaximizedFocus)
    ' This Shells NextPad Giving the New Instance
    ' A maximized focus
End Sub

Private Sub mnuminimized_Click()
    ' Call the ShellnewNextPad Function in Modmain
    ShellNewNextPad (vbMinimizedFocus)
    ' This Shells NextPad Giving the New Instance
    ' A minimized focus

End Sub

Private Sub mnuundo_Click()

    
  SendMessage Form1.ActiveControl.hwnd, EM_UNDO, 0&, 0&
 
     
End Sub

Private Sub mnunewfileitem_Click()
    'this Creates a new File and Propmpts the user
    ' for Any Changes That have been made
    ' ** The Real Code is In Modmain , this just checks the
    ' ** condition of the file state and what action too take
    
    Select Case fstate.dirty ' select a case of TRUE(-1) or FALSE(0)
    
    Case True ' File state is Dirty True (-1)
      newfile 'call newfile procedure in modmain
    Case False ' File state is NOT Dirty False (0)
      Form1.ActiveControl.Text = ("") ' set the forms Active Control's Text Too nothing ("")
      lblfilename.caption = "" ' set lblfilenames caption Property
      fstate.dirty = False 'Set FIle State Too False Since We Didnt Chancge anything
      Form1.caption = "Untitled - NextPad" ' Set the Forms Caption using the Caption Property
    End Select
End Sub

Private Sub mnunormal_Click()
       ' Call the ShellnewNextPad Function in Modmain
     ShellNewNextPad (vbNormalFocus)
    ' This Shells NextPad Giving the New Instance
    ' A Normal focus


End Sub

Private Sub mnuopenfileitem_Click()
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' This menu Item When clicked Will check if the file Open ( If any)
' Nedds to be saved if so , Display the Message Too the user , And Calculate a response .
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Close #1 'Close the file just in case it is open
Dim msg, Response ' declare Variables
Dim regval As String ' Declare variables
Dim caption, Strfilerecent ' Declare Variables
  
   If fstate.dirty = False Then GoTo openfileproc:
   If fstate.dirty = True Then
    
' Msg Variable ///////////////////////////////////////////
' Set Msg Variable Too Display Proper Message Depending on Circumstances ...
  If lblfilename.caption <> "" Then
     msg = "The Text in " _
      & lblfilename.caption & " File has changed" _
       & Chr(13) & Chr(13) & "Do you wish too save The changes ?"
     Else ' * If there isnt then Display The One Below
        msg = "The Text in the untitled file has changed" _
         & Chr(13) & Chr(13) & "Do you wish too save the changes?"
    End If
'//////////////////////////////////////////////////////////
End If
       ' set response variable too display msgbox with the Defined buttons and title
       Response = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton3, "NextPad")

   Select Case Response ' detect Which Button was pressed by the user.....
    
     Case vbYes ' user clicked the yes Button
       Filequicksave ' call procedure in modmain
        CommonDialog1.Filename = ("")
         lblfilename.caption = ""
        
         Case vbNo ' user clicked the No Button
            GoTo openfileproc:
           Case vbCancel ' User clicked The Cancel Button
             Exit Sub
    End Select
     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
openfileproc:
      Reset ' Reset all open Disks
       Close #1 ' close before using it again
        FreeFile (1) 'Free the File
         On Error GoTo Cdlogerror:
     
     With CommonDialog1
        .Flags = Normal_Cdlogflags  ' use the constant in modmain
        .Cancelerror = True ' set Cancel Error Too true When User clicks cancel an error will gernrate
        ' set the filter
        .Filter = "Text Files (*.TXT) |*.TXT| Ini Files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.*"
            ' set the commondialogs title
        .DialogTitle = "Open File"
        .ShowOpen ' Show the open dialog
     End With
                Openfile (CommonDialog1.Filename)



Cdlogerror: ' error that is triggered when the cancel button on the Commondialog is pressed
    If Err.Number = 32755 Then
     Exit Sub
    End If
      'Openfile ' call openfile proc in modmain
End Sub

Private Sub mnuoptionsdialogitem_Click()
     On Error GoTo Objectunloadederror: 'if an error occurs Jump too that line
      Load frmOptions ' Load the form
      frmOptions.Show (vbModal) ' show the form In Vbmodal Mode



Objectunloadederror: '
    If Err.Number <> 0 Then
       ' display the error Message too the user
     ' MsgBox "NextPad Has encountered An Unexpected error" _
      '& Chr(13) & Err.Description, vbCritical, "Error"
      
      Exit Sub 'exit the sub immediately
    End If
End Sub

Private Sub mnupasteitem_Click()
     On Error GoTo Outofmemoryerr:
      ' ** In order too paste We Have too Get Text Form the
      ' ** Clipboard so we Set the Form's active control's Text
      ' ** Too the text That we Get ( If any ) Form the Clipboard
      Form1.ActiveControl.SelText = Clipboard.GetText()


Outofmemoryerr:
     If Err.Number <> 0 Then
    ' display the error Message too the user
      MsgBox "NextPad Has Encountered The Following Error(s) While Pasting Your Selection : " _
      & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Paste Anymore Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
      'exit the sub immediately
      Exit Sub
     End If
End Sub

Private Sub mnuproperties_Click()

Dim R As Long 'declare variables
     
     R = ShowFileproperties(Me.hwnd, lblfilename.caption)
     ' show fileinfo  the hwnd Of this window ; and the lblfilename.caption property
End Sub


Private Sub mnurecentfile_Click(Index As Integer)
  'Call the openrecentfile sub in modmain
  openrecentfile
End Sub


Private Sub mnusaveasfile_Click()
'********************************
'This method of saving will Give the user
'a choice of how too save the file
'With a certain extension
'********************************
   Close #1 ' close the file just in case
   CommonDialog1.Cancelerror = True ' set the commondialogs cancel error too True (-1)
   'set the commondialogs filter what files it should accept
   CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.* "
   CommonDialog1.DialogTitle = "Save As" ' set the commondialogs title
   ' set the common dialogs flags
   CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt

On Error GoTo dialogerror ' if an error occurs goto dialogerror
   CommonDialog1.ShowSave
   If CommonDialog1.Filename <> "" Then ' if The Commondialogs filename is equal too anything other than ""
   Open CommonDialog1.Filename For Output As #1 ' open the file for output
   Print #1, Form1.ActiveControl.Text ' print The form's activecontrols text too the file
   Close #1 ' we are now done with the file so close it
   fstate.dirty = False ' Since we are saving the file state stays dirty
   lblfilename.caption = CommonDialog1.Filename ' the filename Holder's caption is set too the Commondialogs Filename
   Form1.caption = lblfilename.caption & " - NextPad" ' set the forms caption too [ the filename ] - NextPad
dialogerror: ' if an error occurs Vb Will Jump here
   If Err.Number <> 0 Then ' if the error's number is = to anything
   Exit Sub ' exit the sub Immedieately
   Close #1 ' Close the file Just in case
   End If
   End If

End Sub

Private Sub mnusaveitem_Click()
Filequicksave
' call The filequicksave Procedure in module1
' Because Reusing Code OVER AND OVER AND OVER
' Can Get REALLY annoying and can slowdown NextPads
' Start up
fstate.dirty = False
End Sub

Private Sub mnuselectallitem_Click()
   ' set the forms active controls
   ' selection start point to the beginning
    Form1.ActiveControl.SelStart = 0
   '  set the forms active control's selection length  too
   '  the lenth of The Form's active control's Text
    Form1.ActiveControl.SelLength = Len(Form1.ActiveControl.Text)
End Sub

Private Sub mnusetfont_Click()
  On Error GoTo Cancelerror: ' if an error occurs goto that line
   CfontDialog.Flags = cdlCFBoth ' set the Font Dialogs flags
  CfontDialog.ShowFont ' show the Select Font Dialog
  ' set the forms active control's font too the selected font form the dialog
  Form1.ActiveControl.Font = CfontDialog.fontname
  ' set the forms active control's Bold too the Selected font dialog bold
  Form1.ActiveControl.FontBold = CfontDialog.FontBold
  ' set the form's active control's Font size too the selected font size from the dialog
  Form1.ActiveControl.fontsize = CfontDialog.fontsize
  ' save in the registry ; The Font For the Form's Text Box
  ' For later Retrieval From the registry
  SaveRegistryString "Font", "font", Form1.ActiveControl.fontname
  ' save in the registry ; the controls current Font size for later
  ' retrieval from the registry
  SaveRegistryString "Font", "Fontsize", Form1.ActiveControl.fontsize
Cancelerror: ' if an error occurs Vb will Jump here
  Exit Sub ' Exit the sub immediately
End Sub


Private Sub mnuviewclipboardtextitem_Click()
      On Error GoTo Objectunloaded: ' if an error Occurs Vb Will jump too that line
       
       Load frmclpboard ' Load the form
       frmclpboard.Show (vbModal) ' Show the form in VbModal Mode


Objectunloaded:
      Exit Sub ' Exit this Sub Immediately
End Sub

Private Sub mnuwordwrap_Click()
  Dim retval As Boolean
     
    retval = Usewordwrap ' set variable too usewordwrap's value
    
    Usewordwrap = Not Usewordwrap ' usewordwrap =  not usewordwrap
        '*************************************************
    ' ** Heres How it works, ;
    ' ** The Not keyword works like this ;
    ' ** Not uses a Logical return so if, Me.visible = true ,
    ' ** Not will Return the opossite which is False (0)
    '*************************************************
     retval = Usewordwrap
         
         Togglewordwrap (retval) ' togglewordwrap (Retval as BOOLEAN)
         ' This basically toggles wordwrap from the retval of Usewordwrap = Not usewordwrap
    ' ** here we set ; In the Registry ,
    ' ** if the which text box is going too be visible the next time
    ' ** NextPad is started
    
    ' ** OK Here is where it May get tricky let me explain
    ' ** Code Chart  for expression
    ' Abs([Expression]) Returns the absolute value of a number
    ' Cint([Expression]) Converts an expression too an integer
    ' Cbool([Expression]) Converts an expression too a Boolean value of TRUE (-1) or FALSE (0)
    
    ' What we do here is simple We first
    ' Receive the absolute value of the number returned from converting
    ' the expression to an integer from converting it too BOOLEAN
    '; Since NextPad , At startup Doesnt check the wordwrap menu
    ' because if the value Converted too BOOLEAN the integer Was true Would Return
    ' (-1) which would cause the problem , But if we Converted that -1 to 1 using the
    ' ABS Function we would be ok

   SaveRegistryString "Wordwrap", "Wordwrap", Abs(CInt(CBool(Usewordwrap)))

End Sub


Private Sub mainTimer_Timer()
'**********************************************
'// Main timer Events ;
'// Interval 1 millisecond
'**********************************************
    On Error Resume Next ' if there is an error
    ' we dont want too exit the sub because the rest of the code wont be executed
    ' so just go too the next line
    If lblfilename.caption = "" Then ' if the Lblfilename's Caption property tells us
    ' that there is no Filename Loaded Then
     mnuproperties.Enabled = False ' Disable this menu
     ' since we cant use it if no filename is loaded
    Else ' Else If there is Text in the caption property .....
     mnuproperties.Enabled = True ' Enable the menu We ned it now Because
     ' A filename is currently loaded
    End If


    On Error Resume Next ' if there is an error
    ' we dont want too exit the sub because the rest of the code wont be executed
    ' so just go too the next line
     If Clipboard.GetText <> "" Then ' if the clipboard Has text Then ....
      Mnuclearclipboardtextitem.Enabled = True ' enable this menu Because now there is Clipboard text too clear
      mnupasteitem.Enabled = True ' enable this menu because now there is text too paste
      mnuviewclipboardtextitem.Enabled = True ' enable menu this now because now there is text too view
     Else
      Mnuclearclipboardtextitem.Enabled = False ' disable this menu now because now there is NO text too clear
      mnupasteitem.Enabled = False ' Disable this now because now there is no Text Too paste from the clipboard
      mnuviewclipboardtextitem.Enabled = False 'disable this now because now there is  no text Currently in the Clipboard too view
     End If

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
      On Error GoTo Errorhandler: ' if an error occurs Vb will jump too that line
        
        Select Case Button.Key ' Select The button Being pressed by
        ' using the select case statement also on the Buttons Key
     Case "open" ' user Clicked the Open Button
         Call mnuopenfileitem_Click
     Case "save" ' user Clicked the save Button
         Call mnusaveitem_Click
     Case "copy" ' user Clicked the copy Button
         Call mnucopyitem_Click
     Case "cut" ' user Clicked the cut Button
         Call mnucutitem_Click
     Case "paste" ' user Clicked the Paste Button
         Call mnupasteitem_Click
     Case "find" ' user Clicked the Find Button
         Call mnufinditem_Click
     Case "close" ' user Clicked the close Button
         Call mnuclosefileitem_Click
     Case "undo" ' User Clicked Undo Button
         Call mnuundo_Click
     Case "time&date" ' user Clicked the time&date Button
         Call mnuinserttimeanddateitem_Click
     Case "options" ' user Clicked the options Button
         Call mnuoptionsdialogitem_Click
     Case "About" ' user Clicked the about Button
         Call Mnuabout_Click
     Case "new" ' user Clicked the new Button
         Call mnunewfileitem_Click
     Case "properties" ' user Clicked the properties Button
         Call mnuproperties_Click
        End Select


Errorhandler:
    If Err.Number <> 0 Then ' if an error is equal too anything other or above zero then ......
     ' Display the error message too the user
     MsgBox "An Unexpected Error has occured while accessing the Toolbar", vbCritical, "NextPad"
     Exit Sub ' exit the sub immediately
    End If
End Sub

Private Sub Toolbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**********************************************
'// This Event is raised when the user releases the mousebutton
'// over the toolbar
'*********************************************
          
    If Button = 2 Then ' if the user
     ' Right Clicked on the toolbar then ........
           
     PopupMenu mnuoptionsitem, vbPopupMenuCenterAlign, , , mnuhidetoolbaritem
    ' Display the popup menu too the user
    ' Display the popup menu Center Aligned At X
    End If
 
End Sub

Private Sub txt_Change(Index As Integer)
     TextChangecontrol
    ' Call the TextChangecontrol Proc in
    ' Modmain too handle Repitious
    ' code that Just wastes Valuable Coding time and space
    ' on the users Hard-disk

End Sub
