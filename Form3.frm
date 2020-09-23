VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frmtxtfilemanager 
   Caption         =   "Text File Manager v 1.0"
   ClientHeight    =   6165
   ClientLeft      =   630
   ClientTop       =   2175
   ClientWidth     =   9945
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6165
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   5925
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            Text            =   "Text File Manager 1.0"
            TextSave        =   "Text File Manager 1.0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pictoobig 
      Height          =   8295
      Left            =   5280
      ScaleHeight     =   8235
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton CmdOpenwithExternaleditor 
         Caption         =   "Open With External Editor "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Open's The Selected File With The External Editor ( If Any )"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Lblwhy 
         Caption         =   "Error, The File selected Is too large too be displayed here. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.TextBox Txtfile 
      Height          =   8295
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.FileListBox File1 
      ForeColor       =   &H00000000&
      Height          =   4770
      Left            =   120
      Pattern         =   "*.TXT;*.INI;*.lOG"
      TabIndex        =   2
      Top             =   3480
      Width           =   4935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Frame frameleft 
      Caption         =   "Select a file :"
      Height          =   8175
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5175
   End
   Begin VB.PictureBox Picholderleft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   0
      ScaleHeight     =   5925
      ScaleWidth      =   5295
      TabIndex        =   7
      Top             =   0
      Width           =   5295
   End
   Begin VB.PictureBox Picholderright 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   3465
      ScaleHeight     =   5925
      ScaleWidth      =   6480
      TabIndex        =   8
      Top             =   0
      Width           =   6480
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopenwithNextpad 
         Caption         =   "&Open With NextPad"
      End
      Begin VB.Menu mnuopenwithexternaleditor 
         Caption         =   "Open &With External Editor"
      End
      Begin VB.Menu line0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproperties 
         Caption         =   "Properties (File) "
         Shortcut        =   ^P
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnustatusbar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu Mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Frmtxtfilemanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Lastdrive As String

Private Sub Cmdclose_Click()
  Unload Me 'unload the form
End Sub

Private Sub CmdOpenwithExternaleditor_Click()
   Call mnuopenwithexternaleditor_Click 'call this sub (Private)
End Sub

Private Sub Dir1_Change()

       File1.Path = Dir1.Path ' set file1's path too dir1's path
End Sub

Private Sub Drive1_Change()
Dim msg, style, title, Response ' declare variables
   On Error GoTo driveerr: ' if an error occurs go too that line
    
    Dir1.Path = Drive1.Drive ' set dirs path too drive 1's drive

driveerr: ' error Has been Encountered

 If Err.Number <> 0 Then
     Response = MsgBox("Error , Please insert a disk into the selected drive.", vbRetryCancel + vbExclamation + vbDefaultButton2, "Selected Drive error")

   Select Case Response
    Case vbRetry
     Call Drive1_Change
     Exit Sub

    Case vbCancel
     Drive1.Drive = Lastdrive
     Exit Sub
    End Select
   End If

End Sub

Private Sub File1_Click()
Reset ' reset all open disks
Pictoobig.visible = False
 
SelectedFile = File1.Path & "\" & File1.Filename
On Error GoTo cannotfindfileerr:
If FileLen(SelectedFile) > 65000 Then Pictoobig.visible = True: Exit Sub

Close #1

On Error GoTo cannotfindfileerr:

Open SelectedFile For Binary Access Read As #1
Txtfile.Text = Input(LOF(1), 1)
Close #1

cannotfindfileerr:
If Err.Number <> 0 Then
 On Error GoTo Erroutofmemory

   SelectedFile = File1.Path & File1.Filename
    Close #1
     Open SelectedFile For Binary Access Read As #1
      If FileLen(SelectedFile) > 32000 Then:  Pictoobig.visible = True: Exit Sub
       Txtfile.Text = Input(LOF(1), 1)
        Close #1
         Exit Sub
End If

Erroutofmemory:
    If Err.Number = 7 Then
      
      MsgBox "An Unexpected Error Has Occured" & vbNewLine & Err.Description, vbInformation, "NextPad"
       Exit Sub
    End If
End Sub



Private Sub Form_Load()
Lastdrive = Drive1.Drive
End Sub

Private Sub Form_Resize()
        ' If the form is minimized it cant be resized while it is minimized
        'so Exit the Sub Right away
        If Me.WindowState = vbMinimized Then: Exit Sub
        If Me.Height <= 5205 Then: Me.Height = 5400
        If Me.Width <= 8000 Then: Me.Width = 8200
           DoEvents
          ' Set the Text Boxes Height And width
          ' Set thhe Txtfile's Widht too The Forms Width Minus The PicHolderleft's width
          Txtfile.Width = Frmtxtfilemanager.ScaleWidth - Picholderleft.Width
           ' set the Txtfile's Height too the forms height Minus 300
           Txtfile.Height = Frmtxtfilemanager.ScaleHeight - 400
            ' Set the File List Boxes Height , width , Drive list Boxe's Height ,Width ;
            ' and set the Directory List boxes Height and width;
             File1.Height = Frmtxtfilemanager.ScaleHeight - 3900
               frameleft.Height = Frmtxtfilemanager.ScaleHeight - 400
                Pictoobig.Height = Txtfile.Height
                 Pictoobig.Width = Txtfile.Width
                  DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
Set Frmtxtfilemanager = Nothing ' release Memory Used By The Text File Manager
End Sub


Private Sub Mnuabout_Click()
   
   MsgBox "Text File Manager 1.0" & vbNewLine & vbNewLine & "Written By : Jason - Simeone " _
   & vbNewLine & vbNewLine & "Email Address : Cyberarea@Hotmail.com ", vbInformation, "About"

End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuopenwithexternaleditor_Click()
 Dim sFileName As String, regval As String
 
    sFileName = Dir1.Path & "\" & File1.Filename
     On Error GoTo FileDoesNotexist:
       FileDateTime (sFileName)
         GoTo toobig:
      
FileDoesNotexist:
        If Err.Number <> 0 Then
         sFileName = Dir1.Path & File1.Filename
           GoTo toobig:
        End If

toobig:
     regval = GetSetting("NextPad", "UseExternaleditor", "Path", "")

      If UseExternalEditor.use = True And regval = "" Then
No_Externaleditor_Detected: Unload Me
          Exit Sub
      End If

      Select Case UseExternalEditor.use
        
        Case True ' Case is True *******
             ExecuteExternalEditor (GetShortPath(sFileName))
             Close #1
              Reset
               Exit Sub
   
         Case False ' Case is False ******
          MsgBox sFileName _
           & vbNewLine & "External Editor Cannot Be Launched........" _
           & vbNewLine & vbNewLine & "Reason : Be Sure too have " & vbNewLine & Chr$(34) & "Use external editor When opening files too large for NextPad too open." & Chr$(34) & _
           vbNewLine & " Enabled in the options Dialog .", vbExclamation, "Error,Text File Manager"
             Close #1
                Reset
                  Exit Sub
        End Select
        

End Sub




Private Sub mnuopenwithNextpad_Click()
 Dim SelectedFile As String
 
    SelectedFile = Dir1.Path & "\" & File1.Filename
      On Error GoTo FileDoesNotexist:
       FileDateTime (SelectedFile)
       Openfile (SelectedFile)
       Form1.SetFocus
       Exit Sub
FileDoesNotexist:
         bootfile = Dir1.Path & File1.Filename
       Openfile (bootfile)
        Form1.SetFocus
     Exit Sub

End Sub

Private Sub mnuproperties_Click()
 Dim SelectedFile As String, bootfile As String
 
    SelectedFile = Dir1.Path & "\" & File1.Filename
      On Error GoTo FileDoesNotexist:
       FileDateTime (SelectedFile)
         ShowFileproperties Me.Hwnd, SelectedFile
         Exit Sub
FileDoesNotexist:
      If Err.Number <> 0 Then
         bootfile = Dir1.Path & File1.Filename
         ShowFileproperties Me.Hwnd, bootfile
      End If

End Sub

Private Sub mnustatusbar_Click()
   StatusBar1.visible = Not StatusBar1.visible
    mnustatusbar.Checked = StatusBar1.visible
   

End Sub

Private Sub Txtfile_Change()
 If Err Then
      MsgBox "An Unexpected Error Has Occured :" & vbNewLine & Err.Description, vbInformation, "NextPad"
       Exit Sub
 End If
End Sub

Private Sub Txtfile_KeyPress(KeyAscii As Integer)
MsgBox "You cannot Edit Files In the Text File Manager." & _
vbNewLine & "Too Open This File With NextPad : Click " & Chr(34) & "Open With NextPad" & Chr(34) & " In the File Menu...", vbInformation, "NextPad"

End Sub








