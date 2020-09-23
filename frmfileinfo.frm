VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmfileinfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "frmfileinfo.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   240
      ScaleHeight     =   4095
      ScaleWidth      =   5175
      TabIndex        =   10
      Top             =   480
      Width           =   5175
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   120
         Picture         =   "frmfileinfo.frx":000C
         ScaleHeight     =   135
         ScaleWidth      =   4815
         TabIndex        =   19
         Top             =   3000
         Width           =   4815
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   120
         Picture         =   "frmfileinfo.frx":024E
         ScaleHeight     =   135
         ScaleWidth      =   4815
         TabIndex        =   18
         Top             =   2160
         Width           =   4815
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   120
         Picture         =   "frmfileinfo.frx":0490
         ScaleHeight     =   135
         ScaleWidth      =   4815
         TabIndex        =   17
         Top             =   840
         Width           =   4815
      End
      Begin VB.CheckBox chcksystem 
         Caption         =   "&System"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chckarchive 
         Caption         =   "A&rchive"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Lblfiledate 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Lblfiledirectory 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox Lblfilesize 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox Txtlocation 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   4095
      End
      Begin VB.CheckBox chckreadonly 
         Caption         =   "&Read Only"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chckhidden 
         Caption         =   "&Hidden"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Image imgloadicon 
         Height          =   615
         Left            =   0
         Top             =   120
         Width           =   495
      End
      Begin VB.Image imgini 
         Height          =   480
         Left            =   480
         Picture         =   "frmfileinfo.frx":06D2
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgunknown 
         Height          =   480
         Left            =   0
         Picture         =   "frmfileinfo.frx":118C
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgtxt 
         Height          =   480
         Left            =   960
         Picture         =   "frmfileinfo.frx":1C46
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Size :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Location :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblfiletype 
         Caption         =   "Unknown File Type"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Created/modified :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Attributes :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   855
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8070
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmfileinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub Form_Load()
FileattrChange = False
If Form1.lblfilename.caption = "" Then GoTo Fileinfo_Nofilename:
On Error GoTo Fileerror:
Dim filename As String ' declare string variables
Dim filelocation As String ' declare string variables
Dim DetectExtension As String
filename = Form1.lblfilename.caption
'perform Bitwize Comparison too check File attributes
filereadonlyattrs = GetAttr(filename) And vbReadOnly
filehiddenattrs = GetAttr(filename) And vbHidden
filesystemattrs = GetAttr(filename) And vbSystem
filearchiveattrs = GetAttr(filename) And vbArchive


If filename <> "" Then
Me.caption = LCase$(Dir(filename)) & " Properties"

DetectExtension = DetectFextension(Form1.lblfilename.caption)
 
Select Case DetectExtension
Case "TXT" ' if Extension is TXT
lblfiletype.caption = "Text Document"
imgloadicon.Picture = imgtxt.Picture ' load TXT icon
Case "INI" ' if extension is INI
lblfiletype.caption = "INI Configuration Settings File"
imgloadicon = imgini.Picture ' load INI icon
Case "EXE" ' if extension is EXE
lblfiletype.caption = "Executable File"
imgloadicon.Picture = imgunknown.Picture ' load UNKNOWN icon
Case "LOG"   ' if extension is LOG
lblfiletype.caption = "LOG File"
imgloadicon.Picture = imgunknown.Picture ' load UNKNOWN icon
Case Else  ' If The Extension is NULL ""
lblfiletype.caption = "Unknown Type"
imgloadicon.Picture = imgunknown.Picture ' load UNKNOWN icon
End Select

'If DetectExtension = "TXT" Then ' if extension is TXT
'lblfiletype.caption = "Text Document"
'ElseIf DetectExtension = "INI" Then ' if extension is INI
'lblfiletype.caption = "INI Configuration Settings File"

'ElseIf DetectExtension = "EXE" Then ' if extensio is EXE
'lblfiletype.caption = "Executable File"
'ElseIf DetectExtension = "LOG" Then ' if extension is LOG
'lblfiletype.caption = "LOG File"
'ElseIf DetectExtension = "" Then ' If The Extension is NULL ""
'lblfiletype.caption = "Unknown Type"
'End If


Lblfilesize.Text = FileLen(filename) & " Bytes"

Lblfiledirectory.Text = UCase$(Dir(filename))

date_Expression = FileDateTime(filename)

retformat = Format(date_Expression, "dddd, mmmm d , yyyy   hh:mm:ss AMPM")
Lblfiledate.Text = retformat

Txtlocation.Text = CurDir(filename)

If filearchiveattrs = 32 Then
chckarchive.Value = vbChecked
Else
chckarchive.Value = vbUnchecked
End If

If filehiddenattrs = 2 Then
chckhidden.Value = vbChecked
Else
chckhidden.Value = vbUnchecked
End If

If filesystemattrs = 4 Then
chcksystem.Value = vbChecked
Else
chcksystem.Value = vbUnchecked
End If

If filereadonlyattrs = 1 Then
chckreadonly.Value = vbChecked
Else
chckreadonly.Value = vbUnchecked
End If

Fileerror:
If Err.Number <> 0 Then
MsgBox "An unexpected error has occured: " _
& vbCr & "The file May have been Renamed moved or deleted. " _
& vbCr & "Description : " & Err.Description, vbCritical, "Error "
Exit Sub
End If
End If

Fileinfo_Nofilename: '\\ no filename detected
If filename = "" Then
Me.caption = "Untitled File" & " Properties"
Lblfilesize.Text = "Unknown"
Lblfiledirectory.Text = "Untitled File"
Lblfiledate.Text = "Unknown"
Txtlocation.Text = "Unknown"
imgloadicon.Picture = imgunknown.Picture
'\\ disable all check boxes
chckarchive.Enabled = False
chckhidden.Enabled = False
chckreadonly.Enabled = False
chckarchive.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmfileinfo = Nothing
End Sub

Private Sub Lblfiledirectory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Lblfiledirectory
.ToolTipText = Lblfiledirectory.Text
End With
End Sub

Private Sub Txtlocation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Txtlocation
.ToolTipText = Txtlocation.Text
End With
End Sub
