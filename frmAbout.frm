VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NextPad"
   ClientHeight    =   4335
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5310
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2992.095
   ScaleMode       =   0  'User
   ScaleWidth      =   4986.365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picabout 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   9
      Top             =   0
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Written By : Jason - Simeone"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Thanks Too ElitePad for its Registry Module."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please  Email me if you have any suggestions."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled on : Monday, January 29, 2001"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lbllicense 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is licensed too:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225.372
      X2              =   4845.507
      Y1              =   2101.714
      Y2              =   2101.714
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "A Small application for writing and saving text documents.  "
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   225.372
      X2              =   4845.507
      Y1              =   2112.067
      Y2              =   2112.067
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 4.120 Beta 24"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4125
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Email address : Cyberarea@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   3990
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
   
   Unload Me ' unload the form

End Sub
Private Sub Form_Load()

On Error GoTo Picerr: ' if an error occurs vb will jump too that line
  
    Dim license As String ' Declare variables
    Dim company As String ' Declare variables

    ' Set the license string too the setting (If any )
    ' Recieved Form the registry
    license = GetSettingString(HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion", _
    "RegisteredOwner", "")
     
    ' Set the company string too the setting (If any)
    ' Received form the registry
    company = GetSettingString(HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion", _
    "RegisteredOrganization", "")
   
    ' Set the Lbllicenses's Caption Property too the Company
    ' & License String Variables Recieved From the Registry
    lbllicense.caption = license & Chr(13) & Chr(10) & company

  On Error GoTo Picerr: ' if an error occurs vb will jump too that line

   Picabout.Picture = Frmnoreg.Picture1.Picture
' why Bring another hefty 62 kb along
' with us Just load the picture from frmnoreg

Picerr:
    If Err.Number <> 0 Then ' if the error's number is equal too anything above or other than zero then.....
     Exit Sub ' exit this sub immediately
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*************************************************
'This event is triggered when the form is unloaded
'*************************************************
      Set frmAbout.Picabout = Nothing 'set this picture to nothing too release memory it held
      Set frmAbout = Nothing ' set this form too nothing too release memory it held

End Sub


Private Sub Label5_Click()
     MsgBox "Special Thanks Goes out too the Creator And Developer of ElitePad for its Great Registry Module." + vbNewLine + vbNewLine + "Thank you", vbInformation, "Thanks"
End Sub
