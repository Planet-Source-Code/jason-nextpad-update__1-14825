VERSION 5.00
Begin VB.Form FrmDumbyform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TextPad Setup "
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "FrmDumbyform.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   0
      Width           =   3255
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   2895
         Begin VB.Label Label1 
            Caption         =   "TextPad Is Now Finishing Setup... Please Wait......"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2655
         End
      End
   End
   Begin VB.TextBox TxtConvert2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   2895
   End
   Begin VB.TextBox TxtConvert1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmDumbyform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is only for Editing and Formatting strings ;
' the whole Purpose of this form has not been completed yet ,
' therefore this form MUST stay in the project or In modmain ;
' When Detecting an external editor , NextPad may crash ;
' therefore
' !!!!!!!!!!!!!!!!!!!!!!!!!!!
' PLEASE LEAVE THIS FORM IN THE PROJECT !!!!!
' !!!!!!!!!!!!!!!!!!!!!!!!!!!
