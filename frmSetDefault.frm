VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSetDefault 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MSIE Cookie folder:"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdc 
      Left            =   420
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3780
      TabIndex        =   3
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2625
      TabIndex        =   2
      Top             =   945
      Width           =   1485
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select path to MSIE cookie folder:"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   2430
   End
End
Attribute VB_Name = "frmSetDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
    cdc.DialogTitle = "Select folder ..."
    cdc.ShowOpen
    txtPath.Text = cdc.FileName
End Sub
