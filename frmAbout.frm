VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About ..."
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   105
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   645
      ScaleWidth      =   3000
      TabIndex        =   5
      Top             =   105
      Width           =   3000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   330
      Left            =   4935
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4935
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is a Cookie ?"
      Height          =   2850
      Left            =   105
      TabIndex        =   1
      Top             =   1995
      Width           =   6315
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2535
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   4471
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         FileName        =   "C:\Documents and Settings\default\Desktop\Cookies\src\whatiscookie.txt"
         TextRTF         =   $"frmAbout.frx":25DA
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   4410
      Picture         =   "frmAbout.frx":31CE
      ScaleHeight     =   1860
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   105
      Width           =   1995
   End
   Begin VB.Label lblCopy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   945
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    T = "Copyright C 2002 Vladimir S. Pekulas, All Rights Reserved." & vbCrLf
    T = T & "All questions should be directed to: vpekulas@mts.net" & vbCrLf
    lblCopy.Caption = T
End Sub
