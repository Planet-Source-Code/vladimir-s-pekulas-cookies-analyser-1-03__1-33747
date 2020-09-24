VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings ..."
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3345
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chNEXT 
      Caption         =   "Show Next Time"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   4253
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1680
      TabIndex        =   4
      Top             =   4200
      Width           =   1590
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "  Path  "
      TabPicture(0)   =   "frmSettings.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPath"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "  General  "
      TabPicture(1)   =   "frmSettings.frx":0324
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chWWW"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chGF"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chGC"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optStyle"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optStyle1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtexport"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtexport 
         Height          =   285
         Left            =   -74790
         TabIndex        =   14
         Top             =   3045
         Width           =   2745
      End
      Begin VB.OptionButton optStyle1 
         Caption         =   "View All Cookies Only"
         Height          =   225
         Left            =   -74790
         TabIndex        =   11
         Top             =   2205
         Width           =   2010
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "View Files && Cookies"
         Height          =   225
         Left            =   -74790
         TabIndex        =   10
         Top             =   1890
         Width           =   2010
      End
      Begin VB.CheckBox chGC 
         Caption         =   "Gridlines on cookies listing"
         Height          =   225
         Left            =   -74790
         TabIndex        =   9
         Top             =   1260
         Width           =   2850
      End
      Begin VB.CheckBox chGF 
         Caption         =   "Gridlines on file listing"
         Height          =   225
         Left            =   -74790
         TabIndex        =   8
         Top             =   945
         Width           =   2850
      End
      Begin VB.CheckBox chWWW 
         Caption         =   "Remove 'www.' from domain name"
         Height          =   225
         Left            =   -74790
         TabIndex        =   7
         Top             =   630
         Width           =   2850
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cookie Folder: "
         Height          =   3270
         Left            =   210
         TabIndex        =   1
         Top             =   525
         Width           =   2745
         Begin VB.CommandButton cmdFindPath 
            Caption         =   "&Find Cookies Path"
            Height          =   330
            Left            =   105
            TabIndex        =   12
            Top             =   2835
            Width           =   2535
         End
         Begin VB.DriveListBox Drive 
            Height          =   315
            Left            =   105
            TabIndex        =   3
            Top             =   210
            Width           =   2535
         End
         Begin VB.DirListBox Dir 
            Height          =   2115
            Left            =   105
            TabIndex        =   2
            Top             =   590
            Width           =   2535
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Export Path:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   13
         Top             =   2835
         Width           =   870
      End
      Begin VB.Label lblPath 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   3150
         Visible         =   0   'False
         Width           =   2850
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SHITEMID
SHItem As Long
itemID() As Byte
End Type
Private Type ITEMIDLIST
shellID As SHITEMID
End Type
Const DESKTOP = &H0
Const PROGRAMS = &H2
Const MYDOCS = &H5
Const FAVORITES = &H6
Const STARTUP = &H7
Const RECENT = &H8
Const SENDTO = &H9
Const STARTMENU = &HB
Const NETHOOD = &H13
Const FONTS = &H14
Const SHELLNEW = &H15
Const TEMPINETFILES = &H20
Const COOKIES = &H21
Const HISTORY = &H22
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long



Private Sub cmdFindPath_Click()
Dim PATH As String * 256
Dim myid As ITEMIDLIST
Dim rval As Long
rval = SHGetSpecialFolderLocation(Me.hwnd, COOKIES, myid)
If rval = 0 Then
    rval = SHGetPathFromIDList(ByVal myid.shellID.SHItem, ByVal PATH)
    If rval Then
        If MsgBox("Set this path as defaut ?" & vbCrLf & Left(PATH, InStr(PATH, Chr(0)) - 1), vbOKCancel, "Set as defaul ?") = vbOK Then
            Dir.PATH = Left(PATH, InStr(PATH, Chr(0)) - 1)
        End If
    End If
End If
End Sub

Private Sub Command1_Click()
    Unload Me
    SET_GC = chGC.Value
    SET_GF = chGF.Value
    Call Apply_settings
End Sub


Private Sub Dir_Change()
On Error Resume Next
    lblPath.Caption = Dir.PATH
    COOKIE_PATH = Dir.PATH
End Sub

Private Sub Drive_Change()
On Error Resume Next
    Dir.PATH = Drive.Drive
End Sub

Private Sub Form_Load()
On Error Resume Next
    DEFAULT_PATH = mfncGetFromIni("GENERAL", "path", App.PATH & "\custom.ini")
    SHOW_SET = mfncGetFromIni("GENERAL", "show", App.PATH & "\custom.ini")
    REMOVE_WWW = mfncGetFromIni("GENERAL", "remwww", App.PATH & "\custom.ini")
    SET_GF = mfncGetFromIni("GENERAL", "gf", App.PATH & "\custom.ini")
    SET_GC = mfncGetFromIni("GENERAL", "gc", App.PATH & "\custom.ini")
    V_STYLE = mfncGetFromIni("GENERAL", "vstyle", App.PATH & "\custom.ini")
    EXPORT_PATH = mfncGetFromIni("GENERAL", "export", App.PATH & "\custom.ini")
    
    Dir.PATH = DEFAULT_PATH
    chNEXT.Value = Int(SHOW_SET)
    chWWW.Value = REMOVE_WWW
    txtexport.Text = EXPORT_PATH
    
    If V_STYLE = 2 Then optStyle1.Value = True
    If V_STYLE = 1 Then optStyle.Value = True
    
    
    chGC.Value = SET_GC
    chGF.Value = SET_GF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mfncWriteIni("GENERAL", "show", chNEXT.Value, App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "path", Dir.PATH, App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "remwww", chWWW.Value, App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "gc", chGC.Value, App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "gf", chGF.Value, App.PATH & "\custom.ini")
    EXPORT_PATH = txtexport.Text
    Call mfncWriteIni("GENERAL", "export", EXPORT_PATH, App.PATH & "\custom.ini")
    
    
    If optStyle1.Value = True Then V_STYLE = 2
    If optStyle.Value = True Then V_STYLE = 1
    
    Call mfncWriteIni("GENERAL", "vstyle", V_STYLE, App.PATH & "\custom.ini")
    If V_STYLE = "1" Then
        frmMain.mnuStyleOne.Checked = True
        frmMain.mnuAllCookies.Checked = False
        Call frmMain.mnuStyleOne_Click
    Else
        frmMain.mnuStyleOne.Checked = False
        frmMain.mnuAllCookies.Checked = True
        Call frmMain.mnuAllCookies_Click
    End If
    
    Call frmMain.Start_Work
    REMOVE_WWW = chWWW.Value
    COOKIE_PATH = Dir.PATH
End Sub
