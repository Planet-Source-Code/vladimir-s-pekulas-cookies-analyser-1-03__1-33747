VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Cookie Analyser v1.0.3"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9570
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H0080C0FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtrep 
      Height          =   435
      Left            =   4620
      TabIndex        =   8
      Top             =   7245
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   767
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0ECA
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   8085
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":139E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D98
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar STBAR 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   7
      Top             =   6930
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "15/04/2002"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox FL_COOK 
      Height          =   870
      Left            =   4725
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   7875
      Visible         =   0   'False
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   540
      Left            =   315
      TabIndex        =   0
      Top             =   8085
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   953
      _Version        =   393217
      ScrollBars      =   3
      FileName        =   "C:\Documents and Settings\Administrator\Cookies\administrator@www.bridgemart[2].txt"
      TextRTF         =   $"Form1.frx":2139
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   150
      ScaleWidth      =   9840
      TabIndex        =   6
      Top             =   3360
      Width           =   9840
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cookie File Analysed"
      Height          =   3165
      Left            =   105
      TabIndex        =   4
      Top             =   3570
      Width           =   9360
      Begin MSComctlLib.ListView Lst_Cook 
         Height          =   2640
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain  "
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cookie Name  "
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cookie Value  "
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Secure  "
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Source  "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Parent  "
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Cookie Files"
      Height          =   3165
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   9360
      Begin MSComctlLib.ListView lst_all_cookies 
         Height          =   2640
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size (B)"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnusettings 
         Caption         =   "&Settings"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuViewStyle 
      Caption         =   "&View Style"
      Begin VB.Menu mnuStyleOne 
         Caption         =   "File && Cookies"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAllCookies 
         Caption         =   "All Cookies"
      End
   End
   Begin VB.Menu mnuMabout 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteCookieFile 
         Caption         =   "&Delete Cookie File"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh List"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuCookiePopUp 
      Caption         =   "mnuCookiePopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteThisCookie 
         Caption         =   "&Delete Cookie"
      End
      Begin VB.Menu mnuSeeCookieDetails 
         Caption         =   "&See Details"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Cookies Profile"
      End
      Begin VB.Menu mnuQuickStats 
         Caption         =   "&Quick Cookies Stats "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// EXPIRATION OF COOKIE:
'// Ff anyone knows what format of time is presented in cookies please let me know.
Public XX As Integer

'// SET INITIAL VARIABLE
Private Sub Form_Load()
On Error Resume Next
'// RESTORE LAST KNOW POSITION
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)

'// SHOW SETTINGS
    REMOVE_WWW = mfncGetFromIni("GENERAL", "remwww", App.PATH & "\custom.ini")
    P_TOP = mfncGetFromIni("GENERAL", "ptop", App.PATH & "\custom.ini")
    V_STYLE = mfncGetFromIni("GENERAL", "vstyle", App.PATH & "\custom.ini")
    SET_GF = mfncGetFromIni("GENERAL", "gf", App.PATH & "\custom.ini")
    SET_GC = mfncGetFromIni("GENERAL", "gc", App.PATH & "\custom.ini")

    If V_STYLE = "1" Then
        mnuStyleOne.Checked = True
        mnuAllCookies.Checked = False
        mnuQuickStats.Enabled = False
    Else
        mnuStyleOne.Checked = False
        mnuAllCookies.Checked = True
        mnuQuickStats.Enabled = True
    End If
    
    Picture1.Top = P_TOP
    Call Picture1_MouseUp(1, 0, 12, 12)
    
    SHOW_SET = mfncGetFromIni("GENERAL", "show", App.PATH & "\custom.ini")
    If SHOW_SET = "1" Then frmSettings.Show 1, Me
    RESIZE_NOW = True
    Call Apply_settings

'// DEFAULT PATH
    DEFAULT_PATH = mfncGetFromIni("GENERAL", "path", App.PATH & "\custom.ini")
    COOKIE_PATH = DEFAULT_PATH
    STBAR.Panels(2).Text = "No cookie file selected."
    Call Start_Work
End Sub


'// SHOW COOKIES ETC.
Sub Start_Work()
On Error Resume Next
  FL_COOK.Refresh
  lst_all_cookies.Sorted = False
  lst_all_cookies.ListItems.Clear
  If InStr(1, LCase(COOKIE_PATH), "cookie") > 1 Then
    FL_COOK.PATH = COOKIE_PATH
    Call Get_Cookie_files(lst_all_cookies, FL_COOK)
    Me.Caption = "Cookie Analyser v1.0.3"
  Else
    MsgBox "Not an MSIE cookie folder." & vbCrLf & "Select 'File -> Settings' and choose an MSIE cookie folder." & vbCrLf & "Exmple:" & vbCrLf & "C:\Documents and Settings\Administrator\Cookies", vbCritical, "Wrong Path ..."
  End If
      lst_all_cookies.Sorted = True
End Sub

'// GET INDIVIDUAL COOKIES OUT OF COOKIE FILE
Function Analyse_Cookie(ByRef Parent As String)
On Error Resume Next
Dim FULL_SOURCE As String
    FULL_SOURCE = rtf.Text
    UNQ_COOK = Split(FULL_SOURCE, "*")
    For I = 0 To UBound(UNQ_COOK)
       If Not Len(Replace(UNQ_COOK(I), vbCrLf, "")) = 0 Then
            Call DISASSEMBLE_COOKIE(Replace(Trim(UNQ_COOK(I)), vbCrLf, "|~|"), Parent)
       End If
    Next I
End Function


'// ADD DISASSEMBLED COOKIE TO LIST
'Function DISASSEMBLE_COOKIE(ByRef COOK_SRC As String, Parent As String)
'On Error Resume Next
    'XX = XX + 1
    'If Mid(COOK_SRC, 1, 3) = "|~|" Then COOK_SRC = Mid(COOK_SRC, 4, Len(COOK_SRC))
    'UNQ_VALUE = Split(COOK_SRC, "|~|")
    'If UBound(UNQ_VALUE) = 8 Then
        'FILE_NAME = UNQ_VALUE(2)
        'If REMOVE_WWW = "1" Then
            'FILE_NAME = Replace(LCase(FILE_NAME), "www.", "")
        'End If
        'Lst_Cook.ListItems.Add XX, , FILE_NAME, , 3
        'Lst_Cook.ListItems(XX).SubItems(1) = UNQ_VALUE(0)
        'Lst_Cook.ListItems(XX).SubItems(2) = UNQ_VALUE(1)
        'If UNQ_VALUE(3) = 0 Then SEC_I = "False"
        'If UNQ_VALUE(3) = 1 Then SEC_I = "True"
        'Lst_Cook.ListItems(XX).SubItems(3) = SEC_I
        'Lst_Cook.ListItems(XX).SubItems(4) = Replace(COOK_SRC, "|~|", vbCrLf)
        'Lst_Cook.ListItems(XX).SubItems(5) = Parent
    'End If
'End Function

'// Thanks to Michael Doering for his help
Function DISASSEMBLE_COOKIE(ByRef COOK_SRC As String, Parent As String)
'On Error Resume Next
Dim SEC_I As String
XX = XX + 1
'If Mid(COOK_SRC, 1, 3) = "|~|" Then COOK_SRC = Mid(COOK_SRC, 4, Len(COOK_SRC))
If Mid(COOK_SRC, 1, 3) = "|~|" Then
    COOK_SRC = Mid(COOK_SRC, 4, Len(COOK_SRC))
Else
    msarrUnqValue = Split(COOK_SRC, "|~|")
    If InStr(1, COOK_SRC, vbLf, 0) Then
        msarrUnqValue = Split(COOK_SRC, vbLf, -1, 0)
    End If
    '...
    If UBound(msarrUnqValue) < 8 Then Exit Function
    If Len(msarrUnqValue(2)) = 0 Then Exit Function
    '...
    FILE_NAME = msarrUnqValue(2)
    If REMOVE_WWW = "1" Then
        FILE_NAME = Replace(LCase(FILE_NAME), "www.", "")
    End If
    '...
    Lst_Cook.ListItems.Add XX, , FILE_NAME, , 3
    Lst_Cook.ListItems(XX).SubItems(1) = msarrUnqValue(0)
    Lst_Cook.ListItems(XX).SubItems(2) = msarrUnqValue(1)
    If msarrUnqValue(3) = 0 Then SEC_I = "False"
    If msarrUnqValue(3) = 1 Then SEC_I = "True"
    Lst_Cook.ListItems(XX).SubItems(3) = SEC_I
    Lst_Cook.ListItems(XX).SubItems(4) = Replace(COOK_SRC, "|~|", vbCrLf)
    Lst_Cook.ListItems(XX).SubItems(5) = Parent
End If
End Function


'// RESIZE CONTROLS ON FORM
Sub Form_Resize()
On Error Resume Next
    If Not Me.WindowState = 1 Then
        If V_STYLE = "2" Then          '// COOKIES ONLY
           Frame1.Visible = False
           Picture1.Visible = False
           Frame2.Top = Frame1.Top
           Frame2.Height = Me.Height - STBAR.Height - 920
           Frame2.Width = Me.Width - 350
           Lst_Cook.Width = Me.Width - 550
           Lst_Cook.Height = Frame2.Height - 400
           STBAR.Panels(2).Width = (Me.Width - STBAR.Panels(1).Width) - 350
           For C = 1 To 5
               CL_W = CL_W + Lst_Cook.ColumnHeaders(C).Width
           Next C
           Lst_Cook.ColumnHeaders(6).Width = (Lst_Cook.Width - CL_W) - 330
           Call List_All_Cookies_AtOnce
           'STBAR.Panels(2).Text = "No cookie file selected." & "There are " & Lst_Cook.ListItems.Count & " individual in cookies folder."
        Else                           '// BOTH FILES & COOKIES
           Picture1.Top = P_TOP
           Call Picture1_MouseUp(1, 0, 12, 12)
           Frame1.Visible = True
           Picture1.Visible = True
           If RESIZE_NOW = True Then Lst_Cook.ListItems.Clear
           If Me.Height < 5500 Then Me.Height = 5500
           Frame1.Width = Me.Width - 350
           lst_all_cookies.Width = Me.Width - 550
           Frame2.Width = Me.Width - 350
           Lst_Cook.Width = Me.Width - 550
'           lst_all_cookies.ColumnHeaders(4).Width = (lst_all_cookies.Width - lst_all_cookies.ColumnHeaders(1).Width - lst_all_cookies.ColumnHeaders(2).Width - lst_all_cookies.ColumnHeaders(3).Width) - 300
           Frame2.Height = Me.Height - Frame1.Height - 1200
           Lst_Cook.Height = Frame2.Height - 400
           Picture1.Width = Me.Width + 3000
           STBAR.Panels(2).Width = (Me.Width - STBAR.Panels(1).Width) - 350
           For C = 1 To 5
               CL_W = CL_W + Lst_Cook.ColumnHeaders(C).Width
           Next C
           Lst_Cook.ColumnHeaders(6).Width = (Lst_Cook.Width - CL_W) - 110
        End If
    End If
End Sub

'// SAVE RESIZER POSITION AND WINDOW POSITION
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call mfncWriteIni("GENERAL", "ptop", Picture1.Top, App.PATH & "\custom.ini")
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "ViewMode", "vsp"
End Sub

Private Sub lst_all_cookies_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
   lst_all_cookies.SortKey = ColumnHeader.Index - 1
   
   If lst_all_cookies.SortOrder = lvwAscending Then
    lst_all_cookies.SortOrder = lvwDescending
   Else
    lst_all_cookies.SortOrder = lvwAscending
   End If
End Sub

'// ANALISY THE SELECTED COOKIE FILE
Private Sub lst_all_cookies_DblClick()
On Error Resume Next
  If Not lst_all_cookies.ListItems.Count = 0 Then
    FILE_NAME = lst_all_cookies.SelectedItem.ListSubItems(1).Text
    OPENED_FILE = FILE_NAME
    If Not V_STYLE = "2" Then Frame2.Caption = "Analysing Cookie file: " & FILE_NAME
    STBAR.Panels(2).Text = "Analysing Cookie file: " & FILE_NAME
    rtf.LoadFile COOKIE_PATH & "\" & FILE_NAME
    Lst_Cook.ListItems.Clear
    XX = 0
    Lst_Cook.Sorted = False
        Call Analyse_Cookie(OPENED_FILE)
    Lst_Cook.Sorted = True
    RESIZE_NOW = False
  End If
End Sub

'// Open cookie file on enter
Private Sub lst_all_cookies_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 18 Then Call Start_Work
    If Not lst_all_cookies.ListItems.Count = 0 Then
        If KeyAscii = 13 Then
            Call lst_all_cookies_DblClick
        End If
    End If
End Sub

'// SHOW POP-UP MENU
Private Sub lst_all_cookies_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
 If Not lst_all_cookies.ListItems.Count = 0 Then
    If Button = 2 Then
        mnuDeleteCookieFile.Caption = "Delete cookie file: " & lst_all_cookies.SelectedItem.SubItems(1) & " ?"
        Me.PopupMenu mnuPopUp
    End If
 End If
End Sub

'// SORT WHEN REQUESTED
Private Sub Lst_Cook_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
Dim CLMN_NAME As String
   Lst_Cook.SortKey = ColumnHeader.Index - 1
 Select Case ColumnHeader.Index
 Case 1
        CLMN_NAME = "Domain"
 Case 2
        CLMN_NAME = "Cookie Name"
 Case 3
        CLMN_NAME = "Cookie Value"
 Case 4
        CLMN_NAME = "Secure"
 Case 5
        CLMN_NAME = "Source"
 Case 6
        CLMN_NAME = "Parent"
 End Select
   
 If Lst_Cook.SortOrder = lvwAscending Then
    Lst_Cook.SortOrder = lvwDescending
    ColumnHeader.Text = CLMN_NAME & " +"
 Else
    Lst_Cook.SortOrder = lvwAscending
    ColumnHeader.Text = CLMN_NAME & " -"
 End If
   
 For E = 1 To 6
   If Not E = ColumnHeader.Index Then
        HDR_TITLE = Mid(Lst_Cook.ColumnHeaders(E).Text, Len(Lst_Cook.ColumnHeaders(E).Text) - 1, Len(Lst_Cook.ColumnHeaders(E).Text))
        If HDR_TITLE = " +" Then Lst_Cook.ColumnHeaders(E).Text = Mid(Lst_Cook.ColumnHeaders(E).Text, 1, Len(Lst_Cook.ColumnHeaders(E).Text) - 1)
        If HDR_TITLE = " -" Then Lst_Cook.ColumnHeaders(E).Text = Mid(Lst_Cook.ColumnHeaders(E).Text, 1, Len(Lst_Cook.ColumnHeaders(E).Text) - 1)
   End If
 Next E
End Sub

'// SHOW DETAILS OF COOKIE IN POP-UP
Private Sub Lst_Cook_DblClick()
On Error Resume Next
  If Not Lst_Cook.ListItems.Count = 0 Then
    C_WEB = Lst_Cook.SelectedItem.Text
    C_NAME = Lst_Cook.SelectedItem.SubItems(1)
    C_VALUE = Lst_Cook.SelectedItem.SubItems(2)
    C_STATE = Lst_Cook.SelectedItem.SubItems(3)
    C_SOURCE = Lst_Cook.SelectedItem.SubItems(4)
    frmDetails.Show 1, Me
  End If
End Sub

'// OPEN COOKIE DETAILS ON ENTER PRESS
Private Sub Lst_Cook_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not Lst_Cook.ListItems.Count = 0 Then
        If KeyAscii = 13 Then
           Call Lst_Cook_DblClick
        End If
    End If
End Sub

'// SHOW POP-UP MENU FOR COOKIES
Private Sub Lst_Cook_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
 If Not Lst_Cook.ListItems.Count = 0 Then
    If Button = 2 Then
        
        If V_STYLE = "2" Then
            mnuDeleteThisCookie.Enabled = False
            mnuSeeCookieDetails.Enabled = False
        End If
        
        If Lst_Cook.ListItems.Count = 0 Then
            mnuDeleteThisCookie.Enabled = False
            mnuSeeCookieDetails.Enabled = False
        Else
            mnuDeleteThisCookie.Enabled = True
            mnuSeeCookieDetails.Enabled = True
        End If
        mnuDeleteThisCookie.Caption = "Delete cookie: " & UCase(Lst_Cook.SelectedItem.SubItems(1))
        Me.PopupMenu mnuCookiePopUp
    End If
 End If
End Sub

'// Show about pop-up
Private Sub mnuAbout_Click()
    frmAbout.Show 1, Me
End Sub

Sub mnuAllCookies_Click()
    V_STYLE = "2"
    mnuAllCookies.Checked = True
    mnuStyleOne.Checked = False
    Call Form_Resize
    Call List_All_Cookies_AtOnce
    STBAR.Panels(2).Text = "All Cookies Mode"
    Frame2.Caption = "All Cookies available"
    mnuQuickStats.Enabled = True
End Sub

'// DELETE ENTIRE COOKIE FILE
Private Sub mnuDeleteCookieFile_Click()
On Error Resume Next
    COOKIE_FILE = lst_all_cookies.SelectedItem.SubItems(1)
    If MsgBox("Delete cookie file: " & COOKIE_FILE & " ?", vbOKCancel, "Delete File ?") = vbOK Then
        Kill COOKIE_PATH & "\" & COOKIE_FILE
        Lst_Cook.ListItems.Clear
        Call Start_Work
    End If
End Sub

'// EXIT MENU
Private Sub mnuExit_Click()
    If MsgBox("Are you sure you want to exit ?", vbOKCancel, "Exit...") = vbOK Then
        Unload Me
    End If
End Sub

Private Sub mnuExport_Click()
    Call Export_Coookies_Profile(Lst_Cook, EXPORT_PATH)
End Sub

Private Sub mnuQuickStats_Click()
    frmStats.Show 1, Me
End Sub

'// REFRESH FILE LIST
Private Sub mnuRefresh_Click()
    Call Start_Work
End Sub

'// SHOW SETTING WINDOW
Private Sub mnusettings_Click()
    frmSettings.Show 1, Me
End Sub


Sub mnuStyleOne_Click()
    V_STYLE = "1"
    mnuAllCookies.Checked = False
    mnuStyleOne.Checked = True
    RESIZE_NOW = True
    Call Form_Resize
    Frame2.Caption = "No cookie file selected"
    STBAR.Panels(2).Text = "No cookie file selected"
    mnuQuickStats.Enabled = False
End Sub

'// RESIZING FUNCTIONS ETC.
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
    If Button = 1 Then Picture1.Top = Picture1.Top + y
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Picture1.BackColor = &H0&
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim pt&
pt = Picture1.Top
    If pt < 1500 Then
        pt = 1000
        Picture1.Top = pt
    End If
    If pt > (Me.ScaleHeight - 1500) Then
        pt = Me.ScaleHeight - 1500
        Picture1.Top = pt
    End If
    Frame1.Height = pt - Frame1.Top - 5
    lst_all_cookies.Height = Frame1.Height - 450
    Frame2.Top = pt + 190
    Frame2.Height = Me.Height - Frame2.Top - 1000
    Lst_Cook.Height = Frame2.Height - 450
    Picture1.BackColor = &H8000000F
End Sub


'// Delete selected individual cookie only - V_STYLE = "1" only mode
Private Sub mnuDeleteThisCookie_Click()
On Error Resume Next
Dim SEL_INDEX As Integer, COOK_COUNT As Integer, CONTENT As String
  If MsgBox("Delete Cookie: " & UCase(Lst_Cook.SelectedItem.SubItems(1)) & " ?", vbOKCancel, "Delete Cookie ?") = vbOK Then
    If V_STYLE = "2" Then
       Call DeleteCookie_other(COOKIE_PATH & "\" & Lst_Cook.SelectedItem.SubItems(5))
       Lst_Cook.ListItems.Remove (Lst_Cook.SelectedItem.Index)
       Exit Sub
    End If
    CONTENT = ""
    COOK_COUNT = Lst_Cook.ListItems.Count
    If Not COOK_COUNT = 0 Then
        SEL_INDEX = Lst_Cook.SelectedItem.Index
            For C_I = 1 To COOK_COUNT
                If Not C_I = SEL_INDEX Then
                    CONTENT = CONTENT & Lst_Cook.ListItems(C_I).SubItems(4) & "*" & vbCrLf
                End If
            Next C_I
    End If
    Call SaveFile(COOKIE_PATH & "\" & OPENED_FILE, CONTENT)
    Lst_Cook.ListItems.Remove (SEL_INDEX) 'Remove item from list_view
    STBAR.Panels(2).Text = "No cookie file selected." & "There are " & Lst_Cook.ListItems.Count & " individual in cookies folder."
  End If
End Sub


'// Delete selected individual cookie only - V_STYLE = "2" only mode
Function DeleteCookie_other(ByRef SEL_FILE_PATH As String)
Dim REPLACEWITH As String
    txtrep.Text = ""
    txtrep.FileName = SEL_FILE_PATH
    REPLACEWITH = Lst_Cook.SelectedItem.SubItems(4) & "*" & vbCrLf
    txtrep.Text = Replace(txtrep.Text, REPLACEWITH, "")
    Call SaveFile(SEL_FILE_PATH, Trim(txtrep.Text))
End Function




'// See detials of cookie
Private Sub mnuSeeCookieDetails_Click()
    If Not Lst_Cook.ListItems.Count = 0 Then
        Call Lst_Cook_DblClick
    End If
End Sub

Function List_All_Cookies_AtOnce()
 On Error GoTo Error:
    Lst_Cook.ListItems.Clear
    For W = 1 To lst_all_cookies.ListItems.Count
        FILE_NAME = lst_all_cookies.ListItems(W).SubItems(1)
        OPENED_FILE = FILE_NAME
        rtf.LoadFile COOKIE_PATH & "\" & FILE_NAME
        XX = 0
        Lst_Cook.Sorted = False
            Call Analyse_Cookie(OPENED_FILE)
        Lst_Cook.Sorted = True
    Next W

Error:
If Err.Number = 75 Then MsgBox "File sharing violation !" & vbCrLf & "Close all browsers!", vbCritical, "Error)"

End Function

