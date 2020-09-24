VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStats 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Stats ..."
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   330
      Left            =   6090
      TabIndex        =   5
      Top             =   4620
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   4530
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   7260
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   105
         Picture         =   "frmStats.frx":0000
         ScaleHeight     =   360
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   210
         Width           =   270
      End
      Begin MSComctlLib.ListView lstStats 
         Height          =   3165
         Left            =   210
         TabIndex        =   1
         Top             =   735
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   5583
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "# of Cookies"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size (B)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblTotal 
         Height          =   435
         Left            =   210
         TabIndex        =   4
         Top             =   3990
         Width           =   6840
      End
      Begin VB.Label Label1 
         Caption         =   "Cookies Quick Stats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   3
         Top             =   315
         Width           =   3480
      End
      Begin VB.Line Line1 
         X1              =   525
         X2              =   7035
         Y1              =   585
         Y2              =   585
      End
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TOTAL_COOKIES_SIZE As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    TOTAL_COOKIES_SIZE = 0
    Call Create_Stats(frmMain.Lst_Cook, Int(frmMain.Lst_Cook.ListItems.Count))
    Call Get_Num
    '// Size if divided by 1000 instead of 1024 becuase of windows HDD partition
    lblTotal.Caption = lblTotal.Caption & vbCrLf & "Total space occupied by cookie files: " & CCur((TOTAL_COOKIES_SIZE / 1000)) & " KB"
End Sub

Function Get_Num()
On Error Resume Next
Dim C As Integer, X As Integer, I As Integer, FLN As String
  For X = 1 To lstStats.ListItems.Count '// for each unique domain name
    C = 0 '// set counter to 0
    For I = 1 To frmMain.Lst_Cook.ListItems.Count '// for each cookie
        If Trim(LCase(frmMain.Lst_Cook.ListItems(I).Text)) = LCase(Trim(lstStats.ListItems(X).Text)) Then
            C = C + 1
            FLN = frmMain.Lst_Cook.ListItems(I).SubItems(5)
        End If
    Next I
    FLC_SIZE = Get_File_Size(COOKIE_PATH & "\" & FLN)
    TOTAL_COOKIES_SIZE = TOTAL_COOKIES_SIZE + Int(FLC_SIZE)
    lstStats.ListItems(X).SubItems(1) = C
    lstStats.ListItems(X).SubItems(2) = FLC_SIZE
  Next X
End Function

Function Create_Stats(ByRef LST As ListView, TOTAL As Integer)
On Error Resume Next
 lblTotal.Caption = "There is total of " & TOTAL & " cookies saved to your HDD."
    For I = 1 To TOTAL
       Call Check_Agains(LST.ListItems(I).Text)
    Next I
End Function

Function Check_Agains(ByRef DOMS As String)
On Error Resume Next
Dim FOUND As Boolean
FOUND = False
    For X = 1 To lstStats.ListItems.Count
        If Trim(LCase(lstStats.ListItems(X).Text)) = Trim(LCase(DOMS)) Then
            FOUND = True
        End If
    Next X
    If FOUND = False Then
        lstStats.ListItems.Add , , DOMS
    End If
End Function





Private Sub lstStats_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstStats.Sorted = False
lstStats.SortKey = ColumnHeader.Index - 1

If lstStats.SortOrder = lvwAscending Then
    lstStats.SortOrder = lvwDescending
 Else
    lstStats.SortOrder = lvwAscending
 End If
lstStats.Sorted = True
End Sub
