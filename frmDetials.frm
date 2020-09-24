VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cookie Details"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2640
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   5370
      Begin VB.TextBox txtWeb 
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   210
         Width           =   4320
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   525
         Width           =   4320
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtSource 
         Height          =   750
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1785
         Width           =   5160
      End
      Begin VB.ComboBox coSTATE 
         Height          =   315
         Left            =   945
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1155
         Width           =   1170
      End
      Begin VB.CommandButton cmdDecode 
         Caption         =   "Decode"
         Height          =   285
         Left            =   4410
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Domain:"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cookie Source:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   1575
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secure:"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   9
         Top             =   885
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   8
         Top             =   570
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   4095
      TabIndex        =   7
      Top             =   2730
      Width           =   1380
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDecode_Click()
    MsgBox "Note: Works only if the value is a query string." & vbCrLf & Try_Decode(C_VALUE), vbInformation, "Query String Decoder:"
End Sub

Private Sub cmdDecode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub coSTATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub



Private Sub Form_Load()
    coSTATE.AddItem "True"
    coSTATE.AddItem "False"
    txtWeb.Text = C_WEB
    txtName.Text = C_NAME
    txtValue.Text = C_VALUE
    txtSource.Text = C_SOURCE & "*" & vbCrLf
    
    If C_STATE = "False" Then coSTATE.ListIndex = 1
    If C_STATE = "True" Then coSTATE.ListIndex = 0
End Sub


'// Sametimes value of cookie is a portion of url; query string.
'// Special characters are encoded so the url is readable. This function tries
'// to decode the cookie value to a readable format.
'// EXAMPLE OF COOKIE VALUE:
'// Mastering+Maya+3%3A%7C%3Ahttp%3A%2F%2Fwww%2Ebridgemart%2Ecom%2Fclassifieds%2Fview%5Fad%2Easp%3FA%5FID%3D35%26ID%3D6%26cate%3DBooks%26P%3D8%26cate%5Fname%3DOther%2BBooks%0D%0A
'// RESULTS IN:
'// Mastering Maya 3:|:http://www.bridgemart.com/classifieds/view_ad.asp?A_ID=35&ID=6&cate=Books&P=8&cate_name=Other+Books
Function Try_Decode(ByRef SRC As String) As String
    SRC = Replace(SRC, "+", " ")
    SRC = Replace(SRC, "%40", "@")
    SRC = Replace(SRC, "%0D", " ")
    SRC = Replace(SRC, "%0A", " ")
    SRC = Replace(SRC, "%2E", ".")
    SRC = Replace(SRC, "%5F", "_")
    SRC = Replace(SRC, "%2D", "-")
    SRC = Replace(SRC, "%2B", "+")
    SRC = Replace(SRC, "%7C", "|")
    SRC = Replace(SRC, "%7E", "~")
    SRC = Replace(SRC, "%24", "$")
    SRC = Replace(SRC, "%25", "%")
    SRC = Replace(SRC, "%5E", "^")
    SRC = Replace(SRC, "%26", "&")
    SRC = Replace(SRC, "%28", "(")
    SRC = Replace(SRC, "%29", ")")
    SRC = Replace(SRC, "%3D", "=")
    SRC = Replace(SRC, "%60", "`")
    SRC = Replace(SRC, "%2F", "/")
    SRC = Replace(SRC, "%3F", "?")
    SRC = Replace(SRC, "%3C", "<")
    SRC = Replace(SRC, "%3E", ">")
    SRC = Replace(SRC, "%5C", "\")
    SRC = Replace(SRC, "%3A", ":")
    SRC = Replace(SRC, "%21", "!")
    SRC = Replace(SRC, "%5B", "[")
    SRC = Replace(SRC, "%5D", "]")
    SRC = Replace(SRC, "%22", Chr(34))
    Try_Decode = SRC
End Function

Private Sub Form_Paint()
    txtWeb.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txtSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txtWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
