Attribute VB_Name = "Module1"
Public COOKIE_PATH As String, REMOVE_WWW As String, V_STYLE As String
Public C_WEB As String, C_NAME As String, C_VALUE As String, C_STATE As String, C_SOURCE As String
Public OPENED_FILE As String, SET_GF As String, SET_GC As String
Public FILE_NAME As String, RESIZE_NOW As Boolean, EXPORT_PATH As String, P_TOP As Long

Function Get_Cookie_files(ByRef TARGET_LST As ListView, FILETBOX As FileListBox)
On Error Resume Next
    For I = 1 To FILETBOX.ListCount Step 1
       If Not Len(Trim(FILETBOX.List(I))) = 0 Then
         XX = I
         TARGET_LST.ListItems.Add XX, , Get_Site_Only(FILETBOX.List(I)), , 4
         TARGET_LST.ListItems(XX).SubItems(1) = FILETBOX.List(I)
         TARGET_LST.ListItems(XX).SubItems(2) = Get_File_Date(COOKIE_PATH & "\" & FILETBOX.List(I))
         TARGET_LST.ListItems(XX).SubItems(3) = Get_File_Size(COOKIE_PATH & "\" & FILETBOX.List(I))
       End If
    Next I
End Function

Function Get_Site_Only(ByRef FILE_NAME As String) As String
On Error Resume Next
    FILE_NAME = Mid(FILE_NAME, InStr(1, FILE_NAME, "@") + 1, Len(FILE_NAME))
    FILE_NAME = Mid(FILE_NAME, 1, InStr(1, FILE_NAME, ".txt") - 1)
    '// remove 'www.' from domain name if desired
    If REMOVE_WWW = "1" Then
        FILE_NAME = Replace(LCase(FILE_NAME), "www.", "")
    End If
    Get_Site_Only = FILE_NAME
End Function

'// Date & Time of file
Function Get_File_Date(ByRef FILE_PATH As String) As String
On Error Resume Next
    Get_File_Date = FileDateTime(FILE_PATH)
End Function

'// Get File Size (in bytes)
Function Get_File_Size(ByRef FILE_PATH As String) As String
On Error Resume Next
    FL_SIZE = CCur(FileLen(FILE_PATH)) '// Bytes
    Get_File_Size = FL_SIZE
End Function

'// Only legal time zone is GMT
Sub Convert_sec_to_Date()
    MsgBox CDate(880373744 / 86400)
End Sub


'// RE-ORGANIZE COOKIE FILE AFTER 1 COOKIE HAS BEEN DELETED
Function SaveFile(ByVal PATH As String, CONTENT As String)
On Error Resume Next
Dim intFile As Integer
     intFile = FreeFile
    Open PATH For Output As #intFile
        Print #intFile, CONTENT
    Close #intFile
End Function



Function Export_Coookies_Profile(ByRef LST As ListView, PATH As String)
On Error GoTo Error:
Dim X As Integer, Q As Integer, C_CON As String, F_CON As String
    CNT = LST.ListItems.Count
    For Q = 1 To CNT Step 1
            C_CON = ""
            C_CON = "DOMAIN:     " & LST.ListItems(Q).Text & vbCrLf
            C_CON = C_CON & "NAME:        " & LST.ListItems(Q).SubItems(1) & vbCrLf
            C_CON = C_CON & "VALUE:       " & LST.ListItems(Q).SubItems(2) & vbCrLf
            C_CON = C_CON & "SECURE:    " & LST.ListItems(Q).SubItems(3) & vbCrLf
            C_CON = C_CON & "PARENT:     " & LST.ListItems(Q).SubItems(5) & vbCrLf & vbCrLf
            F_CON = F_CON & C_CON
            F_CON = F_CON
    Next Q
    HDR = "There are " & Q & " cookie profiles available below:" & vbCrLf & vbCrLf
    F_CON = HDR & F_CON
    Call SaveFile(PATH, F_CON)
    MsgBox "Your cookie profile is available at:" & vbCrLf & PATH & vbCrLf, vbInformation, "Cookie Profile"
    Exit Function
Error:
 MsgBox "Wrong export path !" & vbCrLf & "Select Settings (F12) and set correct export path." & vbCrLf & "E.g.: " & App.PATH & "\MyCookies.txt", vbCritical
End Function



Function Apply_settings()
On Error Resume Next
    If SET_GF = "1" Then
        frmMain.lst_all_cookies.GridLines = True
    Else
        frmMain.lst_all_cookies.GridLines = False
    End If
    If SET_GC = "1" Then
        frmMain.Lst_Cook.GridLines = True
    Else
        frmMain.Lst_Cook.GridLines = False
    End If
End Function
