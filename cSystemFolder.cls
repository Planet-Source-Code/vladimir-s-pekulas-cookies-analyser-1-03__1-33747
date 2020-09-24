VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystemFolder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MAX_PATH = 260

'this enumeration holds all of the values accepted
'by the SHGetSpecialFolderLocation call, related to the
'folder name that will be retrieved.

Public Enum SystemFolder
    Desktop = 0
    StartMenu_Programs = 2
    My_Documents = 5
    Favorites = 6
    Startup = 7
    Recent = 8
    SentTo = 9
    Start_Menu = 11
    Windows_Desktop = 16
    Network_Neighborhood = 19
    Fonts = 20
    ShellNew = 21
    AllUsers_Desktop = 25
    ApplicationData = 26
    Printhood = 27
    TemporaryInternetFiles = 32
    Cookies = 33
    History = 34
End Enum

Public Function FindSystemFolder(ByVal lngNum As SystemFolder) As String

    On Error GoTo FindSystemFolder_Err
    
    Dim lpStartupPath As String * MAX_PATH
    Dim Pidl As Long
    Dim hResult As Long
    
    'find if a folder does exist with that number
    hResult = SHGetSpecialFolderLocation(0, lngNum, Pidl)


    If hResult = 0 Then 'there is a result
        
        'get the actualy directory name
        hResult = SHGetPathFromIDList(ByVal Pidl, lpStartupPath)


        If hResult = 1 Then
            'strip the string of all miscellaneous and unused characters
            
            lpStartupPath = Left$(Trim$(lpStartupPath), InStr(lpStartupPath, Chr(0)) - 1)
            FindSystemFolder = Trim$(lpStartupPath)
        End If
    End If
    
FindSystemFolder_End:
    
    Exit Function
    
FindSystemFolder_Err:

    'just raise an error is a problem occurs.
    'note that FindSystemFolder will be vbnullstring
    Err.Raise Err.Number, "FindSystemFolder::" & Err.Source, Err.Description
    
End Function

