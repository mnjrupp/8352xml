VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

' define a custom structure
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

' declare API functions we need:
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Function BrowseForFolder(ByVal initdir As Variant, _
                                Optional ByVal message As String = "Choose!", _
                                Optional ByVal flags As Long) As String
    Dim rootID As Long
    Dim PIDL As Long
    Dim Path As String
    Dim nullPos As Integer
    Dim BInfo As BrowseInfo
    Dim myMessage As String
    
    ' convert the message string to ANSI code:
    myMessage = StrConv(message, vbFromUnicode)

    ' check whether user specified a path or a code
    If VarType(initdir) = vbString Then
        ' it's a path!
        ' use undocumented function to create IDList
        ' convert path string to Unicode:
        rootID = SHSimpleIDListFromPath(StrConv(initdir, vbUnicode))
    Else
        ' it's a virtual system folder code
        ' get "real" path
        SHGetSpecialFolderLocation GetDesktopWindow, initdir, rootID
    End If

    ' fill out the BrowseInfo structure:
    BInfo.hwndOwner = GetDesktopWindow
    BInfo.ulFlags = flags
    ' fill in the address of your ANSI message string:
    BInfo.lpszTitle = StrPtr(myMessage)
    
    ' is there a valid rootID? Fill it in!
    If rootID <> 0 Then BInfo.pIDLRoot = rootID
    
    ' open the dialog box, retrieve a PIDL
    ' PIDL is internal identifier for selected folder:
    PIDL = SHBrowseForFolder(BInfo)
    
    ' was there a PIDL returned? Transform to path:
    If PIDL <> 0 Then
        ' reserve space for path name
        Path = String(260, 0)
        
        ' return path name from PIDL:
        SHGetPathFromIDList PIDL, Path
        
        ' manually release the memory:
        Call CoTaskMemFree(PIDL)
        Call CoTaskMemFree(rootID)
        
        ' cut off string at char 0:
        nullPos = InStr(Path, vbNullChar)
        If nullPos <> 0 Then
            Path = Left(Path, nullPos - 1)
        End If
    End If
    
    ' return path:
    BrowseForFolder = Path
End Function




