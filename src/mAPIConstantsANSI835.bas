Attribute VB_Name = "mAPIConstants"
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This code was written by The Frog Prince
'
' If you have questions or comments, I can be reached at
'        TheFrogPrince@hotmail.com
' If you wanna see more cool vb user controls, classes, code,
' and add-ins like this one, or updates to this code, go to
' my web page at
'        http://members.tripod.com/the__frog__prince/
' You are free to use, re-write, or otherwise do as you wish
' with this code.  However, if you do a cool enhancement, I
' would appreciate it if you could e-mail it to me.  I like
' to see what people do with my stuff.  =)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


Option Explicit
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  'system icon index
Public Const SHGFI_LARGEICON = &H0        'large icon
Public Const SHGFI_SMALLICON = &H1        'small icon
Public Const ILD_TRANSPARENT = &H1        'display transparent
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
             SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
             SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Public Const MAX_PATH = 260

Public Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type

Public shinfo As SHFILEINFO

Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type


Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Public Declare Function FindFirstFile _
                            Lib "kernel32" _
                            Alias "FindFirstFileA" ( _
                        ByVal lpFileName As String, _
                        lpFindFileData As WIN32_FIND_DATA) _
                    As Long

Public Declare Function FindNextFile _
                            Lib "kernel32" _
                            Alias "FindNextFileA" ( _
                        ByVal hFindFile As Long, _
                        lpFindFileData As WIN32_FIND_DATA) _
                    As Long

Public Declare Function FindClose _
                                    Lib "kernel32" ( _
                                ByVal hFindFile As Long) _
                            As Long
                            
Public Declare Function SHGetFileInfo Lib "shell32" _
   Alias "SHGetFileInfoA" _
  (ByVal pszPath As String, _
   ByVal dwFileAttributes As Long, _
   psfi As SHFILEINFO, _
   ByVal cbSizeFileInfo As Long, _
   ByVal uFlags As Long) As Long
   
Public Declare Function GetVolumeInformation _
                                    Lib "kernel32" _
                                    Alias "GetVolumeInformationA" ( _
                                ByVal lpRootPathName As String, _
                                ByVal lpVolumeNameBuffer As String, _
                                ByVal nVolumeNameSize As Long, _
                                lpVolumeSerialNumber As Long, _
                                lpMaximumComponentLength As Long, _
                                lpFileSystemFlags As Long, _
                                ByVal lpFileSystemNameBuffer As String, _
                                ByVal nFileSystemNameSize As Long) _
                            As Long



'Public Declare Function PathStripToRoot _
'                                    Lib "SHLWAPI.DLL" _
'                                    Alias "PathStripToRootA" ( _
'                                ByVal pszPath As String) _
'                            As Long
                                
'Public Declare Function PathIsNetworkPath _
'                                    Lib "SHLWAPI.DLL" _
'                                    Alias "PathIsNetworkPathA" ( _
'                                ByVal pszPath As String) _
'                            As Boolean
            

'Public Declare Function PathIsUNCServerShare _
'                                    Lib "SHLWAPI.DLL" _
'                                    Alias "PathIsUNCServerShareA" ( _
'                                ByVal pszPath As String) _
'                            As Boolean


'Public Declare Function PathIsUNCServer _
'                                    Lib "SHLWAPI.DLL" _
'                                    Alias "PathIsUNCServerA" ( _
'                                ByVal pszPath As String) _
'                            As Boolean

'Public Declare Function PathIsUNC _
'                                    Lib "SHLWAPI.DLL" _
'                                    Alias "PathIsUNCA" ( _
'                                ByVal pszPath As String) _
'                            As Boolean

'Public Declare Function OpenFile _
'                            Lib "kernel32" ( _
'                        ByVal lpFileName As String, _
'                        lpReOpenBuff As OFSTRUCT, _
'                        ByVal wStyle As Long) _
'                    As Long

Public Declare Function CloseHandle _
                            Lib "kernel32" ( _
                        ByVal hObject As Long) _
                    As Long

'Public Declare Function GetFileInformationByHandle _
'                            Lib "kernel32" ( _
'                        ByVal hFile As Long, _
'                        lpFileInformation As BY_HANDLE_FILE_INFORMATION) _
'                    As Long

Public Declare Function FileTimeToSystemTime _
                            Lib "kernel32" ( _
                        lpFileTime As FILETIME, _
                        lpSystemTime As SYSTEMTIME) _
                    As Long

'Public Declare Function SystemTimeToFileTime _
'                                    Lib "kernel32" ( _
'                                lpSystemTime As SYSTEMTIME, _
'                                lpFileTime As FILETIME) _
'                            As Long

'Public Declare Function GetExpandedName _
'                            Lib "lz32.dll" _
'                            Alias "GetExpandedNameA" ( _
'                        ByVal lpszSource As String, _
'                        ByVal lpszBuffer As String) _
'                    As Long


'Public Declare Function GetShortPathName _
'                            Lib "kernel32" _
'                            Alias "GetShortPathNameA" ( _
'                        ByVal lpszLongPath As String, _
'                        ByVal lpszShortPath As String, _
'                        ByVal cchBuffer As Long) _
'                    As Long

'Public Declare Function SetFileAttributes _
'                            Lib "kernel32" _
'                            Alias "SetFileAttributesA" ( _
'                        ByVal lpFileName As String, _
'                        ByVal dwFileAttributes As Long) _
'                    As Long

'Public Declare Function GetFileAttributes _
'                            Lib "kernel32" _
'                            Alias "GetFileAttributesA" ( _
'                        ByVal lpFileName As String) _
'                    As enumFileAttributes

'Public Declare Function GetDriveType _
'                            Lib "kernel32" _
'                            Alias "GetDriveTypeA" ( _
'                        ByVal nDrive As String) _
'                    As Long
Public Declare Function UpdateWindow Lib "user32" _
      (ByVal hwnd As Long) As Long

