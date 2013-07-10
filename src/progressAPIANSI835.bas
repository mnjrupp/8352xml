Attribute VB_Name = "progressAPI"
Private Const WM_USER = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

'********************************************************************
'Declare the API functions used in tranposing a progress bar over a status bar
'********************************************************************

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent _
As Long) As Long

Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'**********************************************************************
' Declare a type for RECT
'**********************************************************************
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Sub SetProgressBarToStatusBar(ByVal hWnd_PBar As Long, ByVal hWnd_SBar As Long, _
ByVal nPanel As Long)
    Dim R As RECT
    SendMessage hWnd_SBar, SB_GETRECT, nPanel - 1, R
    SetParent hWnd_PBar, hWnd_SBar
    MoveWindow hWnd_PBar, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, True



End Sub

