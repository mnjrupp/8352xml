Attribute VB_Name = "Module1"
'constant that holds the caption of the main form
Option Explicit
'Public Const MAIN_FORM_CAPTION = "Form1"

 Public strXMLPath As String
 Public strXMLAddin As String
 Public strAccess As String
 Public strcomplete As String
 'Public ttlLen As Long
 Public itemcount As Long
 Public mstrConv2XL As String
 Public mlngTimer2XL As Long
 Public TRNtracer As String
 Public BPRstr As String
 Public CLPUid As Long
 Public CLPstr As String
 Public CLP_num As Long
 Public SVCstr As String
 Public lnItem As Integer
 Public STtrans As String
 Public GStrans As String

 '***************************************
 ' this is used to turn off IDE functions
 ' like progress bars
 
 Public boolIDEMode As Boolean

 '***************************************
 ' Is 5010 version; default is yes
 
 Public boolv5010 As Boolean
 
 '***************************************
'This is needed for the sorting function
 'False if the Date column is in ascending order
 'True if Descending
 Public boolSortDesc As Boolean

 'Public objfile As clsFileSearch 'Private Search Object
 Public objXML As clsConv2XML 'declare the XML class module
 'Public m_booleanEvent As Boolean 'This will be changed on each loop in clsConv2XML
 '**************************************************************************
 'The following is taken from MSDN KB170884
 'This is for sorting dates properly
 'The Listview treats ListItems as strings so sorting is "Out of wack"
 ' This is a workaround
 
      'Structures

      Public Type POINT
        x As Long
        y As Long
      End Type

      Public Type LV_FINDINFO
        flags As Long
        psz As String
        lParam As Long
        pt As POINT
        vkDirection As Long
      End Type

      Public Type LV_ITEM
        mask As Long
        iItem As Long
        iSubItem As Long
        State As Long
        stateMask As Long
        pszText As Long
        cchTextMax As Long
        iImage As Long
        lParam As Long
        iIndent As Long
      End Type
    'For Header Column Items
Public Type HD_ITEM
   mask As Long
   cxy As Long
   pszText As String
   hbm As Long
   cchTextMax As Long
   fmt As Long
   lParam As Long
   iImage As Long
   iOrder As Long
End Type

Public Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum
      'Constants
Public Const LVFI_PARAM = 1
Public Const LVIF_TEXT = &H1

Public Const LVM_FIRST = &H1000
Public Const LVM_FINDITEM = LVM_FIRST + 13
Public Const LVM_GETITEMTEXT = LVM_FIRST + 45
Public Const LVM_SORTITEMS = LVM_FIRST + 48

Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2

Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_IMAGE = &H800
Public Const HDF_STRING = &H4000

Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)
Public Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Public Const HDM_GETIMAGELIST = (HDM_FIRST + 9)


Public Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, _
 ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Declare Function FindWindow Lib "user32" Alias _
 "FindWindowA" (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long
  
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As CBoolean) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As CBoolean) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function SetTimer Lib "user32" _
'(ByVal hwnd As Long, ByVal nIDEvent As Long, _
'ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

  
'Public Declare Function KillTimer Lib "user32" _
'   (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
   
'Continuation of Sorting workaround
'Module Functions and Procedures

      'CompareDates: This is the sorting routine that gets passed to the
      'ListView control to provide the comparison test for date values.

      Public Function CompareDates(ByVal lngParam1 As Long, _
                                   ByVal lngParam2 As Long, _
                                   ByVal hwnd As Long) As Long

'        Dim indx1    As String
'        Dim indx2    As String
        Dim strName1 As String
        Dim strName2 As String
        Dim strPath1 As String
        Dim strPath2 As String
        Dim strSize1 As String
        Dim strSize2 As String
        Dim strType1 As String
        Dim strType2 As String
        Dim dDate1 As Date
        Dim dDate2 As Date

        'Obtain the item names and dates corresponding to the
        'input parameters

        ListView_GetItemData lngParam1, hwnd, strName1, strPath1, strSize1, strType1, dDate1
        ListView_GetItemData lngParam2, hwnd, strName2, strPath2, strSize2, strType2, dDate2

        'Compare the dates
        'Return 0 ==> Less Than
        '       1 ==> Equal
        '       2 ==> Greater Than
        Select Case boolSortDesc
         Case False
            If dDate1 < dDate2 Then
              CompareDates = -1 '0
            ElseIf dDate1 = dDate2 Then
              CompareDates = 0 '1
            Else
              CompareDates = 2
            End If
         Case True
            If dDate1 < dDate2 Then
              CompareDates = 2 '0
            ElseIf dDate1 = dDate2 Then
              CompareDates = 0 '1
            Else
              CompareDates = -1
            End If
        End Select

      End Function

 Public Function CompareValues(ByVal lngParam1 As Long, _
                              ByVal lngParam2 As Long, _
                              ByVal hwnd As Long) As Long
     
  'CompareValues: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for numeric values.
  'Added 11/18/07 MR

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
    'Dim indx1    As String
    'Dim indx2    As String
    Dim strName1 As String
    Dim strName2 As String
    Dim strPath1 As String
    Dim strPath2 As String
    Dim strSize1 As String
    Dim strSize2 As String
    Dim strType1 As String
    Dim strType2 As String
    Dim dDate1 As Date
    Dim dDate2 As Date
    Dim val1 As Double
    Dim val2 As Double
     
  'Obtain the item names and values corresponding
  'to the input parameters
        ListView_GetItemData lngParam1, hwnd, strName1, strPath1, strSize1, strType1, dDate1
        ListView_GetItemData lngParam2, hwnd, strName2, strPath2, strSize2, strType2, dDate2
   'val1 = ListView_GetItemValueStr(hWnd, lParam1)
   'val2 = ListView_GetItemValueStr(hWnd, lParam2)
   ' Debug.Print strSize1 & " = " & strSize2
     val1 = Val(Format(Left$(strSize1, InStr(1, strSize1, "k") - 1), "General Number"))
     val2 = Val(Format(Left$(strSize2, InStr(1, strSize2, "k") - 1), "General Number"))
  'based on the Public variable sOrder set in the
  'columnheader click sub, sort the values appropriately:
        Select Case boolSortDesc
         Case False
            If val1 < val2 Then
              CompareValues = -1 '0
            ElseIf val1 = val2 Then
              CompareValues = 0 '1
            Else
              CompareValues = 2
            End If
         Case True
            If val1 < val2 Then
              CompareValues = 2 '0
            ElseIf val1 = val2 Then
              CompareValues = 0 '1
            Else
              CompareValues = -1
            End If
        End Select
'   Select Case sOrder
'      Case True 'sort descending
'
'            If val1 < val2 Then
'               CompareValues = 0
'            ElseIf val1 = val2 Then
'               CompareValues = 1
'            Else
'               CompareValues = 2
'            End If
'
'      Case Else 'sort ascending
'
'            If val1 > val2 Then
'               CompareValues = 0
'            ElseIf val1 = val2 Then
'               CompareValues = 1
'            Else
'               CompareValues = 2
'            End If
'
'   End Select

End Function

      'GetItemData - Given Retrieves

      Public Sub ListView_GetItemData(lngParam As Long, _
                                      hwnd As Long, _
                                      strName As String, _
                                      strPath As String, _
                                      strSize As String, _
                                      strType As String, _
                                      dDate As Date)
        Dim objFind As LV_FINDINFO
        Dim lngIndex As Long
        Dim objItem As LV_ITEM
        Dim baBuffer(32) As Byte
        Dim lngLength As Long
        Dim iIndex As Integer
        '
        ' Convert the input parameter to an index in the list view
        '
        objFind.flags = LVFI_PARAM
        objFind.lParam = lngParam
        'lngIndex = SendMessage(hwnd, LVM_FINDITEM, -1, VarPtr(objFind))
        lngIndex = SendMessage(hwnd, LVM_FINDITEM, -1, objFind)
        '
        ' Obtain the name of the specified list view item
        For iIndex = 0 To 4
        objItem.mask = LVIF_TEXT
        objItem.iSubItem = iIndex '0
        objItem.pszText = VarPtr(baBuffer(0))
        objItem.cchTextMax = UBound(baBuffer)
        lngLength = SendMessage(hwnd, LVM_GETITEMTEXT, lngIndex, _
                                objItem)
        Select Case iIndex
'        Case 0
'            indx = Left$(StrConv(baBuffer, vbUnicode), lngLength)
'            'Debug.Print indx & " => #"
        Case 0
         strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 1
         strPath = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 2
            strSize = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 3
            strType = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 4
            If lngLength > 0 Then
                dDate = CDate(Left$(StrConv(baBuffer, vbUnicode), lngLength))
            End If
        End Select
         
        Next
        

      End Sub

      'GetListItem - This is a modified version of ListView_GetItemData
      ' It takes an index into the list as a parameter and returns
      ' the appropriate values in the strName and dDate parameters.

      Public Sub ListView_GetListItem(lngIndex As Long, _
                                      hwnd As Long, _
                                      strName As String, _
                                      strPath As String, _
                                      strSize As String, _
                                      strType As String, _
                                      dDate As Date)
        Dim objItem As LV_ITEM
        Dim baBuffer(255) As Byte
        Dim lngLength As Long
        Dim iIndex As Long
        '
        ' Obtain the name of the specified list view item
        '
        

        For iIndex = 0 To 4
        objItem.mask = LVIF_TEXT
        objItem.iSubItem = iIndex 'changed to match the modified date column from Listview
        objItem.pszText = VarPtr(baBuffer(0))
        objItem.cchTextMax = UBound(baBuffer)

        'lngLength = SendMessage(hwnd, LVM_GETITEMTEXT, lngIndex, _
                                VarPtr(objItem))
         lngLength = SendMessage(hwnd, LVM_GETITEMTEXT, lngIndex, _
                                objItem)
        Select Case iIndex
        Case 0
            strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 1
            strPath = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 2
            strSize = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 3
            strType = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        Case 4
            If lngLength > 0 Then
                dDate = CDate(Left$(StrConv(baBuffer, vbUnicode), lngLength))
            End If
        End Select
        Next 'iIndex
      End Sub
' End of Sorting workaround
'***********************************************************************



'NEED TO CHANGE
'Remove this TimerProc since this will be an apartment thread
'app
'Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
'On Error Resume Next
'   Static Busy As Boolean
'   '
'   ' Make sure we're not re-entering here.
'   If Not IsNull(objXML) Then
'   If Not Busy Then
'      '
'      ' Tell global event generator to fire
'      '
'      Busy = True
'      objXML.RaiseEvents
'      Busy = False
'    End If
'   End If
'End Sub

Public Function FARPROC(ByVal pfn As Long) As Long
  
  'A procedure that receives and returns
  'the value of the AddressOf operator.
  'This workaround is needed as you can't assign
  'AddressOf directly to an API when you are also
  'passing the value ByVal in the statement
  '(as is being done with SendMessage)
 
  FARPROC = pfn

End Function

Public Sub Main()
'Will parse the Command line and see if we have passed arguments
' if not then load Form
boolIDEMode = True
If UBound(Split(Command, " ")) < 0 Then

    Form1.Show
 
Else
 On Error GoTo OPPS
    Dim cmdLine() As String
    Dim objFile As FileSystemObject
    Set objFile = New FileSystemObject

    cmdLine = Split(Command, " ")
    'Verify that we have the file to parse into XML and/or database to import into
    Select Case UBound(cmdLine)
        Case 0
            ' Request an XML file by passing 835
            '
            If objFile.FileExists(cmdLine(0)) Then
                boolIDEMode = False
                Form1.ConvertERA False, cmdLine(0)
                End
            Else
                End
            End If
        Case 1
            'Now we have 2 arguments. Need to verify 2 file
            'Paths. first = 835 second = .mdb file
            If objFile.FileExists(cmdLine(0)) Then
                 If objFile.FileExists(cmdLine(1)) And (InStr(1, cmdLine(1), ".mdb") Or InStr(1, cmdLine(1), ".accdb")) Then
                    Dim strXMLPath As String
                   strAccess = cmdLine(1)
                   boolIDEMode = False
                   Form1.ConvertERA True, cmdLine(0)
                   End
                End If
            End If
            End
        Case 2
            If objFile.FileExists(cmdLine(0)) Then
                 If objFile.FileExists(cmdLine(1)) And (InStr(1, cmdLine(1), ".mdb") Or InStr(1, cmdLine(1), ".accdb")) Then
                    If cmdLine(2) = 0 Or CInt(cmdLine(2)) = 1 Or CInt(cmdLine(2)) = 2 Then
                    strAccess = cmdLine(1)
                    boolIDEMode = False
                    Form1.ConvertERA True, cmdLine(0), CInt(cmdLine(2))
                    End
                    End If
                End If
            Else
                End
            End If
            End
    End Select

    Set objFile = Nothing
    Form1.Show
    'Need to verify path name in first argument
    
End If
Exit Sub
OPPS:
' Need to clean up
    Set objFile = Nothing
Form1.Show
End Sub

Public Function Replace09(ByRef Text As String, _
    ByRef sOld As String, ByRef sNew As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal count As Long = 2147483647, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As String
' by Jost Schwider, jost@schwider.de, 20001218

  If LenB(sOld) Then

    If Compare = vbBinaryCompare Then
      Replace09Bin Replace09, Text, Text, _
          sOld, sNew, Start, count
    Else
      Replace09Bin Replace09, Text, LCase$(Text), _
          LCase$(sOld), sNew, Start, count
    End If

  Else 'Suchstring ist leer: Search string is empty
    Replace09 = Text
  End If
End Function

Private Static Sub Replace09Bin(ByRef result As String, _
    ByRef Text As String, ByRef Search As String, _
    ByRef sOld As String, ByRef sNew As String, _
    ByVal Start As Long, ByVal count As Long _
  )
' by Jost Schwider, jost@schwider.de, 20001218
  Dim TextLen As Long
  Dim OldLen As Long
  Dim NewLen As Long
  Dim ReadPos As Long
  Dim WritePos As Long
  Dim CopyLen As Long
  Dim Buffer As String
  Dim BufferLen As Long
  Dim BufferPosNew As Long
  Dim BufferPosNext As Long
  
  'Ersten Treffer bestimmen:
  'Determine if first hit
  If Start < 2 Then
    Start = InStrB(Search, sOld)
  Else
    Start = InStrB(Start + Start - 1, Search, sOld)
  End If
  If Start Then
  
    OldLen = LenB(sOld)
    NewLen = LenB(sNew)
    Select Case NewLen
    Case OldLen 'einfaches Überschreiben:Simply overwrite
    
      result = Text
      For count = 1 To count
        MidB$(result, Start) = sNew
        Start = InStrB(Start + OldLen, Search, sOld)
        If Start = 0 Then Exit Sub
      Next count
      Exit Sub
    
    Case Is < OldLen 'Ergebnis wird kürzer:Result is shorter
    
      'Buffer initialisieren:Initialize buffer
      TextLen = LenB(Text)
      If TextLen > BufferLen Then
        Buffer = Text
        BufferLen = TextLen
      End If
      
      'Ersetzen:Replace
      ReadPos = 1
      WritePos = 1
      If NewLen Then
      
        'Einzufügenden Text beachten:Insert text note
        For count = 1 To count
          CopyLen = Start - ReadPos
          If CopyLen Then
            BufferPosNew = WritePos + CopyLen
            MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
            MidB$(Buffer, BufferPosNew) = sNew
            WritePos = BufferPosNew + NewLen
          Else
            MidB$(Buffer, WritePos) = sNew
            WritePos = WritePos + NewLen
          End If
          ReadPos = Start + OldLen
          Start = InStrB(ReadPos, Search, sOld)
          If Start = 0 Then Exit For
        Next count
      
      Else
      
        'Einzufügenden Text ignorieren (weil leer):Insert text note (because empty)
        For count = 1 To count
          CopyLen = Start - ReadPos
          If CopyLen Then
            MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
            WritePos = WritePos + CopyLen
          End If
          ReadPos = Start + OldLen
          Start = InStrB(ReadPos, Search, sOld)
          If Start = 0 Then Exit For
        Next count
      
      End If
      
      'Ergebnis zusammenbauen:Assemble result
      If ReadPos > TextLen Then
        result = LeftB$(Buffer, WritePos - 1)
      Else
        MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
        result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
      End If
      Exit Sub
    
    Case Else 'Ergebnis wird länger: Result is longer
    
      'Buffer initialisieren:Initialize buffer
      TextLen = LenB(Text)
      BufferPosNew = TextLen + NewLen
      If BufferPosNew > BufferLen Then
        Buffer = Space$(BufferPosNew)
        BufferLen = LenB(Buffer)
      End If
      
      'Ersetzung: Replace
      ReadPos = 1
      WritePos = 1
      For count = 1 To count
        CopyLen = Start - ReadPos
        If CopyLen Then
          'Positionen berechnen:Calculate positions
          BufferPosNew = WritePos + CopyLen
          BufferPosNext = BufferPosNew + NewLen
          
          'Ggf. Buffer vergrößern: If necessary. Expand buffers
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = LenB(Buffer)
          End If
          
          'String "patchen": String "patching"
          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
          MidB$(Buffer, BufferPosNew) = sNew
        Else
          'Position bestimmen: Location
          BufferPosNext = WritePos + NewLen
          
          'Ggf. Buffer vergrößern: If necessary. Expand buffers
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = LenB(Buffer)
          End If
          
          'String "patchen": String patching
          MidB$(Buffer, WritePos) = sNew
        End If
        WritePos = BufferPosNext
        ReadPos = Start + OldLen
        Start = InStrB(ReadPos, Search, sOld)
        If Start = 0 Then Exit For
      Next count
      
      'Ergebnis zusammenbauen: Assemble result
      If ReadPos > TextLen Then
        result = LeftB$(Buffer, WritePos - 1)
      Else
        BufferPosNext = WritePos + TextLen - ReadPos
        If BufferPosNext < BufferLen Then
          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
          result = LeftB$(Buffer, BufferPosNext)
        Else
          result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
        End If
      End If
      Exit Sub
    
    End Select
  
  Else 'Kein Treffer: No hits
    result = Text
  End If
End Sub
