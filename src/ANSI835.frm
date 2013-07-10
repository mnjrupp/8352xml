VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "835->XML 5010"
   ClientHeight    =   4185
   ClientLeft      =   210
   ClientTop       =   975
   ClientWidth     =   7215
   Icon            =   "ANSI835.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7215
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2160
      Top             =   1680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":000C
            Key             =   ""
            Object.Tag             =   "Exists in Database"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":045E
            Key             =   ""
            Object.Tag             =   "Does Not Exists in Database"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":08B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":0B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":0E6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":11C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ANSI835.frx":131B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3810
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "12:16 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   5636
            MinWidth        =   5645
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1482
            MinWidth        =   1482
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Last Updated"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   325
      Left            =   125
      TabIndex        =   1
      Top             =   3818
      Width           =   15
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3795
      Width           =   2895
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "&Open"
      Index           =   1
   End
   Begin VB.Menu mnuExport 
      Caption         =   "&Export"
      Index           =   2
      Begin VB.Menu mnuExcel 
         Caption         =   "E&xcel"
      End
      Begin VB.Menu mnuAccess 
         Caption         =   "&Access"
      End
      Begin VB.Menu mnuXML 
         Caption         =   "&XML"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "O&ptions"
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Export"
      Visible         =   0   'False
      Begin VB.Menu mnupopExcel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnupopupAccess 
         Caption         =   "Access"
      End
      Begin VB.Menu mnupopupXML 
         Caption         =   "XML"
      End
      Begin VB.Menu mnupopupEditor 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mnupopupOpen 
         Caption         =   "Open Directory"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private outptFile As FileSystemObject
Private appobjFile As TextStream
Private objFindFolder As New FileSystemObject
Private objfile835 As New FileSystemObject
Private objFolder As clsBrowseFolder
Attribute objFolder.VB_VarHelpID = -1
Private WithEvents objXML As clsConv2XML
Attribute objXML.VB_VarHelpID = -1
Private stsLeft As Long
Private stsTop As Long
Private varRegData As Variant
Private colFileNames As New Scripting.Dictionary
Private dicVarify As New Scripting.Dictionary
'Private strDefaultMDB As String
'Private flFileExists  As Boolean
Private UpdateFrequency  As Integer
Private fPath As String
' to hold handle for listview
Private m_hwndLV As Long



'      ' For Common Controls 6.0, uncomment the line below and comment the line above.
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As _
              MSComctlLib.ColumnHeader)
              
              Call ColumnClick(ColumnHeader)
               'Refresh the ListView before writing the data
              'Call SyncListviewItemNumbers
       
      End Sub
Private Sub ColumnClick(ByVal ColumnHeader As ColumnHeader)
      'On Error GoTo OOPS
        Dim strName As String
        Dim dDate As Date
        Dim lngItem As Long
        Dim sortImg As Long
        Dim h As Long
  ' the ListView will generate unnecessary NM_CUSTOMDRAW
  ' messages that slows the resizing down dramatically.
  Call SendMessage(m_hwndLV, WM_SETREDRAW, 0, 0)
'        'Handle User click on column header
       Select Case ColumnHeader.Text
       Case "#" 'don't want to sort the numbers
        Exit Sub
       Case "Last Updated"
           ListView1.Sorted = False       'User clicked on the Date header
                                         'Use our sort routine to sort
                                         'by date
            
                                         
           SendMessage ListView1.hwnd, _
                      LVM_SORTITEMS, _
                      ListView1.hwnd, _
                      ByVal FARPROC(AddressOf CompareDates)
            'Need to revert so the compare function will
            'change from Asc to Desc order
            If boolSortDesc = True Then
              boolSortDesc = False
              sortImg = 5
            Else
              boolSortDesc = True
              sortImg = 6
            End If
             'ShowHeaderIcon colNo, imgIndex, showFlag
             ShowHeaderIcon 4, sortImg, True
             
             
             For h = 0 To 3
                ShowHeaderIcon h, 0, 0
             Next 'h
             
            
        Case "Size"
            ListView1.Sorted = False       'User clicked on the Date header
                                         'Use our sort routine to sort
                                         'by date
           SendMessage ListView1.hwnd, _
                      LVM_SORTITEMS, _
                      ListView1.hwnd, _
                      ByVal FARPROC(AddressOf CompareValues)
            'Need to revert so the compare function will
            'change from Asc to Desc order
            If boolSortDesc = True Then
              boolSortDesc = False
              sortImg = 5
            Else
              boolSortDesc = True
              sortImg = 6
            End If
            
            ShowHeaderIcon 2, sortImg, True
            
             For h = 0 To 4
              If h <> 2 Then ShowHeaderIcon h, 0, False
             Next 'h
             
             'SyncListviewItemNumbers
        Case Else
                     ' set sort column to appropriate column:
            ListView1.SortKey = ColumnHeader.Index - 1
        
            ' change sort order to opposite:
            If ListView1.SortOrder = lvwAscending Then
                ListView1.SortOrder = lvwDescending
                sortImg = 5
            Else
                ListView1.SortOrder = lvwAscending
                sortImg = 6
            End If
        
            ' enable sorting
            ShowHeaderIcon ListView1.SortKey, sortImg, True
            
             For h = 0 To 4
              If h <> ListView1.SortKey Then ShowHeaderIcon h, 0, False
             Next 'h
            
            ListView1.Sorted = True
       End Select

        
         Call SendMessage(m_hwndLV, WM_SETREDRAW, 1, 0)
         ListView1.Refresh
        Exit Sub
OOPS:
        MsgBox Err.LastDllError
        Err.Clear
End Sub

Private Sub ListView1_ItemClick(ByVal item As ListItem)
On Error Resume Next
 Me.StatusBar1.Panels(4).Text = ""
 Me.StatusBar1.Panels(3).Text = Me.ListView1.SelectedItem.Text ' item.Text '" Ready"
 'Me.StatusBar1.Panels(2).Width = Me.StatusBar1.Panels(2).Width + 50
 SetProgressBarToStatusBar Me.ProgressBar1.hwnd, Me.StatusBar1.hwnd, 1
 'MsgBox item.SubItems(1)
 'item.Text = 1
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long
Dim lsv As ListItem
    If KeyCode = vbKeyControl Or KeyCode = vbKeyShift Then
        For Each lsv In Me.ListView1.ListItems
            If lsv.Selected Then i = i + 1
        Next
        Me.StatusBar1.Panels(4).Text = " " & i & " items selected "
    Else
        Me.StatusBar1.Panels(4).Text = ""
    End If
End Sub

Private Sub mnuAccess_Click()
    Call Export2Access
End Sub

Private Sub mnuExcel_Click()
On Error Resume Next
Call Export2Excel
End Sub
Private Sub ExportItemToExcel(strPathExcel As String)
On Error Resume Next
'This calles the X835.exe app to do the work for us
Shell (App.Path & "\X835.exe " & strPathExcel)
If Err.Number = 0 Then
 Me.StatusBar1.Panels(3).ToolTipText = strPathExcel
End If
End Sub
Private Sub mnuExport_Click(Index As Integer)
Dim strListViewSelect As String
If Me.ListView1.ListItems.count > 0 Then
    If Me.ListView1.SelectedItem.Selected Then
    If mnuAccess.Enabled = False Then mnuAccess.Enabled = True
    With Me
    strListViewSelect = .ListView1.SelectedItem.SubItems(1)
    strListViewSelect = strListViewSelect & Replace$(.ListView1.SelectedItem, "*", "")
    End With
 
 If colFileNames.Exists(LCase$(strListViewSelect)) Or dicVarify.Exists(LCase$(strListViewSelect)) Then
   mnuAccess.Enabled = False
 End If
End If
End If
End Sub

Private Sub mnupopupEditor_Click()
 Dim i As Long
 Dim strName1 As String
 Dim strPath1 As String
 Dim strSize1 As String
 Dim strType1 As String
 Dim dDate1 As Date
        
If Me.ListView1.ListItems.count > 0 Then
 For i = 1 To ListView1.ListItems.count
    If Me.ListView1.ListItems(i).Selected = True Then
    'Make a call to ListView_GetListItem because the ListItem collection
    'is out of sync
    'the index starts with 0 in WinAPI world but 1 in VB world
    ListView_GetListItem (i - 1), Me.ListView1.hwnd, _
    strName1, strPath1, strSize1, strType1, dDate1
    
    'MsgBox strPath1 & strName1
    'ConvertERA False, strPath1 & strName1
            
        On Error Resume Next

        Shell ("Notepad.exe " & strPath1 & strName1)
        
        DoEvents
    
    End If

Next 'i
End If
End Sub

Private Sub mnupopupOpen_Click()
On Error Resume Next
Dim lsv As ListItem
If Me.ListView1.ListItems.count < 1 Then Exit Sub
Set lsv = ListView1.SelectedItem()
  Shell "Explorer.exe """ & lsv.SubItems(1) & """", vbNormalFocus
   
End Sub

Private Sub mnuXML_Click()
On Error GoTo OPPS
 Dim strListViewSelect As String
 Dim i As Long
 Dim strName1 As String
 Dim strPath1 As String
 Dim strSize1 As String
 Dim strType1 As String
 Dim dDate1 As Date
        
If Me.ListView1.ListItems.count > 0 Then
 For i = 1 To ListView1.ListItems.count
    If Me.ListView1.ListItems(i).Selected = True Then
    'Make a call to ListView_GetListItem because the ListItem collection
    'is out of sync
    'the index starts with 0 in WinAPI world but 1 in VB world
    ListView_GetListItem (i - 1), Me.ListView1.hwnd, _
    strName1, strPath1, strSize1, strType1, dDate1
    
    'MsgBox strPath1 & strName1
    ConvertERA False, strPath1 & strName1
    End If

Next 'i
End If
  Exit Sub
OPPS:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: mnuXML_Click ", vbCritical
 Err.Clear
End Sub

Private Sub objXML_Timer(ByVal ShowProgress As Boolean)
'Need to clean out some memory
  
   Call RefreshProgress(ShowProgress)
End Sub

Public Static Sub RefreshProgress(Optional ShowProgress As Boolean = True)

 If Form1.ProgressBar1.Value < itemcount Then
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 1
    '
    Me.StatusBar1.Panels(3).Text = "Converting to xml " & CStr(Round(Form1.ProgressBar1.Value * 100 / Form1.ProgressBar1.Max, 0)) & " %"
    
 Else
    Form1.ProgressBar1.Value = itemcount
    
    Me.StatusBar1.Panels(3).Text = "Converting to xml " & CStr(Round(Form1.ProgressBar1.Value * 100 / Form1.ProgressBar1.Max, 0)) & " %"
    
 End If
  DoEvents

End Sub

Private Sub Form_Load()
 On Error Resume Next
 LoadSettingsFromFile
 UpdateFrequency = 25

'    g_crl16 = QBColor(3)
'    Set g_IFonts = Label1.Font
    With ListView1
      .SortKey = 0
      '.SmallIcons = ImageList1
    m_hwndLV = .hwnd
    End With
    ' Show header icons
    'ShowHeaderIcon colNo, imgIndex, showFlag
    'ShowHeaderIcon 0, 0, True
    stsLeft = Me.StatusBar1.Left
    stsTop = Me.StatusBar1.Top
    SetProgressBarToStatusBar Form1.ProgressBar1.hwnd, Form1.StatusBar1.hwnd, 1
    DoEvents
    
 With Form2
     If .Text1.Text = "" Then
        .Text1.Text = App.Path & "\dbMcareERA.mdb"
     End If
        .Top = Form1.Top + 1200
        .Left = Form1.Left + 1200
        '.Show vbModal
 End With
 
 ' Call mnuOptions_Click
   
End Sub

Private Sub LoadModelForm()
 Form1.Show vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingsToFile
    Unload Form2
    End

End Sub



Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim strListViewSelect As String
 
If Button = vbRightButton And Shift = 0 Then
    '*******************************************************************
    ' Need to varify if selection was imported during active session
    ' Check against dictionary; if so disable the ACCESS menu
    '*******************************************************************
    If mnupopupAccess.Enabled = False Then mnupopupAccess.Enabled = True
    With Me
    If .ListView1.ListItems.count > 0 Then
    strListViewSelect = .ListView1.SelectedItem.SubItems(1) & .ListView1.SelectedItem
    'Debug.Print strListViewSelect
    End If
    End With
 
    If colFileNames.Exists(LCase$(strListViewSelect)) Then mnupopupAccess.Enabled = False
    If dicVarify.Exists(LCase$(strListViewSelect)) Then mnupopupAccess.Enabled = False
   
    PopupMenu mnupopup, vbPopupMenuRightButton
 
End If
End Sub

Private Sub mnuOptions_Click()
Form2.Show
End Sub

Private Sub mnupopExcel_Click()
On Error Resume Next
Call Export2Excel
End Sub



Private Sub mnupopupAccess_Click()
 Call Export2Access
End Sub

Friend Sub ConvertERA(m_import As Boolean, strPath As String, Optional xmlAccess As Integer = 2)
Dim lstItem As ListItem
Dim opFile As TextStream
Dim rtnStr As String
Dim str835 As Variant
Static i As Long
Dim filelength As Integer
Dim filelength2 As Integer
'On Error GoTo OOPS

'Cannot use the Listview Collection due to the custom sort


If Len(strPath) > 1 Then
    strcomplete = strPath
    strXMLAddin = Form2.Text2.Text
    'strip out the file name by finding the length by position of the period
    filelength = InStrRev(strcomplete, "\", , vbTextCompare) + 1
    filelength2 = InStr(1, strcomplete, ".", vbTextCompare) - 1
    'Changed logic so that the file path of 835 is where the xml file will be created
    strXMLPath = Left$(strcomplete, filelength2)
    'strXMLPath = Left$(Me.ListView1.SelectedItem, InStr(1, Me.ListView1.SelectedItem, ".", vbTextCompare) - 1)
    strXMLPath = strXMLPath & ".xml"
    
    Set objfile835 = New FileSystemObject
    Set opFile = objfile835.OpenTextFile(strcomplete, ForReading)
    rtnStr = opFile.ReadAll
    Set opFile = Nothing
    'Set objfile835 = Nothing
    
     
      'Change to pull delimiter from combo box in Form2
    If Len(Trim(Form2.txtDelim.Text)) = 1 Then
        str835 = Split(rtnStr, Trim(Form2.txtDelim.Text), , vbTextCompare)
    Else
        '~ will be the default if blank
        str835 = Split(rtnStr, "~", , vbTextCompare)
    End If
    itemcount = CLng(UBound(str835))
    'set progress bar
    If boolIDEMode Then
        With Me
        .ProgressBar1.Min = 0
        .ProgressBar1.Max = itemcount
        .ProgressBar1.Value = 0
        End With
    End If
    DoEvents
    
    'Load Form3
    Set objXML = New clsConv2XML
    objXML.diceXML str835
    If boolIDEMode Then
        With Me
        .ProgressBar1.Value = 0
        .StatusBar1.Panels(3).Text = Right$(Mid$(strcomplete, filelength, (filelength2 - filelength) + 1), 24) & " converted to xml. "
        '.StatusBar1.Panels(2).Width = Len(.StatusBar1.Panels(2).Text) + 10
        .StatusBar1.Panels(3).ToolTipText = strXMLPath & " converted to xml"
        SetProgressBarToStatusBar .ProgressBar1.hwnd, .StatusBar1.hwnd, 1
        End With
    End If
     If m_import = True Then
     
        If boolIDEMode Then
               Screen.MousePointer = vbHourglass 'hourglass
               Me.StatusBar1.Panels(3).Text = "Importing into database."
        End If
               
               '***********************************************************
               ' Import XML into Access
               ' 0=structure;1=data and structure;2=data only
               ' xmlAccess = 0 or 1 or 2(default)
               ' strAccess = MDB database full path
               '***********************************************************
               
       If AccessXML(xmlAccess, strXMLPath, strAccess) Then
           If Not dicVarify.Exists(strcomplete) Then
               dicVarify.Add strcomplete, "Key" & i + 1
           End If
            
           
           If boolIDEMode Then
                With Me
                    .StatusBar1.Panels(3).Text = Mid$(strcomplete, filelength) & " imported."
                    '.StatusBar1.Panels(2).Width = .StatusBar1.Panels(2).Width + 50
                    .StatusBar1.Panels(3).ToolTipText = Mid$(strcomplete, filelength) & " import complete."
                    '.StatusBar1.Panels(1).Text = Me.ListView1.SelectedItem & " :: import complete."
                    SetProgressBarToStatusBar .ProgressBar1.hwnd, .StatusBar1.hwnd, 1
                End With
                Screen.MousePointer = vbDefault
           End If
       Else
        
        If boolIDEMode Then
           With Me
               .StatusBar1.Panels(3).Text = Err.Number & " " & Err.Description
               '.StatusBar1.Panels(2).Width = .StatusBar1.Panels(2).Width + 50
               .StatusBar1.Panels(3).ToolTipText = Err.Number & " " & Err.Description
               SetProgressBarToStatusBar .ProgressBar1.hwnd, .StatusBar1.hwnd, 1
           End With
           Screen.MousePointer = vbDefault
        End If
           'Set objfile835 = Nothing
           'MsgBox (Err.Number & " " & Err.Description)
       End If
        objfile835.DeleteFile strXMLPath, True
    End If
     '***************************************************
     ' need to clean up some things
         
          Set objfile835 = Nothing
          strXMLPath = ""
          
          rtnStr = ""
          ReDim str835(0)
     '***************************************************
 End If
 ' Next
  Exit Sub
OOPS:
    Set objfile835 = Nothing
'  MsgBox Err.Number & ": " & Err.Description & vbCrLf _
'  & Err.source & " :: ConvertERA", vbCritical, "Error in conversion"
End Sub

Private Sub mnuOpen_Click(Index As Integer)
'On Error GoTo OPPS
Set colFileNames = New Scripting.Dictionary
'Dim taFiles As mctFileSearchResults
Dim keycount As Integer
Dim strExtfromlabel As String
strExtfromlabel = Form2.Label3.Caption
'Dim splNames() As String
Dim splNames
Dim i As Long
 'Need to clear listviews if items in it
'
 If Me.ListView1.ListItems.count > 0 Then Me.ListView1.ListItems.Clear
 
 '----------------------------------------
 ' clear the header icon
 '----------------------------------------
 Dim h As Long
 For h = 0 To 4
    ShowHeaderIcon h, 0, 0
 Next 'h
 
 Me.StatusBar1.Panels(3).Text = ""
 Me.StatusBar1.Panels(4).Text = ""


 '*****************************************
 ' Pull all ERA Names from database to load in Dictionary
 '*****************************************
  If Form2.Combo2.Text <> "" Then
 
      splNames = Split(GetImportedFilePath(Form2.Combo2.ListIndex), ",", , vbTextCompare)
    If UBound(splNames) >= 0 Then
      For i = 0 To UBound(splNames)
      'Add ERA names to dictionary for later use
        If splNames(i) <> "NULL" Then
         If Not colFileNames.Exists(LCase$(splNames(i))) Then
            colFileNames.Add LCase$(splNames(i)), "Key" & keycount + 1
            'flFileExists = False
          Else
          'flFileExists = True
          End If
         End If
      Next
    End If
    Else
     MsgBox "Please choose database from dropdown menu under Options", vbCritical, "Choose database"
     Exit Sub
 End If
  'Run asynchronously
 Set objFolder = New clsBrowseFolder
    Dim hFile As Long
    Dim fName As String
    Dim fExt As String
    Dim counter As Integer
    Dim WFD As WIN32_FIND_DATA
    Dim dStart As Date
    Dim dFinish As Date
    Dim lScan As Long
    
    dStart = Now
   
    lScan = DateDiff("s", dStart, Now)
    Screen.MousePointer = vbHourglass
    'Check if looking into Application path or Default(My Computer)
    If Form2.Appcheck.Value = 0 Then
        fPath = objFolder.BrowseForFolder("") & "\"
    Else
        fPath = objFolder.BrowseForFolder(App.Path) & "\"
    End If
    
    If Len(fPath) > 0 And Len(strExtfromlabel) > 0 Then
        fName = fPath & strExtfromlabel
        hFile = FindFirstFile(fName, WFD)
         
            If hFile > 0 Then
                  
              counter = 1
              vbAddFileItemView WFD
               Call SendMessage(m_hwndLV, WM_SETREDRAW, 0, 0)
                While FindNextFile(hFile, WFD)
               ' Me.StatusBar1.Panels(2).Text = "Loading grid  |"
                 Me.StatusBar1.Refresh
                 DoEvents
                  counter = counter + 1
                  vbAddFileItemView WFD
                    'Me.StatusBar1.Panels(2).Text = "Loading grid  /"
                    Me.StatusBar1.Refresh
'                    If counter = UpdateFrequency Then
'                        Call UpdateWindow(ListView1.hwnd)
'                        counter = 0
'                    End If
                     
                Wend
                 Call SendMessage(m_hwndLV, WM_SETREDRAW, 1, 0)
                 Call ListView_SetColumnWidth(m_hwndLV, 0, LVSCW_AUTOSIZE)
                 
            End If
      
          FindClose hFile
      
    End If
                Call UpdateWindow(m_hwndLV)

    Me.StatusBar1.Panels(3).Text = ""
    Me.StatusBar1.Panels(4).Text = ListView1.ListItems.count & " files"
    SetProgressBarToStatusBar Me.ProgressBar1.hwnd, Me.StatusBar1.hwnd, 1
    Screen.MousePointer = vbDefault
 
   Exit Sub
OPPS:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: mnuOpen_Click ", vbCritical
  Err.Clear
End Sub

Private Sub vbAddFileItemView(WFD As WIN32_FIND_DATA)
 On Error GoTo OPPS
    Dim sFileName As String
    Dim ListImgKey As String
    Dim fType As String
    Static itemx As Long

    sFileName = TrimNull(WFD.cFileName)
    ' the ListView will generate unnecessary NM_CUSTOMDRAW
  ' messages that slows the resizing down dramatically.
  ' Call SendMessage(m_hwndLV, WM_SETREDRAW, 0, 0)
    
    If sFileName <> "." And sFileName <> ".." Then
              
        Dim hInfo As Long
        Dim tExeType As Long
        Dim itmX As ListItem
        Dim subitemX As ListSubItem
        
        If ListView1.ListItems.count = 0 Then itemx = 0
        itemx = itemx + 1
    
       ' Set itmX = ListView1.ListItems.Add(, , LCase$(sFileName), , 1)
                   
        hInfo = SHGetFileInfo(fPath & sFileName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS)
        
        fType = LCase$(TrimNull(shinfo.szTypeName))
       If colFileNames.Exists(LCase$(fPath & sFileName)) Then
       Set itmX = ListView1.ListItems.Add(, , LCase$(sFileName), , 1)
        'Set subitemX = itmX.ListSubItems.Add(, , LCase$(sFileName))
       'subitemX.ReportIcon = 1
        'itmX.SubItems(1) = LCase$(sFileName)
        Else
            Set itmX = ListView1.ListItems.Add(, , LCase$(sFileName), , 2)
          'Set subitemX = itmX.ListSubItems.Add(, , LCase$(sFileName))
          'subitemX.ReportIcon = 2
           'subitemX.ForeColor = vbRed
        End If
       
        itmX.SubItems(1) = fPath
        
        itmX.SubItems(2) = vbGetFileSizeKBStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
        itmX.SubItems(3) = fType
        itmX.SubItems(4) = vbGetFileDate$(WFD.ftCreationTime)
      
   End If
   ' reset drawing
   'Call SendMessage(m_hwndLV, WM_SETREDRAW, 1, 0)
   'ListView1.Refresh
  Exit Sub
OPPS:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: vbAddFileItemView ", vbCritical
   Err.Clear
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim realWidth As Integer
Dim realHeight As Integer
Dim adjlblTop As Integer
Static rc As RECT
Static cxPrev As Long
realWidth = Me.ScaleWidth
realHeight = Me.ScaleHeight - 200
    
Const offset = 100

 ' If we don't turn the ListView's painting off while resizing columns,
  ' the ListView will generate unnecessary NM_CUSTOMDRAW
  ' messages that slows the resizing down dramatically.
  Call SendMessage(m_hwndLV, WM_SETREDRAW, 0, 0)

 Me.ListView1.Move offset, 2 * offset, realWidth - 2 * offset, realHeight - 2 * offset - Form1.StatusBar1.Height
 'Me.ListView2.Move 20, ListView1.Top + 70, ListView2.Width, ListView1.Height - 400
 adjlblTop = Me.ListView1.Top + Me.ListView1.Height
 Me.StatusBar1.Width = realWidth - 50

 Me.Label1.Left = Form1.ScaleLeft + 100
 Me.Label1.Top = Form1.ScaleHeight - 400
 Me.Label2.Height = Me.Label1.Height - 30
 Me.Label2.Left = Me.Label1.Left + 15
 Me.Label2.Top = Me.Label1.Top + 8
 
  Call GetClientRect(m_hwndLV, rc)
  If (cxPrev <> rc.Right) Then
    
    ' Save the ListView's previous width.
    cxPrev = rc.Right
    
    ' See: "FIX: Problem with ListView's ColumnHeader Width Property"
    ' http://support.microsoft.com/support/kb/articles/q179/9/88.asp
    ' and: "HOWTO: Set the Column Width of Columns in a ListView Control"
    ' http://support.microsoft.com/support/kb/articles/q147/6/66.asp
    'Call ListView_SetColumnWidth(m_hwndLV, 0, rc.Right \ 5)
    Call ListView_SetColumnWidth(m_hwndLV, 1, rc.Right \ 5)
    Call ListView_SetColumnWidth(m_hwndLV, 2, rc.Right \ 5)
    Call ListView_SetColumnWidth(m_hwndLV, 3, rc.Right \ 5)
    'Call ListView_SetColumnWidth(m_hwndLV, 4, rc.Right \ 6)
    Call ListView_SetColumnWidth(m_hwndLV, 4, rc.Right - ((rc.Right \ 5) * 3))
    
    ' Now tell the ListView to generate a NM_CUSTOMDRAW
    ' for all items after painting is turned back on.
    Call InvalidateRect(m_hwndLV, ByVal 0&, CFalse)
  End If
 
  Call SendMessage(m_hwndLV, WM_SETREDRAW, 1, 0)
End Sub

Private Sub mnupopupXML_Click()
 On Error GoTo OPPS
 Dim strListViewSelect As String
 Dim i As Long
 Dim strName1 As String
 Dim strPath1 As String
 Dim strSize1 As String
 Dim strType1 As String
 Dim dDate1 As Date
        
If Me.ListView1.ListItems.count > 0 Then
 For i = 1 To ListView1.ListItems.count
    If Me.ListView1.ListItems(i).Selected = True Then
    'Make a call to ListView_GetListItem because the ListItem collection
    'is out of sync
    'the index starts with 0 in WinAPI world but 1 in VB world
    ListView_GetListItem (i - 1), Me.ListView1.hwnd, _
    strName1, strPath1, strSize1, strType1, dDate1
    
    'MsgBox strPath1 & strName1
    ConvertERA False, strPath1 & strName1
    End If

Next 'i
End If
  Exit Sub
OPPS:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: mnupopupXML_Click ", vbCritical
 Err.Clear
End Sub

Private Sub objXML_writeXML(strXMLPath As String, completeXML As String)
         
Set outptFile = New FileSystemObject
Set appobjFile = outptFile.OpenTextFile(strXMLPath, ForWriting, True)
appobjFile.Write completeXML
appobjFile.Close
Set appobjFile = Nothing
Set outptFile = Nothing
Me.MousePointer = 0


End Sub

 
Friend Function AccessXML(appendtype As Integer, _
xmlPath As String, Accesspath As String) As Boolean
On Error Resume Next
'The appendtype determines what is to be imported
 '0=structure;1=data and structure;2=data only
 'xmlPath=path to xml file
 'Accesspath=complete path to .mdb file
Err.Clear
'Due to the various Office version,Best to late bind using CreateObject
Dim objAccess As Object
 'Dim objAccess As Object
Set objAccess = CreateObject("Access.Application")
objAccess.Visible = False
objAccess.OpenCurrentDatabase Accesspath
If Err.Number <> 0 Then
    AccessXML = False
    Exit Function
End If
objAccess.ImportXml xmlPath, appendtype

If Err.Number <> 0 And Err.Number <> 31550 Then 'We want to ignore 31550:Not All data was imported..
    AccessXML = False
Else
    AccessXML = True
End If
Set objAccess = Nothing

End Function
Private Function GetImportedFilePath(intPath As Integer) As String
On Error Resume Next
Dim objConn As New ADODB.Connection
Dim objRecord As New ADODB.Recordset
Dim objCommand As New ADODB.Command
Dim strSQL As String
Dim strRows As String
Dim varstrFiles As Variant
varstrFiles = Split(Form2.Text1.Text, ";", , vbTextCompare)
'The array index will correspond to the Combo box index passed in
strAccess = CStr(varstrFiles(intPath))

'Build the SQL statement used to pull ERA file names

strSQL = "SELECT SE_file FROM SE;"

If Form2.Check1(0) Then
objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strAccess & ";Mode=Read|Write"
End If
' if either both checkboxes are ticked or just Access 07 in form2 then
' updated connection string info
If Form2.Check1(1) Then
    objConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strAccess & ";Persist Security Info=False;"
End If


'the location of the cursor engine is on the client

objConn.CursorLocation = adUseClient

'open the connection

objConn.Open

'Define the command object

With objCommand
    .ActiveConnection = objConn
    .CommandText = strSQL
    .CommandType = adCmdText
End With


'Defines our recordset object

With objRecord
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open objCommand
   
End With

' will use the GetString Method to pull all ERA names in a string
' to split out and put in an Array
' this way the connection can be closed.
strRows = objRecord.GetString(adClipString, , , ",")

Set objRecord = Nothing
Set objConn = Nothing

GetImportedFilePath = strRows


End Function


Private Sub LoadSettingsFromFile()
On Error Resume Next
Dim tmpObj As New FileSystemObject
Dim tmpTxt As TextStream
Dim strFile As String
Dim varFileArray As Variant
Dim varSettingsArray As Variant
Dim i As Integer, j As Integer, k As Integer
Dim tmpUBound As Integer
' Read text file where configuration is stored
' and place info in its proper place
 If tmpObj.FileExists(App.Path & "/" & App.EXEName & ".ini") Then
    Set tmpTxt = tmpObj.OpenTextFile(App.Path & "/" & App.EXEName & ".ini", ForReading, False)
    strFile = tmpTxt.ReadAll
 End If
 'Clean up
 Set tmpObj = Nothing
 Set tmpTxt = Nothing
 
  'need to split out info based on its line in file
  'by using Split() function
  varFileArray = Split(strFile, vbCrLf, , vbTextCompare)
  If UBound(varFileArray) > 0 Then
     tmpUBound = UBound(varFileArray)
     For k = 0 To tmpUBound
     'Placed in SELECT Case so
     'Can add lines if more info is needed
      Select Case k
          Case 0
              Form2.Text1.Text = CStr(varFileArray(0))
              varSettingsArray = Split(varFileArray(0), ";", , vbTextCompare)
              For i = 0 To UBound(varSettingsArray)
                  Form2.Combo2.AddItem Mid(CStr(varSettingsArray(i)), InStrRev(CStr(varSettingsArray(i)), "\", , vbTextCompare) + 1)
              Next
          Case 1
              varSettingsArray = Split(varFileArray(1), ";", , vbTextCompare)
              Form1.Left = CLng(varSettingsArray(0))
              Form1.Top = CLng(varSettingsArray(1))
              Form1.Width = CLng(varSettingsArray(2))
              Form1.Height = CLng(varSettingsArray(3))
      End Select
    'varFileArray(0) = database names and Paths seperated by ";"
    
     Next 'k
  End If
  

End Sub
Public Sub SaveSettingsToFile()
On Error GoTo SaveErr
Dim tmpObj As New FileSystemObject
Dim tmpFile As TextStream
  'Just add more Writelines for every line of info needed
Set tmpFile = tmpObj.CreateTextFile(App.Path & "\" & App.EXEName & ".ini", True, False)
  tmpFile.WriteLine Form2.Text1.Text
  tmpFile.WriteLine CStr(Form1.Left) & ";" & CStr(Form1.Top) & ";" & CStr(Form1.Width) & ";" & CStr(Form1.Height)
  tmpFile.Close
  
  Set tmpFile = Nothing
  Set tmpObj = Nothing
Exit Sub
SaveErr:
     'Raise an error back to the caller
     Err.Raise vbObjectError + 7000, "SaveSettingsToFile", Err.Description
     Err.Clear
End Sub

Private Function vbGetFileDate(CT As FILETIME) As String
 On Error GoTo OPPS
    Dim ST As SYSTEMTIME
    Dim ds As Single
       
    If FileTimeToSystemTime(CT, ST) Then
       ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
       vbGetFileDate$ = Format$(ds, "Short Date")
    Else
       vbGetFileDate$ = ""
    End If
   Exit Function
OPPS:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: vbGetFileDate ", vbCritical
  Err.Clear
End Function


Private Function vbGetFileSizeKBStr(fsize As Long) As String
On Error Resume Next
    vbGetFileSizeKBStr = Format$(((fsize) / 1000) + 0.5, "#,###,###") & " kb"

End Function

Private Function TrimNull(item As String) As String
    Dim pos As Integer
    pos = InStr(item, Chr$(0))
    If pos Then item = Left$(item, pos - 1)
    TrimNull = item
  
End Function

Public Sub ShowHeaderIcon(colNo As Long, _
                          imgIconNo As Long, _
                          showImage As Long)

   Dim hHeader As Long
   Dim HD As HD_ITEM
   
  'get a handle to the listview header component
   hHeader = SendMessage(ListView1.hwnd, LVM_GETHEADER, 0, ByVal 0)
   
  'set up the required structure members
   With HD
      .mask = HDI_IMAGE Or HDI_FORMAT
      .pszText = ListView1.ColumnHeaders(colNo + 1).Text
      
       If showImage Then
         .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
         .iImage = imgIconNo
       Else
         .fmt = HDF_STRING
      End If


   End With
   
  'modify the header
   Call SendMessage(hHeader, HDM_SETITEM, colNo, HD)
   
End Sub

'Private Sub SyncListviewItemNumbers()
'Dim item As ListItem
'Dim i As Long
'Dim num As Long
'For Each item In ListView1.ListItems
'     num = num + 1
'     item.Text = CInt(num)
'    ' Debug.Print item.Index & " == " & item.Text
'Next
'
'End Sub

'Friend Function WindowProc(hwnd As Long, _
'                           msg As Long, _
'                           wParam As Long, _
'                           lParam As Long) As Long
'
''**************************************
''Subclassing
''**************************************
'
'  'Dim sci As SCROLLINFO
'
'  If hwnd = ListView1.hwnd Then
'
'      Select Case msg
'
'        'message of interest
'         Case WM_VSCROLL
'
''           '------------------------------------------
''           'this block (between the dashed lines)
''           'demonstrates obtaining scrollbar information
''           'and can actually be commented out without
''           'impacting the method!
''
''           'fill a SCROLLINFO structure to receive
''           'scrollbar data from the subclassed
''           'listview and call the GetScrollInfo API
''            With sci
''               .cbSize = Len(sci)
''               .fMask = SIF_ALL
''            End With
''
''            Call GetScrollInfo(hWnd, SB_VERT, sci)
''
''           'Information only: shows the values
''           'returned by the API as tabbed list items
''            With List1
''               .AddItem sci.nMin & vbTab & _
''                        sci.nMax & vbTab & _
''                        sci.nPage & vbTab & _
''                        sci.nPos & vbTab & _
''                        sci.nTrackPos & vbTab & _
''                        wParam & vbTab & lParam
''               .TopIndex = .NewIndex
''            End With
''
''           'If you wanted to provide any special
''           'capability, ie if a user enacted a line up but
''           'you wanted to translate that into a page down,
''           'etc, you could do that here.
''            Select Case wParam
''               Case SB_LINEUP:         '0
''                 'wParam = SB_LINEDOWN  'tweak the action!
''               Case SB_LINEDOWN:       '1
''               Case SB_PAGEUP:         '2
''               Case SB_PAGEDOWN:       '3
''               Case SB_THUMBPOSITION:  '4
''               Case SB_THUMBTRACK:     '5
''               Case SB_TOP:            '6
''               Case SB_BOTTOM:         '7
''               Case SB_ENDSCROLL:      '8
''            End Select
''
''           'SetScrollInfo sets the target window's
''           'scrollbar's characteristics to match the
''           'source. Where a value is outside a valid
''           'range, (ie if the source returned 100 for
''           'nMax, but the target only had 50 items,
''           'the target will set the nMax number to 50).
''            Call SetScrollInfo(ListView2.hWnd, SB_VERT, sci, 1&)
''           '------------------------------------------
'
'           'The actual scrolling method - a one-liner!
'           'On entering this routine, wParam
'           'contains one of the SB_xxx messages
'           'listed above. By passing it directly
'           'to the mirrored listview, the mirror
'           'tracks as the subclassed listview
'           'is scrolled.
''            Call SendMessage(ListView2.hwnd, _
''                             WM_VSCROLL, _
''                             wParam, _
''                             ByVal 0&)
''
'           'If you want to disable scrolling in the
'           'subclassed listview, but want the mirrored
'           'listview to scroll as if its scrollbar had
'           'been used, uncomment the two lines below.
''           WindowProc = 0
''           Exit Function
'
'         Case Else
'      End Select
'
'   End If
'
'  'pass on to the default window procedure
'   WindowProc = CallWindowProc(GetProp(hwnd, "OldWindowProc"), _
'                               hwnd, msg, _
'                               wParam, lParam)
'
'End Function

Private Sub Export2Excel()
    On Error GoTo OOPS

 Dim i As Long
 Dim strName1 As String
 Dim strPath1 As String
 Dim strSize1 As String
 Dim strType1 As String
 Dim dDate1 As Date
        
If Me.ListView1.ListItems.count > 0 Then
 For i = 1 To ListView1.ListItems.count
    
    If Me.ListView1.ListItems(i).Selected = True Then
    'Make a call to ListView_GetListItem because the ListItem collection
    'is out of sync
    'the index starts with 0 in WinAPI world but 1 in VB world
    ListView_GetListItem (i - 1), Me.ListView1.hwnd, _
    strName1, strPath1, strSize1, strType1, dDate1
    
    'MsgBox strPath1 & strName1
    ExportItemToExcel strPath1 & strName1
     Me.StatusBar1.Panels(3).Text = strName1
    Exit For
    'DoEvents
    End If

Next 'i
End If

 Exit Sub
OOPS:
  MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: Export2Excel", vbCritical
End Sub

Private Sub Export2Access()
On Error GoTo OPPS
 Dim strListViewSelect As String
 Dim i As Long
 Dim strName1 As String
 Dim strPath1 As String
 Dim strSize1 As String
 Dim strType1 As String
 Dim dDate1 As Date
        
If Me.ListView1.ListItems.count > 0 Then
 For i = 1 To ListView1.ListItems.count
    If Me.ListView1.ListItems(i).Selected = True Then
    'clean out the status bar text and reset size to make way for next
    With Me.StatusBar1
    .Panels(3).Text = ""
    End With
    'Make a call to ListView_GetListItem because the ListItem collection
    'is out of sync
    'the index starts with 0 in WinAPI world but 1 in VB world
    ListView_GetListItem (i - 1), Me.ListView1.hwnd, _
    strName1, strPath1, strSize1, strType1, dDate1
    ConvertERA True, strPath1 & strName1
    End If
Next 'i
End If

 Exit Sub
OPPS:
  MsgBox Err.Number & ": " & Err.Description & vbCrLf _
  & Err.Source & " :: mnupopupAccess_Click", vbCritical
  Err.Clear
End Sub

Private Sub Timer1_Timer()
    Call mnuOptions_Click
    Timer1.Enabled = False
End Sub
