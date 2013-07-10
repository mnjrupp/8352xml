VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   4515
   ClientLeft      =   1560
   ClientTop       =   1710
   ClientWidth     =   5850
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Misc"
      Height          =   1575
      Left            =   2400
      TabIndex        =   12
      Top             =   2760
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtDelim 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   16
         Text            =   "~"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Access 03"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Access 07"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Delimiter"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         ToolTipText     =   "Delimiter used in 835:~ default"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "ANSI835frm2.frx":0000
      Left            =   3240
      List            =   "ANSI835frm2.frx":0002
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   "Char(s) to append to ERA Name"
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Extensions"
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
      Begin VB.CheckBox Appcheck 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "ANSI835frm2.frx":0004
         Left            =   240
         List            =   "ANSI835frm2.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Application path"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "*.txt"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "XML"
      Height          =   1095
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
      Begin VB.CommandButton cmdFindStyle 
         Caption         =   "..."
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Style sheet: used on import of XML in Access"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "*.xsl stylesheet for remits and imports"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Database"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Place additional Char(s) to ERA Name. This will be in SE Segment"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Database Names:Enter filepath seperated by semicolon"";"""
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Me.Label3.Caption = Combo1.Text
 Form1.Caption = "835->XML 5010     " & Combo2.Text & "    " & Me.Label3.Caption & " Files"
End Sub




Private Sub Combo2_Click()
    Form1.Caption = "835->XML 5010     " & Combo2.Text & "    " & Me.Label3.Caption & " Files"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Me.Hide
End Sub

