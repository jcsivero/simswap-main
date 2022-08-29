VERSION 5.00
Object = "*\AMSVBCldr.vbp"
Begin VB.Form Form2 
   Caption         =   "Calendar Test"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   ScaleHeight     =   5490
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frames 
      Caption         =   "DayBold and DayItalic"
      Height          =   855
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Top             =   1800
      Width           =   3975
      Begin VB.ComboBox cbxDayNum 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   795
      End
      Begin VB.CheckBox chkDayItalic 
         Caption         =   "Italic"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkDayBold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Day Number:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   930
      End
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   3780
      TabIndex        =   2
      Top             =   240
      Width           =   1995
   End
   Begin VB.CommandButton btnSetValue 
      Caption         =   "Set"
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame Frames 
      Caption         =   "Day Name Format"
      Height          =   1095
      Index           =   0
      Left            =   6180
      TabIndex        =   8
      Top             =   600
      Width           =   1575
      Begin VB.OptionButton rbNameFormats 
         Caption         =   "Short"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton rbNameFormats 
         Caption         =   "Medium"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   1035
      End
      Begin VB.OptionButton rbNameFormats 
         Caption         =   "Long"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1035
      End
   End
   Begin VB.Frame Frames 
      Caption         =   "Navigation Options"
      Height          =   1095
      Index           =   1
      Left            =   3780
      TabIndex        =   4
      Top             =   600
      Width           =   2295
      Begin VB.CheckBox chkMonthRO 
         Caption         =   "Month Read-Only"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkYearRO 
         Caption         =   "Year Read-Only"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
      Begin VB.CheckBox chkShowIterration 
         Caption         =   "Show Iterration Buttons"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1995
      End
   End
   Begin VB.ListBox lbxEvents 
      Height          =   2235
      Left            =   0
      TabIndex        =   19
      Top             =   3240
      Width           =   7875
   End
   Begin VB.CheckBox chkShowWillChange 
      Caption         =   "Show WillChangeDate Message"
      Height          =   195
      Left            =   3900
      TabIndex        =   17
      Top             =   2880
      Width           =   2715
   End
   Begin MSVBCalendar.Calendar Calendar1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5106
      Day             =   15
      Month           =   10
      Year            =   1996
      BeginProperty DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Current Date (Value):"
      Height          =   195
      Index           =   0
      Left            =   3780
      TabIndex        =   1
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Events:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   540
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_fIgnoreEvent As Boolean

Private Sub btnSetValue_Click()
    Calendar1.Value = DateValue(txtValue.Text)
End Sub

Private Sub Calendar1_DateChange(ByVal OldDate As Date, ByVal NewDate As Date)
    txtValue.Text = NewDate
    AddEvent "DateChange: OldDate = " & OldDate & ", NewDate = " & NewDate
End Sub

Private Sub Calendar1_DblClick()
    AddEvent "DblClick: Current Date = " & Calendar1.Value
End Sub

Private Sub Calendar1_WillChangeDate(ByVal NewDate As Date, Cancel As Boolean)
    Dim sPrompt As String
    
    AddEvent "WillChangeDate: NewDate = " & NewDate
    If Me.chkShowWillChange Then
        sPrompt = "Date will change from " & Calendar1.Value & " to " & NewDate & "." & vbCrLf & "Will you allow the change?"
        If MsgBox(sPrompt, vbYesNo + vbQuestion, "WillChangeDate Event") = vbNo Then
            AddEvent "Change Denied -- Cancel set to True in WillChange event"
            Cancel = True
        End If
        Calendar1.Refresh
    End If
End Sub

Private Sub cbxDayNum_Click()
    m_fIgnoreEvent = True
    
    If Calendar1.DayBold(cbxDayNum.Text) Then
        chkDayBold.Value = 1
    Else
        chkDayBold.Value = 0
    End If
    
    If Calendar1.DayItalic(cbxDayNum.Text) Then
        chkDayItalic.Value = 1
    Else
        chkDayItalic.Value = 0
    End If
    
    m_fIgnoreEvent = False
End Sub

Private Sub chkDayBold_Click()
    If Not m_fIgnoreEvent Then
        Calendar1.DayBold(cbxDayNum.Text) = CBool(chkDayBold.Value)
        Calendar1.Refresh
    End If
End Sub

Private Sub chkDayItalic_Click()
    If Not m_fIgnoreEvent Then
        Calendar1.DayItalic(cbxDayNum.Text) = CBool(chkDayItalic.Value)
        Calendar1.Refresh
    End If
End Sub

Private Sub chkMonthRO_Click()
    Calendar1.MonthReadOnly = CBool(chkMonthRO.Value)
End Sub

Private Sub chkShowIterration_Click()
    Calendar1.ShowIterrationButtons = CBool(chkShowIterration.Value)
End Sub

Private Sub chkYearRO_Click()
    Calendar1.YearReadOnly = CBool(chkYearRO.Value)
End Sub

Private Sub Form_Load()
    Dim nDay As Long
    
    txtValue.Text = Calendar1.Value
    rbNameFormats(Calendar1.DayNameFormat).Value = True
    chkMonthRO.Value = Abs(Calendar1.MonthReadOnly)
    chkYearRO.Value = Abs(Calendar1.YearReadOnly)
    chkShowIterration.Value = Abs(Calendar1.ShowIterrationButtons)
    Me.Caption = "Calendar Version " & Calendar1.Version
    
    For nDay = 1 To 31
        cbxDayNum.AddItem nDay
    Next nDay
    cbxDayNum.ListIndex = 0
End Sub

Private Sub rbNameFormats_Click(Index As Integer)
    Calendar1.DayNameFormat = Index
End Sub

Private Sub AddEvent(sText As String)
    If lbxEvents.ListCount > 1000 Then
        lbxEvents.Clear
    End If
    lbxEvents.AddItem sText
    lbxEvents.ListIndex = lbxEvents.NewIndex
End Sub


