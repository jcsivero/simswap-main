VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test Icon Handler..."
   ClientHeight    =   2790
   ClientLeft      =   2775
   ClientTop       =   2730
   ClientWidth     =   4575
   Icon            =   "Test_IH.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4575
   Begin VB.ComboBox cmbGuids 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   300
      Width           =   4395
   End
   Begin VB.TextBox txtOut 
      Height          =   1485
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   4395
   End
   Begin VB.CommandButton cmdCallHandler 
      Caption         =   "Call IconHandler From GUID"
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GUID:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   450
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------
Private Sub cmdCallHandler_Click()
'----------------------------------------------------------------
    Dim rf As Long                              ' Return flags
    Dim Idx As Long                             ' Icon Index
    Dim IconFile As String                      ' Icon file output
    Dim Handler As cExtractIcon                 ' Object reference variable
'----------------------------------------------------------------
    Set Handler = New cExtractIcon              ' Instansiate cExtractIcon class.
    txtOut.Text = ""                            ' Clear output textbox
    
    ' Call IconHandler in GUID object...(Method 1)
    Handler.GetIconLocation cmbGuids.Text, FOR_SHELL, Idx, IconFile, rf
    txtOut.Text = txtOut.Text & IconFile        ' Display IconFile output
    txtOut.Text = txtOut.Text & vbCrLf
    
    ' Call IconHandler in GUID object...(Method 2)
    Handler.GetIconLocation cmbGuids.Text, OPEN_ICON, Idx, IconFile, rf
    txtOut.Text = txtOut.Text & IconFile        ' Display IconFile output
    txtOut.Text = txtOut.Text & vbCrLf
    
    Set Handler = Nothing                       ' Destroy handler object.
'----------------------------------------------------------------
End Sub
'----------------------------------------------------------------

'----------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------
    ' Add a few known classid's that are known IconHandlers...
    ' Note that these objects may not be installed on your system...
    cmbGuids.AddItem "{FBF23B40-E3F0-101B-8488-00AA003E56F8}"
    cmbGuids.AddItem "{00021401-0000-0000-C000-000000000046}"
    cmbGuids.AddItem "{0006F045-0000-0000-C000-000000000046}"
    cmbGuids.Text = cmbGuids.List(0)
'----------------------------------------------------------------
End Sub
'----------------------------------------------------------------

