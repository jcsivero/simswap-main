VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "ComCt232.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Begin VB.Form frmSSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VB 5 Saver Setup"
   ClientHeight    =   3285
   ClientLeft      =   2565
   ClientTop       =   2070
   ClientWidth     =   6015
   Icon            =   "SSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frSpeed 
      Caption         =   "Sprite Speed"
      Height          =   645
      Left            =   90
      TabIndex        =   14
      Top             =   2520
      Width           =   5805
      Begin VB.CheckBox chkSpeedRND 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   1125
      End
      Begin ComctlLib.Slider sldSpeed 
         Height          =   330
         Left            =   2100
         TabIndex        =   16
         Top             =   210
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   582
         TickStyle       =   3
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5340
         TabIndex        =   24
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Index           =   6
         Left            =   1710
         TabIndex        =   18
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Index           =   5
         Left            =   4860
         TabIndex        =   17
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame frSpriteSize 
      Caption         =   "Sprite Size %"
      Height          =   645
      Left            =   90
      TabIndex        =   9
      Top             =   1800
      Width           =   5805
      Begin VB.CheckBox chkSizeRND 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Width           =   1125
      End
      Begin ComctlLib.Slider sldSize 
         Height          =   330
         Left            =   2100
         TabIndex        =   11
         Top             =   210
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   582
         TickStyle       =   3
      End
      Begin VB.Label lblSize 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5340
         TabIndex        =   23
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Small"
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   13
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Large"
         Height          =   195
         Index           =   3
         Left            =   4860
         TabIndex        =   12
         Top             =   270
         Width           =   405
      End
   End
   Begin VB.Frame frRefreshRate 
      Caption         =   "Sprite Animation Rate"
      Height          =   645
      Left            =   90
      TabIndex        =   3
      Top             =   1080
      Width           =   5805
      Begin VB.CheckBox chkRefreshRND 
         Caption         =   "Randomize"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   1125
      End
      Begin ComctlLib.Slider sldRefreshRate 
         Height          =   330
         Left            =   2100
         TabIndex        =   4
         Top             =   210
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   582
         TickStyle       =   3
      End
      Begin VB.Label lblRefresh 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5340
         TabIndex        =   22
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Index           =   2
         Left            =   4860
         TabIndex        =   6
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Index           =   1
         Left            =   1710
         TabIndex        =   5
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame fSettings 
      Caption         =   "Sprites"
      Height          =   945
      Left            =   90
      TabIndex        =   2
      Top             =   0
      Width           =   1875
      Begin VB.PictureBox picCount 
         BackColor       =   &H80000005&
         Height          =   315
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   525
         TabIndex        =   20
         Top             =   210
         Width           =   585
         Begin VB.TextBox txtSprites 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "10"
            Top             =   30
            Width           =   195
         End
         Begin ComCtl2.UpDown udCount 
            Height          =   255
            Left            =   330
            TabIndex        =   25
            Top             =   0
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   450
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtSprites"
            BuddyDispid     =   196624
            OrigLeft        =   345
            OrigRight       =   540
            OrigBottom      =   255
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.CheckBox chkTracers 
         Caption         =   "Show Tracers"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Count:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5010
      TabIndex        =   1
      Top             =   540
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5010
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "frmSSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkRefreshRND_Click()
    gRefreshRND = (chkRefreshRND.Value = vbChecked) ' Save rand refresh rate globally
    sldRefreshRate.Enabled = Not gRefreshRND
End Sub
Private Sub chkSizeRND_Click()
    gSizeRND = (chkSizeRND.Value = vbChecked)       ' Save rand sprite size globally
    sldSize.Enabled = Not gSizeRND
End Sub
Private Sub chkSpeedRND_Click()
    gSpeedRND = (chkSpeedRND.Value = vbChecked)     ' Save rand animation rate globally
    sldSpeed.Enabled = Not gSpeedRND
End Sub
Private Sub chkTracers_Click()
    gTracers = (chkTracers.Value = vbChecked)       ' Save use tracers option globally
End Sub
Private Sub cmdCancel_Click()
    Unload Me                                       ' Cancel screen saver setup dialog
End Sub
Private Sub cmdOK_Click()
    SaveSettings                                    ' Save current screen saver settings...
    Unload Me                                       ' Close setup dialog
End Sub

'------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------
    ' Show the screensaver about box...
    If (KeyAscii = Asc("?")) Then AboutBox Me.hwnd
'------------------------------------------------------------
End Sub
'------------------------------------------------------------

'------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------
    ' Load current screen saver registry settings...
    LoadSettings
    
    ' Get Sprite Count Value
    With udCount
        .Max = MAX_SPRITECOUNT
        .Min = MIN_SPRITECOUNT
        .Value = gSpriteCount
    End With
    
    ' Get Refresh Rate Value
    With sldRefreshRate
        .Max = MAX_REFRESHRATE
        .Min = MIN_REFRESHRATE
        .Value = gRefreshRate
        lblRefresh.Caption = CStr(gRefreshRate)
    End With
    
    ' Get Sprite Size Value
    With sldSize
        .Max = MAX_SPRITESIZE
        .Min = MIN_SPRITESIZE
        .Value = gSpriteSize
        lblSize.Caption = CStr(gSpriteSize)
    End With
    
    ' Get Sprite Speed Value
    With sldSpeed
        .Max = MAX_SPRITESPEED
        .Min = MIN_SPRITESPEED
        .Value = gSpriteSpeed
        lblSpeed.Caption = CStr(gSpriteSpeed)
    End With
    
    ' Get Tracers on Value
    If gTracers Then chkTracers.Value = vbChecked
    
    ' Get Rate Random Value
    If gRefreshRND Then chkRefreshRND.Value = vbChecked
    
    ' Get Size Random Value
    If gSizeRND Then chkSizeRND.Value = vbChecked
    
    ' Get Speed Random Value
    If gSpeedRND Then chkSpeedRND.Value = vbChecked
'------------------------------------------------------------
End Sub
'------------------------------------------------------------

Private Sub sldRefreshRate_Change()
    gRefreshRate = sldRefreshRate.Value             ' Save animation refresh rate globally
    lblRefresh.Caption = CStr(gRefreshRate)
End Sub
Private Sub sldRefreshRate_Scroll()
    gRefreshRate = sldRefreshRate.Value             ' Save animation refresh rate globally
    lblRefresh.Caption = CStr(gRefreshRate)
End Sub
Private Sub sldSize_Change()
    gSpriteSize = sldSize.Value                     ' Save active sprite size globally
    lblSize.Caption = CStr(gSpriteSize)
End Sub
Private Sub sldSize_Scroll()
    gSpriteSize = sldSize.Value                     ' Save active sprite size globally
    lblSize.Caption = CStr(gSpriteSize)
End Sub
Private Sub sldSpeed_Change()
    gSpriteSpeed = sldSpeed.Value                   ' Save active sprite speed globally
    lblSpeed.Caption = CStr(gSpriteSpeed)
End Sub
Private Sub sldSpeed_Scroll()
    gSpriteSpeed = sldSpeed.Value                   ' Save active sprite speed globally
    lblSpeed.Caption = CStr(gSpriteSpeed)
End Sub
Private Sub txtSprites_Change()
    gSpriteCount = Val(txtSprites.Text)             ' Save active sprite count globally
End Sub
