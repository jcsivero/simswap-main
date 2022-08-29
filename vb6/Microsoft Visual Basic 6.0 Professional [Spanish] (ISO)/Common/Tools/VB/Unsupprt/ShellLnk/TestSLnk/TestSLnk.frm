VERSION 5.00
Begin VB.Form frmTestSLnk 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tests the IShellLink Typelib Interface."
   ClientHeight    =   3780
   ClientLeft      =   420
   ClientTop       =   720
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtProgramGroup 
      Height          =   285
      Left            =   3120
      TabIndex        =   20
      Top             =   3300
      Width           =   5865
   End
   Begin VB.CommandButton cmdCreateGroup 
      Caption         =   "Create Group"
      Height          =   375
      Left            =   180
      TabIndex        =   18
      Top             =   1050
      Width           =   1125
   End
   Begin VB.CommandButton cmdGetLinkInfo 
      Caption         =   "GetLinkInfo"
      Height          =   375
      Left            =   180
      TabIndex        =   17
      Top             =   540
      Width           =   1125
   End
   Begin VB.ComboBox cmbSysFolders 
      Height          =   315
      Left            =   3120
      TabIndex        =   16
      Top             =   90
      Width           =   5865
   End
   Begin VB.CommandButton cmdGetPath 
      Caption         =   "GetSysPath"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   60
      Width           =   1125
   End
   Begin VB.TextBox txtShowCmd 
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   2895
      Width           =   585
   End
   Begin VB.TextBox txtCmdArgs 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   1740
      Width           =   5865
   End
   Begin VB.TextBox txtIconIndex 
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   2505
      Width           =   585
   End
   Begin VB.TextBox txtIconFile 
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   2130
      Width           =   5865
   End
   Begin VB.TextBox txtWorkDir 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   1365
      Width           =   5865
   End
   Begin VB.TextBox txtExeName 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   975
      Width           =   5865
   End
   Begin VB.TextBox txtLinkName 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   5865
   End
   Begin VB.CommandButton cmdCreateLink 
      Caption         =   "CreateLink"
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start Menu Group:"
      Height          =   195
      Index           =   7
      Left            =   1770
      TabIndex        =   19
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Show Command:"
      Height          =   195
      Index           =   6
      Left            =   1860
      TabIndex        =   14
      Top             =   2940
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cmd Arguments:"
      Height          =   195
      Index           =   5
      Left            =   1890
      TabIndex        =   12
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Icon Index:"
      Height          =   195
      Index           =   4
      Left            =   2265
      TabIndex        =   10
      Top             =   2565
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Icon FileName:"
      Height          =   195
      Index           =   3
      Left            =   1965
      TabIndex        =   8
      Top             =   2175
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Working Directory:"
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   6
      Top             =   1425
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exe Name:"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Link Name:"
      Height          =   195
      Index           =   0
      Left            =   2250
      TabIndex        =   2
      Top             =   660
      Width           =   810
   End
End
Attribute VB_Name = "frmTestSLnk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
Private Sub cmdCreateGroup_Click()
'---------------------------------------------------------------
    MkDir txtProgramGroup.Text                          ' Create Start Menu Program Group...
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub cmdCreateLink_Click()
'---------------------------------------------------------------
    Dim sLnk As cShellLink                              ' ShellLink Variable
'---------------------------------------------------------------
    Set sLnk = New cShellLink                           ' Create ShellLink Instance
    
    sLnk.CreateShellLink txtLinkName.Text, _
                         txtExeName.Text, _
                         txtWorkDir.Text, _
                         txtCmdArgs.Text, _
                         txtIconFile.Text, _
                    CLng(txtIconIndex.Text), _
                    CLng(txtShowCmd.Text)               ' Create a ShellLink (ShortCut)
    
    Set sLnk = Nothing                                  ' Destroy object reference
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub cmdGetLinkInfo_Click()
'---------------------------------------------------------------
    Dim sLnk As cShellLink                              ' ShellLink class variable
    Dim LnkFile As String                               ' Link file name
    Dim ExeFile As String                               ' Link - Exe file name
    Dim WorkDir As String                               '      - Working directory
    Dim ExeArgs As String                               '      - Command line arguments
    Dim IconFile As String                              '      - Icon File name
    Dim IconIdx As Long                                 '      - Icon Index
    Dim ShowCmd As Long                                 '      - Program start state...
'---------------------------------------------------------------
    Set sLnk = New cShellLink                           ' Create new Explorer IShellLink Instance
    
    LnkFile = txtLinkName.Text                          ' Get link file name
    txtExeName.Text = ""                                ' Clear output variables...
    txtWorkDir.Text = ""
    txtCmdArgs.Text = ""
    txtIconFile.Text = ""
    txtIconIndex.Text = ""
    txtShowCmd.Text = ""
    
    sLnk.GetShellLinkInfo LnkFile, _
                          ExeFile, _
                          WorkDir, _
                          ExeArgs, _
                          IconFile, _
                          IconIdx, _
                          ShowCmd                       ' Get Info for shortcut file...
                        
    txtLinkName.Text = LnkFile                          ' Display output...
    txtExeName.Text = ExeFile
    txtWorkDir.Text = WorkDir
    txtCmdArgs.Text = ExeArgs
    txtIconFile.Text = IconFile
    txtIconIndex.Text = Val(IconIdx)
    txtShowCmd.Text = Val(ShowCmd)
    
    Set sLnk = Nothing                                  ' Destroy object reference...
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub cmdGetPath_Click()
'---------------------------------------------------------------
    Dim rc As Long                                      ' return code
    Dim sLnk As cShellLink                              ' ShellLink class object
    Dim sfPath As String                                ' System folder path
    Dim Id As Long                                      ' ID of System folder...
'---------------------------------------------------------------
    ' Create instance of Explorer's IShellLink Interface Base Class
    Set sLnk = New cShellLink
    
    Id = cmbSysFolders.ItemData(cmbSysFolders.ListIndex)  ' Get ID from combo box
    If sLnk.GetSystemFolderPath(Me.hWnd, Id, sfPath) Then ' Get system folder path from id
        SetDefaults sfPath                                ' Update UI with new path
    End If
    
    Set sLnk = Nothing                                  ' Destroy object reference
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------
    SetDefaults (App.Path & "\")                    ' Update UI with current application path
    
    With cmbSysFolders                              ' Add ID's for system folders to combo box...
        .AddItem "DESKTOP"
        .ItemData(.NewIndex) = 0
        .AddItem "PROGRAMS"
        .ItemData(.NewIndex) = &H2
        .AddItem "Controls"
        .ItemData(.NewIndex) = &H3
        .AddItem "Printers"
        .ItemData(.NewIndex) = &H4
        .AddItem "PERSONAL"
        .ItemData(.NewIndex) = &H5
        .AddItem "FAVORITES"
        .ItemData(.NewIndex) = &H6
        .AddItem "STARTUP"
        .ItemData(.NewIndex) = &H7
        .AddItem "RECENT"
        .ItemData(.NewIndex) = &H8
        .AddItem "SENDTO"
        .ItemData(.NewIndex) = &H9
        .AddItem "BITBUCKET: RECYCLE-BIN"
        .ItemData(.NewIndex) = &HA
        .AddItem "STARTMENU"
        .ItemData(.NewIndex) = &HB
        .AddItem "DESKTOPDIRECTORY"
        .ItemData(.NewIndex) = &H10
        .AddItem "DRIVES"
        .ItemData(.NewIndex) = &H11
        .AddItem "NETWORK"
        .ItemData(.NewIndex) = &H12
        .AddItem "NETHOOD"
        .ItemData(.NewIndex) = &H13
        .AddItem "Fonts"
        .ItemData(.NewIndex) = &H14
        .AddItem "TEMPLATES"
        .ItemData(.NewIndex) = &H15
        .AddItem "COMMON_STARTMENU"
        .ItemData(.NewIndex) = &H16
        .AddItem "COMMON_PROGRAMS"
        .ItemData(.NewIndex) = &H17
        .AddItem "COMMON_STARTUP"
        .ItemData(.NewIndex) = &H18
        .AddItem "COMMON_DESKTOPDIRECTORY"
        .ItemData(.NewIndex) = &H19
        .AddItem "APPDATA"
        .ItemData(.NewIndex) = &H1A
        .AddItem "PRINTHOOD"
        .ItemData(.NewIndex) = &H1B
        
        .ListIndex = 0
    End With
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub SetDefaults(pth As String)
'---------------------------------------------------------------
    Dim AppPath As String                                   ' Current Application path
'---------------------------------------------------------------
    AppPath = App.Path                                      ' Get current path
    
    If (Right$(AppPath, 1) <> "\") Then AppPath = AppPath & "\" ' Fix application path if necessary
    If (Right$(pth, 1) <> "\") Then pth = pth & "\"         ' Fix path if necessary
    
    txtLinkName.Text = pth & "testlink.lnk"                 ' Create a full path name for link file
    txtExeName.Text = AppPath & App.EXEName & ".exe"        ' Create a full path name for applicaton exe name
    txtWorkDir.Text = AppPath                               ' Set default working directory
    txtCmdArgs.Text = "-ARG1 -ARG2"                         ' Set default arguments
    txtIconFile.Text = txtExeName.Text                      ' Set default IconFile name to default exename
    txtIconIndex.Text = CStr(1)                             ' Set default Icon Index val
    txtShowCmd.Text = CStr(7)                               ' set default showcommand val
    txtProgramGroup.Text = pth & "Test Link Program Group"  ' Set default Program group name
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------
