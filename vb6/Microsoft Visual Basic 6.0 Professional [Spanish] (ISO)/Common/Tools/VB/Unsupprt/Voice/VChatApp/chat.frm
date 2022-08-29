VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.0#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inter Net Voice"
   ClientHeight    =   3030
   ClientLeft      =   2625
   ClientTop       =   1530
   ClientWidth     =   4170
   FillColor       =   &H00808080&
   Icon            =   "chat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4170
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture3 
      Height          =   2565
      Left            =   60
      ScaleHeight     =   2505
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   390
      Width           =   4005
      Begin VB.CommandButton cmdTalk 
         Caption         =   "&Talk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   540
         TabIndex        =   4
         Top             =   2130
         Width           =   2865
      End
      Begin VB.ListBox ConnectionList 
         Height          =   1065
         ItemData        =   "chat.frx":0442
         Left            =   60
         List            =   "chat.frx":0444
         TabIndex        =   2
         Top             =   840
         Width           =   3825
      End
      Begin VB.ComboBox txtServer 
         Height          =   315
         ItemData        =   "chat.frx":0446
         Left            =   630
         List            =   "chat.frx":0448
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   3255
      End
      Begin VB.Image outLight 
         Height          =   345
         Left            =   3510
         Picture         =   "chat.frx":044A
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   345
      End
      Begin VB.Image inLight 
         Height          =   345
         Left            =   60
         Picture         =   "chat.frx":0754
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conference List:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   600
         Width           =   1425
      End
      Begin VB.Image imgStatus 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   90
         Picture         =   "chat.frx":0A5E
         Stretch         =   -1  'True
         Top             =   30
         Width           =   420
      End
   End
   Begin ComctlLib.Toolbar Tools 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   688
      ImageList       =   "ImgIcons"
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Call"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Hangup"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Auto Answering"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock TCPSocket 
      Index           =   0
      Left            =   3720
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin ComctlLib.ImageList ImgIcons 
      Left            =   4290
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":0EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":11BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":14D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":17EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":1B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":1E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":213C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":2456
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":2770
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":2A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "chat.frx":2DA4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_Base = "0{0E3A2BAD-DE40-11CF-8FDF-D0AF03C10000}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CLOSINGAPPLICATION As Boolean                    ' Application status flag
Public wStream As Object

'--------------------------------------------------------------
Private Sub cmdTalk_Click()                             ' Activates Audio PlayBack
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim iPort As Integer                                ' Local Port
    Dim itm As Integer                                  ' Current listitem
'--------------------------------------------------------------
    If (Not wStream.Playing And wStream.PlayDeviceFree And _
        Not wStream.Recording And wStream.RecDeviceFree) Then ' Validate Audio Device Status
        wStream.Playing = True                          ' Turn Playing Status On
        cmdTalk.Caption = "&Playing"                    ' Modify Button Status Caption
        Screen.MousePointer = vbHourglass               ' Set Pointer To HourGlass
        
        iPort = wStream.StreamInQueue
        Do While (iPort <> NULLPORTID)                  ' While socket ports have data to playback
            inLight.Picture = ImgIcons.ListImages(speakON).Picture ' Flash playback image
            inLight.Refresh                             ' Repaint picture image
            
            For itm = 0 To ConnectionList.ListCount - 1 ' Search for listitem currently playing sound data
                If (ConnectionList.ItemData(itm) = iPort) Then ' If a match is found...
                    ConnectionList.TopIndex = itm       ' Set that listitem to top of listbox
                    ConnectionList.Selected(itm) = True ' Select listitem to show who is currently talking...
                    Exit For                            ' Quit listitem search
                End If
            Next                                        ' Check next listitem
            
            rc = wStream.PlayWave(Me.hWnd, iPort)       ' Play wave data in iPort...
            Call wStream.RemoveStreamFromQueue(iPort)   ' Remove PortID From PlayWave Queue
            iPort = wStream.StreamInQueue
            
            inLight.Picture = ImgIcons.ListImages(speakOFF).Picture ' Show done talking image...
            inLight.Refresh                             ' Repaint image...
        Loop                                            ' Search for next socket in playback queue
        
        ConnectionList.TopIndex = 0                     ' Reset top image...
        If (ConnectionList.ListCount > 0) Then
            ConnectionList.Selected(0) = True           ' Deselect previously listitem
            ConnectionList.Selected(0) = False          ' Deselect currently selected listitem
        End If
        Screen.MousePointer = vbDefault                 ' Set Pointer To Normal
        cmdTalk.Caption = "&Talk"                       ' Modify Button Status Caption
        wStream.Playing = False                         ' Turn Playing Status Off
    End If
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Private Sub cmdTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Activates Audio Recording...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
'--------------------------------------------------------------
    If (Not wStream.Playing And _
        Not wStream.Recording And _
            wStream.RecDeviceFree And _
            wStream.PlayDeviceFree) Then          ' Check Audio Device Status
        wStream.Recording = True                                ' Set Recording Flag
        cmdTalk.Caption = "&Talking"                    ' Update Button Status To "Talking"
        Screen.MousePointer = vbHourglass               ' Set Hourglass
        outLight.Picture = ImgIcons.ListImages(mikeON).Picture ' Show outgoing message image
        outLight.Refresh                                ' Repaint image
        
        rc = wStream.RecordWave(Me.hWnd, TCPSocket)     ' Record voice and send to all connected sockets
        
        outLight.Picture = ImgIcons.ListImages(mikeOFF).Picture ' Show done image
        outLight.Refresh                                ' Repaint image
        Screen.MousePointer = vbDefault                 ' Reset Mouse Pointer
        cmdTalk.Caption = "&Talk"                       ' Reset Button Status
        
        If Not wStream.Playing And _
               wStream.PlayDeviceFree And _
               wStream.RecDeviceFree Then               ' Is Audio Device Free?
            Call cmdTalk_Click                          ' Active Playback Of Any Inbound Messages...
        End If
    End If
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

Private Sub cmdTalk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    wStream.Recording = False                           ' Stop Recording
End Sub

Private Sub connectionlist_Click()
    Tools.Buttons(tbHANGUP).Enabled = True
End Sub

'--------------------------------------------------------------
Private Sub ConnectionList_DblClick()
'--------------------------------------------------------------
    Dim MemberID As String                              ' (Server)(TCPidx)(RemoteIP)
    Dim Idx As Long                                     ' TCP idx
'--------------------------------------------------------------
    If (ConnectionList.Text = "") Then Exit Sub
    MemberID = ConnectionList.List(ConnectionList.ListIndex) ' Get The Conversation MemberID String From List Box
    
    Call GetIdxFromMemberID(TCPSocket, MemberID, Idx)  ' Get TCP idx From Member ID
    Call RemoveConnectionFromList(TCPSocket(Idx), ConnectionList) ' Clear ListBox Entry(s)...
    Call Disconnect(TCPSocket(Idx))                     ' Disconnect Socket Connection
    Unload TCPSocket(Idx)                               ' Destroy socket instance

    cmdTalk.Enabled = (ConnectionList.ListCount > 0)    ' Enable/Disable Talk Button...
    Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    If Not cmdTalk.Enabled Then
        inLight.Picture = ImgIcons.ListImages(speakNO).Picture
        outLight.Picture = ImgIcons.ListImages(mikeNO).Picture
    End If
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim Idx As Long                                     ' Current TCP idx variable
    Dim TCPidx As Long                                  ' Newly created TCP idx value
'--------------------------------------------------------------
    CLOSINGAPPLICATION = False                          ' Set status to not closing
    Call InitServerList(txtServer)                      ' Get Common Servers List
    txtServer.Text = txtServer.List(0)                  ' Display First Name In The List
    imgStatus = ImgIcons.ListImages(phoneHungUp).Picture ' Change Icon To Phone HungUp
    
    Set wStream = CreateObject("WaveStreaming.WaveStream")
    Call wStream.InitACMCodec(WAVE_FORMAT_GSM610, TIMESLICE)
'   Call wStream.InitACMCodec(WAVE_FORMAT_ADPCM, TIMESLICE)
'   Call wStream.InitACMCodec(WAVE_FORMAT_MSN_AUDIO, TIMESLICE)
'   Call wStream.InitACMCodec(WAVE_FORMAT_PCM, TIMESLICE)

    cmdTalk.Enabled = False                             ' Disable Until Connect
    Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    inLight.Picture = ImgIcons.ListImages(speakNO).Picture
    outLight.Picture = ImgIcons.ListImages(mikeNO).Picture

    Call Listen(TCPSocket(0))                           ' Listen For TCP Connection
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------
    Dim Idx As Long                                     ' TCP socket index
    Dim Socket As Winsock                                   ' TCP socket control
'--------------------------------------------------------------
    CLOSINGAPPLICATION = True                           ' Set status flag to closing...
    For Each Socket In TCPSocket                        ' For each socket instance
        Call Disconnect(Socket)                         ' Close connection/listen
    Next                                                ' Next Cntl
    Set wStream = Nothing
    End                                                 ' End Program
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------



'--------------------------------------------------------------
Private Sub TCPSocket_Close(Index As Integer)
' Closing Current TCP Connection...
'--------------------------------------------------------------
    Call RemoveConnectionFromList(TCPSocket(Index), ConnectionList) ' Remove Connection From List
    Call Disconnect(TCPSocket(Index))                           ' Close Port Connection...
    
    cmdTalk.Enabled = (ConnectionList.ListCount > 0)            ' Enable/Disable Talk Button...
    If Not cmdTalk.Enabled Then
        inLight.Picture = ImgIcons.ListImages(speakNO).Picture
        outLight.Picture = ImgIcons.ListImages(mikeNO).Picture
    End If
    
    Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    If cmdTalk.Enabled Then
        imgStatus = ImgIcons.ListImages(phoneHungUp).Picture    ' Show Phone HungUp Icon...
    End If
    
    Unload TCPSocket(Index)                                     ' Destroy socket instance
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Private Sub TCPSocket_Connect(Index As Integer)
' TCP Connection Has Been Accepted And Is Open...
'--------------------------------------------------------------
    Call AddConnectionToList(TCPSocket(Index), ConnectionList) ' Add New Connection To List
    
    imgStatus = ImgIcons.ListImages(phoneRingIng).Picture   ' Show Phone Ringing Icon
    Call ResPlaySound(RingOutId)
    imgStatus = ImgIcons.ListImages(phoneAnswered).Picture  ' Show Phone Answered Icon
    cmdTalk.Enabled = True                                  ' Enabled For Connection...
    Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Private Sub TCPSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
' Accepting Inbound TCP Connection Request...
'--------------------------------------------------------------
    Dim rc As Long
    Dim Idx As Long
    Dim RemHost As String
'--------------------------------------------------------------
    If (TCPSocket(Index).RemoteHost <> "") Then
        RemHost = UCase(TCPSocket(Index).RemoteHost)
    Else
        RemHost = UCase(TCPSocket(Index).RemoteHostIP)
    End If
    
    If (Tools.Buttons(tbAUTOANSWER).Value = tbrUnpressed) Then
        rc = MsgBox("Incomming call from [" & RemHost & "]..." & vbCrLf & _
                    "Do you wish to answer?", vbYesNo)          ' Prompt user to answer...
    Else
        rc = vbYes
    End If
                    
    If (rc = vbYes) Then
        Idx = InstanceTCP(TCPSocket)                            ' Instance TCP Control...
        If (Idx > 0) Then                                       ' Validate that control instance was created...
            TCPSocket(Idx).LocalPort = 0                        ' Set local port to 0, in order to get next available port.
            Call TCPSocket(Idx).Accept(requestID)               ' Accept connection
            Call AddConnectionToList(TCPSocket(Idx), ConnectionList) ' Add New Connection To List
            
            imgStatus = ImgIcons.ListImages(phoneRingIng).Picture  ' Show Phone Ringing Icon
            Call ResPlaySound(RingInId)
            imgStatus = ImgIcons.ListImages(phoneAnswered).Picture ' Show Phone Answered Icon
            cmdTalk.Enabled = True                                 ' Enabled For Connection...
            Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
        End If
    End If
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Private Sub TCPSocket_DataArrival(Index As Integer, ByVal BytesTotal As Long)
' Incomming Buffer On...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim WaveData() As Byte                              ' Byte array of wave data
    Static ExBytes(MAXTCP) As Long                      ' Extra bytes in frame buffer
    Static ExData(MAXTCP) As Variant                    ' Extra bytes from frame buffer
'--------------------------------------------------------------
With wStream
    If (TCPSocket(Index).BytesReceived > 0) Then        ' Validate that bytes where actually received
        Do While (TCPSocket(Index).BytesReceived > 0)   ' While data available...
            If (ExBytes(Index) = 0) Then                ' Was there leftover data from last time
                If (.waveChunkSize <= TCPSocket(Index).BytesReceived) Then ' Can we get and entire wave buffer of data
                    Call TCPSocket(Index).GetData(WaveData, vbByte + vbArray, .waveChunkSize) ' Get 1 wave buffer of data
                    Call .SaveStreamBuffer(Index, WaveData) ' Save wave data to buffer
                    Call .AddStreamToQueue(Index)       ' Queue current stream for playback
                Else
                    ExBytes(Index) = TCPSocket(Index).BytesReceived ' Save Extra bytes
                    Call TCPSocket(Index).GetData(ExData(Index), vbByte + vbArray, ExBytes(Index)) ' Get Extra data
                End If
            Else
                Call TCPSocket(Index).GetData(WaveData, vbByte + vbArray, .waveChunkSize - ExBytes(Index)) ' Get leftover bits
                ExData(Index) = MidB(ExData(Index), 1) & MidB(WaveData, 1) ' Sync wave bits...
                Call .SaveStreamBuffer(Index, ExData(Index)) ' Save the current wave data to the wave buffer
                Call .AddStreamToQueue(Index)           ' Queue the current wave stream
                ExBytes(Index) = 0                      ' Clear Extra byte count
                ExData(Index) = ""                      ' Clear Extra data buffer
            End If
        Loop                                            ' Look for next Data Chunk
        
        If (Not .Playing And .PlayDeviceFree And _
            Not .Recording And .RecDeviceFree) Then     ' Check Audio Device Status
            Call cmdTalk_Click                          ' Start PlayBack...
        End If
    End If
End With
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

Private Sub TCPSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
TCPSocket(Index).Close                                  ' Close down socket
    
    Debug.Print "TCPSocket_Error: Number:", Number
    Debug.Print "TCPSocket_Error: Scode:", Hex(Scode)
    Debug.Print "TCPSocket_Error: Source:", Source
    Debug.Print "TCPSocket_Error: HelpFile:", HelpFile
    Debug.Print "TCPSocket_Error: HelpContext:", HelpContext
    Debug.Print "TCPSocket_Error: Description:", Description
    Call DebugSocket(TCPSocket(Index))
End Sub

'--------------------------------------------------------------
Private Sub Tools_ButtonClick(ByVal Button As Button)
'--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim Idx As Long                                     ' TCP Socket control index
    Dim LocalPort As Long                               ' LocalPort Setting
    Dim RemotePort As Long                              ' RemotePort Setting
'--------------------------------------------------------------
    Select Case Button.Index
    Case tbCALL
        Idx = InstanceTCP(TCPSocket)                        ' Instance TCP Control...
        
        If (Idx > 0) Then                                   ' Did control instance get created???
            Button.Enabled = False                          ' Disable Connect Button
            ConnectionList.Enabled = False                  ' Disable connection list box
            
            On Error Resume Next
            If Not Connect(TCPSocket(Idx), txtServer.Text, VOICEPORT) Then ' Attempt to connect
                Unload TCPSocket(Idx)                       ' Connect failed unload control instance
            End If
            
            ConnectionList.Enabled = True                   ' Renable connection list box
            Button.Enabled = True                           ' Enable Connect Button
        End If
    Case tbHANGUP
        ConnectionList_DblClick
    Case tbAUTOANSWER
        If (Button.Value = tbrPressed) Then
            Button.Image = phoneHungUp
        Else
            Button.Image = phoneAnswered
        End If
    End Select
End Sub

'--------------------------------------------------------------
Private Sub txtServer_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------
    Dim Conn As Long                                        ' Index counter
'--------------------------------------------------------------
    If (KeyAscii = vbKeyReturn) Then                        ' If Return Key Was Pressed...
        For Conn = 0 To txtServer.ListCount                 ' Search Each Entry In ListBox
            If (UCase(txtServer.Text) = UCase(txtServer.List(Conn))) Then Exit Sub
        Next                                                ' If Found Then Exit
        txtServer.AddItem UCase(txtServer.Text)             ' Add Server To List
    End If
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
