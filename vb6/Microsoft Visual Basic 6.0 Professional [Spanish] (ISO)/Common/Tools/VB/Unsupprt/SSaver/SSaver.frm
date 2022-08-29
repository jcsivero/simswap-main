VERSION 5.00
Begin VB.Form frmSSaver 
   BorderStyle     =   0  'None
   Caption         =   "VB 5 - Screen Saver"
   ClientHeight    =   2790
   ClientLeft      =   2460
   ClientTop       =   1935
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SSaver.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer ssTimer 
      Interval        =   50
      Left            =   3930
      Top             =   2250
   End
End
Attribute VB_Name = "frmSSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------
' Declare Variables and Constants
'-----------------------------------------------------------------
Private ssEng As ssEngine                   ' Sprite builder engine
'''Private Sprite() As ssSprite                ' Array of active sprites...

Const BMPXUNITS = 1                         ' # sprite frames on the x axis
Const BMPYUNITS = 46                        ' # sprite frames on the y axis
Const IDB_BITMAP = 101                      ' Res File bitmap image ID

'-----------------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------------
    Dim Idx As Long                         ' Loop index
    Dim ScaleSize As Single                 ' New sprite size (relative to resource size)
'-----------------------------------------------------------------
    InitDeskDC DeskDC, DeskBmp, DispRec     ' Initialize desktop image information...
    LoadSettings                            ' Load saver registry settings...

#If Not DebugOn Then                        ' Don't do if debugging...
    ' Subclass windproc...(not currently used)
'   PrevWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf SubWndProc)
#End If
    
    Set ssEng = New ssEngine                    ' Create new Sprite builder engine
    ReDim gSSprite(gSpriteCount - 1) As ssSprite ' Resize active sprite array...

    For Idx = LBound(gSSprite) To UBound(gSSprite)  ' Initialize each sprite...
        If gSizeRND Then                          ' Determine if sprite size is random...
            ' Randomize sprite size...
            ScaleSize = (((MAX_SPRITESIZE - MIN_SPRITESIZE) * Rnd) + MIN_SPRITESIZE) / 100
        Else
            ScaleSize = gSpriteSize / 100   ' Scale ALL sprite sizes to Registry setting...
        End If
        ' Create new active sprite...
        Set gSSprite(Idx) = ssEng.CreateSprite(Me, DeskDC, IDB_BITMAP, vbBlack, _
                                           BMPXUNITS * BMPYUNITS, BMPXUNITS, BMPYUNITS, _
                                           ScaleSize, ScaleSize, Idx)
                                           
        With gSSprite(Idx)                   ' Initialize sprite settings...
            .BdrX = DispRec.Right - CLng(.uWidth * 0.8)     ' calculate width of display
            .BdrY = DispRec.Bottom - CLng(.uHeight * 0.8)   ' calculate height of display
            
            If gSpeedRND Then               ' Determine if speed of sprite should be random
                .Dx = CLng(((20 * Rnd) + 1) * ScaleSize)    ' Randomize horizontal speed
                .Dy = CLng(((20 * Rnd) + 1) * ScaleSize)    ' Randomize verticle speed
            Else
                .Dx = CLng(gSpriteSpeed * ScaleSize) + 1    ' Use speed setting from registry setting...
                .Dy = .Dx                                   ' Use speed setting from registry setting...
            End If
            
            .x = CLng(.BdrX * Rnd) + 1      ' Randomly place sprite on x axis
            .y = CLng(.BdrY * Rnd) + 1      ' Randomly place sprite on y axis
            .DDx = 1                        ' (Sprite acceleration) Reserved for future use...
            .DDy = 1                        ' (Sprite acceleration) Reserved for future use...
            .TRACERS = gTracers             ' Set tracers option from registry setting
        End With
    Next
    
    If gRefreshRND Then                     ' Set timer animation interval
        ' Use random animation interval
        ssTimer.Interval = CLng((MAX_REFRESHRATE - MIN_REFRESHRATE + 1) * Rnd) + MIN_REFRESHRATE
    Else
        ' Get animation interval from registry setting...
        ssTimer.Interval = (MAX_REFRESHRATE - MIN_REFRESHRATE) + 2 - gRefreshRate
    End If
    
    ssTimer.Enabled = True                  ' Start timer (animate active sprites)
    
    Set ssEng = Nothing                     ' Destroy sprite creation engine
#If Not DebugOn Then                        ' Don't do if debugging...
    If (RunMode = RM_NORMAL) Then ShowCursor 0  ' Hide MousePointer.
#End If
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

Private Sub Form_Click()
    If (RunMode = RM_NORMAL) Then Unload Me ' Terminate if form is clicked
End Sub
Private Sub Form_DblClick()
    If (RunMode = RM_NORMAL) Then Unload Me ' Terminate if form is double clicked
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (RunMode = RM_NORMAL) Then Unload Me ' Terminate if a key is pressed down...
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If (RunMode = RM_NORMAL) Then Unload Me ' Terminate if a key is pressed
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (RunMode = RM_NORMAL) Then Unload Me ' Terminate if form mouse is down
End Sub
'-----------------------------------------------------------------
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'-----------------------------------------------------------------
    Static X0 As Integer, Y0 As Integer
'-----------------------------------------------------------------
    If (RunMode = RM_NORMAL) Then           ' Determine screen saver mode
        If ((X0 = 0) And (Y0 = 0)) Or _
           ((Abs(X0 - x) < 5) And (Abs(Y0 - y) < 5)) Then ' small mouse movement...
            X0 = x                          ' Save current x coordinate
            Y0 = y                          ' Save current y coordinate
            Exit Sub                        ' Exit
        End If
    
        Unload Me                           ' Large mouse movement (terminate screensaver)
    End If
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

Private Sub Form_Paint()
    PaintDeskDC DeskDC, DeskBmp, hwnd           ' Repaint desktop bitmap to form
End Sub

'-----------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------
    Dim Idx As Integer                          ' Array index
'-----------------------------------------------------------------
    ' [* YOU MUST TURN OFF THE TIMER BEFORE DESTROYING THE SPRITE OBJECT *]
    ssTimer.Enabled = False                     ' [* YOU MAY DEADLOCK!!! *]
'   Set gSpriteCollection = Nothing             ' Not sure if this would work...

    For Idx = LBound(gSSprite) To UBound(gSSprite) ' For each active sprite...
        Set gSSprite(Idx) = Nothing               ' Destroy active sprite
    Next

#If Not DebugOn Then                            ' Don't execute when debugging
    ' Subclass windproc...(not currently used)
'   SetWindowLong Me.hwnd, GWL_WNDPROC, PrevWndProc
#End If
    DelDeskDC DeskDC                            ' Cleanup the DeskDC (Memleak will occure if not done)
    
    If (RunMode = RM_NORMAL) Then ShowCursor -1 ' Show MousePointer
    Screen.MousePointer = vbDefault             ' Reset MousePointer
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Private Sub ssTimer_Timer()
'-----------------------------------------------------------------
    Dim Idx As Integer                            ' Array index
'-----------------------------------------------------------------
    For Idx = LBound(gSSprite) To UBound(gSSprite)  ' For each active sprite...
        gSSprite(Idx).AutoMove                      ' Automatically move active sprite
    Next
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------
