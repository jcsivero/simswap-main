VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ssSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------
' Public Variables
'-----------------------------------------------------------------
Public Index As Long                                ' Global array index value
Public idxMin As Long                               ' Minimum sprite frame index value
Public idxMax As Long                               ' Maximum sprite frame index value
Public xUnits As Long                               ' # of horizontal sprite frames
Public yUnits As Long                               ' # of virtical sprite frames
Public uWidth As Long                               ' sprite frame width in pixels
Public uHeight As Long                              ' sprite frame height in pixels
Public DestHDC  As Long                             ' destination window hdc
Public hBitmap As Long                              ' handle to animation bitmap
Public hDisplayBack As Long                         ' handle to background bitmap
Public TRACERS As Boolean                           ' use tracers flag
Public MASKCOLOR As Long                            ' transparency blt color mask
Public SprtH As Long                                ' animation bitmap height in pixels
Public SprtW As Long                                ' animation bitmap width in pixels
Public Mass As Long                                 ' sprite mass(virtual)

'-----------------------------------------------------------------
' AutoMove Programmable Variables
'-----------------------------------------------------------------
Public x As Long                                    ' sprite's current screen x coordinate
Public y As Long                                    ' sprite's current screen y coordinate
Public BdrX As Long                                 ' border width
Public BdrY As Long                                 ' border height
Public Dx As Long                                   ' current x velosity
Public Dy As Long                                   ' current y velosity
Public DDx As Long                                  ' current x acceleration (= 1 not currently used)
Public DDy As Long                                  ' current Y acceleration (= 1 not currently used)
Public ScreenW As Long                              ' width of screen
Public ScreenH As Long                              ' height of screen
Public bmpIdx As Long                               ' current animated bitmap frame index

'-----------------------------------------------------------------
' Private Variables
'-----------------------------------------------------------------
Private LastX As Long                               ' previous x coordinate
Private LastY As Long                               ' previous y coordinate

'-----------------------------------------------------------------
Public Function CollisionTest(Sprite As ssSprite) As Boolean
'-----------------------------------------------------------------
    Dim l1 As Long, r1 As Long, t1 As Long, b1 As Long ' left, right, top, bottom... positions of sprite
    Dim l2 As Long, r2 As Long, t2 As Long, b2 As Long ' left, right, top, bottom... positions of sprite
'-----------------------------------------------------------------
    If (Sprite Is Me) Then Exit Function            ' don't compare sprite with itself
        
    With Me                                         ' current sprite
        l1 = .x
        t1 = .y
        r1 = l1 + .uWidth
        b1 = t1 + .uHeight
    End With
    With Sprite                                     ' other sprite
        l2 = .x
        t2 = .y
        r2 = l2 + .uWidth
        b2 = t2 + .uHeight
    End With
    
    ' Test for sprite collision
    CollisionTest = (((l2 <= l1) And (l1 <= r2)) Or _
                     ((l2 <= r1) And (r1 <= r2))) And _
                    (((t2 <= t1) And (t1 <= b2)) Or _
                     ((t2 <= b1) And (b1 <= b2)))
'-----------------------------------------------------------------
End Function
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Private Function Atn2(y As Double, x As Double) As Double
'-----------------------------------------------------------------
'-  VB implementation of the C runtime ATan2(x,y) function...
'-----------------------------------------------------------------
    If (x <> 0) Then                    ' Prevent divide by zero
        Atn2 = Atn(y / x)               ' Atan2(y,x) = Atn(y/x) when x <> 0
    Else                                ' Handle special case
        Atn2 = 2 * Atn(Sgn(y))          ' as N ~> infinity Atn(N) ~> (sign(N)*PI/2) = 2 * Atn(Sgn(y))
    End If
'-----------------------------------------------------------------
End Function
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Function ResolveCollision() As Boolean
'-----------------------------------------------------------------
'''    Dim Sprite As ssSprite
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim a As Double, cos_a As Double, sin_a As Double
    Dim vn1 As Double, vn2 As Double, vp1 As Double, vp2 As Double
    Dim vx1 As Long, vx2 As Long, vy1 As Long, vy2 As Long
    Dim m1 As Double, m2 As Double, k As Double, e As Double
    Dim vn2p1 As Double, vn2p2 As Double, temp1 As Double
    Dim Idx As Integer, First As Integer, Last As Integer
'-----------------------------------------------------------------
'''    For Each Sprite In gSpriteCollection    ' For each sprite check for...
    First = Me.Index + 1
    Last = UBound(gSSprite)
    
    For Idx = First To Last
'''        Set Sprite = gSprite(Idx)
'''        If CollisionTest(Sprite) Then       ' Check for Collision
        If CollisionTest(gSSprite(Idx)) Then     ' Check for Collision
            With gSSprite(Idx)
                ' Compute the coordinates of the centers of the objects.
                x1 = Me.x + (Me.uWidth \ 2)
                y1 = Me.y + (Me.uHeight \ 2)
                x2 = .x + (.uWidth \ 2)
                y2 = .y + (.uHeight \ 2)
            
                ' Compute the angle of the line joining the centers.
                ' a = atan2((double)(y2 - y1), (double)(x2 - x1)) (C implementation)
                a = Atn2((y2 - y1), (x2 - x1))                 ' (VB implementation)
                cos_a = Cos(a)
                sin_a = Sin(a)
            
                ' Compute the velocities normal and perpendicular
                ' to the center line.
                vx1 = Me.Dx:    vy1 = Me.Dy
                vx2 = .Dx:      vy2 = .Dy
                vn1 = (vx1 * cos_a) + (vy1 * sin_a)
                vp1 = (vy1 * cos_a) - (vx1 * sin_a)
                
                vn2 = (vx2 * cos_a) + (vy2 * sin_a)
                vp2 = (vy2 * cos_a) - (vx2 * sin_a)
    
                ' Compute the momentum along the center line.
                m1 = Me.Mass
                m2 = .Mass
                k = (m1 * vn1) + (m2 * vn2)
            
                ' Compute the energy.
                e = 0.5 * ((m1 * vn1 ^ 2) + (m2 * vn2 ^ 2))
                
                ' There are two possible solutions to the equations.
                ' Compute both and choose.
                ' <<<***Convert to long to fix Floating Point Error Bug.***>>>
                temp1 = Sqr(Fix(k ^ 2 - ((m1 / m2) + 1) * (-2 * e * m1 + k ^ 2)))
                vn2p1 = (k + temp1) / (m1 + m2)
                vn2p2 = (k - temp1) / (m1 + m2)
            
                ' Choose the solution that is not the current state.
                If (vn2p1 = vn2) Then
                    vn2 = vn2p2
                Else
                    vn2 = vn2p1
                End If
            
                ' Compute the new vn1 value.
                vn1 = (k - m2 * vn2) / m1
                
                ' Compute the new x and y velocities.
                vx1 = (vn1 * cos_a) - (vp1 * sin_a)
                vy1 = (vn1 * sin_a) + (vp1 * cos_a)
                vx2 = (vn2 * cos_a) - (vp2 * sin_a)
                vy2 = (vn2 * sin_a) + (vp2 * cos_a)
            
                Me.Dx = vx1                                     ' Save new change in x velosity
                Me.Dy = vy1                                     ' Save new change in Y velosity
                .Dx = vx2                                       ' Save new change in x velosity
                .Dy = vy2                                       ' Save new change in Y velosity

                ' Move the sprites until they are no longer in collision.
                If ((vx1 <> 0) Or (vy1 <> 0) Or (vx2 <> 0) Or (vy2 <> 0)) Then
'''                    Do While CollisionTest(Sprite)
                    Do While CollisionTest(gSSprite(Idx))
                        If ((Dx <> 0) Or (Dy <> 0)) Then        ' if 0 then update wont matter
                            UpdatePosition                      ' Move sprite out of the way
                        ElseIf ((.Dx <> 0) Or (.Dy <> 0)) Then  ' if 0 then update wont matter
                            .UpdatePosition                     ' Move sprite out of the way
                        Else
                            Exit Do                             ' Exit to avoid dead lock(infinite loop)
                        End If
                    Loop
                End If
                ResolveCollision = True                         ' Return success
            End With
        End If
    Next
'-----------------------------------------------------------------
End Function
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub UpdatePosition()
'-----------------------------------------------------------------
    x = x + Dx                                          ' Update x position
    y = y + Dy                                          ' Update y position
    
    If (x < 0) Then Dx = Abs(Dx)                        ' reverse direction when hitting a border.
    If (x > BdrX) Then Dx = -1 * Abs(Dx)
    If (y < 0) Then Dy = Abs(Dy)
    If (y > BdrY) Then Dy = -1 * Abs(Dy)
    
    Dx = Dx * DDx                                       ' acceleration sprite x velocity
    Dy = Dy * DDy                                       ' acceleration sprite y velocity
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub AutoMove()
'-----------------------------------------------------------------
    DrawNext x, y                                       ' Move sprite to next coordinate
    If Not ResolveCollision Then UpdatePosition         ' Check for collision or update current position
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

'-----------------------------------------------------------------
Public Sub DrawNext(PosX As Long, PosY As Long)
'-----------------------------------------------------------------
    Dim rc As Long
    Dim x As Long, y As Long                            ' Source indexed bmp coordinates...
    Dim x1 As Long, y1 As Long, w1 As Long, h1 As Long  ' Repaint Rectangle # 1 screen coordinates
    Dim x2 As Long, y2 As Long, w2 As Long, h2 As Long  ' Repaint Rectangle # 2 screen coordinates
'-----------------------------------------------------------------
    x = CLng((bmpIdx Mod xUnits) * (SprtW / xUnits))    ' Get next indexed bmp x coordinate
    y = CLng((bmpIdx \ xUnits) * (SprtH / yUnits))      ' Get next indexed bmp y coordinate
    
    If (TRACERS) Then                                   ' Tracers? don't clean up previous blt
        DrawTransparentBitmap DestHDC, hBitmap, MASKCOLOR, PosX, PosY, uWidth, uHeight, x, y
    Else                                                ' Clean up & calculate unused sprite space
        Select Case PosX
        Case Is < LastX                                                     ' PosX <=== LastX
            x1 = PosX + uWidth:     w1 = LastX - PosX
            x2 = LastX:             w2 = uWidth
        Case LastX                                                          ' PosX ==== LastX
            x2 = LastX:             w2 = uWidth
        Case Is > LastX                                                     ' PosX ===> LastX
            x1 = LastX:             w1 = PosX - LastX
            x2 = LastX:             w2 = uWidth
        End Select
        
        Select Case PosY
        Case Is < LastY                                                     ' PosY <=== LastY
            y1 = LastY:             h1 = uHeight - (LastY - PosY)
            y2 = PosY + uHeight:    h2 = uHeight - h1
        Case LastY                                                          ' PosY ==== LastY
            y1 = LastY:             h1 = uHeight
        Case Is > LastY                                                     ' PosY ===> LastY
            y1 = PosY:              h1 = uHeight - (PosY - LastY)
            y2 = LastY:             h2 = uHeight - h1
        End Select
    
        ' paint sprite in new position...
        DrawTransparentBitmap DestHDC, hBitmap, MASKCOLOR, PosX, PosY, uWidth, uHeight, x, y, hDisplayBack
    
        If ((LastX <> PosX) Or (LastY <> PosY)) Then ' If sprite has moved...
            ' Repaint previous unoccupied positions...
            If ((w1 > 0) And (h1 > 0)) Then BitBlt DestHDC, x1, y1, w1, h1, hDisplayBack, x1, y1, vbSrcCopy
            If ((w2 > 0) And (h2 > 0)) Then BitBlt DestHDC, x2, y2, w2, h2, hDisplayBack, x2, y2, vbSrcCopy
        End If
    End If
    
    LastX = PosX                                        ' Save previous x position
    LastY = PosY                                        ' Save previous y position
    
    If (bmpIdx < idxMax) Then                           ' Increment bitmap frame index
        bmpIdx = bmpIdx + 1
    Else                                                ' Reset to beginning
        bmpIdx = idxMin
    End If
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------
