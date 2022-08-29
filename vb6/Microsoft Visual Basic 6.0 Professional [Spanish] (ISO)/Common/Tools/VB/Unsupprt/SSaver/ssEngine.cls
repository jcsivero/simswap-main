VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ssEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------
Public Function CreateSprite(Display As Form, hDisplay As Long, _
                             ResID As Long, MaskCol As Long, picCount As Long, _
                             xUnits As Long, yUnits As Long, _
                             Optional ScaleW As Single, _
                             Optional ScaleH As Single, _
                             Optional Idx As Long) As ssSprite
'-----------------------------------------------------------------
    Dim cSprite As ssSprite                         ' ScreenSaver sprite
    Dim RatioX As Single, RatioY As Single          ' X and Y shrinkage factors
'-----------------------------------------------------------------
    Randomize Timer                                 ' Set the randomization seed
    RatioX = 1                                      ' Set the x ratio base
    RatioY = 1                                      ' Set the y ratio base
    Display.ScaleMode = vbPixels                    ' Set the display form scalemode to pixels
    
    Set cSprite = New ssSprite                      ' Create a new screen saver sprite
    With cSprite
        .DestHDC = Display.hdc                      ' Set the display hdc
        .ScreenW = (Screen.Width \ Screen.TwipsPerPixelX)  ' Calculate the screen width in pixels
        .ScreenH = (Screen.Height \ Screen.TwipsPerPixelY) ' Calculate the screen height in pixels

        If (gSprite.ResID <> ResID) Then            ' Only load resource if it isn't already loaded
            gSprite.ResID = ResID                   ' Save resource id
            Set gSprite.Sprite = LoadResPicture(ResID, vbResBitmap) ' Load and save resource bitmap
        End If
        
        If (.ScreenW <> 0) Then RatioX = (DispRec.Right - DispRec.Left) / .ScreenW ' Scale x ratio based on display size
        If (.ScreenH <> 0) Then RatioY = (DispRec.Bottom - DispRec.Top) / .ScreenH ' Scale y ratio based on display size
        If (ScaleW = 0) Then ScaleW = 1             ' Fix scale width
        If (ScaleH = 0) Then ScaleH = 1             ' Fix scale height
        RatioX = RatioX * ScaleW                    ' Adjust X ratio based on sprite scale width
        RatioY = RatioY * ScaleH                    ' Adjust Y ratio based on sprite scale height
        
        .hBitmap = ShrinkBmp(.DestHDC, gSprite.Sprite.Handle, RatioX, RatioY) ' Shrink animated sprite frames...
        
        .SprtW = CLng(Display.ScaleX(gSprite.Sprite.Width, vbHimetric, vbPixels) * RatioX)  ' Calc scaled source image width
        .SprtH = CLng(Display.ScaleY(gSprite.Sprite.Height, vbHimetric, vbPixels) * RatioY) ' Calc scaled source image height
        
        .xUnits = xUnits                            ' Save # of horizontal frames
        .yUnits = yUnits                            ' Save # of verticle frames
        .uWidth = .SprtW \ xUnits                   ' calculate single frame width
        .uHeight = .SprtH \ yUnits                  ' calculate single frame height
        .idxMin = 0                                 ' initialize frame index = 0
        .idxMax = picCount - 1                      ' init max frame index = # frames - 1
        .bmpIdx = (.idxMax - .idxMin) * Rnd + .idxMin ' randomize the first picture frame to be displayed
        .hDisplayBack = hDisplay                    ' save handle to the display bitmap...
        .MASKCOLOR = MaskCol                        ' save the bitmap mask color...
        .Mass = CLng(BASE_MASS * ScaleW * ScaleH)   ' calculate mass based on scaled surface area
        .Index = Idx                                ' save index possition in global array...
'''        gSpriteCollection.Add cSprite               ' add sprite to global collection...
    End With
    
    Set CreateSprite = cSprite                      ' return sprite reference...
    Set cSprite = Nothing                           ' destroy local sprite reference...
'-----------------------------------------------------------------
End Function
'-----------------------------------------------------------------
