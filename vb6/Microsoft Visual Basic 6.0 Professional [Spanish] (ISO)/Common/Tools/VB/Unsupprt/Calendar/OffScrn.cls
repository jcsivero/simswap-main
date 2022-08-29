VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OffScreenDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' OffScreenDC.cls
'----------------------------------------------------------------------
' Implementation file for OffScreenDC class
' This class represents an off screen DC that is useful
' for enabling flick-free and smooth repainting of things
' like controls.
'
' There are also a couple helper methods that do interesting
' GDI operations like drawing 3d rectangles and fast rectangles
'----------------------------------------------------------------------
' Copyright (c) 1996, Microsoft Corporation
'              All Rights Reserved
'
' Information Contained Herin is Proprietary and Confidential
'----------------------------------------------------------------------

Option Explicit

'======================================================================
' Public Enumerations
'======================================================================
Public Enum CaptionAlignments
    caCenterCenter
    cacenterleft
    caCenterRight
    caTopCenter
    caTopLeft
    caTopRight
    caBottomCenter
    caBottomLeft
    caBottomright
End Enum

Public Enum Appearances
    Raised
    Flat
    Sunken
    Selected
End Enum

'======================================================================
' Private Constants
'======================================================================
Private Const BORDER_WIDTH As Long = 3

'======================================================================
' Private Data Members
'======================================================================
Private mhdcWork As Long        'off-screen HDC
Private mhdcCtl As Long         'actual HDC of the control
Private mhbmpOld As Long        'hBmp of the old bitmap in the off-sceen DC
Private mfntCurrent As IFont    'font to use when drawing text
Private mhfntOld As Long        'hFont of the old font in the off-screen dc

Private mcxCtlWidth As Long     'width of the control's surface
Private mcyCtlHeight As Long    'height of the control's surface

Private mrgb3DFace As Long      'color to use for the 3d face
Private mrgb3DHighlight As Long 'color to use for the 3d highlight
Private mrgb3DShadow As Long    'color to use for the 3d shadow

'======================================================================
' Initialize and Terminate Events
'======================================================================

'----------------------------------------------------------------------
' Class_Terminate()
'----------------------------------------------------------------------
' Purpose:  Called when the object is destroyed--do clean-up work
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub Class_Terminate()
    Dim hBmp As Long
    
    'if our handles are NULL then just get out
    If mhdcWork <> 0 Then
        'select the old font back into the off-screen dc
        SelectObject mhdcWork, mhfntOld
        
        'select the old bitmap back into the off-screen DC
        hBmp = SelectObject(mhdcWork, mhbmpOld)
        
        'delete the bitmap we were using
        DeleteObject hBmp
        
        'and now delete the off-screen DC to totally clean up
        DeleteDC mhdcWork
    End If 'we were initialized
    
End Sub 'Class_Terminate()

'======================================================================
' Public Methods and Properties
'======================================================================

'----------------------------------------------------------------------
' BackColor Get/Let
'----------------------------------------------------------------------
' Purpose:  To get and let the current background color of the DC
'----------------------------------------------------------------------
Public Property Get BackColor() As Long
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'return the current background color
    BackColor = GetBkColor(mhdcWork)
End Property 'BackColor Get

Public Property Let BackColor(rgbNew As Long)
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'set the new background color
    SetBkColor mhdcWork, rgbNew
End Property 'BackColor Let

'----------------------------------------------------------------------
' TextColor Get/Let
'----------------------------------------------------------------------
' Purpose:  To get and let the current text color of the DC
'----------------------------------------------------------------------
Public Property Get TextColor() As Long
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'return the current Text color
    TextColor = GetTextColor(mhdcWork)
End Property 'TextColor Get

Public Property Let TextColor(rgbNew As Long)
    'assert that we are initialized
    Debug.Assert mhdcWork <> 0
    
    'set the new text color
    SetTextColor mhdcWork, rgbNew
End Property 'TextColor Let

'----------------------------------------------------------------------
' Font Get/Set
'----------------------------------------------------------------------
' Purpose:  To get and set the current font to use on the DC
'----------------------------------------------------------------------
Public Property Get Font() As StdFont
    'just return the reference we currently are holding
    Set Font = mfntCurrent
End Property 'Font Get

Public Property Set Font(NewFont As StdFont)
    'make sure we're initialized first
    'must call Initialize before setting the font!
    Debug.Assert (mhdcWork <> 0)
    
    'below we will set a local member variable equal to the
    'object passed in.  Even though the type passed in is a
    'StdFont, our member variable is of type IFont.  A StdFont
    'can be casted (QI) to an IFont, and the IFont interface gives
    'us access to the hFont property, which we need in order to
    'set the current font of the off-screen device context.
        
    'if this is the first time the font is being set,
    'grab the existing hFont handle so we can restore it
    'before deleting the DC
    If mfntCurrent Is Nothing Then
        Set mfntCurrent = NewFont
        mhfntOld = SelectObject(mhdcWork, mfntCurrent.hFont)
    Else
        Set mfntCurrent = NewFont
        
        'if this is being set to Nothing, restore the old font
        If mfntCurrent Is Nothing Then
            SelectObject mhdcWork, mhfntOld
        Else
            SelectObject mhdcWork, mfntCurrent.hFont
        End If 'new font is nothing
        
    End If 'first time setting font

End Property 'Font Set

'----------------------------------------------------------------------
' 3D Colors Properties
'----------------------------------------------------------------------
' Purpose:  To return the RGB values for 3d colors
'----------------------------------------------------------------------
Public Property Get ThreeDFaceColor() As Long
    ThreeDFaceColor = mrgb3DFace
End Property

Public Property Get ThreeDHighlightColor() As Long
    ThreeDHighlightColor = mrgb3DHighlight
End Property

Public Property Get ThreeDShadowColor() As Long
    ThreeDShadowColor = mrgb3DShadow
End Property

'----------------------------------------------------------------------
' Initialize()
'----------------------------------------------------------------------
' Purpose:  To initialize the object with the screen DC from which we
'           will create the off-screen DC
' Inputs:   The user control
' Outputs:  none
'----------------------------------------------------------------------
Public Sub Initialize(CtlHdc As Long, CtlWidth As Long, CtlHeight As Long)
    Dim hBmp As Long
    
    'assert that the inputs are valid
    'and that we haven't already called Initialize
    Debug.Assert (CtlHdc <> 0)
    Debug.Assert (mhdcWork = 0)
    
    'store the HDC of the control in our private variable
    mhdcCtl = CtlHdc
    
    'capture the width and height of the control
    mcxCtlWidth = CtlWidth
    mcyCtlHeight = CtlHeight
    
    'create the off-sceen DC
    mhdcWork = CreateCompatibleDC(mhdcCtl)
    
    'create a compatible bitmap from the control DC
    'that is the same size as the control itself
    hBmp = CreateCompatibleBitmap(mhdcCtl, mcxCtlWidth, mcyCtlHeight)
    
    'select that new bitmap into the off-screen DC
    'and save the old bitmap handle so we can reselect
    'it back in before we destroy the off-screen DC
    mhbmpOld = SelectObject(mhdcWork, hBmp)
End Sub 'Initialize()

'----------------------------------------------------------------------
' FillRect()
'----------------------------------------------------------------------
' Purpose:  To fill a rectangle on the off-screen DC with a specified
'           color in a fast way
' Inputs:   The rectangle to fill and color to use
' Outputs:  none
'----------------------------------------------------------------------
Public Sub FillRect(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long, rgbColor As Long, Optional sCaption As String = "", Optional CaptionAlign As CaptionAlignments = caCenterCenter)
    Dim nX As Long              'X for drawing caption text
    Dim nY As Long              'Y for drawing caption text
    Dim rc As RECT              'rect struct to pass to the GDI
    Dim szTextExtent As Size    'pixel size of caption
    
    'assert that we've been initialized already
    'and check the inputs
    Debug.Assert mhdcWork <> 0
    Debug.Assert rgbColor >= 0
    
    'set the back color of the DC to the color desired
    Me.BackColor = rgbColor
    
    'calculate the caption X and Y (centered) if the caption
    'is not an empty string
    If Len(sCaption) > 0 Then
        
        'get the pixel width of the Caption
        GetTextExtentPoint32 mhdcWork, sCaption, Len(sCaption), szTextExtent
        
        'determine the X value based on the alignment chosen
        Select Case CaptionAlign
            Case caCenterCenter, caTopCenter, caBottomCenter
                nX = ((nWidth - szTextExtent.cx) \ 2) + nLeft
            
            Case caCenterRight, caTopRight, caBottomright
                nX = nWidth - BORDER_WIDTH - szTextExtent.cx + nLeft
            
            Case cacenterleft, caTopLeft, caBottomLeft
                nX = nLeft + BORDER_WIDTH
        End Select
        
        'determine the Y value base on the alignment chosen
        Select Case CaptionAlign
            Case caCenterCenter, caCenterRight, cacenterleft
                nY = ((nHeight - szTextExtent.cy) \ 2) + nTop
                
            Case caTopCenter, caTopLeft, caTopRight
                nY = nTop + BORDER_WIDTH
            
            Case caBottomCenter, caBottomLeft, caBottomright
                nY = nHeight - BORDER_WIDTH - szTextExtent.cy + nTop
                
        End Select
    End If 'caption is not ""
    
    'assign the input values to the rect struct
    rc.Left = nLeft
    rc.Top = nTop
    rc.Right = nWidth + nLeft
    rc.Bottom = nHeight + nTop
    
    'ExtTextOut is one of the fastest ways to fill a rectangular
    'area on a DC and is used here to fill our rect
    ExtTextOut mhdcWork, nX, nY, ETO_OPAQUE + ETO_CLIPPED, rc, sCaption, Len(sCaption), 0
    
End Sub 'FillRect

'----------------------------------------------------------------------
' Draw3DRect()
'----------------------------------------------------------------------
' Purpose:  To draw a 3D looking rectangle on the off-screen DC
' Inputs:   The rectangle to make 3d and optionally a caption to
'           display centered in the rect
' Outputs:  none
'----------------------------------------------------------------------
Public Sub Draw3DRect(ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional sCaption As String = "", Optional CaptionAlign As CaptionAlignments = caCenterCenter, Optional Appearance As Appearances = Raised)
    Dim rgbLowerRight As Long   'color to use for the lower right
    Dim rgbUpperLeft As Long    'color to use for the upper left
    
    'assert that we've been initialized already
    'and check the inputs
    Debug.Assert mhdcWork <> 0
    
    'if we haven't gotten the system colors for 3d effects
    'get them first
    If mrgb3DFace = 0 Then
        mrgb3DFace = GetSysColor(COLOR_BTNFACE)
        mrgb3DHighlight = GetSysColor(COLOR_BTNHIGHLIGHT)
        mrgb3DShadow = GetSysColor(COLOR_BTNSHADOW)
    End If
            
    'set the lower-right and upper-left colors based on the
    'desired appearance
    Select Case Appearance
        Case Flat
            rgbLowerRight = mrgb3DShadow
            rgbUpperLeft = mrgb3DShadow
        
        Case Raised
            rgbLowerRight = mrgb3DShadow
            rgbUpperLeft = mrgb3DHighlight
            
        Case Sunken
            rgbLowerRight = mrgb3DHighlight
            rgbUpperLeft = mrgb3DShadow
        
        Case Selected
            rgbLowerRight = mrgb3DHighlight
            rgbUpperLeft = vbBlack
        
    End Select
    
    'fill the rect with the shadow color (or hightlight if sunken)
    Me.FillRect nLeft, nTop, nWidth, nHeight, rgbLowerRight
    
    'now pull the right and bottom edges in by 1 pixel
    nWidth = nWidth - 1
    nHeight = nHeight - 1
    
    'fill the rect with the 3d highlight color (or shadow if sunken)
    Me.FillRect nLeft, nTop, nWidth, nHeight, rgbUpperLeft
    
    'finally pull in the left and top edges by 1 pixel
    nLeft = nLeft + 1
    nTop = nTop + 1
    nWidth = nWidth - 1
    nHeight = nHeight - 1
    
    'change the color to the 3d face color
    'and fill the rect passing the desired caption
    Me.FillRect nLeft, nTop, nWidth, nHeight, mrgb3DFace, sCaption, CaptionAlign

    'if the appearance setting was Selected, invert the rect
    If Appearance = Selected Then
        InvertRect nLeft, nTop, nWidth, nHeight
    End If 'appearance = selected
    
End Sub 'Draw3dRect

'----------------------------------------------------------------------
' InvertRect()
'----------------------------------------------------------------------
' Purpose:  To invert a particular rect on the bitmap
' Inputs:   The area to invert
' Outputs:  none
'----------------------------------------------------------------------
Public Sub InvertRect(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long)
    Dim rc As RECT
    
    rc.Left = nLeft
    rc.Top = nTop
    rc.Right = nLeft + nWidth
    rc.Bottom = nTop + nHeight
    
    Utils.InvertRect mhdcWork, rc
End Sub 'InvertRect()

'----------------------------------------------------------------------
' BlastToScreen()
'----------------------------------------------------------------------
' Purpose:  Blasts the contents of the off-screen DC to the control's
'           on-screen surface
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Public Sub BlastToScreen(Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1)
    If Width = -1 Then Width = mcxCtlWidth
    If Height = -1 Then Height = mcyCtlHeight
    
    'use bitblt to blast the contents of the off-screen dc to the control
    BitBlt mhdcCtl, Left, Top, Width, Height, mhdcWork, _
            Left, Top, SRCCOPY
End Sub 'BlastToScreen()

'======================================================================
' Private Helper Methods
'======================================================================

