VERSION 5.00
Begin VB.UserControl Calendar 
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   EditAtDesignTime=   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "Calendar.ctx":0000
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   232
   ToolboxBitmap   =   "Calendar.ctx":0032
   Begin VB.TextBox ctlFocus 
      Height          =   285
      Left            =   -300
      TabIndex        =   0
      Top             =   900
      Width           =   150
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   3
      ToolTipText     =   "Year"
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cbxMonth 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Month"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnNext 
      Height          =   255
      Left            =   3060
      MaskColor       =   &H000000FF&
      Picture         =   "Calendar.ctx":012C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Go To Next Month"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton btnPrev 
      Height          =   255
      Left            =   60
      MaskColor       =   &H000000FF&
      Picture         =   "Calendar.ctx":020E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Go To Previous Month"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   255
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VB Calendar Control Sample"
'----------------------------------------------------------------------
' Calendar.ctl
'----------------------------------------------------------------------
' Implementation file for the VB Calendar control sample.
' This control displays a month-at-a-time view calendar that the
' developer can use to let users view and adjust date values
'----------------------------------------------------------------------
' Copyright (c) 1996, Microsoft Corporation
'              All Rights Reserved
'
' Information Contained Herin is Proprietary and Confidential
'----------------------------------------------------------------------
Option Explicit

'======================================================================
' Public Event Declarations
'======================================================================
Public Event DateChange(ByVal OldDate As Date, ByVal NewDate As Date)
Public Event WillChangeDate(ByVal NewDate As Date, Cancel As Boolean)
Public Event DblClick()
Public Event Click()

'======================================================================
' Public Enumerations
'======================================================================
Public Enum CalendarMonths  'months of the year
    calJanuary = 1
    calFebruary
    calMarch
    calApril
    calMay
    calJune
    calJuly
    calAugust
    calSeptember
    calOctober
    calNovember
    calDecember
End Enum 'CalendarMonths

Public Enum DaysOfTheWeek
    calUseSystem = 0
    calSunday
    calMonday
    calTuesday
    calWednesday
    calThursday
    calFriday
    calSaturday
End Enum 'DaysOfTheWeek

Public Enum CalendarAreas
    calNavigationArea
    calDayNameArea
    calDateArea
    calUnknownArea
End Enum 'CalendarAreas

'Short = "F"
'Medium = "Fri"
'Long = "Friday"
Public Enum DayNameFormats
    calShortName = 0
    calMediumName
    calLongName
End Enum 'DayNameFormats

'======================================================================
' Private Constants
'======================================================================
Private Const NUMCOLS As Long = 7           'number of cols in grid
Private Const NUMROWS As Long = 6           'number of rows in grid
Private Const NUMMONTHS As Long = 12        'number of months in a year
Private Const NUMDAYS As Long = 7           'number of days in a week
Private Const BORDER3D As Long = 2          'num pixels for good 3d border
Private Const FOCUSBORDER As Long = 1       'num pixels for focus border

Private Enum DaySets
    PrevMonthDays
    CurMonthDays
    NextMonthDays
End Enum 'DaySets

Private Enum DayEffectFlags
    calEffectOff = 1
    calEffectOn = -1
    calEffectDefault = 0
End Enum 'DayEffectFlags

'======================================================================
' Private Data Members
'======================================================================
'Current Date
Private mnDay As Long               'current day number
Private mnYear As Long              'current year number
Private mnMonth As Long             'currnet month number

'Formatting and Appearance Settings
Private mnFirstDayOfWeek As VbDayOfWeek 'first day of the week
Private mnDayNameFormat As DayNameFormats
Private mfntDayNames As StdFont     'font to use for painting day names
Private mclrDayNames As OLE_COLOR   'color for the day names

Private mfShowIterrators As Boolean 'determines if iterrator buttons
                                    'should be shown or not
Private mfMonthReadOnly As Boolean  'month selector or none
Private mfYearReadOnly As Boolean  'month selector or none

'Behavior settings
Private mfLocked As Boolean         'read-only or not

'String Arrays For Month and Day Names
Private masMonthNames(NUMMONTHS - 1) As String 'string array of month names
Private masDayNames(NUMDAYS - 1) As String   'string array of day names

'this should be replaced with day styles eventually
Private mfntDayFont As StdFont      'font to use for painting dates in
                                    'the current month
Private mclrDay As OLE_COLOR        'color for the day numbers

Private mafDayBold(1 To 31) As DayEffectFlags   'array of flags for day being bold
Private mafDayItalic(1 To 31) As DayEffectFlags 'array of flags for day being italic

'Current Column Width and Row Height For Calendar Grid
Private mcxColWidth As Long         'width of each column in the grid
Private mcyRowHeight As Long        'height of each row in the grid

'RECTs For Different Calendar Areas
Private mrcNavArea As RECT          'rect bounding the navigation area
Private mrcDayNameArea As RECT      'rect bounding the day name area
Private mrcCalArea As RECT          'area bounding the calendar days
Private mrcFocusArea As RECT        'current focus area

'General Utility Members
Private mobjRes As ResLoader        'resource loading object (localization)
Private mfIgnoreMonthYearChange As Boolean  'HACKY flag for ignoring the programatic
                                            'change of the month and year navigation
                                            'controls.
Private maRepaintDays(1) As Long    'array of day numbers to repaint
Private mfFastRepaint As Boolean    'boolean flag used to do fast repaint
                                    'when only the day selected is changing

'======================================================================
' Public Property Procedures
'======================================================================

'----------------------------------------------------------------------
' Version Get
'----------------------------------------------------------------------
' Purpose:  Gets the version number of the control
'----------------------------------------------------------------------
Public Property Get Version() As String
Attribute Version.VB_Description = "Returns the version number of this control."
Attribute Version.VB_ProcData.VB_Invoke_Property = ";Misc"
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property 'Get Version()

'----------------------------------------------------------------------
' Day Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and lets the current day value
'----------------------------------------------------------------------
Public Property Get Day() As Long
Attribute Day.VB_Description = "Returns/Sets the Day number of the selected date."
Attribute Day.VB_ProcData.VB_Invoke_Property = ";Data"
    Day = mnDay
End Property 'Get Day()

Public Property Let Day(nNewVal As Long)
    'validate our inputs
    If nNewVal > 0 And nNewVal <= MaxDayInMonth(mnMonth, mnYear) Then
        ChangeValue nNewVal, mnMonth, mnYear
    Else
        mobjRes.RaiseUserError errPropValueRange, Array("Day", "0", CStr(MaxDayInMonth(mnMonth, mnYear)))
    End If
End Property 'Let Day()

'----------------------------------------------------------------------
' Month Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and lets the current month value
'----------------------------------------------------------------------
Public Property Get Month() As CalendarMonths
Attribute Month.VB_Description = "Returns/Sets the month number of the currently selected date."
Attribute Month.VB_ProcData.VB_Invoke_Property = ";Data"
    Month = mnMonth
End Property 'Get Month()

Public Property Let Month(nNewVal As CalendarMonths)
    'validate our inputs
    'note we still need to do this even though we're using
    'an enumeration since VB only treats this as a long value
    If nNewVal > 0 And nNewVal <= 12 Then
        ChangeValue mnDay, nNewVal, mnYear
    Else
        mobjRes.RaiseUserError errPropValueRange, Array("Month", "0", "12")
    End If
End Property 'Let Month()

'----------------------------------------------------------------------
' Year Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and lets the current year value
'----------------------------------------------------------------------
Public Property Get Year() As Long
Attribute Year.VB_Description = "Returns/Sets the year number of the currently selected date."
Attribute Year.VB_ProcData.VB_Invoke_Property = ";Data"
    Year = mnYear
End Property 'Get Year()

Public Property Let Year(nNewVal As Long)
    'validate our inputs
    'year must be between 100 and 9999 due to the restrictions
    'of the date data type in basic
    If nNewVal >= 100 And nNewVal <= 9999 Then
        ChangeValue mnDay, mnMonth, nNewVal
    Else
        mobjRes.RaiseUserError errPropValueRange, Array("Year", "100", "9999")
    End If
End Property 'Let Year()

'----------------------------------------------------------------------
' Value Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and lets the current date value
'----------------------------------------------------------------------
Public Property Get Value() As Date
Attribute Value.VB_Description = "Returns/Sets the currently selected date in the control."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_MemberFlags = "3c"
    Value = DateSerial(mnYear, mnMonth, mnDay)
End Property 'Get Value()

Public Property Let Value(dtNew As Date)
    ChangeValue VBA.Day(dtNew), VBA.Month(dtNew), VBA.Year(dtNew)
End Property 'Let Value()


'----------------------------------------------------------------------
' DayFont Get/Set
'----------------------------------------------------------------------
' Purpose:  Gets or sets the font to use for date numbers
'----------------------------------------------------------------------
Public Property Get DayFont() As Font
Attribute DayFont.VB_Description = "Returns/Sets the font used for the day numbers."
Attribute DayFont.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute DayFont.VB_UserMemId = -512
    Set DayFont = mfntDayFont
End Property 'Get DayFont()

'*** VB BUG Workaround ***
'The fntNew argument is passed in ByVal in order to
'get this property to show in the built-in Font
'property page.  When the bug is fixed, change this
'back to ByRef (remove ByVal)
Public Property Set DayFont(ByVal fntNew As Font)
    Set mfntDayFont = fntNew
    
    UserControl.Refresh
End Property 'Set DayFont()

'----------------------------------------------------------------------
' DayNameFont Get/Set
'----------------------------------------------------------------------
' Purpose:  Gets or sets the font to use for day names
'----------------------------------------------------------------------
Public Property Get DayNameFont() As Font
Attribute DayNameFont.VB_Description = "Returns/Sets the font used for the day names."
Attribute DayNameFont.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set DayNameFont = mfntDayNames
End Property 'Get DayFont()

'*** VB BUG Workaround ***
'The fntNew argument is passed in ByVal in order to
'get this property to show in the built-in Font
'property page.  When the bug is fixed, change this
'back to ByRef (remove ByVal)
Public Property Set DayNameFont(ByVal fntNew As Font)
    Set mfntDayNames = fntNew
    UserControl.Refresh
End Property 'Set DayFont()

'----------------------------------------------------------------------
' DayBold() Get/Let
'----------------------------------------------------------------------
' Purpose:  This property allows the user to set a particular day to bold
'           or not so as to give the effect of a 'special' day
' Inputs:   day number (1 to max day in current month)
' Outputs:  True if it's Bold, False if not
'----------------------------------------------------------------------
Public Property Get DayBold(DayNumber As Long) As Boolean
Attribute DayBold.VB_Description = "Returns/Sets the Bold state for a day in the current month."
    'if the setting for this day is "default" then the
    'value returned depends on the bold state of the
    'DayFont property
    If mafDayBold(DayNumber) = calEffectDefault Then
        DayBold = mfntDayFont.Bold
    Else
        DayBold = (mafDayBold(DayNumber) = calEffectOn)
    End If
End Property 'Get DayBold()

Public Property Let DayBold(DayNumber As Long, NewVal As Boolean)
    If NewVal = True Then
        mafDayBold(DayNumber) = calEffectOn
    Else
        mafDayBold(DayNumber) = calEffectOff
    End If
End Property 'Let DayBold()

'----------------------------------------------------------------------
' DayItalic() Get/Let
'----------------------------------------------------------------------
' Purpose:  This property allows the user to set a particular day italic
'           or not so as to give the effect of a 'special' day
' Inputs:   day number (1 to max day in current month)
' Outputs:  True if it's Italic, False if not
'----------------------------------------------------------------------
Public Property Get DayItalic(DayNumber As Long) As Boolean
Attribute DayItalic.VB_Description = "Returns/Sets the Italic state for a day in the current month."
    'if the setting for this day is "default" then the
    'value returned depends on the italic state of the
    'DayFont property
    If mafDayItalic(DayNumber) = calEffectDefault Then
        DayItalic = mfntDayFont.Italic
    Else
        DayItalic = (mafDayItalic(DayNumber) = calEffectOn)
    End If
End Property 'Get DayItalic()

'**Let
Public Property Let DayItalic(DayNumber As Long, NewVal As Boolean)
    If NewVal = True Then
        mafDayItalic(DayNumber) = calEffectOn
    Else
        mafDayItalic(DayNumber) = calEffectOff
    End If
End Property 'Let DayItalic()

'----------------------------------------------------------------------
' StartOfWeek Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets or lets the first day to display in a week
'----------------------------------------------------------------------
Public Property Get StartOfWeek() As DaysOfTheWeek
Attribute StartOfWeek.VB_Description = "Returns/Sets the first day of the week which will be displayed in the left-most column."
Attribute StartOfWeek.VB_ProcData.VB_Invoke_Property = ";Appearance"
    StartOfWeek = mnFirstDayOfWeek
End Property 'Get StartOfWeek()

Public Property Let StartOfWeek(nNewVal As DaysOfTheWeek)
    'validate our inputs
    If nNewVal >= calUseSystem And nNewVal <= calSaturday Then
        mnFirstDayOfWeek = nNewVal
        
        'do a Refresh to make the control repaint
        UserControl.Refresh
        
    Else
        mobjRes.RaiseUserError errPropValueRange, Array("StartOfWeek", calUseSystem, calSaturday)
    End If 'valid inputs
    
End Property 'Let StartOfWeek()

'----------------------------------------------------------------------
' DayNameFormat Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets or lets the format to use for day names
'           (short, medium, long)
'----------------------------------------------------------------------
Public Property Get DayNameFormat() As DayNameFormats
Attribute DayNameFormat.VB_Description = "Returns/Sets the format to use for the day names (Short = ""M"", Medium = ""Mon"", Long = ""Monday"")."
Attribute DayNameFormat.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayNameFormat = mnDayNameFormat
End Property 'Get DayNameFormat

Public Property Let DayNameFormat(nNewFormat As DayNameFormats)
    'validate the input
    If nNewFormat >= calShortName And nNewFormat <= calLongName Then
        'set the new format and re-load the day names
        mnDayNameFormat = nNewFormat
        LoadDayNames
        UserControl.Refresh
    Else
        mobjRes.RaiseUserError errPropValueRange, Array("DayNameFormat", calShortName, calLongName)
    End If 'valid inputs
End Property 'Let DayNameFormat

'----------------------------------------------------------------------
' ShowIterratorButtons Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets or lets the option for showing or hiding the month
'           iterrator buttons
'----------------------------------------------------------------------
Public Property Get ShowIterrationButtons() As Boolean
Attribute ShowIterrationButtons.VB_Description = "Returns/Sets the visible state of the previous and next month navigation buttons."
Attribute ShowIterrationButtons.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowIterrationButtons = mfShowIterrators
End Property 'Get ShowIterrationButtons()

Public Property Let ShowIterrationButtons(fNew As Boolean)
    'if it's not changing, don't bother
    If fNew = mfShowIterrators Then Exit Property
    
    'assign the new value
    mfShowIterrators = fNew
    
    'and adjust the visible state of the buttons
    btnPrev.Visible = mfShowIterrators
    btnNext.Visible = mfShowIterrators
    
    'trigger the resize event to recalc the widths
    'of the other navigation controls
    UserControl_Resize
End Property 'Let ShowIterrationButtons()

'----------------------------------------------------------------------
' MonthReadOnly Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and lets the option of making the month selector
'           read-only or not
'----------------------------------------------------------------------
Public Property Get MonthReadOnly() As Boolean
Attribute MonthReadOnly.VB_Description = "Returns/Sets the read-only state of the month navigation combo box."
Attribute MonthReadOnly.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MonthReadOnly = mfMonthReadOnly
End Property 'Get MonthReadOnly()

Public Property Let MonthReadOnly(fNew As Boolean)
    'if it's not changing, don't bother
    If fNew = mfMonthReadOnly Then Exit Property
    
    'set the new value and hide or show the month selector
    mfMonthReadOnly = fNew
    cbxMonth.Visible = Not mfMonthReadOnly
End Property 'Let MonthReadOnly()

'----------------------------------------------------------------------
' YearReadOnly Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and lets the option of making the year selector
'           read-only or not
'----------------------------------------------------------------------
Public Property Get YearReadOnly() As Boolean
Attribute YearReadOnly.VB_Description = "Returns/Sets the read-only state of the year navigation text box."
Attribute YearReadOnly.VB_ProcData.VB_Invoke_Property = ";Appearance"
    YearReadOnly = mfYearReadOnly
End Property 'Get YearReadOnly()

Public Property Let YearReadOnly(fNew As Boolean)
    'if it's not changing, don't bother
    If fNew = mfYearReadOnly Then Exit Property
    
    'set the new value and hide or show the month selector
    mfYearReadOnly = fNew
    txtYear.Visible = Not mfYearReadOnly
End Property 'Let YearReadOnly()

'----------------------------------------------------------------------
' Locked Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and sets the Locked option which makes the whole thing
'           read-only or not
'----------------------------------------------------------------------
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/Sets the locked state of the control.  When locked, the user cannot change the selected date."
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Locked = mfLocked
End Property 'Get Locked()

Public Property Let Locked(fNew As Boolean)
    
    'set the private variable
    mfLocked = fNew
    
    'set the locked state of contained controls
    'we'll disable the buttons if locked since
    'there is no locked state for buttons
    cbxMonth.Locked = fNew
    txtYear.Locked = fNew
    btnNext.Enabled = Not fNew
    btnPrev.Enabled = Not fNew
    
End Property 'Let Locked()

'----------------------------------------------------------------------
' DayColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and sets the color used for the day numbers
'----------------------------------------------------------------------
Public Property Get DayColor() As OLE_COLOR
Attribute DayColor.VB_Description = "Returns/Sets the color used for the day numbers."
Attribute DayColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute DayColor.VB_UserMemId = -513
    DayColor = mclrDay
End Property 'Get DayColor()

Public Property Let DayColor(NewVal As OLE_COLOR)
    mclrDay = NewVal
    UserControl.Refresh
End Property 'Let DayColor()

'----------------------------------------------------------------------
' DayNameColor Get/Let
'----------------------------------------------------------------------
' Purpose:  Gets and sets the color used for the day numbers
'----------------------------------------------------------------------
Public Property Get DayNameColor() As OLE_COLOR
Attribute DayNameColor.VB_Description = "Returns/Sets the color used for the day names (i.e. days of the week)."
Attribute DayNameColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayColor = mclrDayNames
End Property 'Get DayNameColor()

Public Property Let DayNameColor(NewVal As OLE_COLOR)
    mclrDayNames = NewVal
    UserControl.Refresh
End Property 'Let DayNameColor()



'======================================================================
' Public Methods
'======================================================================

'----------------------------------------------------------------------
' HitTest()
'----------------------------------------------------------------------
' Purpose:  Does a hit test based on x,y coordinates
' Inputs:   x and y coordinates
' Outputs:  Area of the control and specific date if over one
'----------------------------------------------------------------------
Public Sub HitTest(ByVal X As Long, ByVal Y As Long, Area As Long, HitDate As Date)
Attribute HitTest.VB_Description = "Returns the area and day number (if any) that corresponds to a given X,Y position."
    Dim nRow As Long
    Dim nCol As Long
    
    'assert that the x and y are indeed in our coordinate system
    Debug.Assert (X <= UserControl.ScaleWidth)
    Debug.Assert (Y <= UserControl.ScaleHeight)
    
    'determine the area of the control that x and y are over
    If X > mrcNavArea.Right Then
        Area = calUnknownArea
    Else
        If Y >= mrcNavArea.Top And Y <= mrcNavArea.Bottom Then
            Area = calNavigationArea
        ElseIf Y >= mrcDayNameArea.Top And Y <= mrcDayNameArea.Bottom Then
            Area = calDayNameArea
        ElseIf Y >= mrcCalArea.Top And Y <= mrcCalArea.Bottom Then
            Area = calDateArea
        Else
            Area = calUnknownArea
        End If 'determine area by y
    End If 'x is past right of all areas
    
    'if we are in the date area, calculate the hit date
    If Area = calDateArea Then
        
        'determine the row and column and make them 0-based
        nRow = ((Y - mrcCalArea.Top) \ mcyRowHeight) - 1
        If (Y - mrcCalArea.Top) Mod mcyRowHeight > 0 Then
            nRow = nRow + 1
        End If
        
        nCol = ((X - mrcCalArea.Left) \ mcxColWidth) - 1
        If (X - mrcCalArea.Left) Mod mcxColWidth > 0 Then
            nCol = nCol + 1
        End If
        
        'given the row and column, determine the date
        HitDate = DateForRowCol(nRow, nCol)
        
    End If 'in date area

End Sub 'HitTest

'----------------------------------------------------------------------
' Refresh()
'----------------------------------------------------------------------
' Purpose:  Refreshes/repaints the entire control
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the control by causing a complete repaint."
    'just pass it on...
    UserControl.Refresh
End Sub 'Refresh()

'----------------------------------------------------------------------
' About()
'----------------------------------------------------------------------
' Purpose:  Opens the About box for the control--this is marked hidden
'           so that it doesn't show up in the statement completion
'           but we do mark this with the DispID of AboutBox so that it
'           shows in the property sheet with an elipsis button
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Public Sub About()
Attribute About.VB_Description = "Shows the about box for the control."
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub 'About()


'======================================================================
' Initialize and Terminate Events
'======================================================================
Private Sub UserControl_Initialize()
    
    On Error GoTo Err_Init
    
    'set the resource loader
    'daveste -- 7/31/96
    'TODO: put in code to load a satellite resource DLL based on the
    'locale ID of the ambient host
    Set mobjRes = New ResLoader
    
    'load the month names into the combo box
    LoadMonthNames
    
    'initialize the area rects that don't depend on the
    'size of the control (which are left and top and sometimes bottom)
    'doing this here lets us reduce the code needed to execute
    'when the control is resized which will happen more often
    'than the control being initialized.
    mrcNavArea.Left = 1
    mrcNavArea.Top = 1
    
    'height of navigation area is the height of the month combo
    'plus 4, since we will draw a 3d box around the controls
    mrcNavArea.Bottom = cbxMonth.Height + (2 * BORDER3D)
    mrcDayNameArea.Left = 1
    mrcDayNameArea.Top = mrcNavArea.Bottom
    
    'height of the day name area should be the height of
    'the day name font plus 6 pixels for 3d effects
    mrcDayNameArea.Bottom = mrcDayNameArea.Top + UserControl.TextHeight("A") + 6
    
    mrcCalArea.Left = 1
    mrcCalArea.Top = mrcDayNameArea.Bottom
    
    'set the position and sizes of the navigation controls that
    'don't depend on the size of the control (like left and top
    'values).
    btnPrev.Move mrcNavArea.Left, mrcNavArea.Top, btnPrev.Width, mrcNavArea.Bottom - mrcNavArea.Top
    
    btnNext.Top = mrcNavArea.Top
    btnNext.Height = mrcNavArea.Bottom - mrcNavArea.Top
    
    cbxMonth.Move mrcNavArea.Left + btnPrev.Width + BORDER3D, mrcNavArea.Top + BORDER3D
    txtYear.Height = cbxMonth.Height
    txtYear.Top = mrcNavArea.Top + BORDER3D
    
    'set the disabled picture for the prev and next buttons
    'to be the same as the regular picture--this will let us
    'give a locked effect by disabling the prev and next buttons
    btnPrev.DisabledPicture = btnPrev.Picture
    btnNext.DisabledPicture = btnNext.Picture
    
    Exit Sub

Err_Init:
    Debug.Assert False
    Exit Sub
End Sub 'UserControl_Initialize()

'======================================================================
' Private Event Handles
'======================================================================

'----------------------------------------------------------------------
' InitProperties Event
'----------------------------------------------------------------------
' Purpose:  Called when the control is first put on a form
'           One-time initialization of data members
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    Dim dt As Date
        
    On Error GoTo Err_InitProps
    
    'initialize the day, month and year to the current system date
    dt = Date
    mnDay = VBA.Day(dt)
    mnMonth = VBA.Month(dt)
    mnYear = VBA.Year(dt)
    
    mfIgnoreMonthYearChange = True
    cbxMonth.ListIndex = mnMonth - 1
    txtYear.Text = mnYear
    mfIgnoreMonthYearChange = False
    
    'create new font objects for the day and day name
    'fonts and copy the font attributes from the
    'user control's ambient font into them
    Set mfntDayFont = New StdFont
    CopyFont UserControl.Ambient.Font, mfntDayFont
    
    Set mfntDayNames = New StdFont
    CopyFont UserControl.Ambient.Font, mfntDayNames
    mfntDayNames.Bold = True
    
    'initialize the day and dayname colors to the ambient's
    'fore color value
    mclrDay = vbBlack
    mclrDayNames = vbBlack
    
    'initialize the day name format to medium
    mnDayNameFormat = calMediumName
    LoadDayNames
    
    'init various appearance options
    mfShowIterrators = True
    mfMonthReadOnly = False
    mfYearReadOnly = False
    mfLocked = False
    
    Exit Sub

Err_InitProps:
    Debug.Assert False
    Exit Sub
End Sub 'UserControl_InitProperties()

'----------------------------------------------------------------------
' ReadProperties Event
'----------------------------------------------------------------------
' Purpose:  Called when we need to read property settings back in
' Inputs:   the property bag class for reading
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim dtCurrent As Date
    dtCurrent = Date
    
    On Error Resume Next
    'read in the properties from the property bag
    mnFirstDayOfWeek = PropBag.ReadProperty("StartOfWeek", vbUseSystemDayOfWeek)
    
    ChangeValue PropBag.ReadProperty("Day", VBA.Day(dtCurrent)), _
                PropBag.ReadProperty("Month", VBA.Month(dtCurrent)), _
                PropBag.ReadProperty("Year", VBA.Year(dtCurrent))
    
    Set mfntDayNames = PropBag.ReadProperty("DayNameFont", UserControl.Font)
    Set mfntDayFont = PropBag.ReadProperty("DayFont", UserControl.Font)
    
    mclrDay = PropBag.ReadProperty("DayColor", vbBlack)
    mclrDayNames = PropBag.ReadProperty("DayNameColor", vbBlack)
    
    mnDayNameFormat = PropBag.ReadProperty("DayNameFormat", calMediumName)
    LoadDayNames
    
    Me.ShowIterrationButtons = PropBag.ReadProperty("ShowIterrationButtons", True)
    Me.MonthReadOnly = PropBag.ReadProperty("MonthReadOnly", False)
    Me.YearReadOnly = PropBag.ReadProperty("YearReadOnly", False)
    Me.Locked = PropBag.ReadProperty("Locked", False)
    
    'trigger a resize since this event happens after the initial
    'resize when going to run mode
    UserControl_Resize
    
End Sub 'UserControl_ReadProperties()

'----------------------------------------------------------------------
' WriteProperties Event
'----------------------------------------------------------------------
' Purpose:  Called when we need to write property settings out to disk
' Inputs:   the property bag class for writing
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    'write the current property values to the property bag
    PropBag.WriteProperty "Day", mnDay
    PropBag.WriteProperty "Month", mnMonth
    PropBag.WriteProperty "Year", mnYear
    
    PropBag.WriteProperty "StartOfWeek", mnFirstDayOfWeek, vbUseSystemDayOfWeek
    PropBag.WriteProperty "DayNameFont", mfntDayNames, UserControl.Font
    PropBag.WriteProperty "DayFont", mfntDayFont, UserControl.Font
    PropBag.WriteProperty "DayNameFormat", mnDayNameFormat, calMediumName
    PropBag.WriteProperty "DayColor", mclrDay, vbBlack
    PropBag.WriteProperty "DayNameColor", mclrDayNames, vbBlack
    
    
    PropBag.WriteProperty "ShowIterrationButtons", mfShowIterrators, True
    PropBag.WriteProperty "MonthReadOnly", mfMonthReadOnly, False
    PropBag.WriteProperty "YearReadOnly", mfYearReadOnly, False
    PropBag.WriteProperty "Locked", mfLocked, False
    
End Sub 'UserControl_WriteProperties()

'----------------------------------------------------------------------
' Paint Event
'----------------------------------------------------------------------
' Purpose:  Called when the control needs to be repainted
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_Paint()
    Dim dcWork As OffScreenDC
    
    Dim nTop As Long
    Dim nLeft As Long
    Dim nWidth As Long
    Dim nHeight As Long
    
    Dim nDay As Long
    Dim nRow As Long
    Dim nCol As Long
    Dim nLastDay As Long
    Dim eDaySet As DaySets
    Dim rgbColor As Long
    Dim fDefBold As Boolean
    Dim fDefItalic As Boolean
    
    On Error GoTo Err_Paint
    
    'save the initial bold and italic state of our day font
    fDefBold = mfntDayFont.Bold
    fDefItalic = mfntDayFont.Italic
    
    Set dcWork = New OffScreenDC
    
    dcWork.Initialize UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    'set the text color to be the color chosen for
    'the days of the week names
    OleTranslateColor mclrDayNames, 0, rgbColor
    dcWork.TextColor = rgbColor
    
    If mfFastRepaint Then
        FastRepaint dcWork
        Exit Sub
    End If
    
    'fill the background of the control with the ambient's
    'background color
    nLeft = 0
    nTop = 0
    nWidth = UserControl.ScaleWidth
    nHeight = UserControl.ScaleHeight
    
    'I use the OLE API OleTranslateColor here to translate
    'an OLE color to an RGB value.  VB will return an OLE color
    'value for the ambient's back color and this API will convert
    'it to an RGB value for painting.
    OleTranslateColor UserControl.Ambient.BackColor, 0, rgbColor
    
    dcWork.FillRect nLeft, nTop, nWidth, nHeight, rgbColor
    
    'next fill a black rect that will serve as a thin back outline
    'around the painted part of the control
    nWidth = mrcNavArea.Right + 1
    nHeight = mrcDayNameArea.Bottom + (mcyRowHeight * NUMROWS) + 1
    dcWork.FillRect 0, 0, nWidth, nHeight, vbBlack
    
    'draw a 3d rect around the navigation controls
    nTop = mrcNavArea.Top
    nHeight = mrcNavArea.Bottom - mrcNavArea.Top
    
    If mfShowIterrators Then
        nLeft = mrcNavArea.Left + btnPrev.Width
        nWidth = btnNext.Left - nLeft
    Else
        nLeft = mrcNavArea.Left
        nWidth = mrcNavArea.Right - mrcNavArea.Left
    End If 'mfShowIterrators
    
    dcWork.Draw3DRect nLeft, nTop, nWidth, nHeight
    
    'if the month is read only, draw the month name
    If mfMonthReadOnly Then
        Set dcWork.Font = cbxMonth.Font
        
        'squeeze the width in by one to make a better 3d effect
        dcWork.Draw3DRect cbxMonth.Left, cbxMonth.Top, _
                            cbxMonth.Width - 1, cbxMonth.Height, _
                            cbxMonth.List(cbxMonth.ListIndex), _
                            caCenterCenter, Sunken
    End If 'month is read only
    
    'if the year is read only, draw the year number
    If mfYearReadOnly Then
        Set dcWork.Font = txtYear.Font
        
        dcWork.Draw3DRect txtYear.Left, txtYear.Top, _
                            txtYear.Width, txtYear.Height, _
                            txtYear.Text, caCenterCenter, Sunken
    End If 'year is read only
    
    'paint the day names
    PaintDayNames dcWork
    
    'change the text color to dark gray to paint the previous month days
    'daveste -- 7/31/96
    'TODO: this should be replaced with day styles or at least with
    'a property the control the font and color of these other dates
    dcWork.TextColor = RGB(128, 128, 128)
    
    'get the first and last days of the previous month to paint
    GetPrevMonthDays mnMonth, mnYear, nDay, nLastDay
    eDaySet = PrevMonthDays
    
    Set dcWork.Font = mfntDayFont
    
    'draw a grid of date numbers for the current month
    For nRow = 0 To NUMROWS - 1
        For nCol = 0 To NUMCOLS - 1
            
            'if we've done painting the current set of days
            'switch to the next set
            If nDay > nLastDay Then
                If eDaySet = PrevMonthDays Then
                    OleTranslateColor mclrDay, 0, rgbColor
                    dcWork.TextColor = rgbColor
                    nDay = 1
                    nLastDay = MaxDayInMonth(mnMonth, mnYear)
                    eDaySet = CurMonthDays
                    
                Else
                
                    dcWork.TextColor = RGB(128, 128, 128)
                    nDay = 1
                    nLastDay = 100 'no need to calc the last
                                    'day since the for loops
                                    'will govern when to stop
                    eDaySet = NextMonthDays
                    
                End If 'day set was previous month
            End If 'done painting this day set
            
            'paint the day
            
            'set the font attributes for the day being painted
            If eDaySet = CurMonthDays Then
                If mafDayBold(nDay) = calEffectDefault Then
                    'optimize for the case where no days are bold
                    If mfntDayFont.Bold <> fDefBold Then
                        mfntDayFont.Bold = fDefBold
                        Set dcWork.Font = mfntDayFont
                    End If
                Else
                    mfntDayFont.Bold = (mafDayBold(nDay) = calEffectOn)
                    Set dcWork.Font = mfntDayFont
                End If 'DayBold setting is default
                
                If mafDayItalic(nDay) = calEffectDefault Then
                    'optimize for the case where no days are italic
                    If mfntDayFont.Italic <> fDefItalic Then
                        mfntDayFont.Italic = fDefItalic
                        Set dcWork.Font = mfntDayFont
                    End If
                Else
                    mfntDayFont.Italic = (mafDayItalic(nDay) = calEffectOn)
                    Set dcWork.Font = mfntDayFont
                End If
            End If 'we're in the current month day set
            
            'if it's the current day, draw it selected
            If nDay = mnDay And eDaySet = CurMonthDays Then
                dcWork.Draw3DRect mrcCalArea.Left + (nCol * mcxColWidth), _
                                    mrcCalArea.Top + (nRow * mcyRowHeight), _
                                    mcxColWidth, mcyRowHeight, CStr(nDay), _
                                    caCenterCenter, Selected
                                    
            Else
            
                dcWork.Draw3DRect mrcCalArea.Left + (nCol * mcxColWidth), _
                                    mrcCalArea.Top + (nRow * mcyRowHeight), _
                                    mcxColWidth, mcyRowHeight, CStr(nDay)
            
            End If 'current day
            
            'increment the day number
            nDay = nDay + 1
            
        Next nCol
    Next nRow
    
    'blast the control to the screen
    dcWork.BlastToScreen
    
    'if the dummy control has focus, and we are in run-mode,
    'draw a focus rect around the current focus area
    If UserControl.ActiveControl Is ctlFocus Then
        DrawFocusRect UserControl.hdc, mrcFocusArea
    End If
    
    'restore the initial settings for bold and italic
    'in our day font
    mfntDayFont.Bold = fDefBold
    mfntDayFont.Italic = fDefItalic
    
    Exit Sub
    
Err_Paint:
    Debug.Assert False
    Exit Sub
End Sub 'UserControl_Paint()

'----------------------------------------------------------------------
' Resize Event
'----------------------------------------------------------------------
' Purpose:  Called when the control is resized by the developer
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_Resize()
    Dim nNewWidth As Long       'new scale width
    Dim nNewHeight As Long      'new scale height
    Dim nUsableWidth As Long    'actual width we can use
    
    On Error GoTo Err_Resize
    
    nNewWidth = UserControl.ScaleWidth
    nNewHeight = UserControl.ScaleHeight
    
    'since all the grid cells need to be the same width
    'the usable width is the width we will consume and there
    'maybe unused pixels due to left-overs from division
    nUsableWidth = ((nNewWidth - (2 * mrcCalArea.Left)) \ NUMCOLS) * NUMCOLS
    
    'recalculate the bounding rectangles for the various areas
    'of the control (navigation, day names, and calendar days)
    mrcNavArea.Right = mrcNavArea.Left + nUsableWidth
    mrcDayNameArea.Right = mrcDayNameArea.Left + nUsableWidth
    mrcCalArea.Right = mrcCalArea.Left + nUsableWidth
    mrcCalArea.Bottom = nNewHeight
    
    'Recalculate the width and heights of the grid rows and columns
    mcxColWidth = (nNewWidth - (2 * mrcCalArea.Left)) \ NUMCOLS
    mcyRowHeight = (mrcCalArea.Bottom - mrcCalArea.Top) \ NUMROWS
    
    'resize the month and year selection controls
    btnNext.Left = mrcNavArea.Right - btnNext.Width
    
    'if there's not enough room, just display the buttons
    If (mrcNavArea.Right - mrcNavArea.Left) <= _
        (btnNext.Width + btnPrev.Width + txtYear.Width + 10) _
        And mfShowIterrators Then
        
        cbxMonth.Visible = False
        txtYear.Visible = False
        
    Else
    
        If Not mfMonthReadOnly Then cbxMonth.Visible = True
        If Not mfYearReadOnly Then txtYear.Visible = True
        
        If mfShowIterrators Then
            cbxMonth.Left = mrcNavArea.Left + btnPrev.Width + BORDER3D
            txtYear.Left = btnNext.Left - txtYear.Width - BORDER3D
        Else
            cbxMonth.Left = mrcNavArea.Left + BORDER3D
            txtYear.Left = mrcNavArea.Right - txtYear.Width - BORDER3D
        End If
        
        cbxMonth.Width = txtYear.Left - cbxMonth.Left
    
    End If 'not enough horizontal room
    
    Exit Sub
    
Err_Resize:
    Debug.Assert False
    Exit Sub
    
End Sub 'UserControl_Resize()

'----------------------------------------------------------------------
' MouseDown Event
'----------------------------------------------------------------------
' Purpose:  Called when the mouse button is pushed down while over
'           the control's area
' Inputs:   Which mouse button, shift state and x and y position
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Area As CalendarAreas
    Dim dtOld As Date
    Dim dtNew As Date
        
    On Error GoTo Err_MouseDown
    
    'keep the old date to see if it's changed
    dtOld = Me.Value
    
    'Do a hit test to determine where the user clicked
    Me.HitTest X, Y, Area, dtNew
    
    'if the area was in the date area and the control is not locked,
    'switch to the hit date
    If (Area = calDateArea) And (Not mfLocked) Then
        If dtNew <> dtOld Then
            ChangeValue VBA.Day(dtNew), VBA.Month(dtNew), VBA.Year(dtNew)
        End If 'date did change
    End If 'clicked in date area
    
    'grab focus back if needed
    If Not (UserControl.ActiveControl Is ctlFocus) Then
        ctlFocus.SetFocus
    End If
    
    Exit Sub

Err_MouseDown:
    Debug.Assert False
    Exit Sub
End Sub 'UserControl_MouseDown()

'----------------------------------------------------------------------
' DblClick Event
'----------------------------------------------------------------------
' Purpose:  Called when the user double-clicks on the main control area
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_DblClick()
    On Error GoTo Err_DblClick
    
    'pass this event to the host
    RaiseEvent DblClick
    Exit Sub

Err_DblClick:
    Exit Sub
End Sub 'UserControl_DblClick()

'----------------------------------------------------------------------
' Click Event
'----------------------------------------------------------------------
' Purpose:  Called when the user clicks on the main control area
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub UserControl_Click()
    On Error GoTo Err_Click
    
    'raise our click event to the user
    RaiseEvent Click

    Exit Sub
    
Err_Click:
    Exit Sub
End Sub 'UserControl_Click()

'----------------------------------------------------------------------
' ctlFocus_GotFocus Event
'----------------------------------------------------------------------
' Purpose:  Called when the main calendar area is to get focus.
'           We use a dummy control to capture focus since we are
'           just painting the calendar days and cannot set focus
'           to the entire user control.
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub ctlFocus_GotFocus()
    'draw a focus rect to signify that the calendar
    'area now has focus
    DrawFocusRect UserControl.hdc, mrcFocusArea
End Sub 'ctlFocus_GotFocus()

'----------------------------------------------------------------------
' ctlFocus_LostFocus Event
'----------------------------------------------------------------------
' Purpose:  Called when the main calendar area has lost focus.
'           We use a dummy control to capture focus since we are
'           just painting the calendar days and cannot set focus
'           to the entire user control.
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub ctlFocus_LostFocus()
    'draw a focus rect where the last focus area was
    'drawing a focus rect twice removes it
    DrawFocusRect UserControl.hdc, mrcFocusArea
End Sub 'ctlFocus_LostFocus()

'----------------------------------------------------------------------
' ctlFocus_KeyDown Event
'----------------------------------------------------------------------
' Purpose:  Called when the user presses a key while the dummy control
'           has focus
' Inputs:   Which key, shift state
' Outputs:  None
'----------------------------------------------------------------------
Private Sub ctlFocus_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim dtTemp As Date      'temp date for date arithmetic
    
    Select Case KeyCode
        Case vbKeyLeft
            dtTemp = DateSerial(mnYear, mnMonth, mnDay)
            
            'if shift is down, move by month
            If (Shift And vbShiftMask) > 0 Then
                dtTemp = DateAdd("m", -1, dtTemp)
            
            ElseIf (Shift And vbCtrlMask) > 0 Then
                'else if control is down, move by year
                dtTemp = DateAdd("yyyy", -1, dtTemp)
            
            Else
                'go back on day
                dtTemp = DateAdd("d", -1, dtTemp)
            End If
            
            ChangeValue VBA.Day(dtTemp), VBA.Month(dtTemp), _
                        VBA.Year(dtTemp)
        
        Case vbKeyRight
            dtTemp = DateSerial(mnYear, mnMonth, mnDay)
            
            If (Shift And vbShiftMask) > 0 Then
                dtTemp = DateAdd("m", 1, dtTemp)
            
            ElseIf (Shift And vbCtrlMask) > 0 Then
                'else if control is down, move by year
                dtTemp = DateAdd("yyyy", 1, dtTemp)
            
            Else
                'go forward one day
                dtTemp = DateAdd("d", 1, dtTemp)
            End If
            
            ChangeValue VBA.Day(dtTemp), VBA.Month(dtTemp), _
                        VBA.Year(dtTemp)
            
        Case vbKeyUp
            'go one week back
            dtTemp = DateSerial(mnYear, mnMonth, mnDay)
            dtTemp = DateAdd("ww", -1, dtTemp)
            ChangeValue VBA.Day(dtTemp), VBA.Month(dtTemp), _
                        VBA.Year(dtTemp)
            
        Case vbKeyDown
            'go one week forwad
            dtTemp = DateSerial(mnYear, mnMonth, mnDay)
            dtTemp = DateAdd("ww", 1, dtTemp)
            ChangeValue VBA.Day(dtTemp), VBA.Month(dtTemp), _
                        VBA.Year(dtTemp)
            
        Case vbKeyHome
            'if control is down, go to first day of the year
            If (Shift And vbCtrlMask) > 0 Then
                ChangeValue 1, 1, mnYear
            Else
                'go to the first day of the current month
                ChangeValue 1, mnMonth, mnYear
            End If
            
        Case vbKeyEnd
            'if control is down, go to last day of the year
            If (Shift And vbCtrlMask) > 0 Then
                ChangeValue 31, 12, mnYear
            Else
                'go to the last day of the current month
                ChangeValue MaxDayInMonth(mnMonth, mnYear), _
                            mnMonth, mnYear
            End If
            
    End Select
End Sub 'ctlFocus_KeyDown()

'----------------------------------------------------------------------
' cbxMonth_Click Event
'----------------------------------------------------------------------
' Purpose:  Called when the user changes the item selected in the moth
'           navigation combo box
' Inputs:   none
' Outputs:  None
'----------------------------------------------------------------------
Private Sub cbxMonth_Click()
    If mfIgnoreMonthYearChange Then Exit Sub
    
    'if we are locked, just reset the list index
    'to the current month
    If mfLocked Then
        mfIgnoreMonthYearChange = True
        cbxMonth.ListIndex = mnMonth - 1
        mfIgnoreMonthYearChange = False
    End If
    
    'change the date
    ChangeValue mnDay, cbxMonth.ListIndex + 1, mnYear
    
    RaiseEvent Click
End Sub 'cbxMonth_Click()

'----------------------------------------------------------------------
' txtYear_KeyPress Event
'----------------------------------------------------------------------
' Purpose:  Called when the user presses a key in the year
'           navigation text box
' Inputs:   Key Pressed
' Outputs:  None
'----------------------------------------------------------------------
Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If mfIgnoreMonthYearChange Then Exit Sub
    
    'if they pressed return, process the date change
    If KeyAscii = vbKeyReturn Then
        'change the date
        ChangeValue mnDay, mnMonth, Val(txtYear)
        KeyAscii = 0
    End If
    
End Sub 'txtYear_KeyPress

'----------------------------------------------------------------------
' txtYear_Click Event
'----------------------------------------------------------------------
' Purpose:  Called when the user clicks the year
'           navigation text box
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub txtYear_Click()
    RaiseEvent Click
End Sub 'txtYear_Click()

'----------------------------------------------------------------------
' txtYear_GotFocus Event
'----------------------------------------------------------------------
' Purpose:  Called when the user moved into the year text box
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub txtYear_GotFocus()
    'select all the text that is there
    txtYear.SelStart = 0
    txtYear.SelLength = Len(txtYear.Text)
End Sub

'----------------------------------------------------------------------
' txtYear_LostFocus Event
'----------------------------------------------------------------------
' Purpose:  Called when the user moved out of the year text box
' Inputs:   None
' Outputs:  None
'----------------------------------------------------------------------
Private Sub txtYear_LostFocus()
    If mnYear <> Val(txtYear.Text) Then
        ChangeValue mnDay, mnMonth, Val(txtYear.Text)
    End If
End Sub 'txtYear_LostFocus()


'----------------------------------------------------------------------
' btnNext_Click Event
'----------------------------------------------------------------------
' Purpose:  Called when the user clicks the next month button
' Inputs:   none
' Outputs:  None
'----------------------------------------------------------------------
Private Sub btnNext_Click()
    Dim dtTemp As Date
    dtTemp = DateSerial(mnYear, mnMonth, mnDay)
    dtTemp = DateAdd("m", 1, dtTemp)
    ChangeValue VBA.Day(dtTemp), VBA.Month(dtTemp), VBA.Year(dtTemp)
    ctlFocus.SetFocus
    RaiseEvent Click
End Sub 'btnNext_Click()

'----------------------------------------------------------------------
' btnPrev_Click Event
'----------------------------------------------------------------------
' Purpose:  Called when the user clicks the previous month button
' Inputs:   none
' Outputs:  None
'----------------------------------------------------------------------
Private Sub btnPrev_Click()
    Dim dtTemp As Date
    dtTemp = DateSerial(mnYear, mnMonth, mnDay)
    dtTemp = DateAdd("m", -1, dtTemp)
    ChangeValue VBA.Day(dtTemp), VBA.Month(dtTemp), VBA.Year(dtTemp)
    ctlFocus.SetFocus
    RaiseEvent Click
End Sub 'btnPrev_Click()


'======================================================================
' Private Helper Methods
'======================================================================

'----------------------------------------------------------------------
' PaintDayNames()
'----------------------------------------------------------------------
' Purpose:  Paints names of the week days above the main date grid
' Inputs:   reference to the offscreen dc object
' Outputs:  none
'----------------------------------------------------------------------
Private Sub PaintDayNames(dc As OffScreenDC)
    Dim rc As RECT
    Dim nCol As Long
    Dim fntOld As StdFont
    Dim idx As Long
    
    'make a copy of the day name area rect
    rc.Left = mrcDayNameArea.Left
    rc.Top = mrcDayNameArea.Top
    rc.Right = mrcDayNameArea.Right
    rc.Bottom = mrcDayNameArea.Bottom
    
    'set the current font to use
    Set fntOld = dc.Font
    Set dc.Font = mfntDayNames
    
    'fill a black rect as a border
    dc.FillRect rc.Left, rc.Top, rc.Right - rc.Left, _
                rc.Bottom - rc.Top, vbBlack
                
    'now draw 3d rects for each day name
    rc.Top = rc.Top + 1
    rc.Bottom = rc.Bottom - 1
    
    'initialize idx to be the setting for first day of week
    'and if that setting is "use system", determine what the
    'system is using
    If mnFirstDayOfWeek = vbUseSystemDayOfWeek Then
        '8/4/96 is a Sunday, so if the system says the day
        'of week is other than 1, we'll figure that out
        idx = WeekDay(DateSerial(1996, 8, 4), mnFirstDayOfWeek)
    Else
        idx = mnFirstDayOfWeek
    End If 'first day of week was "use system"
    
    For nCol = 0 To NUMCOLS - 1
        dc.Draw3DRect (nCol * mcxColWidth) + rc.Left, rc.Top, mcxColWidth, _
                        rc.Bottom - rc.Top, masDayNames(idx - 1)
        
        'increment the indexer and if it's past the end
        'wrap it back around to zero
        idx = idx + 1
        If idx > NUMCOLS Then idx = 1
    Next nCol
    
    'reset the old font
    Set dc.Font = fntOld
End Sub 'PaintDayNames()

'----------------------------------------------------------------------
' FastRepaint()
'----------------------------------------------------------------------
' Purpose:  Fast repaint routine for painting when only the day number
'           changes and not the month or year.
' Inputs:   work off screen DC
' Outputs:  none
'----------------------------------------------------------------------
Private Sub FastRepaint(dcWork As OffScreenDC)
    Dim nLeft As Long
    Dim nTop As Long
    Dim rgbColor As Long
    Dim ct As Long
    Dim eAppearance As Appearances
    Dim fDefBold As Boolean
    Dim fDefItalic As Boolean
    
    'save the initial states of bold and italic in our day font
    fDefBold = mfntDayFont.Bold
    fDefItalic = mfntDayFont.Italic
    
    'set the font as the day font and the text
    'color as black
    Set dcWork.Font = mfntDayFont
    OleTranslateColor mclrDay, 0, rgbColor
    dcWork.TextColor = rgbColor
    
    For ct = 0 To 1
        If mafDayBold(maRepaintDays(ct)) = calEffectDefault Then
            'optimize for the case where no days are bold
            If mfntDayFont.Bold <> fDefBold Then
                mfntDayFont.Bold = fDefBold
                Set dcWork.Font = mfntDayFont
            End If
        Else
            mfntDayFont.Bold = (mafDayBold(maRepaintDays(ct)) = calEffectOn)
            Set dcWork.Font = mfntDayFont
        End If 'DayBold setting is default
        
        If mafDayItalic(maRepaintDays(ct)) = calEffectDefault Then
            'optimize for the case where no days are italic
            If mfntDayFont.Italic <> fDefItalic Then
                mfntDayFont.Italic = fDefItalic
                Set dcWork.Font = mfntDayFont
            End If
        Else
            mfntDayFont.Italic = (mafDayItalic(maRepaintDays(ct)) = calEffectOn)
            Set dcWork.Font = mfntDayFont
        End If
        
        'repaint the old day as normal
        nLeft = LeftForDay(maRepaintDays(ct))
        nTop = TopForDay(maRepaintDays(ct))
        
        If ct = 0 Then
            eAppearance = Raised
        Else
            eAppearance = Selected
        End If
        
        dcWork.Draw3DRect nLeft, nTop, _
                            mcxColWidth, mcyRowHeight, _
                            CStr(maRepaintDays(ct)), _
                            caCenterCenter, eAppearance
        
        'blast just this day to the screen
        dcWork.BlastToScreen nLeft, nTop, mcxColWidth, mcyRowHeight
    
    Next ct
    
'    'repaint the newly selected day as selected
'    nLeft = LeftForDay(maRepaintDays(1))
'    nTop = TopForDay(maRepaintDays(1))
'    dcWork.Draw3DRect nLeft, nTop, _
'                        mcxColWidth, mcyRowHeight, _
'                        CStr(maRepaintDays(1)), _
'                        caCenterCenter, Selected
'
'    'blast just this day to the screen
'    dcWork.BlastToScreen nLeft, nTop, mcxColWidth, mcyRowHeight
    
    'draw the focus rect on the selected day if
    'the dummy focus control has focus
    If UserControl.ActiveControl Is ctlFocus Then
        DrawFocusRect UserControl.hdc, mrcFocusArea
    End If
    
    'restore the initial states of bold and italic in our day font
    mfntDayFont.Bold = fDefBold
    mfntDayFont.Italic = fDefItalic
    
    'reset the fast repaint flag to False
    mfFastRepaint = False
    
End Sub 'FastRepaint()

'----------------------------------------------------------------------
' MaxDayInMonth()
'----------------------------------------------------------------------
' Purpose:  Returns the max day number for a given month number and year
' Inputs:   month number
' Outputs:  max day number
'----------------------------------------------------------------------
Private Function MaxDayInMonth(nMonth As Long, Optional nYear As Long = 0) As Long
    Select Case nMonth
        Case 9, 4, 6, 11    '30 days hath September,
                            'April, June, and November
            MaxDayInMonth = 30
        
        Case 2              'February -- check for leapyear
            'The full rule for leap years is that they occur in
            'every year divisible by four, except that they don't
            'occur in years divisible by 100, except that they
            '*do* in years divisible by 400.
            If (nYear Mod 4) = 0 Then
                If nYear Mod 100 = 0 Then
                    If nYear Mod 400 = 0 Then
                        MaxDayInMonth = 29
                    Else
                        MaxDayInMonth = 28
                    End If 'divisible by 400
                Else
                    MaxDayInMonth = 29
                End If 'divisible by 100
            Else
                MaxDayInMonth = 28
            End If 'divisible by 4
        
        Case Else           'All the rest have 31
            MaxDayInMonth = 31
    
    End Select
End Function 'MaxDayInMonth()

'----------------------------------------------------------------------
' ChangeValue()
'----------------------------------------------------------------------
' Purpose:  Changes the control's current value, checking if it's OK
'           and doing the necessary notifications that it's changed
' Inputs:   new day, month and year
' Outputs:  none
'----------------------------------------------------------------------
Private Sub ChangeValue(nDay As Long, nMonth As Long, nYear As Long)
    Dim rc As RECT          'used to invalidate smaller rects
                            'if only the day number changed
    
    Dim fCancel As Boolean  'used in the WillChangeDate event
    Dim dtOld As Date       'old date for raising in event
    
    'give the developer a chance to cancel the date change
    fCancel = False
    RaiseEvent WillChangeDate(DateSerial(nYear, nMonth, nDay), fCancel)
    If fCancel Then Exit Sub
    
    'build a date using the current values
    dtOld = DateSerial(mnYear, mnMonth, mnDay)
    
    'check to see if it's OK to change the value
    If UserControl.CanPropertyChange("Value") Then
        
        'changing the month or year can make the day number
        'invalid, so check the new combination and adjust the day
        'if necessary.
        If nDay > MaxDayInMonth(nMonth, nYear) Then
            nDay = MaxDayInMonth(nMonth, nYear)
        End If
        
        'to avoid unecessary repainting, if only the day number changed
        'just invalidate the two rects where the old and new dates are
        If mnMonth = nMonth And mnYear = nYear Then
            
            'setup a rect for the old day
            rc.Left = LeftForDay(mnDay)
            rc.Top = TopForDay(mnDay)
            rc.Right = rc.Left + mcxColWidth
            rc.Bottom = rc.Top + mcyRowHeight
            
            'invalidate it
            InvalidateRect UserControl.hwnd, rc, 0
            
            'setup a rect for the new day
            rc.Left = LeftForDay(nDay)
            rc.Top = TopForDay(nDay)
            rc.Right = rc.Left + mcxColWidth
            rc.Bottom = rc.Top + mcyRowHeight
            
            'invalidate it
            InvalidateRect UserControl.hwnd, rc, 0
            
            'since we are only changing the current day
            'and not the current month or year, store off
            'the specific days to repaint and set the
            'fast repaint flag to true.  This will cause the
            'paint routing to just repaint these two days
            'which makes the repaint considerably faster.
            'The fast repaint is reset to False automatically.
            maRepaintDays(0) = mnDay
            maRepaintDays(1) = nDay
            mfFastRepaint = True
            
            'change the value and notify those interested
            mnDay = nDay
            
        Else
            'reset the month and year navigators if they need to be
            mfIgnoreMonthYearChange = True
            If cbxMonth.ListIndex <> (nMonth - 1) Then cbxMonth.ListIndex = (nMonth - 1)
            If Val(txtYear.Text) <> nYear Then txtYear.Text = nYear
            mfIgnoreMonthYearChange = False
            
            'change the value and notify those interested
            mnDay = nDay
            mnMonth = nMonth
            mnYear = nYear

            'refresh the entire calendar area since we have to
            're-layout the days
            InvalidateRect UserControl.hwnd, mrcCalArea, 0
        End If 'just changing the day
        
        'update the new focus area based on the new day selected
        mrcFocusArea.Left = LeftForDay(mnDay) + FOCUSBORDER
        mrcFocusArea.Top = TopForDay(mnDay) + FOCUSBORDER
        mrcFocusArea.Right = mrcFocusArea.Left + mcxColWidth - (2 * FOCUSBORDER)
        mrcFocusArea.Bottom = mrcFocusArea.Top + mcyRowHeight - (2 * FOCUSBORDER)
    
        'update the window (usercontrol.refresh will invalidate
        'everything so call UpdateWindow directly)
        UpdateWindow UserControl.hwnd
    
        'notify of the date change
        UserControl.PropertyChanged "Value"
        RaiseEvent DateChange(dtOld, DateSerial(mnYear, mnMonth, mnDay))
        
    Else 'can't change prop
        mobjRes.RaiseUserError errCantChange, Array("Value")
        
    End If 'can change prop
End Sub 'ChangeValue()

'----------------------------------------------------------------------
' LeftForDay()
'----------------------------------------------------------------------
' Purpose:  Returns the left (X) coodinate for a given day in the
'           current month and year
' Inputs:   day number
' Outputs:  left coordinate
'----------------------------------------------------------------------
Private Function LeftForDay(nDay As Long) As Long
    'the left coordinate for a given day is a function of the
    'weekday (column number) of the day, the column width and
    'the grid's left border
    LeftForDay = ((WeekDay(DateSerial(mnYear, mnMonth, nDay), mnFirstDayOfWeek) - 1) _
                    * mcxColWidth) + mrcCalArea.Left
End Function 'LeftForDay()

'----------------------------------------------------------------------
' TopForDay()
'----------------------------------------------------------------------
' Purpose:  Returns the top (Y) coodinate for a given day in the
'           current month and year
' Inputs:   day number
' Outputs:  top coordinate
'----------------------------------------------------------------------
Private Function TopForDay(nDay As Long) As Long
    Dim nRow As Long
    
    'the top coordinate for a given day is a function of the
    'row number of the day (day + column number of first day of month
    'divided by number of columns), the row height, and the top of the
    'entire grid
    
    'we subtract 2 from the left side of the division since the
    'weekday function is 1-based and since we need to subtract an
    'additional one to make zero-base the day
    nRow = (nDay + WeekDay(DateSerial(mnYear, mnMonth, 1), mnFirstDayOfWeek) - 2) \ NUMCOLS
    
    TopForDay = (nRow * mcyRowHeight) + mrcCalArea.Top
    
End Function 'TopForDay()

'----------------------------------------------------------------------
' DateForRowCol()
'----------------------------------------------------------------------
' Purpose:  Returns the Date for a given row and column in the
'           current calendar grid
' Inputs:   row and column number (zero-based)
' Outputs:  corresponding date
'----------------------------------------------------------------------
Private Function DateForRowCol(nRow As Long, nCol As Long) As Date
    Dim dtFirstDay As Date
    Dim nColFirstDay As Long
    Dim ctDaysDiff As Long
    
    Debug.Assert (nRow < NUMROWS)
    Debug.Assert (nCol < NUMCOLS)
    
    'get the column for the first day of the current month
    'first day is always in row 1
    dtFirstDay = DateSerial(mnYear, mnMonth, 1)
    nColFirstDay = WeekDay(dtFirstDay, mnFirstDayOfWeek) - 1
    
    'how many days away is the current row and column?
    ctDaysDiff = (nCol - nColFirstDay) + (NUMDAYS * nRow)
    
    'calc the hit date by using date arithmetic
    DateForRowCol = DateAdd("d", ctDaysDiff, dtFirstDay)
End Function 'DateForRowCol()

'----------------------------------------------------------------------
' GetPrevMonthDays()
'----------------------------------------------------------------------
' Purpose:  Calculates the first and last day of the previous month
'           that should be displayed before the first day of the
'           of the given month and year
' Inputs:   current month and year
' Outputs:  first and last day of prev month to display
'----------------------------------------------------------------------
Private Sub GetPrevMonthDays(ByVal nCurMonth As Long, ByVal nCurYear As Long, nFirst As Long, nLast As Long)
    Dim dtTemp As Date          'temp date
    Dim nColDayOne As Long      'column of 1st day of cur month
    
    'construct a date to do date math
    dtTemp = DateSerial(nCurYear, nCurMonth, 1)
    
    'determine the column of the first day of the current month
    nColDayOne = WeekDay(dtTemp, mnFirstDayOfWeek)
    
    'if the first day of the current month is in column 1, we
    'don't need to paint any days from the prev month, so return
    'zeros and -1 for the first and last value
    If nColDayOne = 1 Then
        nFirst = 0
        nLast = -1
    Else
        'if there are days to paint, calculate the last and
        'first day using date math
        dtTemp = DateAdd("d", -1, dtTemp)
        nLast = VBA.Day(dtTemp)
        
        dtTemp = DateAdd("d", -(nColDayOne - 2), dtTemp)
        nFirst = VBA.Day(dtTemp)
    End If 'no days to paint
    
End Sub 'GetPrevMonthDays()

'----------------------------------------------------------------------
' LoadMonthNames()
'----------------------------------------------------------------------
' Purpose:  Loads the names of the months into the month selector
'           combo box
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Private Sub LoadMonthNames()
    Dim nMonth As Long
    
    'use the format function to return the system specified
    'long month name for each month
    For nMonth = 1 To 12
        masMonthNames(nMonth - 1) = Format(DateSerial(100, nMonth, 1), "mmmm")
        cbxMonth.AddItem masMonthNames(nMonth - 1)
    Next nMonth
End Sub 'LoadMonthNames()

'----------------------------------------------------------------------
' LoadDayNames()
'----------------------------------------------------------------------
' Purpose:  Loads the names of the days into the day name string array
' Inputs:   none
' Outputs:  none
'----------------------------------------------------------------------
Private Sub LoadDayNames()
    Dim nDay As Long
    Dim sFormat As String
    
    Select Case mnDayNameFormat
        Case calShortName, calMediumName
            sFormat = "ddd"
        
        Case calLongName
            sFormat = "dddd"
    End Select
    
    For nDay = 1 To 7
        'if they want the short format, just take the first char
        If mnDayNameFormat = calShortName Then
            masDayNames(nDay - 1) = Left$(Format(DateSerial(1996, 8, 3 + nDay), sFormat), 1)
        Else
            masDayNames(nDay - 1) = Format(DateSerial(1996, 8, 3 + nDay), sFormat)
        End If
    Next nDay
End Sub 'LoadDayNames()

'----------------------------------------------------------------------
' CopyFont
'----------------------------------------------------------------------
' Purpose:  Copies the contents of one StdFont object to another
' Inputs:   source and destination StdFont object
' Outputs:  none
'----------------------------------------------------------------------
Private Sub CopyFont(fntSource As StdFont, fntDest As StdFont)
    'daveste -- 8/14/96
    'REVIEW:  Is there a better way to do this???!!!
    
    'if the destination is nothing, create a new font object
    If fntDest Is Nothing Then Set fntDest = New StdFont
    
    fntDest.Bold = fntSource.Bold
    fntDest.Charset = fntSource.Charset
    fntDest.Italic = fntSource.Italic
    fntDest.Name = fntSource.Name
    fntDest.Size = fntSource.Size
    fntDest.Strikethrough = fntSource.Strikethrough
    fntDest.Underline = fntSource.Underline
    fntDest.Weight = fntSource.Weight
End Sub 'CopyFont()


