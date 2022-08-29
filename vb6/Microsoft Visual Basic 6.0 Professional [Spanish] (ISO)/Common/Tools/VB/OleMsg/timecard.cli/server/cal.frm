VERSION 5.00
Begin VB.Form frmCalender 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Calendar"
   ClientHeight    =   4830
   ClientLeft      =   885
   ClientTop       =   1245
   ClientWidth     =   4230
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Height          =   5235
   KeyPreview      =   -1  'True
   Left            =   825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3.354
   ScaleMode       =   5  'Inch
   ScaleWidth      =   2.937
   Top             =   900
   Width           =   4350
   Begin VB.Frame fraCalender 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dates:"
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      Begin VB.PictureBox picWeekdays 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2415
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.PictureBox picCal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1452
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2415
         TabIndex        =   3
         Top             =   872
         Width           =   2412
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   1092
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   372
         Left            =   1320
         TabIndex        =   1
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Line linDivider 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   120
         X2              =   2640
         Y1              =   2416
         Y2              =   2416
      End
      Begin VB.Line linDivider 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   120
         X2              =   2640
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line linDivider 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   2640
         Y1              =   856
         Y2              =   856
      End
      Begin VB.Line linDivider 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   120
         X2              =   2640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblMonth 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Deciembre  1943"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1812
      End
      Begin VB.Image picGoMonth 
         Height          =   180
         Index           =   1
         Left            =   2880
         Picture         =   "CAL.frx":0000
         Top             =   300
         Width           =   180
      End
      Begin VB.Image picGoMonth 
         Height          =   180
         Index           =   0
         Left            =   600
         Picture         =   "CAL.frx":04D2
         Top             =   300
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmCalender"
Attribute VB_Base = "0{CFF16A29-C697-11CF-A520-00A0D1003923}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Customizable = False
Option Explicit

Dim fDirty%

Dim fRet As Boolean

Const kfMultiselectDates = False  '** can multiple dates be selected at a time?

Const kiDayIndexMax = 41    '** picCal displays 41 visible dates
Private Type SingleDay       '** each visible date has info in a SingleDate rec
    iTop As Integer
    iLeft As Integer
    lForeColor As Long      '** kBlack = current month; kDkGray = prev/next month
    sCaption As String      '** date text ("1"-"31")
End Type

Dim gfrmCal As Form         '** form containing cal frame

'** cal graphic-related vars
Dim giCurYear%, giCurMonth%         '** current month/year visible
Dim giDayWidth%, giDayHeight%       '** dimensions of the 41 visible dates
Dim gsMonthes$(1 To 12)             '** stores month names
Dim gaDays(0 To 41) As SingleDay    '** array of info on visible dates
Dim giTodayIndex%           '** if current month visible then giTodayIndex is graphical inset
Dim gfCreateNewCal%
Dim fFirstClick%
Dim gsUsername$

'** cal date selection vars
'** cal has two kinds of selections
'**     main selection: made by click, shift-click, or drag
'**     ctrl selections: made by ctrl-click
Dim gdSelStart As Date   '** start of main selection block
Dim gdSelEnd As Date     '** end of main selection block
Dim gadCtrlSelect(0 To 100) As Date  '** array of current ctrl-clicked dates; erased on non-ctrl-mousedown
                                        '** if date in main sel then non-selected, else then selected
Dim giMaxCtrlSelectIndex As Integer  '** highest index of gadCtrlSelect in use; init to -1

'** cal mouse vars
Dim giLastSelIndex As Integer    '** last index selected by drag; used to validate MouseOver calls during drags
Dim gdLastDateClicked As Date       '** last index clicked; used as next start for selection block
Dim gfExitedGray%           '** after dragging over gray date to switch month, has mouse left gray dates on new month yet?


'** colors used cal
Const kLtGray = &HC0C0C0
Const kDkGray = &H808080
Const kBlack = &H0&
Const kWhite = &HFFFFFF
Const kBlue = &HFF0000

Private Sub ClearOldSelection(ByVal dStartNew As Date, ByVal dEndNew As Date, ByVal dStartOld As Date, ByVal dEndOld As Date)
    '** redraws all dates between dStartOld & dStartNew but not between dStartNew & dEndNew
    '**     as unselected.
    '** CalMousedown uses ClearOldSelection to deselect dates in the previous selection
    '**     block that are not in the new selection block
    
    Dim dTmp As Date        '** used as utility date
    Dim dFirstDate As Date  '** first date vis in picCal; may be gray from previous month
    Dim iIndex%             '** gaDay index to deselect

    If dEndOld = 0 Or dStartOld = 0 Then Exit Sub
    
    '** switch dStartNew with dEndNew if dStartNew is higher
    If dStartNew > dEndNew Then
        dTmp = dStartNew
        dStartNew = dEndNew
        dEndNew = dTmp
    End If
    '** switch dStartOld with dEndOld if dStartOld is higher
    If dStartOld > dEndOld Then
        dTmp = dStartOld
        dStartOld = dEndOld
        dEndOld = dTmp
    End If
    
    '** if dStartOld comes before the dates visible,
    '** then set dStartOld to first date visible
    If gaDays(0).lForeColor = kDkGray Then
        dFirstDate = DateSerial(giCurYear, giCurMonth - 1, CInt(gaDays(0).sCaption))
    Else
        dFirstDate = DateSerial(giCurYear, giCurMonth, CInt(gaDays(0).sCaption))
    End If
    If dFirstDate > dStartOld Then dStartOld = dFirstDate
    
    '** if dEndOld comes after the dates visible,
    '** then set dEndOld to last date visible
    If gaDays(kiDayIndexMax).lForeColor = kDkGray Then
        dTmp = DateSerial(giCurYear, giCurMonth + 1, CInt(gaDays(kiDayIndexMax).sCaption))
    Else
        dTmp = DateSerial(giCurYear, giCurMonth, CInt(gaDays(kiDayIndexMax).sCaption))
    End If
    If dTmp < dEndOld Then dEndOld = dTmp
    
    '** deselect all dates necessary
    For dTmp = dStartOld To dEndOld
        If dTmp < dStartNew Or dTmp > dEndNew Then
            iIndex = dTmp - dFirstDate
            DrawDay iIndex, kLtGray
        End If
    Next dTmp
End Sub


Private Sub DrawDay(ByVal iIndex%, ByVal lColor&)
    Dim picCal As PictureBox '** vb4 workaround
    
    Set picCal = gfrmCal!picCal
    '** draws an individual day
    
    '** draw background of day
    '** lColor = kBlue if selected, ltGray if unselected
    picCal.Line (gaDays(iIndex).iLeft, gaDays(iIndex).iTop)-(gaDays(iIndex).iLeft + giDayWidth - Screen.TwipsPerPixelX, gaDays(iIndex).iTop + giDayHeight - Screen.TwipsPerPixelY), lColor&, BF
    
    '** if this day is today, inset in 3d
    If iIndex = giTodayIndex Then
        ThreeDRect picCal, gaDays(iIndex).iLeft + Screen.TwipsPerPixelX * 1, gaDays(iIndex).iTop + Screen.TwipsPerPixelY * 1, gaDays(iIndex).iLeft + giDayWidth - Screen.TwipsPerPixelX * 1, gaDays(iIndex).iTop + giDayHeight - Screen.TwipsPerPixelX * 1, True
    End If
    
    '** print the number of the day
    picCal.CurrentX = (giDayWidth - picCal.TextWidth(gaDays(iIndex).sCaption)) / 2 + gaDays(iIndex).iLeft
    picCal.CurrentY = (giDayHeight - picCal.TextHeight(gaDays(iIndex).sCaption)) / 2 + gaDays(iIndex).iTop
    If lColor = kBlue And gaDays(iIndex).lForeColor <> kDkGray Then
        picCal.ForeColor = kWhite '** if selected, kWhite
    Else
        picCal.ForeColor = gaDays(iIndex).lForeColor
    End If
    picCal.Print gaDays(iIndex).sCaption
End Sub





Private Sub fMoreGrayDates()

End Sub
Private Function fIsDateSelected%(ByVal iYear%, ByVal iMonth%, ByVal iDay%)
    Dim dSrc As Date, i%
    
    dSrc = DateSerial(iYear, iMonth, iDay)
    If (dSrc <= gdSelEnd And dSrc >= gdSelStart) Or (dSrc >= gdSelEnd And dSrc <= gdSelStart) Then
        fIsDateSelected = True
    End If
    For i = 0 To giMaxCtrlSelectIndex
        If gadCtrlSelect(i) = dSrc Then
            If fDateInBetween(gadCtrlSelect(i), gdSelStart, gdSelEnd) Then
                fIsDateSelected = False
            Else
                fIsDateSelected = True
            End If
        End If
    Next i
End Function
 

Private Sub InitCalControls()
    Dim i%, sWeekdays$
    Dim iOldScaleMode%, iOnePixelX%, iOnePixelY%
    Dim iRow%, iColumn%
    Dim picWeekdays As PictureBox '** vb4 workaround
    
    Set picWeekdays = gfrmCal!picWeekdays
    iOldScaleMode = gfrmCal.ScaleMode
    gfrmCal.ScaleMode = 1
    iOnePixelX = Screen.TwipsPerPixelX
    iOnePixelY = Screen.TwipsPerPixelY
    
    gfrmCal!lblMonth.Left = (gfrmCal!fraCalender.Width - gfrmCal!lblMonth.Width) / 2
    gfrmCal!picGoMonth(0).Left = gfrmCal!lblMonth.Left - (gfrmCal!picGoMonth(0).Width + 3 * iOnePixelX)
    gfrmCal!picGoMonth(1).Left = gfrmCal!lblMonth.Left + gfrmCal!lblMonth.Width + 3 * iOnePixelX
    
    gfrmCal!cmdOK.Top = gfrmCal!fraCalender.Height - (8 * iOnePixelY + gfrmCal!cmdOK.Height)
    gfrmCal!cmdCancel.Top = gfrmCal!fraCalender.Height - (8 * iOnePixelY + gfrmCal!cmdCancel.Height)
    gfrmCal!picCal.Width = gfrmCal!fraCalender.Width - 16 * iOnePixelX
    gfrmCal!picCal.Height = gfrmCal!cmdOK.Top - (gfrmCal!picCal.Top) - 10 * iOnePixelY

    giDayHeight = gfrmCal!picCal.Height / 6
    giDayWidth = gfrmCal!picCal.Width / 7
    picWeekdays.Width = gfrmCal!picCal.Width
    picWeekdays.Left = gfrmCal!picCal.Left
    
    gfrmCal!linDivider(0).X1 = gfrmCal!picCal.Left
    gfrmCal!linDivider(0).X2 = gfrmCal!picCal.Left + gfrmCal!picCal.Width
    gfrmCal!linDivider(2).Y1 = gfrmCal!picCal.Top + gfrmCal!picCal.Height + iOnePixelY
    gfrmCal!linDivider(2).Y2 = gfrmCal!linDivider(2).Y1
    gfrmCal!linDivider(3).Y1 = gfrmCal!linDivider(2).Y1 + iOnePixelY
    gfrmCal!linDivider(3).Y2 = gfrmCal!linDivider(2).Y1 + iOnePixelY

    For i = 1 To 3
        gfrmCal!linDivider(i).X1 = gfrmCal!linDivider(0).X1
        gfrmCal!linDivider(i).X2 = gfrmCal!linDivider(0).X2
    Next i
    
    sWeekdays = "SMTWTFS"
    For i = 0 To 6
        picWeekdays.CurrentX = i * giDayWidth + giDayWidth / 2
        picWeekdays.Print Mid(sWeekdays, i + 1, 1);
    Next i

    For i = 0 To kiDayIndexMax '41 number of days
        gaDays(i).iLeft = iColumn * giDayWidth
        gaDays(i).iTop = iRow * giDayHeight
        iColumn = iColumn + 1
        If iColumn = 7 Then
            iColumn = 0
            iRow = iRow + 1
        End If
    Next i
    gfrmCal.ScaleMode = iOldScaleMode
End Sub



Function iDayIndex%(iYear%, iMonth%, iDay%)

   iDayIndex = WeekDay(DateSerial(iYear, iMonth, 1)) + iDay - 2
End Function

Private Sub MakeSelection(ByVal dStartNew As Date, ByVal dEndNew As Date, ByVal dStartOld As Date, ByVal dEndOld As Date)
    Dim dTmp
    Dim dFirstDate As Date
    Dim iDayDiff%

    If dEndOld = 0 Or dStartOld = 0 Then Exit Sub
    If dStartNew > dEndNew Then
        dTmp = dStartNew
        dStartNew = dEndNew
        dEndNew = dTmp
    End If
    If dStartOld > dEndOld Then
        dTmp = dStartOld
        dStartOld = dEndOld
        dEndOld = dTmp
    End If
    'reset dStartOld to first of cal if efficient
    If gaDays(0).lForeColor = kDkGray Then
        dFirstDate = DateSerial(giCurYear, giCurMonth - 1, CInt(gaDays(0).sCaption))
    Else
        dFirstDate = DateSerial(giCurYear, giCurMonth, CInt(gaDays(0).sCaption))
    End If
    If dFirstDate > dStartNew Then dStartNew = dFirstDate
    
    'reset dEndOld to first of cal if efficient
     If gaDays(kiDayIndexMax).lForeColor = kDkGray Then
        dTmp = DateSerial(giCurYear, giCurMonth + 1, CInt(gaDays(kiDayIndexMax).sCaption))
    Else
        dTmp = DateSerial(giCurYear, giCurMonth, CInt(gaDays(kiDayIndexMax).sCaption))
    End If
    If dTmp < dEndNew Then dEndNew = dTmp
    
    For dTmp = dStartNew To dEndNew '** ALERT: THIS DOES NOT INCLUDE OLD NOT SELOTHERS!!!
        If dTmp >= dEndOld Or dTmp <= dStartOld Then
            iDayDiff = dTmp - dFirstDate
            DrawDay iDayDiff, kBlue
        End If
    Next dTmp
End Sub

Private Sub DrawCalender()
    '** draws the current dates and selection
    
    Dim dStartDate As Date '** first date of month
    Dim iDayOfWeek%
    Dim iDaysInMonth%
    Dim i%
    Dim iDayInPrevMonth%
    Dim iCurDay%

    gfrmCal!lblMonth = gsMonthes(giCurMonth) & " " & CStr(giCurYear) '** set month label
    dStartDate = DateSerial(giCurYear, giCurMonth, 1)
    
    '** if this is current month, find which index is today
    If (giCurYear = Year(Now)) And (Month(Now) = giCurMonth) Then
        giTodayIndex = iDayIndex(Year(Now), Month(Now), day(Now))
    Else
        giTodayIndex = -1
    End If
    
    '** find how many days are in current month
    '** to get: subtract first day of next month by first day of this month
    If giCurMonth = 12 Then
        iDaysInMonth = DateSerial(giCurYear + 1, 1, 1) - dStartDate
    Else
        iDaysInMonth = DateSerial(giCurYear, giCurMonth + 1, 1) - dStartDate
    End If

    iDayOfWeek = WeekDay(dStartDate)    '** set day of week which the first day of the month falls on
    '** draw all the days of this month
    For i = iDayOfWeek - 1 To (iDayOfWeek - 1) + iDaysInMonth - 1
        iCurDay% = iCurDay% + 1
        gaDays(i).sCaption = Str(iCurDay%)
        If fIsDateSelected(giCurYear, giCurMonth, iCurDay%) Then
            gaDays(i).lForeColor = kBlack
            DrawDay i, kBlue
        Else
            gaDays(i).lForeColor = kBlack
            DrawDay i, kLtGray
        End If
    Next i
    
    '** calculate the number of days in previous month
    If giCurMonth = 1 Then
        iDayInPrevMonth = dStartDate - DateSerial(giCurYear - 1, 12, 1)
    Else
        iDayInPrevMonth = dStartDate - DateSerial(giCurYear, giCurMonth - 1, 1)
    End If

    '** draw in the last gray days of previous month
    For i = 0 To iDayOfWeek - 2
        iCurDay% = iDayInPrevMonth - (iDayOfWeek - i) + 2
        gaDays(i).sCaption = iCurDay%
        gaDays(i).lForeColor = kDkGray
        If fIsDateSelected(giCurYear, giCurMonth - 1, iCurDay%) Then
            DrawDay i, kBlue
        Else
            DrawDay i, kLtGray
        End If
    Next i

    '** draw in the first gray days of next month
    iCurDay% = 0
    For i = (iDayOfWeek - 1) + iDaysInMonth To 41
        iCurDay% = iCurDay% + 1
        gaDays(i).lForeColor = kDkGray
        gaDays(i).sCaption = iCurDay%
        If fIsDateSelected(giCurYear, giCurMonth + 1, iCurDay%) Then
            DrawDay i, kBlue
        Else
            DrawDay i, kLtGray
        End If
    Next i
End Sub

Private Sub CalInitialize(frmCal As Form)
     '** initializes cal vars and controls
    '** frmCal = the form with cal frame control

    fRet = False
    
    Dim i%

    Set gfrmCal = frmCal
    InitCalControls     '** place and initialize controls in cal frame control

    '** init global cal variables
    giCurYear = Year(Now)
    giCurMonth = Month(Now)
    For i = LBound(gadCtrlSelect) To UBound(gadCtrlSelect)
        gadCtrlSelect(i) = 0
    Next i
    
    gdSelStart = DateSerial(Year(Now), Month(Now), day(Now)) '** init main selection to today
    gdSelEnd = gdSelStart
    giMaxCtrlSelectIndex = -1
    
    giLastSelIndex = -1
    gfExitedGray = True

    For i = 1 To 12 '** fill gsMonthes array with month names
        gsMonthes(i) = Format$(DateSerial(giCurYear, i, 1), "mmmm")
    Next

    DrawCalender    '** draw the current month
    fFirstClick = True
End Sub

Private Sub ThreeDRect(picCanvas As PictureBox, iLeft%, iTop%, iRight%, iBottom%, fOut%)
    Dim lColor1&, lColor2&
    
    If fOut Then
        lColor1 = kDkGray
        lColor2 = kWhite
    Else
        lColor1 = kWhite
        lColor2 = kDkGray
    End If

    picCanvas.ForeColor = lColor1
    picCanvas.Line (iLeft - 1, iTop - 2)-(iLeft - 1, iBottom + 2)
    picCanvas.Line (iLeft - 2, iTop - 2)-(iLeft - 2, iBottom + 2)
    picCanvas.Line (iLeft - 2, iTop - 1)-(iRight + 2, iTop - 1)
    picCanvas.Line (iLeft - 2, iTop - 2)-(iRight + 2, iTop - 2)
    
    picCanvas.ForeColor = lColor2
    picCanvas.Line (iRight + 1, iTop - 1)-(iRight + 1, iBottom + 2)
    picCanvas.Line (iRight + 2, iTop - 2)-(iRight + 2, iBottom + 2)
    picCanvas.Line (iLeft - 1, iBottom + 1)-(iRight + 2, iBottom + 1)
    picCanvas.Line (iLeft - 2, iBottom + 2)-(iRight + 2, iBottom + 2)
End Sub



Private Sub CalMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '** select or de-select a date
    '** handles click, shift-click and ctrl-click
    '** MouseOver calls CalMousedown with shift for dragging
    
    Dim dNewDate As Date    '** date selected
    Dim iIndex%             '** gaDay index of date clicked
    Dim iDay%               '** current day (1-31)
    
    '** if not left mouse button then exit
    If (Button And vbLeftButton) <= 0 Then Exit Sub
    
    If fFirstClick = True Then Shift = 0
    fFirstClick = False
    
    '** find the gaDay index of date clicked on
    iIndex = (Int(Y / giDayHeight) * 7) + Int(X / giDayWidth)
    If iIndex < 0 Or iIndex > kiDayIndexMax Then Exit Sub

    iDay = CInt(gaDays(iIndex).sCaption)
    
    '** if the click is on a grayed out date then make new month visible
    If gaDays(iIndex).lForeColor = kDkGray Then
        If iDay < 15 Then
            CalGoMonth 1 '** switch to prev month
        Else
            CalGoMonth 0 '** switch to next month
        End If
        iIndex = iDayIndex(giCurYear, giCurMonth, iDay) '** adjust iIndex to new month
        If (Shift And vbShiftMask) > 0 Then gfExitedGray = False '** set flag to prevent another month switch if new month
    End If                                                       '** has grayed out date under mouse

    dNewDate = DateSerial(giCurYear, giCurMonth, iDay)
    If kfMultiselectDates And (Shift And vbShiftMask) > 0 Then '** shift-key down
        ClearCtrlSelects '** clear all ctrl-key selected dates
        ClearOldSelection gdLastDateClicked, dNewDate, gdSelStart, gdSelEnd
        MakeSelection gdLastDateClicked, dNewDate, gdSelStart, gdSelEnd
        gdSelEnd = dNewDate
        gdSelStart = gdLastDateClicked
    ElseIf kfMultiselectDates And (Shift And vbCtrlMask) > 0 Then '** ctrl-key down
        CtrlSelectDate iIndex, dNewDate
    Else '**simple mouse click, no keys down
        ClearCtrlSelects  '** clear all ctrl-key selected dates
        ClearOldSelection dNewDate, dNewDate, gdSelStart, gdSelEnd
        gdSelStart = dNewDate
        gdSelEnd = dNewDate
        DrawDay iIndex, kBlue
        gdLastDateClicked = dNewDate
    End If
    
End Sub

Private Sub CalGoMonth(iIndex%)
    '** if index = 0, make previous month visible
    '** else, make next month visible
    
    If iIndex% = 0 Then
        giCurMonth = giCurMonth - 1
        If giCurMonth = 0 Then
            giCurMonth = 12
            giCurYear = giCurYear - 1
        End If
    Else
        giCurMonth = giCurMonth + 1
        If giCurMonth = 13 Then
            giCurMonth = 1
            giCurYear = giCurYear + 1
        End If
    End If
    DrawCalender    '** draw new month
End Sub

Private Sub CalMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iIndex% '** index of gaDay that mouse is over
    
    '** set gfExitedGray to true if mouse is not over gray date
    iIndex = (Int(Y / giDayHeight) * 7) + Int(X / giDayWidth) '** calculate index
    If iIndex >= 0 And iIndex <= kiDayIndexMax Then
        If gaDays(iIndex).lForeColor = kBlack Then
            gfExitedGray = True
        End If
    End If
    
    '** if the mouse is not on the same index as last mousemove
    '** and the left mouse button is down
    If kfMultiselectDates And ((Button And vbLeftButton) > 0) And (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 And iIndex <> giLastSelIndex And gfExitedGray = True Then
        giLastSelIndex = iIndex
        CalMouseDown Button, vbShiftMask, X, Y '** simulate mousedown with shiftkey
    End If
        
End Sub

Private Function fDateInBetween(dSrc As Date, dStart As Date, dEnd As Date)
    If (dSrc <= dEnd And dSrc >= dStart) Or (dSrc >= dEnd And dSrc <= dStart) Then
        fDateInBetween = True
    End If
    
End Function

Private Sub ClearCtrlSelects()
    '** clear gadCtrlSelect array; no ctrl-key selection blocks
    '** redraw the ex-ctrl-selected dates
    
    Dim i%, dFirstDate As Date, iIndex%
    
    If gaDays(0).lForeColor = kDkGray Then
        dFirstDate = DateSerial(giCurYear, giCurMonth - 1, CInt(gaDays(0).sCaption))
    Else
        dFirstDate = DateSerial(giCurYear, giCurMonth, CInt(gaDays(0).sCaption))
    End If
    
    For i = 0 To giMaxCtrlSelectIndex '** loop through gadCtrlSelect array
        If gadCtrlSelect(i) <> 0 Then '** if valid ctrl-selection
            If fDateInBetween(gadCtrlSelect(i), gdSelStart, gdSelEnd) Then '** redraw as selected day
                iIndex = gadCtrlSelect(i) - dFirstDate
                If iIndex > -1 And iIndex <= kiDayIndexMax Then
                    DrawDay iIndex, kBlue
                End If
            Else    '** redraw as unselected day (not in selection)
                iIndex = gadCtrlSelect(i) - dFirstDate
                If iIndex > -1 And iIndex <= kiDayIndexMax Then
                    DrawDay iIndex, kLtGray
                End If
            End If
        End If
        gadCtrlSelect(i) = 0 '** clear to 0 (turn off)
    Next i
    giMaxCtrlSelectIndex = -1
End Sub


Private Sub CtrlSelectDate(iIndex%, dNewDate As Date)
    '** perform a ctrl-click on a dNewDate
    '** if this date was selected then deselect; if this date was unselected then select
    '** at least one date MUST be selected at any time
    
    Dim fValid% '** is this ctrl-click valid?
    Dim i%, dTmp As Date '** utility variables
    Dim iExists% '** does this date exist in gadCtrlSelect array? if yes, holds index
    Dim iStep% '** which way do we loop?
    Dim iNumSelMain%, iNumSelCtrl% '** number of dates highlighted in main sel block/ctrl-click array
    
    
    '** first, check if this is a valid ctrl-click
    '** if this causes no dates to be selected than it is INVALID
    
    '** how many dates are selected within the main selection block?
    If gdSelStart > gdSelEnd Then '** do we have to loop through selection backwards?
        iStep = -1  '** yes, gdSelEnd comes first
    Else
        iStep = 1   '** no, gsSelStart comes first
    End If
    
    '** loop through main selection block keeping tally of selected dates within
    For dTmp = gdSelStart To gdSelEnd Step iStep
        If fIsDateSelected(Year(dTmp), Month(dTmp), day(dTmp)) = True Then
            iNumSelMain = iNumSelMain + 1
            If iNumSelMain > 1 Then Exit For
        End If
    Next dTmp
    dTmp = 0    '** clear loop variable
    
    If iNumSelMain > 1 Then  '** multiple dates selected, ok to ctrl-click
        fValid = True
    Else '** if 0 or 1 dates are selected, ctrl-click may not be valid
        '** how many ctrl-click dates are selected? keep tally in iNumSelCtrl
        For i = 0 To giMaxCtrlSelectIndex
            If gadCtrlSelect(i) > 0 And Not fDateInBetween(gadCtrlSelect(i), gdSelStart, gdSelEnd) Then
                iNumSelCtrl = iNumSelCtrl + 1
                If iNumSelCtrl > 1 Then Exit For
            End If
        Next i
        
        If iNumSelMain = 0 And iNumSelCtrl = 1 Then '** if we only have one selected date
                                                    '** and it is a ctrl-click
            '** find that date; store in dTmp
            For i = 0 To giMaxCtrlSelectIndex
                If gadCtrlSelect(i) > 0 And Not fDateInBetween(gadCtrlSelect(i), gdSelStart, gdSelEnd) Then
                    dTmp = gadCtrlSelect(i)
                    Exit For
                End If
            Next i
            If dTmp <> dNewDate Then '** ctrl-click valid if selected date does
                fValid = True        '**    not equal the clicked
            End If
        ElseIf iNumSelMain = 1 And iNumSelCtrl = 0 Then '** if we have one selected date
                                                        '** and it is in the main sel block
            '** if the date just ctrl-clicked isn't the sole selected date than valid
            If Not fIsDateSelected(Year(dNewDate), Month(dNewDate), day(dNewDate)) Then
                fValid = True
            End If
        Else
            fValid = True '** valid; multiple ctrl-click selections
        End If
    End If
    
    If fValid = True Then '** this is a valid ctrl click
        '** does this ctrl-click already exist in the gadCtrlSelect array? if so, find it
        iExists = -1
        For i = 0 To giMaxCtrlSelectIndex
            If gadCtrlSelect(i) = dNewDate Then
                iExists = i
                Exit For
            End If
        Next i
        
        If iExists > -1 Then '** yes, this ctrl-click already exists
            '** since the user is reclicking an already selected ctrl-click,
            '**     this is essentially identical to clearing it
            '** first, draw the selection/deselection
            If fDateInBetween(gadCtrlSelect(iExists), gdSelStart, gdSelEnd) Then
                DrawDay iIndex, kBlue
            Else
                DrawDay iIndex, kLtGray
            End If
            
            gadCtrlSelect(iExists) = 0 '** clear this ctrl-click from array
            '** adjust giMaxCtrlSelectIndex to point to last valid ctrl-click
            '**     in the gadCtrlSelect array
            If iExists = giMaxCtrlSelectIndex Then
                giMaxCtrlSelectIndex = giMaxCtrlSelectIndex - 1
                If giMaxCtrlSelectIndex > -1 Then
                    While giMaxCtrlSelectIndex > 0 And gadCtrlSelect(giMaxCtrlSelectIndex) = 0
                        giMaxCtrlSelectIndex = giMaxCtrlSelectIndex - 1
                    Wend
                    If giMaxCtrlSelectIndex = 0 And gadCtrlSelect(0) = 0 Then giMaxCtrlSelectIndex = -1
                End If
            End If
        Else '** this ctrl-click does not exist already
            '** find the first available (empty) gadCtrlSelect date
            i = 0
            While gadCtrlSelect(i) <> 0
                i = i + 1
            Wend
            
            gadCtrlSelect(i) = dNewDate '** set to new date
            '** draw this ctrl-click
            If fDateInBetween(gadCtrlSelect(i), gdSelStart, gdSelEnd) Then
                DrawDay iIndex, kLtGray
            Else
                DrawDay iIndex, kBlue
            End If
            If i > giMaxCtrlSelectIndex Then giMaxCtrlSelectIndex = i '** reset giMax if necessary
        End If
        gdLastDateClicked = dNewDate
    End If
End Sub



Private Function ValidatePayPeriod(dateToValidate As Date) As Boolean
'in this sample we require that date identifying a pay period be a Friday

If WeekDay(dateToValidate, vbSunday) = vbFriday Then
    ValidatePayPeriod = True
Else
    MsgBox "The date has to be a Friday"
    ValidatePayPeriod = False
End If

End Function







Private Sub JumpToFirstSelected()
    '** set the month/year to the month/year with first selected date
    
    Dim i%
    Dim dFirstDate As Date
    
    If gdSelStart < gdSelEnd Then
        dFirstDate = gdSelStart
    Else
        dFirstDate = gdSelEnd
    End If
         
    For i = 0 To giMaxCtrlSelectIndex
        If gadCtrlSelect(i) < dFirstDate And Not fDateInBetween(gadCtrlSelect(i), gdSelStart, gdSelEnd) Then
            dFirstDate = gadCtrlSelect(i)
        End If
    Next i
    giCurMonth = Month(dFirstDate)
    giCurYear = Year(dFirstDate)
    DrawCalender
End Sub



Public Function GetDate(DateToSet As Date) As Boolean

frmCalender.Show 1

If fRet Then
    DateToSet = gdSelStart
    GetDate = True
Else
    GetDate = False
End If

End Function





Private Sub cmdCancel_Click()

fRet = False

Unload Me

End Sub

Private Sub cmdOK_Click()

If ValidatePayPeriod(gdSelStart) Then
    fRet = True
    Unload Me
End If

End Sub

Private Sub Form_Load()
    CalInitialize Me
End Sub








Private Sub picCal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     fDirty = True
     CalMouseDown Button, Shift, X, Y
End Sub

Private Sub picCal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CalMouseMove Button, Shift, X, Y
End Sub

Private Sub picCal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    giLastSelIndex = -1
End Sub

Private Sub picGoMonth_Click(Index As Integer)
    CalGoMonth Index
End Sub








