VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "FLEXGRID.OCX"
Begin VB.Form formReport 
   Caption         =   "Report"
   ClientHeight    =   4380
   ClientLeft      =   885
   ClientTop       =   2100
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9330
   Begin MSFlexGridLib.MSFlexGrid gridReport 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   1320
      Width           =   1212
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton btnRemind 
      Caption         =   "&Remind"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Caption         =   "Time Report for Pay Period Ending 1/1/2095"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "formReport"
Attribute VB_Base = "0{19C4F559-DF36-11CF-A520-00A0D1003923}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aReport() As Double '3D : days x categories x users
Dim cReceivedReports As Integer 'number of received reports
Dim cReportCategories As Integer 'number of report categories in ReportCategorylist
Dim ReportCategoryList As Variant 'Report categories
Dim ReportPayPeriod As Date     'report payperiod
Dim ReportDate() As Date        'when user sent the report

Public Function CompileReport() As Boolean
'Iterates through all the report messages and extract info
'for the current pay period
On Error GoTo error_olemsg

Dim objReceivFolder As Object
Dim objRepMsg As Object
Dim objmessages As Object

If Not frmCalender.GetDate(ReportPayPeriod) Then
    Exit Function
End If


If objSession Is Nothing Then
    MsgBox "Not logged on"
    CompileReport = False
    Exit Function
End If

'get the receiving folder
GetReceivIPCFolder objReceivFolder
If objReceivFolder Is Nothing Then
    MsgBox "Can't open receive folder"
    CompileReport = False
    Exit Function
End If

'Get message collection from the receiving folder
Set objmessages = objReceivFolder.Messages
If objmessages Is Nothing Then
    MsgBox "Failed to open folder's Messages collection"
    CompileReport = False
    Exit Function
End If
    
'start iterating throuhg the messages
Set objRepMsg = objmessages.getfirst(ReportMsgType)
If objRepMsg Is Nothing Then
    MsgBox "no report msgs found"
    CompileReport = False
    Exit Function
End If


cReceivedReports = 0
Do While Not objRepMsg Is Nothing 'while there are messages
    If Not ProcessMessage(objRepMsg) Then
        CompileReport = False
        Exit Function
    End If
    Set objRepMsg = Nothing
    Set objRepMsg = objmessages.getnext 'next message
Loop


CompileReport = True

Exit Function

error_olemsg:
    MsgBox "Error " & Str(err) & ": " & Error$(err)
    Resume Next

End Function




Function ProcessMessage(objmsg As Object) As Boolean
'If the message is for the right pay period extract and store info

On Error GoTo error_olemsg

Dim tmpPayPeriod As Date
Dim tmpcRepCats As Integer
Dim tmpRepCats As Variant
Dim ind As Integer
Dim PropName As String
Dim var As Variant
Dim day As Integer
Dim userindex As Integer
Dim usrName As String
Dim response As Integer
Dim objFields As Object
Dim msgSentDate As Date
 
'Get msg's fields collection
Set objFields = objmsg.Fields
If objFields Is Nothing Then
    ProcessMessage = True 'ignore this msg
    Exit Function
End If

'get the pay-period
tmpPayPeriod = objFields.Item(PayPeriodPropName)

If tmpPayPeriod <> ReportPayPeriod Then
    ProcessMessage = True   'not intrested in this one
    Exit Function
End If
    
objmsg.Unread = False
objmsg.Update

If cReceivedReports = 0 Then 'first report, has to get the categ. lits
    cReportCategories = objFields.Item(NumCatPropName).Value
    If cReportCategories = 0 Then
        Debug.Print "impossible happend: cReportCats = 0"
        Exit Function
    End If
    ReportCategoryList = objFields.Item(CatPropName).Value
    ReDim aReport(7, cReportCategories, UserList.cUsers)
    ReDim ReportDate(UserList.cUsers)
    
Else 'let's do some validation
    tmpcRepCats = objFields.Item(NumCatPropName).Value
    If tmpcRepCats <> cReportCategories Then
        Debug.Print "number of categories do not match, skipping this message..."
        ProcessMessage = True
        Exit Function
    End If
    tmpRepCats = objFields.Item(CatPropName).Value
    For ind = 0 To tmpcRepCats
        If tmpRepCats(ind) <> ReportCategoryList(ind) Then
            Debug.Print "categories do not match, skipping message..."
            ProcessMessage = True
            Exit Function
        End If
    Next ind
    
End If

usrName = objmsg.sender.Name
'usrName = objFields.Item(NamePropName).Value

userindex = FindUser(usrName)

If E_NOT_FOUND = userindex Then 'the user is not on the list
    response = MsgBox("Received a report from user " & usrName & _
            " who is not on the user list." & Chr(13) & _
            "Would you like to add him/her to the list?", _
            vbYesNo + vbQuestion)
    
    If response = vbYes Then
        'allocate space for the new guy
        ReDim Preserve UserList.aUsers(UserList.cUsers + 1)
        ReDim Preserve aReport(7, cReportCategories, UserList.cUsers + 1)
        ReDim Preserve ReportDate(UserList.cUsers + 1)
        
        'enter him in the list
        UserList.aUsers(UserList.cUsers).DisplayName = usrName
        UserList.aUsers(UserList.cUsers).EntryID = objmsg.sender.id
        UserList.aUsers(UserList.cUsers).ReportIndex = E_NOT_FOUND
        
        'set the index
        userindex = UserList.cUsers
        
        UserList.cUsers = UserList.cUsers + 1
        
    Else
        ProcessMessage = True  'don't care about this one
        Exit Function
    End If
    
End If


'If we are here, everything is cool. Get the data.

'remember when the msg was sent
msgSentDate = objmsg.timesent

If UserList.aUsers(userindex).ReportIndex = E_NOT_FOUND Then
    'if first report from the user
    For ind = 1 To cReportCategories Step 1
        PropName = RepDataPropPrefix & Str(ind)
        var = objFields.Item(PropName)
        For day = 0 To 6 Step 1
            aReport(day, ind - 1, cReceivedReports) = var(day)
        Next day
    Next ind

    UserList.aUsers(userindex).ReportIndex = cReceivedReports
    ReportDate(userindex) = msgSentDate
    cReceivedReports = cReceivedReports + 1
Else
    'if there are more than one report from the same user, user the
    'one that was sent later
    '$
    'make the two loops into one, when sure that they work
    Debug.Print "There is more than one report from " & usrName
    
    If msgSentDate > ReportDate(userindex) Then
        For ind = 1 To cReportCategories Step 1
            PropName = RepDataPropPrefix & Str(ind)
            var = objFields.Item(PropName)
            For day = 0 To 6 Step 1
                aReport(day, ind - 1, UserList.aUsers(userindex).ReportIndex) = var(day)
            Next day
        Next ind
        ReportDate(userindex) = msgSentDate
    End If
    
End If


ProcessMessage = True
Exit Function

error_olemsg:
    MsgBox "Error " & Str(err) & ": " & Error$(err)
    Resume Next

End Function

Function FindUser(strName As String) As Integer
'finds user's positions in the user list given user name
Dim ind As Integer

ind = 0
Do While ind < UserList.cUsers
    If UserList.aUsers(ind).DisplayName = strName Then
        FindUser = ind
        Exit Function
    End If
    ind = ind + 1
Loop

FindUser = E_NOT_FOUND
Exit Function

End Function


Sub ShowGrid()
'uses the extracted data to display the report

Const strNoData As String = "No data"
Const FirstColW As Integer = 2250
Const BorderW As Integer = 30
Dim strDays As Variant
Dim indDays As Integer
Dim indCats As Integer
Dim indUsrs As Integer
Dim indRprt As Integer
Dim sum As Double
Dim total As Double
Dim CellW As Double

strDays = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Total")

gridReport.Cols = 9 'number of elements in strDays+1
gridReport.Rows = UserList.cUsers + 1

'resize columns
CellW = (gridReport.Width - FirstColW - BorderW * gridReport.Cols) _
            / (gridReport.Cols - 1)
gridReport.ColWidth(0) = FirstColW
For indDays = 1 To gridReport.Cols - 1
    gridReport.ColWidth(indDays) = CellW
Next indDays

'display the first row
gridReport.Row = 0
For indDays = 0 To gridReport.Cols - 2
    gridReport.Col = indDays + 1
    gridReport.Text = strDays(indDays)
Next indDays
    
'display the rest of the grid
For indUsrs = 0 To UserList.cUsers - 1 'for all users
    indRprt = UserList.aUsers(indUsrs).ReportIndex
    gridReport.Row = indUsrs + 1
    gridReport.Col = 0
    gridReport.Text = UserList.aUsers(indUsrs).DisplayName
    total = 0
    For indDays = 0 To 6 'for each day
        gridReport.Col = indDays + 1
        If indRprt = E_NOT_FOUND Then
            'no report received from this user
            gridReport.Text = strNoData
            btnRemind.Enabled = True
        Else
            sum = 0 'sum for cats per day
            For indCats = 0 To cReportCategories - 1
                sum = sum + aReport(indDays, indCats, indRprt)
            Next indCats
            gridReport.Text = Str(sum)
            total = total + sum 'total for the week
        End If
    Next indDays
    
    'last column is total
    gridReport.Col = gridReport.Cols - 1
    
    If indRprt <> E_NOT_FOUND Then
        gridReport.Text = Str(total)
    Else
        gridReport.Text = strNoData
    End If
Next indUsrs
    
lblHeader = "Time Report for Pay Period Ending " & ReportPayPeriod

End Sub


Private Sub btnClose_Click()
    Unload Me
    
End Sub


Private Sub btnRemind_Click()
'sends second request message to the users who haven't submitted report

Dim ind As Integer
Dim tmpCats() As String

ReDim tmpCats(cReportCategories)

'put all the cats from variant into a string array
For ind = 0 To cReportCategories - 1
    tmpCats(ind) = ReportCategoryList(ind)
Next ind

formmainsvr.SendRequest cReportCategories, tmpCats, _
            ReportPayPeriod, True
End Sub


Private Sub btnSave_Click()
'save report

On Error GoTo CheckError

Dim indUsrs As Integer
Dim indRprt As Integer
Dim indDays As Integer
Dim indCats As Integer

Open "Report.dat" For Output As #1

Print #1, Tab(24); "Time Report"
Print #1, Tab(20); "Pay period ending " & ReportPayPeriod


For indUsrs = 0 To UserList.cUsers - 1
    Print #1,
    Print #1,
    Print #1, "======================================================================"
    Print #1, "Employee: ", UserList.aUsers(indUsrs).DisplayName
    indRprt = UserList.aUsers(indUsrs).ReportIndex
    If Not indRprt = E_NOT_FOUND Then
        Print #1, Tab(20); _
           "Sun     Mon     Tue     Wed     Thu     Fri     Sat"
        For indCats = 0 To cReportCategories - 1
            Print #1, ReportCategoryList(indCats), Tab(20);
            For indDays = 0 To 6
                Print #1, aReport(indDays, indCats, indRprt); Tab(20 + (1 + indDays) * 8);
            Next indDays
            Print #1,
        Next indCats
    Else
        Print #1, "No data submitted"
    End If
Next indUsrs

Close #1

Exit Sub

CheckError:
MsgBox "Error saving user list"


End Sub


Private Sub Form_Load()
    ShowGrid

End Sub


Private Sub Form_Unload(Cancel As Integer)
'deinit variables global to this module

Dim ind As Integer

For ind = 0 To UserList.cUsers - 1
    UserList.aUsers(ind).ReportIndex = E_NOT_FOUND
Next ind

cReceivedReports = 0
cReportCategories = 0
ReportPayPeriod = Date
ReDim aReport(0, 0, 0)


End Sub


