VERSION 5.00
Begin VB.Form formReport 
   Caption         =   "Time Report Form"
   ClientHeight    =   5295
   ClientLeft      =   930
   ClientTop       =   2175
   ClientWidth     =   10485
   Height          =   5700
   Left            =   870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10485
   Top             =   1830
   Width           =   10605
   Begin VB.TextBox txtTo 
      Height          =   288
      Left            =   1080
      TabIndex        =   25
      Top             =   120
      Width           =   5052
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   8
      Left            =   8760
      TabIndex        =   24
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   7
      Left            =   7680
      TabIndex        =   23
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   6
      Left            =   6600
      TabIndex        =   22
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   5
      Left            =   5520
      TabIndex        =   21
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   4
      Left            =   4440
      TabIndex        =   20
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   3
      Left            =   3360
      TabIndex        =   19
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   2
      Left            =   2280
      TabIndex        =   18
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   1
      Left            =   1200
      TabIndex        =   17
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtCell 
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "&Clear All"
      Height          =   372
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send"
      Height          =   372
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtPayPeriod 
      Height          =   288
      Left            =   7080
      TabIndex        =   16
      Top             =   960
      Width           =   2172
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   4452
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Total"
      Height          =   252
      Left            =   8880
      TabIndex        =   15
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Sat"
      Height          =   252
      Left            =   7800
      TabIndex        =   13
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Fri"
      Height          =   252
      Left            =   6720
      TabIndex        =   12
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Thu"
      Height          =   252
      Left            =   5640
      TabIndex        =   11
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Wed"
      Height          =   252
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Tue"
      Height          =   252
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Mon"
      Height          =   252
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Sun"
      Height          =   252
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   732
   End
   Begin VB.Label lblCategories 
      Alignment       =   2  'Center
      Caption         =   "Categories"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "Pay Period"
      Height          =   252
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   492
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "formReport"
Attribute VB_Base = "0{D624D371-C698-11CF-A520-00A0D1003923}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Customizable = False
Const RowSize As Integer = 9

Dim objRequestMsg As Object 'the request message
Dim ReportCategories  As Variant
Dim CatNum As Integer   'number of report categories in ReportCategories
Dim PayPeriod As Date

Dim ReportData() As WeekDataType

Public Sub Init()
'if there is a request message in the inbox, show the form
If FindRequestMsg Then
    ShowReportForm
End If

End Sub

Function NumFromString(txtstr As String) As Double

If IsNumeric(txtstr) Then
    NumFromString = Val(txtstr)
Else
    NumFromString = 0
End If

End Function

Public Function ShowReportForm() As Boolean
'if can succesfully extract necessary prop from the
'request message show the form

 On Error GoTo error_olemsg

    If objRequestMsg Is Nothing Then
        MsgBox "No  active request message"
        ShowReportForm = False
        Exit Function
    End If
    
    If Not ExtractProps Then
        ShowReportForm = False
        Exit Function
    End If
    
    formReport.Show 1
       
    ShowReportForm = True
    Exit Function
    
error_olemsg:
    MsgBox "Error " & Str(Err) & ": " & Error$(Err)
    Resume Next

End Function

Private Function ExtractProps() As Boolean
'Reads number of report categories, report categiry names
' and pay period from the reques message

Dim objFields As Object

On Error GoTo error_olemsg

If objRequestMsg Is Nothing Then
    MsgBox "no message"
    ExtractProps = False
    Exit Function
End If
    
'get msg's fields collection
Set objFields = objRequestMsg.Fields
If objFields Is Nothing Then
    MsgBox "Error reading request message"
    Exit Function
End If

'number of categories
CatNum = objFields.Item(NumCatPropName).Value

'report categories
ReportCategories = objFields.Item(CatPropName).Value
    
'pay period
PayPeriod = objFields.Item(PayPeriodPropName)
    
ExtractProps = True
Exit Function

error_olemsg:
    MsgBox "Error " & Str(Err) & ": " & Error$(Err)
    ExtractProps = False
    Exit Function
    
End Function


Private Function FindRequestMsg() As Boolean
'finds request message in the inbox
'(request message has message class RequestMsgType)
'RequestMsgType is a const defined in tmcrdcmn.bas
'This functon doesn't deal very well with the situation when
'there are more than one request message in the inbox,
'It just gets the one returned by Inbox.Messges.GetFirst(RequestMsgType)
'This can be changed to showing the listbox with all the request messages
'and letting user choose the one he/she wants to user

On Error GoTo error_olemsg

Dim objInbox As Object
Dim objMessages As Object
Dim objMessage As Object

    If objSession Is Nothing Then
        MsgBox "Not logged on"
        FindRequestMsg = False
        Exit Function
    End If
    
    'get the inbox
    Set objInbox = objSession.Inbox
    If objInbox Is Nothing Then
        MsgBox "Failed to open Inbox"
        FindRequestMsg = False
        Exit Function
    End If
    
    'get the inbox's message collection
    Set objMessages = objInbox.Messages
    If objMessages Is Nothing Then
        MsgBox "Failed to open folder's Messages collection"
        FindRequestMsg = False
        Exit Function
    End If
    
    Set objMessage = objMessages.GetFirst(RequestMsgType)
    If objMessage Is Nothing Then
        MsgBox "no request msg found"
        FindRequestMsg = False
        Exit Function
    End If
    
    Set objRequestMsg = objMessage
        
    FindRequestMsg = True
    Exit Function
    
error_olemsg:
    MsgBox "Error " & Str(Err) & ": " & Error$(Err)
    Resume Next
    
End Function


Private Sub ShowGrid()
'displays the a appropriate number of edit boxes
'on the form

Const initX As Integer = 120
Const initY As Integer = 2160
Const deltaX As Integer = 1080
Const deltaY As Integer = 600

Dim row As Integer
Dim col As Integer
Dim ind As Integer


For row = 1 To CatNum - 1
    For col = 1 To RowSize
        ind = row * RowSize + col - 1
        Load txtCell(ind)
        txtCell(ind).Top = initY + row * deltaY
        txtCell(ind).Left = initX + (col - 1) * deltaX
        txtCell(ind).Visible = True
    Next col
Next row

For row = 0 To CatNum - 1
    txtCell(row * RowSize).Text = ReportCategories(row)
    txtCell(row * RowSize).Enabled = False
    txtCell((row + 1) * RowSize - 1).Enabled = False
Next row

End Sub


Function SumUpRow(RowNum As Integer) As Double
    
Dim ind As Integer
Dim total As Double

total = 0

For ind = 1 To RowSize - 2 Step 1
    total = total + NumFromString(txtCell.Item((RowNum - 1) * RowSize + ind).Text)
Next ind

SumUpRow = total

End Function


Private Sub btnClear_Click()

Dim row As Integer
Dim col As Integer
Dim ind As Integer

For row = 0 To CatNum - 1 Step 1
    For col = 2 To RowSize
        ind = row * RowSize + col - 1
        txtCell(ind).Text = ""
    Next col
Next row
End Sub

Private Sub btnSend_Click()
'generates and sends a report message

On Error GoTo error_olemsg

Dim objReportMsg As Object
Dim obj As Object
Dim objR As Object
Dim prop As Object
Dim objFields As Object

Dim PropName As String
Dim row As Integer
Dim col As Integer
Dim ind As Integer

MousePointer = WaitCursor

ReDim ReportData(CatNum)

Dim dbgstr As String

dbgstr = ""

'get the data
For row = 0 To CatNum - 1 Step 1
    For col = 2 To RowSize - 1 'don't need total
        ind = row * RowSize + col - 1
        ReportData(row).Day(col - 2) = NumFromString(txtCell(ind).Text)
        dbgstr = dbgstr & ReportData(row).Day(col - 2) & " "
    Next col
    Debug.Print dbgstr
    dbgstr = ""
Next row

If objSession Is Nothing Then
    MsgBox "Not logged on"
    Exit Sub
End If

'create a new message in the outbox
Set objReportMsg = objSession.Outbox.Messages.Add
If objReportMsg Is Nothing Then
    MsgBox "Can't add a prop"
    Exit Sub
End If

'set the message class
objReportMsg.Type = ReportMsgType

'address the message to the sender of the request message
Set objR = objReportMsg.Recipients.Add(EntryId:=objRequestMsg.Sender.ID, _
                                        Name:=objRequestMsg.Sender.Name)
If objR Is Nothing Then
    MsgBox "Can't set recipient"
    Exit Sub
End If

'get msg field collection
Set objFields = objReportMsg.Fields
If objFields Is Nothing Then
    MsgBox "Internal error. (can't access msg's field collecton)"
    Exit Sub
End If

'report data is transmitted in named properties.
'name for the property containing data for the i-th category is "i"
'i = 1, 2, ..., NumberOfCategories
For row = 1 To CatNum Step 1
    PropName = RepDataPropPrefix & Str(row)
    'we can't write:
    'Set obj = objFields.Add(Name:=PropName, _
                             Class:=vbDouble + vbArray, _
                             Value:=ReportData(row - 1.Day)
    'because of the way VB passes array parameters
    'so we first add a property and then set its value
    Set obj = objFields.Add(Name:=PropName, _
                                    Class:=vbDouble + vbArray)
    If obj Is Nothing Then
        MsgBox "Can't add a prop"
        Exit Sub
    End If
    obj.Value = ReportData(row - 1).Day
Next row

Set obj = objFields.Add(Name:=CatPropName, _
                                    Class:=vbString + vbArray)
If obj Is Nothing Then
        MsgBox "Can't add a prop"
        Exit Sub
    End If
obj.Value = ReportCategories

Set obj = objFields.Add(Name:=NumCatPropName, _
            Class:=vbInteger, _
            Value:=CatNum)
If obj Is Nothing Then
        MsgBox "Can't add a prop"
        Exit Sub
End If

Set prop = objFields.Add(Name:=PayPeriodPropName, _
            Class:=vbDate, _
            Value:=PayPeriod)
If prop Is Nothing Then
        MsgBox "Can't add a prop"
        Exit Sub
End If

'$for testing only, later this field (txtName)
'will be read-only
'Set obj = objFields.Add(Name:=NamePropName, _
            Class:=vbString, _
            Value:=txtName.Text)
'If obj Is Nothing Then
'        MsgBox "Can't add a prop"
'        Exit Sub
'End If

objReportMsg.Send showDialog:=False

MousePointer = DefaultCursor

Unload Me

Exit Sub

error_olemsg:
    MsgBox "Error " & Str(Err) & ": " & Error$(Err)
    Resume Next

End Sub

Private Sub Categories_Click()

End Sub

Private Sub Form_Load()
    txtTo.Text = objRequestMsg.Sender.Name
    txtTo.Enabled = False
    
    txtName.Text = objSession.CurrentUser.Name
    txtName.Enabled = False
    
    txtPayPeriod.Text = PayPeriod
    txtPayPeriod.Enabled = False
    
    ShowGrid
    
End Sub





Private Sub Form_Unload(Cancel As Integer)

    CatNum = 0
    Set objRequestMsg = Nothing
    
End Sub




Private Sub txtCell_LostFocus(Index As Integer)
'do some validation
Dim indTot As Integer
    If (Index Mod RowSize = 0) Or ((Index + 1) Mod RowSize = 0) Then
        Debug.Print "LostFocus from a disable control"
        Exit Sub
    End If
    
    If txtCell.Item(Index).Text = "" Then
        Exit Sub
    End If
    
    If IsNumeric(txtCell.Item(Index).Text) And _
        Val(txtCell.Item(Index).Text) >= 0 And _
        Val(txtCell.Item(Index).Text) <= 24 Then
        
        indTot = (Index \ RowSize) * RowSize + RowSize - 1
        txtCell.Item(indTot).Text = SumUpRow(Index \ RowSize + 1)
    Else
        MsgBox "Has to be number of hours." + Chr(13) + _
                "(Can not be greater than 24)"
        txtCell(Index).SetFocus
    End If
    
End Sub


