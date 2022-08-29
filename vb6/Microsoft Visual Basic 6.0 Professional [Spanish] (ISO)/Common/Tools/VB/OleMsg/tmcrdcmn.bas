Attribute VB_Name = "common"
Option Explicit

'named property names
Global Const CatPropName As String = "ReportCategories"
Global Const NumCatPropName As String = "NumReportCategories"
Global Const PayPeriodPropName As String = "PayPeriod"
Global Const RepDataPropPrefix As String = "ReportedTime"
'$for testing
'Global Const NamePropName As String = "UserName"

'messages class names
'request is an IPM message so that a user can see it in the inbox
Global Const RequestMsgType = "IPM.TimeCardSample.Request"
'report is an IPC message
Global Const ReportMsgType = "IPC.TimeCardSample.Report"

'vb constants for mouse pointer
Global Const WaitCursor As Integer = 11
Global Const DefaultCursor As Integer = 0
