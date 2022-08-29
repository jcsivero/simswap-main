VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' PENDIENTE: declarar objeto Recordset ADO local. Por ejemplo:
'Private WithEvents rs As ADODB.Recordset

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    ' PENDIENTE:  devolver el recordset apropiado según DataMember. Por ejemplo:
    
    'Select Case DataMember
    'Case ""             ' Predeterminado
    '    Set Data = Nothing
    'Case Else           ' Predeterminado
    '    Set Data = rs
    'End Select
End Sub

