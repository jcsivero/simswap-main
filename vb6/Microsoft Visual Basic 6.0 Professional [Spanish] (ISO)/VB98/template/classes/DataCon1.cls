VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComplexDataConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' PENDIENTE: declarar una variable recordset como origen de datos para esta clase. Por ejemplo:

'Private WithEvents RS As ADODB.RecordSet

Public Property Get DataSource() As DataSource
    ' PENDIENTE: devolver el objeto DataSource para esta clase. Por ejemplo:
   
    'Set DataSource = RS.DataSource
End Property
Public Property Set DataSource(ByVal objDataSource As DataSource)
    '  PENDIENTE: establecer el objeto DataSource para esta clase. Por ejemplo:
    
    'Set RS.DataSource = objDataSource
    'PropertyChanged "DataSource"
End Property
Public Property Get DataMember() As DataMember
    ' PENDIENTE: devolver DataMember. Por ejemplo:
    
    'Set DataMember = RS.DataMember
End Property
Public Property Let DataMember(ByVal DataMember As DataMember)
    ' PENDIENTE: asignar DataMember nuevo. Por ejemplo:
    
    'Set RS.DataMember = DataMember
    'PropertyChanged "DataMember"
End Property

