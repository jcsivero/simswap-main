Attribute VB_Name = "client"
Option Explicit


Global objSession As Object

Type WeekDataType
    Day(7) As Double
End Type



Sub Main()

Dim bFlag As Boolean

On Error GoTo error_olemsg

'create session and logon
bFlag = Util_CreateSessionAndLogon()

If Not bFlag Then Exit Sub

'go try to find request message and show the report form
formReport.Init

If Not objSession Is Nothing Then
    'logoff
    objSession.Logoff
End If


Exit Sub

    
error_olemsg:
    If Not bFlag Then
        MsgBox "Error " & Str(Err) & ": " & Error$(Err)
        End
    End If
    

End Sub

Function Util_CreateSessionAndLogon() As Boolean
    'create session and logon
    On Error GoTo err_CreateSessionAndLogon

    Set objSession = CreateObject("MAPI.Session")
    objSession.Logon
    Util_CreateSessionAndLogon = True
    Exit Function

err_CreateSessionAndLogon:
    Set objSession = Nothing
    
    If Not (Err = -2147221229) Then  'if not user cancel
        MsgBox "Unrecoverable Error:" & Err
    End If
    
    Util_CreateSessionAndLogon = False
    Exit Function
    
error_olemsg:
    MsgBox "Error " & Str(Err) & ": " & Error$(Err)
    Resume Next

End Function




