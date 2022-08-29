Attribute VB_Name = "modDHTML"
'PutProperty: almacena informaci�n en un cookie llamando a esta funci�n.
'             Los datos necesarios son la propiedad nombrada 
'             y el valor de la propiedad que desea almacenar.
'
'             Los datos opcionales son:
'               expires : especifica una fecha que define el tiempo de vida v�lido
'                         de la propiedad. Cuando haya caducado, la propiedad
'                         dejar� de estar almacenada ni se podr� utilizar.

Public Sub PutProperty(objDocument As HTMLDocument, strName As String, vntValue As Variant, Optional Expires As Date)

     objDocument.cookie = strName & "=" & CStr(vntValue) & _
        IIf(CLng(Expires) = 0, "", "; expires=" & Format(CStr(Expires), "ddd, dd-mmm-aa hh:mm:ss") & " GMT") ' & _

End Sub

'GetProperty: obtiene el valor de una propiedad llamando a esta funci�n.
'             Los datos necesarios son la propiedad nombrada y el valor
'             de retorno de la funci�n es el valor actual de la propidad.
'             Si no se puede encontrar la propiedad o ha caducado,
'             el valor de retorno ser� una cadena vac�a.
'
Public Function GetProperty(objDocument As HTMLDocument, strName As String) As Variant
    Dim aryCookies() As String
    Dim strCookie As Variant
    On Local Error GoTo NextCookie

    'Divide el objeto cookie en una matriz de cookies.
    aryCookies = Split(objDocument.cookie, ";")
    For Each strCookie In aryCookies
        If Trim(VBA.Left(strCookie, InStr(strCookie, "=") - 1)) = Trim(strName) Then
            GetProperty = Trim(Mid(strCookie, InStr(strCookie, "=") + 1))
            Exit Function
        End If
NextCookie:
        Err = 0
    Next strCookie
End Function


