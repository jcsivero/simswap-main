Attribute VB_Name = "modLoadRes"
' Este procedimiento cargar� cadenas de recursos asociadas a controles
' de un formulario seg�n el Id. de recurso almacenado en la propiedad
' Tag de un control.

' La cadena de recursos se cargar� en la propiedad de un control as�:
' Objeto      Propiedad
' Form        Caption
' Menu        Caption
' TabStrip    Caption, ToolTipText
' Toolbar     ToolTipText
' ListView    ColumnHeader.Text

Sub LoadResStrings(frm As Form)
  On Error Resume Next
  
  Dim ctl As Control
  Dim obj As Object
  
  'establecer el t�tulo del formulario
  If IsNumeric(frm.Tag) Then
    frm.Caption = LoadResString(CInt(frm.Tag))
  End If
  
  'establecer los t�tulos de los controles con la
  'propiedad Caption para los t�tulos de men� y la propiedad 
  'Tag para los dem�s controles
  For Each ctl In frm.Controls
    Err.Clear
    If TypeName(ctl) = "Menu" Then
      If IsNumeric(ctl.Caption) Then
        If Err = 0 Then
          ctl.Caption = LoadResString(CInt(ctl.Caption))
        End If
      End If
    ElseIf TypeName(ctl) = "TabStrip" Then
      For Each obj In ctl.Tabs
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.Caption = LoadResString(CInt(obj.Tag))
        End If
        'comprobar si hay informaci�n sobre herramientas
        If IsNumeric(obj.ToolTipText) Then
          If Err = 0 Then
            obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
          End If
        End If
      Next
    ElseIf TypeName(ctl) = "Toolbar" Then
      For Each obj In ctl.Buttons
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.ToolTipText = LoadResString(CInt(obj.Tag))
        End If
      Next
    ElseIf TypeName(ctl) = "ListView" Then
      For Each obj In ctl.ColumnHeaders
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.Text = LoadResString(CInt(obj.Tag))
        End If
      Next
    Else
      If IsNumeric(ctl.Tag) Then
        If Err = 0 Then
          ctl.Caption = LoadResString(CInt(ctl.Tag))
        End If
      End If
      'comprobar si hay informaci�n sobre herramientas
      If IsNumeric(ctl.ToolTipText) Then
        If Err = 0 Then
          ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
        End If
      End If
    End If
  Next

End Sub
