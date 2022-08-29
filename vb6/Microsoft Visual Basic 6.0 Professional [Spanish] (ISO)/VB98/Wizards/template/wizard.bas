Attribute VB_Name = "modWizard"
Option Explicit

Global Const WIZARD_NAME = "WizardTemplate"

Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

'WinHelp Commands
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELP_QUIT = &H2              '  Cerrar la Ayuda
Public Const HELP_CONTENTS = &H3&         '  Mostrar el índice y el contenido
Public Const HELP_CONTEXT = &H1           '  Mostrar el tema de ulTopic
Public Const HELP_INDEX = &H3             '  Mostrar el índice

Global Const APP_CATEGORY = "Wizards"

Global Const CONFIRM_KEY = "ConfirmScreen"
Global Const DONTSHOW_CONFIRM = "DontShow"


'--------------------------------------------------------------------------
'debe ejecutar este procedimiento desde la ventana Inmediato
'para que agregue la entrada a VBADDIN.INI si no existe ya
'de forma que el componente esté disponible la próxima vez que cargue VB 
'--------------------------------------------------------------------------
Sub AddToINI()
    Debug.Print WritePrivateProfileString("Add-Ins32", WIZARD_NAME & ".Wizard", "0", "VBADDIN.INI")
End Sub

Function GetResString(nRes As Integer) As String
    Dim sTmp As String
    Dim sRetStr As String
  
    Do
        sTmp = LoadResString(nRes)
        If Right(sTmp, 1) = "_" Then
            sRetStr = sRetStr + VBA.Left(sTmp, Len(sTmp) - 1)
        Else
            sRetStr = sRetStr + sTmp
        End If
        nRes = nRes + 1
    Loop Until Right(sTmp, 1) <> "_"
    GetResString = sRetStr
  
End Function

Function GetField(sBuffer As String, sSep As String) As String
    Dim p As Integer
    
    p = InStr(sBuffer & sSep, sSep)
    GetField = VBA.Left(sBuffer, p - 1)
    sBuffer = Mid(sBuffer, p + Len(sSep))
  
End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next
    
    Dim ctl As Control
    Dim obj As Object
    
    'establecer el título del formulario
    If IsNumeric(frm.Tag) Then
        frm.Caption = LoadResString(CInt(frm.Tag))
    End If
    
    'establecer los títulos de los controles con la propiedad
    'Caption para elementos de menú y la propiedad Tag 
    'para los demás controles
    For Each ctl In frm.Controls
        If TypeName(ctl) = "Menu" Then
            If IsNumeric(ctl.Caption) Then
                If Err = 0 Then
                    ctl.Caption = LoadResString(CInt(ctl.Caption))
                Else
                    Err = 0
                End If
            End If
        ElseIf TypeName(ctl) = "TabStrip" Then
            For Each obj In ctl.Tabs
                If IsNumeric(obj.Tag) Then
                    obj.Caption = LoadResString(CInt(obj.Tag))
                End If
                'comprobar si hay información sobre herramientas
                If IsNumeric(obj.ToolTipText) Then
                    If Err = 0 Then
                        obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
                    Else
                        Err = 0
                    End If
                End If
            Next
        ElseIf TypeName(ctl) = "Toolbar" Then
            For Each obj In ctl.Buttons
                If IsNumeric(obj.Tag) Then
                    obj.ToolTipText = LoadResString(CInt(obj.Tag))
                End If
            Next
        ElseIf TypeName(ctl) = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                If IsNumeric(obj.Tag) Then
                    obj.Text = LoadResString(CInt(obj.Tag))
                End If
            Next
        Else
            If IsNumeric(ctl.Tag) Then
                If Err = 0 Then
                    ctl.Caption = GetResString(CInt(ctl.Tag))
                Else
                    Err = 0
                End If
            End If
            'comprobar si hay información sobre herramientas
            If IsNumeric(ctl.ToolTipText) Then
                If Err = 0 Then
                    ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
                Else
                    Err = 0
                End If
            End If
        End If
    Next

End Sub

'==================================================
'Propósito: Reemplace las cadenas <TOPIC_TEXT> en 
'           la cadena de archivo res para colocar 
'           correctamente los tokens localizados
'
'Entradas:  sString = Cadena para buscar y reemplazar
'           sReplacement = Cadena con la que reemplazar el token
'           sReplacement2 = Segunda cadena con la que reemplazar el token
'
'Resultados: Nueva cadena con el token reemplazado
'==================================================
Function ReplaceTopicTokens(sString As String, _
                            sReplacement As String, _
                            sReplacement2 As String) As String
    On Error Resume Next
    
    Dim p As Integer
    Dim sTmp As String
    
    Const TOPIC_TEXT = "<TOPIC_TEXT>"
    Const TOPIC_TEXT2 = "<TOPIC_TEXT2>"
    
    sTmp = sString
    Do
        p = InStr(sTmp, TOPIC_TEXT)
        If p Then
            sTmp = VBA.Left(sTmp, p - 1) + sReplacement + Mid(sTmp, p + Len(TOPIC_TEXT))
        End If
    Loop While p
    
    If Len(sReplacement2) > 0 Then
        Do
            p = InStr(sTmp, TOPIC_TEXT2)
            If p Then
                sTmp = VBA.Left(sTmp, p - 1) + sReplacement2 + Mid(sTmp, p + Len(TOPIC_TEXT2))
            End If
        Loop While p
    End If
    
    ReplaceTopicTokens = sTmp
  
End Function

Public Function GetResData(sResName As String, sResType As String) As String
    Dim sTemp As String
    Dim p As Integer
  
    sTemp = StrConv(LoadResData(sResName, sResType), vbUnicode)
    p = InStr(sTemp, vbNullChar)
    If p Then sTemp = VBA.Left$(sTemp, p - 1)
    GetResData = sTemp
End Function

Function AddToAddInCommandBar(VBInst As Object, sCaption As String, oBitmap As Object) As Object   'Office.CommandBarControl
    On Error GoTo AddToAddInCommandBarErr
    
    Dim c As Integer
    Dim cbMenuCommandBar As Object   'Office.CommandBarControl  'objeto de barra de comandos
    Dim cbMenu As Object
    
    'ver si se encuentra el menú Complementos
    Set cbMenu = VBInst.CommandBars(1).Controls("Complementos")
    If cbMenu Is Nothing Then
        'no está disponible; salimos de la función
        Exit Function
    End If
    
    'agregarlo a la barra de comandos
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    c = cbMenu.Controls.Count - 1
    If cbMenu.Controls(c).BeginGroup And _
        Not cbMenu.Controls(c - 1).BeginGroup Then
        'this s the first addin being added so it needs a separator
        cbMenuCommandBar.BeginGroup = True
    End If
    'establecer el título
    cbMenuCommandBar.Caption = sCaption
    'sin hacer: establecer onaction (necesario en este punto)
    cbMenuCommandBar.OnAction = "hola"
    'copiar el icono al Portapapeles
    Clipboard.SetData oBitmap
    'establecer el icono para el botón
    cbMenuCommandBar.PasteFace
  
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
AddToAddInCommandBarErr:
  
End Function



