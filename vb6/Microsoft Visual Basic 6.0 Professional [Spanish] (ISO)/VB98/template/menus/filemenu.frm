VERSION 5.00
Begin VB.Form frmFileMenu 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   -9996
   ClientTop       =   1980
   ClientWidth     =   6684
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   6684
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Cerrar"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Gua&rdar como..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Guardar &todo"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Propiedades"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "C&onfigurar impresora..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "&Vista preliminar"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimir..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "E&nviar..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileClose_Click()
  MsgBox "El código de Cerrar va aquí"
End Sub

Private Sub mnuFileExit_Click()
  'descargar el formulario
  Unload Me
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "El código de Nuevo archivo va aquí"
End Sub

Private Sub mnuFileOpen_Click()
  MsgBox "El código de Abrir va aquí"
End Sub

Private Sub mnuFilePrint_Click()
  MsgBox "El código de Imprimir va aquí"
End Sub

Private Sub mnuFilePrintPreview_Click()
  MsgBox "El código de Vista preliminar va aquí"
End Sub

Private Sub mnuFilePrintSetup_Click()
  MsgBox "El código de Configurar impresora va aquí"
End Sub

Private Sub mnuFileProperties_Click()
  MsgBox "El código de propiedades va aquí"
End Sub

Private Sub mnuFileSave_Click()
  MsgBox "El código de Guardar archivo va aquí"
End Sub

Private Sub mnuFileSaveAll_Click()
  MsgBox "El código de Guardar todo va aquí"
End Sub

Private Sub mnuFileSaveAs_Click()
  MsgBox "El código de Guardar como va aquí"
End Sub

Private Sub mnuFileSend_Click()
  MsgBox "El código de Enviar va aquí"
End Sub
