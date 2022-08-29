VERSION 5.00
Begin VB.Form frmExpFileMenu 
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
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "En&viar a"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Ca&mbiar nombre"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Propiedades"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Cerrar"
      End
   End
End
Attribute VB_Name = "frmExpFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileClose_Click()
  'descargar el formulario
  Unload Me
End Sub

Private Sub mnuFileDelete_Click()
  MsgBox "El código de Eliminar va aquí"
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "El código de Nuevo va aquí"
End Sub

Private Sub mnuFileOpen_Click()
  MsgBox "El código de Abrir va aquí"
End Sub

Private Sub mnuFileProperties_Click()
  MsgBox "El código de Propiedades va aquí"
End Sub

Private Sub mnuFileRename_Click()
  MsgBox "El código de Cambiar nombre va aquí"
End Sub

Private Sub mnuFileSend_Click()
  MsgBox "El código de Enviar va aquí"
End Sub
