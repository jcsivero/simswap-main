VERSION 5.00
Begin VB.Form frmEditMenu 
   Caption         =   "Form1"
   ClientHeight    =   2868
   ClientLeft      =   -9996
   ClientTop       =   2880
   ClientWidth     =   4332
   LinkTopic       =   "Form1"
   ScaleHeight     =   2868
   ScaleWidth      =   4332
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Co&rtar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Pegad&o especial..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "&Seleccionar todo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invertir selección"
      End
   End
End
Attribute VB_Name = "frmEditMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuEditCopy_Click()
  MsgBox "Coloque el código de Copiar aquí"
End Sub

Private Sub mnuEditCut_Click()
  MsgBox "Coloque el código de Cortar aquí"
End Sub

Private Sub mnuEditDSelectAll_Click()
  MsgBox "Coloque el código de Seleccionar todo aquí"
End Sub

Private Sub mnuEditInvertSelection_Click()
  MsgBox "Coloque el código de Invertir selección aquí"
End Sub

Private Sub mnuEditPaste_Click()
  MsgBox "Coloque el código de Pegar aquí"
End Sub

Private Sub mnuEditPasteSpecial_Click()
  MsgBox "Coloque el código de Pegado especial aquí"
End Sub

Private Sub mnuEditUndo_Click()
  MsgBox "Coloque el código de Deshacer aquí"
End Sub
