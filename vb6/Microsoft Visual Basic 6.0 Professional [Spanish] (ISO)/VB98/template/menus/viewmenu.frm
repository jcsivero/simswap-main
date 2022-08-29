VERSION 5.00
Begin VB.Form frmViewMenu 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   -9990
   ClientTop       =   285
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   6675
   Begin VB.Menu mnuView 
      Caption         =   "&Ver"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Barra de herramientas"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Barra de e&stado"
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLargeIcons 
         Caption         =   "&Iconos grandes"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "Ico&nos pequeños"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&Lista"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Detalles"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Org&anizar iconos"
         Begin VB.Menu mnuVAIByName 
            Caption         =   "por &Nombre"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "por &Tipo"
         End
         Begin VB.Menu mnuVAIBySize 
            Caption         =   "por T&amaño"
         End
         Begin VB.Menu mnuVAIByDate 
            Caption         =   "por &Fecha"
         End
      End
      Begin VB.Menu mnuViewLineUpIcons 
         Caption         =   "Alin&ear iconos"
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "Actuali&zar"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Opciones..."
      End
   End
End
Attribute VB_Name = "frmViewMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

Private Sub mnuVAIByDate_Click()
'  lvListView.SortKey = DATE_COLUMN
End Sub

Private Sub mnuVAIByName_Click()
'  lvListView.SortKey = NAME_COLUMN
End Sub

Private Sub mnuVAIBySize_Click()
'  lvListView.SortKey = SIZE_COLUMN
End Sub

Private Sub mnuVAIByType_Click()
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewDetails_Click()
'  lvListView.View = lvwReport
End Sub

Private Sub mnuViewLargeIcons_Click()
'  lvListView.View = lvwIcon
End Sub

Private Sub mnuViewLineUpIcons_Click()
'  lvListView.Arrange = lvwAutoLeft
End Sub

Private Sub mnuViewList_Click()
'  lvListView.View = lvwList
End Sub

Private Sub mnuViewOptions_Click()
'  frmOptions.Show vbModal
End Sub

Private Sub mnuViewRefresh_Click()
  MsgBox "El código de Actualizar va aquí"
End Sub

Private Sub mnuViewSmallIcons_Click()
'  lvListView.View = lvwSmallIcon
End Sub

Private Sub mnuViewStatusBar_Click()
  If mnuViewStatusBar.Checked Then
'    sbStatusBar.Visible = False
    mnuViewStatusBar.Checked = False
  Else
'    sbStatusBar.Visible = True
    mnuViewStatusBar.Checked = True
  End If
End Sub

Private Sub mnuViewToolbar_Click()
  If mnuViewToolbar.Checked Then
'    tbToolBar.Visible = False
    mnuViewToolbar.Checked = False
  Else
'    tbToolBar.Visible = True
    mnuViewToolbar.Checked = True
  End If
End Sub
