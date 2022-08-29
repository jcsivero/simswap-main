VERSION 5.00
Begin VB.Form frmHelpMenu 
   Caption         =   "Form1"
   ClientHeight    =   2652
   ClientLeft      =   -9996
   ClientTop       =   4056
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2652
   ScaleWidth      =   4800
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Buscar Ayuda acerca de..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de MiApl..."
      End
   End
End
Attribute VB_Name = "frmHelpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Private Sub mnuHelpAbout_Click()
  MsgBox "El código del cuadro Acerca de va aquí"
'  frmAbout.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub

Private Sub mnuHelpSearch_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub
