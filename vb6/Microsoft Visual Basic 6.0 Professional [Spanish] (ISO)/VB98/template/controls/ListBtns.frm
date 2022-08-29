VERSION 5.00
Begin VB.Form frmListButtons 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   2880
   ClientTop       =   3210
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   3330
   Begin VB.ListBox lstItems 
      DragIcon        =   "Cuadro de lista de botón.frx":0000
      Height          =   2895
      IntegralHeight  =   0   'False
      Left            =   450
      TabIndex        =   4
      Top             =   165
      Width           =   2280
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2790
      Picture         =   "Cuadro de lista de botón.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "5011"
      Top             =   1215
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2790
      Picture         =   "Cuadro de lista de botón.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "5012"
      Top             =   1695
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDelete 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2790
      Picture         =   "Cuadro de lista de botón.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "5010"
      Top             =   735
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   330
      Left            =   2790
      Picture         =   "Cuadro de lista de botón.frx":0748
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "5009"
      Top             =   255
      UseMaskColor    =   -1  'True
      Width           =   330
   End
End
Attribute VB_Name = "frmListButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Dim sTmp As String
  sTmp = InputBox("Escriba el nuevo elemento que desea agregar:")
  If Len(sTmp) = 0 Then Exit Sub
  lstItems.AddItem sTmp
End Sub

Private Sub cmdDelete_Click()
  If lstItems.ListIndex > -1 Then
    If MsgBox("Eliminar '" & lstItems.Text & "'?", vbQuestion + vbYesNo) = vbYes Then
      lstItems.RemoveItem lstItems.ListIndex
    End If
  End If
End Sub

Private Sub cmdUp_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstItems
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'no se puede subir el primer elemento
    'subir elemento
    .AddItem .Text, nItem - 1
    'quitar elemento antiguo
    .RemoveItem nItem + 1
    'seleccionar el elemento que se ha movido
    .Selected(nItem - 1) = True
  End With
End Sub

Private Sub cmdDown_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstItems
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'no se puede bajar el último elemento
    'bajar elemento
    .AddItem .Text, nItem + 2
    'quitar elemento antiguo
    .RemoveItem nItem
    'seleccionar el elemento que se ha movido
    .Selected(nItem + 1) = True
  End With
End Sub

Private Sub lstItems_DragDrop(Source As Control, X As Single, Y As Single)
  Dim i As Integer
  Dim nID As Integer
  Dim sTmp As String
  
  If Source.Name <> "lstItems" Then Exit Sub
  If lstItems.ListCount = 0 Then Exit Sub

  With lstItems
    i = (Y \ TextHeight("A")) + .TopIndex
    If i = .ListIndex Then
      'colocar encima de sí mismo
      Exit Sub
    End If
    If i > .ListCount - 1 Then i = .ListCount - 1
    nID = .ListIndex
    sTmp = .Text
    If (nID > -1) Then
      sTmp = .Text
    .RemoveItem nID
    .AddItem sTmp, i
    .ListIndex = .NewIndex
    End If
  End With
  SetListButtons
End Sub

Sub lstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then lstItems.Drag
End Sub

Private Sub lstItems_Click()
  SetListButtons
End Sub

Sub SetListButtons()
  Dim i As Integer
  i = lstItems.ListIndex
  'establecer el estado de los botones de desplazamiento
  cmdUp.Enabled = (i > 0)
  cmdDown.Enabled = ((i > -1) And (i < (lstItems.ListCount - 1)))
  cmdDelete.Enabled = (i > -1)
End Sub

