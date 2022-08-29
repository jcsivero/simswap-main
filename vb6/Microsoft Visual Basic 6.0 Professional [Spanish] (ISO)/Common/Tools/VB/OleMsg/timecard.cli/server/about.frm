VERSION 5.00
Begin VB.Form formAbout 
   Caption         =   "About"
   ClientHeight    =   2925
   ClientLeft      =   1290
   ClientTop       =   1545
   ClientWidth     =   5790
   Height          =   3330
   Left            =   1230
   LinkTopic       =   "formAbout"
   ScaleHeight     =   2925
   ScaleWidth      =   5790
   Top             =   1200
   Width           =   5910
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Copyright (c) 1995 Microsoft Corporation"
      Height          =   372
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   4092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OLE Messaging Sample"
      Height          =   252
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2172
   End
End
Attribute VB_Name = "formAbout"
Attribute VB_Base = "0{CFF16A23-C697-11CF-A520-00A0D1003923}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btnOK_Click()
    Unload Me

End Sub


