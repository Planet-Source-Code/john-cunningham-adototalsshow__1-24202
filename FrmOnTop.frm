VERSION 5.00
Begin VB.Form FrmOnTop 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   525
   ClientLeft      =   5730
   ClientTop       =   375
   ClientWidth     =   1095
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   1095
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Hi There!"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
