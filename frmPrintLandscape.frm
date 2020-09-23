VERSION 5.00
Begin VB.Form frmPrintLandscape 
   BorderStyle     =   0  'None
   ClientHeight    =   1170
   ClientLeft      =   4875
   ClientTop       =   390
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Left            =   2520
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   480
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmPrintLandscape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Timer1.Interval = 500   ' blink label every half second
    
    Timer2.Interval = 5000  'allow blind for five seconds
    
    'is printer orientation Portrait?
    If ckIt = True Then
        'if Portrait then set correct label
        Label1.Caption = "Report will print out in Portriat Orientation "
        'reset ckIt's value
        ckIt = False
        
    Else   'Print in Landscape Orientation
        Label1.Caption = "Report will print out in Landscape Orientation "
        'reset ckIt's value
        ckIt = False
    End If
End Sub

Private Sub Form_Resize()
    Me.Height = Label1.Height
    Me.Width = Label1.Width
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()

 If Label1.ForeColor = &HFF& Then       'Red
        Label1.ForeColor = &HC00000     'Blue
      Else
        Label1.ForeColor = &HFF&
    End If
    
End Sub

Private Sub Timer2_Timer()
If Timer2.Interval = 5000 Then Unload Me
End Sub
