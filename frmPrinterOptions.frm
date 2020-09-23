VERSION 5.00
Begin VB.Form frmPrinterOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Landscape Orientation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Portrait Orientation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Report Printer Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   248
      Width           =   2655
   End
End
Attribute VB_Name = "frmPrinterOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbForReport_file As String
Dim cn_ForReport As ADODB.Connection
Dim rs_ForReport As ADODB.Recordset

Private Sub Check1_Click()
   
'Print Report in Portrait Orientation

    ckIt = True

    rptParts.Caption = "                               ADO Totals Show Test Report"
    rptParts.WindowState = vbMaximized
    rptParts.Show
    
    'Put the printer orientation form on top of the DataReport Preview
    modPutWindowOnTop
   
    Unload Me

End Sub

Private Sub Check2_Click()

'Print Report in Landscape Orientation
    rptParts.Caption = "                                  ADO Totals Show Test Report"
    rptParts.WindowState = vbMaximized
    rptParts.Show
   
    'Put the printer orientation form on top of the DataReport Preview
    modPutWindowOnTop
    rptParts.Orientation = rptOrientLandscape
    
    Unload Me
    
End Sub

Private Sub Form_Load()
  ' Get the data.
    dbForReport_file = App.Path
    If Right$(dbForReport_file, 1) <> "\" Then dbForReport_file = dbForReport_file & "\"
    dbForReport_file = dbForReport_file & "parts.mdb"

    ' Open a connection.
    Set cn_ForReport = New ADODB.Connection
    cn_ForReport.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & dbForReport_file & ";" & _
        "Persist Security Info=False"
    rptParts.WindowState = vbMaximized
    cn_ForReport.Open
    ' Open the Recordset.
   
    Set rs_ForReport = cn_ForReport.Execute("SELECT Item,Cost,Quantity,[Cost]* [Quantity] AS Total FROM query1 ORDER BY Item", , adCmdText)
    ' Connect the Recordset to the DataReport.
    Set rptParts.DataSource = rs_ForReport
End Sub

Private Sub Form_Resize()

    Me.Height = Shape1.Height
    Me.Width = Shape1.Width
    'Center the Form on the Screen
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Set rs_ForReport = Nothing
    Set cn_ForReport = Nothing
    Unload Me

End Sub
