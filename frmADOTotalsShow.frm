VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmADOTotalsShow 
   Caption         =   "ADO - Totals - DataGrid Test"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "frmADOTotalsShow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   6615
      Begin VB.TextBox txtDB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
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
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   20
         Text            =   "txtDB(2)"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
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
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   18
         Text            =   "txtDB(1)"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDB 
         BackColor       =   &H00FFC0C0&
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
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   17
         Text            =   "txtDB(0)"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtRecordCount 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Text            =   "txtRecordCount"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdMoveBack 
         Height          =   495
         Left            =   120
         Picture         =   "frmADOTotalsShow.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move Back"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdMoveNext 
         Height          =   495
         Left            =   5760
         Picture         =   "frmADOTotalsShow.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move Forward"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdReportShow 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblQty 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCost 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cost"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   390
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      DefColWidth     =   80
      ForeColor       =   16744576
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTotalSales 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtTotalQty 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   6600
      Width           =   6375
   End
   Begin VB.Label lblTotalSales 
      Caption         =   "Total Sales"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblTotalQty 
      Caption         =   "Total Qty"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "frmADOTotalsShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'*                     Project: ADO Totals                                  *
'****************************************************************************
' Modules:        OnTopModule.bas - Make API call to put a Window on Top
'                 modtxtEffect.bas - see About Form
'                 modWebEmail.bas  - send email
'****************************************************************************
' Description: This project demonstrates a number of important ADO Data Base
'              - VB methods and utilities.
'
'1.  Paramount of these utilites, from my perspective, is how to get the totals
'    from a database field as well as how to multiply the items in two DB Fields
'    and then create another field to show the answer.  In addition a running
'    total is kept, showing the Sum of the multiplied fields.  This is a feature
'    that any inventory data base requires.

'2.  Make an ADO Connection without using the ADO Data Control.  When a
'    program uses the ADO Data Control(s) it hard connects the
'    data base to a given directory ("C:\My Documents\MyDataBase.mdb"), or
'    worse yet the directory that the program is design / built in;
'
'    (C:\Program Files\Microsoft Visual Studio\VB98\MyDataBaseProgram\etc").
'
'    This connection causes havoc when the product is installed to a client's
'    computer.  Of course a work around for this is to make a DSN Connection,
'    but this requires the inclusion of numerous files in your setup/install
'    program.  When the connection is made through the use of code as opposed
'    to using the ADO Data Control the Database file is located / installed in
'    the program's deployed directory and is eaisly connected to.

'3.  Deploy and show all of the data using the MS DataGrid Control.  Program
'    shows how to fill and set the Data Grid.  This demo allows adding and
'    deleting of records directly on the Data Grid (Although this method is
'    not recommended due to lack of ease, the better method is to use the
'    programs Text Boxes).

'4.  Program sets up a TextBox and recordset movement procedure that
'    emulates the familiar data control methods, i.e., The user can move
'    through the records - movement is apparent in both the Data Grid and
'    the Text Boxes, and the Current Record and Number of Records are continously
'    displayed and updated when records are added or deleted.

'5.  Program makes use of the MS Report Designer in lieu of Crystal Reports.
'    I have used Crystal Reports in the past and had numerous problems. I now
'    use the MS Data Report exclusively.  If you have any questions/problems
'    with this demo it will be in this area.  Study the
'    DataEnviornment and DataReport Designers Structure

' ==========================================================================
' ====           Full Credit to Colin T. Green Email: cgts@cgts.com.au  ====
' ====           for his help with;                                     ====
' ====           the DataGrid Methods, the Requery method and the       ====
' ====           display method for the totals.                         ====
' ==========================================================================
' ====           John P. Cunningham - johnpc7@home.com                  ====
' ====          Web Site:  http://johnpc.freeservers.com                ====
'===========================================================================

'***************************************************************************
'**** NB!                                                               ****
'**** Make sure to open Project-References and select                   ****
'**** Microsoft ActiveX Data Object Library 2.0 or higher.              ****
'***************************************************************************

Option Explicit


                       '**** Form Level Declarations ****

'********************************************************************************
'* Be sure to add a Reference to Ms ActiveX Data Objects 2.x Library to Project *
'********************************************************************************

Dim db_file As String                'Name of DataBase
Dim SQLstmt As String                'SQL Statement String(s)
Dim cn As ADODB.Connection           'Connect to the Main ADO Data Type
Attribute cn.VB_VarHelpID = -1
Dim cn1 As ADODB.Connection          'Connect to secondary ADO Data Type
Dim rs1 As ADODB.Recordset           'Primary Record Source Name
Attribute rs1.VB_VarHelpID = -1
Dim rs2 As ADODB.Recordset           'Aggregate Record Source Name
Attribute rs2.VB_VarHelpID = -1
Dim rs3 As ADODB.Recordset           'Record Source to allow updates
Dim RcCount As Integer               'Record Counter
Dim RetVal As Integer                'Variable used for Messages Boxes
Dim strCnn As String                 'String to connect cn1
Dim strItem As String                'String to represent Item Field/Record Addition
Dim strCost As String                'String to represent Cost Field/Record Addition
Dim strQuantity As String            'String to represent Quantity Field/Record Addition
Dim booRecordAdded As Boolean        'Record added test/check
Dim ij As Integer                    'index for Text Box Array

Private Sub OpenData()
'   App.Path never has a \ character at the end
    db_file = App.Path & "\Parts.mdb"
'   Open main connection.
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    cn.Open
'   Once this connection is open you can use it throughout your application
    SQLstmt = "SELECT Item,Cost,Quantity,[Cost]* [Quantity] AS Total FROM query1"
'   Get the records.
    Set rs1 = New ADODB.Recordset
    rs1.Open SQLstmt, cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Set DataGrid1.DataSource = rs1
    
    With rs1
        .MoveFirst
        txtDB(0) = !Item
    End With
'---------------------------------------------------------------------------------------
'   Get an agregate recordset to calculate the totals
    SQLstmt = "SELECT sum(Quantity) as SumOfQuantities, sum(Totals) "
    SQLstmt = SQLstmt + " as SumOfTotal FROM query1 "
    Set rs2 = New ADODB.Recordset
    rs2.Open SQLstmt, cn, adOpenStatic, adLockOptimistic, adCmdText
'---------------------------------------------------------------------------------------
End Sub

Private Sub Update_Controls()
'In order to update the Total Quantities and Total Sales
    rs1.Requery
    rs2.Requery
    
    txtTotalQty.Text = rs2.Fields("sumOfQuantities")
    txtTotalSales.Text = Format(rs2.Fields("SumOfTotal"), "$##0.00")
    
End Sub

Private Sub formatDataGrid()
    With DataGrid1
        .Columns(0).Width = 1800 'Set for maximum anticipated width
        .Columns(1).NumberFormat = "Currency"
        .Columns(3).NumberFormat = "Currency"
        .HoldFields
    End With
End Sub

Private Sub setTotalBoxes()
    'set the location for the Total Quantity and Total Sales TextBoxes
    txtTotalQty.Top = DataGrid1.Top + DataGrid1.Height + 200
    txtTotalQty.Left = lblTotalQty.Left + 1300
    txtTotalQty.Width = DataGrid1.Columns(2).Width
    
    lblTotalQty.Top = txtTotalQty.Top
    lblTotalQty.Left = 150
    lblTotalQty = "Total Quantity:"
    
    
    txtTotalSales.Top = DataGrid1.Top + DataGrid1.Height + 200
    txtTotalSales.Left = DataGrid1.Columns(3).Left
    txtTotalSales.Width = DataGrid1.Columns(3).Width
    
    lblTotalSales.Top = txtTotalSales.Top
    lblTotalSales.Left = 3400
    lblTotalSales = "Total Sales:"
    
End Sub

Private Sub About_Click()

    frmAbout.Show
    
End Sub

Private Sub cmdAdd_Click()

'Clear out the TextBoxes to accept Input of new record
For ij = 0 To 2
    txtDB(ij) = ""
    txtDB(ij).BackColor = &HFFC0FF
Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VB-6 Message Box Generator Add-In
'By: John P. Cunningham
 RetVal = MsgBox("Please fill in the Item, Cost and Quantity Text Boxes.")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 txtDB(0).SetFocus
'From here program branches to txtDB_LostFocus Sub via the Enter Key
'The txtDB_LostFocus calls the AddNewRecord Subroutine
End Sub
Public Sub AddNewRecord()

   'Open a new connection.
    Set cn1 = New ADODB.Connection
    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    cn1.Open strCnn
   
   Set rs3 = New ADODB.Recordset
   rs3.CursorType = adOpenKeyset
   rs3.LockType = adLockOptimistic
   rs3.Open "tblParts", cn1, , , adCmdTable

   ' Get data from the TextBoxes/user.
    strItem = txtDB(0)
    strCost = txtDB(1)
    strQuantity = txtDB(2)

   

      rs3.AddNew
      rs3!Item = strItem
      rs3!Cost = strCost
      rs3!Quantity = strQuantity
       
      booRecordAdded = True
     rs3.Update
     Call Update_Controls
      
     ' Show the newly added data.
     'VB-6 Message Box Generator Add-In
     'By: John P. Cunningham
      RetVal = MsgBox("New Record - " & rs3!Item & vbCr & _
         "Cost = " & Format(rs3!Cost, "$####.##") & vbCr & _
         "Quantity = " & rs3!Quantity & vbCr & vbCr & _
         "Click on Update Command Button When this Box Closes", 64, "ADO Totals Demo")
         
   'Change the Background Color of the Text Boxes Back
   For ij = 0 To 2
     txtDB(ij).BackColor = &HFFC0C0
   Next
   
   'Update the Database & DataGrid
    cmdUpdate_Click
    
End Sub

Private Sub cmdAddNewRecord_Click()

If txtDB(0) = "" Or txtDB(1) = "" Or txtDB(2) = "" Then
    'VB-6 Message Box Generator Add-In
    'By: John P. Cunningham
    RetVal = MsgBox("You must fill in all of the Text Inputs", 16, "ADO Totals Show")
    Exit Sub
Else
    AddNewRecord
End If

End Sub

Private Sub cmdDelete_Click()

'VB-6 Message Box Generator Add-In
'By: John P. Cunningham
RetVal = MsgBox("Are you sure that you want to Delete this record?", 52, "ADO Totals Show")

Select Case RetVal
     Case 6     'Yes
          rs1.Delete
          rs1.MovePrevious
          Update_Controls
          
     Case 7     'No - do nothing
          
End Select

'Update the TextBoxes to show the record previous to the one just deleted
With rs1
        .MovePrevious
        
      If rs1.BOF Then
        .MoveFirst
         txtDB(0).Text = !Item
         txtDB(1).Text = Format(!Cost, "$####.00")
         txtDB(2).Text = !Quantity
         txtRecordCount = "No. " & rs1.AbsolutePosition & " of " & RcCount & " Records"
         Exit Sub
     End If
      
    End With

End Sub

Private Sub cmdMoveBack_Click()

    With rs1
        .MovePrevious
        
      If rs1.BOF Then
        .MoveFirst
         txtDB(0).Text = !Item
         txtDB(1).Text = Format(!Cost, "$####.00")
         txtDB(2).Text = !Quantity
         txtRecordCount = "No. " & rs1.AbsolutePosition & " of " & RcCount & " Records"
        Exit Sub
        
      Else
      
         txtDB(0).Text = !Item
         txtDB(1).Text = Format(!Cost, "$####.00")
         txtDB(2).Text = !Quantity
         txtRecordCount = "No. " & rs1.AbsolutePosition & " of " & RcCount & " Records"
         
      End If
      
    End With
    
End Sub
Private Sub cmdMoveNext_Click()

  With rs1
        .MoveNext
        
      If rs1.EOF Then
        .MoveFirst
         txtDB(0).Text = !Item
         txtDB(1).Text = Format(!Cost, "$####.00")
         txtDB(2).Text = !Quantity
         txtRecordCount = "No. " & rs1.AbsolutePosition & " of " & RcCount & " Records"
         Exit Sub
        
      Else
      
         txtDB(0).Text = !Item
         txtDB(1).Text = Format(!Cost, "$####.00")
         txtDB(2).Text = !Quantity
         txtRecordCount = "No. " & rs1.AbsolutePosition & " of " & RcCount & " Records"
         
      End If
      
    End With
    
End Sub
Private Sub cmdReportShow_Click()
   
    frmPrinterOptions.Show

End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
'Dim BottomRow, BRowCheck
'This sub would only be used if updates were via the DataGrid
    'BRowCheck = RcCount
    'BottomRow = DataGrid1.Columns(0).CellText(DataGrid1.RowBookmark(DataGrid1.VisibleRows - 1))
    'Label1.Caption = "Records " & _
   BottomRow & " are currently displayed." & RcCount & " Records  " & BottomRow
End Sub

Private Sub DataGrid1_AfterUpdate()

    rs1.Update
    
End Sub

Private Sub DataGrid1_Click()
'When user clicks on a DataGrid Item, show the values in the TextBoxes
With rs1
          
            txtDB(0).Text = !Item
            txtDB(1).Text = Format(!Cost, "$####.00")
            txtDB(2).Text = !Quantity
            txtRecordCount = "No. " & rs1.AbsolutePosition & " of " & RcCount & " Records"
         End With
End Sub

Private Sub DataGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    setTotalBoxes
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'To allow the Enter Key to be used as the Tab Key
'Set the Form's Key Preview Property to "True" then
'set the Form's Controls Tab Order and put the following
'in the Form's KeyPress Subroutine


If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If

End Sub

Private Sub Form_Load()
 
    Call OpenData               'Open and Set the ADO Connection
    Call Update_Controls        'Update the Recordsets
    Call formatDataGrid         'Set the DataGrid's Format
    
     'Move to the First Record & show it in TextBoxes
     With rs1
          .MoveFirst
            txtDB(0).Text = !Item
            txtDB(1).Text = Format(!Cost, "$####.00")
            txtDB(2).Text = !Quantity
     End With
     
     'How many records in the Database
     RcCount = rs1.RecordCount
     'Show the Total Number of Records and the Current Record Number
     txtRecordCount.Text = "No. " & rs1.AbsolutePosition & " of " & RcCount & "  Records"
     
End Sub

Private Sub cmdExit_Click()
    frmAbout.Show
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    rs1.Update
    rs2.Update
    Update_Controls
    cmdMoveBack_Click
End Sub

Private Sub Form_Resize()
    Dim iMaxHeight As Integer
  
    If Me.Width < 5600 Then Me.Width = 5600
    If Me.Height < 3000 Then Me.Height = 3000
    
    iMaxHeight = Me.Height - Frame1.Height - txtTotalQty.Height - 800
    DataGrid1.Move 0, 100, Me.ScaleWidth, iMaxHeight
    
    With Frame1
        .Top = Me.ScaleHeight - Frame1.Height
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
    End With
        
    setTotalBoxes
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
    Set rs1 = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
    Set cn = Nothing
    Set cn1 = Nothing
Unload Me
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
'Highlight each TextBox when it has the Focus
For ij = 0 To 2
 With txtDB(ij)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   Next
End Sub

Private Sub txtDB_LostFocus(Index As Integer)
'When the Focus is off the last TextBox execute the AddNewRecord Subroutine

    'Check to see that all of the TextBoxes are filled in
    If Index = 2 Then
        If txtDB(0) = "" Or txtDB(1) = "" Or txtDB(2) = "" Then
        
           'VB-6 Message Box Generator Add-In
           'By: John P. Cunningham
            RetVal = MsgBox("You must fill in all of the Text Inputs", 16, "ADO Totals Show")
            Exit Sub
        Else
            'When the Focus is off the last TextBox execute the AddNewRecord Subroutine
            AddNewRecord
        End If
    End If
    
End Sub
