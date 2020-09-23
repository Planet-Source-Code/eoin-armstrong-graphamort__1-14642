VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mortgage Calculator"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optGraph 
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optGraph 
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3975
      Left            =   4200
      OleObjectBlob   =   "frmMain.frx":0000
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdChart 
      Caption         =   "View &Chart"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdAmortise 
      Caption         =   "&Amortisation Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtResult 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Text            =   "0"
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Monthly Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1980
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyy"
      Format          =   24444931
      CurrentDate     =   36908
   End
   Begin VB.TextBox txtTerm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1380
      Width           =   735
   End
   Begin VB.TextBox txtInterest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   780
      Width           =   735
   End
   Begin VB.TextBox txtLoan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   11
      Top             =   2647
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Date of First Payment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Term (Whole Years):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Interest Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Loan Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iPayments As Integer ' iPayments = no. of MONTHLY payments payable

Private Sub cmdAmortise_Click()

'This sub populates the MSFlexgrid

    On Error Resume Next ' to compensate for the error sometimes given in the cell colour changing routine
    Dim sCurrMonth As String 'for the Date column
    Dim iDay, iMonth, iYear, iCountPay As Integer ' 1st 3 are for the date column, 4th is the loop counter
    Dim dCurrBal, dCurrInt, dCurrPrin, dEndBal, dCumInt, dTotPaid As Double 'Period's starting loan balance, interest this period,
                                                                            'principal paid this period, loan balance at end of period,
                                                                            'cumulative interest at end of period, total paid to date at end of period
                                                                            '; respectively!
    Dim sWhen As String 'the commencement date
    
    frmMain.MousePointer = vbHourglass ' cpu busy
    sWhen = dtpStart ' assign commencement date
    iDay = Val(Left(sWhen, 2)) ' the day
    iMonth = Val(Mid(sWhen, 4, 2)) ' the month
    iYear = Val(Right(sWhen, 2)) ' the year - latter 2 are required to increment period
    dTotPaid = txtResult.Text ' assign monthly payment
    dCurrBal = txtLoan.Text ' assign loan balance for 1st period
    dCurrInt = dCurrBal / 12 * (txtInterest / 100) ' assign interest for the first period
    dCurrPrin = dTotPaid - dCurrInt ' assign principal paid for the first period
    dEndBal = dCurrBal - dCurrPrin ' assign balance at end of the first period
    dCumInt = dCurrInt ' assign cumulative interest for the first period
    dTotPaid = txtResult.Text ' assign running total paid for first period
    
    With MSFlexGrid1
        .Rows = iPayments + 1 ' number of rows in the FlexGrid. + 1, to cater for the header row
        .TextMatrix(1, 2) = Format(dCurrBal, "Currency", 2)  ' print current balance,
        .TextMatrix(1, 3) = Format(dCurrInt, "Currency", 2)  ' current interest,
        .TextMatrix(1, 4) = Format(dCurrPrin, "Currency", 2) ' current principal,
        .TextMatrix(1, 5) = Format(dEndBal, "Currency", 2)   ' end balance,
        .TextMatrix(1, 6) = Format(dCumInt, "Currency", 2)   ' cumulative interest, and
        .TextMatrix(1, 7) = Format(dTotPaid, "Currency", 2)  ' running total paid
    End With                                                 ' for the first period.
                                                             ' Remeber that the FlexGrid row and column indices both begin at 0.
                                                             ' So the current balance will print in the 3rd column from the left,
                                                             ' the current interest will appear in the 4th, and so on.
    ' the all-important loop
    For iCountPay = 1 To iPayments ' 1 to the number of rows
        With MSFlexGrid1
            .TextMatrix(iCountPay, 0) = iCountPay ' populates the No. of payments column i.e. the first column on the left
            If iMonth > 12 Then   ' If we're past December, we revert
                iMonth = 1        ' back to January,
                iYear = iYear + 1 ' and add a year.
            End If
            .TextMatrix(iCountPay, 1) = Format(Val(iDay), "00") & _
            "/" & Format(Val(iMonth), "00") & "/" & Format(Val(iYear), "00")
            ' The above split line populates the 2nd column with the payment dates.
            ' I get the feeling I could have done this more simply...
            ' Basically, I converted the bits of the month back to little strings,
            ' and concatenated them with slashes to form 1 big, lovely string.
            iMonth = iMonth + 1 ' increment the month
            If iCountPay > 1 Then ' the loop for the financial figures really begins
                                  ' from the 3rd row (1st row=headers, 2nd was the first period,
                                  ' and was already populated above.
                
                ' starting balance for this period is read from ending balance from previous period
                dCurrBal = .TextMatrix(iCountPay - 1, 5)
                ' print the starting balance for this period
                .TextMatrix(iCountPay, 2) = Format(dCurrBal, "Currency", 2)
                ' multiply the starting balance by the interst rate (which is divided by 100 to provide a percentage)
                ' and divide by 12 to get the interest paid for this month.
                dCurrInt = .TextMatrix(iCountPay, 2) / 12 * (txtInterest / 100)
                ' print the interest paid this period
                .TextMatrix(iCountPay, 3) = Format(dCurrInt, "Currency", 2)
                ' this periods principal paid is the monthly total minus the interest paid
                dCurrPrin = txtResult - dCurrInt
                ' print the principal paid for this period
                .TextMatrix(iCountPay, 4) = Format(dCurrPrin, "Currency", 2)
                ' this periods end balance is the start balance minus principal paid this period
                dEndBal = dCurrBal - dCurrPrin
                ' print the end balance for this period
                .TextMatrix(iCountPay, 5) = Format(dEndBal, "Currency", 2)
                ' cumulative interest for this period is the cumulative interst paid in the last period
                ' plus interest paid this period
                dCumInt = .TextMatrix(iCountPay - 1, 6) + dCurrInt
                ' print the cumulative interest for this period
                .TextMatrix(iCountPay, 6) = Format(dCumInt, "Currency", 2)
                ' the total paid to date this period is simply the monthly payment
                ' multiplied by the payment number
                dTotPaid = .TextMatrix(1, 7) * iCountPay
                ' print the total paid at the end of this period
                .TextMatrix(iCountPay, 7) = Format(dTotPaid, "Currency", 2)
            End If
        End With
    Next iCountPay ' increment the all-important loop
    
    With MSFlexGrid1
        .ColWidth(0) = 640   ' Adjust the column widths
        .ColWidth(1) = 1220  ' in 'twips'.
        .ColWidth(2) = 1700
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 1700
        .ColWidth(6) = 1700
        .ColWidth(7) = 1700
        Dim i
        For i = 0 To 7           ' All columns
            .ColAlignment(i) = 3 ' to be centrally aligned
        Next i
    End With
    
'    Dim iCols As Integer
'    Do Until MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
'        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'        For C = 0 To MSFlexGrid1.Cols - 1
'            MSFlexGrid1.Col = C
'            MSFlexGrid1.CellBackColor = &HC0FFC0
'        Next
'        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'    Loop

    ' This section colours alternate rows, making the
    ' grid a little easier to read.
    Dim iCols As Integer
     ' for each row (select by row first)
    Do Until MSFlexGrid1.Row = iPayments
        ' skip the header row & select the next.
        ' start the first row white,
        ' and increment
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        ' for all columns (remember that the Columns index
        ' begins with 0, but that the .Col property starts at 1
        ' hence the - 1
        For iCols = 0 To 7 'MSFlexGrid1.Cols - 1
            ' select each column in turn, after having already
            ' selected the row above
            MSFlexGrid1.Col = iCols
            ' change the cell colour (light green, in this case)
            MSFlexGrid1.CellBackColor = &HC0FFC0
        Next iCols
        ' next column, please
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    Loop
    
    MSFlexGrid1.Visible = True      ' show the grid
    cmdChart.Enabled = True         ' show the chart button
    frmMain.MousePointer = vbNormal ' revert mouse pointer
End Sub

Private Sub cmdCalculate_Click()
'This calculates the monthly payment
    
    On Error GoTo BadData ' If you forget to fill in a field
    Dim dRate, dLoan, dMonthlyPayment As Double ' The rate, the principal, the eventual result
    
    ' This is the monthly interest rate. The interest rate per
    ' period is required for VB's own 'Pmt' function.
    ' It's divided by 1200 (a) to get the monthly rate: 12; and
    ' (b) to get a percentage: 100
    dRate = txtInterest.Text / 1200
    ' assign the principal
    dLoan = txtLoan.Text
    ' calculate the ubiquitous iPayments variable, which
    ' is required in the 'Pmt' function, and helps determine
    ' the number of rows within the FlexGrid
    iPayments = txtTerm * 12
    ' The 'Pmt' function calculates the monthly payment
    dMonthlyPayment = Pmt(dRate, iPayments, -dLoan)
    ' assign the result to the relevant text box and format
    ' it as currency to 2 decimal places
    txtResult.Text = Format(dMonthlyPayment, "Currency", 2)
    ' show the print amortisation schedule button
    cmdAmortise.Enabled = True
    ' exit the sub, so the program doesn't continue on to
    ' the 'BadData' error-handling section
    Exit Sub

' What can I say? I'm a lazy swine. I've used the routine for
' handling all errors...
BadData:
    MsgBox "You have input your data incorrectly!", vbOKOnly + vbCritical, "Error"
    cmdAmortise.Enabled = False
    cmdChart.Enabled = False
    txtLoan.SetFocus
End Sub

Private Sub cmdChart_Click()
' this plots and shows the chart

    ' Charts can be thought of as graphic representations
    ' of a matrix of figures. And generally they are handled
    ' this way. A chart's data is based on a grid. You enter
    ' data in rows and columns, thereby affecting the chart's
    ' appearance.
    Dim iRowCount, iColumnCount As Integer ' the chart's rows and columns
    Dim dCurrData As Double ' I'll assign a chart 'cells' data with this variable
    'Dim sDateLabel As String
    
    frmMain.MousePointer = vbHourglass ' cpu busy
    ' This ensures that the chart will not redraw itself
    ' every time data is placed within a cell using the
    ' loop below.
    MSChart1.Repaint = False
    MSChart1.RowCount = iPayments ' See? Just like the FlexGrid (well alomst)
    
    ' You can either select columns or rows to begin this
    ' little nested loop. I chose columns in this case.
    ' I don't know why...
    ' Beginning first loop
    ' Only 2 columns needed - each periods interest & principal payments
    For iColumnCount = 1 To 2
        ' You must also plot each row
        For iRowCount = 1 To iPayments
            ' assign a FlexGrid cell's data to dCurrData
            ' I use iColumnCount + 2, because the first loop
            ' will seek for data in column number 4 (interst),
            ' then column 5 (principal) - again remember that
            ' a FlexGrid's column and row indices begin at 0,
            ' so you have to compensate...
            dCurrData = MSFlexGrid1.TextMatrix(iRowCount, iColumnCount + 2)
            ' MSChart's SetData method allocates dCurrData to
            ' the chart's current cell
            MSChart1.DataGrid.SetData iRowCount, iColumnCount, dCurrData, False
        Next iRowCount
    Next iColumnCount
    
    ' These are the series' legends
    MSChart1.Plot.SeriesCollection(1) = "Interest"
    MSChart1.Plot.SeriesCollection(2) = "Principal"
    
    ' show the chart
    MSChart1.Visible = True
    ' repaint the chart
    MSChart1.Repaint = True
    ' show chart options
    Frame1.Visible = True
    ' clicking the chart will now not produce any ugle lines:
    MSChart1.Enabled = False
    frmMain.MousePointer = vbNormal 'cpu finished
End Sub

Private Sub Form_Load()
    ' the FlexGrid's column headers...
    With MSFlexGrid1
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Begin Bal."
        .TextMatrix(0, 3) = "Interest"
        .TextMatrix(0, 4) = "Principal"
        .TextMatrix(0, 5) = "End Bal."
        .TextMatrix(0, 6) = "Cum. Interest"
        .TextMatrix(0, 7) = "Total Paid"
    End With
    ' assign today's date to the DateTimePicker control
    dtpStart = Date
    ' start off with an Area type graph
    optGraph(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub

Private Sub optGraph_Click(Index As Integer)
' sub to switch graph types
    If Index = 0 Then
        MSChart1.chartType = VtChChartType2dArea
        ' one series appears on top of the other
        MSChart1.Stacking = True
        ' redraw the chart
        MSChart1.Repaint = True
    Else
        MSChart1.chartType = VtChChartType2dLine
        ' series shown apart - necessary for a meaningful line graph
        MSChart1.Stacking = False
        ' redraw the chart
        MSChart1.Repaint = True
    End If
End Sub

Private Sub txtInterest_Change()
'altering any inputs will reset results
    txtResult.Text = 0
    cmdAmortise.Enabled = False
    cmdChart.Enabled = False
    MSFlexGrid1.Visible = False
    MSChart1.Visible = False
End Sub

Private Sub txtLoan_Change()
'altering any inputs will reset results
    txtResult.Text = 0
    cmdAmortise.Enabled = False
    cmdChart.Enabled = False
    MSFlexGrid1.Visible = False
    MSChart1.Visible = False
End Sub

Private Sub txtTerm_Change()
'altering any inputs will reset results
    txtResult.Text = 0
    cmdAmortise.Enabled = False
    cmdChart.Enabled = False
    MSFlexGrid1.Visible = False
    MSChart1.Visible = False
End Sub
