VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amortization Calculator"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8340
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   5160
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid mflGrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1920
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   5741
      _Version        =   393216
      GridColor       =   0
      Redraw          =   -1  'True
      GridLines       =   0
   End
   Begin VB.Frame fraInformation 
      Caption         =   "Amortization Schedule"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8130
      Begin MSMask.MaskEdBox mdtExtra 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mdtEscrow 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "&Calculate"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mdtPeriods 
         Height          =   285
         Left            =   4320
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mdtTerm 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mdtRate 
         Height          =   285
         Left            =   6360
         TabIndex        =   4
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mdtLoanAmount 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSIM 
         Caption         =   "Interest Saved"
         Height          =   245
         Left            =   4320
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblSI 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4320
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblExtra 
         Caption         =   "Extra Principal"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblEscrow 
         Caption         =   "Escrow"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Loan Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblInterestRate 
         Caption         =   "APR"
         Height          =   255
         Left            =   6360
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTerm 
         Caption         =   "Term"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNper 
         Caption         =   "Periods"
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   5190
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9525
            TextSave        =   ""
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "9:51 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "7/10/2001"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents X As clAMORT
Attribute X.VB_VarHelpID = -1

'---------------------------------------------
' Sub cmdCalculate_Click
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:42
'
' Purpose:
'
'
'---------------------------------------------
Private Sub cmdCalculate_Click()
    Dim Years As Long
    Dim z As New clPBPANEL
    
    On Error GoTo proc_err
    
    Set X = New clAMORT
    Dim Index As Long
    Dim Max As Long
    z.ShowProgressInStatusBar True, SB.hWnd, PB, SB, Screen, 1
    X.Clear
    Years = CLng(mdtTerm)
    Max = (CLng(mdtTerm) * CLng(mdtPeriods))
    PB.Max = Max
    X.Calculate CDbl(mdtRate), CLng(mdtTerm), CLng(mdtPeriods), CDbl(mdtLoanAmount), CDbl(mdtExtra), CDbl(mdtEscrow), Max
    mflGrid.Redraw = False
    mflGrid.Visible = False
    Call InitGrid(Max)
    For Index = 1 To Max
        Call CalcRow(X.GetInterestRate(Index), X.GetLoanAmount(Index), X.GetPayment(Index), X.GetPrincipalPayment(Index), X.GetInterestPayment(Index), X.GetTotalInterestPaid(Index), X.GetTotalPrincipalPaid(Index), X.GetTotalExtraPrincipal(Index), Index)
    Next Index
    lblSI.Caption = FormatCurrency(X.GetSavedInterest, 2)
    mflGrid.Visible = True
    mflGrid.Redraw = True
    z.ShowProgressInStatusBar False, SB.hWnd, PB, SB, Screen, 1
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
    
End Sub

'---------------------------------------------
' Sub Form_Load
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:43
'
' Purpose:
'
'
'---------------------------------------------
Private Sub Form_Load()
    On Error GoTo proc_err
    
    Call InitGrid(0)
    PB.Visible = False
    mdtEscrow = 0
    mdtExtra = 0
    mdtTerm = 0
    mdtPeriods = 0
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Function IsEven
'
' Created by Unknown
'
' Purpose:
'
'    iNUM:
'    Return value:
'
'---------------------------------------------
Function IsEven(iNUM As Long) As Boolean
    On Error GoTo proc_err
    
    IsEven = iNUM Mod 2 = 0
proc_exit:
    Exit Function
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Function

'---------------------------------------------
' Sub InitGrid
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'    Max:
'
'---------------------------------------------
Sub InitGrid(Max As Long)
    On Error GoTo proc_err
    
    mflGrid.Cols = 9
    mflGrid.Rows = (1 + Max)
    mflGrid.ColWidth(0) = 400
    mflGrid.Col = 0
    mflGrid.Text = "#"
    mflGrid.ColWidth(1) = 1250
    mflGrid.Col = 1
    mflGrid.Text = "Interest Rate"
    mflGrid.ColWidth(2) = 1250
    mflGrid.Col = 2
    mflGrid.Text = "Loan Amount"
    mflGrid.ColWidth(3) = 1250
    mflGrid.Col = 3
    mflGrid.Text = "Payment"
    mflGrid.ColWidth(4) = 1250
    mflGrid.Col = 4
    mflGrid.Text = "Applied Principal"
    mflGrid.ColWidth(5) = 1250
    mflGrid.Col = 5
    mflGrid.Text = "Applied Interest"
    mflGrid.ColWidth(6) = 1250
    mflGrid.Col = 6
    mflGrid.Text = "Total Interest"
    mflGrid.ColWidth(7) = 1250
    mflGrid.Col = 7
    mflGrid.Text = "Total Principal"
    mflGrid.ColWidth(8) = 1250
    mflGrid.Col = 8
    mflGrid.Text = "Total Extra"
    mflGrid.AllowBigSelection = False
    mflGrid.AllowUserResizing = flexResizeColumns
    mflGrid.SelectionMode = flexSelectionByRow
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
    
End Sub

'---------------------------------------------
' Sub CalcRow
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'    I:
'    A:
'    P:
'    AP:
'    AI:
'    TI:
'    TP:
'    TE:
'    Index:
'
'---------------------------------------------
Sub CalcRow(I As Double, A As Double, P As Double, AP As Double, AI As Double, TI As Double, TP As Double, TE As Double, Index As Long)
    On Error GoTo proc_err
    
    mflGrid.Row = Index
    mflGrid.Col = 0
    mflGrid.Text = CStr(Index)
    mflGrid.Col = 1
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(I, "#0.000000")
    mflGrid.Col = 2
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(A, "$###,###,###.00")
    mflGrid.Col = 3
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(P, "$###,###,###.00")
    mflGrid.Col = 4
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(AP, "$###,###,###.00")
    mflGrid.Col = 5
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(AI, "$###,###,###.00")
    mflGrid.Col = 6
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(TI, "$###,###,###.00")
    mflGrid.Col = 7
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(TP, "$###,###,###.00")
    mflGrid.Col = 8
    If IsEven(Index) Then
        mflGrid.CellBackColor = &HC0FFC0
    End If
    mflGrid.CellAlignment = flexAlignRightCenter
    mflGrid.Text = Format(TE, "$###,###,###.00")
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub SelectAllText
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'    Text:
'
'---------------------------------------------
Sub SelectAllText(Text As Object)
    On Error GoTo proc_err
    
    Text.SelStart = 0
    Text.SelLength = Len(Text.Text)
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub mdtEscrow_GotFocus
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'
'---------------------------------------------
Private Sub mdtEscrow_GotFocus()
    On Error GoTo proc_err
    
    SelectAllText mdtEscrow
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub mdtExtra_GotFocus
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:45
'
' Purpose:
'
'
'---------------------------------------------
Private Sub mdtExtra_GotFocus()
    On Error GoTo proc_err
    
    SelectAllText mdtExtra
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub mdtLoanAmount_GotFocus
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:45
'
' Purpose:
'
'
'---------------------------------------------
Private Sub mdtLoanAmount_GotFocus()
    On Error GoTo proc_err
    
    SelectAllText mdtLoanAmount
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
    
End Sub

'---------------------------------------------
' Sub mdtPeriods_GotFocus
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:45
'
' Purpose:
'
'
'---------------------------------------------
Private Sub mdtPeriods_GotFocus()
    On Error GoTo proc_err
    
    SelectAllText mdtPeriods
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub mdtRate_GotFocus
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'
'---------------------------------------------
Private Sub mdtRate_GotFocus()
    On Error GoTo proc_err
    
    SelectAllText mdtRate
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub mdtTerm_GotFocus
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'
'---------------------------------------------
Private Sub mdtTerm_GotFocus()
    On Error GoTo proc_err
    
    SelectAllText mdtTerm
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub

'---------------------------------------------
' Sub x_CurrentCount
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:44
'
' Purpose:
'
'    Count:
'
'---------------------------------------------
Private Sub x_CurrentCount(ByVal Count As Long)
    On Error GoTo proc_err
    
    PB.Value = Count
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
End Sub
