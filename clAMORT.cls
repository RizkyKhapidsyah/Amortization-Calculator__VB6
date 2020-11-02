VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clAMORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------
' Class: clAMORT
' Created by Brad Gosdin & Rizky Khapidsyah
' Date: 07-10-2001    Time: 21:20
'
' Purpose: Class to handle Amortization Schedules
'---------------------------------------------
Option Explicit

Private mvarRate() As Double
Private mvarPayment() As Double
Private mvarTerm() As Long
Private mvarPrincipalPayment() As Double
Private mvarInterestPayment() As Double
Private mvarLoanAmount() As Double
Private mvarTotalInterestPaid() As Double
Private mvarTotalPrincipalPaid() As Double
Private mvarTotalExtraPrincipal() As Double
Private mvarSavedInterest As Double
Private mvarCount As Long

Event CurrentCount(ByVal Count As Long)

'---------------------------------------------
' Function MonthlyPayment
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:20
'
' Purpose:
'
'    dAMOUNT:
'    dRATE:
'    dYEARS:
'    dPERIODS:
'    Return value:
'
'---------------------------------------------
Private Function MonthlyPayment(dAMOUNT As Double, dRATE As Double, dYEARS As Double, dPERIODS As Double) As Double
    Dim Interest As Double
    
    On Error GoTo proc_err
    
    Interest = (dRATE / 100) / dPERIODS
    MonthlyPayment = dAMOUNT * ((Interest * (1 + Interest) ^ (dYEARS * dPERIODS)) / ((1 + Interest) ^ (dYEARS * dPERIODS) - 1))

proc_exit:
    Exit Function
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "MonthlyPayment"
    GoTo proc_exit
    
End Function

'---------------------------------------------
' Sub Calculate
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:21
'
' Purpose:
'
'    Rate:
'    Years:
'    Periods:
'    LoanAmount:
'    ExtraPrincipal:
'    Escrow:
'    IMax:
'
'---------------------------------------------
Sub Calculate(Rate As Double, Years As Long, Periods As Long, LoanAmount As Double, ExtraPrincipal As Double, Escrow As Double, IMax As Long)
    Dim dInterest As Double
    Dim dPayment As Double
    Dim dPrincipal As Double
    Dim dInterestPayment As Double
    Dim dLoanAmount As Double
    Dim dExtraPrincipal As Double
    Dim dTotalInterest As Double
    Dim dTotalPrincipal As Double
    Dim dTotalExtra As Double
    Dim dEscrow As Double
    Dim dSavedInterest As Double
    Dim dNumberPeriods As Long
    Dim dOuterIndex As Long
    Dim dInnerIndex As Long
    
    On Error GoTo proc_err
    
    dNumberPeriods = (Years * Periods)
    mvarCount = dNumberPeriods
    
    ReDim mvarRate(dNumberPeriods)
    ReDim mvarPayment(dNumberPeriods)
    ReDim mvarTerm(dNumberPeriods)
    ReDim mvarPrincipalPayment(dNumberPeriods)
    ReDim mvarInterestPayment(dNumberPeriods)
    ReDim mvarLoanAmount(dNumberPeriods)
    ReDim mvarTotalInterestPaid(dNumberPeriods)
    ReDim mvarTotalPrincipalPaid(dNumberPeriods)
    ReDim mvarTotalExtraPrincipal(dNumberPeriods)
    
    dLoanAmount = LoanAmount
    dInterest = MonthlyInterest(Rate, CDbl(Periods))
    dPayment = MonthlyPayment(LoanAmount, Rate, CDbl(Years), CDbl(Periods))
    dEscrow = Escrow / Periods
    dExtraPrincipal = ExtraPrincipal
    
    For dOuterIndex = 1 To dNumberPeriods
        dInterestPayment = dLoanAmount * dInterest
        dPrincipal = dPayment - (dLoanAmount * dInterest)
        dTotalInterest = dTotalInterest + (dLoanAmount * dInterest)
        dLoanAmount = dLoanAmount - dPrincipal
    Next dOuterIndex
    
    dSavedInterest = dTotalInterest
    dPrincipal = 0
    dInterestPayment = 0
    dTotalInterest = 0
    dOuterIndex = 0
    dLoanAmount = LoanAmount
    
    For dOuterIndex = 1 To dNumberPeriods
        
        RaiseEvent CurrentCount(dOuterIndex)
        
        dInterestPayment = dLoanAmount * dInterest
        dTotalExtra = dTotalExtra + dExtraPrincipal
        dPrincipal = dPayment + dExtraPrincipal - (dLoanAmount * dInterest)
        dTotalInterest = dTotalInterest + (dLoanAmount * dInterest)
        dLoanAmount = dLoanAmount - dPrincipal
        dTotalPrincipal = dTotalPrincipal + dPrincipal
        
        mvarRate(dOuterIndex) = dInterest
        mvarPayment(dOuterIndex) = dPayment + dEscrow
        mvarTerm(dOuterIndex) = dNumberPeriods
        mvarPrincipalPayment(dOuterIndex) = dPrincipal
        mvarInterestPayment(dOuterIndex) = dInterestPayment
        mvarLoanAmount(dOuterIndex) = dLoanAmount
        mvarTotalExtraPrincipal(dOuterIndex) = dTotalExtra
        mvarTotalPrincipalPaid(dOuterIndex) = dTotalPrincipal
        mvarTotalInterestPaid(dOuterIndex) = dTotalInterest
        
        If (dLoanAmount < 0) Or (dOuterIndex = dNumberPeriods) Then
            mvarLoanAmount(dOuterIndex) = 0
            mvarPrincipalPayment(dOuterIndex) = dPrincipal + dLoanAmount
            mvarSavedInterest = dSavedInterest - dTotalInterest
            
            For dInnerIndex = dOuterIndex To dNumberPeriods
                RaiseEvent CurrentCount(dInnerIndex)
            Next dInnerIndex
            
            IMax = dOuterIndex
            Exit For
        End If
        
    Next dOuterIndex
    
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "Calculate"
    GoTo proc_exit
    
End Sub

'---------------------------------------------
' Function MonthlyInterest
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:21
'
' Purpose:
'
'    dRATE:
'    dPERIODS:
'    Return value:
'
'---------------------------------------------
Private Function MonthlyInterest(dRATE As Double, dPERIODS As Double) As Double
    
    On Error GoTo proc_err
    
    MonthlyInterest = dRATE / (dPERIODS * 100)
proc_exit:
    Exit Function
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "MonthlyInterest"
    GoTo proc_exit
    
End Function

'---------------------------------------------
' Sub Clear
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:21
'
' Purpose:
'
'
'---------------------------------------------
Sub Clear()
    
    On Error GoTo proc_err
    
    Erase mvarInterestPayment
    Erase mvarPrincipalPayment
    Erase mvarTerm
    Erase mvarPayment
    Erase mvarInterestPayment
    Erase mvarLoanAmount
    Erase mvarTotalInterestPaid
    Erase mvarTotalPrincipalPaid
    Erase mvarTotalExtraPrincipal
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "Clear"
    GoTo proc_exit
    
End Sub

'---------------------------------------------
' Function Count
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:25
'
' Purpose:
'
'    Return value:
'
'---------------------------------------------
Function Count() As Long
    Count = mvarCount
End Function

'---------------------------------------------
' Function GetTotalPrincipalPaid
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:26
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetTotalPrincipalPaid(Index) As Double
    GetTotalPrincipalPaid = mvarTotalPrincipalPaid(Index)
End Function

'---------------------------------------------
' Function GetTotalExtraPrincipal
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:26
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetTotalExtraPrincipal(Index) As Double
    GetTotalExtraPrincipal = mvarTotalExtraPrincipal(Index)
End Function

'---------------------------------------------
' Function GetTotalInterestPaid
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:26
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetTotalInterestPaid(Index) As Double
    GetTotalInterestPaid = mvarTotalInterestPaid(Index)
End Function

'---------------------------------------------
' Function GetInterestPayment
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:27
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetInterestPayment(Index) As Double
    GetInterestPayment = mvarInterestPayment(Index)
End Function

'---------------------------------------------
' Function GetPrincipalPayment
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:27
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetPrincipalPayment(Index) As Double
    GetPrincipalPayment = mvarPrincipalPayment(Index)
End Function

'---------------------------------------------
' Function GetTerm
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:27
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetTerm(Index) As Long
    GetTerm = mvarTerm(Index)
End Function

'---------------------------------------------
' Function GetPayment
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:28
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetPayment(Index) As Double
    GetPayment = mvarPayment(Index)
End Function

'---------------------------------------------
' Function GetInterestRate
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:28
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetInterestRate(Index) As Double
    GetInterestRate = mvarRate(Index)
End Function

'---------------------------------------------
' Function GetLoanAmount
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:28
'
' Purpose:
'
'    Index:
'    Return value:
'
'---------------------------------------------
Function GetLoanAmount(Index) As Double
    GetLoanAmount = mvarLoanAmount(Index)
End Function

'---------------------------------------------
' Function GetSavedInterest
'
' Created by Brad Gosdin
' Date: 07-10-2001    Time: 21:28
'
' Purpose:
'
'    Return value:
'
'---------------------------------------------
Function GetSavedInterest() As Double
    GetSavedInterest = mvarSavedInterest
End Function
