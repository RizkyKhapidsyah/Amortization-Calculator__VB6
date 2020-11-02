VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clPBPANEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

'---------------------------------------------
' Sub ShowProgressInStatusBar
'
' Created by Unknown
'
' Purpose: Place the progress bar into the status bar
'
'    bShowProgressBar:
'    hWnd:
'    PB:
'    SB:
'    Screen:
'    Panel:
'
'---------------------------------------------
Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean, hWnd As Long, PB As Control, SB As Control, Screen As Object, Panel As Long)
    Dim tRC As RECT
    
    On Error GoTo proc_err
    
    If bShowProgressBar Then
        SendMessageAny SB.hWnd, SB_GETRECT, (Panel - 1), tRC
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
        With PB
            SetParent .hWnd, SB.hWnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
    Else
        SetParent PB.hWnd, hWnd
        PB.Visible = False
    End If
proc_exit:
    Exit Sub
    
proc_err:
    MsgBox Err.Number & " " & Err.Description & " An Error Occurred In Class", vbCritical, "ShowProgressInStatusBar"
    GoTo proc_exit
    
End Sub
