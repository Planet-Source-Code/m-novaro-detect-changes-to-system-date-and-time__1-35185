VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateTimeChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Const WM_TIMECHANGE = &H1E
Private Const HWND_TOPMOST = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Event DateTimeChanged(ByVal newDate As Date, ByVal newTime As Date)

Private m_bInitialized As Boolean
Private m_hWnd As Long

Private Sub Class_Initialize()
    m_bInitialized = False
End Sub

Public Sub StartCaptureChange(hwnd As Long)
    AttachMessage Me, hwnd, WM_TIMECHANGE
    m_hWnd = hwnd
    m_bInitialized = True
End Sub

Private Sub Class_Terminate()
    If m_bInitialized Then
        DetachMessage Me, m_hWnd, WM_TIMECHANGE
    End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    '
    ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case iMsg
    
        Case WM_TIMECHANGE
            RaiseEvent DateTimeChanged(Date, Time)

        Case Else
            '

    End Select

End Function

Public Sub ShowChangeDateTime()
   
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,0", vbNormalFocus)

End Sub

Public Sub SendChangeNotification()

    Call SendMessage(HWND_TOPMOST, WM_TIMECHANGE, 0, ByVal 0)

End Sub
