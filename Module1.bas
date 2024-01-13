Attribute VB_Name = "Formtrongsuot"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1


'--------------------------------------------
Public Const GWL_EXSTYLE = (-20)

Public Declare Function SetWindowLong Lib "user32" _
       Alias "SetWindowLongA" (ByVal hWnd As Long, _
       ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong Lib "user32" _
       Alias "GetWindowLongA" (ByVal hWnd As Long, _
       ByVal nIndex As Long) As Long
       
'--------------------------------------------------

Public Declare Function SetLayeredWindowAttributes Lib "user32" _
       (ByVal hWnd As Long, ByVal crKey As Long, _
       ByVal bAlpha As Integer, ByVal dwFlags As Long) As Long

Public Const WS_EX_LAYERED = &H80000
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2


Public Function SetWindow(hWnd As Long, crKey As Long, _
                bAlpha As Integer, dwFlags As Long) As Long
Dim ExStyle As Long
Dim i As Integer
Dim result As Long
    'thay doi ExStyle cua form
    ExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    ExStyle = ExStyle Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, ExStyle
    
    result = SetLayeredWindowAttributes(hWnd, crKey, bAlpha, dwFlags)
    SetWindow = result
End Function

