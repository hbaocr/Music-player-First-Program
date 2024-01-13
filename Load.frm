VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   LinkTopic       =   "Form2"
   Picture         =   "Load.frx":0000
   ScaleHeight     =   2730
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub Form_DblClick()
    Unload Me

End Sub

Private Sub Form_Load()
    SetWindow Me.hwnd, Form2.BackColor, 0, LWA_COLORKEY
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
