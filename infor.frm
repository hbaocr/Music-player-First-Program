VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Song"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   5160
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim POI As Boolean
Dim CHUOI2, chuoi3
Private Sub Command1_Click()
Dim gh, KJ
Open Ds(Baiso) For Input As #1
Do While Not EOF(1)
Line Input #1, gh
KJ = gh
XOAT (KJ)
Text1.Text = Text1.Text + Chr(13) + Chr(10) + chuoi3
Loop
Close #1
End Sub
Function XOAT(chuoi As String)
Dim OI, UI, JKL1, WER, WER1, WER2
Dim i, P
Dim CHUOI1
CHUOI1 = ""
For i = Len(chuoi) To 1 Step -1
OI = Right(chuoi, i)
UI = UCase(Left(OI, 1))
'If UI = " " Then
'JKL1 = I
Select Case UI
Case "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "A", "S", "D", "F", "G", "H", "J", "K", "L", "Z", "X", "C", "V", "B", "N", "M", " ", "'"
CHUOI1 = CHUOI1 + Left(OI, 1)
End Select
Next i
CHUOI2 = CHUOI1
XOAT1 (CHUOI2)
End Function
Function XOAT1(ad As String)
Dim WER, WER1
Dim CHUOI2, chuoi
Dim das
chuoi = ad
For i = 1 To Len(chuoi)
das = das + 1
WER = Left(LTrim(chuoi), das)
WER1 = Right(WER, 1)
If WER1 = " " Then
das = 0
chuoi = Right(chuoi, Len(chuoi) - Len(WER))
          If Len(Trim(WER)) < 11 Then
          CHUOI2 = CHUOI2 + LTrim(WER)
         End If
End If
Next i
chuoi3 = CHUOI2 + chuoi
End Function


Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Command1_Click
Timer1.Enabled = False
End Sub
