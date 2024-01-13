VERSION 5.00
Begin VB.Form song 
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
Attribute VB_Name = "song"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim POI As Boolean
Dim chuoi, chuoi3
Private Sub Command1_Click()
Dim gh, KJ
Open Ds(Baiso) For Input As #1
Do While Not EOF(1)
Line Input #1, gh
XOAT (gh)
Text1.Text = Text1.Text + Chr(13) + Chr(10) + chuoi3
Loop
Close #1
End Sub
Function XOAT(ad As String)
Dim WER, WER1
Dim CHUOI2
chuoi = ad

For i = 1 To Len(chuoi)
WER = Left(chuoi, i)
WER1 = Right(WER, 1)
If WER1 = " " Then
chuoi = Right(chuoi, Len(chuoi) - Len(WER) + 1)
          If Len(Trim(WER)) < 15 Then
          CHUOI2 = CHUOI2 + WER
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
