VERSION 5.00
Begin VB.Form Save1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save lists"
   ClientHeight    =   2955
   ClientLeft      =   3135
   ClientTop       =   4140
   ClientWidth     =   4620
   Icon            =   "Save1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4620
   Begin VB.CommandButton MDSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cancel"
      Height          =   315
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   840
   End
   Begin VB.CommandButton MDSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Save"
      Height          =   315
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   840
   End
   Begin VB.TextBox ten 
      BackColor       =   &H00FFC0C0&
      Height          =   345
      Left            =   855
      TabIndex        =   2
      Text            =   "danhsach"
      Top             =   2520
      Width           =   2760
   End
   Begin VB.DirListBox d2 
      BackColor       =   &H00FFFFC0&
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   4530
   End
   Begin VB.DriveListBox d1 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   0
      Width           =   3345
   End
   Begin VB.Label khanh 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Look in"
      Height          =   210
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   30
      Width           =   750
   End
   Begin VB.Label dc 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   2130
      Width           =   3675
   End
   Begin VB.Label khanh 
      BackColor       =   &H00FFC0C0&
      Caption         =   "File name"
      Height          =   210
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   2565
      Width           =   750
   End
End
Attribute VB_Name = "Save1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tamdia
Private Sub d1_Change()
    On Error GoTo kt
        d2 = StrConv(d1, 3)
    Exit Sub
kt:
    If MsgBox("Ban hay coi lai dia cua ban, No bi loi o dau do" & Chr(13) & Chr(10) & "Ban co muon tiep tuc cong viec khong ?", vbYesNo, "Canh bao") = 6 Then
        d1_Change
    Else
        d1 = tamdia
    End If
End Sub
Private Sub d2_Change()
    dc.Caption = d2
End Sub
Private Sub Form_Load()
    tamdia = d1
    d2 = StrConv(d1, 3)
    dc.Caption = d2
End Sub

Private Sub MDSave_Click(Index As Integer)
    On Error GoTo kt
    If Index = 0 Then
        If Len(ten.Text) > 0 Then
            If InStr(1, ten.Text, ".", 1) > 0 Then ten.Text = Left(ten.Text, InStr(1, ten.Text, ".", 1) - 1)
                ten.Text = ten.Text & "." & TypeOfExtension
                If Len(d2) < 4 Then
                    ddan = d2 & ten.Text
                Else
                    ddan = d2 & "\" & ten.Text
                End If
                Hamsave
               Unload Save1
            Else
                MsgBox "Ban phai danh ten can luu vao hop File name", vbOKOnly, "Thong Bao"
        End If
    Else
        Unload Save1
    End If
    Exit Sub
kt:
End Sub


