VERSION 5.00
Begin VB.Form Open1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Files lists"
   ClientHeight    =   2745
   ClientLeft      =   2790
   ClientTop       =   3795
   ClientWidth     =   6465
   Icon            =   "Open1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6465
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   4440
   End
   Begin VB.DriveListBox DR 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   1905
   End
   Begin VB.DirListBox DR1 
      BackColor       =   &H00FFFFC0&
      Height          =   2340
      Left            =   60
      TabIndex        =   7
      Top             =   360
      Width           =   1920
   End
   Begin VB.FileListBox DR2 
      BackColor       =   &H00FFFFC0&
      Height          =   2430
      Left            =   1980
      MultiSelect     =   1  'Simple
      TabIndex        =   6
      Top             =   0
      Width           =   4440
   End
   Begin VB.ComboBox KieuFile 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "Open1.frx":0442
      Left            =   2000
      List            =   "Open1.frx":045B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2400
      Width           =   2625
   End
   Begin VB.ListBox DSFile 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton OpenAdd 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   5640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   810
   End
   Begin VB.CommandButton OpenAdd 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Add all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   4680
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   810
   End
   Begin VB.CommandButton OpenAdd 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Remove all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5880
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton OpenOk 
      BackColor       =   &H00FFC0FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "Open1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codem(1000) As Boolean, soluong As Integer
Dim f, v, f1, v1, f2, v2
Dim st
Dim down As Boolean

Private Sub DR2_Click()
    Open1.Caption = DR2.ListIndex
End Sub
'Private Sub DR2_DblClick()
'     If Right(DR2.Path, 1) = "\" Then
'        DSFile.AddItem DR2.Path & DR2.list(DR2.ListIndex)
'     Else
'        DSFile.AddItem DR2.Path & "\" & DR2.list(DR2.ListIndex)
'    End If
'    End Sub


Private Sub DR2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo kt
codem(DR2.ListIndex) = DR2.Selected(DR2.ListIndex)
Exit Sub
kt:
End Sub

Private Sub DR2_KeyPress(KeyAscii As Integer)
On Error GoTo kt
codem(DR2.ListIndex) = DR2.Selected(DR2.ListIndex)
Exit Sub
kt:
End Sub

Private Sub DR2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo kt
codem(DR2.ListIndex) = DR2.Selected(DR2.ListIndex)
Exit Sub
kt:
End Sub

Private Sub DSFile_Click()
'    gobo
End Sub


Private Sub Form_Load()
On Error GoTo kt
down = False
For i = 0 To 1000
codem(i) = False
Next i
    SLT = 0
    KieuFile.ListIndex = 3
    Exit Sub
kt:
End Sub
Private Sub DR_Change()
On Error Resume Next
Dim loi
        DR1 = DR
        If Err <> 0 Then
       loi = MsgBox("Disk not ready", vbOKOnly, "Ready")
        If loi = vbOK Then
        Exit Sub
        DR1.Path = "c:"
        End If
        Else
        DR1 = DR
        End If
End Sub
Private Sub DR1_Change()
        DR2 = DR1
End Sub

Private Sub KieuFile_Click()
On Error GoTo kt
DR2.Pattern = Right(KieuFile, 5)
If Right(KieuFile, 5) = "iles)" Then DR2.Pattern = "*.wav;*.mp3;*.Dat;*.Avi;*.mid;*.snd;*.au;*.aif*.aifc;*.aiff;*.cda;*.wma"
If Right(KieuFile, 5) = "e *.*" Then DR2.Pattern = "*.*"
Exit Sub
kt:
End Sub

Private Sub OpenAdd_Click(Index As Integer)
On Error GoTo kt
Select Case Index
Case 0 ' Add
soluong = DR2.ListCount
goi
Case 1 'Add all
          For i = 0 To DR2.ListCount - 1
            If Right(DR2.Path, 1) = "\" Then
                DSFile.AddItem DR2.Path & DR2.list(i)
            Else
                DSFile.AddItem DR2.Path & "\" & DR2.list(i)
            End If
    thaydoi1 = False
                 
          Next i
Case 2 'Remove all
           DSFile.Clear
End Select
OpenOk_Click
Exit Sub
kt:
End Sub

Private Sub OpenOk_Click()
    For i = 1 To DSFile.ListCount
            DsT(i) = DSFile.list(i - 1)
    Next i
    SLT = DSFile.ListCount
    Unload Open1
    XlFile
    thaydoi = True
    End Sub
Function gobo()
    On Error GoTo kt
        Dim mang(1000) As String, dem As Integer
        dem = 0
        For i = 0 To DSFile.ListCount
            If i <> DSFile.ListIndex Then
                mang(i) = DSFile.list(i)
                dem = dem + 1
                Else
                    mang(i) = ""
            End If
        Next i
        DSFile.Clear
        For i = 0 To dem
            If mang(i) <> "" Then DSFile.AddItem mang(i)
        Next i
Exit Function
kt:
End Function
Function goi()
On Error GoTo kt
For i = 0 To soluong
    If codem(i) = True Then
        If Right(DR2.Path, 1) = "\" Then
                DSFile.AddItem DR2.Path & DR2.list(i)
            Else
                DSFile.AddItem DR2.Path & "\" & DR2.list(i)
        End If
    End If
Next i
Exit Function
kt:
End Function
Private Sub dsfile_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyDown
md
Case vbKeyUp
mu
End Select
End Sub

Private Sub mu()
Dim st
For f2 = 0 To DSFile.ListCount - 1
If DSFile.Selected(f2) = True Then
If f2 > 0 Then
st = DSFile.list(f2)
DSFile.RemoveItem f2
DSFile.AddItem st, f2 - 1
DSFile.Selected(f2) = True
End If
Exit For
End If
Next f2
End Sub
Private Sub md()
Dim st
For v2 = 0 To DSFile.ListCount - 1
          If DSFile.Selected(v2) = True Then
If v2 < DSFile.ListCount - 1 Then
st = DSFile.list(v)
DSFile.RemoveItem v2
DSFile.AddItem st, v2 + 1
DSFile.Selected(v2) = True
End If
 Exit For
        End If
        Next v2
End Sub

Private Sub Timer1_Timer()
If DSFile.list(0) = "" Then
OpenOk.Enabled = False
Else
OpenOk.Enabled = True
End If
End Sub
