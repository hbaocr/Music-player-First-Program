VERSION 5.00
Begin VB.Form WinSeek 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add folder"
   ClientHeight    =   1935
   ClientLeft      =   4125
   ClientTop       =   2805
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3990
   Begin VB.ComboBox KieuFile 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "addfolder.frx":0000
      Left            =   2040
      List            =   "addfolder.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   1905
   End
   Begin VB.Timer gh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1560
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   1080
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   360
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1185
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1320
      ScaleHeight     =   1695
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox lstFoundFiles 
         Height          =   1230
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblfound 
         Caption         =   "&Files Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   2055
      End
      Begin VB.DirListBox dirList 
         Height          =   990
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin VB.FileListBox filList 
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtSearchSpec 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Text            =   "*.*"
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "WinSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SearchFlag As Integer

Private Sub cmdExit_Click()
cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  
    If dirList.Path <> dirList.list(dirList.ListIndex) Then
        dirList.Path = dirList.list(dirList.ListIndex)
        Exit Sub
    End If
  Picture2.Move 0, 0
  Picture1.Visible = False
  Picture2.Visible = True
    
        KieuFile_Change
    FirstPath = dirList.Path
    DirCount = dirList.ListCount
    NumFiles = 0                       '
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True
    DirDiver = False
    retval = DoEvents()
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
                        dirList.Path = dirList.list(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
                DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
        If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then
            ThePath = dirList.Path
        Else
            ThePath = dirList.Path + "\"
            End If
        For ind = 0 To filList.ListCount - 1
            entry = ThePath + filList.list(ind)
            lstFoundFiles.AddItem entry
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
       gh.Enabled = True
          Exit Function
           End If
    If BackUp <> "" Then
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then
        DirDiver = True
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function
    Else
        MsgBox Error
        End
    End If
End Function

Private Sub DirList_Change()
        filList.Path = dirList.Path
End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.list(dirList.ListIndex)
End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub Form_Load()
gh.Enabled = False
    Picture2.Move 0, 0
    Picture2.Width = WinSeek.ScaleWidth
    Picture2.BackColor = WinSeek.BackColor
    lblCount.BackColor = WinSeek.BackColor
'    lblCriteria.BackColor = WinSeek.BackColor
    lblfound.BackColor = WinSeek.BackColor
    Picture1.Move 0, 0
    Picture1.Width = WinSeek.ScaleWidth
    Picture1.BackColor = WinSeek.BackColor
End Sub
Private Sub ResetSearch()
        lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False
    Picture2.Visible = False
    Picture1.Visible = True
    dirList.Path = CurDir: drvList.Drive = dirList.Path
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub gh_Timer()

On Error GoTo kt
    For i = 1 To lstFoundFiles.ListCount
            DsT(i) = lstFoundFiles.list(i - 1)
    Next i
    SLT = lstFoundFiles.ListCount
    Unload WinSeek
    XlFile
kt:
End Sub

Private Sub KieuFile_Change()
      On Error GoTo kt
filList.Pattern = Right(KieuFile, 5)
If Right(KieuFile, 5) = "iles)" Then filList.Pattern = "*.wav;*.mp3;*.Dat;*.Avi;*.mid;*.snd;*.au;*.aif*.aifc;*.aiff;*.cda"
If Right(KieuFile, 5) = "e *.*" Then filList.Pattern = "*.*"
Exit Sub
kt:

End Sub

Private Sub txtSearchSpec_Change()
      ' filList.Pattern = txtSearchSpec.Text

End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub

