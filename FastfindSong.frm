VERSION 5.00
Begin VB.Form FastfindSong 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Xin NHap Ten Bai Hat"
   ClientHeight    =   615
   ClientLeft      =   2940
   ClientTop       =   4260
   ClientWidth     =   4110
   LinkTopic       =   "Form3"
   ScaleHeight     =   615
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Line Line4 
      X1              =   4080
      X2              =   4080
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Nhap Ten Bai Hat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FastfindSong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
Unload FastfindSong
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbEnter Then
Unload FastfindSong
End If
End Sub

Private Sub Text1_Change()
TimNhanh Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbEnter Then
Unload FastfindSong
End If

End Sub
