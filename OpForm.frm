VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OpForm 
   Caption         =   "Option Form"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4035
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4035
   Begin VB.CommandButton ICON 
      Caption         =   "Chon Icon"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin MSComDlg.CommonDialog Cbo 
      Left            =   4320
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Apply 
      Caption         =   "Thiet Lap Lien Ket"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton SelectAll 
      Caption         =   "Chon Het"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chon File Lien Ket Voi Ung Dung"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox TypeExtension 
         Caption         =   "mpg"
         Height          =   435
         Index           =   8
         Left            =   1800
         TabIndex        =   12
         Top             =   1680
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "wmv"
         Height          =   435
         Index           =   7
         Left            =   1800
         TabIndex        =   11
         Top             =   1320
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "dat"
         Height          =   435
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "cda"
         Height          =   435
         Index           =   6
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "mp3"
         Height          =   435
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "mid"
         Height          =   435
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "asf"
         Height          =   435
         Index           =   3
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "wma"
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   700
      End
      Begin VB.CheckBox TypeExtension 
         Caption         =   "wav"
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   700
      End
   End
End
Attribute VB_Name = "OpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CBOPATH As String

Public Function CheckType(TOE As String) As Boolean
'TOE = viet tat TypeOfExtension
Dim i As Byte
CheckType = False
For i = 0 To 6
If UCase(Trim(TOE)) = UCase(Trim(OpForm.TypeExtension(i).Caption)) Then
CheckType = True
Exit Function
End If
Next i
End Function
Private Sub SavSetting()
For i = 0 To 8
SaveSetting App.EXEName, "AsociationFile", CStr(i), CStr(OpForm.TypeExtension(i).Caption) & CStr(OpForm.TypeExtension(i).Value)
Next i
End Sub

Private Sub SettingChosenType()
Dim ValCheck As String
Dim i As Byte
'0= ko
'1 = chon

For i = 0 To 8
ValCheck = GetSetting(App.EXEName, "AsociationFile", CStr(i))
OpForm.TypeExtension(i).Value = Val(Right(ValCheck, 1))
Next i

End Sub
Private Sub DelSetting()
DeleteSetting App.EXEName, "AsociationFile"
End Sub

Private Sub Apply_Click()
Dim i As Integer, j As Integer
Dim count As Boolean
Dim MyValdefaulT As String
SavSetting
    '------------------------
    MyValdefaulT = GetSetting(App.EXEName, "AsociationFile", "IconPathDefault")

    If CBOPATH = "" Then
    CBOPATH = MyValdefaulT
    End If
       For i = 0 To 8
            If OpForm.TypeExtension(i).Value = 1 Then
              ASSOCIATE_TYPE LCase(OpForm.TypeExtension(i).Caption), App.EXEName, CBOPATH
             Else
             ASSOCIATE_TYPE LCase(OpForm.TypeExtension(i).Caption), "", CBOPATH
             End If
        Next i
       Unload OpForm
   
End Sub

Private Sub Form_Load()
SettingChosenType
End Sub

Private Sub ICON_Click()
Cbo.DialogTitle = "Chon Icon cho File Lien Ket Ung Dung"
Cbo.ShowOpen
CBOPATH = Cbo.FileName
End Sub

Private Sub SelectAll_Click()
Dim i As Byte
For i = 0 To 8
TypeExtension(i).Value = 1
Next i
End Sub
