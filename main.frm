VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form main 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My song in your Computer"
   ClientHeight    =   1005
   ClientLeft      =   2445
   ClientTop       =   2340
   ClientWidth     =   5190
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5190
   Begin VB.CommandButton Image3 
      Height          =   255
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1680
      Width           =   405
   End
   Begin VB.CommandButton Image2 
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1680
      Width           =   405
   End
   Begin VB.PictureBox pic1 
      Height          =   160
      Left            =   45
      ScaleHeight     =   105
      ScaleWidth      =   5055
      TabIndex        =   26
      Top             =   0
      Width           =   5115
      Begin VB.PictureBox pic2 
         BackColor       =   &H8000000D&
         FillColor       =   &H00FF8080&
         ForeColor       =   &H8000000D&
         Height          =   160
         Left            =   0
         ScaleHeight     =   105
         ScaleWidth      =   435
         TabIndex        =   27
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog OpIcon 
      Left            =   4680
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   480
      TabIndex        =   25
      Top             =   2880
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5400
      Top             =   720
   End
   Begin VB.Timer tgchay 
      Interval        =   500
      Left            =   7200
      Top             =   960
   End
   Begin VB.CommandButton DGPlayPause 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   315
      Width           =   405
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   420
      Index           =   1
      Left            =   45
      TabIndex        =   12
      Top             =   240
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   741
      BandCount       =   6
      FixedOrder      =   -1  'True
      _CBWidth        =   5115
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      BandBackColor1  =   65280
      MinHeight1      =   255
      Width1          =   405
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      BandEmbossPicture1=   -1  'True
      MinHeight2      =   360
      NewRow2         =   0   'False
      BandStyle2      =   1
      MinHeight3      =   360
      NewRow3         =   0   'False
      BandStyle3      =   1
      MinHeight4      =   360
      NewRow4         =   0   'False
      BandStyle4      =   1
      MinHeight5      =   360
      NewRow5         =   0   'False
      BandStyle5      =   1
      MinHeight6      =   360
      FixedBackground6=   0   'False
      NewRow6         =   0   'False
      BandStyle6      =   1
      AllowVertical6  =   0   'False
      Begin VB.CheckBox Check1 
         Caption         =   "Loop"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   75
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   75
         Width           =   405
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   270
         Left            =   4200
         TabIndex        =   22
         Top             =   75
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
         _Version        =   393216
         LargeChange     =   20
         Max             =   100
         SelStart        =   50
         TickFrequency   =   25
         Value           =   50
      End
      Begin VB.CommandButton Image4 
         Height          =   255
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   75
         Width           =   405
      End
      Begin VB.CommandButton Image7 
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   75
         Width           =   405
      End
      Begin VB.CommandButton DGN_B 
         Height          =   255
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   75
         Width           =   405
      End
      Begin VB.CommandButton DGN_B 
         Height          =   255
         Index           =   0
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   75
         Width           =   405
      End
      Begin VB.CommandButton DGPlayPause 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   555
         Width           =   405
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   575
         Width           =   575
      End
      Begin VB.CommandButton DGPlayPause 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   75
         Width           =   405
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   435
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   3000
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   767
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   5475
      _CBHeight       =   435
      _Version        =   "6.0.8169"
      BandBackColor1  =   65280
      Child1          =   "DGthanhchay"
      MinHeight1      =   375
      Width1          =   2715
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      BandEmbossPicture1=   -1  'True
      MinHeight2      =   360
      FixedBackground2=   0   'False
      NewRow2         =   0   'False
      BandStyle2      =   1
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Slider DGthanhchay 
         DragIcon        =   "main.frx":0442
         Height          =   375
         Left            =   30
         TabIndex        =   11
         ToolTipText     =   """"""
         Top             =   30
         Visible         =   0   'False
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   2
         TickStyle       =   3
         TickFrequency   =   0
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   960
   End
   Begin MSComctlLib.StatusBar sta1 
      Height          =   255
      Left            =   45
      Negotiate       =   -1  'True
      TabIndex        =   21
      Top             =   720
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
            MinWidth        =   1235
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label dhb 
      Caption         =   "Label5"
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   11655
   End
   Begin VB.Image h1 
      Height          =   390
      Index           =   0
      Left            =   6000
      Picture         =   "main.frx":074C
      ToolTipText     =   "Play"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image h6 
      Height          =   345
      Left            =   8040
      Picture         =   "main.frx":1016
      Stretch         =   -1  'True
      ToolTipText     =   "App Path"
      Top             =   960
      Width           =   360
   End
   Begin VB.Image h4 
      Height          =   375
      Left            =   7560
      Picture         =   "main.frx":1418
      ToolTipText     =   "Loop"
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image h5 
      Height          =   375
      Left            =   7560
      Picture         =   "main.frx":1C8E
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Heïn  giôø"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "0 Min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "AÂm thanh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Vol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Laëp laïi"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Thôøi gian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image h3 
      Height          =   390
      Left            =   7800
      Picture         =   "main.frx":2504
      ToolTipText     =   "Playlist"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image h2 
      Height          =   390
      Index           =   1
      Left            =   7320
      Picture         =   "main.frx":2DCE
      ToolTipText     =   "Next"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image h2 
      Height          =   390
      Index           =   0
      Left            =   6840
      Picture         =   "main.frx":3698
      ToolTipText     =   "Previous"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image h1 
      Height          =   390
      Index           =   1
      Left            =   6480
      Picture         =   "main.frx":3F62
      ToolTipText     =   "Stop"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5880
      Picture         =   "main.frx":482C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1335
   End
   Begin MediaPlayerCtl.MediaPlayer Chay 
      CausesValidation=   0   'False
      DragIcon        =   "main.frx":1776E
      DragMode        =   1  'Automatic
      Height          =   765
      Left            =   5400
      TabIndex        =   0
      Top             =   1200
      Width           =   675
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   0   'False
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   0
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   0   'False
      SendWarningEvents=   0   'False
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   0   'False
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -5000
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu dsad 
      Caption         =   "&File"
      Begin VB.Menu mmopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mmsave 
         Caption         =   "&Save list"
         Shortcut        =   ^Z
      End
      Begin VB.Menu qwe 
         Caption         =   "-"
      End
      Begin VB.Menu mmexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu adas 
      Caption         =   "&View"
      Begin VB.Menu mmlist 
         Caption         =   "&List"
         Shortcut        =   ^L
      End
      Begin VB.Menu mmoption 
         Caption         =   "&Option"
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mmselect 
         Caption         =   "&Select File"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MZ 
      Caption         =   "&Zoom"
      Visible         =   0   'False
      Begin VB.Menu MZoom 
         Caption         =   "50%"
         Index           =   0
      End
      Begin VB.Menu MZoom 
         Caption         =   "100%"
         Index           =   1
      End
      Begin VB.Menu MZoom 
         Caption         =   "150%"
         Index           =   2
      End
      Begin VB.Menu MZoom 
         Caption         =   "200%"
         Index           =   3
      End
      Begin VB.Menu MZoom 
         Caption         =   "Full screen"
         Index           =   4
      End
   End
   Begin VB.Menu vblist 
      Caption         =   "&List"
      Visible         =   0   'False
      Begin VB.Menu mmflist 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu vblist1 
      Caption         =   "&List"
      Visible         =   0   'False
      Begin VB.Menu mmflist1 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Option 
      Caption         =   "&Option"
      Begin VB.Menu SettingList 
         Caption         =   "Setting List Asociation"
         Begin VB.Menu Apply 
            Caption         =   "Apply Asociation List"
         End
         Begin VB.Menu Cancel 
            Caption         =   "Cancel Asociation List"
         End
      End
      Begin VB.Menu SettingFile 
         Caption         =   "Setting File Asociation"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Dim laplai
Dim giay0, phut0, giay1, phut1
Dim hen
Dim TAT1, TAT2, TAT3
Dim jhg

Private Sub Apply_Click()
Dim IconNamePath As String
Dim AplicationName As String
OpIcon.DialogTitle = "Chon Icon Cho File Lien Ket Ung Dung"
OpIcon.ShowOpen
IconNamePath = OpIcon.FileName
'luu default Icon Path
SaveSetting App.EXEName, "AsociationFile", "IconPathDefault", IconNamePath
'==================
AplicationName = App.Path & "\" & App.EXEName
MakeFileAssociation TypeOfExtension, App.Path, App.EXEName, App.EXEName, IconNamePath
End Sub

Private Sub Cancel_Click()
DeleteFileAssociation TypeOfExtension
End Sub

Private Sub Chay_Click(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
    If Chay.DisplaySize = mpFullScreen Then Chay.DisplaySize = mpOneHalfScreen
End Sub


Private Sub Command1_Click()
Load WinSeek
WinSeek.Show
End Sub


Private Sub Command2_Click()
Load Form1
Form1.Show
End Sub



Private Sub Command3_Click()
Dim Chieudaiten As Long, kq As Long, K(5) As Long
Dim k1 As Long, k2 As Long, k3 As Long, k4 As Long, k5 As Long, k6 As Long
Dim Info As WNDCLASS
Dim ten As String
ten = String(30, Chr$(0))
Chieudaiten = GetClassName(Me.hwnd, ten, Len(ten))
ten = Left(ten, Chieudaiten)
kq = GetClassInfo(App.hInstance, ten, Info)
With Info
K(0) = .style
K(1) = .lpfnwndproc
K(2) = .cbClsextra
K(3) = .hInstance
K(4) = .hCursor
K(5) = .hIcon
End With
'dhb.Caption = Str$(k(i))
'i = i + 1
i = App.UnattendedApp
dhb.Caption = i
End Sub



Private Sub DGN_B_Click(Index As Integer)
On Error GoTo kt
    Select Case Index
                Case 0
                    Baiso = Baiso - 1
                    If Baiso <= 0 Then
                        Baiso = 1
                        DGN_B(0).Enabled = False
                    End If
                Case 1
                    Baiso = Baiso + 1
                    If Baiso > SL Then
                        Baiso = SL
                        DGN_B(1).Enabled = False
                    End If
               End Select
               Cochay = True
Exit Sub
kt:
End Sub

Private Sub DGthanhchay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Chay.Pause
Timer1.Enabled = False
tgchay.Enabled = False

End Sub

Private Sub DGthanhchay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
tgchay.Enabled = True
Chay.Play
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If Shift <> 0 Then
If KeyCode = 112 Then
Load optionw
optionw.Show
End If
End Sub

Private Sub Form_Load()
Dim LinkPath As String
Dim DuongDanFileTam As String
pic2.Width = 0
TypeOfExtension = "ngh"
DuongDanFileTam = ""
LinkPath = Command$
If (UCase(Right(LinkPath, 3)) <> UCase("exe")) And (UCase(Right(LinkPath, 3)) <> UCase(TypeOfExtension)) Then
DuongDanFileTam = TaoFileTam(LinkPath, TypeOfExtension)
End If
Me.Caption = "Chuong trinh nghe nhac mini "
dhbtn = 1
 Dim Buffer As String
 'Load Form2
 'Form2.Show
 
  FormatMessage &H1000, ByVal 0&, GetCurrentThread, &H0, Buffer, 200, ByVal 0&
    'Show the message
'    MsgBox ("Xin chao cac ban")
On Error GoTo kt
DGPlayPause(0).Picture = h1(0).Picture
DGPlayPause(1).Picture = h1(1).Picture
DGN_B(0).Picture = h2(0).Picture
DGN_B(1).Picture = h2(1).Picture
Image7.Picture = h3.Picture
Image3.Picture = h4.Picture
Image2.Picture = h5.Picture
Image4.Picture = h6.Picture
sta1.Panels(1).Text = "Name        "
thaydoi = False

Chay.Volume = -Slider1.Value * 10
Label7.Caption = Str(giay1) + " : " + Str(phut1) + " / " + Str(giay0) + " : " + Str(phut0)
TAT1 = 60
 SL = 0: SLT = 0: Baiso = 0: Cochay = False
 Load MDDanhsach
 MDDanhsach.Show
Dim m As String

If Right(LinkPath, 3) <> TypeOfExtension Then
     If UCase(Right(LinkPath, 3)) = UCase("exe") Or (LinkPath = "") Then
          m = App.Path
          If Right(m, 1) = "\" Then
          m = Left(m, Len(m) - 1)
          End If
          m = m & "\" & "danhsach." & TypeOfExtension
         ddan = m
      Else
      m = DuongDanFileTam
      ddan = m
      End If
 Else
 
 m = LinkPath
 ddan = m
 End If
         
          gtc = Chay.Volume
          
 
    i = FreeFile
    
     Open m For Input As #i
        Input #i, m
            If Len(m) > 0 Then
                Cochay = True
                Baiso = 1
                XLM3u ddan
            End If
        Close #i
           Hienan
            DGPlayPause(0).Enabled = False
            'Unload Form2
            '++++++++++++++++++++++++++++
            'kill fileTAm ngay tai day
            If DuongDanFileTam <> "" Then
            XoaFileTam DuongDanFileTam
            End If
            
            '++++++++++++++++++++++++++++
            Call list
       Exit Sub
kt:
            SL = 0: SLT = 0: Baiso = 0: Cochay = False
           Hienan
          DGPlayPause(0).Enabled = False
          
       '   Unload Form2
       
End Sub
Private Sub DGPlayPause_Click(Index As Integer)
On Error GoTo kt
Dim loi
    Cochay = False
    Select Case Index
        Case 0
           If Baiso > 0 Then Chay.Play
           If Baiso = 0 Then
                Baiso = 1
                Cochay = True
           End If
           'Call list
               ' For loi = 1 To SL
               '      Load mmflist(loi)
               ' Next loi
               ' For loi = 0 To SL
               '      mmflist(loi).Caption = (loi) & " - " & boduoi(Xlten(Ds(loi)))
               ' Next loi
               '       vblist.Visible = True
               '       mmflist(0).Visible = False
        Case 1
            Chay.Pause
        End Select
        Exit Sub
kt:
End Sub

Private Sub Form_Resize()
If main.WindowState = 1 Then
    MDDanhsach.Visible = False
Else
If kiemtra3 = False Then
    Load MDDanhsach
    MDDanhsach.Visible = True
    
    Else
    MDDanhsach.Visible = False
    End If
End If
End Sub

'Private Sub Image2_Click()
'Chay.PlayCount = 999
'Image3.Visible = True
'Image2.Visible = False
'DGthanhchay.SelectRange = True
'End Sub
'
'Private Sub Image3_Click()
''Chay.PlayCount = 1
'image2.Visible = True
'Image3.Visible = False
'DGthanhchay.SelectRange = False
'End Sub

Private Sub Image4_Click()
Command1_Click
End Sub

Private Sub Image7_Click()
mmlist_Click
End Sub

Private Sub Label1_Click()
hen = InputBox("Nhap vao khoang thoi gian ma ban muon thoat khoi khoi chuong trinh (tinh bang phut )", "Time")
If hen <> "" And Val(hen) <> 0 Then
        Timer3.Enabled = True
        If hen < 10 Then
             Label2.Caption = Trim("0" + Str(hen - 1)) + ":" + Str(TAT1)
        Else
             Label2.Caption = Str(hen - 1) + ":" + Str(TAT1)
        End If
End If
End Sub

Private Sub Label4_Click()
laplai = Not laplai
If laplai Then
           Label4.Caption = "On"
Else
           Label4.Caption = "Off"
End If
End Sub

Private Sub Label8_Click()
Dim exe
exe = Shell("sndvol32.exe", vbNormalFocus)
End Sub

Private Sub Label9_Click()
Load Form2
Form2.Show
End Sub

Private Sub mmflist_Click(Index As Integer)
Dim i
      Cochay = True
      Baiso = Val(Index)
      mmflist(Index).Checked = True
   For i = 0 To Index - 1
      mmflist(i).Checked = False
   Next i
   For i = Index + 1 To SL
      mmflist(i).Checked = False
   Next i
End Sub

Private Sub mmflist1_Click(Index As Integer)
Dim i
      Cochay = True
      Baiso = Val(Index)
      mmflist1(Index).Checked = True
   For i = 0 To Index - 1
      mmflist1(i).Checked = False
   Next i
   For i = Index + 1 To dem2
      mmflist1(i).Checked = False
   Next i

End Sub

Private Sub mmopen_Click()
            Open1.Show
            Hienan
End Sub
Private Sub mmsave_Click()
            Save1.Show
End Sub
Private Sub mmexit_Click()
            Unload main
            Unload MDDanhsach
End Sub
Private Sub mmlist_Click()
On Error GoTo kt
kiemtra3 = False
            If MDDanhsach.Visible = True Then
                            Unload MDDanhsach
                            MDDanhsach.Show
                        Else
                             MDDanhsach.Show
                            Unload MDDanhsach
                            MDDanhsach.Show
                    End If
            Exit Sub
kt:
End Sub
Private Sub mmoption_Click()
'            Chay.ShowDialog mpShowDialogOptions
Load optionw
optionw.Show
End Sub
Private Sub mmselecfile_Click()
Dim t As String, ii As String
On Error GoTo kt
    ii = InputBox("Ban muon mo bai thu may trong danh sach" & Chr(13) & Chr(10) & "Ban chi duoc nhap con so" & Chr(13) & Chr(10) & "Trong danh sach co tat ca la " & SL & " bai", "Chon bai hat theo thu tu", Baiso)
    If IsNumeric(ii) Then
            If Val(ii) > SL Or Val(ii) < 1 Then
                    If MsgBox("So lon qua khong co trong danh sach" & Chr(13) & Chr(10) & "Danh sach chi co " & SL & " bai" & Chr(13) & Chr(10) & "Ban co muon tiep tuc cong viec hay khong ?", vbYesNo, "Thong Bao") = 6 Then
                             mmselecfile_Click
                        Else
                            GoTo kt
                    End If
            Else
                    Baiso = Val(ii)
                    Cochay = True
            End If
        Else
            If ii <> "" Then
                MsgBox "Ban chi nhap vao duoi dang con so", vbOKOnly, "Canh bao"
                mmselecfile_Click
            End If
        End If
Exit Sub
kt:
Hienan
End Sub


Private Sub mmselect_Click()
Dim t As String, ii As String
On Error GoTo kt
    ii = InputBox("Ban muon mo bai thu may trong danh sach" & Chr(13) & Chr(10) & "Ban chi duoc nhap con so" & Chr(13) & Chr(10) & "Trong danh sach co tat ca la " & SL & " bai", "Chon bai hat theo thu tu", Baiso)
    If IsNumeric(ii) Then
            If Val(ii) > SL Or Val(ii) < 1 Then
                    If MsgBox("So lon qua khong co trong danh sach" & Chr(13) & Chr(10) & "Danh sach chi co " & SL & " bai" & Chr(13) & Chr(10) & "Ban co muon tiep tuc cong viec hay khong ?", vbYesNo, "Thong Bao") = 6 Then
                             mmselecfile_Click
                        Else
                            GoTo kt
                    End If
            Else
                    Baiso = Val(ii)
                    Cochay = True
            End If
        Else
            If ii <> "" Then
                MsgBox "Ban chi nhap vao duoi dang con so", vbOKOnly, "Canh bao"
                mmselecfile_Click
            End If
        End If
Exit Sub
kt:
Hienan

End Sub



Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic2.Width = X
Chay.CurrentPosition = Chay.SelectionEnd * X / pic1.Width
End Sub

Private Sub pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic2.Width = X
Chay.CurrentPosition = Chay.SelectionEnd * X / pic1.Width
End Sub

Private Sub SettingFile_Click()
Load OpForm
OpForm.Show
End Sub

Private Sub Slider1_Change()
If Slider1.Value <= 0 Then
Chay.Mute = True
Else
Chay.Mute = False
Chay.Volume = -(Slider1.Max - Slider1.Value) * 20
End If
End Sub

Private Sub tgchay_Timer()
On Error GoTo kt
gtc = Chay.Volume
If coOpen = True Then mmlist_Click
sta1.Panels(1).Text = Right(sta1.Panels(1).Text, 1) + Left(sta1.Panels(1).Text, Len(sta1.Panels(1).Text) - 1)
            If Chay.FileName <> "" Then DGthanhchay = Chay.CurrentPosition
                phut1 = Int(Chay.CurrentPosition \ 60)
                giay1 = Int(Chay.CurrentPosition - phut1 * 60)
                phut0 = Int(Chay.SelectionEnd / 60)
                giay0 = Int(Chay.SelectionEnd - phut0 * 60)
                    If phut1 >= 0 And giay1 >= 0 And phut0 >= 0 And giay0 >= 0 Then
                    sta1.Panels(3).Text = Str(phut0) + " :" + Str(giay0)
                    sta1.Panels(2).Text = Str(phut1) + " :" + Str(giay1)

                    End If
                     pic2.Width = pic1.Width / DGthanhchay.Max * DGthanhchay
                     
                 '-------------------------------------------------
            If DGthanhchay = DGthanhchay.Max Then
                          If Check1.Value = 0 Then
                             Baiso = Baiso + 1
                           End If
                    Cochay = True
                    If Baiso > SL Then
                        'If Label4.Caption = "On" Then
                       Baiso = 0
                       Cochay = False
                       'End If
                       End If
            End If
         '-------------------------------------------------
          If Cochay = True Then
                   Chay.FileName = Ds(Baiso)
                   DGthanhchay.Max = Chay.SelectionEnd
                  sta1.Panels(1).Text = Space(20) & boduoi(Xlten(Ds(Baiso)))
                  Cochay = False
                   xlloi Ds(Baiso)
                   kiemtra = True
                   '     If MDDanhsach.Visible = True Then
                   '         Unload MDDanhsach
                   '         MDDanhsach.Show
                   '     Else
                   '         Unload MDDanhsach
                   '     End If
                              ':"":":::::::::::::::::::::::
                    Hienan
          End If
         
   '---------------Khoi dau ham with----------------------------------
        With Chay
            If SL > 0 Then
                    If .PlayState = mpPlaying Then
                        DGPlayPause(0).Enabled = False
                        DGPlayPause(1).Enabled = True
                        DGthanhchay.Enabled = True
                          DGN_B(1).Enabled = True
                    DGN_B(0).Enabled = True
                    Else
                        DGPlayPause(0).Enabled = True
                        DGPlayPause(1).Enabled = False
                        DGthanhchay.Enabled = False
                    End If
             Else
                    DGPlayPause(0).Enabled = False
                    DGPlayPause(1).Enabled = False
                    DGN_B(1).Enabled = False
                    DGN_B(0).Enabled = False
                    DGthanhchay.Enabled = False
           End If
        End With
    '---------------Ket thuc ham with----------------------------------
    coOpen = False
    Exit Sub
kt:
'If MsgBox("Can't play this song") = vbOK Then
Baiso = 0: Cochay = False: DGthanhchay.Value = 0
'End If

End Sub
Function Hamchay(TenFile As String)
    On Error GoTo kt
                Chay.FileName = TenFile
    Exit Function
kt:
End Function
Private Sub DGthanhchay_Scroll()
Chay.CurrentPosition = DGthanhchay
End Sub
Function Hienan()
'-----------Khoi dau ham With -------------------------------------
On Error GoTo kt
    With Chay
               .AllowChangeDisplaySize = True
               .AutoRewind = False
              .EnableContextMenu = False
              .ClickToPlay = False
              .SendMouseClickEvents = True
              .SendErrorEvents = True
              .SendMouseMoveEvents = False
              .SendOpenStateChangeEvents = False
              .SendPlayStateChangeEvents = False
              .SendWarningEvents = False
              .ShowAudioControls = False
              .ShowCaptioning = False
              .ShowControls = False
              .ShowDisplay = False
              .ShowGotoBar = False
              .ShowPositionControls = False
              .ShowStatusBar = False
              .ShowTracker = False
              .TransparentAtStart = False
              .VideoBorder3D = False
              .WindowlessVideo = False
'------------------------------------------------
'------------------------------------------------
              If SL > 0 Then
                    If Baiso <= 1 Then
                        DGN_B(0).Enabled = False
                    Else
                        DGN_B(0).Enabled = True
                    End If
                    If Baiso >= SL Then
                        DGN_B(1).Enabled = False
                    Else
                        DGN_B(1).Enabled = True
                    End If
            End If
'------------------------------------------------
                End With
'----------------Ket thuc Ham Whith chay------------------
Exit Function
kt:
End Function
Function xlloi(s As String)
If Chay.ErrorCode Or Chay.ErrorDescription <> "" Then
    If Len(s) <= 0 Then s = " 'Da chon' "
    MsgBox "Khong the doc duoc File " & s & Chr(13) & Chr(10) & "Co the duong dan toi File tren bi sai" & Chr(13) & Chr(10) & "Cung co khi Windows Media khong ho tro File " & s, vbOKOnly, "Thong bao"
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
Hamsavetam
   Tatsinhdong MDDanhsach.hwnd
    Unload MDDanhsach
    Tatsinhdong Open1.hwnd
    Unload Open1
    Tatsinhdong Save1.hwnd
    Unload Save1
    Tatsinhdong Form1.hwnd
    Unload Form1
    Tatsinhdong OpForm.hwnd
    Unload OpForm
    Tatsinhdong WinSeek.hwnd
    Unload WinSeek
   ' Tatsinhdong MDDanhsach.hwnd
    Tatsinhdong main.hwnd
'    Unload Form2
End Sub
'Private Sub Timer1_Timer()
'On Error Resume Next
'If MDDanhsach.WindowState <> 1 Or MDDanhsach.WindowState <> 2 Then
'MDDanhsach.Top = main.Top + main.Height
'MDDanhsach.Left = main.Left
'MDDanhsach.Width = main.Width
'MDDanhsach.Height = main.Height
'Else
'Exit Sub
'End If
'Chay.Volume = -giatri / 2
'If main.WindowState = 1 Then
'Unload MDDanhsach
'Else
'Load MDDanhsach
' End If
'kt:
'End Sub
'
Private Sub Timer2_Timer()
If thaydoi = True Then
If thaydoi1 = True Then
Call List1
Else
Call list
End If
thaydoi1 = False
thaydoi = False
End If

End Sub

Private Sub Timer3_Timer()
TAT1 = TAT1 - 1
If hen < 10 Then
         If TAT1 < 10 Then
             Label2.Caption = Trim("0" + Str(hen - 1)) + " :0" + Str(TAT1)
         Else
             Label2.Caption = Trim("0" + Str(hen - 1)) + " :" + Str(TAT1)
         End If
Else
         If TAT1 < 10 Then
             Label2.Caption = Str(hen - 1) + " :0" + Str(TAT1)
         Else
             Label2.Caption = Str(hen - 1) + " :" + Str(TAT1)
         End If
End If
If TAT1 <= 0 Then
            TAT1 = 60
            hen = hen - 1
     If hen <= 0 Then
            Unload main
     End If
End If
End Sub
Public Sub list()
On Error Resume Next
If SL <> 0 Then
For loi = 1 To SL
Unload mmflist(loi)
Next loi
For loi = 1 To SL
  Load mmflist(loi)
  mmflist(loi).Visible = True
  mmflist(loi).Caption = (loi) & " - " & boduoi(Xlten(Ds(loi)))
                Next loi
                      vblist.Visible = True
                      vblist1.Visible = False
                      mmflist(0).Visible = False
Else
vblist1.Visible = False
End If
End Sub
Public Sub List1()
On Error Resume Next
If SL <> 0 Then
For loi = 1 To 500
Unload mmflist1(loi)
Next loi
loi = 0
For loi = 1 To 500
Load mmflist1(loi)
mmflist1(loi).Visible = True
mmflist1(loi).Caption = (loi) & " - " & boduoi(Xlten(Ds(loi)))
Next loi
For loi = 1 To 500
If loi > dem2 Then
Unload mmflist1(loi)
End If
Next loi
vblist1.Visible = True
vblist.Visible = False
mmflist1(0).Visible = False
Else
vblist1.Visible = False
vblist.Visible = False
End If
End Sub

