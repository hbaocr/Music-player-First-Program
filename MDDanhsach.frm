VERSION 5.00
Begin VB.MDIForm MDDanhsach 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Files List"
   ClientHeight    =   4275
   ClientLeft      =   7950
   ClientTop       =   1725
   ClientWidth     =   2955
   Icon            =   "MDDanhsach.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox anh1 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6795
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   0
      Width           =   2955
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   4800
         Top             =   720
      End
      Begin VB.Timer hamgoilai 
         Interval        =   1
         Left            =   4800
         Top             =   1920
      End
      Begin VB.Timer goihamtg 
         Interval        =   1
         Left            =   4800
         Top             =   1200
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00800000&
         Height          =   4155
         Left            =   0
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.Menu mmformat 
      Caption         =   "&Format"
      Visible         =   0   'False
      Begin VB.Menu mmdelect 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mmremove 
         Caption         =   "&Remove all"
      End
      Begin VB.Menu mmcolor 
         Caption         =   "&Colour"
         Visible         =   0   'False
         Begin VB.Menu mmtc 
            Caption         =   "&Text colour"
         End
         Begin VB.Menu mmlc 
            Caption         =   "&List colour"
         End
      End
      Begin VB.Menu gufulghlbb 
         Caption         =   "-"
      End
      Begin VB.Menu mmadfiles 
         Caption         =   "&Add files"
      End
      Begin VB.Menu mmadpath 
         Caption         =   "&Add Path"
      End
      Begin VB.Menu mmto 
         Caption         =   "&Move to"
      End
      Begin VB.Menu FastFind 
         Caption         =   "Fast Find Song In Your Playing List"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "MDDanhsach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'+++++++++++++++++++++++++++++++++++++
  Private Const GWL_WNDPROC As Long = (-4&)
   
  
  Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, _
                                                              ByVal nIndex&, ByVal dwNewLong&)
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Dim Xx As Integer, Yy As Integer, Mcb As Integer
'Private Sub goihamtg_Timer()
'If cobaomau = True Then hamgoilai.Interval = 1
'End Sub
Dim x1, x2, l1, l2
Dim i22, i23, v2, f2
Dim t, t1, t2 As Integer
Dim op
Dim qw
Dim vitri


Private Sub anh1_KeyDown(KeyCode As Integer, Shift As Integer)
If ((KeyCode <= vbKeyZ) And (KeyCode >= vbKeyA)) Or ((vbKey1 < KeyCode) And (vbkeycode < 9)) Then
TimNhanh KeyCode
End If
End Sub

Private Sub FastFind_Click()
Load FastfindSong
FastfindSong.Show
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mmformat
End If
End Sub

Private Sub MDIForm_Load()
On Error GoTo kt
Dim gio1
Dim strfile As String
Dim Ret
Dim kl
Dim i23
'+++++++++++++++++++BAt cac su kien DragDrop++++++++++++++++++++++++++++
  DragAcceptFiles MDDanhsach.List1.hwnd, 1&


  procOld = SetWindowLong(MDDanhsach.List1.hwnd, GWL_WNDPROC, AddressOf WindowProc)
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
List1.Clear
kiemtra1 = True
            Dim i As Integer
            List1.Top = 45
            List1.Left = 45
            For i = 1 To SL
            List1.AddItem i & " <> " & boduoi(Xlten(Ds(i)))
            Next i
            List1.Selected(Baiso - 1) = True
           For i = 1 To List1.ListCount
                tentam(i) = Str(i)
                tentam2(i) = tim(List1.list(i))
           Next i
'DDanhsach.Top = main.Top + main.Height
'MDDanhsach.Left = main.Left
'MDDanhsach.Width = main.Width
'MDDanhsach.Height = main.Height
  Exit Sub
kt:
End Sub
Private Sub list1_DblClick()
On Error GoTo kt
Dim jkl

    For jkl = 1 To List1.ListCount
            DsT(jkl) = List1.list(jkl - 1)
            Next jkl
    SLT = List1.ListCount
Cochay = True
  Baiso = Val(List1.list(List1.ListIndex))
kt:
End Sub
Private Sub MMdelete_Click()
gobo
End Sub

Private Sub MDIForm_Resize()
On Error GoTo kt
    If MDDanhsach.WindowState <> 1 Or MDDanhsach.WindowState <> 2 Then
        anh1.Height = MDDanhsach.Height
        List1.Height = MDDanhsach.Height - 500
        List1.Width = MDDanhsach.Width - 300
        Else
'        MDDanhsach.Width = main.Width
'        MDDanhsach.Height = main.Height

    End If
    Exit Sub
kt:
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Call SetWindowLong(MDDanhsach.List1.hwnd, GWL_WNDPROC, procOld)
Unload MDDanhsach
kiemtra3 = True
End Sub

Private Sub mmadfiles_Click()
Load Open1
Open1.Show
End Sub

Private Sub mmadpath_Click()
Load WinSeek
WinSeek.Show
End Sub

Private Sub mmdelect_Click()
'gobo
xoanhieu
End Sub

Private Sub mmlc_Click()
common.ShowColor
List1.BackColor = common.Color
End Sub

Private Sub mmremove_Click()
On Error GoTo kt
For i = 0 To 1000
    Ds(i) = ""
Next i
    List1.Clear
   SL = 0
    Baiso = 0
    Cochay = False
   'MDIForm_Load
        Exit Sub
kt:
End Sub
Function gobo()
    On Error GoTo kt
    Dim jiu
        Dim mang(1000) As String, dem As Integer, tam(1000) As String
        For i = 0 To List1.ListCount
                   If i <> List1.ListIndex Then
                mang(i) = List1.list(i)
                dem = dem + 1
                Else
                jiu = i
                       If i <> Baiso - 1 Then
                    mang(i) = ""
                       Else
                       MsgBox ("Can't remove it")
                       Exit Function
                       End If
            End If
            Next i
       List1.Clear
        SL = 0
        For i = 0 To dem
             If mang(i) <> "" Then
                SL = SL + 1
                 List1.AddItem mang(i)
                Ds(SL) = Ds(Val(mang(i)))
            End If
            
        Next i
        
         MDIForm_Load
         List1.Selected(jiu) = True
Exit Function
kt:
End Function

Private Sub mmtc_Click()
common.ShowColor
List1.ForeColor = common.Color
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
For i22 = 0 To List1.ListCount - 1
     If List1.Selected(i22) = True Then
        t = i22 + 1
        Exit For
     End If
Next i22
If Shift <> 0 Then
     If KeyCode = vbKeyUp Then
      mu
      List1.SetFocus
    End If
If KeyCode = vbKeyDown Then
md
List1.SetFocus
End If
End If

If Shift <> 0 And KeyCode = vbKeyA Then
    Dim pao As Long
     For pao = 0 To List1.ListCount - 1
     List1.Selected(pao) = True
     Next pao
End If

If KeyCode = vbKeyDelete Then
'gobo
xoanhieu
List1.SetFocus
End If
If KeyCode = 13 Or KeyCode = 32 Then
list1_DblClick
List1.SetFocus
End If
If KeyCode = vbKeyM Then
mmto_Click
List1.SetFocus
End If
End Sub
Private Sub mu()
Dim st, st2
thaydoi = True
For f2 = 0 To List1.ListCount - 1
      If List1.Selected(f2) = True Then
            If f2 - 1 > 0 Then
                st = tim(List1.list(f2))
                st2 = tim(List1.list(f2 - 1))
                List1.RemoveItem f2
                List1.AddItem (t - 1) & " <> " & LTrim(st), f2 - 1
                List1.list(f2) = (f2 + 1) & " <> " & LTrim(st2)
                hjkl = Ds(t - 1)
                Ds(t - 1) = Ds(t)
                Ds(t) = hjkl
                List1.Selected(f2) = True
            Else
                MsgBox ("Can't move up")
            End If
      End If
Next f2
End Sub
Private Sub md()
Dim st, st2
  'List1.AddItem i & " - " & Xlten(Ds(i))
  thaydoi = True
For v2 = 0 To List1.ListCount - 1
         If List1.Selected(v2) = True Then
            If v2 + 1 < List1.ListCount - 1 Then
                 st = tim(List1.list(v2))
                 st2 = tim(List1.list(v2 + 1))
                 List1.RemoveItem v2
                 List1.AddItem (t + 1) & " <> " & LTrim(st), v2 + 1
                 List1.list(t - 1) = (t) & " <> " & LTrim(st2)
                 hjkl = Ds(t)
                 Ds(t) = Ds(t + 1)
                 Ds(t + 1) = hjkl
                 List1.Selected(v2) = True
            Else
                 MsgBox ("Can't move down")
            End If
       End If
Next v2
End Sub
Private Sub mmto_Click()
On Error Resume Next
op = Val(InputBox("Let input the new location which you want to replace !  ", "Move to"))
If op >= 1 Then
            For i = 0 To List1.ListCount - 1
               If List1.Selected(i) = True Then
                 qw = i + 1
               Exit For
              End If
            Next i
                          If op > qw Then
                                    For i = qw To op - 1
                                         If List1.Selected(i - 1) = True Then
                                             t = i + 1
                                         End If
                                         If kiemtra1 = True Then
                                         mdto
                                         Else
                                         kiemtra1 = True
                                         End If
                                    Next i
                          End If
                          If op < qw Then
                                    For i = qw To op + 1 Step -1
                                         If List1.Selected(i - 1) = True Then
                                             t = i + 1
                                         End If
                                         If kiemtra1 = True Then
                                         muto
                                         Else
                                         kiemtra1 = True
                                         End If
                                    Next i
                          End If
                          If op = qw Then
                              MsgBox ("The same location")
                          End If
Else
 If MsgBox("Can't move", vbOKOnly) = vbOK Then
 Exit Sub
 End If
End If
End Sub
Private Sub mdto()
Dim st, st2
thaydoi = True
     If i + 1 <= List1.ListCount Then
               st = tim(List1.list(i - 1))
               st2 = tim(List1.list(i))
               List1.RemoveItem i - 1
               List1.AddItem (i + 1) & " <> " & LTrim(st), i
               List1.list(i - 1) = (i) & " <> " & LTrim(st2)
               hjkl = Ds(i)
               Ds(i) = Ds(i + 1)
               Ds(i + 1) = hjkl
               List1.Selected(i) = True
    Else
              If MsgBox("Can't move", vbOKOnly) = vbOK Then kiemtra1 = False
              
                         
    End If
'kt:
End Sub
Private Sub muto()
Dim st, st2
Dim ter
thaydoi = True
ter = i - 1
If ter - 1 >= 0 Then
         st = tim(List1.list(ter))
         st2 = tim(List1.list(ter - 1))
         List1.RemoveItem ter
         List1.AddItem (ter) & " <> " & LTrim(st), ter - 1
         List1.list(ter) = (ter + 1) & " <> " & LTrim(st2)
         hjkl = Ds(ter + 1)
         Ds(ter + 1) = Ds(ter)
         Ds(ter) = hjkl
         List1.Selected(ter - 1) = True
        Else
                      If MsgBox("Can't move", vbOKOnly) = vbOK Then kiemtra1 = False

End If
 End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If kiemtra = True Then
For i = 0 To List1.ListCount - 1
 If i = Baiso - 1 Then
 List1.Selected(i) = True
 Else
 List1.Selected(i) = False
 End If
 Next i
   kiemtra = False
End If
End Sub
Function xoanhieu()
        Dim mang(1000) As String, tam(1000) As String
        dem = 1
        For i = 0 To List1.ListCount - 1
If List1.Selected(i) = False Then
mang(dem) = Ds(i + 1)
dem = dem + 1
End If
Next i
List1.Clear
For i = 0 To SL
Ds(i) = ""
Next i
For it = 1 To dem
If mang(it) <> "" Then
Ds(it) = mang(it)
List1.list(it - 1) = Str(it) + " <> " + boduoi(Xlten(Ds(it)))
End If
Next it
dem2 = List1.ListCount
SL = List1.ListCount
thaydoi = True
thaydoi1 = True
End Function
