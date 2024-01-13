Attribute VB_Name = "MDmain"
Option Explicit
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public gtc, giatri
Public luutam As Boolean
Public Ds(1000) As String, SL As Integer, ddan As String, coOpen As Boolean
Public DsT(1000) As String, SLT As Integer
Public Cochay As Boolean, Baiso As Integer, i As Integer
Public Rcolor As Integer, Gcolor As Integer, Bcolor As Integer, cobaomau As Boolean
Public RTcolor As Integer, GTcolor As Integer, BTcolor As Integer, cochu As Boolean
Public tentam(1000)
Public tentam2(1000)
Public tentam3(1000)
Public tenddu(1000)
Public kiemtra, kiemtra1, kiemtra3 As Boolean
Public thaydoi, thaydoi1 As Boolean
Public dem, dem2



Function Xlten(ten As String) As String
On Error GoTo kt
    For i = Len(ten) To 0 Step -1
        If Mid(ten, i, 1) = "\" Then
            Xlten = StrConv(Mid(ten, i + 1, Len(ten)), 3)
            i = 0
        Else
            Xlten = StrConv(ten, 3)
        End If
    Next i
Exit Function
kt:
End Function
Function XlFile()
On Error GoTo kt
    Dim co As Boolean, M3u(100) As String, dem As Integer
    dem = 0
    co = False
        For i = 1 To SLT
               If StrConv(Right(DsT(i), 3), 3) <> "ngh" Then
                    Ds(i + SL - dem) = DsT(i)
               Else
                    dem = dem + 1
                    M3u(dem) = DsT(i)
                    co = True
               End If
        Next i
        SL = SL + SLT - dem
        If co = True Then
            For i = 1 To dem
                XLM3u M3u(i)
            Next i
        End If
        coOpen = True
Exit Function
kt:
End Function
Function XLM3u(Fm3u As String)
On Error GoTo kt
Dim h As String, t As Integer
        t = FreeFile
        Open Fm3u For Input As #t
            Do While Not EOF(t)
                Input #t, h
                If Len(h) > 3 Then
                        If InStr(1, h, ":", 1) = 0 Then
                            SL = SL + 1
                            If Left(h, 1) = "\" Then
                                Ds(SL) = Left(Fm3u, 2) & h
                            Else
                                Ds(SL) = Left(Fm3u, 3) & h
                            End If
                        Else
                            SL = SL + 1
                            Ds(SL) = h
                        End If
                End If
            Loop
         Close #t
         coOpen = True
    Exit Function
kt:
End Function
Function Hamsave()
Dim h As Integer, m As String
    On Error GoTo kt
        For i = 1 To SL
            h = FreeFile
            m = Ds(i)
            Open ddan For Append As #h
                If Len(m) > 3 Then
                    Print #h, m
                End If
            Close #h
        Next i
Exit Function
kt:
MsgBox "Co loi khi luu", vbOKOnly, "Thong bao"
End Function
Function Hamsavetam()
'On Error GoTo kt
Dim h As Integer, m As String
Dim strf
    'On Error GoTo kt
    On Error Resume Next
    strf = App.PATH + "\danhsach." & TypeOfExtension
    Kill strf
        For i = 1 To SL
            h = FreeFile
            m = Ds(i)
        
           Open strf For Append As #h
            If Len(m) > 3 Then
                    Print #h, m
                    End If
            Close #h
        Next i
'Exit Function
'kt:
'MsgBox "Co loi khi luu", vbOKOnly, "Thong bao"
End Function
Public Function tim(ByVal chuoi As String) As String
Dim a, l, K, j
Dim s, D, s1
Dim kq As Boolean
For a = 0 To Len(chuoi)
s = Right(chuoi, a)
s1 = Left(s, 2)
If s1 = "<>" Then
l = a - 2
kq = True
Exit For
Else
kq = False
End If
Next a
If kq = True Then
D = Right(chuoi, l)
tim = D
Else
D = chuoi
End If
End Function


Public Function boduoi(chuoi As String)
Dim a, l, K, j
Dim s, D, s1
Dim kq As Boolean
For a = 0 To Len(chuoi)
s = Left(chuoi, a)
s1 = Right(s, 1)
If s1 = "." Then
l = a - 1
kq = True
Exit For
Else
kq = False
End If
Next a
If kq = True Then
D = Left(chuoi, l)
boduoi = D
Else
boduoi = chuoi
End If

End Function
Public Function TaoFileTam(ByVal PATH As String, ByVal Extension As String)
  Dim h
  Dim FilePath As String
  FilePath = Trim$(App.PATH & "FileTam." & Extension)
  h = FreeFile
 Open FilePath For Append As #h
            If Len(PATH) > 3 Then
                    Print #h, PATH
            End If
 Close #h
 TaoFileTam = FilePath
End Function
Public Sub XoaFileTam(ByVal DelPathFile)
Kill DelPathFile
End Sub
