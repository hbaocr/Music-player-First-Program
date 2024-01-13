Attribute VB_Name = "DRAGandDROP"
  Option Explicit
  Public Const MAX_PATH As Long = 260&
  
  Public Const WM_DROPFILES As Long = &H233&


  Public procOld As Long
  'bien toan cuc tra ve duong dan cua file duoc DRAG&DROP (keo va tha)
  Public sFileName$
  
  
  Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, _
                                                    ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
                                                    
  Public Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd&, ByVal fAccept&)
                               
  Public Declare Function DragQueryFile& Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop&, ByVal iFile&, _
                                                                                  ByVal lpszFile$, ByVal cch&)
  Public Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop&)
'ham nhan bat nhung thong diep cua window gui toi
  
Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                                              ByVal wParam As Long, ByVal lParam As Long) As Long
    'ham nhan bat nhung thong diep cua window gui toi

  Select Case iMsg
    
    Case WM_DROPFILES
     DropFiles wParam
      WindowProc = False
      Exit Function
      
  End Select
  
  WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)

End Function



Public Sub DropFiles(ByVal hDrop&)

  Dim nCharsCopied&
  
  sFileName = String$(MAX_PATH, vbNullChar)
  nCharsCopied = DragQueryFile(hDrop, 0&, sFileName, MAX_PATH)

  '
  DragFinish hDrop
  
 
  If nCharsCopied Then
    sFileName = Left$(sFileName, nCharsCopied)
  
 
    On Error GoTo Kohople
    
    XuLyFileDragDrop sFileName
  End If
  
  Exit Sub
  
Kohople:

  
End Sub




Public Sub XuLyFileDragDrop(ByVal Filepath As String)
Dim Cantren
MDDanhsach.List1.Clear
Cantren = MDDanhsach.List1.ListCount
DsT(Cantren + 1) = Filepath
    SLT = Cantren + 1
        XlFile
    thaydoi = True
End Sub





