Attribute VB_Name = "CacTienIchLatVat"
Const AW_CENTER = &H10 'Hi?u ?ng Window "?n vào" trong n?u AW_HIDE dý?c dùng chung và "hi?n ra" n?u AW_HIDE không dý?c dùng.
Const AW_HIDE = &H10000 'Qui d?nh ?n Window.
Const AW_ACTIVATE = &H20000 'Qui d?nh kích ho?t m?t Window
Const AW_SLIDE = &H40000 'Qui d?nh dùng hi?u ?ng Slide
Const AW_BLEND = &H80000 'Dùng hi?u ?ng Fade. Ch? có hi?u l?c khi Window là Top_level.
Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean
Const LB_FINDSTRING = &H18F
Public Sub TimNhanh(KyTuGoVao)
MDDanhsach.List1.ListIndex = SendMessage(MDDanhsach.List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(KyTuGoVao))
End Sub
Public Sub Tatsinhdong(ByVal Handle_Of_OBJ)
  AnimateWindow Handle_Of_OBJ, 1000, AW_CENTER Or AW_HIDE Or AW_BLEND
End Sub
