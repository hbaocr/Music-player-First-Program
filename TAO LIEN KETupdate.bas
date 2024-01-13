Attribute VB_Name = "FileAssoc"

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)

Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_IDLIST = &H0&
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1&
Public Const ERROR_BADKEY = 2&
Public Const ERROR_CANTOPEN = 3&
Public Const ERROR_CANTREAD = 4&
Public Const ERROR_CANTWRITE = 5&
Public Const ERROR_OUTOFMEMORY = 6&
Public Const ERROR_INVALID_PARAMETER = 7&
Public Const ERROR_ACCESS_DENIED = 8&

Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const KEY_SET_VALUE = &H2&
Public Const MAX_PATH = 260&
Public Const REG_DWORD As Long = 4
Public Const REG_SZ = 1
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL

Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public TypeOfExtension As String

Public Sub MakeFileAssociation(Duoi_Mo_Rong_File_Link_Ung_Dung As String, Duong_Dan_Den_Ung_Dung_Lien_Ket As String, Ten_Ung_Dung_Lien_Ket As String, Ghi_Chu As String, Optional FulliconPath As String)
Dim Ret&
If Left(Duong_Dan_Den_Ung_Dung_Lien_Ket, 1) <> "\" Then Duong_Dan_Den_Ung_Dung_Lien_Ket = Duong_Dan_Den_Ung_Dung_Lien_Ket & "\"
'tao 1 duong dan tao moi lien ket giua file co duoi la *.xyz (Duoi_Mo_Rong_File_Link_Ung_Dung) voi ung dung can chay cua no
sKeyName = "." & Duoi_Mo_Rong_File_Link_Ung_Dung
sKeyValue = Ten_Ung_Dung_Lien_Ket
Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
'thiet lap KEY cua ung dung va mo ta them (Ghi_Chu)
sKeyName = Ten_Ung_Dung_Lien_Ket
sKeyValue = Ghi_Chu
Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
'thiet lap bieu tuong mac ding (default) cho file co duoi mo rong la *.xyz
If FulliconPath <> "" Then
    sKeyName = Ten_Ung_Dung_Lien_Ket & "\DefaultIcon"
    sKeyValue = FulliconPath
    Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0

    End If
'thiet lap khoa(KEy) de tao lien ket toi file lien ket *.xyz
sKeyName = Ten_Ung_Dung_Lien_Ket & "\shell\open\command"
sKeyValue = Chr(34) & Duong_Dan_Den_Ung_Dung_Lien_Ket & Ten_Ung_Dung_Lien_Ket & ".exe" & Chr(34) & " %1"
Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)

End Sub
Public Sub ASSOCIATE_TYPE(TypeEx As String, ValueApplicationName As String, IconPath As String)
Dim sKeyName As String
Dim sKeyValue As String

Dim Ret&
Dim lphKey&
Dim Path As String

Path = App.Path
If Right(Path, 1) <> "\" Then
Path = Path & "\"
End If
'tao ghi chu
sKeyName = App.EXEName
sKeyValue = App.EXEName & " 's file"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
'tao lien ket
sKeyName = Trim("." & TypeEx)
sKeyValue = ValueApplicationName
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
'set icon
sKeyName = ValueApplicationName
sKeyValue = IconPath
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, _
sKeyValue, MAX_PATH)

'Ðoi Icon
SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0



End Sub

Public Sub DeleteFileAssociation(Duoi_Mo_Rong_File_Link_Ung_Dung As String)
Dim Application As String
Dim Ret&
'Kiem tra xem file nay da duoc dang ky lien ket voi ung dung nao chua
Application = ReadKey(HKEY_CLASSES_ROOT, "." & Duoi_Mo_Rong_File_Link_Ung_Dung, "", "")
If Application <> "" Then
    'Huy dang ky lien ket cua file co fan mo rong la *.xyz
    Ret& = DeleteKey(HKEY_CLASSES_ROOT, "." & Duoi_Mo_Rong_File_Link_Ung_Dung)
    'Huy lien ket voi Ung dung can chay
    Ret& = DeleteKey(HKEY_CLASSES_ROOT, Application)
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0

   End If
End Sub

Public Function CheckFileAssociation(ByVal Duoi_Mo_Rong_File_Link_Ung_Dung As String) As String
Duoi_Mo_Rong_File_Link_Ung_Dung = "." & Duoi_Mo_Rong_File_Link_Ung_Dung
'Kiem tra cac Ung dung lien ket voi file *.xyz
CheckFileAssociation = ReadKey(HKEY_CLASSES_ROOT, Duoi_Mo_Rong_File_Link_Ung_Dung, "", "")
End Function

Public Function ReadKey(ByVal KeyName As String, ByVal SubKeyName As String, ByVal ValueName As String, ByVal DefaultValue As String) As String
Dim sBuffer As String
Dim lBufferSize As Long
Dim Ret&
sBuffer = Space(255)
lBufferSize = Len(sBuffer)
Ret& = RegOpenKey(KeyName, SubKeyName, 0, KEY_READ, lphKey&)
If Ret& = ERROR_SUCCESS Then
    Ret& = RegQueryValue(lphKey&, ValueName, 0, REG_SZ, sBuffer, lBufferSize)
    Ret& = RegCloseKey(lphKey&)
    Else
    Ret& = RegCloseKey(lphKey&)
    End If
sBuffer = Trim(sBuffer)
If sBuffer <> "" Then
    sBuffer = Left(sBuffer, Len(sBuffer) - 1)
    Else
    sBuffer = DefaultValue
    End If
ReadKey = sBuffer
End Function

Public Function WriteKey(ByVal KeyName As String, ByVal SubKeyName As String, ByVal ValueName As String, ByVal KeyValue As String) As Long
Dim Ret&
Ret& = RegCreateKey&(KeyName, SubKeyName, lphKey&)
If Ret& = ERROR_SUCCESS Then
    Ret& = RegSetValue&(lphKey&, ValueName, REG_SZ, KeyValue, 0&)
    Else
    Ret& = RegCloseKey(lphKey&)
    End If
WriteKey = Ret&
End Function

Public Function DeleteKey(ByVal KeyName As String, ByVal SubKeyName As String) As Long
Dim Ret&
Ret& = RegOpenKey(KeyName, SubKeyName, 0, KEY_WRITE, lphKey&)
If Ret& = ERROR_SUCCESS Then
    Ret& = RegDeleteKey(lphKey&, "")
    Ret& = RegCloseKey(lphKey&)
End If
DeleteKey = Ret&
End Function

Public Sub CheckFileClick(PathOfFileClick As String)

End Sub
