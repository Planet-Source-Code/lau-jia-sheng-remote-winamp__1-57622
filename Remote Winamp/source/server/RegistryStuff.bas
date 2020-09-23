Attribute VB_Name = "RegistryStuff"

'Declarations required to edit the Windows Registry:
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Public Const REG_NONE = (0)                         'No value type
Public Const REG_SZ = (1)                           'Unicode nul terminated string
Public Const REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
Public Const REG_BINARY = (3)                       'Free form binary
Public Const REG_DWORD = (4)                        '32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
Public Const REG_LINK = (6)                         'Symbolic Link (unicode)
Public Const REG_MULTI_SZ = (7)                     'Multiple Unicode strings
Public Const REG_RESOURCE_LIST = (8)                'Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)
Const READ_CONTROL = &H20000
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
End Type
Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
End Type
Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long

Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Const LB_ITEMFROMPOINT = &H1A9

 Sub WriteRegistry(hKey As Long, SubKey As String, _
    ValueName As String, vNewValue As String)
    Dim Result As Long, RetVal As Long
    
    RetVal = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, Result)
    RetVal = RegSetValueEx(Result, ValueName, 0, REG_SZ, vNewValue, CLng(Len(vNewValue) + 1))
    
    RegCloseKey hKey
    RegCloseKey Result

End Sub
Sub RemoveFromRegistry(KeyName As String)
Dim RetVal As Long, hKey As Long, ValueName As String, _
        SubKey As String, Result As Long, SA As SECURITY_ATTRIBUTES, _
        Create As Long
    RetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", _
        0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        SA, Result, Create)
    RetVal = RegDeleteValue(Result, KeyName)
    RegCloseKey Result
End Sub
Function PullValue(hKey As Long, Place As String, SubValue As String)
 Dim szBuffer As String, dataBuff As String, ldataBuffSize As Long, phkResult As Long, RetVal As Long, _
        Value As String, RegEnumIndex As Long
    dataBuff = Space(255)
    ldataBuffSize = Len(dataBuff)
    szBuffer = Place
    RetVal = RegOpenKeyEx(hKey, szBuffer, 0, KEY_ALL_ACCESS, phkResult)
    Value = SubValue
    RetVal = RegQueryValueEx(phkResult, Value, 0, 0, dataBuff, ldataBuffSize)
    If RetVal = ERROR_SUCCESS Then
            's$ = ConvertString(dataBuff, ldataBuffSize)
   'PullValue = s$
    MsgBox "The Pullvalue was tried", 16, "Error"
    
    Else
        MsgBox "Error in retreiving that value.", 16, "Error"
    End If
End Function


