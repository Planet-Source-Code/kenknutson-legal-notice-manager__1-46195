Attribute VB_Name = "RegPlus"
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Global Cretnew As Boolean

Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
    REG_NONE = (0)
    REG_SZ = (1)
    REG_EXPAND_SZ = (2)
    REG_BINARY = (3)
    REG_DWORD = (4)
    REG_DWORD_LITTLE_ENDIAN = (4)
    REG_DWORD_BIG_ENDIAN = (5)
    REG_LINK = (6)
    REG_MULTI_SZ = (7)
    REG_RESOURCE_LIST = (8)
    REG_FULL_RESOURCE_DESCRIPTOR = (9)
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Public Const ERROR_SUCCESS = 0
Public Const ERROR_REG = 1

Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))


Function RegQueryStringValue(ByVal hKey As ERegistryClassConstants, ByVal strValueName As String)
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim ERROR_SUCCESS
    On Error GoTo 0
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function
Public Function GETSTRING(hKey As ERegistryClassConstants, strpath As String, strvalue As String)
Dim keyhand&
Dim datatype&
R = RegOpenKey(hKey, strpath, keyhand&)
GETSTRING = RegQueryStringValue(keyhand&, strvalue)
R = RegCloseKey(keyhand&)
End Function
'Function StripTerminator(ByVal strString As String) As String
'    Dim intZeroPos As Integer
'
'    intZeroPos = InStr(strString, Chr$(0))
'    If intZeroPos > 0 Then
'        StripTerminator = Left$(strString, intZeroPos - 1)
'    Else
'        StripTerminator = strString
'    End If
'End Function

Public Sub SAVESTRING(hKey As ERegistryClassConstants, strpath As String, strvalue As String, strdata As String)
Dim keyhand&
R = RegCreateKey(hKey, strpath, keyhand&)
R = RegSetValueEx(keyhand&, strvalue, 0, REG_SZ, ByVal strdata, Len(strdata))
R = RegCloseKey(keyhand&)
End Sub
Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop
End Sub
Public Function Short_Name(Long_Path As String) As String

       '     'Returns short path of the passed long path
       Dim Short_Path As String
       Dim Answer As Long
       Short_Path = Space(250)
       Answer = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))

              If Answer Then
                     Short_Name = Left$(Short_Path, Answer)
              End If

End Function

Public Sub DeleteRegistryValue(ByVal hKey As ERegistryClassConstants, ByVal strpath As String, ByVal strvalue As String)
  Dim lRtn As Long
  Dim retKey As Long
  
  lRtn = RegOpenKeyEx(hKey, strpath, 0&, KEY_ALL_ACCESS, retKey)
  
  If lRtn <> ERROR_SUCCESS Then
    RegCloseKey (retKey)
    Exit Sub
  End If
  
  lRtn = RegDeleteValue(retKey, strvalue)
  
  If lRtn <> ERROR_SUCCESS Then
    RegCloseKey (retKey)
    Exit Sub
  End If
  
  RegCloseKey (retKey)
  
End Sub

