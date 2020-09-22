Attribute VB_Name = "modWindows"
'-----------------------------------------------------------------------------------
'--------------------SUB FOR CHANGING THE AVAILABLE BUTTONS IN----------------------
'--------------------------------THE WINDOWS SECURITY DIALOG------------------------
'--------------------------------------------BY CHANGING THE REGISTRY BASED---------
'-----------------------------------------------------ON THE SETTING YOU SELECT-----
'-----------------------------------------------------------------------------------

Public Sub WinSecurity(ByVal regSET As regKey, ByVal Enabled As Boolean)

    'Declare the variables
    Dim Command As String
    'Select the key
    Select Case regSET
           Case Logoff
                  Command = "NoLogoff"
           Case Shutdown
                  Command = "NoClose"
           Case ChangePassword
                  Command = "DisableChangePassword"
           Case TaskMgr
                  Command = "DisableTaskMgr"
           Case LockWorkstation
                  Command = "DisabeLockWorkstation"
    End Select
    'Set the value of the keys
    If Command = "NoLogoff" Then Call CreateKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", Command, Not Enabled): GoTo SKIPOUT
    If Command = "NoClose" Then Call CreateKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", Command, Not Enabled): GoTo SKIPOUT
    Debug.Print "GOTOIT"
    Call CreateKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", Command, Not Enabled)

SKIPOUT:

End Sub

'-----------------------------------------------------------------------------------
'-----------------------------THESE ARE THE DECLARATIONS----------------------------
'------------------------------------------WHERE THE API'S ARE DECLARED-------------
'------------------------------------------AND ENUM ARE DECLARED TOO----------------
'-----------------------------------------------------------------------------------

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" _
                (ByVal HKey As Long) _
                As Long
                        
Declare Function RegCreateKey Lib "advapi32.dll" _
                Alias "RegCreateKeyA" _
                (ByVal HKey As Long, _
                ByVal lpSubKey As String, _
                phkResult As Long) _
                As Long
                
Declare Function RegOpenKeyEx Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" _
                (ByVal HKey As Long, _
                ByVal lpSubKey As String, _
                ByVal ulOptions As Long, _
                ByVal samDesired As Long, _
                phkResult As Long) _
                As Long

Enum regKey

    Logoff = 0
    Shutdown = 1
    ChangePassword = 2
    TaskMgr = 3
    LockWorkstation = 4

End Enum

Enum RegistryErrorCodes

    ERROR_ACCESS_DENIED = 5&
    ERROR_INVALID_PARAMETER = 87
    ERROR_MORE_DATA = 234
    ERROR_NO_MORE_ITEMS = 259
    ERROR_SUCCESS = 0&

End Enum

Enum RegistryLongTypes
    REG_BINARY = 3              ' Binary Type
    REG_DWORD = 4               ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
End Enum

Enum RegistryKeyAccess
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_EVENT = &H1
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    READ_CONTROL = &H20000
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_REQUIRED = &HF0000
    SYNCHRONIZE = &H100000
    STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    STANDARD_RIGHTS_READ = (READ_CONTROL)
    STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL + KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
End Enum

Enum RegistryHives
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

'-----------------------------------------------------------------------------------
'---------------------SUBS FOR CREATING REGISTRY------------------------------------
'------------------------------------KEYS, TO MODIFY BUTTONS------------------------
'------------------------------------------IN WINDOWS SECURITY DIALOG---------------
'-----------------------------------------------------------------------------------

Public Sub CreateKey(ByVal EnmHive As Long, ByVal StrSubKey As String, ByVal strValueName As String, ByVal LngData As Long, Optional ByVal EnmType As RegistryLongTypes = REG_DWORD_LITTLE_ENDIAN)

    Dim HKey As Long 'Holds a pointer to the registry key
    'Create the Registry Key
    Call CreateSubKey(EnmHive, StrSubKey)
    'Open the registry key
    HKey = GetSubKeyHandle(EnmHive, StrSubKey, KEY_ALL_ACCESS)
    'Set the registry value
    RegSetValueEx HKey, strValueName, 0, EnmType, LngData, 4
    'Close the registry key
    RegCloseKey HKey

End Sub

Public Sub CreateSubKey(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String)

    Dim HKey As Long 'Holds the handle from the created key.
    'Create the Key'
    RegCreateKey EnmHive, StrSubKey & Chr(0), HKey
    'Close the key
    RegCloseKey HKey

End Sub

Private Function GetSubKeyHandle(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, Optional ByVal EnmAccess As RegistryKeyAccess = KEY_READ) As Long

    Dim HKey As Long 'Holds the handle of the specified key
    Dim RetVal As Long 'Holds the data returned from the registry key
    'Open the registry key
    RetVal = RegOpenKeyEx(EnmHive, StrSubKey, 0, EnmAccess, HKey)
    If RetVal <> ERROR_SUCCESS Then
    'Unable to create key
    HKey = 0
    End If
    GetSubKeyHandle = HKey

End Function
