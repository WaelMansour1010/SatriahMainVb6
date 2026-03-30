VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LngCurrentOpenKey As Long

Private Declare Function RegCreateKey _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       phkResult As Long) As Long   'Create Key in the Registry

Private Declare Function RegSetValueEx _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal reserved As Long, _
                                        ByVal dwType As Long, _
                                        lpData As Any, _
                                        ByVal cbData As Long) As Long      'Enter a value in the Registry
        
Private Declare Function RegCloseKey _
                Lib "advapi32.dll" (ByVal hKey As Long) As Long     'Close the Registry

Private Declare Function RegOpenKeyEx _
                Lib "advapi32" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       ByRef phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
                Lib "advapi32" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          ByRef lpType As Long, _
                                          ByVal lpData As String, _
                                          ByRef lpcbData As Long) As Long
    
Private Declare Function RegSetValueExString _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal reserved As Long, _
                                        ByVal dwType As Long, _
                                        ByVal lpValue As String, _
                                        ByVal cbData As Long) As Long

Private Declare Function RegCreateKeyEx _
                Lib "advapi32" _
                Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                         ByVal lpSubKey As String, _
                                         ByVal reserved As Long, _
                                         ByVal lpClass As String, _
                                         ByVal dwOptions As Long, _
                                         ByVal samDesired As Long, _
                                         ByVal lpSecurityAttributes As Long, _
                                         phkResult As Long, _
                                         lpdwDisposition As Long) As Long

Public Enum RegSections
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001  'Hex For HKEY_CURRENT_USER
    HKEY_LOCAL_MACHINE = &H80000002 'Hex for HKEY_LOCAL_MACHINE
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Private Const REG_SZ = 1                 'Constant for a string variable type

Private Const REG_DWORD = 4              'Constant for a long variable type

'Private Const HKEY_CURRENT_USER = &H80000001   'Hex For HKEY_CURRENT_USER
'Private Const HKEY_LOCAL_MACHINE = &H80000002  'Hex for HKEY_LOCAL_MACHINE
' Reg Key Security Options...
Private Const READ_CONTROL = &H20000

Private Const KEY_QUERY_VALUE = &H1

Private Const KEY_SET_VALUE = &H2

Private Const KEY_CREATE_SUB_KEY = &H4

Private Const KEY_ENUMERATE_SUB_KEYS = &H8

Private Const KEY_NOTIFY = &H10

Private Const KEY_CREATE_LINK = &H20

Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Private Const REG_OPTION_NON_VOLATILE = 0

Private Const ERROR_SUCCESS = 0

Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"

Private Const gREGVALSYSINFOLOC = "MSINFO"

Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"

Private Const gREGVALSYSINFO = "PATH"

Public Function RegDoesKeyExist(dwKeyType As RegSections, _
                                sRegPath As String) As Boolean
    Dim hKey As Long
    hKey = RegKeyOpen(dwKeyType, sRegPath)
    RegDoesKeyExist = (hKey <> 0)
    RegCloseKey hKey
End Function

Public Function RegCreateNewKey(ByVal dwKeyType As RegSections, _
                                ByVal sNewKeyName As String) As Long

    Dim hKey As Long
    Dim result As Long
    Call RegCreateKeyEx(dwKeyType, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, result)
    
    RegCreateNewKey = hKey
End Function

Public Function RegKeyOpen(dwKeyType As RegSections, _
                           sKeyPath As String) As Long
    Dim hKey As Long
    Dim dwOptions As Long
    dwOptions = 0&

    If RegOpenKeyEx(dwKeyType, sKeyPath, dwOptions, KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
        RegKeyOpen = hKey
        LngCurrentOpenKey = hKey
    End If

End Function

Public Function RegWriteStringValue(ByVal sValue, _
                                    ByVal dwDataType, _
                                    ByVal sNewValue) As Long
    Dim success As Long
    Dim dwNewValue As Long
    dwNewValue = Len(sNewValue)

    If dwNewValue > 0 Then
        RegWriteStringValue = RegSetValueExString(LngCurrentOpenKey, sValue, 0&, REG_SZ, sNewValue, dwNewValue)
    End If

End Function

Public Function RegReadKeyValue(KeyRoot As RegSections, _
                                KeyName As String, _
                                SubKeyRef As String, _
                                ByRef KeyVal As String) As Boolean
    'Retrun TRUE if the Function Succed
    'False if Function Faild
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          'Handle Error...

    tmpVal = String$(1024, 0)                               'Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...

        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = tmpVal                                     ' Copy String Value

        Case REG_DWORD                                          ' Double Word Registry Key Data Type

            For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
            Next

            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select

    RegReadKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:
    ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    RegReadKeyValue = False                                 ' Return Failure
    rc = RegCloseKey(hKey)
End Function

Private Sub Class_Terminate()

    If LngCurrentOpenKey <> 0 Then
        RegCloseKey LngCurrentOpenKey
    End If

End Sub
