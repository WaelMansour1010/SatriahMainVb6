Attribute VB_Name = "ModExpire"
Option Explicit

Public Const EnKey = "FPR7IGVY7L3RIQQUJ3LV4RLZIIS9FQC4DL3A0I0S3V8AD14PVEQ2GYDGJZJI60ZU8XA0XL9GOD1X0833WVO775GDJUDQ8MU6J55ZAMY4HGCEH8X45W5YE82V1I2MMH78-Byte-bisegypt"

Public Type FileRec
    
    CustomerName As String * 100
    phone As String * 100
    Mobile As String * 100
    Emial As String * 100
    Address As String * 255
    ComputerID As String * 255
    SerialNumber As String * 255
    ActivateNumber As String * 255
    HardDisk_ID As String * 255
    Processor_ID As String * 255
    MaxNumToRun As Long
    CurRumNumber As Long
    VersionType As Integer
    FristRunDate As Date
    LastRunDate As Date
End Type

Public Enum ExireTypes
    DevelopVersion
    DemoRun
    DemoStop
    Registered
    UnErrorOccured
End Enum

Public RecSave As FileRec

Public RecRead As FileRec

Public RegSaveRec As FileRec

Private Const FLAG_ICC_FORCE_CONNECTION = &H1

Private Declare Function InternetCheckConnection _
                Lib "wininet.dll" _
                Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, _
                                                  ByVal dwFlags As Long, _
                                                  ByVal dwReserved As Long) As Long

Private Const GENERIC_READ = &H80000000

Private Const GENERIC_WRITE = &H40000000

Private Const FILE_SHARE_READ = &H1

Private Const FILE_SHARE_WRITE = &H2

Private Const OPEN_EXISTING = 3

Private Const CREATE_NEW = 1

Private Const INVALID_HANDLE_VALUE = -1

Private Const VER_PLATFORM_WIN32_NT = 2

Private Const IDENTIFY_BUFFER_SIZE = 512

Private Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

'GETVERSIONOUTPARAMS contains the data returned
'From the Get Driver Version function
Private Type GETVERSIONOUTPARAMS
    bVersion       As Byte 'Binary driver version.
    bRevision      As Byte 'Binary driver revision
    bReserved      As Byte 'Not used
    bIDEDeviceMap  As Byte 'Bit map of IDE devices
    fCapabilities  As Long 'Bit mask of driver capabilities
    dwReserved(3)  As Long 'For future use
End Type

'IDE registers
Private Type IDEREGS
    bFeaturesReg     As Byte 'Used for specifying SMART "commands"
    bSectorCountReg  As Byte 'IDE sector count register
    bSectorNumberReg As Byte 'IDE sector number register
    bCylLowReg       As Byte 'IDE low order cylinder value
    bCylHighReg      As Byte 'IDE high order cylinder value
    bDriveHeadReg    As Byte 'IDE drive/head register
    bCommandReg      As Byte 'Actual IDE command
    bReserved        As Byte 'reserved for future use - must be zero
End Type

'SENDCMDINPARAMS contains the input parameters for the
'Send Command to Drive function
Private Type SENDCMDINPARAMS
    cBufferSize     As Long     'Buffer size in bytes
    irDriveRegs     As IDEREGS  'Structure with drive register values.
    bDriveNumber    As Byte     'Physical drive number to send command to (0,1,2,3).
    bReserved(2)    As Byte     'Bytes reserved
    dwReserved(3)   As Long     'DWORDS reserved
    bBuffer()      As Byte      'Input buffer.
End Type

'Valid values for the bCommandReg member of IDEREGS.
Private Const IDE_ID_FUNCTION = &HEC            'Returns ID sector for ATA.

Private Const IDE_EXECUTE_SMART_FUNCTION = &HB0 'Performs SMART cmd.
'Requires valid bFeaturesReg,
'bCylLowReg, and bCylHighReg

'Cylinder register values required when issuing SMART command
Private Const SMART_CYL_LOW = &H4F

Private Const SMART_CYL_HI = &HC2

'Status returned From driver
Private Type DRIVERSTATUS
    bDriverError  As Byte          'Error code From driver, or 0 if no error
    bIDEStatus    As Byte          'Contents of IDE Error register
    'Only valid when bDriverError is SMART_IDE_ERROR
    bReserved(1)  As Byte
    dwReserved(1) As Long
End Type

Private Type IDSECTOR
    wGenConfig                 As Integer
    wNumCyls                   As Integer
    wReserved                  As Integer
    wNumHeads                  As Integer
    wBytesPerTrack             As Integer
    wBytesPerSector            As Integer
    wSectorsPerTrack           As Integer
    wVendorUnique(2)           As Integer
    sSerialNumber(19)          As Byte
    wBufferType                As Integer
    wBufferSize                As Integer
    wECCSize                   As Integer
    sFirmwareRev(7)            As Byte
    sModelNumber(39)           As Byte
    wMoreVendorUnique          As Integer
    wDoubleWordIO              As Integer
    wCapabilities              As Integer
    wReserved1                 As Integer
    wPIOTiming                 As Integer
    wDMATiming                 As Integer
    wBS                        As Integer
    wNumCurrentCyls            As Integer
    wNumCurrentHeads           As Integer
    wNumCurrentSectorsPerTrack As Integer
    ulCurrentSectorCapacity    As Long
    wMultSectorStuff           As Integer
    ulTotalAddressableSectors  As Long
    wSingleWordDMA             As Integer
    wMultiWordDMA              As Integer
    bReserved(127)             As Byte
End Type

'Structure returned by SMART IOCTL commands
Private Type SENDCMDOUTPARAMS
    cBufferSize   As Long         'Size of Buffer in bytes
    DRIVERSTATUS  As DRIVERSTATUS 'Driver status structure
    bBuffer()    As Byte          'Buffer of arbitrary length for data read From drive
End Type

'Vendor specific feature register defines
'for SMART "sub commands"
Private Const SMART_ENABLE_SMART_OPERATIONS = &HD8

'Status Flags Values
Public Enum STATUS_FLAGS
    PRE_FAILURE_WARRANTY = &H1
    ON_LINE_COLLECTION = &H2
    PERFORMANCE_ATTRIBUTE = &H4
    ERROR_RATE_ATTRIBUTE = &H8
    EVENT_COUNT_ATTRIBUTE = &H10
    SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

'IOCTL commands
Private Const DFP_GET_VERSION = &H74080

Private Const DFP_SEND_DRIVE_COMMAND = &H7C084

Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Private Type ATTR_DATA
    AttrID As Byte
    AttrName As String
    AttrValue As Byte
    ThresholdValue As Byte
    WorstValue As Byte
    StatusFlags As STATUS_FLAGS
End Type

Private Type DRIVE_INFO
    bDriveType As Byte
    SerialNumber As String
    Model As String
    FirmWare As String
    Cilinders As Long
    Heads As Long
    SecPerTrack As Long
    BytesPerSector As Long
    BytesperTrack As Long
    NumAttributes As Byte
    Attributes() As ATTR_DATA
End Type

Private Enum IDE_DRIVE_NUMBER
    PRIMARY_MASTER
    PRIMARY_SLAVE
    SECONDARY_MASTER
    SECONDARY_SLAVE
    TERTIARY_MASTER
    TERTIARY_SLAVE
    QUARTIARY_MASTER
    QUARTIARY_SLAVE
End Enum

Private Declare Function CreateFile _
                Lib "kernel32" _
                Alias "CreateFileA" (ByVal lpFileName As String, _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal dwShareMode As Long, _
                                     lpSecurityAttributes As Any, _
                                     ByVal dwCreationDisposition As Long, _
                                     ByVal dwFlagsAndAttributes As Long, _
                                     ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle _
                Lib "kernel32" (ByVal hObject As Long) As Long
  
Private Declare Function DeviceIoControl _
                Lib "kernel32" (ByVal hDevice As Long, _
                                ByVal dwIoControlCode As Long, _
                                lpInBuffer As Any, _
                                ByVal nInBufferSize As Long, _
                                lpOutBuffer As Any, _
                                ByVal nOutBufferSize As Long, _
                                lpBytesReturned As Long, _
                                lpOverlapped As Any) As Long
  
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (hpvDest As Any, _
                                       hpvSource As Any, _
                                       ByVal cbCopy As Long)
  
Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx _
                Lib "kernel32" _
                Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Function GetDriveInfo(drvNumber As IDE_DRIVE_NUMBER) As DRIVE_INFO
  
    Dim hDrive As Long
    Dim DI As DRIVE_INFO
   
    hDrive = SmartOpen(drvNumber)
   
    If hDrive <> INVALID_HANDLE_VALUE Then
   
        If SmartGetVersion(hDrive) = True Then
      
            With DI
                .bDriveType = 0
                .NumAttributes = 0
                ReDim .Attributes(0)
                .bDriveType = 1
            End With
         
            If SmartCheckEnabled(hDrive, drvNumber) Then
            
                If IdentifyDrive(hDrive, IDE_ID_FUNCTION, drvNumber, DI) = True Then
         
                    GetDriveInfo = DI
               
                End If   'IdentifyDrive
            End If   'SmartCheckEnabled
        End If   'SmartGetVersion
    End If   'hDrive <> INVALID_HANDLE_VALUE
   
    CloseHandle hDrive
   
End Function

Private Function IdentifyDrive(ByVal hDrive As Long, _
                               ByVal IDCmd As Byte, _
                               ByVal drvNumber As IDE_DRIVE_NUMBER, _
                               DI As DRIVE_INFO) As Boolean
    
    'Function: Send an IDENTIFY command to the drive
    'drvNumber = 0-3
    'IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
    Dim SCIP As SENDCMDINPARAMS
    Dim IDSEC As IDSECTOR
    Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
    Dim cbBytesReturned As Long
   
    With SCIP
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        .bDriveNumber = CByte(drvNumber)
        
        With .irDriveRegs
            .bFeaturesReg = 0
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = 0
            .bCylHighReg = 0
            .bDriveHeadReg = &HA0 'compute the drive number

            If Not IsWinNT4Plus Then
                .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If

            'the command can either be IDE
            'identify or ATAPI identify.
            .bCommandReg = CByte(IDCmd)
        End With
    End With
   
    If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&) Then
                      
        CopyMemory IDSEC, bArrOut(16), Len(IDSEC)

        DI.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
        DI.SerialNumber = StrConv(SwapBytes(IDSEC.sSerialNumber), vbUnicode)
      
        IdentifyDrive = True
      
    End If
    
End Function

Private Function IsWinNT4Plus() As Boolean

    'returns True if running Windows NT4 or later
    Dim osv As OSVERSIONINFO

    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
   
        IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor >= 4)
 
    End If

End Function

Private Function SmartCheckEnabled(ByVal hDrive As Long, _
                                   drvNumber As IDE_DRIVE_NUMBER) As Boolean
   
    'SmartCheckEnabled - Check if SMART enable
    'FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
    'bDriveNum = 0-3
    Dim SCIP As SENDCMDINPARAMS
    Dim SCOP As SENDCMDOUTPARAMS
    Dim cbBytesReturned As Long
   
    With SCIP
   
        .cBufferSize = 0
      
        With .irDriveRegs
            .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI

            .bDriveHeadReg = &HA0

            If Not IsWinNT4Plus Then
                .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If

            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
           
        End With
       
        .bDriveNumber = drvNumber
       
    End With
   
    SmartCheckEnabled = DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, SCIP, Len(SCIP) - 4, SCOP, Len(SCOP) - 4, cbBytesReturned, ByVal 0&)
End Function

Private Function SmartGetVersion(ByVal hDrive As Long) As Boolean
   
    Dim cbBytesReturned As Long
    Dim GVOP As GETVERSIONOUTPARAMS
   
    SmartGetVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, ByVal 0&, 0, GVOP, Len(GVOP), cbBytesReturned, ByVal 0&)
   
End Function

Private Function SmartOpen(drvNumber As IDE_DRIVE_NUMBER) As Long

    'Open SMART to allow DeviceIoControl
    'communications and return SMART handle

    If IsWinNT4Plus() Then
      
        SmartOpen = CreateFile("\\.\PhysicalDrive" & CStr(drvNumber), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&)

    Else
      
        SmartOpen = CreateFile("\\.\SMARTVSD", 0&, 0&, ByVal 0&, CREATE_NEW, 0&, 0&)
    End If
   
End Function

Private Function SwapBytes(b() As Byte) As Byte()
   
    'Note: VB4-32 and VB5 do not support the
    'return of arrays From a function. For
    'developers using these VB versions there
    'are two workarounds to this restriction:
    '
    '1) Change the return data type ( As Byte() )
    '   to As Variant (no brackets). No change
    '   to the calling code is required.
    '
    '2) Change the function to a sub, remove
    '   the last line of code (SwapBytes = b()),
    '   and take advantage of the fact the
    '   original byte array is being passed
    '   to the function ByRef, therefore any
    '   changes made to the passed data are
    '   actually being made to the original data.
    '   With this workaround the calling code
    '   also requires modification:
    '
    '      di.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
    '
    '   ... to ...
    '
    '      Call SwapBytes(IDSEC.sModelNumber)
    '      di.Model = StrConv(IDSEC.sModelNumber, vbUnicode)
   
    Dim bTemp As Byte
    Dim cnt As Long

    For cnt = LBound(b) To UBound(b) Step 2
        bTemp = b(cnt)
        b(cnt) = b(cnt + 1)
        b(cnt + 1) = bTemp
    Next cnt
      
    SwapBytes = b()
      
End Function

Public Function IsInternetConnected(Optional BolShowMsg As Boolean) As Boolean
    Dim BolConnected As Boolean
    Dim Msg As String
    On Error GoTo ErrTrap
    '
    'If SystemOptions.SysConnectionType = ConnecLocal Then
    '    IsInternetConnected = True
    '    Exit Function
    'End If

    If InternetCheckConnection("http://" & SystemOptions.SysServerIP & "/", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        BolConnected = False
    Else
        BolConnected = True
    End If

    IsInternetConnected = (BolConnected)

    If BolShowMsg = True Then
        If IsInternetConnected = False Then
            Msg = "·«ÌÊÃœ ≈ ’«· »«·√‰ —‰ "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
        End If
    End If

    Exit Function
ErrTrap:
    IsInternetConnected = False
End Function

Public Function CheckExpireation() As ExireTypes
    Dim Msg As String
    Dim StrHardDiskData As String
    Dim StrProcessorData As String
    Dim RetVal As Long
    Dim cRegistery As New ClsRegistry
    On Error GoTo hErr
    WriteInLogFile "Check The Path oh the RegFile"

    If Dir(SystemOptions.SysRegFilePath) = "" Then
        WriteInLogFile "Check the Registery"

        If CheckFromRegistery = True Then
            WriteInLogFile "Creat the Key in Registery"
            cRegistery.RegCreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\SystemOperations"
            cRegistery.RegCreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\SystemOperations\S_A"
            WriteInLogFile "Creat New RegFile"
            CreateNewRegFile
            WriteInLogFile "RegFile Created"
        End If
    End If

    If Dir(SystemOptions.SysRegFilePath) = "" Then
        Msg = "⁄›Ê« „·› Õ„«Ì… «·»—‰«„Ã €Ì— „ÊÃÊœ"
        Msg = Msg & Chr(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã "
        Msg = Msg & Chr(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckExpireation = DemoStop
        Exit Function
    End If

    WriteInLogFile "Load the Regfile"

    If LoadRegFile = False Then
        Msg = "⁄›Ê« „·› Õ„«Ì… «·»—‰«„Ã €Ì— „·«∆„"
        Msg = Msg & Chr(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã "
        Msg = Msg & Chr(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckExpireation = UnErrorOccured
        Exit Function
    End If

    WriteInLogFile "Check RecRead.VersionType"
    WriteInLogFile "RecRead.VersionType=" & RecRead.VersionType

    If RecRead.VersionType = 0 Then

        'This is a Demo Version
        'In The Demo Version
        'We check that the
        If RecRead.CurRumNumber >= RecRead.MaxNumToRun Then
            '⁄›Ê« ·ﬁœ «‰ Â  «·‰”Œ… «· Ã—»Ì… „‰ «·»—‰«„Ã
            Msg = "⁄›Ê« ·ﬁœ «‰ Â  «·‰”Œ… «· Ã—»Ì… „‰ «·»—‰«„Ã"
            Msg = Msg & Chr(13) & "RecRead.MaxNumToRun=" & RecRead.MaxNumToRun
            Msg = Msg & Chr(13) & "RecRead.CurRumNumber=" & RecRead.CurRumNumber
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CheckExpireation = DemoStop
        Else
            CheckExpireation = DemoRun
        End If

    ElseIf RecRead.VersionType = 1 Then
        'This Is as Registered Verison
        WriteInLogFile "GetHardDiskData"
        StrHardDiskData = GetHardDiskData(False)
        WriteInLogFile "StrHardDiskData"
        WriteInLogFile "GetProcessorData"
        StrProcessorData = GetProcessorData(False)
        WriteInLogFile StrProcessorData

        If (Trim(RecRead.HardDisk_ID) <> StrHardDiskData) And (Trim(RecRead.Processor_ID) <> StrProcessorData) Then
            'Msg = "„·› «·Õ„«Ì… «·„ÊÃÊœ.."
            'Msg = Msg & Chr(13) & "·«Ì‰«”» Â–« «·ÃÂ«“"
            'Msg = Msg & Chr(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
            Msg = "„·› «·Õ„«Ì… «·„Õœœ €Ì— ’«·Õ ··⁄„· „⁄ Â–« «·ÃÂ«“"
            MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CheckExpireation = UnErrorOccured
            Exit Function
        Else
            CheckExpireation = Registered
        End If
    End If

    SystemOptions.SysRunNumber = RecRead.CurRumNumber
    RecRead.CurRumNumber = RecRead.CurRumNumber + 1
    RetVal = cRegistery.RegKeyOpen(HKEY_LOCAL_MACHINE, "SOFTWARE\SystemOperations\S_A")
    cRegistery.RegWriteStringValue "ex", 1, RecRead.CurRumNumber

    SaveInRegFile RecRead
    Exit Function
hErr:
    Msg = Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    Msg = Msg & Chr(13) & "Function CheckExpireation"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    CheckExpireation = UnErrorOccured
End Function

Private Function CheckFromRegistery() As Boolean
    Dim cRegistery As ClsRegistry
    Dim RetVal As Long

    On Error GoTo ErrTrap
    Set cRegistery = New ClsRegistry

    If cRegistery.RegDoesKeyExist(HKEY_LOCAL_MACHINE, "SOFTWARE\SystemOperations\S_A") = False Then
        'this is new system copy
        CheckFromRegistery = True
    Else
        'the system is installed before
        'and the user delete the RegFile
        CheckFromRegistery = False
    End If

    Exit Function
ErrTrap:
    CheckFromRegistery = False
End Function

Public Function CreateNewRegFile() As Boolean

    Dim IntFreeFile As Integer
    Dim StrFilePath As String
    Dim StrNewPath As String
    Dim StrNewEncrPath As String
    Dim Msg As String
    Dim Encryptor As ImpulseEncryption

    On Error GoTo ErrTrap
    IntFreeFile = FreeFile
    StrFilePath = App.path & "\TempRegFile.txt" 'SystemOptions.SysRegTempFilePath

    If Dir(StrFilePath, vbNormal) <> "" Then
        Kill StrFilePath
    End If

    RecSave.CustomerName = ""
    RecSave.phone = ""
    RecSave.Mobile = ""
    RecSave.Emial = ""
    RecSave.Address = ""
    RecSave.ComputerID = ""
    RecSave.SerialNumber = ""
    RecSave.ActivateNumber = ""
    RecSave.HardDisk_ID = ""
    RecSave.Processor_ID = ""
    RecSave.MaxNumToRun = 50
    RecSave.CurRumNumber = 0
    RecSave.VersionType = 0
    RecSave.FristRunDate = Date
    RecSave.LastRunDate = Date

    Open StrFilePath For Random As #IntFreeFile Len = Len(RecSave)
    Put #IntFreeFile, 1, RecSave
    Close #IntFreeFile
    Set Encryptor = New ImpulseEncryption
    Encryptor.EncryptionKey = EnKey
    StrNewEncrPath = App.path & "\RegFile.txt" 'SystemOptions.SysRegFilePath

    If Dir(StrNewEncrPath) <> "" Then
        Kill StrNewEncrPath
    End If

    Encryptor.EncryptFile StrFilePath, StrNewEncrPath

    If Dir(StrFilePath) <> "" Then
        Kill StrFilePath
    End If

    Msg = "„·ÕÊŸ… Â«„…:-"
    Msg = Msg & Chr(13) & " „ ≈‰‘«¡ „·› «·Õ„«Ì… «·Œ«’… »«·»—‰«„Ã ⁄·Ï «·„”«— "
    Msg = Msg & Chr(13) & ""
    Msg = Msg & Chr(13) & StrNewEncrPath
    Msg = Msg & Chr(13) & ""
    Msg = Msg & Chr(13) & "≈” Œœ„ Â–« «·„·› ·«Õﬁ« ›Ï  ‰‘Ìÿ «·»—‰«„Ã œÊ‰"
    Msg = Msg & Chr(13) & "«·Õ«Ã… ≈·Ï «·√ ’«· »«·√‰ —‰  «Ê «·Â« ›"
    'MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

    CreateNewRegFile = True
    Exit Function
ErrTrap:
    Msg = "CreateNewRegFile Fialed"
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    CreateNewRegFile = False
End Function

Public Function GetHardDiskData(BolWithCaption As Boolean) As String
    Dim DI As DRIVE_INFO
    Dim drvNumber As Long
    Dim Temp As String
    Dim xx As String

    For drvNumber = PRIMARY_MASTER To QUARTIARY_SLAVE
        DI = GetDriveInfo(drvNumber)

        With DI

            Select Case .bDriveType

                Case 0

                Case 1
                    xx = Trim(.SerialNumber)
                    xx = Replace(xx, Chr(0), vbNullString, , , vbBinaryCompare)
                    Temp = Temp & Trim(xx) & ";"
                    Exit For

                Case 2

                Case Else
            End Select

        End With

    Next

    Temp = Mid(Temp, 1, Len(Temp) - 1)
    GetHardDiskData = Temp
End Function

Public Function LoadRegFile() As Boolean
    Dim Encryptor As ImpulseEncryption
    Dim StrDecrPath As String
    Dim IntFreeFile As Integer
    Dim StrAllFile  As String
    Dim VarTemp As Variant
    Dim StrFilePath As String
    Dim Msg As String

    On Error GoTo ErrTrap

    If Dir(Trim(SystemOptions.SysRegFilePath)) = "" Then
        Exit Function
    End If

    Set Encryptor = New ImpulseEncryption

    Encryptor.EncryptionKey = EnKey
    StrDecrPath = SystemOptions.SysRegTempFilePath

    If Dir(StrDecrPath, vbNormal) <> "" Then
        Kill StrDecrPath
    End If

    Encryptor.DecryptFile SystemOptions.SysRegFilePath, StrDecrPath

    IntFreeFile = FreeFile

    Open StrDecrPath For Random As #IntFreeFile Len = Len(RecRead)
    Get #IntFreeFile, 1, RecRead
    Close #IntFreeFile

    If Dir(StrDecrPath, vbNormal) <> "" Then
        Kill StrDecrPath
    End If

    WriteInLogFile "LoadRegFile = True"
    LoadRegFile = True
    Exit Function
ErrTrap:
    Msg = "Error IN LoadRegFile"
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    LoadRegFile = False
    WriteInLogFile "Error IN LoadRegFile"
End Function

Public Function GetProcessorData(BolWithCaption As Boolean) As String
    Dim objNameSpace As SWbemServices, ObjCPUSet As SWbemObjectSet
    Dim objCpu As SWbemObject
    Set objNameSpace = GetObject("winmgmts:")
    Set ObjCPUSet = objNameSpace.InstancesOf("Win32_Processor")
    Dim ss As ImpulseGlobals.ImpulseProcessor
    Set ss = New ImpulseGlobals.ImpulseProcessor

    For Each objCpu In ObjCPUSet

        If BolWithCaption = True Then
            GetProcessorData = "»Ì«‰«  «·»—Ê””Ê—" & Chr(13) & Chr(10) & objCpu.GetObjectText_
        Else
            'GetProcessorData = objCpu.SerialNumber
            GetProcessorData = objCpu.ProcessorId
        End If

        Exit For
    Next

    Set objCpu = Nothing
    Set ObjCPUSet = Nothing
    Set objNameSpace = Nothing
End Function

Public Function SaveInRegFile(XSavedRec As FileRec)
    Dim Encryptor As ImpulseEncryption 'salim2
    Dim StrDecrPath As String
    Dim IntFreeFile As Integer
    Dim StrAllFile  As String
    Dim VarTemp As Variant
    Dim StrFilePath As String
    Dim Msg As String

    On Error GoTo ErrTrap

    If Dir(Trim(SystemOptions.SysRegFilePath)) = "" Then
        Exit Function
    End If

    Set Encryptor = New ImpulseEncryption

    Encryptor.EncryptionKey = EnKey
    StrDecrPath = SystemOptions.SysRegTempFilePath

    If Dir(StrDecrPath, vbNormal) <> "" Then
        Kill StrDecrPath
    End If

    'Decrypt the file to enable to write the data
    Encryptor.DecryptFile SystemOptions.SysRegFilePath, StrDecrPath

    IntFreeFile = FreeFile
    'Write the data
    Open StrDecrPath For Random As #IntFreeFile Len = Len(XSavedRec)
    Put #IntFreeFile, 1, XSavedRec
    Close #IntFreeFile

    If Dir(SystemOptions.SysRegFilePath, vbNormal) <> "" Then
        Kill SystemOptions.SysRegFilePath
    End If

    Encryptor.EncryptFile StrDecrPath, SystemOptions.SysRegFilePath

    If Dir(StrDecrPath, vbNormal) <> "" Then
        Kill StrDecrPath
    End If

    SaveInRegFile = True
    Exit Function
ErrTrap:
    Msg = "Error In SaveInRegFile"
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    SaveInRegFile = False

End Function
