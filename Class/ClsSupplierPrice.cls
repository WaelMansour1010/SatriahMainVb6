VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSupplierPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const CBS As String = "CryptBinaryToString"
Private Const CSB As String = "CryptStringToBinary"
Private Const CRYPT_STRING_BASE64 As Long = 1

Private sBase64Buf As String
Private m_bData() As Byte

Private Declare Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "Crypt32.dll" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

'================================
'EVENTS
'================================
Public Event Error(ByVal Number As Long, Description As String, ByVal Source As String)

Public Property Get bBuffer() As Byte()
    bBuffer = m_bData
End Property

Public Property Let bBuffer(bNewValue() As Byte)
    m_bData = bNewValue
End Property

Public Property Get Base64Buf() As String
    Base64Buf = sBase64Buf
End Property

Public Property Let Base64Buf(sNewValue As String)
    sBase64Buf = sNewValue
End Property

Public Sub Base64Decode()
    Const Routine As String = "Base64.Base64Decode"
    Const CRYPT_STRING_BASE64 As Long = 1
    Const CRYPT_STRING_NOCRLF As Long = &H40000000
    Dim bTmp() As Byte
    Dim lLen As Long
    Dim dwActualUsed As Long
    'Get output buffer length
    If CryptStringToBinary(StrPtr(sBase64Buf), Len(sBase64Buf), CRYPT_STRING_BASE64, StrPtr(vbNullString), lLen, 0&, dwActualUsed) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSB, Routine)
        GoTo ReleaseHandles
    End If
    'Convert Base64 to binary.
    ReDim bTmp(lLen - 1)
    If CryptStringToBinary(StrPtr(sBase64Buf), Len(sBase64Buf), CRYPT_STRING_BASE64, VarPtr(bTmp(0)), lLen, 0&, dwActualUsed) = 0 Then
        RaiseEvent Error(Err.LastDllError, CSB, Routine)
        GoTo ReleaseHandles
    Else
        m_bData = bTmp
    End If
ReleaseHandles:
End Sub

Public Sub Base64Encode()
    Const Routine As String = "Base64.Base64Encode"
    Dim lLen As Long
    'Determine Base64 output String length required.
    If CryptBinaryToString(m_bData(0), UBound(m_bData) + 1, CRYPT_STRING_BASE64, StrPtr(vbNullString), lLen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CBS, Routine)
        GoTo ReleaseHandles
    End If
    'Convert binary to Base64.
    sBase64Buf = String$(lLen - 1, CHR$(0))
    If CryptBinaryToString(m_bData(0), UBound(m_bData) + 1, CRYPT_STRING_BASE64, StrPtr(sBase64Buf), lLen) = 0 Then
        RaiseEvent Error(Err.LastDllError, CBS, Routine)
        GoTo ReleaseHandles
    End If
ReleaseHandles:
End Sub

Public Property Get sBuffer() As String
    sBuffer = ByteToStr(m_bData)
End Property

Public Property Let sBuffer(sNewValue As String)
    Dim bTmp() As Byte
    bTmp = StrToByte(sNewValue)
    m_bData = bTmp
End Property

Public Property Get uBuffer() As String
    uBuffer = ByteToUni(m_bData)
End Property

Public Property Let uBuffer(sNewValue As String)
    Dim bTmp() As Byte
    bTmp = sNewValue
    uBuffer = bTmp
End Property






'
'
'Public Function SupplierPrice(SupplierID As Integer)
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As New CRAXDRT.Report
'    Dim StrSQL As String
'    Dim StrRS As New ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'    On Error GoTo ErrTrap
'
'    If SupplierID = Null Then
'        Exit Function
'    End If
'
'    StrSQL = "select * From Supplier_Price where CusID=" & SupplierID
'    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
'
'    If Dir(App.path & "\Reports\SupplierPrice.rpt") = "" Then
'        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
'        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
'        Msg = "\Reports\SupplierPrice.rpt"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Exit Function
'    Else
'        Screen.MousePointer = vbArrowHourglass
'        Set xReport = xApp.OpenReport(App.path & "\Reports\SupplierPrice.rpt")
'        xReport.Database.SetDataSource rs
'        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
'        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
'        xReport.ParameterFields(3).AddCurrentValue user_name
'        Screen.MousePointer = vbDefault
'    End If
'
'    Set FrmReport = New FrmReportViewer
'    FrmReport.CRViewer.ReportSource = xReport
'    FrmReport.CRViewer.viewReport
'    FrmReport.show
'    Screen.MousePointer = vbDefault
'    SendKeys "{RIGHT}"
'    Exit Function
'ErrTrap:
'    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
'    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    Screen.MousePointer = vbDefault
'End Function
'

