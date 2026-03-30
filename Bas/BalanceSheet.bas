Attribute VB_Name = "BalanceSheet"



' $Id: basQRcode.bas
'/**
' The VBA/VB6 interface to diQRcode.dll.
'
' @author dai
' @version 4.0.0
'**/
' $Date: 2021-06-16 10:07:00 $

'*********************** COPYRIGHT NOTICE***********************
' Copyright (c) 2021 DI Management Services Pty Limited.
' <www.di-mgt.com.au> <www.cryptosys.net>. All rights reserved.
' This code may only be used by licensed users and in
' accordance with the licence conditions.
' The latest version of diQRcode and a licence
' may be obtained from <https://www.cryptosys.net/qrcode>.
' This copyright notice must always be left intact.
'******************** END OF COPYRIGHT NOTICE*******************


Option Explicit
Option Base 0

'//////////////////////////////////////////////////
' ECC LEVELS
Public Const QRCODE_ECC_M As Long = 0  ' Default
Public Const QRCODE_ECC_L As Long = 1
Public Const QRCODE_ECC_Q As Long = 2
Public Const QRCODE_ECC_H As Long = 3
' OPTION FLAGS
Public Const QRCODE_ESCAPED As Long = &H1000&
Public Const QRCODE_BASE64  As Long = &H2000&
Public Const QRCODE_NO_NAMECHECK As Long = &H10000
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'/**
' Image type
'**/
Public Enum ImageType
    '/**
    ' Output image as a GIF file (default)
    '**/
    GIF = 0
    '/**
    ' Output image as an SVG file
    '**/
    SVG = &H10
End Enum

'/**
' Error-correction code
'**/
Public Enum Ecc
    '/**
    ' EC level M (default)
    '**/
    mNr = 0
    '/**
    ' EC level L
    '**/
    lNr = 1
    '/**
    ' EC level Q
    '**/
    QNr = 2
    '/**
    ' EC level H
    '**/
    HNr = 3
End Enum

' Local constants
Private Const OUT_OF_RANGE_ERROR As Long = 11

'#If VBA7 Then
'' Declarations for 64-bit Office
'' (In VB6 these will appear red. Turn off "Auto Syntax Check" in Tools > Options to avoid annoying warnings)
'Private Declare PtrSafe Function QRCODE_CreateImage Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
'Private Declare PtrSafe Function QRCODE_CreateImageFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateImage" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
'Private Declare PtrSafe Function QRCODE_CreateGif Lib "diQRcode.dll" (ByVal strOutputFile As String, ByVal strInput As String, ByVal nPixelsPerModule As Long, ByVal strParams As String, ByVal nOptions As Long) As Long
'Private Declare PtrSafe Function QRCODE_CreateGifFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateGif" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
'Private Declare PtrSafe Function QRCODE_CreatePdf Lib "diQRcode.dll" (ByVal strOutputFile As String, ByVal strInput As String, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
'Private Declare PtrSafe Function QRCODE_CreatePdfFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreatePdf" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
'Public Declare PtrSafe Function QRCODE_Version Lib "diQRcode.dll" () As Long
'Private Declare PtrSafe Function QRCODE_DllInfo Lib "diQRcode.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal nOptions As Long) As Long
'Private Declare PtrSafe Function QRCODE_ErrorLookup Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nErrCode As Long) As Long
'Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
'    ByVal CodePage As Long, _
'    ByVal dwFlags As Long, _
'    ByVal lpWideCharStr As Long, _
'    ByVal cchWideChar As Long, _
'    ByVal lpMultiByteStr As Long, _
'    ByVal cbMultiByte As Long, _
'    ByVal lpDefaultChar As Long, _
'    ByVal lpUsedDefaultChar As Long _
'    ) As Long
'#Else
' Declarations for VB6 and 32-bit Office
Private Declare Function QRCODE_CreateImage Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreateImageFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateImage" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreateGif Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreateGifFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateGif" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreatePdf Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreatePdfFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreatePdf" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_ErrorLookup Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nErrCode As Long) As Long

'/**
' Get version number of core DLL.
' @return Version number as an integer in form `Major*10000 + Minor*100 + Release`. For example, version 2.10.3 would return 21003.
'**/
Public Declare Function QRCODE_Version Lib "diQRcode.dll" () As Long
Private Declare Function QRCODE_DllInfo Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nOptions As Long) As Long

'///////////////////////////////////////////////////////////////
' INTERNAL UTF-8 MAPPING
''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
'#End If

' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
' Ref: [How to convert VBA/VB6 Unicode strings to UTF-8](https://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html)
Private Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

' ------------------
' EXPORTED FUNCTIONS
' ------------------

'/**
' Create an image file of a QR code.
' @param szOutputFile Name of output image file to be created.
' @param szInput Text input to be encoded (ANSI characters only).
' @param imgType Image type (default = GIF)
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param nMargin Size of margin in modules (default = 4 modules)
' @param EccLevel Error correction level (default = Ecc.mNr)
' @param nOptions Option flags. Use the `Or` operator to combine: <br>
' `QRCODE_ESCAPED` to indicate #-escaped sequences in string <br>
' `QRCODE_BASE64` to encode output as base64 text <br>
' `QRCODE_NO_NAMECHECK` do not check filename extension against file type <br>
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'**/
Public Function qrcodeCreateImage(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nOptions As Long = 0) As Long
    Dim i As Long
    Dim opts As Long
    opts = nOptions Or EccLevel Or imgType
    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
    For i = 1 To Len(szInput)
        If AscW(mId(szInput, i, 1)) > &HFF Then
            qrcodeCreateImage = -OUT_OF_RANGE_ERROR
            Exit Function
        End If
    Next
    ' Do the business
    qrcodeCreateImage = QRCODE_CreateImage(szOutputFile, szInput, nPixelsPerModule, nMargin, opts)
End Function

'/**
' Create a GIF file of a QR code (_deprecated_).
' @param szOutputFile Name of output GIF file to be created.
' @param szInput Text input to be encoded (ANSI characters only).
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param szParams Optional parameters.  Set as `"margin=N"` to change the margin to `N` modules (default=4).
' @param nOptions Option flags. Set as `QRCODE_ESCAPED` to indicate #-escaped sequences in string.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
' @deprecated Use {@link qrcodeCreateImage}.
'**/
Public Function qrcodeCreateGif(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional szParams As String = "", Optional nOptions As Long = 0) As Long
    Dim i As Long
    Dim opts As Long
    Const OUT_OF_RANGE_ERROR As Long = 11
    opts = nOptions Or EccLevel
    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
    For i = 1 To Len(szInput)
        If AscW(mId(szInput, i, 1)) > &HFF Then
            qrcodeCreateGif = -OUT_OF_RANGE_ERROR
            Exit Function
        End If
    Next
    ' Do the business
    qrcodeCreateGif = QRCODE_CreateGif(szOutputFile, szInput, nPixelsPerModule, szParams, opts)
End Function

'/**
' Create a PDF file of a QR code.
' @param szOutputFile Name of output PDF file to be created.
' @param szInput Text input to be encoded (ANSI characters only).
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nPageWidth  Width of PDF page in pixels (default = 0 : set width to fit QRcode image).
' @param nPageHeight Height of PDF page in pixels (default = 0 : set height to fit QRcode image).
' @param nX  X-coordinate in pixels of bottom-left of QRcode image (default = 0 : at left side).
' @param nY  Y-coordinate in pixels of bottom-left of QRcode image (default = 0 : at bottom).
' @param nOptions Option flags. Set as `QRCODE_ESCAPED` to indicate #-escaped sequences in string.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'**/
Public Function qrcodeCreatePdf(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nPageWidth As Long = 0, Optional nPageHeight As Long = 0, Optional nX As Long = 0, Optional nY As Long = 0, Optional nOptions As Long = 0) As Long
    Dim i As Long
    Dim opts As Long
    Const OUT_OF_RANGE_ERROR As Long = 11
    opts = nOptions Or EccLevel
    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
    For i = 1 To Len(szInput)
        If AscW(mId(szInput, i, 1)) > &HFF Then
            qrcodeCreatePdf = -OUT_OF_RANGE_ERROR
            Exit Function
        End If
    Next
    ' Do the business
    qrcodeCreatePdf = QRCODE_CreatePdf(szOutputFile, szInput, nPixelsPerModule, nPageWidth, nPageHeight, nX, nY, opts)
End Function

'/**
' Create a image file of a QR code, encoding input in UTF-8.
' @param szOutputFile Name of output image file to be created.
' @param szInput Text input to be encoded.
' @param imgType Image type (default = GIF)
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param nMargin Size of margin in modules (default = 4 modules)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nOptions Option flags. Use the `Or` operator to combine: <br>
' `QRCODE_ESCAPED` to indicate #-escaped sequences in string <br>
' `QRCODE_BASE64` to encode output as base64 text <br>
' `QRCODE_NO_NAMECHECK` do not check filename extension against file type <br>
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'**/
'Public Function qrcodeCreateImageInUtf8(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
'    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nOptions As Long = 0) As Long
'    Dim b() As Byte
'    Dim nLen As Long
'    Dim opts As Long
'    opts = nOptions Or EccLevel Or imgType
'    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
'    b = Utf8BytesFromString(szInput)
'    ' Add an extra NUL byte to the array
'    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
'    ReDim Preserve b(nLen)
'    b(nLen - 1) = 0
'    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
'    qrcodeCreateImageInUtf8 = QRCODE_CreateImageFromBytes(szOutputFile, b(0), nPixelsPerModule, nMargin, opts)
'End Function

Public Function qrcodeCreateImageInUtf8(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nOptions As Long = 0) As Long
    Dim b() As Byte
    Dim nLen As Long
    Dim opts As Long
    opts = nOptions Or EccLevel Or imgType
    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
    b = Utf8BytesFromString(szInput)
    ' Add an extra NUL byte to the array
    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
    ReDim Preserve b(nLen)
    b(nLen - 1) = 0
    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
    qrcodeCreateImageInUtf8 = QRCODE_CreateImageFromBytes(szOutputFile, b(0), nPixelsPerModule, nMargin, opts)
End Function

'/**
' Create a GIF file of a QR code, encoding input in UTF-8 (_deprecated_).
' @param szOutputFile Name of output GIF file to be created.
' @param szInput Text input to be encoded.
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param szParams Optional parameters. Set as `"margin=N"` to change the margin to `N` modules (default=4).
' @param nOptions Option flags. Set as 0 for default.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check).
' @remark Any non-ASCII characters in `szInput` will be encoded in UTF-8 before processing.
' @deprecated Use {@link qrcodeCreateImageInUtf8}.
'**/
Public Function qrcodeCreateGifInUtf8(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional szParams As String = "", Optional nOptions As Long = 0) As Long
    Dim b() As Byte
    Dim nLen As Long
    Dim opts As Long
    opts = nOptions Or EccLevel
    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
    b = Utf8BytesFromString(szInput)
    ' Add an extra NUL byte to the array
    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
    ReDim Preserve b(nLen)
    b(nLen - 1) = 0
    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
    qrcodeCreateGifInUtf8 = QRCODE_CreateGifFromBytes(szOutputFile, b(0), nPixelsPerModule, szParams, opts)
End Function


'/**
' Create a PDF file of a QR code, encoding input in UTF-8.
' @param szOutputFile Name of output GIF file to be created.
' @param szInput Text input to be encoded.
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nPageWidth  Width of PDF page in pixels (default = 0 : set width to fit QRcode image).
' @param nPageHeight Height of PDF page in pixels (default = 0 : set height to fit QRcode image).
' @param nX  X-coordinate in pixels of bottom-left of QRcode image (default = 0 : at left side).
' @param nY  Y-coordinate in pixels of bottom-left of QRcode image (default = 0 : at bottom).
' @param nOptions Option flags. Set as 0 for default.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check).
' @remark Any non-ASCII characters in `szInput` will be encoded in UTF-8 before processing.
'**/
Public Function qrcodeCreatePdfInUtf8(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nPageWidth As Long = 0, Optional nPageHeight As Long = 0, Optional nX As Long = 0, Optional nY As Long = 0, Optional nOptions As Long = 0) As Long
    Dim b() As Byte
    Dim nLen As Long
    Dim opts As Long
    opts = nOptions Or EccLevel
    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
    b = Utf8BytesFromString(szInput)
    ' Add an extra NUL byte to the array
    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
    ReDim Preserve b(nLen)
    b(nLen - 1) = 0
    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
    qrcodeCreatePdfInUtf8 = QRCODE_CreatePdfFromBytes(szOutputFile, b(0), nPixelsPerModule, nPageWidth, nPageHeight, nX, nY, opts)
End Function

'/**
' Get information about the core native DLL.
' @param  nOptions For future use.
' @return Information about the core DLL module.
' @remark Example result.
' `Platform` is the platform the core DLL was compiled for: `Win32` or `X64`.
' {@code
' "Platform=Win32; Compiled=Mar 27 2021 03:50:03; Licence=T"
' }
'**/
Public Function qrcodeDllInfo(Optional nOptions As Long = 0) As String
    Dim nc As Long
    nc = QRCODE_DllInfo("", 0, nOptions)
    If nc <= 0 Then Exit Function
    qrcodeDllInfo = String(nc, " ")
    nc = QRCODE_DllInfo(qrcodeDllInfo, Len(qrcodeDllInfo), nOptions)
    qrcodeDllInfo = left$(qrcodeDllInfo, nc)
End Function

'/**
' Look up description for error code.
' @param  nErrCode Value of error code to lookup (may be positive or negative).
' @return Error message, or empty string if no corresponding error code.
'**/
Public Function qrcodeErrorLookup(nErrCode As Long) As String
    Dim nc As Long
    nc = QRCODE_ErrorLookup("", 0, nErrCode)
    If nc <= 0 Then Exit Function
    qrcodeErrorLookup = String(nc, " ")
    nc = QRCODE_ErrorLookup(qrcodeErrorLookup, nc, nErrCode)
    qrcodeErrorLookup = left$(qrcodeErrorLookup, nc)
End Function



Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean, Optional actualTo As Boolean = False) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    actualTo = True
    If actualTo = True Then
   lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = SystemOptions.cdoSMTPServer '  "mail.sattaryah.com"
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = SystemOptions.cdoSMTPServerPort  '25
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = SystemOptions.cdoSMTPUseSSL '  False
     lobj_cdomsg.BodyPart.Charset = "Windows-1256"
     
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = 1
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = SystemOptions.cdoSendUserName ' "a.s@sattaryah.com"
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = SystemOptions.cdoSendPassword '"spamkiller"
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = 2
    lobj_cdomsg.Configuration.Fields.update
    If sTo = "" Then
    lobj_cdomsg.To = "a.s@sattaryah.com,e.m@sattaryah.com"
    Else
    lobj_cdomsg.To = sTo
    End If
    lobj_cdomsg.From = SystemOptions.TxtFromName & "<" & SystemOptions.txtFromEmail & " >"  '  "Dynamic ERP<info@sattaryah.com>"
    lobj_cdomsg.subject = sSubject
    lobj_cdomsg.TextBody = sBody
    Else
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.subject = sSubject
    lobj_cdomsg.TextBody = sBody
    End If
    
    If Trim$(sFilePath) <> vbNullString Then
         lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function
Public Function GETContractDateDATE(Emp_id As Integer, Optional novalue As Boolean) As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT   Contract_date AS MaxDate from dbo.Contract WHERE     (emp_id = " & Emp_id & ")"
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    If Not IsNull(rs("MaxDate").value) Then
 GETContractDateDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
novalue = False
Else
 GETContractDateDATE = Date
 novalue = True
 End If
 Else
 GETContractDateDATE = Date
 novalue = True
    End If

End Function
Public Function GETlASTiSSUEDATENew(Emp_id As Integer, Optional novalue As Boolean, Optional Str_Last As Integer = 0) As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
 If Str_Last = 0 Then
  sql = "SELECT     MAX(BignDateWork) AS MaxDate from dbo.TblEmployee WHERE     (Emp_ID = " & Emp_id & ")"
  Else
  
 sql = "SELECT     MAX(lastHolidaydate) AS MaxDate from dbo.TblEmployee WHERE     (Emp_ID = " & Emp_id & ")"
 End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    If Not IsNull(rs("MaxDate").value) Then
 GETlASTiSSUEDATENew = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
novalue = False
Else
 'GETlASTiSSUEDATE = Date
 GETlASTiSSUEDATENew = GETContractDateDATE(Emp_id)
 
 novalue = True
 End If
 Else
 'GETlASTiSSUEDATE = Date
 GETlASTiSSUEDATENew = GETContractDateDATE(Emp_id)
 
 novalue = True
    End If

End Function

Public Function getFinancialEquationData(FinancialEquationsId As Integer, _
                                         Opr As String, _
                                         generalvalue As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim AccountsCodes As String
    AccountsCodes = ""
    sql = "Select  *  from FinancialEquations where FinancialEquationsId=" & FinancialEquationsId
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then: Exit Function
    Opr = IIf(IsNull(Rs3("FinancialEquationsOpr").value), "", Rs3("FinancialEquationsOpr").value)
    generalvalue = IIf(Not IsNumeric(Rs3("GeneralValue").value), 0, Rs3("GeneralValue").value)
   
    Rs3.Close

End Function

Public Function BalanceSheetAccount(BalanceSheetId As Integer, _
                                    filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim AccountsCodes As String
    AccountsCodes = ""
    sql = "Select  " & filed & " from BalanceSheetViewAccounts where BalanceSheetId=" & BalanceSheetId
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then BalanceSheetAccount = "": Exit Function
    Dim i As Long
    For i = 1 To Rs3.RecordCount
        AccountsCodes = AccountsCodes & "'" & Rs3(filed).value & "',"
        Rs3.MoveNext
    Next i

    AccountsCodes = mId(AccountsCodes, 1, Len(AccountsCodes) - 1)
    BalanceSheetAccount = AccountsCodes
    Rs3.Close

End Function

Public Function get_last_month_day(pDate As Date) As String

    Select Case Month(pDate)

        Case 1, 3, 5, 7, 8, 10, 12
            get_last_month_day = "31"

        Case 4, 6, 9, 11
            get_last_month_day = "30"

        Case 2

            If (year(pDate) Mod 4) = 0 Then
                get_last_month_day = "29"
            Else
                get_last_month_day = "28"
            End If

    End Select
 
End Function


Public Function ByteToStr(bArray() As Byte) As String
    Dim lPntr As Long
    Dim bTmp() As Byte
    On Error GoTo ByteErr
    ReDim bTmp(UBound(bArray) * 2 + 1)
    For lPntr = 0 To UBound(bArray)
        bTmp(lPntr * 2) = bArray(lPntr)
    Next lPntr
    Let ByteToStr = bTmp
    Exit Function
ByteErr:
    ByteToStr = ""
End Function

Public Function ByteToUni(bArray() As Byte) As String
    ByteToUni = bArray
End Function

Public Sub DebugPrintByte(sDescr As String, bArray() As Byte)
    Dim lPtr As Long
    Debug.Print sDescr & ":"
    If GetbSize(bArray) = 0 Then Exit Sub
    For lPtr = 0 To UBound(bArray)
        Debug.Print right$("0" & Hex$(bArray(lPtr)), 2) & " ";
        If (lPtr + 1) Mod 16 = 0 Then Debug.Print
    Next lPtr
    Debug.Print
End Sub

Public Function GetbSize(bArray() As Byte) As Long
    On Error GoTo GetSizeErr
    GetbSize = UBound(bArray) + 1
    Exit Function
GetSizeErr:
    GetbSize = 0
End Function

Public Function PeekB(ByVal lpdwData As Long) As Byte
    CopyMemory PeekB, ByVal lpdwData, 1
End Function
 
Public Sub DebugPrintString(sDescr As String, strToPrint As String)
    Dim lPtr As Long
    Dim sSep As String * 1
    Debug.Print sDescr & ":"
    For lPtr = 0 To LenB(strToPrint) - 1
        Debug.Print right$("0" & Hex$(PeekB(StrPtr(strToPrint) + lPtr)), 2) & " ";
        If (lPtr + 2) Mod 16 = 0 Then Debug.Print
    Next lPtr
    Debug.Print
End Sub

Public Function StrToByte(strInput As String) As Byte()
    Dim lPntr As Long
    Dim bTmp() As Byte
    Dim bArray() As Byte
    If Len(strInput) = 0 Then Exit Function
    ReDim bTmp(LenB(strInput) - 1) 'Memory length
    ReDim bArray(Len(strInput) - 1) 'String length
    CopyMemory bTmp(0), ByVal StrPtr(strInput), LenB(strInput)
    'Examine every second byte
    For lPntr = 0 To UBound(bArray)
        If bTmp(lPntr * 2 + 1) > 0 Then
            'bArray(lPntr) = Asc(Mid$(strInput, lPntr + 1, 1))
            StrToByte = bTmp
            Exit Function
        Else
            bArray(lPntr) = bTmp(lPntr * 2)
        End If
    Next lPntr
    StrToByte = bArray
End Function

Public Function UniToByte(strInput As String) As Byte()
    UniToByte = strInput
End Function


Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long

    For lIdx = 1 To Len(sText)
        'ToHexDump = ToHexDump & right$("0" & Hex(Asc(mId(sText, lIdx, 1))), 2)
        ToHexDump = ToHexDump & right$("0" & Hex(Asc(mId(sText, lIdx, 1))), 2)
    Next
End Function
Public Function createTLV(TLVtag As String, TLVvalue As String) As String

 

createTLV = ((TLVtag)) & zeropadding((Hex(Len(TLVvalue))), 2) & ToHexDump(TLVvalue)
End Function


Private Function BytesToHex(ByRef Bytes() As Byte) As String
    'Quick and dirty Byte array to hex String, format:
    '
    '   "HH HH HH"

    Dim LB As Long
    Dim ByteCount As Long
    Dim BytePos As Integer

    LB = LBound(Bytes)
    ByteCount = UBound(Bytes) - LB + 1
    If ByteCount < 1 Then Exit Function
    BytesToHex = Space$(3 * (ByteCount - 1) + 2)
    For BytePos = LB To UBound(Bytes)
        Mid$(BytesToHex, 3 * (BytePos - LB) + 1, 2) = _
            right$("0" & Hex$(Bytes(BytePos)), 2)
    Next
End Function

Public Function HexToBytes(ByVal HexString As String) As Byte()
    'Quick and dirty hex String to Byte array.  Accepts:
    '
    '   "HH HH HH"
    '   "HHHHHH"
    '   "H HH H"
    '   "HH,HH,     HH" and so on.

    Dim Bytes() As Byte
    Dim HexPos As Integer
    Dim HexDigit As Integer
    Dim BytePos As Integer
    Dim Digits As Integer

    ReDim Bytes(Len(HexString) \ 2)  'Initial estimate.
    For HexPos = 1 To Len(HexString)
        HexDigit = InStr("0123456789ABCDEF", _
                         UCase$(mId$(HexString, HexPos, 1))) - 1
        If HexDigit >= 0 Then
            If BytePos > UBound(Bytes) Then
                'Add some room, we'll add room for 4 more to decrease
                'how often we end up doing this expensive step:
                ReDim Preserve Bytes(UBound(Bytes) + 4)
            End If
            Bytes(BytePos) = Bytes(BytePos) * &H10 + HexDigit
            Digits = Digits + 1
        End If
        If Digits = 2 Or HexDigit < 0 Then
            If Digits > 0 Then BytePos = BytePos + 1
            Digits = 0
        End If
    Next
    If Digits = 0 Then BytePos = BytePos - 1
    If BytePos < 0 Then
        Bytes = "" 'Empty.
    Else
        ReDim Preserve Bytes(BytePos)
    End If
    HexToBytes = Bytes
End Function

Public Function HexToString(value As String)
    Dim szTemp As String
    szTemp = value
    
    Dim szData As String
    szData = ""
    While Len(szTemp) > 0
        szData = CHR(CLng("&h" & right(szTemp, 2))) & szData
        If (Len(szTemp) = 1) Then
            szTemp = left(szTemp, Len(szTemp) - 1)
        Else
            szTemp = left(szTemp, Len(szTemp) - 2)
        End If
    Wend
    HexToString = szData
End Function





'
'
'' $Id: basQRcode.bas
''/**
'' The VBA/VB6 interface to diQRcode.dll.
''
'' @author dai
'' @version 4.0.0
''**/
'' $Date: 2021-06-16 10:07:00 $
'
''*********************** COPYRIGHT NOTICE***********************
'' Copyright (c) 2021 DI Management Services Pty Limited.
'' <www.di-mgt.com.au> <www.cryptosys.net>. All rights reserved.
'' This code may only be used by licensed users and in
'' accordance with the licence conditions.
'' The latest version of diQRcode and a licence
'' may be obtained from <https://www.cryptosys.net/qrcode>.
'' This copyright notice must always be left intact.
''******************** END OF COPYRIGHT NOTICE*******************
'
'
'Option Explicit
'Option Base 0
'
''//////////////////////////////////////////////////
'' ECC LEVELS
'Public Const QRCODE_ECC_M As Long = 0  ' Default
'Public Const QRCODE_ECC_L As Long = 1
'Public Const QRCODE_ECC_Q As Long = 2
'Public Const QRCODE_ECC_H As Long = 3
'' OPTION FLAGS
'Public Const QRCODE_ESCAPED As Long = &H1000&
'Public Const QRCODE_BASE64  As Long = &H2000&
'Public Const QRCODE_NO_NAMECHECK As Long = &H10000
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
''/**
'' Image type
''**/
'Public Enum ImageType
'    '/**
'    ' Output image as a GIF file (default)
'    '**/
'    GIF = 0
'    '/**
'    ' Output image as an SVG file
'    '**/
'    SVG = &H10
'End Enum
'
''/**
'' Error-correction code
''**/
'Public Enum Ecc
'    '/**
'    ' EC level M (default)
'    '**/
'    mNr = 0
'    '/**
'    ' EC level L
'    '**/
'    lNr = 1
'    '/**
'    ' EC level Q
'    '**/
'    QNr = 2
'    '/**
'    ' EC level H
'    '**/
'    HNr = 3
'End Enum
'
'' Local constants
'Private Const OUT_OF_RANGE_ERROR As Long = 11
'
''#If VBA7 Then
''' Declarations for 64-bit Office
''' (In VB6 these will appear red. Turn off "Auto Syntax Check" in Tools > Options to avoid annoying warnings)
''Private Declare PtrSafe Function QRCODE_CreateImage Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
''Private Declare PtrSafe Function QRCODE_CreateImageFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateImage" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
''Private Declare PtrSafe Function QRCODE_CreateGif Lib "diQRcode.dll" (ByVal strOutputFile As String, ByVal strInput As String, ByVal nPixelsPerModule As Long, ByVal strParams As String, ByVal nOptions As Long) As Long
''Private Declare PtrSafe Function QRCODE_CreateGifFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateGif" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
''Private Declare PtrSafe Function QRCODE_CreatePdf Lib "diQRcode.dll" (ByVal strOutputFile As String, ByVal strInput As String, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
''Private Declare PtrSafe Function QRCODE_CreatePdfFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreatePdf" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
''Public Declare PtrSafe Function QRCODE_Version Lib "diQRcode.dll" () As Long
''Private Declare PtrSafe Function QRCODE_DllInfo Lib "diQRcode.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal nOptions As Long) As Long
''Private Declare PtrSafe Function QRCODE_ErrorLookup Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nErrCode As Long) As Long
''Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
''    ByVal CodePage As Long, _
''    ByVal dwFlags As Long, _
''    ByVal lpWideCharStr As Long, _
''    ByVal cchWideChar As Long, _
''    ByVal lpMultiByteStr As Long, _
''    ByVal cbMultiByte As Long, _
''    ByVal lpDefaultChar As Long, _
''    ByVal lpUsedDefaultChar As Long _
''    ) As Long
''#Else
'' Declarations for VB6 and 32-bit Office
'Private Declare Function QRCODE_CreateImage Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
'Private Declare Function QRCODE_CreateImageFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateImage" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
'Private Declare Function QRCODE_CreateGif Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
'Private Declare Function QRCODE_CreateGifFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateGif" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
'Private Declare Function QRCODE_CreatePdf Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
'Private Declare Function QRCODE_CreatePdfFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreatePdf" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
'Private Declare Function QRCODE_ErrorLookup Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nErrCode As Long) As Long
'
''/**
'' Get version number of core DLL.
'' @return Version number as an integer in form `Major*10000 + Minor*100 + Release`. For example, version 2.10.3 would return 21003.
''**/
'Public Declare Function QRCODE_Version Lib "diQRcode.dll" () As Long
'Private Declare Function QRCODE_DllInfo Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nOptions As Long) As Long
'
''///////////////////////////////////////////////////////////////
'' INTERNAL UTF-8 MAPPING
'''' WinApi function that maps a UTF-16 (wide character) string to a new character string
'Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
'    ByVal CodePage As Long, _
'    ByVal dwFlags As Long, _
'    ByVal lpWideCharStr As Long, _
'    ByVal cchWideChar As Long, _
'    ByVal lpMultiByteStr As Long, _
'    ByVal cbMultiByte As Long, _
'    ByVal lpDefaultChar As Long, _
'    ByVal lpUsedDefaultChar As Long _
'    ) As Long
''#End If
'
'' CodePage constant for UTF-8
'Private Const CP_UTF8 = 65001
'
'''' Return byte array with VBA "Unicode" string encoded in UTF-8
'' Ref: [How to convert VBA/VB6 Unicode strings to UTF-8](https://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html)
'Private Function Utf8BytesFromString(strInput As String) As Byte()
'    Dim nBytes As Long
'    Dim abBuffer() As Byte
'    ' Catch empty or null input string
'    Utf8BytesFromString = vbNullString
'    If Len(strInput) < 1 Then Exit Function
'    ' Get length in bytes *including* terminating null
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
'    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
'    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
'    Utf8BytesFromString = abBuffer
'End Function
'
'' ------------------
'' EXPORTED FUNCTIONS
'' ------------------
'
''/**
'' Create an image file of a QR code.
'' @param szOutputFile Name of output image file to be created.
'' @param szInput Text input to be encoded (ANSI characters only).
'' @param imgType Image type (default = GIF)
'' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
'' @param nMargin Size of margin in modules (default = 4 modules)
'' @param EccLevel Error correction level (default = Ecc.mNr)
'' @param nOptions Option flags. Use the `Or` operator to combine: <br>
'' `QRCODE_ESCAPED` to indicate #-escaped sequences in string <br>
'' `QRCODE_BASE64` to encode output as base64 text <br>
'' `QRCODE_NO_NAMECHECK` do not check filename extension against file type <br>
'' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
''**/
'Public Function qrcodeCreateImage(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
'    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nOptions As Long = 0) As Long
'    Dim i As Long
'    Dim opts As Long
'    opts = nOptions Or EccLevel Or imgType
'    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
'    For i = 1 To Len(szInput)
'        If AscW(mId(szInput, i, 1)) > &HFF Then
'            qrcodeCreateImage = -OUT_OF_RANGE_ERROR
'            Exit Function
'        End If
'    Next
'    ' Do the business
'    qrcodeCreateImage = QRCODE_CreateImage(szOutputFile, szInput, nPixelsPerModule, nMargin, opts)
'End Function
'
''/**
'' Create a GIF file of a QR code (_deprecated_).
'' @param szOutputFile Name of output GIF file to be created.
'' @param szInput Text input to be encoded (ANSI characters only).
'' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
'' @param EccLevel Error correction level (default = Ecc.M)
'' @param szParams Optional parameters.  Set as `"margin=N"` to change the margin to `N` modules (default=4).
'' @param nOptions Option flags. Set as `QRCODE_ESCAPED` to indicate #-escaped sequences in string.
'' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'' @deprecated Use {@link qrcodeCreateImage}.
''**/
'Public Function qrcodeCreateGif(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional szParams As String = "", Optional nOptions As Long = 0) As Long
'    Dim i As Long
'    Dim opts As Long
'    Const OUT_OF_RANGE_ERROR As Long = 11
'    opts = nOptions Or EccLevel
'    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
'    For i = 1 To Len(szInput)
'        If AscW(mId(szInput, i, 1)) > &HFF Then
'            qrcodeCreateGif = -OUT_OF_RANGE_ERROR
'            Exit Function
'        End If
'    Next
'    ' Do the business
'    qrcodeCreateGif = QRCODE_CreateGif(szOutputFile, szInput, nPixelsPerModule, szParams, opts)
'End Function
'
''/**
'' Create a PDF file of a QR code.
'' @param szOutputFile Name of output PDF file to be created.
'' @param szInput Text input to be encoded (ANSI characters only).
'' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
'' @param EccLevel Error correction level (default = Ecc.M)
'' @param nPageWidth  Width of PDF page in pixels (default = 0 : set width to fit QRcode image).
'' @param nPageHeight Height of PDF page in pixels (default = 0 : set height to fit QRcode image).
'' @param nX  X-coordinate in pixels of bottom-left of QRcode image (default = 0 : at left side).
'' @param nY  Y-coordinate in pixels of bottom-left of QRcode image (default = 0 : at bottom).
'' @param nOptions Option flags. Set as `QRCODE_ESCAPED` to indicate #-escaped sequences in string.
'' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
''**/
'Public Function qrcodeCreatePdf(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nPageWidth As Long = 0, Optional nPageHeight As Long = 0, Optional nX As Long = 0, Optional nY As Long = 0, Optional nOptions As Long = 0) As Long
'    Dim i As Long
'    Dim opts As Long
'    Const OUT_OF_RANGE_ERROR As Long = 11
'    opts = nOptions Or EccLevel
'    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
'    For i = 1 To Len(szInput)
'        If AscW(mId(szInput, i, 1)) > &HFF Then
'            qrcodeCreatePdf = -OUT_OF_RANGE_ERROR
'            Exit Function
'        End If
'    Next
'    ' Do the business
'    qrcodeCreatePdf = QRCODE_CreatePdf(szOutputFile, szInput, nPixelsPerModule, nPageWidth, nPageHeight, nX, nY, opts)
'End Function
'
''/**
'' Create a image file of a QR code, encoding input in UTF-8.
'' @param szOutputFile Name of output image file to be created.
'' @param szInput Text input to be encoded.
'' @param imgType Image type (default = GIF)
'' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
'' @param nMargin Size of margin in modules (default = 4 modules)
'' @param EccLevel Error correction level (default = Ecc.M)
'' @param nOptions Option flags. Use the `Or` operator to combine: <br>
'' `QRCODE_ESCAPED` to indicate #-escaped sequences in string <br>
'' `QRCODE_BASE64` to encode output as base64 text <br>
'' `QRCODE_NO_NAMECHECK` do not check filename extension against file type <br>
'' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
''**/
''Public Function qrcodeCreateImageInUtf8(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
''    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nOptions As Long = 0) As Long
''    Dim b() As Byte
''    Dim nLen As Long
''    Dim opts As Long
''    opts = nOptions Or EccLevel Or imgType
''    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
''    b = Utf8BytesFromString(szInput)
''    ' Add an extra NUL byte to the array
''    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
''    ReDim Preserve b(nLen)
''    b(nLen - 1) = 0
''    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
''    qrcodeCreateImageInUtf8 = QRCODE_CreateImageFromBytes(szOutputFile, b(0), nPixelsPerModule, nMargin, opts)
''End Function
'
'Public Function qrcodeCreateImageInUtf8(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
'    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nOptions As Long = 0) As Long
'    Dim b() As Byte
'    Dim nLen As Long
'    Dim opts As Long
'    opts = nOptions Or EccLevel Or imgType
'    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
'    b = Utf8BytesFromString(szInput)
'    ' Add an extra NUL byte to the array
'    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
'    ReDim Preserve b(nLen)
'    b(nLen - 1) = 0
'    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
'    qrcodeCreateImageInUtf8 = QRCODE_CreateImageFromBytes(szOutputFile, b(0), nPixelsPerModule, nMargin, opts)
'End Function
'
''/**
'' Create a GIF file of a QR code, encoding input in UTF-8 (_deprecated_).
'' @param szOutputFile Name of output GIF file to be created.
'' @param szInput Text input to be encoded.
'' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
'' @param EccLevel Error correction level (default = Ecc.M)
'' @param szParams Optional parameters. Set as `"margin=N"` to change the margin to `N` modules (default=4).
'' @param nOptions Option flags. Set as 0 for default.
'' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check).
'' @remark Any non-ASCII characters in `szInput` will be encoded in UTF-8 before processing.
'' @deprecated Use {@link qrcodeCreateImageInUtf8}.
''**/
'Public Function qrcodeCreateGifInUtf8(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional szParams As String = "", Optional nOptions As Long = 0) As Long
'    Dim b() As Byte
'    Dim nLen As Long
'    Dim opts As Long
'    opts = nOptions Or EccLevel
'    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
'    b = Utf8BytesFromString(szInput)
'    ' Add an extra NUL byte to the array
'    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
'    ReDim Preserve b(nLen)
'    b(nLen - 1) = 0
'    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
'    qrcodeCreateGifInUtf8 = QRCODE_CreateGifFromBytes(szOutputFile, b(0), nPixelsPerModule, szParams, opts)
'End Function
'
'
''/**
'' Create a PDF file of a QR code, encoding input in UTF-8.
'' @param szOutputFile Name of output GIF file to be created.
'' @param szInput Text input to be encoded.
'' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
'' @param EccLevel Error correction level (default = Ecc.M)
'' @param nPageWidth  Width of PDF page in pixels (default = 0 : set width to fit QRcode image).
'' @param nPageHeight Height of PDF page in pixels (default = 0 : set height to fit QRcode image).
'' @param nX  X-coordinate in pixels of bottom-left of QRcode image (default = 0 : at left side).
'' @param nY  Y-coordinate in pixels of bottom-left of QRcode image (default = 0 : at bottom).
'' @param nOptions Option flags. Set as 0 for default.
'' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check).
'' @remark Any non-ASCII characters in `szInput` will be encoded in UTF-8 before processing.
''**/
'Public Function qrcodeCreatePdfInUtf8(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.mNr, Optional nPageWidth As Long = 0, Optional nPageHeight As Long = 0, Optional nX As Long = 0, Optional nY As Long = 0, Optional nOptions As Long = 0) As Long
'    Dim b() As Byte
'    Dim nLen As Long
'    Dim opts As Long
'    opts = nOptions Or EccLevel
'    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
'    b = Utf8BytesFromString(szInput)
'    ' Add an extra NUL byte to the array
'    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
'    ReDim Preserve b(nLen)
'    b(nLen - 1) = 0
'    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
'    qrcodeCreatePdfInUtf8 = QRCODE_CreatePdfFromBytes(szOutputFile, b(0), nPixelsPerModule, nPageWidth, nPageHeight, nX, nY, opts)
'End Function
'
''/**
'' Get information about the core native DLL.
'' @param  nOptions For future use.
'' @return Information about the core DLL module.
'' @remark Example result.
'' `Platform` is the platform the core DLL was compiled for: `Win32` or `X64`.
'' {@code
'' "Platform=Win32; Compiled=Mar 27 2021 03:50:03; Licence=T"
'' }
''**/
'Public Function qrcodeDllInfo(Optional nOptions As Long = 0) As String
'    Dim nc As Long
'    nc = QRCODE_DllInfo("", 0, nOptions)
'    If nc <= 0 Then Exit Function
'    qrcodeDllInfo = String(nc, " ")
'    nc = QRCODE_DllInfo(qrcodeDllInfo, Len(qrcodeDllInfo), nOptions)
'    qrcodeDllInfo = left$(qrcodeDllInfo, nc)
'End Function
'
''/**
'' Look up description for error code.
'' @param  nErrCode Value of error code to lookup (may be positive or negative).
'' @return Error message, or empty string if no corresponding error code.
''**/
'Public Function qrcodeErrorLookup(nErrCode As Long) As String
'    Dim nc As Long
'    nc = QRCODE_ErrorLookup("", 0, nErrCode)
'    If nc <= 0 Then Exit Function
'    qrcodeErrorLookup = String(nc, " ")
'    nc = QRCODE_ErrorLookup(qrcodeErrorLookup, nc, nErrCode)
'    qrcodeErrorLookup = left$(qrcodeErrorLookup, nc)
'End Function
'
'
'
'Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
'    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
'    sSmtpUser As String, sSmtpPword As String, _
'    sFilePath As String, bSmtpSSL As Boolean, Optional actualTo As Boolean = False) As String
'
'    On Error GoTo SendMail_Error:
'    Dim lobj_cdomsg      As CDO.Message
'    Set lobj_cdomsg = New CDO.Message
'    actualTo = True
'    If actualTo = True Then
'   lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = SystemOptions.cdoSMTPServer '  "mail.sattaryah.com"
'    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = SystemOptions.cdoSMTPServerPort  '25
'    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = SystemOptions.cdoSMTPUseSSL '  False
'     lobj_cdomsg.BodyPart.Charset = "Windows-1256"
'
'    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = 1
'    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = SystemOptions.cdoSendUserName ' "a.s@sattaryah.com"
'    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = SystemOptions.cdoSendPassword '"spamkiller"
'    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
'    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = 2
'    lobj_cdomsg.Configuration.Fields.update
'    If sTo = "" Then
'    lobj_cdomsg.To = "a.s@sattaryah.com,e.m@sattaryah.com"
'    Else
'    lobj_cdomsg.To = sTo
'    End If
'    lobj_cdomsg.From = SystemOptions.TxtFromName & "<" & SystemOptions.txtFromEmail & " >"  '  "Dynamic ERP<info@sattaryah.com>"
'    lobj_cdomsg.subject = sSubject
'    lobj_cdomsg.TextBody = sBody
'    Else
'    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
'    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
'    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
'    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
'    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
'    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
'    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
'    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
'    lobj_cdomsg.Configuration.Fields.update
'    lobj_cdomsg.To = sTo
'    lobj_cdomsg.From = sFrom
'    lobj_cdomsg.subject = sSubject
'    lobj_cdomsg.TextBody = sBody
'    End If
'
'    If Trim$(sFilePath) <> vbNullString Then
'         lobj_cdomsg.AddAttachment (sFilePath)
'    End If
'    lobj_cdomsg.send
'    Set lobj_cdomsg = Nothing
'    SendMail = "ok"
'    Exit Function
'
'SendMail_Error:
'    SendMail = Err.Description
'End Function
'Public Function GETContractDateDATE(Emp_id As Integer, Optional novalue As Boolean) As Date
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
'
'  sql = "SELECT   Contract_date AS MaxDate from dbo.Contract WHERE     (emp_id = " & Emp_id & ")"
'
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'    If Not IsNull(rs("MaxDate").value) Then
' GETContractDateDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
'novalue = False
'Else
' GETContractDateDATE = Date
' novalue = True
' End If
' Else
' GETContractDateDATE = Date
' novalue = True
'    End If
'
'End Function
'Public Function GETlASTiSSUEDATENew(Emp_id As Integer, Optional novalue As Boolean, Optional Str_Last As Integer = 0) As Date
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
' If Str_Last = 0 Then
'  sql = "SELECT     MAX(BignDateWork) AS MaxDate from dbo.TblEmployee WHERE     (Emp_ID = " & Emp_id & ")"
'  Else
'
' sql = "SELECT     MAX(lastHolidaydate) AS MaxDate from dbo.TblEmployee WHERE     (Emp_ID = " & Emp_id & ")"
' End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'    If Not IsNull(rs("MaxDate").value) Then
' GETlASTiSSUEDATENew = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
'novalue = False
'Else
' 'GETlASTiSSUEDATE = Date
' GETlASTiSSUEDATENew = GETContractDateDATE(Emp_id)
'
' novalue = True
' End If
' Else
' 'GETlASTiSSUEDATE = Date
' GETlASTiSSUEDATENew = GETContractDateDATE(Emp_id)
'
' novalue = True
'    End If
'
'End Function
'
'Public Function getFinancialEquationData(FinancialEquationsId As Integer, _
'                                         Opr As String, _
'                                         generalvalue As Double)
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    Dim AccountsCodes As String
'    AccountsCodes = ""
'    sql = "Select  *  from FinancialEquations where FinancialEquationsId=" & FinancialEquationsId
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then: Exit Function
'    Opr = IIf(IsNull(Rs3("FinancialEquationsOpr").value), "", Rs3("FinancialEquationsOpr").value)
'    generalvalue = IIf(Not IsNumeric(Rs3("GeneralValue").value), 0, Rs3("GeneralValue").value)
'
'    Rs3.Close
'
'End Function
'
'Public Function BalanceSheetAccount(BalanceSheetId As Integer, _
'                                    filed As String) As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    Dim AccountsCodes As String
'    AccountsCodes = ""
'    sql = "Select  " & filed & " from BalanceSheetViewAccounts where BalanceSheetId=" & BalanceSheetId
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then BalanceSheetAccount = "": Exit Function
'    Dim i As Long
'    For i = 1 To Rs3.RecordCount
'        AccountsCodes = AccountsCodes & "'" & Rs3(filed).value & "',"
'        Rs3.MoveNext
'    Next i
'
'    AccountsCodes = mId(AccountsCodes, 1, Len(AccountsCodes) - 1)
'    BalanceSheetAccount = AccountsCodes
'    Rs3.Close
'
'End Function
'
'Public Function get_last_month_day(pDate As Date) As String
'
'    Select Case Month(pDate)
'
'        Case 1, 3, 5, 7, 8, 10, 12
'            get_last_month_day = "31"
'
'        Case 4, 6, 9, 11
'            get_last_month_day = "30"
'
'        Case 2
'
'            If (year(pDate) Mod 4) = 0 Then
'                get_last_month_day = "29"
'            Else
'                get_last_month_day = "28"
'            End If
'
'    End Select
'
'End Function
'
'
'Public Function ByteToStr(bArray() As Byte) As String
'    Dim lPntr As Long
'    Dim bTmp() As Byte
'    On Error GoTo ByteErr
'    ReDim bTmp(UBound(bArray) * 2 + 1)
'    For lPntr = 0 To UBound(bArray)
'        bTmp(lPntr * 2) = bArray(lPntr)
'    Next lPntr
'    Let ByteToStr = bTmp
'    Exit Function
'ByteErr:
'    ByteToStr = ""
'End Function
'
'Public Function ByteToUni(bArray() As Byte) As String
'    ByteToUni = bArray
'End Function
'
'Public Sub DebugPrintByte(sDescr As String, bArray() As Byte)
'    Dim lPtr As Long
'    Debug.Print sDescr & ":"
'    If GetbSize(bArray) = 0 Then Exit Sub
'    For lPtr = 0 To UBound(bArray)
'        Debug.Print right$("0" & Hex$(bArray(lPtr)), 2) & " ";
'        If (lPtr + 1) Mod 16 = 0 Then Debug.Print
'    Next lPtr
'    Debug.Print
'End Sub
'
'Public Function GetbSize(bArray() As Byte) As Long
'    On Error GoTo GetSizeErr
'    GetbSize = UBound(bArray) + 1
'    Exit Function
'GetSizeErr:
'    GetbSize = 0
'End Function
'
'Public Function PeekB(ByVal lpdwData As Long) As Byte
'    CopyMemory PeekB, ByVal lpdwData, 1
'End Function
'
'Public Sub DebugPrintString(sDescr As String, strToPrint As String)
'    Dim lPtr As Long
'    Dim sSep As String * 1
'    Debug.Print sDescr & ":"
'    For lPtr = 0 To LenB(strToPrint) - 1
'        Debug.Print right$("0" & Hex$(PeekB(StrPtr(strToPrint) + lPtr)), 2) & " ";
'        If (lPtr + 2) Mod 16 = 0 Then Debug.Print
'    Next lPtr
'    Debug.Print
'End Sub
'
'Public Function StrToByte(strInput As String) As Byte()
'    Dim lPntr As Long
'    Dim bTmp() As Byte
'    Dim bArray() As Byte
'    If Len(strInput) = 0 Then Exit Function
'    ReDim bTmp(LenB(strInput) - 1) 'Memory length
'    ReDim bArray(Len(strInput) - 1) 'String length
'    CopyMemory bTmp(0), ByVal StrPtr(strInput), LenB(strInput)
'    'Examine every second byte
'    For lPntr = 0 To UBound(bArray)
'        If bTmp(lPntr * 2 + 1) > 0 Then
'            'bArray(lPntr) = Asc(Mid$(strInput, lPntr + 1, 1))
'            StrToByte = bTmp
'            Exit Function
'        Else
'            bArray(lPntr) = bTmp(lPntr * 2)
'        End If
'    Next lPntr
'    StrToByte = bArray
'End Function
'
'Public Function UniToByte(strInput As String) As Byte()
'    UniToByte = strInput
'End Function
'
'
'Public Function ToHexDump(sText As String) As String
'    Dim lIdx            As Long
'
'    For lIdx = 1 To Len(sText)
'        ToHexDump = ToHexDump & right$("0" & Hex(Asc(mId(sText, lIdx, 1))), 2)
'    Next
'End Function
'Public Function createTLV(TLVtag As String, TLVvalue As String) As String
'
'
'createTLV = ((TLVtag)) & zeropadding((Hex(Len(TLVvalue))), 2) & ToHexDump(TLVvalue)
'End Function
'
'
'Private Function BytesToHex(ByRef Bytes() As Byte) As String
'    'Quick and dirty Byte array to hex String, format:
'    '
'    '   "HH HH HH"
'
'    Dim LB As Long
'    Dim ByteCount As Long
'    Dim BytePos As Integer
'
'    LB = LBound(Bytes)
'    ByteCount = UBound(Bytes) - LB + 1
'    If ByteCount < 1 Then Exit Function
'    BytesToHex = Space$(3 * (ByteCount - 1) + 2)
'    For BytePos = LB To UBound(Bytes)
'        Mid$(BytesToHex, 3 * (BytePos - LB) + 1, 2) = _
'            right$("0" & Hex$(Bytes(BytePos)), 2)
'    Next
'End Function
'
'Public Function HexToBytes(ByVal HexString As String) As Byte()
'    'Quick and dirty hex String to Byte array.  Accepts:
'    '
'    '   "HH HH HH"
'    '   "HHHHHH"
'    '   "H HH H"
'    '   "HH,HH,     HH" and so on.
'
'    Dim Bytes() As Byte
'    Dim HexPos As Integer
'    Dim HexDigit As Integer
'    Dim BytePos As Integer
'    Dim Digits As Integer
'
'    ReDim Bytes(Len(HexString) \ 2)  'Initial estimate.
'    For HexPos = 1 To Len(HexString)
'        HexDigit = InStr("0123456789ABCDEF", _
'                         UCase$(mId$(HexString, HexPos, 1))) - 1
'        If HexDigit >= 0 Then
'            If BytePos > UBound(Bytes) Then
'                'Add some room, we'll add room for 4 more to decrease
'                'how often we end up doing this expensive step:
'                ReDim Preserve Bytes(UBound(Bytes) + 4)
'            End If
'            Bytes(BytePos) = Bytes(BytePos) * &H10 + HexDigit
'            Digits = Digits + 1
'        End If
'        If Digits = 2 Or HexDigit < 0 Then
'            If Digits > 0 Then BytePos = BytePos + 1
'            Digits = 0
'        End If
'    Next
'    If Digits = 0 Then BytePos = BytePos - 1
'    If BytePos < 0 Then
'        Bytes = "" 'Empty.
'    Else
'        ReDim Preserve Bytes(BytePos)
'    End If
'    HexToBytes = Bytes
'End Function
'
'Public Function HexToString(value As String)
'    Dim szTemp As String
'    szTemp = value
'
'    Dim szData As String
'    szData = ""
'    While Len(szTemp) > 0
'        szData = CHR(CLng("&h" & right(szTemp, 2))) & szData
'        If (Len(szTemp) = 1) Then
'            szTemp = left(szTemp, Len(szTemp) - 1)
'        Else
'            szTemp = left(szTemp, Len(szTemp) - 2)
'        End If
'    Wend
'    HexToString = szData
'End Function
'
'
'
'
