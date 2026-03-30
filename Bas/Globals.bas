Attribute VB_Name = "Globals"
Option Explicit

Public InvType As Integer

Public BillType As Integer

Public AlertCount As Integer

Public Due_Date As Date

Public AlertCountFree As Integer

Private Const PI    As Double = 3.14159265358979
Public Const WebUrl4_hisms = "https://www.hisms.ws/api.php"
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

Private Type POINTAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Declare Function Polygon _
                Lib "gdi32" (ByVal hDC As Long, _
                             lpPoint As POINTAPI, _
                             ByVal nCount As Long) As Long

' API Declarations
Private Declare Function GetSystemMetrics& _
                Lib "user32" (ByVal nIndex As Long)

Private Declare Function sndPlaySound _
                Lib "winmm.dll" _
                Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                       ByVal uFlags As Long) As Long
' Constants
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

' Declarations
Private ClsGradient As New CGradient

Public fX As Long

Public fY As Long

Public lngScaleX As Long

Public lngScaleY As Long

Public Declare Function GetSystemDirectory _
               Lib "kernel32" _
               Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                            ByVal nSize As Long) As Long

'and this function use this constant MAX_PATH
Public Const MAX_PATH = 260

Public system_path As String

Public connection_string As String

Public Sub DrawAngle(picDraw As PictureBox, _
                     ByVal fAngle As Single)
    Dim iSize       As Integer
    Dim iFillStyle  As Integer
    Dim lFillColor  As Long
    Dim lForeColor  As Long
    Dim lRet        As Long
    Dim uaPts(3)    As POINTAPI
    'Size arrow to best fit picDraw at any angle
    iSize = IIf(picDraw.ScaleHeight < picDraw.ScaleWidth, Int(picDraw.ScaleHeight / PI), Int(picDraw.ScaleWidth / PI))
    
    'Setup the 4 points of the arrow using the first point
    'as the center and the other points offset From the center.
    uaPts(0).X = picDraw.ScaleWidth / 2
    uaPts(0).Y = picDraw.ScaleHeight / 2
    uaPts(1).X = uaPts(0).X - iSize
    uaPts(1).Y = uaPts(0).Y - iSize
    uaPts(2).X = uaPts(0).X + iSize
    uaPts(2).Y = uaPts(0).Y
    uaPts(3).X = uaPts(0).X - iSize
    uaPts(3).Y = uaPts(0).Y + iSize
    
    'Rotate the arrow to the correct angle
    Call RotatePoints(uaPts(0), uaPts, fAngle)
    
    'Save picDraw settings
    iFillStyle = picDraw.FillStyle
    lFillColor = picDraw.FillColor
    lForeColor = picDraw.ForeColor
    
    'Setup picDraw to fill the arrow
    picDraw.FillStyle = vbFSSolid   'Solid Fill
    picDraw.FillColor = &HFFFFFF    'Inside = White
    picDraw.ForeColor = &H0&        'Border = Black
    
    'Draw the filled arrow
    lRet = Polygon(picDraw.hDC, uaPts(0), 4)
    
    'Restore picDraw settings
    picDraw.FillStyle = iFillStyle
    picDraw.FillColor = lFillColor
    picDraw.ForeColor = lForeColor

    'Free the memory
    Erase uaPts
    
End Sub

Private Sub RotatePoints(uAxisPt As POINTAPI, _
                         uRotatePts() As POINTAPI, _
                         fDegrees As Single)

    'Rotates an array of PointAPI points around a center point by fDegrees

    Dim lIdx        As Long
    Dim fDX         As Single
    Dim fDY         As Single
    Dim fRadians    As Single

    fRadians = fDegrees * RADS
    
    For lIdx = 0 To UBound(uRotatePts)
        fDX = uRotatePts(lIdx).X - uAxisPt.X
        fDY = uRotatePts(lIdx).Y - uAxisPt.Y
        uRotatePts(lIdx).X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
        uRotatePts(lIdx).Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    Next lIdx
    
End Sub
'Public Function ConArNum(ByVal strStringToConvert As String) As String
'
'On Error GoTo ErrorHandler
'//"0123456789 ””"
'strStringToConvert = Replace$(strStringToConvert, "0", "0")
'strStringToConvert = Replace$(strStringToConvert, "1", ChrW$(1633))
'strStringToConvert = Replace$(strStringToConvert, "2", ChrW$(1634))
'strStringToConvert = Replace$(strStringToConvert, "3", ChrW$(1635))
'strStringToConvert = Replace$(strStringToConvert, "4", ChrW$(1636))
'strStringToConvert = Replace$(strStringToConvert, "5", ChrW$(1637))
'strStringToConvert = Replace$(strStringToConvert, "6", ChrW$(1638))
'strStringToConvert = Replace$(strStringToConvert, "7", ChrW$(1639))
'strStringToConvert = Replace$(strStringToConvert, "8", ChrW$(1640))
'strStringToConvert = Replace$(strStringToConvert, "9", ChrW$(1641))
'
'ConArNum = strStringToConvert
'
'Exit Function
'ErrorHandler:
'ConArNum = vbNullString
'
'End Function
Public Function StringDotFormat(ByVal StrFormat As String, _
                                ParamArray aryPlaceHolders()) As String

    Dim intPlaceHolderIndex As Integer
    Dim strOutput As String
    strOutput = StrFormat
    For intPlaceHolderIndex = LBound(aryPlaceHolders) To UBound(aryPlaceHolders)
        strOutput = Replace(strOutput, "{" & intPlaceHolderIndex & "}", aryPlaceHolders(intPlaceHolderIndex) & "")
    Next
    StringDotFormat = strOutput
End Function
Public Sub SendSMSToClient(code, Msg As String)
    Dim s         As String
    Dim RsOptions As New ADODB.Recordset
    s = "select Cus_mobile from TblCustemers where  CusID = " & val(code)
    RsOptions.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If RsOptions.EOF Then
        MsgBox "⁄„Ì· €Ì— ’ÕÌÕ "
        Exit Sub
    End If
    Dim custPhone As String
    custPhone = RsOptions!Cus_mobile & ""
    RsOptions.Close
    
    If custPhone = "" Then
        MsgBox "  —ﬁ„  ·Ì›Ê‰ «·⁄„Ì· €Ì— ’ÕÌÕ "
        Exit Sub
    End If
    s = "SELECT SMSUserName, "
    s = s & "       SMSPassWord, "
    s = s & "       SenderName, "
    s = s & "       OPTWEB "
    s = s & "FROM dbo.TblOptions;"
   
    RsOptions.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    If val(RsOptions!OPTWEB & "") <> 4 Then
        MsgBox "·« Ì„ﬂ‰ «—”«· «·—”«·Â "
        Exit Sub
    End If
    
    If RsOptions!SMSUserName & "" = "" Then
        MsgBox "«”„ «·„” Œœ„ ·Œœ„Â «·—”«∆· €Ì— ’ÕÌÕ "
        Exit Sub
    End If
    
    If RsOptions!SMSPassWord & "" = "" Then
        MsgBox "ﬂ·„Â «·”—  ·Œœ„Â «·—”«∆· €Ì— ’ÕÌÕ "
        Exit Sub
    End If
    If RsOptions!SenderName & "" = "" Then
        MsgBox "«”„ «·„—”· ·Œœ„Â «·—”«∆· €Ì— ’ÕÌÕ "
        Exit Sub
    End If
    
    s = WebUrl4_hisms & "?send_sms&username=" & RsOptions!SMSUserName & "&password=" _
       & RsOptions!SMSPassWord & "&numbers=" & custPhone & "&sender=" & RsOptions!SenderName & "&message=" & Msg
    Dim Req
    Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")
    Req.Open "get", s, async:=False
    Req.send
    Dim Result As String
    Result = Req.responseText
    If InStr(1, Result, "-", vbTextCompare) Then
        Result = Split(Result, "-")(0)
    End If

    Select Case Result
        Case "1"
            MsgBox "«”„ «·„” Œœ„ €Ì— ’ÕÌÕ"
        Case "2"
            MsgBox "ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ… "
        Case "404"
            MsgBox "·„ Ì „ «œŒ«· Ã„Ì⁄ «·»—„ —«  «·„ÿ·Ê»… "
        Case "504"
            MsgBox "«·Õ”«» „⁄ÿ· "
        Case "3"
            MsgBox " „ «·«—”«· "
        Case "4"
            MsgBox " ·«  ÊÃœ «—ﬁ«„ "
        Case "5"
            MsgBox "·«  ÊÃœ —”«·Â"
        Case "6"
            MsgBox "„—”· Œ«ÿÏ¡ "
        Case "7"
            MsgBox " „—”· €Ì— „›⁄· "
        Case "8"
            MsgBox " «·—”«·Â »Â« ﬂ·„«  „„‰Ê⁄Â "
        Case "9"
            MsgBox " ·« ÌÊÃœ —’Ìœ "
        Case "10"
            MsgBox " «—ÌŒ Œ«ÿÏ¡ "
        Case "11"
            MsgBox "Êﬁ  Œ«ÿÏ¡ "

        Case Else
            MsgBox Req.responseText
    End Select
End Sub

