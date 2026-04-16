VERSION 5.00
Begin VB.Form FrmActivation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "License Activaton"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Command5 
      Caption         =   "Activate"
      Height          =   495
      Left            =   7200
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox SQlTxt 
      Height          =   2175
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtDexrypted 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   6480
      Width           =   6975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paste"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TxtLicense 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1320
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ›⁄Ì·"
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "GetCode By"
      Height          =   1695
      Left            =   9000
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton OptActtype 
         Caption         =   "Direct"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton OptActtype 
         Caption         =   "Email"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton OptActtype 
         Caption         =   "Sms"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtCode 
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ›⁄Ì·"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Activation Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label License 
      Caption         =   "License"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbl 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageAsLong Lib "user32" _
     Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long
Private Type tGUID
   l1 As Long
   l2 As Long
   l3 As Long
   l4 As Long
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" ( _
      lpGuid As tGUID _
   ) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" ( _
      lpGuid As tGUID, _
      ByVal lpString As String, _
      ByVal cbBytes As Integer _
   ) As Integer
Public Function GetNetworkConnectionMACAddress() As String

' Return the currently used network adapter's MAC address

' Syntax
'
' GetNetworkConnectionMACAddress()

    Dim oWMIService As Object
    Dim vAdapters As Variant
    Dim oAdapter As Object
    Dim lIndex As Long
    Dim lMatchIndex As Long
    Dim vResult As Variant
    
    ' Adapters are pulled from the Windows Management Instrumentation database
    ' The currently used adapter has a MAC address and an IP address that is not 0.0.0.0
    Set oWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    Set vAdapters = oWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each oAdapter In vAdapters
        If Not IsNull(oAdapter.MACAddress) And IsArray(oAdapter.IPAddress) Then
            lMatchIndex = -1
            For lIndex = 0 To UBound(oAdapter.IPAddress)
                If Not oAdapter.IPAddress(lIndex) = "0.0.0.0" Then
                    lMatchIndex = lIndex
                    Exit For
                End If
            Next lIndex
            If Not lMatchIndex < 0 Then
                GetNetworkConnectionMACAddress = oAdapter.MACAddress
            End If
        End If
   Next

End Function

 


Public Function CreateGUID() As String

' Create and return a unique GUID string.

   Dim GUID As tGUID
   Dim Temp As String
   Dim Result As Long
   Dim Length As Long
   
   Result = CoCreateGuid(GUID)
   If (Result = 0) Then
      Temp = StrConv(String(38, Chr(0)), vbUnicode)
      Length = StringFromGUID2(GUID, Temp, Len(Temp))
      Temp = StrConv(Temp, vbFromUnicode)
      If (Length > 0) Then
         If (Left(Temp, 1) = "{") Then Temp = Right(Temp, Len(Temp) - 1)
         If (Right(Temp, 1) = "}") Then Temp = Left(Temp, Len(Temp) - 1)
         Length = InStr(Temp, "-")
         Do While (Length <> 0)
            Temp = Left(Temp, Length - 1) & Right(Temp, Len(Temp) - Length)
            Length = InStr(Temp, "-")
         Loop
      Else
         Temp = ""
      End If
   End If
   CreateGUID = Temp

End Function
Function URLEncode(ByVal str As String) As String
    Dim intLen As Integer
    Dim X As Integer
    Dim curChar As Long
    Dim newStr As String

    intLen = Len(str)
    newStr = ""

    For X = 1 To intLen
        curChar = Asc(Mid$(str, X, 1))
          
        If (curChar < 48 Or curChar > 57) And (curChar < 65 Or curChar > 90) And (curChar < 97 Or curChar > 122) Then
            newStr = newStr & "%" & Hex(curChar)
        Else
            newStr = newStr & Chr(curChar)
        End If

    Next X
              
    URLEncode = newStr
End Function


Public Sub SendMessage(Optional msgstr As String = "", _
                       Optional Numbers As String = "")
    Dim t As String

    If msgstr = "" Then
        msgstr = txtMessage.Text
    End If

    If Numbers = "" Then
        Numbers = txtNumbers.Text
    End If

    ''t = send(UserName, URLEncode(Password), ConvertToUnicode(ConvertString(txtMessage.Text)), txtSender.Text, txtNumbers.Text)
    't = Send("966550015230 ", URLEncode("aljazeera10"), ConvertToUnicode(msgstr), txtSender.Text, Numbers)
 
    If msgstr = "" Then
        ShowResult (t)
    Else
        ShowResult t, 1
    End If

End Sub
Private Sub ShowResult(val As String, _
                       Optional outme As Integer = 0)

    If outme <> 0 Then Exit Sub

    Select Case val

        Case "1": MsgBox ("·ﬁœ  „   ⁄„·Ì… «—”«· «·—”«·…  »‰Ã«Õ") 'sent

        Case "2": MsgBox ("≈‰ —’Ìœﬂ ·œÏ „Ê»«Ì·Ì ﬁœ ≈‰ ÂÏ Ê·„ Ì⁄œ »Â √Ì —”«∆·. (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance = 0

        Case "3": MsgBox ("≈‰ —’Ìœﬂ «·Õ«·Ì ·« Ìﬂ›Ì ·≈ „«„ ⁄„·Ì… «·≈—”«·. (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance  not  enough"

        Case "4": MsgBox ("≈‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ··œŒÊ· ≈·Ï Õ”«» «·—”«∆· €Ì— ’ÕÌÕ ( √ﬂœ „‰ √‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ÂÊ ‰›”Â «·–Ì  ” Œœ„Â ⁄‰œ œŒÊ·ﬂ ≈·Ï „Êﬁ⁄ „Ê»«Ì·Ì)") 'mobile not found

        Case "5": MsgBox ("Â‰«ﬂ Œÿ√ ›Ì ﬂ·„… «·„—Ê— ( √ﬂœ „‰ √‰ ﬂ·„… «·„—Ê— «· Ì  „ ≈” Œœ«„Â« ÂÌ ‰›”Â« «· Ì  ” Œœ„Â« ⁄‰œ œŒÊ·ﬂ „Êﬁ⁄ „Ê»«Ì·Ì,≈–« ‰”Ì  ﬂ·„… «·„—Ê— ≈÷€ÿ ⁄·Ï —«»ÿ ‰”Ì  ﬂ·„… «·„—Ê— · ’·ﬂ —”«·… ⁄·Ï ÃÊ«·ﬂ »—ﬁ„ «·„—Ê— «·Œ«’ »ﬂ)") 'password error

        Case "6": MsgBox ("≈‰ ’›Õ… «·≈—”«· ·« ÃÌ» ›Ì «·Êﬁ  «·Õ«·Ì (ﬁœ ÌﬂÊ‰ Â‰«ﬂ ÿ·» ﬂ»Ì— ⁄·Ï «·’›Õ… √Ê  Êﬁ› „ƒﬁ  ··’›Õ… ›ﬁÿ Õ«Ê· „—… √Œ—Ï √Ê  Ê«’· „⁄ «·œ⁄„ «·›‰Ì ≈–« ≈” „— «·Œÿ√)") 'page not response try send again

        Case "12": MsgBox ("≈‰ Õ”«»ﬂ »Õ«Ã… ≈·Ï  ÕœÌÀ Ì—ÃÏ „—«Ã⁄… «·œ⁄„ «·›‰Ì")

        Case "13": MsgBox ("≈‰ ≈”„ «·„—”· «·–Ì ≈” Œœ„ Â ›Ì Â–Â «·—”«·… ·„ Ì „ ﬁ»Ê·Â. (Ì—ÃÏ ≈—”«· «·—”«·… »≈”„ „—”· ¬Œ— √Ê  ⁄—Ì› ≈”„ «·„—”· ·œÏ „Ê»«Ì·Ì)") 'sender not accept

        Case "14": MsgBox "≈‰ ≈”„ «·„—”· «·–Ì ≈” Œœ„ Â €Ì— „⁄—› ·œÏ „Ê»«Ì·Ì. (Ì„ﬂ‰ﬂ  ⁄—Ì› ≈”„ «·„—”· „‰ Œ·«· ’›Õ… ≈÷«›… ≈”„ „—”·)" 'sender name not activated

        Case "15": MsgBox "ÌÊÃœ —ﬁ„ ÃÊ«· Œ«ÿ∆ ›Ì «·√—ﬁ«„ «· Ì ﬁ„  »«·≈—”«· ·Â«. ( √ﬂœ „‰ ’Õ… «·√—ﬁ«„ «· Ì  —Ìœ «·≈—”«· ·Â« Ê√‰Â« »«·’Ì€… «·œÊ·Ì…)"

        Case "16": MsgBox "«·—”«·… «· Ì ﬁ„  »≈—”«·Â« ·«  Õ ÊÌ ⁄·Ï ≈”„ „—”·. (√œŒ· ≈”„ „—”· ⁄‰œ ≈—”«·ﬂ «·—”«·…)"

        Case "17": MsgBox "·„ Ì „ «—”«· ‰’ «·—”«·…. «·—Ã«¡ «· √ﬂœ „‰ «—”«· ‰’ «·—”«·… Ê«· √ﬂœ „‰  ÕÊÌ· «·—”«·… «·Ï ÌÊ‰Ì ﬂÊœ («·—Ã«¡ «· √ﬂœ „‰ «” Œœ«„ «·œ«·… ConvertToUnicode)"

        Case "-1": MsgBox "·„ Ì „ «· Ê«’· „⁄ Œ«œ„ (Server) «·≈—”«· „Ê»«Ì·Ì »‰Ã«Õ. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"

        Case "-2": MsgBox "·„ Ì „ «·—»ÿ „⁄ ﬁ«⁄œ… «·»Ì«‰«  (Database) «· Ì  Õ ÊÌ ⁄·Ï Õ”«»ﬂ Ê»Ì«‰« ﬂ ·œÏ „Ê»«Ì·Ì. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
    
        Case Else: MsgBox (val)
    End Select

End Sub

Private Sub Command1_Click()
TxtCode = CreateGUID
'SendMessage TxtCode, "966541793243"


End Sub
Public Function CryptRC4(sText As String, sKey As String) As String
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap     As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lIdx        As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    lI = 0
    lJ = 0
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
    Next
End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
End Function

Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long

    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
    Next
End Function

Public Function FromHexDump(sText As String) As String
    Dim lIdx            As Long

    For lIdx = 1 To Len(sText) Step 2
        FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid(sText, lIdx, 2)))
    Next
End Function
Private Sub Command2_Click()
    
 
Dim myWMI As Object, myObj As Object, Itm

Set myWMI = GetObject("winmgmts:\\.\root\cimv2")
Set myObj = myWMI.ExecQuery("SELECT * FROM " & _
                 "Win32_NetworkAdapterConfiguration " & _
                 "WHERE IPEnabled = True")
For Each Itm In myObj
    'MsgBox (Itm.IPAddress(0))
    TxtCode = (Itm.MACAddress)
      Dim sSecret     As String

    sSecret = ToHexDump(CryptRC4(TxtCode, "10111982"))
   TxtCode = sSecret
    'Debug.Print sSecret
    'Debug.Print CryptRC4(FromHexDump(sSecret), "10111982")
    
    Exit For
Next
End Sub
 
Private Sub Command3_Click()
'Clipboard.Clear
'Clipboard.SetText "Hello", vbCFText

If Clipboard.GetFormat(vbCFText) Then
Me.TxtLicense = Clipboard.GetText(vbCFText)
 
End If

Me.TxtDexrypted.Text = CryptRC4(FromHexDump(TxtLicense.Text), TxtCode.Text)

Me.SQlTxt.Text = Replace(TxtDexrypted.Text, "%%", vbNewLine)
End Sub

Private Sub Command4_Click()
Clipboard.Clear
Clipboard.SetText TxtCode.Text, vbCFText
 
End Sub

Private Sub Command5_Click()
On Error GoTo errortrap
    Dim lCount As Long
    Const EM_GETLINECOUNT = 186

    lCount = SendMessageAsLong(SQlTxt.hWnd, EM_GETLINECOUNT, 0, 0)
'    MsgBox lCount
    
For i = 0 To lCount - 1
   Dim myParas As Variant
    myParas = Split(SQlTxt, vbNewLine)
 StrSQL = myParas(i)
   If StrSQL <> "" Then
   Debug.Print StrSQL
 Cn.Execute StrSQL
End If
Next i
 
 MsgBox "Done", vbInformation, Me.Caption
Exit Sub
errortrap:
MsgBox "Error in Activation"
End Sub

Private Sub Form_Load()
Command2_Click
End Sub

Private Sub TxtCode_Change()
lbl.Caption = Len(TxtCode)
End Sub
