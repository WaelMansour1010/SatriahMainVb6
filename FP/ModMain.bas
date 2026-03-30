Attribute VB_Name = "ModMain"
Public Cn As New ADODB.Connection
Public SysSQLServerType As Integer
Public SysSQLServerName As String
Public SysSQLServerTypeTechnical As String
Public StrAppRegPath As String
Public SysSQLServerDataBaseName As String
Public SysSQLServerUserId As String
Public SysSQLServerUserpassword As String
Public FullConnectionString As String
  
 Public Sub Main()
 

     If Dir(App.Path & "\Connection.txt", vbNormal) <> "" Then
  Open App.Path & "\Connection.txt" For Input As #1
  Dim I As Integer

    Do Until EOF(1)
        Line Input #1, A
        'subsequent lines
If I = 0 Then
SysSQLServerName = A
 
  
ElseIf I = 1 Then
SysSQLServerDataBaseName = A
ElseIf I = 2 Then
SysSQLServerUserpassword = A
ElseIf I = 3 Then
SysSQLServerUserId = A

End If
       I = I + 1
    
    Loop

    Close #1
Else
MsgBox "Connection File Not Exist", vbCritical, ""
 End If
 
 

   Set Cn = New ADODB.Connection
    With Cn
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15


 
'SysSQLServerUserpassword = "Admin@123"
'SysSQLServerUserId = "sa"
'SysSQLServerDataBaseName = "Byte"
'SysSQLServerName = "allajazera.gotdns.com\mssqlserver,1433"
.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & SysSQLServerDataBaseName & ";Data Source=" & SysSQLServerName
'.ConnectionString = FullConnectionString
.Open
MsgBox "Êã ÇáÇÊÕÇá ÈäÍÇÍ", vbInformation, ""
End With

 FPFRM.Show

 End Sub
 Public Function LoadPictureFromDB(PictControl As Object, _
                                  rs As Object, _
                                  FieldName As String, _
                                  Optional ByRef StrFileName As String) As Boolean

    Dim oPict As StdPicture
    Dim sDir As String
    Dim sTempFile As String
    Dim iFileNum As Integer
    Dim lFileLength As Long
    Dim abBytes() As Byte
    Dim iCtr As Integer

    On Error GoTo ErrorHandler

    If Not TypeOf rs Is ADODB.Recordset Then Exit Function

    'sDir = GetTempDir
    If sDir = "" Then sDir = "C:\"
    sTempFile = sDir & "0X2341KLZX.dat"

    If Len(Dir$(sTempFile)) > 0 Then
        Kill sTempFile
    End If

    iFileNum = FreeFile
    Open sTempFile For Binary As #iFileNum
    lFileLength = LenB(rs(FieldName))
    abBytes = rs(FieldName).GetChunk(lFileLength)
    Put #iFileNum, , abBytes()
    Close #iFileNum

    If Not PictControl Is Nothing Then
        PictControl.Picture = LoadPicture(sTempFile)
        Kill sTempFile
    Else
        StrFileName = sTempFile
    End If

    LoadPictureFromDB = True
ErrorHandler:
End Function
Public Function SavePictureToDB(PictControl As Object, _
                                rs As Object, _
                                FieldName As String) As Boolean
    Dim oPict As StdPicture

    Dim sDir As String
    Dim sTempFile As String
    Dim iFileNum As Integer
    Dim lFileLength As Long

    Dim abBytes() As Byte
    Dim iCtr As Integer

    On Error GoTo ErrorHandler

    If Not TypeOf rs Is ADODB.Recordset Then Exit Function
    Set oPict = PictControl.Picture

    If oPict Is Nothing Then Exit Function

    'Save picture to temp file
    'sDir = App.Path
    If sDir = "" Then sDir = "C:\"
    sTempFile = sDir & "0X2341KLZX.dat"

    If Dir(sTempFile) <> "" Then
        Kill sTempFile
    End If

    SavePicture oPict, sTempFile

    'read file contents to byte array
    iFileNum = FreeFile
    Open sTempFile For Binary Access Read As #iFileNum
    lFileLength = LOF(iFileNum)
    ReDim abBytes(lFileLength)
    Get #iFileNum, , abBytes()
    'put byte array contents into db field
    rs.Fields(FieldName).AppendChunk abBytes()
    Close #iFileNum
    'Don't return false if file can't be deleted
    On Error Resume Next
    Kill sTempFile
    SavePictureToDB = True
ErrorHandler:
End Function


 
 Public Function SQLDate(ConvertDate As Date, Optional BolPutChar As Boolean = False, _
Optional BolPutSep As Boolean = True) As String

Dim StrTemp As String
Dim StrRes As String
Dim IntMonthB As Integer
Dim IntMonthE As Integer
Dim IntDay As Integer
Dim IntMonth As Integer
Dim IntYear As Integer
Dim StrMonthPrev As String
Dim StrTempqq As String

    IntDay = Day(ConvertDate)
    IntMonth = Month(ConvertDate)
    IntYear = Year(ConvertDate)
    StrMonthPrev = Choose(IntMonth, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    If BolPutSep = True Then
        StrRes = "" & Format(IntDay, "00") & "-" & StrMonthPrev & "-" & IntYear & ""
    Else
        StrRes = "" & Format(IntDay, "00") & " " & StrMonthPrev & " " & IntYear & ""
    End If
    If BolPutChar = False Then
        SQLDate = StrRes
    ElseIf BolPutChar = True Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrTemp = "#" & StrRes & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrTemp = "'" & StrRes & "'"
        End If
        SQLDate = StrTemp
    End If
 End Function



