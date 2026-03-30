Attribute VB_Name = "mMain"
Option Explicit
 Private funcHolder As New Dictionary
Private mlngStart As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
' Specify the position where the tracing instructions are inserted
Public Enum ProcPosition
    ProcEnter
    ProcExit
    ProcInside
End Enum

' Text indentation information
Private m_Indent As Long

' File information
Private m_iFile  As Integer

Public Const E_ERR_CS_TRACING_INIT = vbObjectError + 1265
'Private funcHolder As New Dictionary
Private Const S_ERR_CS_TRACING_INIT = "Could not initiate tracing"

Private Declare Sub InitCommonControls _
   Lib "comctl32.dll" ()

Public Sub xMain()
   
    InitCommonControls
   
    '   Dim f As New FRMPOS
    '   f.Show
   
End Sub
'********************
'Trace


' Member      : AxCsInitiateTrace
' Description : Instantiate the text file data
Public Sub AxCsInitiateTrace()

    Static bErrReported As Boolean

    On Error GoTo hErr
    
    ' Create a new file
    On Error Resume Next

    m_iFile = FreeFile
    
    Open "C:\log\TracingResults.txt" For Append As #m_iFile
    
    If Err.Number <> 0 Then
        If Not bErrReported Then
            bErrReported = True
            Err.Raise E_ERR_CS_TRACING_INIT + 1, "AxCsInitiateTrace", _
               "Could not open the text file."
        End If
        Exit Sub
    End If
    
    Exit Sub

hErr:
    Err.Raise E_ERR_CS_TRACING_INIT, "AxCsInitiateTrace", S_ERR_CS_TRACING_INIT

End Sub

' Member      : AxCsTerminateTrace
' Description : Used for cleaning purposes
Public Sub AxCsTerminateTrace()
    
    Close #m_iFile
    
    m_iFile = 0
    
End Sub

' Member      : AxCsTrace
' Description : Send information to the text file regarding the current method being processed
' Parameters  : ProjectName   - Name of the project which contains the method
'               ComponentName - Name of the component which contains the method
'               MemberName    - Name of the method being processed
'               TracePosition - Indicates the position within method body, either at start, at exit,
'                               or inside it, where the AxCsTxtTraceWatch method is called
' Notes       : For inside-member calls of AxCsTrace method, you can use the ProjectName
'               parameter to send tracing information.
Public Sub AxCsTrace(ByVal projectName As String, _
                     Optional ByVal componentname As String = "", _
                     Optional ByVal MemberName As String = "", _
                     Optional ByVal TracePosition As ProcPosition = ProcInside)

    Dim sTemp    As String
    Dim funcsign As String
    Dim Datestr  As String
    funcsign = projectName & "." & componentname & "."
    If InStr(1, MemberName, "(") > 0 Then
        funcsign = funcsign & Trim(Split(MemberName, "(")(0))
    Else
        funcsign = funcsign & Trim(MemberName)
    End If
    
    ' Make sure that tracing is initiated
    If m_iFile = 0 Then AxCsInitiateTrace

    ' The color, bold state and suffix can be customized below
    '    Dim t As SYSTEMTIME
    '    GetSystemTime t
    If TracePosition = ProcEnter Then
        sTemp = " - enter "
        If Not funcHolder.Exists(funcsign) Then
            mlngStart = GetTickCount
            funcHolder.Add funcsign, mlngStart
        End If
    ElseIf TracePosition = ProcExit Then
        Dim str As String
        str = "[not]"
        sTemp = " - exit"
        If funcHolder.Exists(funcsign) Then
            Dim EndTimer As Long
            Dim FuncTime As Long
            FuncTime = funcHolder(funcsign)
            EndTimer = (GetTickCount - FuncTime)
            str = "[" & EndTimer & "]"
            funcHolder.Remove funcsign
        End If
       
    Else
        sTemp = ""
    End If
   
    ' Decrease the indent for the current level of tracing information
    If TracePosition = ProcExit Then
        m_Indent = m_Indent - 4
    End If

    If Not (TracePosition = ProcInside) Then
        sTemp = projectName & "." & componentname & "." & funcsign & sTemp
    Else
        sTemp = projectName
    End If
    
    ' Add tracing information to the text file
    
    On Error Resume Next
    Print #m_iFile, "" & str & " [" & Format(Now, "HH:MM:SS") & "]   " & String$(m_Indent, ".") & sTemp
    
    ' Increase the indent for the next level of tracing information
    If TracePosition = ProcEnter Then
        m_Indent = m_Indent + 4
    End If
    
    AxCsTerminateTrace
    
End Sub

' Member      : AxCsDumpParamValue
' Description : Returns a formatted information about each method parameter
' Parameters  : sParamName    - Parameter name
'               vParamValue   - Parameter value
Public Function AxCsDumpParamValue(sParamName As String, vParamValue As Variant) As String

    Dim sRet$
    
    If IsObject(vParamValue) Then
    
        sRet = "Object"
        
    ElseIf IsArray(vParamValue) Then
    
        sRet = "Array"
        
    Else
    
        On Error Resume Next
        sRet = CStr(vParamValue)
        
        If Err.Number <> 0 Then
            sRet = "Not determined"
        Else
            If Len(sRet) > 13 Then sRet = left$(sRet, 10) & "..."
        End If
        
    End If
    
    AxCsDumpParamValue = sParamName & ": [" & sRet & "]"

End Function

Public Function closeRs(rs As ADODB.Recordset)
    On Error Resume Next
    If rs.State <> 0 Then
        rs.Close
    End If
End Function

