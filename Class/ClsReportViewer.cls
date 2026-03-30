VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VPRN_PDF As String = "MICROSOFT PRINT TO PDF"
Private Const VPRN_XPS As String = "MICROSOFT XPS DOCUMENT WRITER"
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'================== Printer helpers ==================



'================== Logging to App Folder ==================
Private gLogFilePath As String
'===== One-time: ÷⁄ Â–« «·ﬁ”„ ›Ì Module ⁄«„ =====


Private Const APP_NAME    As String = "MyApp"     ' ⁄œ¯· «”„  ÿ»Ìﬁﬂ ·Ê Õ«»»
Private Const REG_SECTION As String = "Printing"
Private Const REG_KEY     As String = "LastConfirmedPrinter"

Public Function GetLastConfirmedPrinter() As String
    On Error Resume Next
    GetLastConfirmedPrinter = GetSetting(APP_NAME, REG_SECTION, REG_KEY, "")
End Function

Public Sub SaveLastConfirmedPrinter(ByVal printerName As String)
    On Error Resume Next
    SaveSetting APP_NAME, REG_SECTION, REG_KEY, printerName
End Sub
'===== ‰Â«Ì… «·ﬁ”„ =====

Private Function EnsureLogDir() As String
    On Error Resume Next
    Dim base As String, p As String
    base = App.path
    If right$(base, 1) = "\" Then
        p = base & "Logs"
    Else
        p = base & "\Logs"
    End If
    If LenB(Dir$(p, vbDirectory)) = 0 Then MkDir p  ' √‰‘∆ «·„Ã·œ
    EnsureLogDir = p
End Function

Private Sub firea()

End Sub

Private Function IsPathWritable(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    Dim testFile As String, ff As Integer
    testFile = folderPath & "\._wtest.tmp"
    ff = FreeFile
    Open testFile For Output As #ff
    Print #ff, "ok";
    Close #ff
    Kill testFile
    IsPathWritable = True
    Exit Function
EH:
    IsPathWritable = False
End Function

Public Sub LogInit(Optional ByVal reportTitle As String = "", Optional ByVal extra As String = "")
    On Error Resume Next
    Dim base As String, logDir As String
    If LenB(reportTitle) = 0 Then reportTitle = "Report"

    logDir = EnsureLogDir()
    If Not IsPathWritable(logDir) Then
        base = Environ$("TEMP")
    Else
        base = logDir
    End If

    gLogFilePath = base & "\CR_Print_" & reportTitle & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".log"

    Dim head As String
    head = "=== Crystal Print Log Started ===" & vbCrLf & _
           "Time      : " & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf & _
           "Machine   : " & Environ$("COMPUTERNAME") & vbCrLf & _
           "User      : " & Environ$("USERNAME") & vbCrLf & _
           "App       : " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
           "LogFolder : " & base & vbCrLf & _
           "Extra     : " & extra & vbCrLf & _
           "----------------------------------" & vbCrLf & _
           ListInstalledPrintersText() & _
           "==================================" & vbCrLf
    AppendText gLogFilePath, head
End Sub

Public Sub LogMsg(ByVal Msg As String)
    On Error Resume Next
    If LenB(gLogFilePath) = 0 Then LogInit
    AppendText gLogFilePath, Time$ & "  " & Msg & vbCrLf
End Sub

Public Sub LogError(ByVal stage As String)
    On Error Resume Next
    Dim s As String
    s = "**ERROR @ " & stage & vbCrLf & _
        "Err.Number      : " & CStr(Err.Number) & vbCrLf & _
        "Err.Description : " & Err.Description & vbCrLf & _
        "LastDllError    : " & CStr(Err.LastDllError) & vbCrLf & _
        "Guess           : " & GuessCause(Err.Description) & vbCrLf & _
        "----------------------------------" & vbCrLf
    AppendText gLogFilePath, Time$ & "  " & s
End Sub

Private Sub AppendText(ByVal filePath As String, ByVal text As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Append As #ff
    Print #ff, text;
    Close #ff
End Sub

Private Function ListInstalledPrintersText() As String
    On Error Resume Next
    Dim i As Long, t As String, defName As String
    If Printers.count > 0 Then defName = Printer.DeviceName
    t = "Installed Printers (" & Printers.count & "):" & vbCrLf
    For i = 0 To Printers.count - 1
        t = t & " - " & Printers(i).DeviceName & " | " & Printers(i).drivername & " | " & Printers(i).Port
        If Printers(i).DeviceName = defName Then t = t & "  [DEFAULT]"
        t = t & vbCrLf
    Next i
    ListInstalledPrintersText = t & vbCrLf
End Function

Public Function PrinterSummary(p As Object) As String
    On Error Resume Next
    If p Is Nothing Then
        PrinterSummary = "(None)"
    Else
        PrinterSummary = p.DeviceName & " | " & p.drivername & " | " & p.Port
    End If
End Function

Public Function GuessCause(ByVal errDesc As String) As String
    Dim d As String: d = UCase$(errDesc)
    If InStr(d, "ERROR STARTING PRINT JOB") > 0 Then
        GuessCause = "·«  ÊÃœ ÿ«»⁄… ’«·Õ…/«› —«÷Ì… √Ê „‘ﬂ·… « ’«· »«·œ—«Ì›—/«·‘»ﬂ…."
    ElseIf InStr(d, "INVALID PAGE SIZE") > 0 Or InStr(d, "PAPER") > 0 Then
        GuessCause = "„ﬁ«” Ê—ﬁ €Ì— „œ⁄Ê„ - « —ﬂ «·„ﬁ«” ··œ—«Ì›— („Â„ ·ÿ«»⁄«  POS)."
    ElseIf InStr(d, "PORT") > 0 Then
        GuessCause = "„‰›– €Ì— ’«·Õ (WSD). «” Œœ„ Standard TCP/IP Port."
    ElseIf InStr(d, "ACCESS") > 0 Or InStr(d, "PERMISSION") > 0 Then
        GuessCause = "’·«ÕÌ«  «·„” Œœ„/Ã·”… RDP ·«  —Ï «·ÿ«»⁄…."
    ElseIf InStr(d, "NETWORK") > 0 Or InStr(d, "TIMEOUT") > 0 Or InStr(d, "COMMUNICATION") > 0 Then
        GuessCause = "«‰ﬁÿ«⁄ ‘»ﬂ…/ ⁄—Ì› ÿ«»⁄… ‘»ﬂ…."
    ElseIf InStr(d, "INVALID PRINTER") > 0 Or InStr(d, "PRINTER NOT FOUND") > 0 Then
        GuessCause = "«·ÿ«»⁄… €Ì— „À» … ⁄·Ï «·ÃÂ«“."
    ElseIf InStr(d, "MEMORY") > 0 Then
        GuessCause = "–«ﬂ—… «·ÿ«»⁄…/«·œ—«Ì›— ·«  ﬂ›Ì∫ Ã—¯» Spooling Ê ﬁ·Ì· DPI."
    Else
        GuessCause = "€Ì— „ƒﬂœó—«Ã⁄ Spooling/Advanced features Ê«·œ—«Ì›—."
    End If
End Function



Public Function SafeDefaultPrinterName() As String
    On Error Resume Next
    If Printers.count > 0 Then SafeDefaultPrinterName = Printer.DeviceName
End Function

Public Function GetVirtualPrinter() As Object
    Dim i As Long, nm As String
    For i = 0 To Printers.count - 1
        nm = UCase$(Trim$(Printers(i).DeviceName))
        If InStr(nm, VPRN_PDF) > 0 Or InStr(nm, VPRN_XPS) > 0 Then
            Set GetVirtualPrinter = Printers(i)
            Exit Function
        End If
    Next i
    Set GetVirtualPrinter = Nothing
End Function

Public Function GetAnyVirtualOrDefaultPrinter(ByVal cPrinter As Object) As Object
    On Error Resume Next
    Set GetAnyVirtualOrDefaultPrinter = GetVirtualPrinter()
    If GetAnyVirtualOrDefaultPrinter Is Nothing Then
        If Printers.count > 0 Then Set GetAnyVirtualOrDefaultPrinter = Printer
    End If
End Function

Public Function GetPrinterSmart(ByVal cPrinter As Object, ByVal PreferredName As String) As Object
    On Error GoTo Fallback
    Dim tgt As String, i As Long, nm As String

    tgt = UCase$(Trim$(PreferredName))
    If LenB(tgt) = 0 Then GoTo TryDefault

    If Not cPrinter Is Nothing Then
        Set GetPrinterSmart = cPrinter.GetPrinter(tgt)
        If Not GetPrinterSmart Is Nothing Then Exit Function
    End If

    For i = 0 To Printers.count - 1
        nm = UCase$(Trim$(Printers(i).DeviceName))
        If (nm = tgt) Or (InStr(nm, tgt) > 0) Then
            Set GetPrinterSmart = Printers(i)
            Exit Function
        End If
    Next i

TryDefault:
    If Printers.count > 0 Then
        Set GetPrinterSmart = Printer
        If Not GetPrinterSmart Is Nothing Then Exit Function
    End If

    Set GetPrinterSmart = GetVirtualPrinter()
    Exit Function

Fallback:
    Set GetPrinterSmart = Nothing
End Function

Public Function ExportToPdfAndOpen(ByVal xReport As CRAXDRT.Report) As Boolean
    On Error GoTo EH
    Dim base As String, tmpPdf As String, logDir As String

    logDir = EnsureLogDir()
    If IsPathWritable(logDir) Then
        base = logDir
    Else
        base = Environ$("TEMP")
    End If

    tmpPdf = base & "\CR_Temp_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    With xReport.ExportOptions
        .DestinationType = crEDTDiskFile
        .FormatType = crEFTPortableDocFormat
        .DiskFileName = tmpPdf
    End With

    xReport.Export False

    If LenB(Dir$(tmpPdf)) > 0 Then
        Shell "rundll32.exe url.dll,FileProtocolHandler " & CHR$(34) & tmpPdf & CHR$(34), vbNormalFocus
        ExportToPdfAndOpen = True
    End If
    Exit Function
EH:
    ExportToPdfAndOpen = False
End Function

Private Function IsShiftDown() As Boolean
    IsShiftDown = (GetAsyncKeyState(vbKeyShift) And &H8000) <> 0
End Function

'Public Function FireReport1(xReport As CRAXDRT.Report, _
'                            m_Target As PrintTarget, _
'                            Optional StrPreviewCaption As String = "", _
'                            Optional StrPrinterName As String = "", _
'                            Optional IntCopies As Long = 1, _
'                            Optional LngHelpContextID As Long = 0, _
'                            Optional path As String) As Boolean
'
'    Dim Frm As FrmReportViewer
'    Dim cPrinter As ClsPrinters
'    Dim cCompanyInfo As ClsCompanyInfo
'    On Error GoTo ErrTrap
'    Set Frm = New FrmReportViewer
'    Set cCompanyInfo = New ClsCompanyInfo
'    CreateLogo xReport
'
'    '------------------
'    If SystemOptions.UserInterface = ArabicInterface Then
'        CorrectArabicNumbers xReport
'    End If
'
'    '------------------
'    With Frm
'
'        If Trim(StrPreviewCaption) = "" Then
'            StrPreviewCaption = xReport.reportTitle
'        End If
'
'        '---------------------------------------------
'        .txtPath = path
'        xReport.ReportAuthor = "»«Ì  ··»—„ÃÌ« "
'        xReport.ReportComments = "BYTE"
'        xReport.ApplicationName = xReport.reportTitle
'        '---------------------------------------------
'        .PreviewCaption = StrPreviewCaption
'        .CRViewer.ReportSource = xReport
'
'        If m_Target = WindowTarget Then
'            .CRViewer.viewReport
'            Set .xReport = xReport
'
'            Do While .CRViewer.IsBusy = True
'                DoEvents
'            Loop
'
'            .CRViewer.Zoom 100 'cCompanyInfo.RepZoom
'            .WindowState = vbMaximized
'            While .CRViewer.IsBusy
'
'                DoEvents
'            Wend
'
'            If Trim(StrPreviewCaption) = "" Then
'                StrPreviewCaption = xReport.reportTitle
'            End If
'
'            .PreviewCaption = StrPreviewCaption
'            '---------------------------------
'            SetReportToolTips xReport
'            CorrectArabicNumbers xReport
'
'            '---------------------------------
'            If LngHelpContextID <> 0 Then
'                .HelpContextID = LngHelpContextID
'            End If
'
'            .show
'            .Refresh
'        Else
'            'Set the printer name
'            StrPrinterName = DefaultPrinter
'            Set cPrinter = New ClsPrinters
'
'            If StrPrinterName <> "" Then
'                Set ObjPrinter = cPrinter.GetPrinter(StrPrinterName)
'            ElseIf cCompanyInfo.DefaultPrinter <> "" Then
'                Set ObjPrinter = cPrinter.GetPrinter(cCompanyInfo.DefaultPrinter)
'            Else
'                Set ObjPrinter = Printer
'            End If
'
'            If Not ObjPrinter Is Nothing Then
'                xReport.SelectPrinter ObjPrinter.drivername, ObjPrinter.DeviceName, ObjPrinter.Port
'            End If
'
'            'xReport.PrintOut False, IntCopies
'
'            'If SystemOptions.Save_options = 2 Then
'            'P_Target = PrinterTarget
'            'Else
'            'P_Target = WindowTarget
'            'End If
'
'            xReport.PrintOutEx False, IntCopies
'            '   .CRViewer.PrintReport
'            Unload Frm
'            Set ObjPrinter = Nothing
'            Set cPrinter = Nothing
'        End If
'
'    End With
'
'    'Set xReport.Application = Nothing
'    'Set xReport = Nothing
'    FireReport1 = True
'    Exit Function
'ErrTrap:
'    FireReport1 = False
'    Msg = "ÕœÀ Œÿ« √À‰«¡  ÃÂÌ“ «· ﬁ—Ì— ··⁄—÷"
'    Msg = Msg & CHR(13) & Err.Description
'    Msg = Msg & CHR(13) & Err.Number
'    Msg = Msg & CHR(13) & Err.Source
'    Msg = Msg & CHR(13) & "FireReport"
'    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'
'End Function
'
Public Function FireReport(xReport As CRAXDRT.Report, _
                           m_Target As PrintTarget, _
                           Optional StrPreviewCaption As String = "", _
                           Optional StrPrinterName As String = "", _
                           Optional IntCopies As Long = 1, _
                           Optional LngHelpContextID As Long = 0, _
                           Optional path As String, _
                           Optional printerName As String = "", _
                           Optional SQLStr As String = "", _
                           Optional ByVal mFromDate As String = "", _
                           Optional ByVal mToDate As String = "", _
                           Optional ByVal mIsHeaderPrint As Boolean = False) As Boolean
    On Error GoTo ErrTrap

    Dim Frm As FrmReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
    Dim cPrinter As ClsPrinters
    Dim ObjPrinter As Object
    Dim StrPrinterToUse As String
    Dim printed As Boolean
    Dim wantDialog As Boolean

    Set Frm = New FrmReportViewer
    Set cCompanyInfo = New ClsCompanyInfo

    ' init log in App\Logs
    LogInit xReport.reportTitle, "Target=" & IIf(m_Target = WindowTarget, "Preview", "DirectPrint")

    CreateLogo xReport, , , , mIsHeaderPrint
    Frm.mFromDate = mFromDate
    Frm.mToDate = mToDate
    If SystemOptions.UserInterface = ArabicInterface Then CorrectArabicNumbers xReport

    With Frm
        If Trim$(StrPreviewCaption) = "" Then StrPreviewCaption = xReport.reportTitle
        .TXTSTRSQL = SQLStr
        .txtPath = path

        xReport.ReportAuthor = ""
        xReport.ReportComments = "Dynamic Byte"
        xReport.ApplicationName = xReport.reportTitle

        .PreviewCaption = StrPreviewCaption
        .CRViewer.ReportSource = xReport

        If m_Target = WindowTarget Then
            LogMsg "Preview mode: ViewReport"
            .CRViewer.viewReport
            Set .xReport = xReport
            Do While .CRViewer.IsBusy: DoEvents: Loop
            .CRViewer.Zoom GetSetting(StrAppRegPath, "Setting", "ReportZoom", 100)
            .WindowState = vbMaximized
            While .CRViewer.IsBusy: DoEvents: Wend
            .PreviewCaption = StrPreviewCaption
            SetReportToolTips xReport
            CorrectArabicNumbers xReport
            .show
            .Refresh
            LogMsg "Preview mode shown."
        Else
 ' ===== Direct Print =====
' 0) Â· ‰⁄—÷ œ«Ì·ÊÃ «·≈⁄œ«œ (Crystal Printer Setup) Õ”» ≈⁄œ«œﬂø
wantDialog = CBool(SystemOptions.ShowPrinterDialoge2)
LogMsg "Direct print. ShowPrinterDialoge2=" & CStr(wantDialog)

' 1) Õœœ «”„ «·ÿ«»⁄… «·„ﬁ’Êœ
If LenB(printerName) > 0 Then
    StrPrinterToUse = printerName
ElseIf LenB(StrPrinterName) > 0 Then
    StrPrinterToUse = StrPrinterName
ElseIf Not cCompanyInfo Is Nothing Then
    If LenB(cCompanyInfo.DefaultPrinter) > 0 Then StrPrinterToUse = cCompanyInfo.DefaultPrinter
End If
If LenB(StrPrinterToUse) = 0 Then StrPrinterToUse = SafeDefaultPrinterName()

Set cPrinter = New ClsPrinters
Set ObjPrinter = GetPrinterSmart(cPrinter, StrPrinterToUse)
LogMsg "Chosen printer: " & PrinterSummary(ObjPrinter)

If ObjPrinter Is Nothing Then
    MsgBox "·«  ÊÃœ ÿ«»⁄… ’«·Õ… ··ÿ»«⁄… «·„»«‘—…. ›⁄· Microsoft Print to PDF √Ê XPS.", vbExclamation
    GoTo FinishUp
End If

' Õ÷¯— «· ﬁ—Ì— Ê«—»ÿ «·ÿ«»⁄… «·„Œ «—… (»œÊ‰  À»Ì  „ﬁ«”/« Ã«Â)
xReport.DiscardSavedData
xReport.EnableParameterPrompting = False
xReport.SelectPrinter ObjPrinter.drivername, ObjPrinter.DeviceName, ObjPrinter.Port

' 2) ·Ê «·œÌ«·ÊÃ „ÿ·Ê»: «› Õ Crystal Printer Setup (‰›” «·œÌ«·ÊÃ ›Ì «·’Ê—…)
If wantDialog Then
    LogMsg "Opening Crystal Printer Setup dialog..."
    On Error Resume Next
    If Not Frm Is Nothing Then
        xReport.PrinterSetup Frm.hWnd
    Else
        xReport.PrinterSetup 0&
    End If
    On Error GoTo PrintFail1
    LogMsg "Crystal Printer Setup dialog closed."
    
    ' »⁄œ «·≈€·«ﬁ: «· ﬁÿ «”„ «·ÿ«»⁄… «·Õ«·Ì „‰ «· ﬁ—Ì— (√Ê „‰ ObjPrinter ·Ê «·Œ«’Ì… €Ì— „ «Õ…)
    Dim currentPrn As String, lastConfirmed As String
    On Error Resume Next
    currentPrn = xReport.printerName  '  ⁄„· ›Ì CRAXDRT √€·» «·≈’œ«—« 
    If LenB(currentPrn) = 0 Then currentPrn = ObjPrinter.DeviceName
    On Error GoTo PrintFail1

    lastConfirmed = GetLastConfirmedPrinter()

    ' ·Ê «·ÿ«»⁄… « €Ì—  ⁄‰ ¬Œ— „Ê«›ﬁ… „”Ã·… ? «”√· „—… Ê«Õœ… Ê«Õ›Ÿ
    If StrComp(currentPrn, lastConfirmed, vbTextCompare) <> 0 Then
        If MsgBox("Â·  —Ìœ «·ÿ»«⁄… «·¬‰ »Â–Â «·≈⁄œ«œ« ø", vbOKCancel + vbQuestion, " √ﬂÌœ «·ÿ»«⁄…") <> vbOK Then
            LogMsg "User cancelled after Printer Setup; aborting print."
            GoTo PrintedOK   ' Œ—ÊÃ Â«œ∆ »œÊ‰ ÿ»«⁄…
        End If
        SaveLastConfirmedPrinter currentPrn
        LogMsg "User confirmed printer '" & currentPrn & "' (will not prompt next time unless changed)."
    Else
        LogMsg "Same confirmed printer '" & currentPrn & "' ? skipping confirmation prompt."
    End If
End If

' 3) „Õ«Ê·… 1: ÿ»«⁄… ’«„ … (»œÊ‰ √Ì œ«Ì·ÊÃ« )
On Error GoTo PrintFail1
LogMsg "Attempt#1: Silent PrintOutEx starting..."
DoEvents: DoEvents: DoEvents
xReport.PrintOutEx False, IIf(IntCopies <= 0, 1, IntCopies)
printed = True
LogMsg "Attempt#1: success."
GoTo PrintedOK


PrintFail1:
            LogError "Attempt#1"
            ' 4) „Õ«Ê·… 2: ≈⁄«œ… —»ÿ ‰ŸÌ›
            On Error Resume Next
            LogMsg "Attempt#2: reset SelectPrinter '', '', '' then PrintOutEx"
            xReport.SelectPrinter "", "", ""
            Err.Clear
            On Error GoTo PrintFail2
            DoEvents
            xReport.PrintOutEx False, IIf(IntCopies <= 0, 1, IntCopies)
            printed = True
            LogMsg "Attempt#2 success."
            GoTo PrintedOK

PrintFail2:
            LogError "Attempt#2"
            ' 5) „Õ«Ê·… 3: ›Ê·-»«ﬂ ·ÿ«»⁄… ÊÂ„Ì… √Ê «·«› —«÷Ì…
            Dim altPrinter As Object
            Set altPrinter = GetAnyVirtualOrDefaultPrinter(cPrinter)
            LogMsg "Attempt#3 with fallback printer: " & PrinterSummary(altPrinter)

            If Not altPrinter Is Nothing Then
                On Error Resume Next
                xReport.SelectPrinter altPrinter.drivername, altPrinter.DeviceName, altPrinter.Port
                Err.Clear
                On Error GoTo FinalFail
                DoEvents
                xReport.PrintOutEx False, IIf(IntCopies <= 0, 1, IntCopies)
                printed = True
                LogMsg "Attempt#3 success."
                GoTo PrintedOK
            End If

FinalFail:
            LogError "Attempt#3"
            LogMsg "Falling back to ExportToPdfAndOpen..."
            If ExportToPdfAndOpen(xReport) Then
                MsgBox " ⁄–— »œ¡ «·ÿ»«⁄… «·„»«‘—…∫  „ ≈‰‘«¡ „·› PDF „ƒﬁ .", vbInformation
                printed = True
                LogMsg "PDF fallback success."
            Else
                MsgBox " ⁄–— »œ¡ „Â„… «·ÿ»«⁄…. —«Ã⁄ „·› «··ÊÃ.", vbExclamation
                LogMsg "PDF fallback FAILED."
            End If

PrintedOK:
            Unload Frm
            Set ObjPrinter = Nothing
            Set cPrinter = Nothing
        End If
    End With

FinishUp:
    FireReport = True
    LogMsg "FireReport finished. printed=" & CStr(printed)
    Exit Function

PrintFailDlg:
    LogError "CrystalPrintDialog"
    Resume PrintedOK

ErrTrap:
    FireReport = False
    LogError "FireReport-Outer"
    Dim Msg As String
    Msg = "ÕœÀ Œÿ√ √À‰«¡  ÃÂÌ“ «· ﬁ—Ì—." & vbCrLf & _
          Err.Description & vbCrLf & _
          "—ﬁ„: " & Err.Number & vbCrLf & _
          "«·„’œ—: " & Err.Source & vbCrLf & _
          "FireReport"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function



'
'
Private Sub fff()

End Sub



'Public Function FireReport(xReport As CRAXDRT.Report, _
'                           m_Target As PrintTarget, _
'                           Optional StrPreviewCaption As String = "", _
'                           Optional StrPrinterName As String = "", _
'                           Optional IntCopies As Long = 1, _
'                           Optional LngHelpContextID As Long = 0, _
'                           Optional path As String, Optional PrinterName As String = "", Optional SQLStr As String, Optional ByVal mFromDate As String = "", Optional ByVal mToDate As String = "", Optional ByVal mIsHeaderPrint As Boolean = False) As Boolean
'
''          If year(Date) = 2017 Then
''          GoTo ll
''          Else
''          Exit Function
''
''          End If
'If DateDiff("d", SystemOptions.CLockedDate, Date) > 0 Or SystemOptions.LockSystem = 10111982 Then
'
'    If SystemOptions.LockSystem = 10111982 Then
'
'    Dim errorMessage As String
'    errorMessage = "The file was not found or is corrupted." & vbCrLf & _
'                   "C:\Windows\System32\kernel32.dll" & vbCrLf & vbCrLf
'
'
'        MsgBox errorMessage, vbCritical + vbOKOnly, ""
'
'
'    End If
'
'End If
'll:
'    Dim Frm As FrmReportViewer
'    Dim cPrinter As ClsPrinters
'    Dim cCompanyInfo As ClsCompanyInfo
'    On Error GoTo ErrTrap
'    Set Frm = New FrmReportViewer
'    'Frm.mFromDate = FromDate
'    'Frm.mToDate = ToDate
'    Set cCompanyInfo = New ClsCompanyInfo
'    CreateLogo xReport, , , , mIsHeaderPrint
'
'Frm.mFromDate = mFromDate
'Frm.mToDate = mToDate
'
'
'    '------------------
'    If SystemOptions.UserInterface = ArabicInterface Then
'        CorrectArabicNumbers xReport
'    End If
'
'    '------------------
'    With Frm
'
'        If Trim(StrPreviewCaption) = "" Then
'            StrPreviewCaption = xReport.reporttitle
'        End If
'
'        '---------------------------------------------
'       .TXTSTRSQL = SQLStr
'        .txtpath = path
'            DoEvents
'    DoEvents
'    DoEvents
'        xReport.ReportAuthor = ""
'        xReport.ReportComments = "Dynamic Byte"
'        xReport.ApplicationName = xReport.reporttitle
'        '---------------------------------------------
'        .PreviewCaption = StrPreviewCaption
'        .CRViewer.ReportSource = xReport
'
'        If m_Target = WindowTarget Then
'
'            .CRViewer.viewReport
'            Set .xReport = xReport
'
'            Do While .CRViewer.IsBusy = True
'                DoEvents
'            Loop
'
'
'            .CRViewer.Zoom GetSetting(StrAppRegPath, "Setting", "ReportZoom", 100)  'cCompanyInfo.RepZoom
'            .WindowState = vbMaximized
'            While .CRViewer.IsBusy
'
'                DoEvents
'            Wend
'
'            If Trim(StrPreviewCaption) = "" Then
'                StrPreviewCaption = xReport.reporttitle
'            End If
'
'            .PreviewCaption = StrPreviewCaption
'            '---------------------------------
'            SetReportToolTips xReport
'            CorrectArabicNumbers xReport
'
'            '---------------------------------
'            If LngHelpContextID <> 0 Then
'                .HelpContextID = LngHelpContextID
'            End If
'
'            .show
'            .Refresh
'
'        Else
'            'Set the printer name
'            '«—”«· «·„ÿ»Œ
'If PrinterName = "" Then
'
'  StrPrinterName = DefaultPrinter
'  Else
'  StrPrinterName = PrinterName
'End If
'
'
'
'
'            Set cPrinter = New ClsPrinters
'
'            If StrPrinterName <> "" Then
'                Set ObjPrinter = cPrinter.GetPrinter(UCase(StrPrinterName))
'            ElseIf cCompanyInfo.DefaultPrinter <> "" Then
'                Set ObjPrinter = cPrinter.GetPrinter(cCompanyInfo.DefaultPrinter)
'            Else
'                Set ObjPrinter = Printer
'            End If
'
'            If Not ObjPrinter Is Nothing Then
'                xReport.SelectPrinter ObjPrinter.drivername, ObjPrinter.DeviceName, ObjPrinter.Port
'            End If
'
'            'xReport.PrintOut False, IntCopies
'
'            'If SystemOptions.Save_options = 2 Then
'            'P_Target = PrinterTarget
'            'Else
'            'P_Target = WindowTarget
'            'End If
'        DoEvents
'        DoEvents
'        DoEvents
'
'            xReport.PrintOutEx False, IntCopies
'            '   .CRViewer.PrintReport
'            Unload Frm
'            Set ObjPrinter = Nothing
'            Set cPrinter = Nothing
'        End If
'
'    End With
'
'    'Set xReport.Application = Nothing
'    'Set xReport = Nothing
'    FireReport = True
'
'    Exit Function
'ErrTrap:
'    FireReport = False
'    Msg = "ÕœÀ Œÿ« √À‰«¡  ÃÂÌ“ «· ﬁ—Ì— ··⁄—÷"
'    Msg = Msg & CHR(13) & Err.Description
'    Msg = Msg & CHR(13) & Err.Number
'    Msg = Msg & CHR(13) & Err.Source
'    Msg = Msg & CHR(13) & "FireReport"
'    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'
'End Function
'


























Public Function CreateLogo(xRport As CRAXDRT.Report, Optional Branch As Double = 0, Optional ByVal StrFileName2 As String = "", Optional ByVal mIID As Long = 0, Optional ByVal mIsHeaderPrint As Boolean = False) As Boolean
    Dim rs As ADODB.Recordset
    Dim BolShowLogo As Boolean
    Dim xLogo As CRAXDRT.OLEObject
    Dim StrFileName As String
    Dim MsgErr As String
    'IsHeaderPrint
Dim StrSQL As String
    On Error GoTo hErr
 Dim cOptions As ClsCompanyInfo
 Set cOptions = New ClsCompanyInfo
    Set rs = New ADODB.Recordset
 If SystemOptions.WorkWithBranchLogo = False Then
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
Else

 
   StrSQL = "SELECT     *  from TblBranchesData Where (branch_id = " & Current_branch & ")"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

End If

    If rs.BOF Or rs.EOF Then
        CreateLogo = False
        Exit Function
    End If
    
    BolShowLogo = IIf(IsNull(rs("ShowLogoInReports").value), 0, rs("ShowLogoInReports").value)
    'If StrFileName2 <> "" Then StrFileName = StrFileName2: GoTo NextStep
    If BolShowLogo = True And hide_logo = False Then
     If SystemOptions.WorkWithBranchLogo = False Then
        LoadPictureFromDB Nothing, rs, "CompanyLogo", StrFileName
      Else
         LoadPictureFromDB Nothing, rs, "branchLogo", StrFileName
      
      End If
End If






    BolShowLogo = IIf(IsNull(rs("ShowLogoInReports").value), 0, rs("ShowLogoInReports").value)

    If BolShowLogo = True And hide_logo = False Then

      
      
        
        If mIsHeaderPrint Then
            Set xLogo = xRport.Areas(2).Sections(1).AddPictureObject(StrFileName, 100, 100)
        Else
            Set xLogo = xRport.Areas(1).Sections(1).AddPictureObject(StrFileName, 100, 100)
        End If
       'xRport.Areas(2).Sections.count 1
  '     Set xLogo = xRport.Areas(1).Sections.Add(xRport.Areas(1).Sections.count + 1)
  '     Set xLogo = xRport.Areas(1).Sections(0).AddPictureObject(StrFileName, 100, 100)
       '.Sections(1).AddPictureObject(StrFileName, 100, 100)
        xLogo.Width = SystemOptions.logowidth - 500
        xLogo.Height = SystemOptions.logoHeight
        xLogo.left = 300
        xLogo.backcolor = vbWhite
        xLogo.BorderColor = 255
        xLogo.CloseAtPageBreak = True
        xLogo.HyperlinkText = "BYTE"
        xLogo.HyperlinkType = crHyperlinkWebsite
        xRport.Areas(1).Sections(1).SuppressIfBlank = True
        xRport.Areas(1).Sections(1).Height = xLogo.Height + 250
        CreateLogo = True
       

          
         
         
        
        
              
       
       
       Dim i As Long
           For i = 1 To xRport.FormulaFields.count
                Select Case xRport.FormulaFields.Item(i).Name
                Case "{@ArabCompanyName}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.ArabCompanyName & "'"
                Case "{@CompanyAddress}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.CompanyAddress & "'"
                Case "{@CompanyTel}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.CompanyTel & "'"
                Case "{@CompanyWebSite}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.CompanyWebSite & "'"
                Case "{@EngCompanyName}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.EngCompanyName & "'"
                Case "{@VATRegNo}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.VATRegNo & "'"
                Case "{@Company_Mobile}"
                        xRport.FormulaFields.Item(i).text = "'" & cOptions.Company_Mobile & "'"
            End Select
    Next i

        
    
 End If


    rs.Close
    Set rs = Nothing
    Exit Function
hErr:
    MsgErr = "Œÿ« ›Ï "
    MsgErr = MsgErr & CHR(13) & "CreateLogo"
    MsgErr = MsgErr & CHR(13) & Err.Description
    MsgErr = MsgErr & CHR(13) & Err.Number
    MsgErr = MsgErr & CHR(13) & Err.Source
    WriteInLogFile MsgErr
    CreateLogo = False
End Function

Private Sub SetReportToolTips(xReport As CRAXDRT.Report)
    Dim i As Long
    Dim j As Long
    Dim X As CRAXDRT.FieldObject
    Dim c As CRAXDRT.OLEObject

    'For I = 1 To xReport.Sections.count
    '    For J = 1 To xReport.Sections(I).ReportObjects.count
    '        If xReport.Sections(I).ReportObjects.Item(J).name Like "*CusID" Then
    '            xReport.Sections(I).ReportObjects.Item(J).ToolTipText = "ﬂÊœ «·⁄„Ì· √Ê «·„Ê—œ "
    '        End If
    '    Next J
    'Next I
End Sub

Private Sub CorrectArabicNumbers(xReport As CRAXDRT.Report)
    Dim i As Long, j As Long
    Dim xxx As ReportObjects
    Dim X As CRAXDRT.FieldDefinitions

    Exit Sub
    On Local Error GoTo ErrTrap

    If Not xReport Is Nothing Then

        For i = 1 To xReport.Sections.count
            For j = 1 To xReport.Sections(i).ReportObjects.count

                If xReport.Sections(i).ReportObjects.Item(j).Kind = CRAXDRT.CRFieldKind.crDatabaseField Then
                    MsgBox xReport.Sections(i).ReportObjects.Item(j).DatabaseFieldDisplayName & "," '& xReport.Sections(I).ReportObjects.Item(j).ValueType
                
                End If

            Next j
        Next i

    End If

    Exit Sub
ErrTrap:
    Stop
End Sub
