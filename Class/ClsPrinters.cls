VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function OpenPrinter _
                Lib "winspool.drv" _
                Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                                      phPrinter As Long, _
                                      pDefault As Any) As Long

Private Declare Function ClosePrinter _
                Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Declare Function PrinterProperties _
                Lib "winspool.drv" (ByVal hWnd As Long, _
                                    ByVal hPrinter As Long) As Long

Private Declare Function DeviceCapabilities _
                Lib "winspool.drv" _
                Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
                                             ByVal lpPort As String, _
                                             ByVal iIndex As Long, _
                                             lpOutput As Any, _
                                             lpDevMode As Any) As Long

'ó ›Ì √⁄·Ï «·„ÊœÌÊ·:
Private Const VPRN_PDF As String = "Microsoft Print to PDF"
Private Const VPRN_XPS As String = "Microsoft XPS Document Writer"

Public Function PrintReportDirect(ByVal xReport As CRAXDRT.Report, _
                                  Optional ByVal PrinterName As String = "", _
                                  Optional ByVal Copies As Integer = 1) As Boolean
    On Error GoTo EH

    Dim cPrinter As New ClsPrinters
    Dim ObjPrinter As Object
    Dim nameToTry As String

    ' 0)  √ﬂœ ≈‰ ›ÌÂ √Ì ÿ«»⁄… „ «Õ…
    If Printers.count = 0 Then
        MsgBox "·«  ÊÃœ √Ì ÿ«»⁄… „ıÀ» … ⁄·Ï Â–« «·ÃÂ«“. ›⁄¯· Microsoft Print to PDF √Ê √÷› ÿ«»⁄….", vbExclamation
        Exit Function
    End If

    ' 1) Õœœ «”„ «·ÿ«»⁄… «·„›÷¯·… (·Ê « »⁄ ) √Ê «·œÌ›Ê· 
    If LenB(PrinterName) > 0 Then
        nameToTry = PrinterName
    Else
        nameToTry = DefaultPrinter
    End If

    ' 2) Ã—¯» »«· — Ì»: «·„ÿ·Ê» -> PDF -> XPS -> Default
    Set ObjPrinter = cPrinter.GetPrinter(UCase$(nameToTry))
    If ObjPrinter Is Nothing Then Set ObjPrinter = cPrinter.GetPrinter(UCase$(VPRN_PDF))
    If ObjPrinter Is Nothing Then Set ObjPrinter = cPrinter.GetPrinter(UCase$(VPRN_XPS))
    If ObjPrinter Is Nothing Then Set ObjPrinter = Printer ' ﬁœ  ﬂÊ‰ Nothing ·Ê „›Ì‘ œÌ›Ê· 

    If ObjPrinter Is Nothing Then
        MsgBox "·„ √Ãœ ÿ«»⁄… ’«·Õ…. —Ã«¡ ≈÷«›… ÿ«»⁄… (Virtual) À„ ≈⁄«œ… «·„Õ«Ê·….", vbExclamation
        Exit Function
    End If

    ' 3) «—»ÿ «· ﬁ—Ì— »«·ÿ«»⁄… «·„Œ «—…
    xReport.DiscardSavedData
    xReport.EnableParameterPrompting = False
    xReport.SelectPrinter ObjPrinter.drivername, ObjPrinter.DeviceName, ObjPrinter.Port

    ' 4) À»¯  «·Ê—ﬁ/«·« Ã«Â (⁄œ¯· Õ”»  ﬁ—Ì—ﬂ)
    On Error Resume Next
    xReport.PaperSize = crPaperA4
    xReport.PaperOrientation = crPortrait
    On Error GoTo EH

    DoEvents: DoEvents ' √ÕÌ«‰« „›Ìœ… „⁄ »⁄÷ «·œ—«Ì›—“

    ' 5) ÿ»«⁄… „»«‘—… »œÊ‰ Dialog
    xReport.PrintOutEx False, IIf(Copies <= 0, 1, Copies)
    PrintReportDirect = True
    GoTo Clean

EH:
    ' ó Œÿ… B:  ’œÌ— PDF „ƒﬁ  ·Ê «·ÿ»«⁄… «·„»«‘—… ›‘· 
    On Error Resume Next
    Dim tmpPdf As String
    tmpPdf = Environ$("TEMP") & "\CR_Temp_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    With xReport.ExportOptions
        .DestinationType = crEDTDiskFile
        .FormatType = crEFTPortableDocFormat
        .DiskFileName = tmpPdf
    End With
    xReport.Export False

    If LenB(Dir$(tmpPdf)) > 0 Then
        Shell "rundll32.exe url.dll,FileProtocolHandler " & CHR$(34) & tmpPdf & CHR$(34), vbNormalFocus
        MsgBox " ⁄–— «·ÿ»«⁄… «·„»«‘—…∫  „ ≈‰‘«¡ PDF „ƒﬁ  »œ·« „‰ –·ﬂ.", vbInformation
        PrintReportDirect = True
    Else
        MsgBox " ⁄–— »œ¡ „Â„… «·ÿ»«⁄…. " & Err.Number & " - " & Err.Description, vbExclamation
    End If

Clean:
    Set ObjPrinter = Nothing
    Set cPrinter = Nothing
End Function

Public Function GetPrinter(Optional StrPrintrName As String = "") As Object
    Dim i As Integer, nm As String, tgt As String
    On Error GoTo ErrTrap

    If LenB(StrPrintrName) = 0 Then
        Set GetPrinter = Printer
        Exit Function
    End If

    tgt = UCase$(Trim$(StrPrintrName))
    For i = 0 To Printers.count - 1
        nm = UCase$(Trim$(Printers(i).DeviceName))
        If (nm = tgt) Or (InStr(nm, tgt) > 0) Then
            Set GetPrinter = Printers(i)
            Exit Function
        End If
    Next i

    Exit Function
ErrTrap:
    Set GetPrinter = Nothing
End Function


Public Sub AddPrinter()
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    Shell "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
End Sub

Public Function ShowPrinterProperties(StrPrinterName As String, _
                                      LngParantHwnd As Long) As Boolean

    Dim ObjPrinter As Object
    Dim hPrinter As Long

    Set ObjPrinter = GetPrinter(StrPrinterName)

    If Not ObjPrinter Is Nothing Then
        Screen.MousePointer = vbArrowHourglass
        OpenPrinter ObjPrinter.DeviceName, hPrinter, ByVal 0&
        PrinterProperties LngParantHwnd, hPrinter
        ClosePrinter hPrinter
        Screen.MousePointer = vbDefault
        Exit Function
    End If

End Function
