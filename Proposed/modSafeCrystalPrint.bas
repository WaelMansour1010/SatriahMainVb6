Attribute VB_Name = "modSafeCrystalPrint"
Option Explicit

' ============================================================
' Proposed safe Crystal printing wrapper for VB6 / TSplus
' This file is NOT wired into the project automatically.
' Copy it manually into a new BAS module if you decide to use it.
' ============================================================

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum SafeCrystalTarget
    scPreview = 0
    scPrint = 1
End Enum

Private m_PrintBusy As Boolean
Private m_LastPrintTick As Long
Private m_LastPrintKey As String

Public Function SafeNormalizeCopies(ByVal RequestedCopies As Variant, _
                                    Optional ByVal ForceOneCopy As Boolean = True) As Long
    On Error GoTo EH

    If ForceOneCopy Then
        SafeNormalizeCopies = 1
    ElseIf IsNumeric(RequestedCopies) Then
        SafeNormalizeCopies = CLng(RequestedCopies)
        If SafeNormalizeCopies < 1 Then SafeNormalizeCopies = 1
        If SafeNormalizeCopies > 3 Then SafeNormalizeCopies = 3
    Else
        SafeNormalizeCopies = 1
    End If
    Exit Function

EH:
    SafeNormalizeCopies = 1
End Function

Public Function SafeShouldBlockDuplicate(ByVal PrintKey As String, _
                                         Optional ByVal CooldownMs As Long = 2500) As Boolean
    Dim nowTick As Long
    Dim age As Long

    nowTick = GetTickCount()
    age = nowTick - m_LastPrintTick
    If age < 0 Then age = 0

    If m_PrintBusy Then
        Debug.Print "SafeCrystalPrint: blocked because another print is still running."
        SafeShouldBlockDuplicate = True
        Exit Function
    End If

    If LenB(m_LastPrintKey) > 0 Then
        If StrComp(m_LastPrintKey, PrintKey, vbTextCompare) = 0 Then
            If age <= CooldownMs Then
                Debug.Print "SafeCrystalPrint: blocked duplicate print call. Age(ms)=" & CStr(age)
                SafeShouldBlockDuplicate = True
                Exit Function
            End If
        End If
    End If
End Function

Private Sub SafeRememberPrint(ByVal PrintKey As String)
    m_LastPrintKey = PrintKey
    m_LastPrintTick = GetTickCount()
End Sub

Private Function SafeReportKey(ByVal xReport As CRAXDRT.Report, _
                               ByVal Target As SafeCrystalTarget, _
                               ByVal PreferredPrinter As String) As String
    On Error Resume Next
    SafeReportKey = UCase$(Trim$(xReport.ReportTitle)) & "|" & _
                    CStr(Target) & "|" & _
                    UCase$(Trim$(PreferredPrinter))
End Function

Public Function SafeGetPrinter(ByVal PreferredPrinter As String) As Object
    Dim i As Long
    Dim tgt As String
    Dim nm As String

    On Error GoTo EH

    tgt = UCase$(Trim$(PreferredPrinter))

    If LenB(tgt) > 0 Then
        For i = 0 To Printers.Count - 1
            nm = UCase$(Trim$(Printers(i).DeviceName))
            If nm = tgt Or InStr(1, nm, tgt, vbTextCompare) > 0 Then
                Set SafeGetPrinter = Printers(i)
                Exit Function
            End If
        Next i
    End If

    If Printers.Count > 0 Then
        Set SafeGetPrinter = Printer
    End If
    Exit Function

EH:
    Set SafeGetPrinter = Nothing
End Function

Private Function SafeValidatePrinter(ByVal xPrinter As Object, _
                                     Optional ByVal ShowMsg As Boolean = True) As Boolean
    On Error GoTo EH

    If xPrinter Is Nothing Then
        If ShowMsg Then MsgBox "لا توجد طابعة صالحة متاحة لهذه الجلسة.", vbExclamation, App.Title
        Exit Function
    End If

    If LenB(Trim$(xPrinter.DeviceName)) = 0 Then
        If ShowMsg Then MsgBox "اسم الطابعة غير صالح.", vbExclamation, App.Title
        Exit Function
    End If

    SafeValidatePrinter = True
    Exit Function

EH:
    If ShowMsg Then MsgBox "تعذر التحقق من الطابعة الحالية.", vbExclamation, App.Title
End Function

Private Sub SafePrepareReport(ByVal xReport As CRAXDRT.Report, _
                              ByVal xPrinter As Object, _
                              Optional ByVal ForcePaperSize As Integer = 0, _
                              Optional ByVal ForceOrientation As Integer = 0)
    On Error Resume Next

    xReport.DiscardSavedData
    xReport.EnableParameterPrompting = False

    If Not xPrinter Is Nothing Then
        xReport.SelectPrinter xPrinter.DriverName, xPrinter.DeviceName, xPrinter.Port
    End If

    ' Do not force paper/orientation unless the caller explicitly asks.
    ' Forced mismatched settings are one of the common causes of abnormal pagination.
    If ForcePaperSize <> 0 Then xReport.PaperSize = ForcePaperSize
    If ForceOrientation <> 0 Then xReport.PaperOrientation = ForceOrientation
End Sub

Public Function SafeCrystalPreviewOrPrint(ByVal xReport As CRAXDRT.Report, _
                                          ByVal Target As SafeCrystalTarget, _
                                          Optional ByVal PreviewCaption As String = "", _
                                          Optional ByVal PreferredPrinter As String = "", _
                                          Optional ByVal RequestedCopies As Variant = 1, _
                                          Optional ByVal ForceOneCopy As Boolean = True, _
                                          Optional ByVal AllowPrinterSetupDialog As Boolean = False, _
                                          Optional ByVal CooldownMs As Long = 2500, _
                                          Optional ByVal ShowTrace As Boolean = True, _
                                          Optional ByVal ForcePaperSize As Integer = 0, _
                                          Optional ByVal ForceOrientation As Integer = 0) As Boolean
    On Error GoTo EH

    Dim Frm As FrmReportViewer
    Dim xPrinter As Object
    Dim copiesToPrint As Long
    Dim printKey As String

    If xReport Is Nothing Then
        If ShowTrace Then MsgBox "التقرير غير مهيأ للطباعة.", vbExclamation, App.Title
        Exit Function
    End If

    copiesToPrint = SafeNormalizeCopies(RequestedCopies, ForceOneCopy)
    printKey = SafeReportKey(xReport, Target, PreferredPrinter)

    If SafeShouldBlockDuplicate(printKey, CooldownMs) Then
        If ShowTrace Then Debug.Print "SafeCrystalPrint: duplicate call ignored. key=" & printKey
        Exit Function
    End If

    m_PrintBusy = True
    SafeRememberPrint printKey

    If ShowTrace Then
        Debug.Print "SafeCrystalPrint: start"
        Debug.Print "SafeCrystalPrint: report=" & xReport.ReportTitle
        Debug.Print "SafeCrystalPrint: target=" & CStr(Target)
        Debug.Print "SafeCrystalPrint: requested copies=" & CStr(RequestedCopies)
        Debug.Print "SafeCrystalPrint: actual copies=" & CStr(copiesToPrint)
    End If

    Set Frm = New FrmReportViewer
    If LenB(Trim$(PreviewCaption)) = 0 Then PreviewCaption = xReport.ReportTitle

    Frm.PreviewCaption = PreviewCaption
    Frm.CRViewer.ReportSource = xReport

    If Target = scPreview Then
        Frm.CRViewer.ViewReport

        Do While Frm.CRViewer.IsBusy
            DoEvents
        Loop

        Frm.WindowState = vbMaximized
        Frm.Show
        Frm.Refresh

        SafeCrystalPreviewOrPrint = True
        GoTo CleanExit
    End If

    Set xPrinter = SafeGetPrinter(PreferredPrinter)
    If Not SafeValidatePrinter(xPrinter, True) Then GoTo CleanExit

    SafePrepareReport xReport, xPrinter, ForcePaperSize, ForceOrientation

    If ShowTrace Then
        Debug.Print "SafeCrystalPrint: printer=" & xPrinter.DeviceName & " | " & xPrinter.DriverName & " | " & xPrinter.Port
    End If

    If AllowPrinterSetupDialog Then
        On Error Resume Next
        xReport.PrinterSetup Frm.hWnd
        On Error GoTo EH
    End If

    DoEvents
    DoEvents

    ' One direct print call only.
    ' No retry here, because some Crystal/TSplus paths can already have queued the job.
    xReport.PrintOutEx False, copiesToPrint

    If ShowTrace Then Debug.Print "SafeCrystalPrint: PrintOutEx sent successfully."
    SafeCrystalPreviewOrPrint = True

CleanExit:
    On Error Resume Next
    Set xPrinter = Nothing
    If Target <> scPreview Then Unload Frm
    Set Frm = Nothing
    m_PrintBusy = False
    Exit Function

EH:
    If ShowTrace Then
        Debug.Print "SafeCrystalPrint: error " & Err.Number & " - " & Err.Description
    End If
    m_PrintBusy = False
    MsgBox "فشلت عملية الطباعة الآمنة." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, App.Title
End Function

' ============================================================
' Level 1
' Use only to sanitize copies if you cannot change the print flow yet.
' Example:
'   xReport.PrintOutEx False, SafeNormalizeCopies(SystemOptions.NOOFPRINTCOPIESSALES, True)
' ============================================================

' ============================================================
' Level 2
' Use this if your current code already prepares the report/printer and
' you only want duplicate protection + one print call.
' ============================================================
Public Function SafeCrystalPrintDirect(ByVal xReport As CRAXDRT.Report, _
                                       Optional ByVal PreferredPrinter As String = "", _
                                       Optional ByVal RequestedCopies As Variant = 1, _
                                       Optional ByVal ShowTrace As Boolean = True) As Boolean
    SafeCrystalPrintDirect = SafeCrystalPreviewOrPrint(xReport, scPrint, "", PreferredPrinter, RequestedCopies, True, False, 2500, ShowTrace)
End Function

' ============================================================
' Level 3
' Full wrapper for both preview and direct print.
' Example:
'   Call SafeCrystalPreviewOrPrint(xReport, scPreview, "Invoice Preview")
'   Call SafeCrystalPreviewOrPrint(xReport, scPrint, "", cCompanyInfo.DefaultPrinter, 1, True, False)
' ============================================================
