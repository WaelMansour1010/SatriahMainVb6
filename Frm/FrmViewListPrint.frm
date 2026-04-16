VERSION 5.00
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Begin VB.Form FrmViewListPrint 
   Caption         =   "Form2"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   8610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Height          =   3675
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   4005
      _cx             =   7064
      _cy             =   6482
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   18.0871212121212
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSPDF8LibCtl.VSPDF8 VSPDF81 
      Left            =   1740
      Top             =   900
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
   Begin VSReport8LibCtl.VSReport VSReport1 
      Left            =   2310
      Top             =   900
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "FrmViewListPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FG As Object

Private Sub Form_Resize()
    On Error Resume Next

    With Me.VSPrinter1
        .Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
    End With

End Sub

Private Sub VSPrinter1_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)

    If Button = vbRightButton Then

        'initialize pdf control
        With Me.VSPDF81
            .Title = " ﬁ—Ì— „Ã„⁄"
            .Creator = "»«Ì  ··»—„ÃÌ« "
            .Author = "»«Ì  ··»—„ÃÌ« "
            .Subject = "«· ﬁ—Ì— «·„Ã„⁄"
            .Keywords = "BYTE"
        End With
    
        Dim fname$
        fname = App.path & "\test.pdf"

        If Dir(fname) <> "" Then
            Kill fname
        End If
    
        VSPrinter1.Font.Charset = 178
        VSPDF81.ConvertDocument Me.VSPrinter1, fname
    
        ExportToRTF
    
        OpenFile fname
        OpenFile App.path & "\test.rtf"
    End If

End Sub

Private Sub ExportToRTF()
    VSPrinter1.Font.Charset = 178
    VSPrinter1.ExportFormat = vpxRTF
    VSPrinter1.ExportFile = App.path & "\test.rtf"

    With Me.VSPrinter1
        'set up styles
        .PaperBin = binAuto + &H1000&
        .FontName = "Tahoma"
        .FontSize = 18
        .IndentLeft = 0
        .SpaceAfter = "8pt"
        .Styles.Add "Title", vpsContent
        .FontSize = 12
        .IndentLeft = "0.25in"
        .Styles.Add "Normal", vpsContent
    
        ' set up page
        .Header = "Export Pictures||Page %d"
        .Footer = "Page %d||Export Pictures"
        .PageBorder = pbTopBottom
        .MarginLeft = "1.5in"
        .HdrFontSize = 9
        .HdrFontBold = True
        'start document
        .StartDoc
    
        .RenderControl = Me.FG.hWnd
        'export RTF header/footer
        ExportHeaderFooter VSPrinter1
    
        .EndDoc
    End With

End Sub

Private Sub ExportHeaderFooter(vp As VSPrinter)
    
    ' no RTF export file? no work!
    If Len(vp.ExportFile) = 0 Then Exit Sub
    If vp.ExportFormat < vpxRTF Then Exit Sub

    ' build rtf style string for headers and foooters
    Dim rtfStyle$
    rtfStyle = "\rtf1\ansi\ansicpg1252\deff0\deflang1033 " & "{\fonttbl{\f999 {{fname}};}}\li0\tqc\tx{{center}}\tqr\tx{{right}}\f999\fs{{fsize}}"
    vp.GetMargins
    rtfStyle = Replace(rtfStyle, "{{center}}", (vp.X1 + vp.X2) / 2 - vp.X1)
    rtfStyle = Replace(rtfStyle, "{{right}}", vp.X2 - vp.X1)
    rtfStyle = Replace(rtfStyle, "{{fname}}", vp.HdrFontName)
    rtfStyle = Replace(rtfStyle, "{{fsize}}", CInt(2 * vp.HdrFontSize))

    If vp.HdrFontBold Then rtfStyle = rtfStyle & "\b"
    If vp.HdrFontItalic Then rtfStyle = rtfStyle & "\i"
    If vp.HdrFontUnderline Then rtfStyle = rtfStyle & "\ul"

    ' output header field
    Dim rtf$, s$, v
    s = vp.Header

    If Len(s) Then
        s = Replace(s, "\", "\\")
        s = Replace(s, "%d", "{\field{\*\fldinst PAGE}}")
        v = Split(s, "|")
        rtf = "{\header{" & rtfStyle & " {{left}} \tab {{center}} \tab {{right}}\par }}"
        rtf = Replace(rtf, "{{left}}", v(0))

        If UBound(v) >= 1 Then rtf = Replace(rtf, "{{center}}", v(1))
        If UBound(v) >= 2 Then rtf = Replace(rtf, "{{right}}", v(2))
        rtf = Replace(rtf, "{{center}}", "")
        rtf = Replace(rtf, "{{right}}", "")
        vp.ExportRaw = rtf
    End If

    ' output footer field
    s = vp.Footer

    If Len(s) Then
        s = Replace(s, "\", "\\")
        s = Replace(s, "%d", "{\field{\*\fldinst PAGE}}")
        v = Split(s, "|")
        rtf = "{\footer{" & rtfStyle & " {{left}} \tab {{center}} \tab {{right}}\par }}"
        rtf = Replace(rtf, "{{left}}", v(0))

        If UBound(v) >= 1 Then rtf = Replace(rtf, "{{center}}", v(1))
        If UBound(v) >= 2 Then rtf = Replace(rtf, "{{right}}", v(2))
        rtf = Replace(rtf, "{{center}}", "")
        rtf = Replace(rtf, "{{right}}", "")
        vp.ExportRaw = rtf
    End If

End Sub

