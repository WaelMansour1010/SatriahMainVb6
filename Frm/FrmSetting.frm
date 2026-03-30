VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmSetting 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "÷»ÿ Œ’«∆’ «·»«—ﬂÊœ"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "FrmSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ Ì«— «·ÿ«»⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   0
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   4290
      Width           =   4875
      Begin VB.ComboBox CboPrinters 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   300
         Width           =   4605
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»«—ﬂÊœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1305
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   75
      Width           =   4875
      Begin VB.CommandButton CmdStkFrst 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   300
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   637
         Width           =   300
      End
      Begin VB.ComboBox CbStk 
         Height          =   315
         ItemData        =   "FrmSetting.frx":038A
         Left            =   1815
         List            =   "FrmSetting.frx":038C
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   1695
      End
      Begin VB.ComboBox CbBarType 
         Height          =   315
         ItemData        =   "FrmSetting.frx":038E
         Left            =   1815
         List            =   "FrmSetting.frx":0390
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·«” Ìﬂ— : »—„“ ≈·Ï ⁄œœ «·«” Ìﬂ—«  ›Ì Ê—ﬁ… «·ÿ»«⁄…"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   3
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   990
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4530
         Picture         =   "FrmSetting.frx":0392
         Top             =   990
         Width           =   240
      End
      Begin VB.Label LblStkBegin 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   300
         Left            =   645
         TabIndex        =   28
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "” Ìﬂ— «·»œ«Ì…"
         Height          =   195
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·«” Ìﬂ—"
         Height          =   270
         Index           =   1
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   645
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·»«—ﬂÊœ"
         Height          =   270
         Index           =   0
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈ŸÂ«—  ‰’"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1035
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1410
      Width           =   4875
      Begin VB.TextBox TxtdownCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   623
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox TxtUpCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   225
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "√ŸÂ— «·‰’"
         Height          =   315
         Index           =   1
         Left            =   780
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   600
         Width           =   1125
      End
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "√ŸÂ— «·‰’"
         Height          =   315
         Index           =   0
         Left            =   750
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1125
      End
      Begin VB.ComboBox CboUpText 
         Height          =   315
         ItemData        =   "FrmSetting.frx":071C
         Left            =   1920
         List            =   "FrmSetting.frx":071E
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   1485
      End
      Begin VB.ComboBox CboDownText 
         Height          =   315
         ItemData        =   "FrmSetting.frx":0720
         Left            =   1920
         List            =   "FrmSetting.frx":0722
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   638
         Width           =   1485
      End
      Begin VB.CheckBox ChkDNLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‰’ «·”›·Ï"
         Height          =   240
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   1590
      End
      Begin VB.CheckBox ChkUpLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‰’ «·⁄·ÊÏ"
         Height          =   240
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈⁄œ«œ «·Œÿ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1050
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2475
      Width           =   4875
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   400
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   525
         Width           =   1830
         Begin VB.CheckBox ChkItalic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„«∆·"
            Height          =   225
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   135
            Width           =   810
         End
         Begin VB.CheckBox ChkBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—Ì÷"
            Height          =   225
            Left            =   945
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   135
            Width           =   810
         End
      End
      Begin VB.CommandButton CmdFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   300
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   615
         Width           =   300
      End
      Begin VB.Label LblFont 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   135
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰„ÿ"
         Height          =   210
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblFontSize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2505
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   615
         Width           =   420
      End
      Begin VB.Label LblFontName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2955
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   615
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÕÃ„"
         Height          =   195
         Index           =   1
         Left            =   2565
         TabIndex        =   11
         Top             =   345
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄"
         Height          =   195
         Index           =   0
         Left            =   4275
         TabIndex        =   10
         Top             =   345
         Width           =   255
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈⁄œ«œ «··Ê‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   705
      Index           =   1
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3555
      Width           =   4875
      Begin VB.CommandButton CmdBColor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   330
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton CmdFColor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   330
         Left            =   2670
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   255
         Width           =   390
      End
      Begin VB.Label LblBColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Left            =   660
         TabIndex        =   16
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Ê‰ «·Œ·›Ì…"
         Height          =   210
         Left            =   1650
         TabIndex        =   15
         Top             =   300
         Width           =   810
      End
      Begin VB.Label LblFColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   3090
         TabIndex        =   14
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Ê‰ «·Œÿ"
         Height          =   210
         Left            =   4155
         TabIndex        =   13
         Top             =   300
         Width           =   600
      End
   End
   Begin ImpulseButton.ISButton CmdEnd 
      Height          =   375
      Left            =   180
      TabIndex        =   29
      Top             =   5280
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmSetting.frx":0724
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   375
      Left            =   1155
      TabIndex        =   30
      Top             =   5280
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmSetting.frx":0ABE
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdPreview 
      Height          =   375
      Left            =   2160
      TabIndex        =   31
      Top             =   5280
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„⁄«Ì‰…"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmSetting.frx":0E58
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   0
      Top             =   810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StkCount As Integer
Dim StkCol As Integer
Dim TTP As clstooltip
Dim M_PrintType As Integer

Private Sub CboDownText_Change()
    On Error GoTo ErrTrap

    If CboDownText.ListIndex = 2 Then
        TxtdownCaption.Visible = True
    Else
        TxtdownCaption.Visible = False
    End If

    If Me.CboDownText.ListIndex = 0 Or Me.CboDownText.ListIndex = 1 Then
        Me.Chk(1).Visible = True
    Else
        Me.Chk(1).Visible = False
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub CboDownText_Click()
    CboDownText_Change
End Sub

Private Sub CboUpText_Change()

    If CboUpText.ListIndex = 2 Then
'        TxtUpCaption.Visible = True
    Else
    '    TxtUpCaption.Visible = False
    End If

    If Me.CboUpText.ListIndex = 0 Or Me.CboUpText.ListIndex = 1 Then
    '    Me.Chk(0).Visible = True
    Else
    '    Me.Chk(0).Visible = False
    End If

End Sub

Private Sub CboUpText_Click()
    CboUpText_Change
End Sub

Private Sub CbStk_Click()
    LblStkBegin.Caption = 1
End Sub

Private Sub ChkDNLbl_Click()
    On Error GoTo ErrTrap

    If ChkDNLbl.value = vbChecked Then
        CboDownText.locked = False
        TxtdownCaption.locked = False
    Else
        CboDownText.locked = True
        TxtdownCaption.locked = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkUpLbl_Click()
    On Error GoTo ErrTrap

    If ChkUpLbl.value = vbChecked Then
        CboUpText.locked = False
        TxtUpCaption.locked = False
    Else
        CboUpText.locked = True
        TxtUpCaption.locked = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdBColor_Click()
    On Error GoTo ErrTrap

    CDlg.color = LblBColor.backcolor
    CDlg.ShowColor
    LblBColor.backcolor = CDlg.color
ErrTrap:

End Sub

Private Sub CmdFColor_Click()
    On Error GoTo ErrTrap

    CDlg.color = LblFColor.backcolor
    CDlg.ShowColor
    LblFColor.backcolor = CDlg.color
ErrTrap:
End Sub

Private Sub CmdFont_Click()
    On Error GoTo ErrTrap
    CDlg.Flags = cdlCFBoth + cdlCFLimitSize
    CDlg.Min = 4
    CDlg.Max = 14
    CDlg.FontBold = LblFont.FontBold
    CDlg.FontItalic = LblFont.FontItalic
    CDlg.FontName = LblFont.FontName
    CDlg.fontsize = LblFont.fontsize
    CDlg.ShowFont

    If CDlg.FontName <> "" Then
        LblFont.FontName = CDlg.FontName
        LblFont.FontBold = CDlg.FontBold
        LblFont.FontItalic = CDlg.FontItalic
        LblFont.fontsize = CDlg.fontsize
        LblFontName.Caption = CDlg.FontName
        LblFontSize.Caption = CDlg.fontsize
        ChkBold.value = CInt(Abs(CDlg.FontBold))
        ChkItalic.value = CInt(Abs(CDlg.FontItalic))
    End If

ErrTrap:
End Sub

Private Sub CmdEnd_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
End Sub

Private Sub CmdOk_Click()
    On Error GoTo ErrTrap
    Dim i As Integer
    cBarcode.backcolor = LblBColor.backcolor
    cBarcode.ForeColor = LblFColor.backcolor
    LblFont.FontBold = CBool(Abs(ChkBold.value))
    LblFont.FontItalic = CBool(Abs(ChkItalic.value))

    Set cBarcode.Font = LblFont.Font
    cBarcode.ShowUpText = CBool(ChkUpLbl.value)
    cBarcode.ShowDNText = CBool(ChkDNLbl.value)
    cBarcode.BarCodeTyp = CbBarType.ItemData(CbBarType.ListIndex)
    cBarcode.StickerSize = CbStk.ItemData(CbStk.ListIndex)
    AddItems
    cBarcode.PrintPage 0
    Unload Me

    'cBarcode.Preview
    'Set cBarcode = Nothing
ErrTrap:
End Sub

Private Sub CmdPreview_Click()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim cPrinter As ClsPrinters
    cBarcode.backcolor = LblBColor.backcolor
    cBarcode.ForeColor = LblFColor.backcolor
    LblFont.FontBold = CBool(Abs(ChkBold.value))
    LblFont.FontItalic = CBool(Abs(ChkItalic.value))
    Set cBarcode.Font = LblFont.Font
    cBarcode.ShowUpText = CBool(ChkUpLbl.value)
    cBarcode.ShowDNText = CBool(ChkDNLbl.value)
    cBarcode.BarCodeTyp = CbBarType.ItemData(CbBarType.ListIndex)
    cBarcode.StickerSize = CbStk.ItemData(CbStk.ListIndex)

    If Me.CboPrinters.ListIndex >= 0 Then
        Set cPrinter = New ClsPrinters
        Set Printer = cPrinter.GetPrinter(Me.CboPrinters.text)
    End If

    'cBarcode.ClearItems
    'cBarcode.StickerBegin = Val(LblStkBegin.Caption)
    AddItems
    'cBarcode.AddItem FG.TextMatrix(2, FG.ColIndex("Code")), "yasser", "ahmed"

    'For I = 1 To 80
    '    If I <> 20 Then
    '        cBarcode.AddItem Left("12" & CStr(I) & "0000000000000000", 16), "BarCode" & I, "1254" & CStr(I)
    '        Else
    '        cBarcode.AddItem ""
    '    End If
    'Next
    cBarcode.MoveFirst
    'FrmParent.LblPageNum.Caption = cBarcode.PageNumber
    'FrmParent.LblPageCount.Caption = cBarcode.PageCount
    'BarCode.Top = 0
    'BarCode.Left = 1800
    'Unload Me
    'FrmParent.Show vbModal
    'SetParent BarCode.hWnd, FrmParent.PicParent.hWnd
ErrTrap:
End Sub

Private Sub CmdStkFrst_Click()
    On Error GoTo ErrTrap
    cBarcode.StickerSize = CbStk.ItemData(CbStk.ListIndex)
   ' FrmFirstStk.GridOfBarCode cBarcode.StickerCount, cBarcode.PaperCols
'    FrmFirstStk.show vbModal
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim StkSize As Long
    Dim cBarcode As ClsBarcode
    Dim IntDefIndex As Integer

    Set cBarcode = New ClsBarcode

    With CbBarType
        .AddItem "Code39"
        .ItemData(0) = 8
        .AddItem "Code128"
        .ItemData(1) = 9
        .AddItem "ExCode39"
        .ItemData(2) = 18
        .AddItem "Code93"
        .ItemData(3) = 36
        .AddItem "ExCode93"
        .ItemData(4) = 37
    End With

    LblBColor.backcolor = cBarcode.backcolor
    LblFColor.backcolor = cBarcode.ForeColor
    Set LblFont.Font = cBarcode.Font
    LblFontName.Caption = LblFont.FontName
    LblFontSize.Caption = Int(LblFont.fontsize)
    ChkBold.value = Abs(CInt(LblFont.FontBold))
    ChkItalic.value = Abs(CInt(LblFont.FontItalic))
    ChkUpLbl.value = Abs(CInt(cBarcode.ShowUpText))
    ChkDNLbl.value = Abs(CInt(cBarcode.ShowDNText))

    Select Case cBarcode.BarCodeTyp

        Case 8
            CbBarType.ListIndex = 0

        Case 9
            CbBarType.ListIndex = 1

        Case 18
            CbBarType.ListIndex = 2

        Case 36
            CbBarType.ListIndex = 3

        Case 37
            CbBarType.ListIndex = 4
    End Select

    'CbBarType.ListIndex = cBarcode.BarCodeTyp
    ChkUpLbl_Click
    ChkDNLbl_Click
    FillStkCount
    StkSize = cBarcode.StickerSize

    For i = 0 To CbStk.ListCount - 1

        If StkSize = CbStk.ItemData(i) Then
            CbStk.ListIndex = i
            Exit For
        End If

    Next

    With CboUpText
        .AddItem "ﬂÊœ «·’‰›"
        .AddItem "”⁄— «·’‰›"
        .AddItem "«”„ «·’‰› "
        .AddItem "»œÊ‰ "
        .ListIndex = 0
    End With

    With CboDownText
        .AddItem "ﬂÊœ «·’‰›"
        .AddItem "”⁄— «·’‰›"
        .AddItem "«”„ «·’‰› "
        .AddItem "»œÊ‰ "
        .ListIndex = 0
    End With

    GetBarcodeSetting
    Set CmdPreview.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Preview").Picture
    Set CmdEnd.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture

    'CbStk.ListIndex = CbStk.List(CbStk.ItemData(144))
    '------------------
    If Printers.count > 0 Then

        For i = 0 To Printers.count - 1
            CboPrinters.AddItem Printers(i).DeviceName

            If Printer.DeviceName = Printers(i).DeviceName Then
                IntDefIndex = i
                'Lbl(10).Caption = Printers(I).DriverName
            End If

        Next i

        CboPrinters.ListIndex = IntDefIndex
    Else
        'Lbl(3).Visible = True
        'DisableAll
    End If

    '------------------
    AddTip
    Set cBarcode = Nothing
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveBarcodeSetting
    Set cBarcode = Nothing
End Sub

Private Sub LblBColor_dblClick()
    CmdBColor_Click
End Sub

Private Sub LblFColor_dblClick()
    CmdFColor_Click
End Sub

Private Sub LblFontName_DblClick()
    CmdFont_Click
End Sub

Private Sub AddItems()
    'Adding Items
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim ItemCount As Integer
    Dim UpText As String
    Dim DownText As String
    cBarcode.ClearItems
    cBarcode.StickerBegin = Me.LblStkBegin.Caption

    If Me.PrintType = 0 Then

        With FrmPrintBarcode

            For RowNum = 1 To .Fg.Rows - 1

                If .Fg.Cell(flexcpChecked, RowNum, .Fg.ColIndex("Print")) = flexChecked Then

                    Select Case CboUpText.ListIndex

                        Case 0

                            If Me.Chk(0).Visible = True And Me.Chk(0).value = vbChecked Then
                                UpText = "Code" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("barcodeno"))
                            Else
                                UpText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("barcodeno"))
                            End If

                            UpText = UpText & "/" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("PartNo"))
                        
                        Case 1

                            If Me.Chk(0).Visible = True And Me.Chk(0).value = vbChecked Then
                                UpText = "price " & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Cost"))
                            Else
                                UpText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Cost"))
                            End If

                        Case 2
                          '  UpText = TxtUpCaption.text

                              If Me.Chk(0).Visible = True And Me.Chk(0).value = vbChecked Then
                                UpText = "Name " & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Name"))
                            Else
                                UpText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Name"))
                            End If
                            
                        Case 3
                            UpText = ""
                    End Select

                    Select Case CboDownText.ListIndex

                        Case 0

                            If Me.Chk(1).Visible = True And Me.Chk(1).value = vbChecked Then

                                DownText = "ﬂÊœ «·’‰›" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Code"))
                            Else

                                DownText = "ﬂÊœ «·’‰›" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Code"))
                            End If

                        Case 1

                            DownText = "price " & .Fg.TextMatrix(RowNum, .Fg.ColIndex("Cost"))

                        Case 2

                            DownText = TxtdownCaption.text

                        Case 3

                            DownText = ""
                    End Select

                    If val(.Fg.TextMatrix(RowNum, .Fg.ColIndex("Qty"))) > 0 Then

                        For ItemCount = 1 To val(.Fg.TextMatrix(RowNum, .Fg.ColIndex("Qty")))
                            cBarcode.AddItem .Fg.TextMatrix(RowNum, .Fg.ColIndex("Code")), UpText, DownText
                        Next ItemCount

                    End If
                End If

            Next RowNum

        End With

    ElseIf Me.PrintType = 1 Then

        With FrmPrintItemsBarcodes

            For RowNum = 1 To .Fg.Rows - 1

                If .Fg.Cell(flexcpChecked, RowNum, .Fg.ColIndex("Print")) = flexChecked Then

                    Select Case CboUpText.ListIndex

                        Case 0

                            If Me.Chk(0).Visible = True And Me.Chk(0).value = vbChecked Then
                                UpText = "ﬂÊœ «·’‰›" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("ItemCode"))
                            Else
                                UpText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("ItemCode"))
                            End If

                        Case 1

                            If Me.Chk(0).Visible = True And Me.Chk(0).value = vbChecked Then
                                UpText = "«·”⁄— " & .Fg.TextMatrix(RowNum, .Fg.ColIndex("SallingPrice"))
                            Else
                                UpText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("SallingPrice"))
                            End If

                        Case 2
                            UpText = TxtUpCaption.text

                        Case 3
                            UpText = ""
                    End Select

                    Select Case CboDownText.ListIndex

                        Case 0

                            If Me.Chk(1).Visible = True And Me.Chk(1).value = vbChecked Then

                                DownText = "ﬂÊœ «·’‰›" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("ItemCode"))
                            Else

                                DownText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("ItemCode"))
                            End If

                        Case 1

                            If Me.Chk(1).Visible = True And Me.Chk(1).value = vbChecked Then

                                DownText = "«·”⁄— " & .Fg.TextMatrix(RowNum, .Fg.ColIndex("SallingPrice"))
                            Else

                                DownText = "" & .Fg.TextMatrix(RowNum, .Fg.ColIndex("SallingPrice"))
                            End If

                        Case 2

                            DownText = TxtdownCaption.text

                        Case 3

                            DownText = ""
                    End Select

                    If val(.Fg.TextMatrix(RowNum, .Fg.ColIndex("Qty"))) > 0 Then

                        For ItemCount = 1 To val(.Fg.TextMatrix(RowNum, .Fg.ColIndex("Qty")))
                            'cBarcode.AddItem .Fg.TextMatrix(RowNum, .Fg.ColIndex("Code")), UpText, DownText
                            cBarcode.AddItem .Fg.TextMatrix(RowNum, .Fg.ColIndex("ItemCode")), UpText, DownText
                        Next ItemCount

                    End If
                End If

            Next RowNum

        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveBarcodeSetting()
    SaveSetting StrAppRegPath, "Barcode", "UptextType", CboUpText.ListIndex
    SaveSetting StrAppRegPath, "Barcode", "DowntextType", CboDownText.ListIndex
    SaveSetting StrAppRegPath, "Barcode", "UpCaption", TxtUpCaption.text
    SaveSetting StrAppRegPath, "Barcode", "DownCaption", TxtdownCaption.text
End Sub

Private Sub GetBarcodeSetting()
    CboUpText.ListIndex = GetSetting(StrAppRegPath, "Barcode", "UptextType", 0)
    CboDownText.ListIndex = GetSetting(StrAppRegPath, "Barcode", "DowntextType", 0)
    TxtUpCaption.text = GetSetting(StrAppRegPath, "Barcode", "UpCaption", "")
    TxtdownCaption.text = GetSetting(StrAppRegPath, "Barcode", "DownCaption", "")
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "Œ’«∆’ «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdStkFrst, "” Ìﬂ— «·»œ«Ì… ..." & Wrap & "· ÕœÌœ ” Ìﬂ— „⁄Ì‰ · »œ√ „‰Â ⁄„·Ì… «·ÿ»«⁄…" & Wrap & "  ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "Œ’«∆’ «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdFont, " ‰”Ìﬁ «·Œÿ ..." & Wrap & "·÷»ÿ  ‰”Ìﬁ«  «·Œÿ" & Wrap & "  ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "Œ’«∆’ «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdFColor, "·Ê‰ «·Œÿ ..." & Wrap & "· ÕœÌœ ·Ê‰ Œÿ «·»«—ﬂÊœ" & Wrap & "  ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "Œ’«∆’ «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdBColor, "·Ê‰ «·Œ·›Ì…..." & Wrap & "· ÕœÌœ ·Ê‰ Œ·›Ì… «·»«—ﬂÊœ" & Wrap & "  ≈÷€ÿ Â‰«", True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub FillStkCount()
    On Error GoTo ErrTrap
    CbStk.AddItem "1 «” Ìﬂ—", 0
    CbStk.ItemData(0) = 1
    CbStk.AddItem "2 «” Ìﬂ—", 1
    CbStk.ItemData(1) = 2
    CbStk.AddItem "4 «” Ìﬂ—", 2
    CbStk.ItemData(2) = 4
    CbStk.AddItem "8 «” Ìﬂ—", 3
    CbStk.ItemData(3) = 8
    CbStk.AddItem "10 «” Ìﬂ—", 4
    CbStk.ItemData(4) = 10
    CbStk.AddItem "12 «” Ìﬂ—", 5
    CbStk.ItemData(5) = 12
    CbStk.AddItem "14 «” Ìﬂ—", 6
    CbStk.ItemData(6) = 14
    CbStk.AddItem "16 (2*8)«” Ìﬂ—", 7
    CbStk.ItemData(7) = 16
    CbStk.AddItem "16 (4*4) «” Ìﬂ—", 8
    CbStk.ItemData(8) = 17
    CbStk.AddItem "21 «” Ìﬂ—", 9
    CbStk.ItemData(9) = 21
    CbStk.AddItem "24 (3*8) «” Ìﬂ—", 10
    CbStk.ItemData(10) = 24
    CbStk.AddItem "24 (4*6) «” Ìﬂ—", 11
    CbStk.ItemData(11) = 25
    CbStk.AddItem "28 «” Ìﬂ—", 12
    CbStk.ItemData(12) = 28
    CbStk.AddItem "36 «” Ìﬂ—", 13
    CbStk.ItemData(13) = 36
    CbStk.AddItem "40 «” Ìﬂ—", 14
    CbStk.ItemData(14) = 40
    CbStk.AddItem "48 «” Ìﬂ—", 15
    CbStk.ItemData(15) = 48
    CbStk.AddItem "56 «” Ìﬂ—", 16
    CbStk.ItemData(16) = 56
    CbStk.AddItem "72 «” Ìﬂ—", 17
    CbStk.ItemData(17) = 72
    CbStk.AddItem "96 «” Ìﬂ—", 18
    CbStk.ItemData(18) = 96
    CbStk.AddItem "102 «” Ìﬂ—", 19
    CbStk.ItemData(19) = 102
    CbStk.AddItem "108 (6*18) «” Ìﬂ—", 20
    CbStk.ItemData(20) = 108
    CbStk.AddItem "108(9*12) «” Ìﬂ—", 21
    CbStk.ItemData(21) = 109
    CbStk.AddItem "120 «” Ìﬂ—", 22
    CbStk.ItemData(22) = 120
    CbStk.AddItem "144 «” Ìﬂ—", 23
    CbStk.ItemData(23) = 144
    Exit Sub
ErrTrap:
End Sub

Public Property Get PrintType() As Integer
    PrintType = M_PrintType
End Property

Public Property Let PrintType(ByVal vNewValue As Integer)
    M_PrintType = vNewValue
End Property
