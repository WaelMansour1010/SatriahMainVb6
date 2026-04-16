VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInstallMent 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «·√ﬁ”«ÿ"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   ControlBox      =   0   'False
   Icon            =   "FrmInstallMent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   6435
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
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  «·œ›⁄Â «·„ﬁœ„…"
      Height          =   615
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   1560
      Width           =   3135
      Begin VB.TextBox TxtAdvPayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   50
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·œ›⁄Â «·„ﬁœ„…"
         Height          =   345
         Index           =   4
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.TextBox TxtStartQast 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   0
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   4200
         Width           =   1380
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   60
         Width           =   1410
      End
      Begin VB.ComboBox CboPrecenType 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   420
         Width           =   1410
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   60
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   780
         Width           =   1410
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         Index           =   4
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1140
         Width           =   1410
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   5
         Left            =   60
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2220
         Width           =   1410
      End
      Begin VB.TextBox TxtDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   0
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   4680
         Width           =   1380
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   585
         Left            =   2130
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   7260
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker Dtp_First 
         Height          =   345
         Left            =   60
         TabIndex        =   13
         Top             =   2700
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarForeColor=   0
         CalendarTitleBackColor=   0
         CalendarTitleForeColor=   51455
         CustomFormat    =   "yyyy/M/d"
         Format          =   91684867
         CurrentDate     =   38031
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   975
         Index           =   3
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3120
         Width           =   3090
         _cx             =   5450
         _cy             =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   210
            Index           =   0
            Left            =   2415
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   345
            Width           =   630
         End
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘Â—"
            Height          =   225
            Index           =   1
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   345
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.TextBox Txt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   7
            Left            =   30
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   585
            Width           =   915
         End
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   225
            Index           =   2
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   345
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·› —…"
            Height          =   195
            Index           =   16
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   345
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —… »Ì‰ œ›⁄«  «· ﬁ”Ìÿ"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   0
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   0
            Width           =   1980
         End
      End
      Begin ImpulseButton.ISButton Cmd_Cal 
         Height          =   390
         Left            =   -30
         TabIndex        =   21
         Top             =   4980
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "√Õ”» «·√ﬁ”«ÿ"
         BackColor       =   8421504
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmInstallMent.frx":038A
         ColorButton     =   8421504
         ColorHoverText  =   16777215
         DrawFocusRectangle=   0   'False
         ColorToggledText=   16777215
         ColorToggledHoverText=   16777215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ’„ «·„ ›ﬁ ⁄·Ì…"
         Height          =   315
         Index           =   3
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "ﬁÌ„Â Ì „ «·« ›«ﬁ ⁄·ÌÂ« Œ·«›« ·ﬁÌ„… «·ﬁ”ÿ «·„Õ ”»…"
         Top             =   4680
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„… «·„ ›ﬁ ⁄·ÌÂ«"
         Height          =   315
         Index           =   2
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "ﬁÌ„Â Ì „ «·« ›«ﬁ ⁄·ÌÂ« Œ·«›« ·ﬁÌ„… «·ﬁ”ÿ «·„Õ ”»…"
         Top             =   0
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€ «·√”«”Ï"
         Height          =   225
         Index           =   31
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·“Ì«œ… («·›«∆œ…)"
         Height          =   225
         Index           =   35
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   1470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·›«∆œ…"
         Enabled         =   0   'False
         Height          =   225
         Index           =   10
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€ «·ﬂ·Ï"
         Height          =   225
         Index           =   14
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1260
         Width           =   1470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√ﬁ”«ÿ"
         Height          =   345
         Index           =   22
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2220
         Width           =   1470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ √Ê· ﬁ”ÿ"
         Height          =   345
         Index           =   39
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2820
         Width           =   1470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„… «·„ ›ﬁ ⁄·ÌÂ«"
         Height          =   315
         Index           =   19
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "ﬁÌ„Â Ì „ «·« ›«ﬁ ⁄·ÌÂ« Œ·«›« ·ﬁÌ„… «·ﬁ”ÿ «·„Õ ”»…"
         Top             =   4200
         Width           =   1425
      End
      Begin VB.Label LblID 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Left            =   1290
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   7260
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LblNoteID 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   7260
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Fg 
      Height          =   3300
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   3210
      _cx             =   5662
      _cy             =   5821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmInstallMent.frx":0724
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin ImpulseButton.ISButton Cmdsave 
      Height          =   375
      Left            =   1050
      TabIndex        =   3
      Top             =   5580
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      BackStyle       =   0
      ButtonImage     =   "FrmInstallMent.frx":087B
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5580
      Width           =   855
      _ExtentX        =   1508
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
      BackStyle       =   0
      ButtonImage     =   "FrmInstallMent.frx":0C15
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   6315
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Ã„«·Ï «·√ﬁ”«ÿ"
      Height          =   405
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3285
      Width           =   3210
   End
   Begin VB.Label LblTotalQasts 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1830
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   3210
   End
End
Attribute VB_Name = "FrmInstallMent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldGrdValue As Variant

Public Frm As Form

Private Sub CboPrecenType_Change()
    On Error GoTo ErrTrap
    With CboPrecenType
        If .ListIndex > -1 Then
            Select Case .ItemData(.ListIndex)
                Case 1
                    lbl(10).Caption = "‰”»… «·›«∆œ…"

                Case 2
                    lbl(10).Caption = "ﬁÌ„… «·“Ì«œ…"
            End Select

            CalPre
        End If
    End With
    ''//////////////////

    If CboPrecenType.ListIndex > -1 And CboPrecenType.ListIndex < 2 Then
        lbl(10).Enabled = True
        Txt(3).Enabled = True
    Else
        lbl(10).Enabled = False
        Txt(3).Enabled = False
    End If

    Fg.Clear flexClearScrollable, flexClearEverything
    LblTotalQasts.Caption = ""
    Exit Sub
ErrTrap:
End Sub

Private Sub CboPrecenType_Click()
    CboPrecenType_Change
End Sub

Private Sub Cmd_Cal_Click()
    Calculations True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim RowNum As Integer
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsDetalis As ADODB.Recordset
    Dim XFg As VSFlex8UCtl.vsFlexGrid
    Dim StrTemp As String
    Dim i As Long

    With Fg

        If .TextMatrix(1, .ColIndex("Serial")) = "" Then
            Msg = "ÌÃ» Õ”«» ﬁÌ„… «·√ﬁ”«ÿ ﬁ»· «·Õ›Ÿ" & Chr(13)
            Msg = Msg + "·Õ”«» ﬁÌ„… «·√ﬁ”«ÿ «÷€ÿ ›Êﬁ («Õ”» «·√ﬁ”«ÿ)"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
        If Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value")), Decimal_Places1) < Round((val(Me.Txt(4).text)), Decimal_Places1) - val(Me.TxtDiscount.text) - val(Me.TxtAdvPayment.text) Then
            Msg = "„Ã„Ê⁄ «·√ﬁ”«ÿ «·„”Ã·… ·« ”«ÊÏ «·„»·€ «·ﬂ·Ï " & Chr(13) & "«·„›—Ê÷  ﬁ”ÌÿÂ... " & Chr(13) & "»—Ã«¡ „—«Ã⁄… ﬁÌ„ «·√ﬁ”ÿ «·„”Ã·…"
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

    End With

    If Not Frm Is Nothing Then
        If val(Me.Txt(1).text) <> val(Me.Txt(4).text) Then
            Msg = "ÌÃ» „·«ÕŸ… «‰ «·„»·€ «·√”«”Ï «·√Ã· ⁄·Ï «·⁄„Ì·"
            Msg = Msg & Chr(13) & "√’»Õ ·«Ì”«ÊÏ ≈Ã„«·Ï «·√ﬁ”«ÿ »⁄œ(Õ”«» ﬁÌ„… «·›«∆œ…)"
            Msg = Msg & Chr(13) & "·–« ›«‰ «·»—‰«„Ã ”Ê› ÌﬁÊ„ "
        
        End If

        Frm.LblPrecenType.Caption = Me.CboPrecenType.text
        Frm.LblPrecenType.Tag = Me.CboPrecenType.ListIndex
        Frm.LblPrecenValue.Caption = val(Me.Txt(3).text)
        Frm.LblInstallTotal.Caption = val(LblTotalQasts.Caption) + val(TxtAdvPayment.text)
        Frm.LblInstallCount.Caption = Me.Fg.Aggregate(flexSTCount, Fg.FixedRows, Fg.ColIndex("Value"), Fg.Rows, Fg.ColIndex("Value"))
        Frm.LblFirstInstallDate.Caption = DisplayDate(Dtp_First.value)
    
        For i = 0 To Me.OptInt.count

            If Me.OptInt(i).value = True Then
                Frm.LblInstallmentType.Caption = Me.OptInt(i).Caption
                Frm.LblInstallmentType.Tag = i
                Exit For
            End If

        Next

        Frm.LblStartValue.Caption = val(TxtStartQast.text)
        Frm.LblInstallSeprator.Caption = val(Me.Txt(7).text)
        Frm.LblDiscount.Caption = val(Me.TxtDiscount)
     frmsalebill.TotalQest.text = val(Txt(4).text)
     If val(val(Txt(5).text)) <> 0 Then
     frmsalebill.QstValue.text = val(Txt(4).text) / val(Txt(5).text)
     End If
     frmsalebill.QstNo.text = val(Txt(5).text)
 
     frmsalebill.QestStartDate.value = Dtp_First.value
     If OptInt(0).value = True Then
     frmsalebill.OptInt(0).value = True
     frmsalebill.QestEndtDate.value = DateAdd("d", val(Txt(5).text) - 1, Dtp_First.value)
     ElseIf OptInt(1).value = True Then
     frmsalebill.OptInt(1).value = True
     frmsalebill.QestEndtDate.value = DateAdd("m", val(Txt(5).text) - 1, Dtp_First.value)
     ElseIf OptInt(2).value = True Then
     frmsalebill.OptInt(2).value = True
     frmsalebill.QestEndtDate.value = DateAdd("yyyy", val(Txt(5).text) - 1, Dtp_First.value)
     End If
      frmsalebill.QestEndtDate_Change
      frmsalebill.QestStartDate_Change
        Frm.LblAdvPayment.Caption = val(Me.TxtAdvPayment)
     
        With Frm.FgInstallments
            .Rows = Me.Fg.Rows

            For i = 1 To Me.Fg.Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("Value")) = Me.Fg.TextMatrix(i, Me.Fg.ColIndex("Value"))
                .TextMatrix(i, .ColIndex("Due_Date")) = Me.Fg.TextMatrix(i, Me.Fg.ColIndex("Due_Date"))
            Next i

            Frm.FgInstallments.AutoSize 0, Frm.FgInstallments.Cols - 1, False
        End With

    End If

    Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    With Me.Fg
        OldGrdValue = .TextMatrix(Row, Col)
    End With

End Sub

Private Sub Form_Activate()

    If Me.Tag = "R" Then
        Fram.Enabled = False
        Cmdsave.Visible = False
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BGround As ClsBackGroundPic
    CenterForm Me

    FormPostion Me, GetPostion

    With CboPrecenType
        .AddItem "‰”»… „∆ÊÌ…", 0
        .ItemData(0) = 1
        .AddItem "ﬁÌ„… À«» …", 1
        .ItemData(1) = 2
        .AddItem "·«ÌÊÃœ", 2
        .ItemData(2) = 3
        .ListIndex = 2
    End With

    Set BGround = New ClsBackGroundPic
    Fg.WallPaper = BGround.Picture
    Dtp_First.value = Date
    OptInt(1).value = True
    Txt(7).text = 1
    'Txt(5).text = 12

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub CalPre()
    On Error GoTo ErrTrap
    Dim SngUseValue As Single  'ﬁÌ„… «·›«∆œ…
    Dim SngAllValue As Single

    'Õ”«» ﬁÌ„… «·›«∆œ…
    If Me.CboPrecenType.ListIndex > -1 Then
        If Me.CboPrecenType.ItemData(CboPrecenType.ListIndex) = 1 Then
            SngUseValue = (val(Txt(1).text) * val(Txt(3).text)) / 100
        ElseIf Me.CboPrecenType.ItemData(CboPrecenType.ListIndex) = 2 Then
            SngUseValue = val(Me.Txt(3).text)
        ElseIf Me.CboPrecenType.ItemData(CboPrecenType.ListIndex) = 3 Then
            SngUseValue = 0
        End If
    End If

    Txt(4).text = (SngUseValue)
    '«·„»·€ «·ﬂ·Ï («·–Ï ”Ê› Ìﬁ”ÿ) Ì”«ÊÏ Õ”«» ﬁÌ„…
    '«·›«∆œ… „⁄ ﬁÌ„… «·„»·€ «·„ »ﬁÏ
    SngAllValue = SngUseValue + val(Txt(1).text)
    Txt(4).text = (SngAllValue)

    frmsalebill.lblInstComm = SngUseValue
    Exit Sub
ErrTrap:
End Sub

Private Sub Calculations(Optional WithMsg As Boolean = False)
    On Error GoTo ErrTrap
    Dim SngAllValue As Single
    Dim i  As Integer
    Dim IntNoOFQast As Integer
    Dim IntRes As Integer
    Dim SngOnePor As Single
    Dim FirstDate As Date
    Dim PreDate As Date
    Dim NewDate As Date
    Dim DateInterval As String
    Dim DateNumber As Integer
    Dim Msg As String

    If CboPrecenType.ListIndex = 0 Or CboPrecenType.ListIndex = 1 Then
        If Txt(3).text = "" Then
            Msg = "›Ì Õ«·… ÊÃÊœ ›«∆œ… ÌÃ»  ÕœÌœ ﬁÌ„… √Ê ‰”»… Â–Â «·›«∆œ…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If Txt(3).Enabled = True Then
                Txt(3).SetFocus
            End If

            Exit Sub
        End If

        If Not IsNumeric(Txt(3).text) Then
            Msg = " ﬁÌ„… √Ê ‰”»… Â–Â «·›«∆œ… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If Txt(3).Enabled = True Then
                Txt(3).SetFocus
            End If

            Exit Sub
        End If
    End If

    If TxtStartQast.text = "" Then
        If Me.Txt(5).text = "" Then
            Msg = "ÌÃ» ≈œŒ«· ⁄œœ «·√ﬁ”«ÿ"

            If WithMsg = True Then
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Txt(5).SetFocus
            End If

            Exit Sub
        End If

        If Not IsNumeric(Me.Txt(5).text) Then
            Msg = " ⁄œœ «·√ﬁ”«ÿ ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"

            If WithMsg = True Then
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Txt(5).SetFocus
            End If

            Exit Sub
        End If
    End If

    SngAllValue = val(Txt(4).text)

    If Txt(7).text = "" Then
        Msg = "ÌÃ» ≈œŒ«· „œ… › —…«·ﬁ”ÿ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Txt(7).SetFocus
        Exit Sub
    End If

    If Not IsNumeric(Txt(7).text) Then
        Msg = "„œ… › —…«·ﬁ”ÿ ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Txt(7).SetFocus
        Exit Sub
    End If

    IntNoOFQast = val(Txt(5).text)

    If TxtStartQast.text <> "" Then
        If Not IsNumeric(TxtStartQast.text) Then
            Msg = "«·ﬁÌ„… «·„»œ∆Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtStartQast.SetFocus
            Exit Sub
        End If
    End If

    If TxtDiscount.text <> "" Then
        If Not IsNumeric(TxtDiscount.text) Then
            Msg = "ﬁÌ„… «·Œ’„   ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtDiscount.SetFocus
            Exit Sub
        End If
    End If

    If Dtp_First.value = Date Then
        Msg = " «—ÌŒ √Ê· ﬁ”ÿ ÂÊ  «—ÌŒ «·ÌÊ„ " & Chr(13)
        Msg = Msg & "Â· «‰  „ √ﬂœ „‰ «·√” „—«—...øø" & Chr(13)

        If WithMsg = True Then
            IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

            If IntRes = vbNo Then
                Exit Sub
            End If
        End If
    End If

    If IsNumeric(TxtDiscount.text) Then
        SngAllValue = val(Txt(4).text) - val(TxtDiscount.text)
    End If

    If TxtAdvPayment.text <> "" Then
        If Not IsNumeric(TxtAdvPayment.text) Then
            Msg = "ﬁÌ„… «·œ›⁄Â «·„ﬁœ„…   ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtDiscount.SetFocus
            Exit Sub
        End If
    End If

    If IsNumeric(TxtAdvPayment.text) Then
        SngAllValue = SngAllValue - val(TxtAdvPayment.text)
    End If

    If val(Me.TxtStartQast.text) > 0 Then
        IntNoOFQast = SngAllValue \ val(Me.TxtStartQast.text)
        SngOnePor = val(Me.TxtStartQast.text)
    Else
        SngOnePor = SngAllValue / IntNoOFQast
    End If

    If OptInt(0).value = True Then
        DateInterval = "d"
    ElseIf OptInt(1).value = True Then
        DateInterval = "M"
    ElseIf OptInt(2).value = True Then
        DateInterval = "yyyy"
    End If

    NewDate = Dtp_First.value
    DateNumber = val(Txt(7).text)

    'End If
    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + IntNoOFQast

        For i = 1 To IntNoOFQast

            DoEvents
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("Value")) = Round(SngOnePor, Decimal_Places1)

            If i = 1 Then
                NewDate = NewDate
                '        ElseIf I = 2 Then
                '            PreDate = CDate("1" & "/" & Month(Dtp_First.Value) + 1 & "/" & Year(Dtp_First.Value))
                '            NewDate = PreDate
            Else
                PreDate = CDate(Trim(.TextMatrix(i - 1, .ColIndex("Due_Date"))))
                NewDate = DateAdd(DateInterval, DateNumber, PreDate)
            End If

            .TextMatrix(i, .ColIndex("Due_Date")) = Format(NewDate, "yyyy/M/d")
            Due_Date = Format(NewDate, "yyyy/M/d")
        Next i

        .AutoSize 1, .Cols - 1, False
        Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

    If (SngAllValue Mod SngOnePor > 0) And val(Me.TxtStartQast.text) > 0 Then

        With Fg
            .Rows = .Rows + 1
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("Value")) = (SngAllValue Mod SngOnePor)
            PreDate = CDate(Trim(.TextMatrix(i - 1, .ColIndex("Due_Date"))))
            NewDate = DateAdd(DateInterval, DateNumber, PreDate)
            .TextMatrix(i, .ColIndex("Due_Date")) = Format(NewDate, "yyyy/M/d")
            Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
            Due_Date = Format(NewDate, "yyyy/M/d")
            
        End With

    End If

    frmsalebill.DtpDelayDate = Due_Date
    'BolQastCal = True
    Exit Sub
ErrTrap:
End Sub



Private Sub OptInt_Click(Index As Integer)
    Fg.Clear flexClearScrollable, flexClearEverything
    LblTotalQasts.Caption = ""
End Sub

Private Sub Txt_Change(Index As Integer)
    Fg.Clear flexClearScrollable, flexClearEverything
    LblTotalQasts.Caption = ""
    CalPre
End Sub

Private Sub TxtStartQast_Change()
    Fg.Clear flexClearScrollable, flexClearEverything
    LblTotalQasts.Caption = ""
End Sub

Public Sub Retrive(NoteID As Long)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RowNum As Integer
    '«·»Ì«‰«  «·√”«”Ì…
    StrSQL = "SELECT * FROM InstallMent WHERE NoteID=" & NoteID
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        LblID.Caption = IIf(IsNull(RsTemp("PartID").value), "", (RsTemp("PartID").value))
        Txt(1).text = IIf(IsNull(RsTemp("BasicAmmount").value), "", (RsTemp("BasicAmmount").value))

        If RsTemp("InterestType").value <> "" Then
            CboPrecenType.ListIndex = RsTemp("InterestType").value
        Else
            CboPrecenType.ListIndex = 2
        End If

        Txt(3).text = IIf(IsNull(RsTemp("InterestVal").value), "", (RsTemp("InterestVal").value))
        Txt(4).text = IIf(IsNull(RsTemp("Total").value), "", (RsTemp("Total").value))
        Txt(5).text = IIf(IsNull(RsTemp("InstallCount").value), "", (RsTemp("InstallCount").value))
    
        TxtDiscount.text = IIf(IsNull(RsTemp("Discount").value), "", (RsTemp("Discount").value))
        Me.TxtAdvPayment.text = IIf(IsNull(RsTemp("AdvPayment").value), "", (RsTemp("AdvPayment").value))
       
        Me.TxtStartQast.text = IIf(IsNull(RsTemp("StartValue").value), "", RsTemp("StartValue").value)
        '      LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
    
        Dtp_First.value = IIf(IsNull(RsTemp("FirstInstallDate").value), Date, (RsTemp("FirstInstallDate").value))

        If RsTemp("InstallmentType").value <> "" Then

            Select Case RsTemp("InstallmentType").value

                Case 0
                    OptInt(0).value = True

                Case 1
                    OptInt(1).value = True

                Case 2
                    OptInt(2).value = True
            End Select

        End If

        Txt(7).text = IIf(IsNull(RsTemp("InstallSeprator").value), "", (RsTemp("InstallSeprator").value))
        TxtStartQast.text = IIf(IsNull(RsTemp("StartValue").value), "", (RsTemp("StartValue").value))
        TxtDiscount.text = IIf(IsNull(RsTemp("Discount").value), "", (RsTemp("Discount").value))
    End If

    '»Ì«‰«  «·√ﬁ”«ÿ
    If LblID.Caption <> "" Then
        StrSQL = "select * From InstallMentDetails where PartID= " & LblID.Caption
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then

            With Fg
                .Rows = RsTemp.RecordCount + 1

                For RowNum = 1 To RsTemp.RecordCount '
                    .RowData(RowNum) = IIf(IsNull(RsTemp("QestID").value), "", (RsTemp("QestID").value))
                    .TextMatrix(RowNum, .ColIndex("Serial")) = IIf(IsNull(RsTemp("QeqtNum").value), "", (RsTemp("QeqtNum").value))
                    .TextMatrix(RowNum, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", (RsTemp("Value").value))
                    .TextMatrix(RowNum, .ColIndex("Due_Date")) = IIf(IsNull(RsTemp("DueDate").value), "", Format((RsTemp("DueDate").value), "yyyy/mm/dd"))
                    Debug.Print .RowData(RowNum)
                    RsTemp.MoveNext
                Next RowNum

                Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
            End With

        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    Select Case Me.Fg.ColKey(Col)

        Case "Value"
            Cancel = False

        Case "Due_Date"
            Cancel = False

        Case Else
            Cancel = True
    End Select

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With Fg

        Select Case Fg.ColKey(Col)

            Case "Due_Date"

                If .TextMatrix(Row, Col) <> "" Then
                    If IsDate(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), "YYYY/M/D")
                    Else
                        .TextMatrix(Row, Col) = OldGrdValue
                    End If
                End If

            Case "Value"
                CalSum
        End Select

    End With

End Sub

Private Sub CalSum()

    With Fg
        Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

End Sub

