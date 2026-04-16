VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMatrialsOp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… „Ê«œ «·⁄„·ÌÂ"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17220
   Icon            =   "FrmMatrialsOp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   17220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      BackColor       =   &H00C0FFC0&
      Caption         =   "„Ê«œ «·⁄„·Ì… —ﬁ„"
      Height          =   4455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   14055
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3960
         Width           =   1965
      End
      Begin VB.TextBox TxtTotalApro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   5400
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3960
         Width           =   1965
      End
      Begin VB.TextBox XPTxtDiscountVal 
         Height          =   375
         Left            =   11400
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   3960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox XPCboDiscountType 
         Height          =   315
         Left            =   10560
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   360
         Left            =   720
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   -120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox TxtFillData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   -120
         Visible         =   0   'False
         Width           =   690
      End
      Begin C1SizerLibCtl.C1Elastic TxtCodeAother 
         Height          =   690
         Index           =   2
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   13875
         _cx             =   24474
         _cy             =   1217
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
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
         Begin VB.TextBox TxtCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11940
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   300
            Width           =   1860
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4770
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3390
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   705
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   4125
            MaxLength       =   20
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   -300
            Width           =   2370
         End
         Begin VB.TextBox TxtQuantity 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2055
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   300
            Width           =   1185
         End
         Begin VB.ComboBox CboItemCase 
            Height          =   315
            Left            =   6570
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   -420
            Width           =   1785
         End
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   375
            Left            =   75
            TabIndex        =   18
            Top             =   240
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmMatrialsOp.frx":038A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   8880
            TabIndex        =   1
            Top             =   300
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCatalog 
            Height          =   315
            Left            =   6000
            TabIndex        =   36
            Top             =   300
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ ·ÊÃ"
            Height          =   255
            Index           =   2
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   0
            Width           =   2400
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…  ﬁœÌ—Ì"
            Height          =   255
            Index           =   19
            Left            =   4860
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   0
            Width           =   1140
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—  ﬁœÌ—Ì"
            Height          =   255
            Index           =   12
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄— «·›⁄·Ì"
            Height          =   255
            Index           =   26
            Left            =   705
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   0
            Width           =   1305
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì… «·›⁄·Ì…"
            Height          =   255
            Index           =   27
            Left            =   2115
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ì—Ì«·"
            Height          =   255
            Index           =   28
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   -600
            Width           =   2025
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’‰›"
            Height          =   255
            Index           =   29
            Left            =   6735
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   -720
            Width           =   1620
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈”„ «·’‰›"
            Height          =   255
            Index           =   30
            Left            =   9090
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   0
            Width           =   2400
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   11565
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   0
            Width           =   2385
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   2865
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   13875
         _cx             =   24474
         _cy             =   5054
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
         BackColorFixed  =   14871017
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmMatrialsOp.frx":0724
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
         WallPaperAlignment=   0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   8
         Left            =   12360
         TabIndex        =   34
         Top             =   3960
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   476
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–›"
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmMatrialsOp.frx":0A67
         ColorButton     =   12648384
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã„«·Ì «· ﬁœÌ—Ì"
         Height          =   255
         Index           =   1
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   3960
         Width           =   1365
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã„«·Ì «·›⁄·Ì"
         Height          =   255
         Index           =   0
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3960
         Width           =   1365
      End
      Begin VB.Label LblTotalQty 
         Caption         =   "Label38"
         Height          =   135
         Left            =   12120
         TabIndex        =   29
         Top             =   3960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LblItemsCount 
         Caption         =   "Label27"
         Height          =   135
         Left            =   240
         TabIndex        =   28
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   6
      Top             =   4920
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   ""
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin MSComctlLib.TreeView TreeItems 
      Height          =   4455
      Left            =   14040
      TabIndex        =   35
      Top             =   480
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   7858
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "„Ê«œ «·⁄„·ÌÂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   17175
   End
End
Attribute VB_Name = "FrmMatrialsOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim NewGrid As New ClsGrid
Dim currentterms As String
Private Sub AddNewFgAttachRow()
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If val(Me.DCboItemsName.BoundText) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «Ú”„ «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DCboItemsName.SetFocus
        Exit Sub
    End If



    If val(Me.Text16.text) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬂ„ÌÂ «· ﬁœÌ—ÌÂ ··’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Text16.SetFocus
        Exit Sub
    End If



   ' With Me.FgAttachs
   '     LngFindRow = .FindRow(val(Me.DcboItemID1.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)
'
'        If LngFindRow <> -1 Then
'            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·« ...!!!"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
''            .SetFocus
 '           Exit Sub
 '       End If
'
'    End With

    LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("itemid"))

    With Me.FG
        .TextMatrix(LngNewRow, .ColIndex("catalogID")) = Me.DcbCatalog.BoundText
        .TextMatrix(LngNewRow, .ColIndex("catalog")) = DcbCatalog.text
        
        .TextMatrix(LngNewRow, .ColIndex("itemid")) = Me.DCboItemsName.BoundText
        .TextMatrix(LngNewRow, .ColIndex("Code")) = TxtCode.text
        .TextMatrix(LngNewRow, .ColIndex("Name")) = DCboItemsName.text
        .TextMatrix(LngNewRow, .ColIndex("Quntapro")) = val(Text16.text)
        .TextMatrix(LngNewRow, .ColIndex("priceapro")) = val(Text15.text)
        .TextMatrix(LngNewRow, .ColIndex("totalapro")) = val(Text15.text) * val(Text16.text)
        .TextMatrix(LngNewRow, .ColIndex("Count")) = val(TxtQuantity.text)
        .TextMatrix(LngNewRow, .ColIndex("Price")) = val(TxtPrice.text)
        .TextMatrix(LngNewRow, .ColIndex("Valu")) = val(TxtPrice.text) * val(TxtQuantity.text)
        .AutoSize 0, .Cols - 1, False
    End With

 Text16.text = ""
Text15.text = ""
    Me.DCboItemsName.text = ""
    Me.TxtCode.text = ""
   TxtPrice.text = ""
 TxtQuantity.text = ""
    Me.TxtCode.SetFocus
ReLineGrid
End Sub
Sub fillCombo(Optional ItemID As Double = 0, Optional ByRef ItemName As String)
Dim Rs1 As ADODB.Recordset
Dim sql As String
Set Rs1 = New ADODB.Recordset
sql = " select * from TblItems where ItemID =" & ItemID & ""
Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
ItemName = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
Else
ItemName = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
End If
TxtCode.text = IIf(IsNull(Rs1("ItemCode").value), "", Rs1("ItemCode").value)
DCboItemsName.BoundText = val(IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value))
End If
End Sub
Sub GEtCatlog(Optional ItemID As Double = 0, Optional ByRef CatlogName As String)
Dim Rs1 As ADODB.Recordset
Dim sql As String
Set Rs1 = New ADODB.Recordset
sql = " select * from TblItemCatalog where id =" & ItemID & ""
Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then

CatlogName = IIf(IsNull(Rs1("CatlogName").value), "", Rs1("CatlogName").value)

End If
End Sub


Sub FillGrid(Optional ID As Double = 0)
Dim rs As ADODB.Recordset
Dim sql As String
Dim i As Integer
              Set rs = New ADODB.Recordset

sql = " SELECT     TOP 100 PERCENT dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEFDetails.ItemId, dbo.TblProcessDEFDetails.UnitID, dbo.TblProcessDEFDetails.Price,"
sql = sql & "                      dbo.TblProcessDEFDetails.cost , dbo.TblItems.itemcode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode"
sql = sql & "  FROM         dbo.TblProcessDEF INNER JOIN"
sql = sql & "                       dbo.TblProcessDEFDetails ON dbo.TblProcessDEF.TblProcessDEFID = dbo.TblProcessDEFDetails.TblProcessDEFID INNER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.TblProcessDEFDetails.ItemId = dbo.TblItems.ItemID"
sql = sql & "  Where (dbo.TblProcessDEF.TblProcessDEFID = " & ID & ")"
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
With FG
.rows = .rows + rs.RecordCount
rs.MoveFirst
For i = 1 To rs.RecordCount
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
.TextMatrix(i, .ColIndex("itemid")) = IIf(IsNull(rs("ItemId").value), "", rs("ItemId").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)

Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)

End If
.TextMatrix(i, .ColIndex("Quntapro")) = IIf(IsNull(rs("Cost").value), "", rs("Cost").value)
.TextMatrix(i, .ColIndex("priceapro")) = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
.TextMatrix(i, .ColIndex("totalapro")) = IIf(IsNull(rs("Cost").value), 0, rs("Cost").value) * IIf(IsNull(rs("Price").value), 0, rs("Price").value)
rs.MoveNext
Next i
End With

End If
ReLineGrid


End Sub

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
    save
    Unload Me

       Case 24
     '  AddNewFgRowother
       Case 8
            DeleteFgRowAther
    End Select

End Sub
Sub save()
Dim str As Variant
Dim i As Integer
str = ""
Dim xx As Variant


With Me.FG
For i = 1 To .rows - 1
 If .TextMatrix(i, .ColIndex("Name")) <> "" Then
 str = str & Trim(.TextMatrix(i, .ColIndex("itemid"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Count"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Price"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Valu"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Quntapro"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("priceapro"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("catalogID"))) & "#"
  str = str & Trim(.TextMatrix(i, .ColIndex("monthly"))) & "#"
 str = str & Trim("@")
  str = str & CHR(13)
  str = Trim(str)
 End If
Next

'xx = StrConv(str, vbUnicode)



Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("FlgOper")) = 1
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("matrials")) = str
'Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("matrials")) = val(TxtTotalApro.Text)
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("total_items")) = val(TxtTotalApro.text)
End With
End Sub

Sub Retrive(Optional Pand As Double = 0, Optional Opr As Double = 0, Optional ProjectID As Double = 0)
  Dim astrSplit2tems2() As String
  Dim astrSplitItems() As String
  Dim ItemName As String
  Dim i As Integer
  Dim j As Integer
  Dim st As String
  Dim nElements As Integer
  Dim Catalogname As String
 
          
      With Me.FG
    If Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("matrials")) <> "" Then
       st = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("matrials"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
      nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
        .rows = .FixedRows + nElements

        For j = 0 To nElements - 1
    astrSplit2tems2 = Split(astrSplitItems(j), "#")
    i = j + 1
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("itemid")) = val(astrSplit2tems2(0))
             .TextMatrix(i, .ColIndex("Code")) = GetItemCode(val(astrSplit2tems2(0)))
             fillCombo val(astrSplit2tems2(0)), ItemName
            .TextMatrix(i, .ColIndex("Name")) = ItemName
            .TextMatrix(i, .ColIndex("Quntapro")) = val(astrSplit2tems2(4))
            .TextMatrix(i, .ColIndex("priceapro")) = val(astrSplit2tems2(5))
            .TextMatrix(i, .ColIndex("Count")) = val(astrSplit2tems2(1))
            .TextMatrix(i, .ColIndex("Price")) = val(astrSplit2tems2(2))
            GEtCatlog val(astrSplit2tems2(6)), Catalogname
            .TextMatrix(i, .ColIndex("catalog")) = Catalogname
            .TextMatrix(i, .ColIndex("catalogID")) = val(astrSplit2tems2(6))
            .TextMatrix(i, .ColIndex("monthly")) = val(astrSplit2tems2(7))
            .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Count")))
            .TextMatrix(i, .ColIndex("totalapro")) = val(.TextMatrix(i, .ColIndex("priceapro"))) * val(.TextMatrix(i, .ColIndex("Quntapro")))
             '  End If
         '   RsDetails.MoveNext
         
        Next j
        End If
    'ReLineGridCount
    ReLineGrid
    
End With

End Sub
Private Sub DeleteFgRowAther()

    With Me.FG

        If .row = -1 Then Exit Sub
        If .row = 0 Then Exit Sub
        .RemoveItem .row
        '.AutoSize 0, .Cols - 1, False
     ReLineGrid
    End With

End Sub

Private Sub cmdAdd_Click()
AddNewFgAttachRow
End Sub

Private Sub DCboItemsName_Change()
Dim Dcombos As ClsDataCombos
 Set Dcombos = New ClsDataCombos
 Dim UnitID As Long
     Dim StrSQL  As String
    Me.TxtCode.text = GetItemCode(val(Me.DCboItemsName.BoundText))
     Dim RsUnitData As New ADODB.Recordset
            StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
            StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
            StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & val(Me.DCboItemsName.BoundText)
            StrSQL = StrSQL + " AND DefaultUnit=1"
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
       
                    UnitID = RsUnitData("UnitID").value
                     
       Else
       UnitID = 1
        
            End If

            RsUnitData.Close
            Set RsUnitData = Nothing
           
           
    Text15.text = GetItemPrice(val(Me.DCboItemsName.BoundText), 1, UnitID)
    TxtPrice.text = Text15.text
    
Dcombos.GetCatalogItem Me.DcbCatalog, val(Me.DCboItemsName.BoundText)
End Sub

Private Sub DCboItemsName_Click(Area As Integer)
DCboItemsName_Change
End Sub

Private Sub FG_AfterEdit(ByVal row As Long, ByVal Col As Long)
 Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG

        Select Case .ColKey(Col)
        Case "monthly"
        .TextMatrix(row, .ColIndex("Quntapro")) = Projects.Monthly * val(.TextMatrix(row, .ColIndex("monthly")))
         .TextMatrix(row, .ColIndex("totalapro")) = val(.TextMatrix(row, .ColIndex("Quntapro"))) * val(.TextMatrix(row, .ColIndex("priceapro")))
            Case "catalog"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("catalogID"), False, True)
                .TextMatrix(row, .ColIndex("catalogID")) = StrAccountCode
         End Select
       End With
ReLineGrid
End Sub

Private Sub fg_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim ItemID As Long

    With FG

        Select Case .ColKey(Col)
        

            Case "catalog"
            If .TextMatrix(row, .ColIndex("itemid")) <> "" Then
            ItemID = val(.TextMatrix(row, .ColIndex("itemid")))
                StrSQL = "select * from TblItemCatalog where ItemID=" & ItemID & ""
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = FG.BuildComboList(rs, "CatlogName", "ID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
           
               End If
        End Select

    End With
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub TreeItems_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim NodeKey As Double
    On Error GoTo ErrTrap

   

    If right(Node.key, 1) = "G" Then
        Exit Sub
    End If

    NodeKey = left(Node.key, Len(Node.key) - 1)

    If NodeKey <> 0 Then
    fillCombo NodeKey
        'Retrive (NodeKey)
                
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
   txtTotal.text = 0
   TxtTotalApro.text = 0
    
    IntCounter = 0
    Dim i As Integer

    With Me.FG

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
               .TextMatrix(i, .ColIndex("FullCode")) = currentterms & "-" & .TextMatrix(i, .ColIndex("Ser"))
                .TextMatrix(i, .ColIndex("totalapro")) = val(.TextMatrix(i, .ColIndex("priceapro"))) * val(.TextMatrix(i, .ColIndex("Quntapro")))
                .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Count")))
  txtTotal.text = val(txtTotal.text) + val(.TextMatrix(i, .ColIndex("Valu")))
  TxtTotalApro.text = val(TxtTotalApro.text) + val(.TextMatrix(i, .ColIndex("totalapro")))
  
            End If

        Next i
   
    End With
    End Sub
    
Private Sub Form_Load()
Dim rs As ADODB.Recordset
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim Xpid As Double
Dim rw As Double
Dim rwOp As Double
Dim rwpand As Double
    Set Dcombos = New ClsDataCombos
Dcombos.GetItemsNames Me.DCboItemsName
Dcombos.GetCatalogItem Me.DcbCatalog

'//
 TreeItems.ImageList = mdifrmmain.ImgLstTree
   Set rs = New ADODB.Recordset
    rs.Open "[TblItems]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   '  LoadMenus
    LoadTreeGroups TreeItems
''//''
   ' Set NewGrid.DtpBillDate = Me.XPDtbBill
If Projects.TxtModFlg.text <> "R" Then
Cmd(0).Enabled = True
Else
Cmd(0).Enabled = False

End If

    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    currentterms = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("fullcode"))
    If SystemOptions.UserInterface = ArabicInterface Then
                    Frame1.Caption = "„Ê«œ «·⁄„·ÌÂ —ﬁ„ —ﬁ„ : " & currentterms
                Else
                    Frame1.Caption = "Matrials For Process No: " & currentterms
                End If
      
             
    Xpid = val(Projects.txt_project_id.text)
    rwOp = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("id")))
    rw = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("OPRIDD")))
    rwpand = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("ProjectDes_ID")))
  
If Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("matrials")) <> "" Then
Retrive rwpand, rwOp, Xpid
 End If
If Projects.TxtModFlg.text = "N" And val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("FlgOper"))) <> 1 And Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("matrials")) = "" Then
If (Xpid <> 0 And rw) <> 0 Then
Me.FillGrid rw
End If
End If
    Set GrdBack = New ClsBackGroundPic

 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
'

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Save"
    Cmd(2).Caption = "Exit"
    Cmd(8).Caption = "Delete"
  Me.Caption = "Distribution Expenses on Items"
  
Label5.Caption = Me.Caption
lbl(31).Caption = "Code"
lbl(30).Caption = "name"
lbl(2).Caption = "Catalog"
lbl(19).Caption = "Exp. Qty"
lbl(12).Caption = "Exp Price"
'Me.LblClientName.Caption = "ClientName"

lbl(27).Caption = "Act.. Qty"
lbl(26).Caption = "Act Price"

lbl(1).Caption = "Exp Total"
lbl(0).Caption = "Act Total"

'lbl(4).Caption = "From"
'lbl(3).Caption = "To"
'Cmd(24).Caption = "Add"
'Cmd(25).Caption = "Delete"

 
'Lbl(51).Caption = "Type Value"
'Lbl(41).Caption = "Value  "
 
'Lbl(39).Caption = "Count"
'Me.lbreg.Caption = "Date Registration"

     With Me.FG
     
 
        .TextMatrix(0, .ColIndex("fullcode")) = "Opr Code"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("Quntapro")) = "Exp Qty"
        .TextMatrix(0, .ColIndex("priceapro")) = "Exp Price"
        .TextMatrix(0, .ColIndex("totalapro")) = "Exp Totals"
        .TextMatrix(0, .ColIndex("count")) = "Act. Qty"
        .TextMatrix(0, .ColIndex("Price")) = "Act. Price"
        .TextMatrix(0, .ColIndex("Valu")) = "Act .Totals"
        .TextMatrix(0, .ColIndex("Catalog")) = "Catalog"
        .TextMatrix(0, .ColIndex("monthly")) = "Monthly Qty"
 
    End With
  '
End Sub


Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtCode.text = "" Then
            Me.DCboItemsName.BoundText = ""
        Else
            Me.DCboItemsName.BoundText = GetItemID(Trim$(Me.TxtCode.text))
        End If
    End If
End Sub
