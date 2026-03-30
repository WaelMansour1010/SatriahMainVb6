VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmFillItems 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "…"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4515
   HelpContextID   =   1000
   Icon            =   "FrmFillItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox XPTxtCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2280
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   15
      Width           =   1215
   End
   Begin VB.ComboBox XPCboItemCase 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1185
   End
   Begin VB.TextBox XPTxtPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2280
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1425
      Width           =   1215
   End
   Begin VB.TextBox XPTxtQuantity 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2280
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1065
      Width           =   1215
   End
   Begin VB.TextBox XPTxtDiscountValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   2280
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2430
      Width           =   1215
   End
   Begin VB.ComboBox XPCboDiscountType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2085
      Width           =   1185
   End
   Begin VB.TextBox TxtGuaranteeTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2280
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1770
      Width           =   1215
   End
   Begin VB.TextBox XPTxtSerial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2280
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2790
      Width           =   1215
   End
   Begin VSFlex8UCtl.VSFlexGrid FgSerials 
      Height          =   2055
      Left            =   0
      TabIndex        =   27
      Top             =   1050
      Width           =   2235
      _cx             =   3942
      _cy             =   3625
      Appearance      =   2
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmFillItems.frx":038A
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
      Editable        =   0
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
   Begin VB.TextBox TxtJeneralPart 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   1380
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3450
      Width           =   1395
   End
   Begin VB.CheckBox XPChkQuantity 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "  ”ÃÌ· ﬂ„Ì…"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3165
      Width           =   1515
   End
   Begin MSDataListLib.DataCombo DCboItemsName 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Left            =   30
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   330
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmFillItems.frx":03DA
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton CmdAdd 
      Height          =   375
      Left            =   975
      TabIndex        =   10
      Top             =   3930
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈÷«›…"
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
      ButtonImage     =   "FrmFillItems.frx":0974
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
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3930
      Width           =   795
      _ExtentX        =   1402
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
      ButtonImage     =   "FrmFillItems.frx":0D0E
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
   Begin ImpulseButton.ISButton XPBtnRemove 
      Height          =   285
      Left            =   450
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3150
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   503
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
      ButtonImage     =   "FrmFillItems.frx":10A8
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      LowerToggledContent=   0   'False
   End
   Begin ImpulseButton.ISButton CmdClearAll 
      Height          =   285
      Left            =   30
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3150
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   503
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
      ButtonImage     =   "FrmFillItems.frx":1442
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      LowerToggledContent=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄œœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   10
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3150
      Width           =   405
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   9
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3150
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5205
      X2              =   -30
      Y1              =   3840
      Y2              =   3855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì„ﬂ‰ﬂ ≈ŸÂ«— ‘«‘… «·»ÕÀ ⁄‰ «·√’‰«› »«·÷€ÿ ⁄·Ï F7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   8
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3885
      Width           =   2655
   End
   Begin VB.Label LblRequireSerial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2790
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ã“¡ «·À«»  „‰ «·”Ì—Ì«·"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3465
      Width           =   1665
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”Ì—Ì«·"
      Height          =   315
      Index           =   4
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2805
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„œ… «·÷„«‰"
      Height          =   315
      Index           =   0
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1770
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·Œ’„"
      Height          =   315
      Index           =   2
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2085
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·Œ’„"
      Height          =   315
      Index           =   7
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”⁄—"
      Height          =   315
      Index           =   1
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1425
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬂ„Ì…"
      Height          =   315
      Index           =   3
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·’‰›"
      Height          =   315
      Index           =   4
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   735
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   315
      Index           =   5
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   45
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰›"
      Height          =   315
      Index           =   6
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   390
      Width           =   945
   End
   Begin VB.Label LblSerial 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì” ·Â ”Ì—Ì«·"
      Height          =   225
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   765
      Width           =   1155
   End
End
Attribute VB_Name = "FrmFillItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo  As clsDCboSearch

Private m_DealingForm As GridTransType
Dim TTP As clstooltip

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "≈÷«›… ’‰› ›Ï «·›« Ê—…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnRemove, "Õ–› «·”Ì—Ì«·  «·„Õœœ „‰ ﬁ«∆„…" & Wrap & "«·”Ì—Ì«·«  «·„÷«›…", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "≈÷«›… ’‰› ›Ï «·›« Ê—…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdClearAll, "„”Õ ﬂ· «·”Ì—Ì«·«  «·„÷«›…", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "≈÷«›… ’‰› ›Ï «·›« Ê—…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl FgSerials, "√—ﬁ«„ «·”Ì—Ì«·  «·Œ«’… »«·’‰› «·„÷«›", BolRtl
        End With

    End If

ErrTrap:
End Sub

Private Sub cmdAdd_Click()

    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer
    Dim IntRes As Integer

    On Error GoTo ErrTrap

    If DCboItemsName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If XPTxtQuantity.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtQuantity.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(XPTxtQuantity.text) Then
        Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtQuantity.SetFocus
        Exit Sub
    End If

    If XPTxtPrice.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·”⁄—"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtPrice.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(XPTxtPrice.text) Then
        Msg = "«·”⁄— ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtPrice.SetFocus
        Exit Sub
    End If

    If Me.DealingForm <> ShowPrice Then
        If TxtGuaranteeTime.text <> "" Then
            If Me.TxtGuaranteeTime.Enabled = True Then
                If Not (IsNumeric(TxtGuaranteeTime.text)) Or val(TxtGuaranteeTime.text) < 1 Or val(TxtGuaranteeTime.text) > 36 Or val(TxtGuaranteeTime.text) <> Int(val(TxtGuaranteeTime.text)) Then
                    Msg = "„œ… «·÷„«‰ €Ì— ’«·Õ…" & Chr(13)
                    Msg = Msg + "√œŒ· ﬁÌ„… —ﬁ„Ì… ’ÕÌÕ… »Ì‰ 1 -36 " & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    TxtGuaranteeTime.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountValue.text = "" Then
            Msg = "ÌÃ»  ÕœÌœ ﬁÌ„… «·Œ’„"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountValue.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtDiscountValue.text) Then
            Msg = "ﬁÌ„… «·Œ’„ ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountValue.SetFocus
            Exit Sub
        End If
    End If

    Select Case Me.DealingForm

        Case InvoiceTransaction

            With mdifrmmain.ActiveForm.FG

                If XPChkQuantity.value = Checked Then
                    If Trim$(TxtJeneralPart.text) = "" Then
                        Msg = "·„  ﬁ„ »≈œŒ«· «·Ã“¡ «·À«»  „‰ «·”Ì—Ì«· ...!!!"
                        Msg = Msg & Chr(13) & "Â·  —Ìœ «·„ «»⁄… ..øøø"
                        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                        If IntRes = vbNo Then
                            TxtJeneralPart.SetFocus
                            Exit Sub
                        End If
                    End If

                    mdifrmmain.ActiveForm.TxtFillData.text = "T"
                    Screen.MousePointer = vbArrowHourglass

                    For ItemCount = 1 To XPTxtQuantity

                        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                            .Rows = .Rows + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("HaveSerial")) = True

                        If XPCboItemCase.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                        End If

                        VarNum = XPTxtSerial.text + ItemCount - 1
                        StrSerial = TxtJeneralPart & VarNum
                        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = StrSerial
                        .TextMatrix(.Rows - 1, .ColIndex("Count")) = 1
                        .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                        If XPCboDiscountType.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
                        .TextMatrix(.Rows - 1, .ColIndex("guaranteeTime")) = IIf(TxtGuaranteeTime.text = "", "", TxtGuaranteeTime.text)
                    Next ItemCount

                    Screen.MousePointer = vbDefault
                    mdifrmmain.ActiveForm.TxtFillData.text = "F"
                Else

                    If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                        .Rows = .Rows + 1
                    End If
        
                    .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                    .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

                    If XPCboItemCase.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("Serial")) = XPTxtSerial.text
                    .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
                    .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                    If XPCboDiscountType.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
                    .TextMatrix(.Rows - 1, .ColIndex("guaranteeTime")) = IIf(TxtGuaranteeTime.text = "", "", TxtGuaranteeTime.text)

                    If LblSerial.Tag = "T" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
                    ElseIf LblSerial.Tag = "F" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
                    End If
                End If

                .AutoSize 0, .Cols - 1, False
            End With

        Case MoveItems

            With FrmMoving.FG

                If XPChkQuantity.value = Checked Then
                    If TxtJeneralPart.text = "" Then
                        Msg = "ÌÃ»  ÕœÌœ «·Ã“¡ «·À«»  „‰ «·”Ì—Ì«· "
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        TxtJeneralPart.SetFocus
                        Exit Sub
                    End If

                    FrmMoving.TxtFillData.text = "T"
                    Screen.MousePointer = vbArrowHourglass

                    For ItemCount = 1 To XPTxtQuantity

                        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                            .Rows = .Rows + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("HaveSerial")) = True

                        If XPCboItemCase.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                        End If

                        VarNum = XPTxtSerial.text + ItemCount - 1
                        StrSerial = TxtJeneralPart & VarNum
                        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = StrSerial
                        .TextMatrix(.Rows - 1, .ColIndex("Count")) = 1
                        .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                        If XPCboDiscountType.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
                        .TextMatrix(.Rows - 1, .ColIndex("guaranteeTime")) = IIf(TxtGuaranteeTime.text = "", "", TxtGuaranteeTime.text)
                    Next ItemCount

                    Screen.MousePointer = vbDefault
                    FrmMoving.TxtFillData.text = "F"
                Else

                    If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                        .Rows = .Rows + 1
                    End If
        
                    .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                    .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

                    If XPCboItemCase.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("Serial")) = XPTxtSerial.text
                    .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
                    .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                    If XPCboDiscountType.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
                    .TextMatrix(.Rows - 1, .ColIndex("guaranteeTime")) = IIf(TxtGuaranteeTime.text = "", "", TxtGuaranteeTime.text)

                    If LblSerial.Tag = "T" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
                    ElseIf LblSerial.Tag = "F" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
                    End If
                End If

                .AutoSize 0, .Cols - 1, False
            End With

        Case PurchaseTransaction

            With mdifrmmain.ActiveForm.FG

                If XPChkQuantity.value = Checked Then
                    If Trim$(TxtJeneralPart.text) = "" Then
                        Msg = "·„  ﬁ„ »≈œŒ«· «·Ã“¡ «·À«»  „‰ «·”Ì—Ì«· ...!!!"
                        Msg = Msg & Chr(13) & "Â·  —Ìœ «·„ «»⁄… ..øøø"
                        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                        If IntRes = vbNo Then
                            TxtJeneralPart.SetFocus
                            Exit Sub
                        End If
                    End If

                    mdifrmmain.ActiveForm.TxtFillData.text = "T"
                    Screen.MousePointer = vbArrowHourglass

                    For ItemCount = 1 To XPTxtQuantity.text

                        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                            .Rows = .Rows + 1
                        End If
        
                        .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("HaveSerial")) = True

                        If XPCboItemCase.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                        End If

                        VarNum = XPTxtSerial.text + ItemCount - 1
                        StrSerial = TxtJeneralPart & VarNum
                        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = StrSerial
                        .TextMatrix(.Rows - 1, .ColIndex("Count")) = 1
                        .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                        If XPCboDiscountType.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)

                        If ItemCount = XPTxtQuantity.text Then
                            mdifrmmain.ActiveForm.TxtFillData.text = "F"
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("guaranteeTime")) = IIf(TxtGuaranteeTime.text = "", "", TxtGuaranteeTime.text)
                    Next ItemCount

                    Screen.MousePointer = vbDefault
                Else

                    If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                        .Rows = .Rows + 1
                    End If
    
                    .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                    .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

                    If XPCboItemCase.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("Serial")) = XPTxtSerial.text
                    .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
                    .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                    If XPCboDiscountType.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
                    .TextMatrix(.Rows - 1, .ColIndex("guaranteeTime")) = IIf(TxtGuaranteeTime.text = "", "", TxtGuaranteeTime.text)

                    If LblSerial.Tag = "T" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
                    ElseIf LblSerial.Tag = "F" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
                    End If
                End If

                .AutoSize 0, .Cols - 1, False
            End With

        Case ReturnTransaction

            With FrmReturnpurchases.FG

                If XPChkQuantity.value = Checked Then
                    If TxtJeneralPart.text = "" Then
                        Msg = "ÌÃ»  ÕœÌœ «·Ã“¡ «·À«»  „‰ «·”Ì—Ì«· "
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        TxtJeneralPart.SetFocus
                        Exit Sub
                    End If

                    FrmReturnpurchases.TxtFillData.text = "T"
                    Screen.MousePointer = vbArrowHourglass

                    For ItemCount = 1 To XPTxtQuantity.text

                        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                            .Rows = .Rows + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("HaveSerial")) = True

                        If XPCboItemCase.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                        End If

                        VarNum = XPTxtSerial.text + ItemCount - 1
                        StrSerial = TxtJeneralPart & VarNum
                        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = StrSerial
                        .TextMatrix(.Rows - 1, .ColIndex("Count")) = 1
                        .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                        If ItemCount = XPTxtQuantity.text Then
                            FrmReturnpurchases.TxtFillData.text = "F"
                        End If

                        Debug.Print ItemCount
                    Next ItemCount

                    Screen.MousePointer = vbDefault
                Else

                    If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                        .Rows = .Rows + 1
                    End If
    
                    .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                    .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

                    If XPCboItemCase.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("Serial")) = XPTxtSerial.text
                    .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
                    .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                    If LblSerial.Tag = "T" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
                    ElseIf LblSerial.Tag = "F" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
                    End If
                End If

                .AutoSize 0, .Cols - 1, False
            End With

            '⁄—Ê÷ «·√”⁄«—
        Case ShowPrice

            With FrmShowPrice.FG

                If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                    .Rows = .Rows + 1
                End If

                .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

                If XPCboItemCase.ListIndex <> -1 Then
                    .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                End If

                .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
                .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                If XPCboDiscountType.ListIndex <> -1 Then
                    .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
                End If

                .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
                .AutoSize 0, .Cols - 1, False
            End With

            '«·⁄—Ê÷ «·Ã«Â“…
        Case Template

           ' With FrmTemplate.FG

           '     If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
           '         .Rows = .Rows + 1
           '     End If
'
'                .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
'                .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
'
'                If XPCboItemCase.ListIndex <> -1 Then
'                    .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
'                End If
'
'                .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
'                .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text
'
'                If XPCboDiscountType.ListIndex <> -1 Then
'                    .TextMatrix(.Rows - 1, .ColIndex("DiscountType")) = XPCboDiscountType.ListIndex + 1
'                End If

'                .TextMatrix(.Rows - 1, .ColIndex("DiscountVal")) = val(XPTxtDiscountValue.text)
'                .AutoSize 0, .Cols - 1, False
'            End With

            '„— Ã⁄ «·„»Ì⁄« 
        Case ReturnSalling

            With FrmReturnSalling.FG

                If XPChkQuantity.value = Checked Then
                    If TxtJeneralPart.text = "" Then
                        Msg = "ÌÃ»  ÕœÌœ «·Ã“¡ «·À«»  „‰ «·”Ì—Ì«· "
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        TxtJeneralPart.SetFocus
                        Exit Sub
                    End If

                    FrmReturnSalling.TxtFillData.text = "T"
                    Screen.MousePointer = vbArrowHourglass

                    For ItemCount = 1 To XPTxtQuantity.text

                        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                            .Rows = .Rows + 1
                        End If

                        .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
                        .TextMatrix(.Rows - 1, .ColIndex("HaveSerial")) = True

                        If XPCboItemCase.ListIndex <> -1 Then
                            .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                        End If

                        VarNum = XPTxtSerial.text + ItemCount - 1
                        StrSerial = TxtJeneralPart & VarNum
                        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = StrSerial
                        .TextMatrix(.Rows - 1, .ColIndex("Count")) = 1
                        .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                        If ItemCount = XPTxtQuantity.text Then
                            FrmReturnSalling.TxtFillData.text = "F"
                        End If

                        Debug.Print ItemCount
                    Next ItemCount

                    Screen.MousePointer = vbDefault
                Else

                    If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                        .Rows = .Rows + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
                    .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

                    If XPCboItemCase.ListIndex <> -1 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("Serial")) = XPTxtSerial.text
                    .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
                    .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

                    If LblSerial.Tag = "T" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
                    ElseIf LblSerial.Tag = "F" Then
                        .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
                    End If
                End If

                .AutoSize 0, .Cols - 1, False
            End With

    End Select

    clear_all Me
    XPTxtCode.SetFocus
    XPCboDiscountType.ListIndex = 0
    XPCboItemCase.ListIndex = 0
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdItemSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DCboItemsName
    FrmItemSearch.show vbModal
End Sub

'
Private Sub DCboItemsName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DCboItemsName.BoundText <> "" Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "select * From TblItems where ItemID=" & DCboItemsName.BoundText
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        XPTxtSerial.Enabled = RsTemp("HaveSerial").value
        XPTxtQuantity.Enabled = Not (RsTemp("HaveSerial").value)
        LblRequireSerial.Visible = RsTemp("HaveSerial").value
    
        XPTxtCode.text = RsTemp("ItemCode").value
        XPLbl(1).Enabled = False
    
        If XPTxtSerial.Enabled = False Then
            XPTxtSerial.text = ""
        End If

        If XPTxtQuantity.Enabled = False Then
            XPTxtQuantity.text = 1
        End If

        If RsTemp("HaveGuarantee").value = True Then
            TxtGuaranteeTime.Enabled = True
        Else
            TxtGuaranteeTime.Enabled = False
        End If

        If RsTemp("HaveSerial").value = True Then
            XPChkQuantity.Enabled = True
            XPChkQuantity.value = False
            XPLbl(4).Caption = "«·”Ì—Ì«·"
            LblSerial.ForeColor = vbRed
            LblSerial.Caption = "·Â ”Ì—Ì«·"
            LblSerial.Tag = "T"
        
        Else
            XPChkQuantity.Enabled = False
            XPChkQuantity.value = Unchecked
            LblSerial.ForeColor = vbBlue
            LblSerial.Caption = "·Ì” ·Â ”Ì—Ì«·"
            LblSerial.Tag = "F"
            TxtGuaranteeTime.text = ""
      
        End If

        Me.TxtGuaranteeTime.text = IIf(IsNull(RsTemp("GuaranteeValue").value), "", RsTemp("GuaranteeValue").value)
        RsTemp.Close
        PutPrice
        XPChkQuantity_Click
    Else
        XPChkQuantity.Enabled = False
        XPChkQuantity.value = Unchecked
        XPLbl(1).Enabled = False
        TxtJeneralPart.Enabled = False
        XPChkQuantity_Click
        LblSerial.Caption = ""
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()

    Select Case Me.DealingForm

        Case InvoiceTransaction
            Me.HelpContextID = 160

        Case PurchaseTransaction
            Me.HelpContextID = 100

        Case ReturnTransaction
            Me.HelpContextID = 240
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CmdItemSearch_Click
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrTrap

    If KeyAscii = vbKeyReturn Then
        If Not Me.ActiveControl Is XPTxtSerial Then
            KeyAscii = 0
            SendKeys "{TAB}"
        Else

            If val(Me.lbl(9).Caption) < val(Me.XPTxtQuantity.text) Then
                'KeyAscii = 0
                XPTxtSerial.SetFocus
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Dim GrdBack  As ClsBackGroundPic
    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic
    Set Me.FgSerials.WallPaper = GrdBack.Picture
    FgSerials.AutoSize 0, FgSerials.Cols - 1, False
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DCboItemsName
    LblSerial.Caption = ""
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboItemsName

    With XPCboDiscountType
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
        .AddItem "„Ã«‰Ì"
    End With

    With XPCboItemCase
        .AddItem "ÃœÌœ"
        .AddItem "„” ⁄„·"
    End With

    If SystemOptions.UserInvoiceChangePrice = 0 Then
        Me.XPTxtPrice.Enabled = False
    End If

    XPCboDiscountType.ListIndex = 0
    XPCboItemCase.ListIndex = 0
    CmdItemSearch.TabStop = False
    AddTip
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Resize()
    'Debug.Print Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
    Set TTP = Nothing
End Sub

Private Sub TxtGuaranteeTime_GotFocus()
    TxtGuaranteeTime.SelStart = 0
    TxtGuaranteeTime.SelLength = Len(TxtGuaranteeTime.text) + 1
End Sub

Private Sub XPCboDiscountType_Change()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = -1 Or XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Then
        XPTxtDiscountValue.Enabled = False
        XPTxtDiscountValue.text = ""
    Else
        XPTxtDiscountValue.Enabled = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboDiscountType_Click()
    XPCboDiscountType_Change
End Sub

Private Sub XPChkQuantity_Click()
    On Error GoTo ErrTrap

    If XPChkQuantity.value = Checked Then
        XPLbl(4).Caption = "»œ«Ì… «· —ﬁÌ„"
        XPTxtQuantity.Enabled = True
        XPTxtQuantity.text = ""
        XPLbl(1).Enabled = True
        TxtJeneralPart.Enabled = True
        TxtJeneralPart.Enabled = True
        XPLbl(1).Enabled = True
        TxtJeneralPart.text = ""
    Else
        XPLbl(4).Caption = "«·”Ì—Ì«·"
        '    XPTxtQuantity.Enabled = False
        '    XPTxtQuantity.Text = ""
        XPLbl(1).Enabled = False
        TxtJeneralPart.Enabled = False
        TxtJeneralPart.Enabled = False
        XPLbl(1).Enabled = False
        TxtJeneralPart.text = ""
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtCode_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset

    If XPTxtCode.text <> "" Then
        StrSQL = "select * From TblItems where ItemCode='" & XPTxtCode.text & "'"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            DCboItemsName.BoundText = RsTemp("ItemID").value
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)
    'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
    'End If
End Property

Private Sub XPTxtPrice_GotFocus()
    On Error Resume Next
    XPTxtPrice.SelStart = 0
    XPTxtPrice.SelLength = Len(XPTxtPrice.text) + 1

End Sub

Private Sub XPTxtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtPrice.text, 0)
End Sub

Private Sub XPTxtQuantity_GotFocus()
    SelectText XPTxtQuantity
End Sub

Private Sub XPTxtQuantity_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtQuantity.text, 0)
End Sub

Private Sub AddFgRow(StrRowText As String)
    Dim LngRow As Long
    Dim i As Integer
    Dim IntSerialCount As Integer

    With Me.FgSerials
        .Rows = .Rows + 1
        LngRow = .Rows - 1
        .TextMatrix(LngRow, .ColIndex("ItemSerial")) = StrRowText
        ReSerialGrid FgSerials, .ColIndex("Counter")
        .AutoSize 0, .Cols - 1, False
        .ShowCell LngRow, .ColIndex("ItemSerial")
        IntSerialCount = .Rows - 1
        Me.lbl(9).Caption = IntSerialCount
        Me.XPTxtQuantity.text = IntSerialCount
    End With

End Sub

Private Sub XPTxtQuantity_LostFocus()

    If val(Me.XPTxtQuantity.text) > 0 Then
        PutPrice
    Else
        Me.XPTxtPrice.text = "0"
    End If

End Sub

Private Sub XPTxtSerial_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(Me.XPTxtSerial.text) <> "" Then
            KeyAscii = 0
            cmdAdd_Click
            '        AddFgRow Trim(Me.XPTxtSerial.Text)
            '        Me.XPTxtSerial.Text = ""
            '        If Val(Me.lbl(9).Caption) < Val(Me.XPTxtQuantity.Text) Then
            '            XPTxtSerial.SetFocus
            '        End If
        End If
    End If

End Sub

Private Sub PutPrice()
    Dim m_DealFrom As GridTransType
    m_DealFrom = Me.DealingForm

    If m_DealFrom = InvoiceTransaction Or m_DealFrom = ReturnSalling Or m_DealFrom = ShowPrice Then
        Me.XPTxtPrice.text = GetItemPrice(val(Me.DCboItemsName.BoundText), val(Me.XPTxtQuantity.text))
    ElseIf m_DealFrom = PurchaseTransaction Or m_DealFrom = MoveItems Or m_DealFrom = ReturnTransaction Then
        Me.XPTxtPrice.text = GetCostItemPrice(Me.DCboItemsName.BoundText, 2)
    End If

End Sub
