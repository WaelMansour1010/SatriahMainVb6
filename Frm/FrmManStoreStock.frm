VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManStoreStock 
   Caption         =   "Ã—œ „Œ“‰ «·’Ì«‰…"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "FrmManStoreStock.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   10575
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _cx             =   18653
      _cy             =   13176
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   1
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
      GridRows        =   6
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmManStoreStock.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1125
         Index           =   3
         Left            =   2715
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   30
         Width           =   3105
         _cx             =   5477
         _cy             =   1984
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ForeColor       =   128
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "ŒÌ«—«  ⁄„·Ì… «·Ã—œ"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
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
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ· «·√’‰«› «·„ÊÃÊœ… ›Ï «·„Œ“‰"
            Height          =   255
            Index           =   2
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   780
            Width           =   2775
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√’‰«› —Ã⁄  „‰ «·÷„«‰ Ê·„  ”·„"
            Height          =   255
            Index           =   1
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   540
            Width           =   2775
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√’‰«› œŒ·  ’Ì«‰… Ê·„ Ì „  ’Ì·ÕÂ«"
            Height          =   345
            Index           =   0
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   2925
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1125
         Index           =   2
         Left            =   5835
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   4710
         _cx             =   8308
         _cy             =   1984
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
         Caption         =   "⁄Ê«„· «·»ÕÀ"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
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
         Begin VB.ComboBox CboReportType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   270
            Width           =   3495
         End
         Begin MSDataListLib.DataCombo DcboStores 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   660
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   225
            Index           =   1
            Left            =   3630
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   255
            Index           =   6
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   660
            Width           =   765
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1125
         Index           =   1
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   1230
         _cx             =   2170
         _cy             =   1984
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
         ForeColor       =   128
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   " ÕœÌœ  «—ÌŒ «·»ÕÀ "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
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
         Begin MSComCtl2.DTPicker DtpFrom 
            Height          =   345
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   100073473
            CurrentDate     =   39209
         End
         Begin MSComCtl2.DTPicker DtpTO 
            Height          =   345
            Left            =   360
            TabIndex        =   6
            Top             =   630
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   100073473
            CurrentDate     =   39209
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   255
            Index           =   4
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   300
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   255
            Index           =   5
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   690
            Width           =   525
         End
      End
      Begin MSComctlLib.ProgressBar PrgBar 
         Height          =   300
         Left            =   30
         TabIndex        =   9
         Top             =   6570
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   570
         Left            =   30
         TabIndex        =   10
         Top             =   585
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1005
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
         ButtonImage     =   "FrmManStoreStock.frx":062D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   5385
         Left            =   30
         TabIndex        =   11
         Top             =   1170
         Width           =   10515
         _cx             =   18547
         _cy             =   9499
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmManStoreStock.frx":09C7
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   0
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   6885
         Width           =   10515
         _cx             =   18547
         _cy             =   979
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
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   90
            Width           =   1470
            _ExtentX        =   2593
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
            ButtonImage     =   "FrmManStoreStock.frx":0B25
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   390
            Index           =   3
            Left            =   8655
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   90
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·√’‰«›:-"
            ForeColor       =   &H00000080&
            Height          =   390
            Index           =   0
            Left            =   9435
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   90
            Width           =   1005
         End
         Begin VB.Image Img 
            Height          =   240
            Left            =   1665
            Picture         =   "FrmManStoreStock.frx":0EBF
            Top             =   150
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin ImpulseButton.ISButton CmdDo 
         Height          =   540
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   953
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ‰›Ì–"
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
         ButtonImage     =   "FrmManStoreStock.frx":1249
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
End
Attribute VB_Name = "FrmManStoreStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dcombos As ClsDataCombos

Private Sub CboReportType_Change()

    If Me.CboReportType.ListIndex = 0 Then
        Dcombos.GetStores Me.DcboStores

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(6).Caption = "«”„ «·„Œ“‰"
        Else
            Me.lbl(6).Caption = "Store Name"
        End If

        Me.FG.ColHidden(FG.ColIndex("CusName")) = True
        Me.FG.ColHidden(FG.ColIndex("DateGoIN")) = False
        Me.FG.TextMatrix(0, FG.ColIndex("DateGoIN")) = " «—ÌŒ «·⁄„·Ì…"
        Me.Ele(3).Visible = True
    
    ElseIf Me.CboReportType.ListIndex = 1 Then
        Dcombos.GetCustomersSuppliers 2, Me.DcboStores, True

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(6).Caption = "«”„ «·„Ê—œ"
        Else
            Me.lbl(6).Caption = "Supplier Name"
        End If

        Me.FG.ColHidden(FG.ColIndex("CusName")) = False
        Me.FG.ColHidden(FG.ColIndex("DateGoIN")) = False
        Me.Ele(3).Visible = False
    ElseIf Me.CboReportType.ListIndex = 2 Then
        Dcombos.GetStores Me.DcboStores

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(6).Caption = "«”„ «·„Œ“‰"
        Else
            Me.lbl(6).Caption = "Store Name"
        End If
    End If

End Sub

Private Sub CboReportType_Click()
    CboReportType_Change
End Sub

Private Sub CmdDo_Click()
    Dim Msg As String

    If Me.CboReportType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboReportType.SetFocus
        Exit Sub
    ElseIf Me.CboReportType.ListIndex = 0 Then

        If Me.DcboStores.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboStores.SetFocus
            Exit Sub
        End If

        StoreStock
    ElseIf Me.CboReportType.ListIndex = 1 Then
        SupManStock
    End If

End Sub

Private Sub CmdPrint_Click()
    Dim cItemsReport As ClsItemsReport

    If CboReportType.ListIndex = -1 Then
    ElseIf Me.CboReportType.ListIndex = 0 Then
        Set cItemsReport = New ClsItemsReport

        If Me.Opt(0).value = True Then
            cItemsReport.ManStockReports val(Me.DcboStores.BoundText), 1
        ElseIf Me.Opt(1).value = True Then
            cItemsReport.ManStockReports val(Me.DcboStores.BoundText), 3
        Else
            cItemsReport.ManStockReports val(Me.DcboStores.BoundText), 0
        End If
    
        Set cItemsReport = Nothing
    ElseIf Me.CboReportType.ListIndex = 1 Then
        Set cItemsReport = New ClsItemsReport
        cItemsReport.SupManStock val(Me.DcboStores.BoundText)
        Set cItemsReport = Nothing
    End If

End Sub

Private Sub Form_Load()

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.CboReportType
        .Clear
        .AddItem "Ã—œ „Œ“‰ «·’Ì«‰…"
        .AddItem "«·√’‰«› «·„ÊÃÊœ… ›Ï «·÷„«‰"
        '.AddItem "√’‰«› «— Ã⁄  „‰ «·÷„«‰ Ê·„  ”·„"
    End With

    With Me.FG
        .RowHeightMin = 320
        .AutoSizeMode = flexAutoSizeColWidth
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Set Dcombos = New ClsDataCombos
    Dcombos.GetStores Me.DcboStores
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
    Me.Height = 9240
    Me.Width = 11100
    Resize_Form Me
End Sub

Private Sub StoreStock()
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrStartDate As String
    Dim StrStartQtyDate As String
    Dim StrEndDate As String
    Dim SngItemCostPrice As Single
    Dim LngItemID As Long
    Dim BolInProgress As Boolean
    Dim Msg As String

    If IsNull(Me.DTPFrom.value) Then
        StrStartDate = SQLDate(#1/1/1901#, True)
        StrStartQtyDate = SQLDate(#1/1/1901#, True)
    Else
        StrStartDate = SQLDate(Me.DTPFrom.value, True)
        StrStartQtyDate = SQLDate(Me.DTPFrom.value - 1, True)
    End If

    If IsNull(Me.DTPTo.value) Then
        StrEndDate = SQLDate(#1/1/2079#, True)
    Else
        StrEndDate = SQLDate(Me.DTPTo.value, True)
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT     QryManStockComplete.QTY, QryManStockComplete.ItemID, QryManStockCompl" & "ete.ItemCode, QryManStockComplete.ItemName,                        QryManStockCo" & "mplete.HaveSerial, QryManStockComplete.ItemSerial, QryManStockComplete.TicketNO," & " QryManStockComplete.StoreID,                        QryManStockComplete.StoreNa" & "me, dbo.TblMaintenece.MaintananceID, dbo.TblMaintenece.DateGoIN, dbo.TblMaintene" & "ce.ReciptNumber,                        dbo.TblMaintenece.ManOperationTypeID FRO" & "M         dbo.QryManStockComplete(0) QryManStockComplete INNER JOIN             " & "          dbo.QryManLastTicketTrans ON QryManStockComplete.TicketNO = dbo.QryMan" & "LastTicketTrans.TicketNO INNER JOIN                       dbo.TblMaintenece ON d" & "bo.QryManLastTicketTrans.LastTrans = dbo.TblMaintenece.MaintananceID"
        StrSQL = StrSQL + " Where QryManStockComplete.StoreID=" & val(Me.DcboStores.BoundText)

        If Me.Opt(0).value = True Then
            StrSQL = StrSQL + " AND dbo.TblMaintenece.ManOperationTypeID=1"
        ElseIf Me.Opt(1).value = True Then
            StrSQL = StrSQL + " AND dbo.TblMaintenece.ManOperationTypeID=3"
        Else
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Exit Sub
    End If

    With Me.FG
        .Rows = .FixedRows
    End With

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    Else

        With Me.FG
            Me.PrgBar.Visible = True
            Me.PrgBar.Max = rs.RecordCount
            .Rows = .FixedRows + rs.RecordCount

            For i = .FixedRows To .Rows - 1
                BolInProgress = True

                DoEvents
                Me.PrgBar.value = i
                Me.lbl(3).Caption = i
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)

                If Not IsNull(rs("DateGoIN").value) Then
                    .TextMatrix(i, .ColIndex("DateGoIN")) = DisplayDate(rs("DateGoIN").value)
                End If
            
                rs.MoveNext
                .ShowCell i, IIf(.Col = -1, 1, .Col)
            Next i

            BolInProgress = False
            .AutoSize 0, .Cols - 1, False
            Me.PrgBar.Visible = False
            Me.PrgBar.value = 0
        End With

    End If

End Sub

Private Sub SupManStock()
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrStartDate As String
    Dim StrStartQtyDate As String
    Dim StrEndDate As String
    Dim SngItemCostPrice As Single
    Dim LngItemID As Long
    Dim BolInProgress As Boolean
    Dim Msg As String

    If IsNull(Me.DTPFrom.value) Then
        StrStartDate = SQLDate(#1/1/1901#, True)
        StrStartQtyDate = SQLDate(#1/1/1901#, True)
    Else
        StrStartDate = SQLDate(Me.DTPFrom.value, True)
        StrStartQtyDate = SQLDate(Me.DTPFrom.value - 1, True)
    End If

    If IsNull(Me.DTPTo.value) Then
        StrEndDate = SQLDate(#1/1/2079#, True)
    Else
        StrEndDate = SQLDate(Me.DTPTo.value, True)
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblMaintenece.MaintananceID, dbo.TblMaintenece.ReciptNumber, dbo.TblMaintenece.CashCustomerName, dbo.TblCustemers.CusName, "
        StrSQL = StrSQL + "dbo.TblMaintenece.DateGoIN, QryManSupStockComplete.QTY, QryManSupStockComplete.ItemID, QryManSupStockComplete.ItemCode,"
        StrSQL = StrSQL + "QryManSupStockComplete.ItemName, QryManSupStockComplete.ItemSerial, QryManSupStockComplete.TicketNO,"
        StrSQL = StrSQL + "QryManSupStockComplete.HaveSerial , QryManSupStockComplete.CusID"
        StrSQL = StrSQL + " FROM         dbo.QryManSupStockComplete(0) QryManSupStockComplete INNER JOIN"
        StrSQL = StrSQL + " dbo.TblCustemers ON QryManSupStockComplete.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL + " dbo.TblMainteneceDetails INNER JOIN"
        StrSQL = StrSQL + " dbo.TblMaintenece ON dbo.TblMainteneceDetails.MaintananceID = dbo.TblMaintenece.MaintananceID ON"
        StrSQL = StrSQL + " QryManSupStockComplete.TicketNO = dbo.TblMainteneceDetails.TicketNO"
        StrSQL = StrSQL + " Where (dbo.TblMaintenece.ManOperationTypeID = 2)"

        If Me.DcboStores.BoundText <> "" Then
            StrSQL = StrSQL + " AND QryManSupStockComplete.CusID=" & Me.DcboStores.BoundText
        End If

        StrSQL = StrSQL + " Order By QryManSupStockComplete.ItemName"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Exit Sub
    End If

    With Me.FG
        .Rows = .FixedRows
    End With

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    Else

        With Me.FG
            Me.PrgBar.Visible = True
            Me.PrgBar.Max = rs.RecordCount
            .Rows = .FixedRows + rs.RecordCount

            For i = .FixedRows To .Rows - 1
                BolInProgress = True

                DoEvents
                Me.PrgBar.value = i
                Me.lbl(3).Caption = i
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)

                If Not IsNull(rs("DateGoIN").value) Then
                    .TextMatrix(i, .ColIndex("DateGoIN")) = DisplayDate(rs("DateGoIN").value)
                End If

                rs.MoveNext
                .ShowCell i, IIf(.Col = -1, 1, .Col)
            Next i

            BolInProgress = False
            .AutoSize 0, .Cols - 1, False
            Me.PrgBar.Visible = False
            Me.PrgBar.value = 0
        End With

    End If

End Sub
