VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmEvaluation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…"
   ClientHeight    =   9840
   ClientLeft      =   4395
   ClientTop       =   2295
   ClientWidth     =   12135
   Icon            =   "FrmEvaluation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   12135
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9840
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12135
      _cx             =   21405
      _cy             =   17357
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
      Begin C1SizerLibCtl.C1Elastic pnlGrid 
         Height          =   4635
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3690
         Width           =   11895
         _cx             =   20981
         _cy             =   8176
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
         BackColor       =   -2147483633
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   4485
            Left            =   45
            TabIndex        =   32
            Top             =   -30
            Width           =   11820
            _cx             =   20849
            _cy             =   7911
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16776960
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   25
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEvaluation.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
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
            Begin C1SizerLibCtl.C1Elastic frm_Chart 
               Height          =   4212
               Left            =   480
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   360
               Visible         =   0   'False
               Width           =   9972
               _cx             =   17595
               _cy             =   7435
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
               BackColor       =   -2147483633
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
               Begin VB.TextBox txtStandered 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1920
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   3720
                  Width           =   2796
               End
               Begin VB.TextBox txtEmp 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   6120
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   3720
                  Width           =   2796
               End
               Begin VB.CommandButton bClose 
                  BackColor       =   &H000000FF&
                  Caption         =   "X"
                  Height          =   375
                  Left            =   9600
                  Style           =   1  'Graphical
                  TabIndex        =   53
                  Top             =   0
                  Width           =   372
               End
               Begin MSChart20Lib.MSChart chrt_emp 
                  Height          =   3540
                  Left            =   0
                  OleObjectBlob   =   "FrmEvaluation.frx":0769
                  TabIndex        =   52
                  Top             =   120
                  Width           =   9888
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„⁄Ì«—"
                  Height          =   336
                  Index           =   1
                  Left            =   4356
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   3720
                  Width           =   1296
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ÊŸ›"
                  Height          =   336
                  Index           =   0
                  Left            =   8556
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   3720
                  Width           =   1296
               End
            End
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   ">>"
            Height          =   228
            Left            =   96
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   0
            Visible         =   0   'False
            Width           =   324
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   510
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   8400
         Width           =   6360
         _cx             =   11218
         _cy             =   900
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
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   3675
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   120
            Width           =   750
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   2
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   4
            Left            =   750
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   1110
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   735
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   12015
         _cx             =   21193
         _cy             =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…     "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   1
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   5
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation.frx":2C21
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation.frx":2FBB
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   7
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation.frx":3355
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   3
            Left            =   615
            TabIndex        =   8
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation.frx":36EF
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
      End
      Begin C1SizerLibCtl.C1Elastic pnlHeader 
         Height          =   2784
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   11832
         _cx             =   20876
         _cy             =   4921
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
         Begin VB.TextBox Emp_Code 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9864
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   600
            Width           =   840
         End
         Begin VB.ComboBox YearID 
            Height          =   288
            Left            =   4872
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   600
            Width           =   1860
         End
         Begin VB.ComboBox MonthID 
            Height          =   288
            Left            =   1896
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   600
            Width           =   1872
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   1452
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1200
            Width           =   11532
            _cx             =   20346
            _cy             =   2566
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
            Caption         =   "·„‘—Ê⁄ „Õœœ"
            Align           =   0
            AutoSizeChildren=   7
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
            Begin VB.CheckBox opt_Branch 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·›—⁄ „Õœœ"
               Height          =   312
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   360
               Width           =   1212
            End
            Begin VB.CheckBox opt_Project 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·„‘—Ê⁄ „Õœœ"
               Height          =   312
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   720
               Width           =   1212
            End
            Begin VB.CheckBox opt_Managerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·√œ«—… „Õœœ…"
               Height          =   312
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1080
               Width           =   1212
            End
            Begin VB.TextBox oneEmp_Code 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   3648
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   720
               Width           =   840
            End
            Begin VB.CommandButton btnView 
               Caption         =   "⁄—÷"
               Height          =   372
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   480
               Width           =   1200
            End
            Begin VB.OptionButton opt_AllEmployees 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·ﬂ· «·„ÊŸ›Ì‰"
               Height          =   192
               Left            =   4344
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   360
               Width           =   1428
            End
            Begin VB.OptionButton opt_OneEmployee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·„ÊŸ› „Õœœ"
               Height          =   192
               Left            =   4476
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   720
               Width           =   1320
            End
            Begin MSDataListLib.DataCombo Branch 
               Height          =   288
               Left            =   6588
               TabIndex        =   38
               Top             =   360
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Mangerial 
               Height          =   288
               Left            =   6576
               TabIndex        =   39
               Top             =   1080
               Width           =   2712
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Project 
               Height          =   288
               Left            =   6576
               TabIndex        =   40
               Top             =   720
               Width           =   2712
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo OneEmployee 
               Height          =   288
               Left            =   1740
               TabIndex        =   41
               Top             =   720
               Width           =   1884
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox ID 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7956
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   2796
         End
         Begin MSComCtl2.DTPicker SDate 
            Height          =   312
            Left            =   4908
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95551491
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo BranchID 
            Height          =   288
            Left            =   1896
            TabIndex        =   27
            Top             =   120
            Width           =   1872
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo EmployeeID 
            Height          =   288
            Left            =   7956
            TabIndex        =   29
            Top             =   600
            Width           =   1884
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘Â— "
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3924
            TabIndex        =   34
            Top             =   600
            Width           =   852
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   6876
            TabIndex        =   33
            Top             =   600
            Width           =   744
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ«∆„ »«· ﬁÌÌ„"
            Height          =   312
            Index           =   6
            Left            =   10404
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   1296
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   24
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   120
            Width           =   636
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·ÌÊ„"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   6912
            TabIndex        =   26
            Top             =   120
            Width           =   744
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   336
            Index           =   8
            Left            =   10392
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   1296
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   750
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   9090
         Width           =   12135
         _cx             =   21405
         _cy             =   1323
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
         Align           =   2
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   10830
            TabIndex        =   16
            Top             =   120
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
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
            ButtonImage     =   "FrmEvaluation.frx":3A89
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   9525
            TabIndex        =   17
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
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
            ButtonImage     =   "FrmEvaluation.frx":A2EB
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   8010
            TabIndex        =   18
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   873
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
            ButtonImage     =   "FrmEvaluation.frx":10B4D
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   3
            Left            =   6900
            TabIndex        =   19
            Top             =   120
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
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
            ButtonImage     =   "FrmEvaluation.frx":173AF
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   4890
            TabIndex        =   20
            Top             =   120
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
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
            ButtonImage     =   "FrmEvaluation.frx":1DC11
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   6
            Left            =   1245
            TabIndex        =   21
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            ButtonImage     =   "FrmEvaluation.frx":24473
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   495
            Left            =   105
            TabIndex        =   22
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "«·„—›ﬁ« "
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
            ButtonImage     =   "FrmEvaluation.frx":4E095
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   7
            Left            =   3915
            TabIndex        =   23
            Top             =   120
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
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
            ButtonImage     =   "FrmEvaluation.frx":548F7
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   9
            Left            =   2355
            TabIndex        =   24
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
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
            ButtonImage     =   "FrmEvaluation.frx":5B159
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
      End
   End
End
Attribute VB_Name = "FrmEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp2 As ADODB.Recordset

Dim TTP As clstooltip
Dim Direction As Integer
Private Sub C1Elastic4_Click()

End Sub

Private Sub Fill_Grid()

If YearID.ListIndex = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«Œ — «·”‰… «Ê·« ")
    Else
        MsgBox ("Please , Select the year first")
    End If
       Exit Sub
End If


If MonthID.ListIndex = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«Œ — «·‘Â— «Ê·« ")
    Else
        MsgBox ("Please , Select the Month first")
    End If
       Exit Sub
End If



Dim i As Integer
Dim WeakFrom  As Double, WeakTo As Double, InterFrom As Double, InterTo As Double, GoodFrom As Double, GoodTo As Double, VeryGFrom As Double, VeryGTo As Double, ExcelFrom As Double, ExcelTo As Double

Grid.Rows = Grid.FixedRows

Rs_Temp.Open Selection_Query, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs_Temp.RecordCount > 0 Then
    Grid.Rows = Grid.FixedRows + Rs_Temp.RecordCount
    With Grid
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
                .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
                .TextMatrix(i, .ColIndex("MaxDgree")) = IIf(IsNull(Rs_Temp("MaxDgree").value), "", Rs_Temp("MaxDgree").value)
                .TextMatrix(i, .ColIndex("StanderedID")) = IIf(IsNull(Rs_Temp("StanderedID").value), "", Rs_Temp("StanderedID").value)
                .TextMatrix(i, .ColIndex("emp_name")) = IIf(IsNull(Rs_Temp("emp_name").value), "", Rs_Temp("emp_name").value)
                .TextMatrix(i, .ColIndex("StanderedName")) = IIf(IsNull(Rs_Temp("StanderedName").value), "", Rs_Temp("StanderedName").value)
                .TextMatrix(i, .ColIndex("PreDegree")) = IIf(IsNull(Rs_Temp("Eval_Degree").value), "", Rs_Temp("Eval_Degree").value)
                 .TextMatrix(i, .ColIndex("Curr_Dynamic")) = Dynamic_Degree(val(.TextMatrix(i, .ColIndex("Emp_ID"))), _
                 val(.TextMatrix(i, .ColIndex("StanderedID"))), val(YearID.ListIndex), val(MonthID.ListIndex))
                   .TextMatrix(i, .ColIndex("sum_Degrees")) = val(.TextMatrix(i, .ColIndex("Manual_Degree"))) + val(.TextMatrix(i, .ColIndex("Curr_Dynamic")))
                   
            '      .TextMatrix(i, .ColIndex("sum_Degrees")) = val(.TextMatrix(i, .ColIndex("sum_Degrees"))) + val(.TextMatrix(i, .ColIndex("Manual_Degree")))
                 If (val(.TextMatrix(i, .ColIndex("sum_Degrees"))) > val(.TextMatrix(i, .ColIndex("MaxDgree")))) Then
                              .TextMatrix(i, .ColIndex("sum_Degrees")) = val(.TextMatrix(i, .ColIndex("MaxDgree")))
                 End If
                 
                 
                        WeakFrom = IIf(IsNull(Rs_Temp("WeakFrom").value), 0, Rs_Temp("WeakFrom").value)
                        WeakTo = IIf(IsNull(Rs_Temp("WeakTo").value), 0, Rs_Temp("WeakTo").value)
                        InterFrom = IIf(IsNull(Rs_Temp("InterFrom").value), 0, Rs_Temp("InterFrom").value)
                        InterTo = IIf(IsNull(Rs_Temp("InterTo").value), 0, Rs_Temp("InterTo").value)
                        GoodFrom = IIf(IsNull(Rs_Temp("GoodFrom").value), 0, Rs_Temp("GoodFrom").value)
                        GoodTo = IIf(IsNull(Rs_Temp("GoodTo").value), 0, Rs_Temp("GoodTo").value)
                        VeryGFrom = IIf(IsNull(Rs_Temp("VeryGFrom").value), 0, Rs_Temp("VeryGFrom").value)
                        VeryGTo = IIf(IsNull(Rs_Temp("VeryGTo").value), 0, Rs_Temp("VeryGTo").value)
                        ExcelFrom = IIf(IsNull(Rs_Temp("ExcelFrom").value), 0, Rs_Temp("ExcelFrom").value)
                        ExcelTo = IIf(IsNull(Rs_Temp("ExcelTo").value), 0, Rs_Temp("ExcelTo").value)
                        
                        .TextMatrix(i, .ColIndex("WeakFrom")) = WeakFrom
                        .TextMatrix(i, .ColIndex("WeakTo")) = WeakTo
                        .TextMatrix(i, .ColIndex("GoodFrom")) = GoodFrom
                        .TextMatrix(i, .ColIndex("GoodTo")) = GoodTo
                        .TextMatrix(i, .ColIndex("InterFrom")) = InterFrom
                        .TextMatrix(i, .ColIndex("InterTo")) = InterTo
                        .TextMatrix(i, .ColIndex("VeryGFrom")) = VeryGFrom
                        .TextMatrix(i, .ColIndex("VeryGTo")) = VeryGTo
                        .TextMatrix(i, .ColIndex("ExcelFrom")) = ExcelFrom
                        .TextMatrix(i, .ColIndex("ExcelTo")) = ExcelTo
                        
                        .TextMatrix(i, .ColIndex("Final_Evaluation")) = Eval_Title(val(.TextMatrix(i, .ColIndex("sum_Degrees"))), WeakFrom, WeakTo, InterFrom, InterTo, GoodFrom, GoodTo, VeryGFrom, VeryGTo, ExcelFrom, ExcelTo)
                Rs_Temp.MoveNext
                
            Next
    End With
End If


End Sub

Private Function Eval_Title(sum_Degrees As Double, WeakFrom As Double, WeakTo As Double, InterFrom As Double, InterTo As Double, GoodFrom As Double, GoodTo As Double, VeryGFrom As Double, VeryGTo As Double, ExcelFrom As Double, ExcelTo As Double) As String

Dim str  As String

If sum_Degrees <= WeakTo Then
            str = " ÷⁄Ì› "
            Eval_Title = str
End If

If sum_Degrees >= InterFrom And sum_Degrees <= InterTo Then
            str = " „ Ê”ÿ "
            Eval_Title = str
End If

If sum_Degrees >= GoodFrom And sum_Degrees <= GoodTo Then
            str = " ÃÌœ "
            Eval_Title = str
End If

If sum_Degrees >= VeryGFrom And sum_Degrees <= VeryGTo Then
            str = " ÃÌœ Ãœ« "
            Eval_Title = str
End If

If sum_Degrees >= ExcelFrom Then
            str = " „„ «“ "
            Eval_Title = str
End If

End Function



Private Function Dynamic_Degree(Emp_id As Integer, stndr As Integer, yr As Integer, mnth As Integer) As Double


Dim str As String
Set Rs_Temp2 = New ADODB.Recordset
Dim NoofHour  As Integer, NoofDays As Integer

str = str & " SELECT  NoofHour , NoofDays ,   dbo.TblEvaluationStandered.ID, dbo.TblChangedComponentRegister.month, dbo.TblChangedComponentRegister.year, dbo.TblChangedComponentRegister.ComponentID,"
str = str & " dbo.TblChangedComponentRegisterDetails.Emp_id, dbo.TblEvaluationStandered_Details.AllowanceID, dbo.TblEvaluationStandered_Details.AllowanceName,"
str = str & " dbo.TblEvaluationStandered_Details.InfluenceType , dbo.TblEvaluationStandered_Details.points"
str = str & " FROM     dbo.TblChangedComponentRegisterDetails INNER JOIN"
str = str & " dbo.TblChangedComponentRegister ON"
str = str & " dbo.TblChangedComponentRegisterDetails.ChangedComponentid = dbo.TblChangedComponentRegister.ChangedComponentid INNER JOIN"
str = str & " dbo.TblEvaluationStandered_Details INNER JOIN"
str = str & " dbo.TblEvaluationStandered ON dbo.TblEvaluationStandered_Details.HID = dbo.TblEvaluationStandered.ID ON"
str = str & " dbo.TblChangedComponentRegister.ComponentID = dbo.TblEvaluationStandered_Details.AllowanceID"

str = str & "       where 1 =1 "
str = str & " and TblChangedComponentRegisterDetails.Emp_id = " & Emp_id
str = str & " and  dbo.TblEvaluationStandered.ID = " & stndr
str = str & " and dbo.TblChangedComponentRegister.month = " & mnth
str = str & " and dbo.TblChangedComponentRegister.year = " & yr

Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
Dim points  As Double
Dim Allpoints  As Double
Dim inf As Integer
Dim i As Integer

If Rs_Temp2.RecordCount > 0 Then
    
    For i = 0 To Rs_Temp2.RecordCount - 1
    
            points = IIf(IsNull(Rs_Temp2("Points").value), 0, Rs_Temp2("Points").value)
            NoofHour = IIf(IsNull(Rs_Temp2("NoofHour").value), 0, Rs_Temp2("NoofHour").value)
            NoofDays = IIf(IsNull(Rs_Temp2("NoofDays").value), 0, Rs_Temp2("NoofDays").value)
            inf = IIf(IsNull(Rs_Temp2("InfluenceType").value), 0, Rs_Temp2("InfluenceType").value)
            If inf = 1 Then
                    points = points
            ElseIf inf = 2 Then
                    points = points * (-1)
            End If
            
            If NoofDays > 0 Then
                    points = points * NoofDays
            ElseIf NoofDays > 0 Then
                    points = points * NoofDays
            End If
            
            Allpoints = Allpoints + points
            Rs_Temp2.MoveNext
    Next
    
End If

Dynamic_Degree = Allpoints

End Function


Private Function Selection_Query() As String

Dim str As String
Set Rs_Temp = New ADODB.Recordset

str = str & "  select H.FullCode , H.Emp_ID , H.emp_name , D.ID StanderedID ,  D.Ename StanderedName , d.MaxDgree   ,"
str = str & "  D.WeakFrom ,D.WeakTo ,D.InterFrom , D.InterTo ,D.GoodFrom , D.GoodTo , D.VeryGFrom , D.VeryGTo , D.ExcelFrom , D.ExcelTo  ,"
str = str & "  (select Eval_Degree from TblEvaluation_Employee where  StanderedID = D.ID  and Emp_ID  = H.emp_id and  YearNo = " & val(YearID.ListIndex) & " and  MonthNo =  " & val(MonthID.ListIndex) - 1 & " )  Eval_Degree"
str = str & "   from tblemployee H  ,TblEvaluationStandered D "

str = str & " where 1 = 1 "

If opt_OneEmployee.value = True Then
        str = str & " and   H.Emp_ID =  " & val(OneEmployee.BoundText)
ElseIf opt_Branch.value = True Then
        str = str & "  and  H.branchid  =  " & val(Branch.BoundText)
ElseIf opt_Project = True Then
        str = str & "  and H.project_id =  " & val(Project.BoundText)
ElseIf opt_Managerial = True Then
        str = str & "  and  H.RegionID =  " & val(Mangerial.BoundText)
End If

str = str & "  and H.Emp_ID <>  " & val(EmployeeID.BoundText)
str = str & "  order by h.Emp_Name  "

Selection_Query = str
End Function

Private Sub bClose_Click()
frm_Chart.Visible = False
End Sub

Private Sub Branch_Change()

Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set Project.RowSource = Rs_Temp

str = " select id,Project_name from projects where branch_no = " & val(Branch.BoundText)

fill_combo Project, str
Project.Refresh

OneEmployee.BoundText = ""
oneEmp_Code.Text = ""

LoadEmployee
End Sub

Private Sub btnView_Click()
        Fill_Grid
End Sub

Private Sub CheckUSer()

Dim str As String

Set Rs_Temp = New ADODB.Recordset
str = " SELECT * from tblusers where userid =  " & user_id
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
Dim II As Integer
If Rs_Temp.RecordCount > 0 Then
            If Not IsNull(Rs_Temp("EmpID").value) Then
                        II = IIf(IsNull(Rs_Temp("EmpID").value), "", Rs_Temp("EmpID").value)
                        EmployeeID.BoundText = II
            End If
End If

End Sub


Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Select Case Index
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            ID.Text = CStr(new_id("TblEmpEvaluation", "ID", "", True))
           opt_AllEmployees.value = True
           Grid.Rows = Grid.FixedRows
           'Grid.Rows = Grid.FixedRows + 10
           
          Dim mm As Integer
          mm = Month(DateTime.Now)
           
           YearID.ListIndex = year(DateTime.Now) - 2006
           MonthID.ListIndex = mm - 1
           CheckUSer
    
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Grid.Rows = Grid.Rows + 1
        Case 2

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Action

        Case 5

        Case 6
                Unload Me
         Case 7
                 print_report2
         Case 9
         
         Unload FrmInsurancesSearch
         FrmInsurancesSearch.SendForm = 0
         FrmInsurancesSearch.show vbModal
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

 



Private Sub FlightNo_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.FlightNo.text, 1)
End Sub

Private Sub Emp_Code_Change()


Dim val1, val2, str As String
If Emp_Code.Text = "" Then Exit Sub
'EmployeeID.BoundText = ""

    str = " select * From TblEmployee where  fullcode = '" & Emp_Code.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
     Else
        val1 = ""
    End If
    EmployeeID.BoundText = val1
End Sub

Private Sub EmployeeID_Change()
    Dim val1, val2, str As String
    If EmployeeID.BoundText = "" Then Exit Sub
    Emp_Code.Text = ""
    
        str = " select * From TblEmployee where  Emp_ID = " & val(EmployeeID.BoundText)
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs_Temp.RecordCount > 0 Then
            Rs_Temp.MoveFirst '
            val1 = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
         Else
            val1 = ""
        End If
    
    Emp_Code.Text = val1
    OneEmployee.BoundText = ""

    oneEmp_Code.Text = ""
    Grid.Rows = Grid.FixedRows
    LoadEmployee
    
End Sub

Private Sub Form_Activate()
'    txtid.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fill_Combos()
    Dim Dcombos As ClsDataCombos
    Dim str As String
    
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetBranches BranchID
    Dcombos.GetBranches Branch
    
   ' Dcombos.GetEmployees OneEmployee
    
    Dcombos.GetEmployees EmployeeID
   
  If SystemOptions.UserInterface = ArabicInterface Then
        str = "SELECT Emp_ID,Emp_Name From TblEmployee Where 1=1  "
Else
        str = "SELECT Emp_ID,Emp_Namee From TblEmployee Where 1=1  "
End If
 fill_combo OneEmployee, str
   
    Dcombos.GetSection Me.Mangerial
    
    str = " select id,Project_name from projects"
    fill_combo Project, str
    
  '  Dcombos getpro
    Dim i As Integer
     For i = 2006 To 2050
        YearID.AddItem i
    Next
   
       For i = 1 To 12
        MonthID.AddItem MonthName(i)
    Next
   
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub

Private Sub LoadEmployee()
Dim StrSQL As String

Set Rs_Temp = New ADODB.Recordset
Set OneEmployee.RowSource = Rs_Temp

If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Emp_ID,Emp_Name From TblEmployee Where 1=1  "
Else
        StrSQL = "SELECT Emp_ID,Emp_Namee From TblEmployee Where 1=1  "
End If

If opt_Branch.value = 1 Then
        StrSQL = StrSQL & " and BranchId = " & val(Branch.BoundText)
End If

If opt_Project.value = 1 Then
        StrSQL = StrSQL & " and Project_ID = " & val(Project.BoundText)
End If

If opt_Managerial.value = 1 Then
        StrSQL = StrSQL & " and SectionID = " & val(Mangerial.BoundText)
End If

StrSQL = StrSQL & " and  Emp_ID  <>  " & val(EmployeeID.BoundText)

fill_combo OneEmployee, StrSQL
OneEmployee.Refresh
    
    
End Sub

Private Sub Form_Load()
    
   On Error GoTo ErrTrap
       
    Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
    ChangeLang
    SetInterface Me

    End If
''
    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  „·› «·„œ«—”  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""


'
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
'
    
    
    Dim StrSQL As String
    StrSQL = ""

     If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = "SELECT  *  From TblEmpEvaluation    "
     Else
            StrSQL = "SELECT  *  From TblEmpEvaluation"
     End If
     rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    
    If OPEN_NEW_SCREEN = True Then
            Cmd_Click (0)
    End If

    Exit Sub

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
    EleHeader.Caption = "Manual Evaluation "
    Me.Caption = EleHeader.Caption
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(9).Caption = "Search"
    Cmd(6).Caption = "Exit"
    Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"
    
    lbl(8).Caption = "Ser"
    Label3.Caption = "Today date"
    lbl(24).Caption = "Branch"
    lbl(6).Caption = "Evaluation supervisor"
    Label1.Caption = "Year"
    Label2.Caption = "Month"
    
    C1Elastic2.Caption = "For Specific Project"
    opt_Branch.Caption = "Selected branch"
    opt_Project.Caption = "Selected Branch"
    opt_Managerial.Caption = "Selected department"
    opt_AllEmployees.Caption = "All Employees"
    opt_OneEmployee.Caption = "Specific Employee"
    btnView.Caption = "Show"
   
    With Grid
        .TextMatrix(0, .ColIndex("FullCode")) = "Emp Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Name"
        .TextMatrix(0, .ColIndex("StanderedName")) = "Standereds"
        .TextMatrix(0, .ColIndex("MaxDgree")) = "Max Mark"
        .TextMatrix(0, .ColIndex("PreDegree")) = "Previous mark"
        .TextMatrix(0, .ColIndex("Curr_Dynamic")) = "Aurrent automatic mark"
        .TextMatrix(0, .ColIndex("Manual_Degree")) = "Manual mark"
        .TextMatrix(0, .ColIndex("sum_Degrees")) = "Total Mark"
        .TextMatrix(0, .ColIndex("Final_Evaluation")) = "Evaluation"
        .TextMatrix(0, .ColIndex("btn")) = "Flowchart"
    End With
   
    lbl(0).Caption = "Employee"
    lbl(1).Caption = "Standered"
    
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"
    




'lbl(9).Caption = "Last Contract"



End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub


Private Sub txtCount_Change()

End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode As String
Dim Msg As String
'  Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
Dim sql As String
Dim count As Integer
Dim Rate As Double
 
    With Grid

     Select Case .ColKey(Col)

             Case "MaxDgree"
                      '  StrAccountCode = .ComboData
                        .TextMatrix(Row, .ColIndex("sum_Degrees")) = val(.TextMatrix(Row, .ColIndex("MaxDgree")))
                        'Grid.Rows = Grid.Rows + 1
             Case "Manual_Degree"
                          .TextMatrix(Row, .ColIndex("sum_Degrees")) = val(.TextMatrix(Row, .ColIndex("Manual_Degree"))) + val(.TextMatrix(Row, .ColIndex("Curr_Dynamic")))
                         '.TextMatrix(Row, .ColIndex("Final_Evaluation")) = val(.TextMatrix(Row, .ColIndex("sum_Degrees"))) + val(.TextMatrix(Row, .ColIndex("Manual_Degree")))
                        
                         If (val(.TextMatrix(Row, .ColIndex("sum_Degrees"))) > val(.TextMatrix(Row, .ColIndex("MaxDgree")))) Then
                              .TextMatrix(Row, .ColIndex("sum_Degrees")) = val(.TextMatrix(Row, .ColIndex("MaxDgree")))
                            End If
                         
                          .TextMatrix(Row, .ColIndex("Final_Evaluation")) = Eval_Title(val(.TextMatrix(Row, .ColIndex("sum_Degrees"))), val(.TextMatrix(Row, .ColIndex("WeakFrom"))), val(.TextMatrix(Row, .ColIndex("WeakTo"))), _
                         val(.TextMatrix(Row, .ColIndex("InterFrom"))), val(.TextMatrix(Row, .ColIndex("InterTo"))), val(.TextMatrix(Row, .ColIndex("GoodFrom"))), val(.TextMatrix(Row, .ColIndex("GoodTo"))), _
                         val(.TextMatrix(Row, .ColIndex("VeryGFrom"))), val(.TextMatrix(Row, .ColIndex("VeryGTo"))), val(.TextMatrix(Row, .ColIndex("ExcelFrom"))), val(.TextMatrix(Row, .ColIndex("ExcelTo"))))
                                          
     End Select
End With




End Sub






Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Grid

If TxtModFlg.Text = "R" And .ColKey(Col) <> "btn" Then
          .ComboList = ""
          Cancel = True
End If


Select Case .ColKey(Col)
 Case "Emp_Name"
         .ComboList = ""
          Cancel = True
    Case "StanderedName"
         .ComboList = ""
          Cancel = True
    Case "PreDegree"
            .ComboList = ""
            Cancel = True
            
    Case "Curr_Dynamic"
         .ComboList = ""
          Cancel = True
    Case "MaxDgree"
            .ComboList = ""
            Cancel = True
            
    Case "sum_Degrees"
         .ComboList = ""
          Cancel = True
    Case "Final_Evaluation"
            .ComboList = ""
            Cancel = True
         Case "btn"
                .ComboList = ""
            
    End Select
 End With

End Sub

Private Sub Grid_Click()
 '   Charting (val(Grid.TextMatrix(Grid.Row, Grid.ColIndex("Emp_ID"))))
 '   frm_Chart.Visible = True
 
If Grid.Col = Grid.ColIndex("btn") Then
            frm_Chart.Visible = True
            txtEmp.Text = Grid.TextMatrix(Grid.Row, Grid.ColIndex("Emp_Name"))
            txtStandered.Text = Grid.TextMatrix(Grid.Row, Grid.ColIndex("StanderedName"))
End If
 
End Sub

Private Sub Charting(Emp_id As Integer)
Dim str As String, i As Integer
str = " select EName , ENameE , final_evaluation  , sum_Degrees  from TblEvaluation_Details  E , TblEvaluationStandered S where E.StanderedID = S.ID and Emp_ID  = " & Emp_id & " and E.HID  =  " & val(ID.Text)
Set Rs_Temp = New ADODB.Recordset

Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
    
    chrt_emp.ShowLegend = True
    chrt_emp.ColumnCount = Rs_Temp.RecordCount
    
    For i = 0 To Rs_Temp.RecordCount - 1
             
                chrt_emp.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_emp.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_emp.RowLabel = "Chart"
                End If
                
                chrt_emp.Column = i + 1
                chrt_emp.Row = 1
                chrt_emp.Data = IIf(IsNull(Rs_Temp("sum_Degrees").value), 0, Rs_Temp("sum_Degrees").value)
                chrt_emp.ColumnLabel = IIf(IsNull(Rs_Temp("EName").value), 0, Rs_Temp("EName").value)
                Rs_Temp.MoveNext
    Next
End If

End Sub


Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


'Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With Grid
     
     Select Case .ColKey(Col)
     
       Case "SID"
          Set Rs_Temp = New ADODB.Recordset
          
          StrSQL = " Select ID,EName From TblEmpEvaluationStandered  "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          
           StrComboList = Grid.BuildComboList(Rs_Temp, "EName", "ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
                            
        Case "ToCity"
           
           Set Rs_Temp = New ADODB.Recordset
          StrSQL = " Select GovernmentID,GovernmentName From TblCountriesGovernments  "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = Grid.BuildComboList(Rs_Temp, "GovernmentName", "GovernmentID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
           
          Case "btn"
               .ColComboList(.ColIndex("btn")) = "..."
           
     End Select
   End With
   



End Sub

 









Private Sub Label4_Click()


'chrt_emp.Visible = False
'Grid.Width = pnlGrid.Width
'Grid.left = 10
'Grid.Refresh
'Grid.Redraw = flexRDDirect

RepaintGrid

If Direction = 1 Then
        Direction = 0
ElseIf Direction = 0 Then
        Direction = 1
End If


End Sub

Private Sub RepaintGrid()

If Direction = 1 Then
        chrt_emp.Visible = True
        Grid.Width = pnlGrid.Width - chrt_emp.Width
        Grid.left = chrt_emp.Width
        Grid.Refresh
    '    Dir = 0
ElseIf Direction = 0 Then
        chrt_emp.Visible = False
        Grid.Width = pnlGrid.Width
        Grid.left = 10
        Grid.Refresh
     '   Dir = 1
End If


End Sub



Private Sub Mangerial_Change()
     LoadEmployee
     OneEmployee.BoundText = ""
     oneEmp_Code.Text = ""
End Sub


Private Sub oneEmp_Code_Change()


Dim val1, val2, str As String
If oneEmp_Code.Text = "" Then Exit Sub
'EmployeeID.BoundText = ""

    str = " select * From TblEmployee where  fullcode = '" & oneEmp_Code.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
     Else
        val1 = ""
    End If
    OneEmployee.BoundText = val1

End Sub

Private Sub OneEmployee_Change()
Dim val1, val2, str As String
If OneEmployee.BoundText = "" Then Exit Sub
oneEmp_Code.Text = ""

    str = " select * From TblEmployee where  Emp_ID = " & val(OneEmployee.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
     Else
        val1 = ""
    End If
    
    oneEmp_Code.Text = val1

End Sub

Private Sub opt_AllEmployees_Click()

'OneEmployee.BoundText = ""
'Branch.BoundText = ""
'Project.BoundText = ""
'Mangerial.BoundText = ""

OneEmployee.Enabled = True
oneEmp_Code.Enabled = True


End Sub

Private Sub opt_Branch_Click()
OneEmployee.BoundText = ""
oneEmp_Code.Text = ""
End Sub

Private Sub opt_Managerial_Click()
        OneEmployee.BoundText = ""
        oneEmp_Code.Text = ""
End Sub

Private Sub opt_OneEmployee_Click()

OneEmployee.Enabled = True
oneEmp_Code.Enabled = True

End Sub

Private Sub opt_Project_Click()
    OneEmployee.BoundText = ""
    oneEmp_Code.Text = ""
End Sub

Private Sub Project_Change()

LoadEmployee
OneEmployee.BoundText = ""
oneEmp_Code.Text = ""

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…"
            Else
                Me.Caption = "School  Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(9).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            ID.locked = True
      

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            pnlHeader.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
            pnlHeader.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… «·ÕÃ“ (  ⁄œÌ· )"
            Else
                Me.Caption = "Booking Request Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            ID.locked = True
           pnlHeader.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    
        
MySQL = MySQL & "                SELECT dbo.TblEmpEvaluation.ID, dbo.TblEmpEvaluation.SDate, dbo.TblBranchesData.branch_name, dbo.TblEmpEvaluation.EmployeeID,"
MySQL = MySQL & "                dbo.TblEmployee.Emp_Code AS EvalEmpCode, dbo.TblEmployee.Emp_Name AS EvalEmpName, dbo.TblEmpEvaluation.YearID, dbo.TblEmpEvaluation.MonthID,"
MySQL = MySQL & "                TblEmployee_1.Emp_Code AS ECode, TblEmployee_1.Emp_Name AS EName, dbo.TblEmpEvaluation.Branch, B.branch_name AS BName, dbo.TblEmpEvaluation.Project,"
MySQL = MySQL & "                dbo.TblEmpEvaluation.Mangerial, dbo.TblSection.name AS SectionName, dbo.TblEvaluation_Details.Emp_ID AS GridEmployeeID,"
MySQL = MySQL & "                GridEmployee.Emp_Code AS GridEmployeeCode, GridEmployee.Emp_Name AS GridEmployeeName, dbo.TblEvaluation_Details.EvalTitle, dbo.TblEvaluation_Details.Remarks,"
MySQL = MySQL & "                dbo.TblEvaluation_Details.Final_Evaluation, dbo.TblEvaluation_Details.Manual_Degree, dbo.TblEvaluation_Details.sum_Degrees, dbo.TblEvaluation_Details.MaxDgree,"
MySQL = MySQL & "                dbo.TblEvaluation_Details.PreDegree, dbo.TblEvaluation_Details.StanderedID, dbo.TblEvaluation_Details.Curr_Dynamic,"
MySQL = MySQL & "                dbo.TblEvaluationStandered.EName AS StanderedName, dbo.TblEvaluationStandered.ENameE AS StanderedNameE, dbo.TblEmpEvaluation.opt_AllEmployees,"
MySQL = MySQL & "                dbo.TblEmpEvaluation.opt_OneEmployee , dbo.TblEmpEvaluation.opt_Branch, dbo.TblEmpEvaluation.opt_Project, dbo.TblEmpEvaluation.opt_Managerial"
MySQL = MySQL & "                , TblEmpEvaluation.YearTitle ,  TblEmpEvaluation.MonthTitle   "
MySQL = MySQL & "                FROM     dbo.TblEmployee AS GridEmployee INNER JOIN"
MySQL = MySQL & "                dbo.TblEvaluation_Details ON GridEmployee.Emp_ID = dbo.TblEvaluation_Details.Emp_ID INNER JOIN"
MySQL = MySQL & "                dbo.TblEvaluationStandered ON dbo.TblEvaluation_Details.StanderedID = dbo.TblEvaluationStandered.ID RIGHT OUTER JOIN"
MySQL = MySQL & "                dbo.TblSection RIGHT OUTER JOIN"
MySQL = MySQL & "                dbo.TblEmpEvaluation ON dbo.TblSection.Id = dbo.TblEmpEvaluation.Mangerial RIGHT OUTER JOIN"
MySQL = MySQL & "                dbo.TblEmployee ON dbo.TblEmpEvaluation.EmployeeID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                dbo.TblBranchesData AS B ON dbo.TblEmpEvaluation.Branch = B.branch_id RIGHT OUTER JOIN"
MySQL = MySQL & "                dbo.TblBranchesData ON dbo.TblEmpEvaluation.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                dbo.TblEmployee AS TblEmployee_1 ON dbo.TblEmpEvaluation.OneEmployee = TblEmployee_1.Emp_ID ON"
MySQL = MySQL & "                dbo.TblEvaluation_Details.HID = dbo.TblEmpEvaluation.ID LEFT OUTER JOIN"
MySQL = MySQL & "                dbo.projects ON dbo.TblEmpEvaluation.Project = dbo.projects.id"
        
MySQL = MySQL & "        where   TblEmpEvaluation.ID = " & val(ID.Text)
                  
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EvaluationEmployee.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EvaluationEmployee.rpt"
    End If
    
 Dim mm As String

       
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's No data"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault



 
End Function
Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
     If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
   
   chrt_emp.ColumnCount = 0
   
   
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    BranchID.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    
    EmployeeID.BoundText = IIf(IsNull(rs("EmployeeID").value), "", Trim(rs("EmployeeID").value))
    YearID.ListIndex = IIf(IsNull(rs("YearID").value), -1, Trim(rs("YearID").value))
    MonthID.ListIndex = IIf(IsNull(rs("MonthID").value), -1, Trim(rs("MonthID").value))
    
    
    

    opt_Branch.value = IIf(IsNull(rs("opt_Branch").value), 0, rs("opt_Branch").value)
    opt_Project.value = IIf(IsNull(rs("opt_Project").value), 0, rs("opt_Project").value)
    opt_Managerial.value = IIf(IsNull(rs("opt_Managerial").value), 0, rs("opt_Managerial").value)
    opt_AllEmployees.value = IIf(IsNull(rs("opt_AllEmployees").value), False, CBool(rs("opt_AllEmployees").value))
    opt_OneEmployee.value = IIf(IsNull(rs("opt_OneEmployee").value), False, CBool(rs("opt_OneEmployee").value))
     
    Branch.BoundText = IIf(IsNull(rs("Branch").value), "", Trim(rs("Branch").value))
    Project.BoundText = IIf(IsNull(rs("Project").value), "", Trim(rs("Project").value))
    Mangerial.BoundText = IIf(IsNull(rs("Mangerial").value), "", Trim(rs("Mangerial").value))
    OneEmployee.BoundText = IIf(IsNull(rs("OneEmployee").value), "", Trim(rs("OneEmployee").value))
    Grid.Rows = Grid.FixedRows
    
    Set Rs_Temp = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = StrSQL & "  SELECT TblEmployee.Emp_ID ,  TblEmployee.FullCode ,  dbo.TblEvaluationStandered.EName, dbo.TblEvaluationStandered.ENameE, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEvaluation_Details.ID,"
    StrSQL = StrSQL & "  dbo.TblEvaluation_Details.HID, dbo.TblEvaluation_Details.StanderedID, dbo.TblEvaluation_Details.Emp_ID, dbo.TblEvaluation_Details.PreDegree,"
    StrSQL = StrSQL & "  dbo.TblEvaluation_Details.MaxDgree, dbo.TblEvaluation_Details.Curr_Dynamic, dbo.TblEvaluation_Details.sum_Degrees, dbo.TblEvaluation_Details.Manual_Degree,"
    StrSQL = StrSQL & "  dbo.TblEvaluation_Details.Final_Evaluation , dbo.TblEvaluation_Details.Remarks  , TblEvaluation_Details.EvalTitle"
    StrSQL = StrSQL & "  FROM     dbo.TblEvaluation_Details INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblEmpEvaluation ON dbo.TblEvaluation_Details.HID = dbo.TblEmpEvaluation.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblEvaluation_Details.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblEvaluationStandered ON dbo.TblEvaluation_Details.StanderedID = dbo.TblEvaluationStandered.ID"
    
    StrSQL = StrSQL & "  where TblEmpEvaluation.ID = " & val(ID.Text)
    
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
     Rs_Temp.MoveFirst
     With Grid
        .Rows = Rs_Temp.RecordCount + 1
        Dim j As Integer
        For j = 1 To .Rows - 1
                
                .TextMatrix(j, .ColIndex("id")) = IIf(IsNull(Rs_Temp("id").value), "", Rs_Temp("id").value)
                .TextMatrix(j, .ColIndex("Emp_ID")) = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
                .TextMatrix(j, .ColIndex("FullCode")) = IIf(IsNull(Rs_Temp("FullCode").value), 0, Rs_Temp("FullCode").value)
                .TextMatrix(j, .ColIndex("Emp_Name")) = IIf(IsNull(Rs_Temp("Emp_Name").value), 0, Rs_Temp("Emp_Name").value)
                .TextMatrix(j, .ColIndex("StanderedName")) = IIf(IsNull(Rs_Temp("EName").value), "", Rs_Temp("EName").value)
                .TextMatrix(j, .ColIndex("PreDegree")) = IIf(IsNull(Rs_Temp("PreDegree").value), "", Rs_Temp("PreDegree").value)
                .TextMatrix(j, .ColIndex("Curr_Dynamic")) = IIf(IsNull(Rs_Temp("Curr_Dynamic").value), "", Rs_Temp("Curr_Dynamic").value)
                .TextMatrix(j, .ColIndex("MaxDgree")) = IIf(IsNull(Rs_Temp("MaxDgree").value), "", Rs_Temp("MaxDgree").value)
                .TextMatrix(j, .ColIndex("sum_Degrees")) = IIf(IsNull(Rs_Temp("sum_Degrees").value), "", Rs_Temp("sum_Degrees").value)
                .TextMatrix(j, .ColIndex("Manual_Degree")) = IIf(IsNull(Rs_Temp("Manual_Degree").value), "", Rs_Temp("Manual_Degree").value)
                
                .TextMatrix(j, .ColIndex("Final_Evaluation")) = IIf(IsNull(Rs_Temp("EvalTitle").value), "", Rs_Temp("EvalTitle").value)
                
                Rs_Temp.MoveNext
         Next
        End With
        
        Grid.Row = 1
        Charting (val(Grid.TextMatrix(1, Grid.ColIndex("Emp_ID"))))
    End If
    
    

    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub




Private Sub TxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub


Private Sub txtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub



Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Grid.Rows = Grid.FixedRows
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
        Grid.Rows = Grid.FixedRows
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
 

End Function
 
Private Sub SaveData()
 Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If Trim(BranchID.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Managerial Area"
            Else
                Msg = "Õœœ «·›—⁄ «Ê·« "
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            BranchID.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
           If Trim(EmployeeID.BoundText) = "" Then
            
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Employee Evalution"
            Else
                Msg = "Õœœ «·ﬁ«∆„ »«· ﬁÌÌ„ "
            End If
            
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            EmployeeID.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
            End If
            
            
           If Grid.Rows <= Grid.FixedRows Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ «œŒ· ⁄‰«’— «· ﬁÌÌ„ «Ê·« "
            Else
                Msg = "Please evaluation elements first"
            End If
                 MsgBox (Msg)
                 Exit Sub
           End If
    
    
   ' If Trim(EName.text) = "" Then
   '         If SystemOptions.UserInterface = EnglishInterface Then
   '             Msg = "Specify Standered Name "
   '         Else
   '             Msg = "«œŒ· «”„ «·„⁄Ì«— «Ê·« "
   '         End If
''
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            EName.SetFocus
'   '         SendKeys "{F4}"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
    
        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text

           Case "N"
                rs.AddNew
                ID.Text = CStr(new_id("TblEmpEvaluation", "ID", "", True))
           Case "E"
                StrSQL = "delete From TblEvaluation_Details where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
           End Select

       
        

       
        rs("ID").value = val(ID.Text)
        rs("SDate").value = SDate.value
        rs("BranchID").value = IIf(BranchID.BoundText = "", Null, BranchID.BoundText)
        rs("EmployeeID").value = IIf(EmployeeID.BoundText = "", Null, EmployeeID.BoundText)
        rs("YearID").value = IIf(YearID.ListIndex = -1, Null, YearID.ListIndex)
        rs("MonthID").value = IIf(MonthID.ListIndex = -1, Null, MonthID.ListIndex)
        rs("opt_AllEmployees").value = opt_AllEmployees.value
        rs("opt_OneEmployee").value = opt_OneEmployee.value
        rs("opt_Branch").value = opt_Branch.value
        rs("opt_Project").value = opt_Project.value
        rs("opt_Managerial").value = opt_Managerial.value
        
        rs("OneEmployee").value = IIf(OneEmployee.BoundText = "", Null, OneEmployee.BoundText)
        rs("Project").value = IIf(Project.BoundText = "", Null, Project.BoundText)
        rs("Branch").value = IIf(Branch.BoundText = "", Null, Branch.BoundText)
        rs("Mangerial").value = IIf(Mangerial.BoundText = "", Null, Mangerial.BoundText)
        
        rs("CreationUserID").value = user_id
        rs("CreationDate").value = Date
        
        rs("YearTitle").value = YearID.Text
        rs("MonthTitle").value = MonthID.Text
        
        
        rs.update
        
        
        Dim ss As String
        Dim Rs_Temp As ADODB.Recordset
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = " select * from TblEvaluation_Details  where 1 = -1 "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Grid
        Dim j As Integer
        For j = 1 To Grid.Rows - 1
           If .TextMatrix(j, .ColIndex("Emp_ID")) <> "" Then
                    Rs_Temp.AddNew
                    Rs_Temp("ID") = CStr(new_id("TblEvaluation_Details", "ID", "", True))
                    Rs_Temp("HID") = val(ID.Text)
                    Rs_Temp("StanderedID") = .TextMatrix(j, .ColIndex("StanderedID"))
                    Rs_Temp("Emp_ID") = val(.TextMatrix(j, .ColIndex("Emp_ID")))
                    Rs_Temp("PreDegree") = val(.TextMatrix(j, .ColIndex("PreDegree")))
                    Rs_Temp("Curr_Dynamic") = val(.TextMatrix(j, .ColIndex("Curr_Dynamic")))
                    Rs_Temp("MaxDgree") = val(.TextMatrix(j, .ColIndex("MaxDgree")))
                    Rs_Temp("sum_Degrees") = val(.TextMatrix(j, .ColIndex("sum_Degrees")))
                    Rs_Temp("Manual_Degree") = val(.TextMatrix(j, .ColIndex("Manual_Degree")))

                    
                    Rs_Temp("EvalTitle") = .TextMatrix(j, .ColIndex("Final_Evaluation"))
                     
                    
                    Set Rs_Temp2 = New ADODB.Recordset
                    StrSQL = " select * from TblEvaluation_Employee   where 1 = -1 "
                    Rs_Temp2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    
                                    Rs_Temp2.AddNew
                                    Rs_Temp2("ID") = CStr(new_id("TblEvaluation_Employee", "ID", "", True))
                                    Rs_Temp2("HID") = val(ID.Text)
                                    Rs_Temp2("StanderedID") = .TextMatrix(j, .ColIndex("StanderedID"))
                                    Rs_Temp2("Emp_ID") = val(.TextMatrix(j, .ColIndex("Emp_ID")))
                                    Rs_Temp2("Eval_Degree") = val(.TextMatrix(j, .ColIndex("Final_Evaluation")))
                                    Rs_Temp2("YearNo") = val(YearID.ListIndex)
                                    Rs_Temp2("MonthNo") = val(MonthID.ListIndex)
                                    Rs_Temp2("EvalTitle") = .TextMatrix(j, .ColIndex("Final_Evaluation"))
                                    Rs_Temp2.update
                                    
                    Rs_Temp.update
                 End If
           Next
        End With
         
        
        
    
        Dim StrDes As String

     

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be daved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)
Grid.Rows = Grid.FixedRows
        Case "E"
            rs.find " ID='" & val(ID.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Action()
  
        Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If ID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…  —ﬁ„ " & CHR(13)
        Msg = Msg + (ID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Delete Booking Request File ? " & CHR(13)
        Msg = Msg + (ID.Text) & CHR(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
        
            If Not rs.RecordCount < 1 Then
                                
                 StrSQL = "delete From TblEvaluation_Details where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                           
                           StrSQL = "delete From TblEvaluation_Employee where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                 
                           
                StrSQL = "delete From TblEmpEvaluation where  ID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From TblEmpEvaluation "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   
                   
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                    Grid.Rows = Grid.FixedRows
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Grid.Rows = Grid.FixedRows
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
         Msg = "this process Not Aailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ‰ﬁ«ÿ «· ﬁÌÌ„ «·ÌœÊÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub


