VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmSearchSerial 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15285
   HelpContextID   =   10
   Icon            =   "FrmSearchSerial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   15285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   8430
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15285
      _cx             =   26961
      _cy             =   14870
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   1
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmSearchSerial.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   720
         Index           =   2
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7695
         Width           =   15255
         _cx             =   26908
         _cy             =   1270
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
         BorderWidth     =   1
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   390
            Left            =   8025
            TabIndex        =   53
            Top             =   150
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œ…"
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   0
            Left            =   12405
            TabIndex        =   54
            Top             =   150
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   688
            ButtonStyle     =   1
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   1
            Left            =   10170
            TabIndex        =   55
            Top             =   150
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   688
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   2
            Left            =   285
            TabIndex        =   56
            Top             =   150
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   7
            Left            =   3300
            TabIndex        =   63
            Top             =   120
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   953
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   7080
         Index           =   1
         Left            =   15
         TabIndex        =   2
         Top             =   600
         Width           =   15255
         _cx             =   26908
         _cy             =   12488
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›|≈” ⁄·«„ ⁄‰ ﬂ„Ì… „Ã„Ê⁄…|’‰› „Ã„⁄|«” ⁄·«„ ⁄‰  «—ÌŒ ’·«ÕÌ… «’‰«›|«” ⁄·«„ ⁄‰ «·«’‰«› «· Ì »·€  Õœ «·ÿ·»"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6705
            Index           =   5
            Left            =   16200
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   15165
            _cx             =   26749
            _cy             =   11827
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
            Begin MSDataListLib.DataCombo DcboStores 
               Height          =   315
               Left            =   150
               TabIndex        =   61
               Top             =   1035
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈” Œœ„ «·ﬂ„Ì«  «·„ «Õ… ›Ï «·„Œ“‰ «·„Õœœ"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   6570
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   1005
               Width           =   8445
            End
            Begin VB.TextBox TxtAssbliedItemCode 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   6345
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   105
               Width           =   6480
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItemsA 
               Height          =   4755
               Left            =   75
               TabIndex        =   5
               Top             =   1545
               Width           =   15015
               _cx             =   26485
               _cy             =   8387
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSearchSerial.frx":040C
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
            Begin MSDataListLib.DataCombo DcboAssbliedItems 
               Height          =   315
               Left            =   1590
               TabIndex        =   6
               Top             =   615
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdItemSearch 
               Height          =   390
               Index           =   1
               Left            =   150
               TabIndex        =   7
               Top             =   600
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "..."
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
               ButtonImage     =   "FrmSearchSerial.frx":0544
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õœœ «·„Œ“‰"
               Height          =   360
               Index           =   28
               Left            =   4605
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   1080
               Width           =   1740
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   270
               Index           =   27
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   6345
               Width           =   2340
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì… «· Ï Ì„ﬂ‰ﬂ  Ã„Ì⁄Â« „‰ «·’‰›:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   270
               Index           =   26
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   6345
               Width           =   6870
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   435
               Index           =   25
               Left            =   12300
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   105
               Width           =   2715
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·’‰›"
               Height          =   435
               Index           =   24
               Left            =   12300
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   540
               Width           =   2715
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6705
            Index           =   4
            Left            =   15900
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   330
            Width           =   15165
            _cx             =   26749
            _cy             =   11827
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
            Begin VB.TextBox TxtGroupCode 
               Alignment       =   1  'Right Justify
               Height          =   465
               Left            =   3315
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   30
               Width           =   8910
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   5310
               Left            =   75
               TabIndex        =   12
               Top             =   1050
               Width           =   15015
               _cx             =   26485
               _cy             =   9366
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSearchSerial.frx":0ADE
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
            Begin MSDataListLib.DataCombo DcboGroupID 
               Height          =   315
               Left            =   675
               TabIndex        =   13
               Top             =   555
               Width           =   11550
               _ExtentX        =   20373
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Ã„«·Ï «·ﬂ„Ì« : "
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
               Height          =   285
               Index           =   23
               Left            =   3855
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   6375
               Width           =   3165
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   22
               Left            =   75
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   6435
               Width           =   3540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
               Height          =   510
               Index           =   21
               Left            =   11925
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   75
               Width           =   3015
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄…"
               Height          =   390
               Index           =   20
               Left            =   11925
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   540
               Width           =   3015
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6705
            Index           =   3
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   15165
            _cx             =   26749
            _cy             =   11827
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
            Begin VB.TextBox TxtStoreID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   11010
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox XPTxtCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6180
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   60
               Width           =   6645
            End
            Begin C1SizerLibCtl.C1Tab TabMain 
               Height          =   5205
               Index           =   0
               Left            =   75
               TabIndex        =   20
               Top             =   1290
               Width           =   14940
               _cx             =   26352
               _cy             =   9181
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
               Appearance      =   2
               MousePointer    =   0
               Version         =   801
               BackColor       =   12648447
               ForeColor       =   -2147483630
               FrontTabColor   =   14871017
               BackTabColor    =   12648447
               TabOutlineColor =   -2147483632
               FrontTabForeColor=   -2147483630
               Caption         =   "«·ﬂ„Ì«  Ê«·„Œ“Ê‰|«·√”⁄«—|„·«ÕŸ« "
               Align           =   0
               CurrTab         =   0
               FirstTab        =   0
               Style           =   3
               Position        =   1
               AutoSwitch      =   -1  'True
               AutoScroll      =   -1  'True
               TabPreview      =   -1  'True
               ShowFocusRect   =   -1  'True
               TabsPerPage     =   0
               BorderWidth     =   0
               BoldCurrent     =   0   'False
               DogEars         =   -1  'True
               MultiRow        =   0   'False
               MultiRowOffset  =   200
               CaptionStyle    =   0
               TabHeight       =   0
               TabCaptionPos   =   4
               TabPicturePos   =   0
               CaptionEmpty    =   ""
               Separators      =   0   'False
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   37
               Picture(0)      =   "FrmSearchSerial.frx":0BCD
               Picture(1)      =   "FrmSearchSerial.frx":0F67
               Picture(2)      =   "FrmSearchSerial.frx":1301
               Flags(2)        =   2
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   4740
                  Index           =   6
                  Left            =   15885
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   14850
                  _cx             =   26194
                  _cy             =   8361
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
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   4740
                  Index           =   7
                  Left            =   15585
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   14850
                  _cx             =   26194
                  _cy             =   8361
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
                  Begin VB.Frame Fra 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·”⁄— ›Ï «Œ— ›« Ê—… »Ì⁄"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000080&
                     Height          =   1425
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   30
                     Width           =   2865
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "—ﬁ„ «·›« Ê—…:"
                        ForeColor       =   &H00000040&
                        Height          =   255
                        Index           =   12
                        Left            =   1620
                        RightToLeft     =   -1  'True
                        TabIndex        =   31
                        Top             =   300
                        Width           =   1185
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   " «—ÌŒ «·›« Ê—…:"
                        ForeColor       =   &H00000040&
                        Height          =   255
                        Index           =   13
                        Left            =   1620
                        RightToLeft     =   -1  'True
                        TabIndex        =   30
                        Top             =   570
                        Width           =   1185
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«”„ «·⁄„Ì·:"
                        ForeColor       =   &H00000040&
                        Height          =   255
                        Index           =   14
                        Left            =   1620
                        RightToLeft     =   -1  'True
                        TabIndex        =   29
                        Top             =   840
                        Width           =   1185
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "”⁄— «·’‰›:"
                        ForeColor       =   &H00000040&
                        Height          =   255
                        Index           =   15
                        Left            =   1620
                        RightToLeft     =   -1  'True
                        TabIndex        =   28
                        Top             =   1110
                        Width           =   1185
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   16
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   27
                        Top             =   300
                        Width           =   1545
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   17
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   26
                        Top             =   570
                        Width           =   1545
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   18
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   25
                        Top             =   840
                        Width           =   1545
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   19
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   24
                        Top             =   1110
                        Width           =   1545
                     End
                  End
                  Begin ImpulseButton.ISButton ISButton1 
                     Height          =   405
                     Index           =   0
                     Left            =   270
                     TabIndex        =   32
                     Top             =   1590
                     Visible         =   0   'False
                     Width           =   2565
                     _ExtentX        =   4524
                     _ExtentY        =   714
                     Caption         =   "«”⁄«— ›Ê« Ì— «·»Ì⁄"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin ImpulseButton.ISButton ISButton2 
                     Height          =   405
                     Index           =   0
                     Left            =   270
                     TabIndex        =   33
                     Top             =   2070
                     Visible         =   0   'False
                     Width           =   2565
                     _ExtentX        =   4524
                     _ExtentY        =   714
                     Caption         =   "«”⁄«— ›Ê« Ì— «·‘—«¡"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FgItemPriceList 
                     Height          =   1425
                     Left            =   2910
                     TabIndex        =   34
                     Top             =   1200
                     Width           =   2625
                     _cx             =   4630
                     _cy             =   2514
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
                     Rows            =   2
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmSearchSerial.frx":169B
                     ScrollTrack     =   0   'False
                     ScrollBars      =   2
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H008080FF&
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
                     Height          =   825
                     Index           =   0
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   4395
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(„” Â·ﬂ)"
                     Height          =   285
                     Index           =   3
                     Left            =   4110
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   60
                     Width           =   1395
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(⁄„Ì·)"
                     Height          =   285
                     Index           =   4
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   360
                     Width           =   1395
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(œÌ·—)"
                     Height          =   285
                     Index           =   7
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   660
                     Width           =   1395
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   8
                     Left            =   3420
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   60
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   9
                     Left            =   3390
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   360
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   10
                     Left            =   3390
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   660
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·”⁄— „‰ Œ·«· ﬁ«∆„… «·√”⁄«—"
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
                     Height          =   225
                     Index           =   11
                     Left            =   3450
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   960
                     Width           =   2055
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   4740
                  Index           =   8
                  Left            =   45
                  TabIndex        =   42
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   14850
                  _cx             =   26194
                  _cy             =   8361
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
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   2415
                     Left            =   30
                     TabIndex        =   43
                     Top             =   30
                     Width           =   14745
                     _cx             =   26009
                     _cy             =   4260
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
                     Rows            =   15
                     Cols            =   11
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmSearchSerial.frx":1720
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
                  Begin VSFlex8UCtl.VSFlexGrid FgSum 
                     Height          =   1335
                     Left            =   30
                     TabIndex        =   44
                     Top             =   2460
                     Width           =   14745
                     _cx             =   26009
                     _cy             =   2355
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
                     Rows            =   15
                     Cols            =   3
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmSearchSerial.frx":18AF
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
               End
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   1440
               TabIndex        =   45
               Top             =   405
               Width           =   11385
               _ExtentX        =   20082
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdItemSearch 
               Height          =   390
               Index           =   0
               Left            =   75
               TabIndex        =   46
               Top             =   390
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "..."
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
               ButtonImage     =   "FrmSearchSerial.frx":1937
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseAniLabel.ISAniLabel LblLink 
               Height          =   270
               Left            =   75
               TabIndex        =   47
               Top             =   105
               Width           =   6030
               _ExtentX        =   10636
               _ExtentY        =   476
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "FrmSearchSerial.frx":1ED1
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   "⁄—÷ ‘«‘…  ﬁ«—Ì— «·’‰›"
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin MSDataListLib.DataCombo DCboStoreName 
               Height          =   315
               Left            =   4305
               TabIndex        =   65
               Top             =   720
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Œ“‰  "
               Height          =   285
               Index           =   29
               Left            =   12975
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   720
               Width           =   1740
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·’‰›"
               Height          =   405
               Index           =   6
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   435
               Width           =   2340
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   465
               Index           =   5
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   60
               Width           =   2340
            End
            Begin VB.Label LblHaveSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Â–« «·’‰› ·Â ”Ì—Ì«·"
               ForeColor       =   &H000000FF&
               Height          =   345
               Left            =   -1890
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   960
               Width           =   5745
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
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
               Height          =   285
               Index           =   1
               Left            =   225
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   6540
               Width           =   11625
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Index           =   2
               Left            =   5955
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1080
               Width           =   4905
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   6705
            Left            =   16500
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   330
            Width           =   15165
            _cx             =   26749
            _cy             =   11827
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
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6030
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   600
               Width           =   6645
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   3240
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   120
               Width           =   1440
            End
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   255
               Left            =   9735
               TabIndex        =   68
               Top             =   120
               Width           =   5130
               _Version        =   786432
               _ExtentX        =   9049
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ŸÂ«— «·«’‰«› «· Ì ·Ì” ·Â«  «—ÌŒ «‰ Â«¡ "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   4995
               Left            =   150
               TabIndex        =   70
               Top             =   1515
               Width           =   14865
               _cx             =   26220
               _cy             =   8811
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
               Cols            =   13
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSearchSerial.frx":2033
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
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   2490
               TabIndex        =   82
               Top             =   945
               Width           =   10185
               _ExtentX        =   17965
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Â–« «·’‰› ·Â ”Ì—Ì«·"
               ForeColor       =   &H000000FF&
               Height          =   345
               Left            =   -2340
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   960
               Width           =   4680
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   225
               Index           =   35
               Left            =   13050
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   600
               Width           =   1740
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·’‰›"
               Height          =   285
               Index           =   30
               Left            =   12975
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   975
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   330
               Index           =   34
               Left            =   1515
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   345
               Width           =   2865
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   33
               Left            =   1515
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   30
               Width           =   2865
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   330
               Index           =   32
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   120
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁ»·  «—ÌŒ «·«‰ Â«¡ »"
               Height          =   285
               Index           =   31
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   120
               Width           =   1890
            End
         End
         Begin C1SizerLibCtl.C1Elastic EleMain 
            Height          =   6705
            Left            =   16800
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   330
            Width           =   15165
            _cx             =   26749
            _cy             =   11827
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   5100
               Left            =   150
               TabIndex        =   87
               Top             =   840
               Width           =   14940
               _cx             =   26352
               _cy             =   8996
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
               BackColorBkg    =   16777215
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSearchSerial.frx":226F
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
            Begin MSDataListLib.DataCombo DcbStore 
               Bindings        =   "FrmSearchSerial.frx":243E
               Height          =   315
               Left            =   7845
               TabIndex        =   88
               Top             =   360
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ListField       =   "account_name"
               BoundColumn     =   "code"
               Text            =   ""
               RightToLeft     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   585
               Index           =   9
               Left            =   150
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   6000
               Width           =   14865
               _cx             =   26220
               _cy             =   1032
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·√’‰«›"
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Index           =   38
                  Left            =   2475
                  TabIndex        =   92
                  Top             =   150
                  Width           =   345
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Index           =   37
                  Left            =   2205
                  TabIndex        =   91
                  Top             =   150
                  Width           =   255
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Œ“‰"
               Height          =   270
               Index           =   36
               Left            =   13275
               TabIndex        =   89
               Top             =   360
               Width           =   840
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   570
         Index           =   1
         Left            =   15
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   15
         Width           =   15255
         _cx             =   26908
         _cy             =   1005
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
         BorderWidth     =   1
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
            Caption         =   "«” ⁄·«„ ⁄‰  «—ÌŒ ’·«ÕÌ… «’‰«›"
            Height          =   195
            Index           =   3
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   360
            Width           =   4170
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«” ⁄·«„ ⁄‰ «·«’‰«› «· Ì »·€  Õœ «·ÿ·»"
            Height          =   195
            Index           =   4
            Left            =   2220
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   120
            Width           =   4200
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰› „Ã„⁄"
            Height          =   255
            Index           =   2
            Left            =   6915
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   45
            Width           =   3855
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… „Ã„Ê⁄…"
            Height          =   255
            Index           =   1
            Left            =   10935
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   300
            Width           =   4260
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
            Height          =   255
            Index           =   0
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   30
            Value           =   -1  'True
            Width           =   4395
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   90
            Visible         =   0   'False
            Width           =   2160
         End
      End
   End
End
Attribute VB_Name = "FrmSearchSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rs As New ADODB.Recordset
Dim cSearchDcbo(3) As clsDCboSearch
Dim m_LngGridRow As Long
Dim FirstPeriodDateInthisYear  As Date

Private Sub CheckBox1_Click()
Cmd_Click (0)
End Sub

Private Sub Chk_Click()
    Me.lbl(28).Enabled = CBool(Me.Chk.value)
    Me.DcboStores.Enabled = CBool(Me.Chk.value)
End Sub

Public Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    Dim StSQL As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
    
            If rs.State = adStateOpen Then
                rs.Close
            End If

            If Opt(0).value = True Then
                If DCboItemsName.BoundText = "" Then
                    If Trim(XPTxtCode.Text) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "»—Ã«¡  ÕœÌœ «·’‰›..!!"
                            '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        ElseIf SystemOptions.UserInterface = EnglishInterface Then
                            Msg = "Please choose an Item Name....!"
                            '    MsgBox Msg, vbOKOnly + vbExclamation, App.Title
                        End If

                        DCboItemsName.SetFocus
                        SendKeys "{F4}"
                        'Exit Sub
                    End If
                End If

                rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount < 1 Then
                    'Exit Sub
                End If

                Retrive
            ElseIf Opt(1).value = True Then

                If val(Me.DcboGroupID.BoundText) = 0 Then
                    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·„Ã„Ê⁄…....!!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount < 1 Then
            
                End If

                RetriveGroup
            ElseIf Opt(2).value = True Then

                If val(Me.DcboAssbliedItems.BoundText) = 0 Then
                    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·’‰› «·„Ã„⁄....!!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                If Me.Chk.value = vbChecked Then
                    If Me.DcboStores.BoundText = "" Then
                        Msg = "ÌÃ»  ÕœÌœ «”„ «·„Œ“‰...!!!"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub
                    End If
                End If

                rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                RetriveAssbliedItem
                
            ElseIf Opt(3).value = True Then
                VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                VSFlexGrid1.Rows = 2
                FillItemExperDate
                
            ElseIf Opt(4).value = True Then
                FillGrid
            End If
            
        Case 1
            clear_all Me
            ClearData
            Opt(0).value = True

        Case 2
            Unload Me

        Case 7
            If Opt(3).value = True Then
                print_report
            ElseIf Opt(4).value = True Then
                print_reportReq
            Else
                printing
            End If
            
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Function printing()
 
    Dim VReport As ClsGardReport
 
    Set VReport = New ClsGardReport
 
    VReport.ShowGardData2 Build_Sql, FirstPeriodDateInthisYear, Date, DCboItemsName.Text
    
End Function

Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
        sql = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS Qty, "
        sql = sql & "               dbo.Transaction_Details.Item_ID , dbo.TblItems.itemcode, dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode"
        sql = sql & "    FROM         dbo.Transaction_Details INNER JOIN"
        sql = sql & "              dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        sql = sql & "              dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
        sql = sql & "              dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
       If DataCombo1.BoundText <> "" Then
            If CheckBox1.value = vbChecked Then
                My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " or dbo.Transaction_Details.ExpiryDate is null) And ( Item_ID = " & DataCombo1.BoundText & ") "
            Else
                My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " ) And ( Item_ID = " & DataCombo1.BoundText & ") "
            End If
        Else
            If CheckBox1.value = vbChecked Then
                My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " or dbo.Transaction_Details.ExpiryDate is null) "
            Else
                My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " ) "
            End If
        End If
        sql = sql & "    GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
        sql = sql & "               dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode"
        sql = sql & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) > 0)"
        sql = sql & " ORDER BY dbo.Transaction_Details.ExpiryDate"
    
 
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmExperDateOfItemsSerial.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmExperDateOfItemsSerialE.rpt"
        End If

   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "No Date"
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
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
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
ErrTrap:
  End Function


Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdItemSearch_Click(Index As Integer)

    Select Case Index

        Case 0
            PutFormOnTop Me.hwnd, False
            ModOpenScreen.ShowDialogItemsSearch Me.DCboItemsName
            PutFormOnTop Me.hwnd, True

        Case 1
            PutFormOnTop Me.hwnd, False
            ModOpenScreen.ShowDialogItemsSearch Me.DcboAssbliedItems
            PutFormOnTop Me.hwnd, True
    End Select

End Sub

Private Sub DataCombo1_Click(Area As Integer)
DataCombo1_Change
End Sub

Private Sub DcboAssbliedItems_Change()
    Dim StrItemCode As String

    If Me.DcboAssbliedItems.BoundText = "" Then
        Me.TxtAssbliedItemCode.Text = ""
        Exit Sub
    Else
        StrItemCode = GetItemCode(val(Me.DcboAssbliedItems.BoundText))

        If StrItemCode <> Trim$(Me.TxtAssbliedItemCode.Text) Then
            TxtAssbliedItemCode.Text = StrItemCode
        End If
    End If

End Sub

Private Sub DcboAssbliedItems_Click(Area As Integer)
DcboAssbliedItems_Change
End Sub

Private Sub DCboItemsName_Change()
    Dim StrItemCode As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Me.DCboItemsName.BoundText = "" Then
        Me.XPTxtCode.Text = ""
        ClearData
    Else
        Me.LblHaveSerial.Visible = True
        StrSQL = "Select * From TblItems Where ItemID=" & val(Me.DCboItemsName.BoundText) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrItemCode = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        End If

        If StrItemCode <> Trim(Me.XPTxtCode.Text) Then
            Me.XPTxtCode.Text = StrItemCode
        End If

        rs.Close
        Set rs = Nothing
    End If

End Sub
Private Sub DataCombo1_Change()
    Dim StrItemCode As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Me.DataCombo1.BoundText = "" Then
        Me.Text2.Text = ""
        ClearData
    Else
        Me.Label1.Visible = True
        StrSQL = "Select * From TblItems Where ItemID =" & Me.DataCombo1.BoundText & " "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrItemCode = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        End If

        If StrItemCode <> Trim(Me.Text2.Text) Then
            Me.Text2.Text = StrItemCode
        End If

        rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub DCboItemsName_KeyDown(KeyCode As Integer, _
                                  Shift As Integer)
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.Tag = "" Then
            GetItemData val(Me.DCboItemsName.BoundText), ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, _
                                  Shift As Integer)
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.Tag = "" Then
            GetItemData val(Me.DataCombo1.BoundText), ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 1
        Set FrmItemSearch.DcboItems = Me.DCboItemsName
        FrmItemSearch.show vbModal
    End If

End Sub
Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 1
        Set FrmItemSearch.DcboItems = Me.DataCombo1
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub Fg_DblClick()
    Dim LngRow As Long
    On Error GoTo ErrTrap

    If fg.Row < fg.FixedRows Then Exit Sub
    If fg.TextMatrix(fg.Row, fg.ColIndex("Serial")) = "·«ÌÊÃœ" Then Exit Sub
    If mdifrmmain.ActiveForm Is Nothing Then Exit Sub

    LngRow = Me.LngGridRow

    If Txt.Text = "Search" Then
        If Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("Serial")) <> "" Then
            mdifrmmain.ActiveForm.fg.TextMatrix(LngRow, mdifrmmain.ActiveForm.fg.ColIndex("Serial")) = fg.TextMatrix(fg.Row, fg.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Retutn" Then

        If Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("Serial")) <> "" Then
            FrmReturnpurchases.fg.TextMatrix(FrmReturnpurchases.fg.Row, FrmReturnpurchases.fg.ColIndex("Serial")) = fg.TextMatrix(fg.Row, fg.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Moving" Then

        If Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("Serial")) <> "" Then
            FrmMoving.fg.TextMatrix(FrmMoving.fg.Row, FrmMoving.fg.ColIndex("Serial")) = fg.TextMatrix(fg.Row, fg.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Destruction" Then

        If Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("Serial")) <> "" Then
            FrmDestruction.fg.TextMatrix(FrmDestruction.fg.Row, FrmDestruction.fg.ColIndex("Serial")) = fg.TextMatrix(fg.Row, fg.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Replace" Then

        If Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("Serial")) <> "" Then
            FrmReplace.TxtNewSerial.Text = fg.TextMatrix(fg.Row, fg.ColIndex("Serial"))
            FrmReplace.DCboStoreName.BoundText = fg.TextMatrix(fg.Row, fg.ColIndex("StoreName"))
            Unload Me
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, _
                       Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Fg_DblClick
    End If

End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyF3 Then
        CmdItemSearch_Click 0
    End If

    Exit Sub
ErrTrap:
End Sub


Public Sub FillItemExperDate()
Dim My_SQL As String
Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        My_SQL = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS Qty, "
        My_SQL = My_SQL & "               dbo.Transaction_Details.Item_ID , dbo.TblItems.itemcode, dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode"
        My_SQL = My_SQL & "    FROM         dbo.Transaction_Details INNER JOIN"
        My_SQL = My_SQL & "              dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        My_SQL = My_SQL & "              dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
        My_SQL = My_SQL & "              dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
      If CheckBox1.value = vbChecked Then
        My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " or dbo.Transaction_Details.ExpiryDate is null) "
        Else
        My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " ) "
        End If
       If val(DataCombo1.BoundText) Then
       My_SQL = My_SQL & " and dbo.Transaction_Details.Item_ID=" & val(DataCombo1.BoundText) & ""
       End If
        My_SQL = My_SQL & "    GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
        My_SQL = My_SQL & "               dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode"
        My_SQL = My_SQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) > 0)"
        My_SQL = My_SQL & " ORDER BY dbo.Transaction_Details.ExpiryDate"
        rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
   With VSFlexGrid1
   .Rows = .Rows + rs.RecordCount - 1
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Serial")) = i
                        .TextMatrix(i, .ColIndex("ExpiryDate")) = IIf(IsNull(rs.Fields("ExpiryDate").value), "", rs.Fields("ExpiryDate").value)
                        .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs.Fields("Qty").value), 0, rs.Fields("Qty").value)
                        .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)
                        If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value)
                        Else
                        .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value)
                        End If
                         rs.MoveNext
                    Next i
   End With
                    rs.Close
   End If
     
End Sub

Private Sub Text1_Change()
Cmd_Click (0)
End Sub

Private Sub Text2_Change()
    GetItemData 0, Trim(Me.Text2.Text)
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub


Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
 
     
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim RsTemp As New ADODB.Recordset
    Dim RsNote As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrList As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Dim Msg As String
    On Error GoTo ErrTrap
    Me.LblHaveSerial.Caption = ""

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        Msg = "Note:-"
        Msg = Msg & CHR(13) & "You can display a item cart report "
        Msg = Msg & "which show all the item transactions "
        Msg = Msg & "you can display this report From the Report Screen"
        lbl(0).Caption = Msg
        Msg = "Press F7 to show item search..."
        lbl(1).Caption = Msg
    Else
        Msg = "„·ÕÊŸ…:-"
        Msg = Msg & CHR(13) & "Ì„ﬂ‰ﬂ ⁄—÷  ﬁ—Ì— »ﬂ«—  «·’‰› «·–Ï Ì⁄—÷ ·ﬂ "
        Msg = Msg & "Ã„Ì⁄ «·Õ—ﬂ«  «·Œ«’… »«·’‰› „‰ Ê—«œ Ê’«œ— "
        Msg = Msg & "„‰ Œ·«· ‘«‘… «· ﬁ«—Ì— «·⁄«„… À„  ﬁ«—Ì— «·√’‰«› Ê√Œ —  ﬁ—Ì— ﬂ«—  «·’‰›"
        lbl(0).Caption = Msg
        Msg = "··„”«⁄œ… ›Ï «·⁄ÀÊ— ⁄·Ï «”„ «·’‰› ≈÷€ÿ F7"
        lbl(1).Caption = Msg
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    StrSQL = "select * From TblStore"
    RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With fg
        StrList = .BuildComboList(RsNote, "StoreName", "StoreID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("StoreName")) = "|" & StrList
        End If

    End With

    CenterForm Me
    fg.WallPaper = BG.SearchWallpaper
    FgItems.WallPaper = BG.SearchWallpaper
    FgItemsA.WallPaper = BG.SearchWallpaper
    FgItemsA.AutoSize 0, FgItemsA.Cols - 1, False
    Set FgItemPriceList.WallPaper = BG.Picture
    Set FgSum.WallPaper = BG.Picture
    FgItemPriceList.AutoSize 0, FgItemPriceList.Cols - 1, False

    FormPostion Me, GetPostion
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DCboItemsName, 0
    Dcombos.GetItemsNames Me.DataCombo1, 0
    Dcombos.GetStores Me.DcbStore
    Dcombos.GetItemSGroups Me.DcboGroupID
     Dcombos.GetStores Me.DCboStoreName
     
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DataCombo1
    'Set cSearchDcbo(4).Client = Me.DCboItemsName
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboGroupID
    Dcombos.GetItemsNames Me.DcboAssbliedItems, , 1
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DcboAssbliedItems
    Dcombos.GetStores Me.DcboStores
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboStores

    Opt(0).value = True
    Opt_Click 0
    Me.Chk.value = vbUnchecked
    Chk_Click
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion
    Me.Tag = ""

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo(0) = Nothing
    Set cSearchDcbo(1) = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String

    'On Error GoTo ErrTrap
    If Opt(0).value = True Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "select * From QryGardComplete"
            StrSQL = StrSQL + " where ItemCode='" & XPTxtCode.Text & "'"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
       
            StrSQL = "SELECT  Transaction_Details.lotno,   dbo.TblUnites.UnitId, SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.TblStore.StoreName, dbo.TblUnites.UnitName, "
            StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
            StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
            StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
            StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
            StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
            StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
            StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
            StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
            StrSQL = StrSQL + " and Item_ID =" & val(DCboItemsName.BoundText)
     If val(Me.DCboStoreName.BoundText) <> 0 Then
                    StrSQL = StrSQL + " AND Transactions.StoreID=" & val(Me.DCboStoreName.BoundText) & ""
     End If
            StrSQL = StrSQL & "  GROUP BY dbo.TblStore.StoreName, dbo.TblUnites.UnitName, dbo.TblUnites.UnitId, dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName,"
            StrSQL = StrSQL & "  dbo.TblItemsColors.ColorName,Transaction_Details.lotno"
            StrSQL = StrSQL & "  HAVING      (SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) <> 0)"
            ' StrSQL = "SELECT * From dbo.QryGardComplete(0)"
            ' StrSQL = StrSQL + " where ItemCode='" & XPTxtCode.text & "'"
            ' StrSQL = StrSQL + " Order By StoreName"
    
        End If

    ElseIf Opt(1).value = True Then

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "select * From QryGardComplete"
            StrSQL = StrSQL + " where GroupID=" & val(Me.DcboGroupID.BoundText) & ""
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT * From dbo.QryGardComplete(0)"
            StrSQL = StrSQL + " where GroupID=" & val(Me.DcboGroupID.BoundText) & ""
            StrSQL = StrSQL + " Order By ItemName"
        End If

    ElseIf Opt(2).value = True Then

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT dbo.TblItemsParts.TableID, dbo.TblItemsParts.ItemID," & "dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty," & "dbo.TblItemsParts.PartItemPrice,QryGARDShort.QTY, QryGARDShort.ItemCode," & "QryGARDShort.ItemName, QryGARDShort.StoreID, QryGARDShort.StoreName," & "QryGARDShort.GroupID"
            StrSQL = StrSQL + " FROM         dbo.TblItemsParts LEFT OUTER JOIN "
            StrSQL = StrSQL + " dbo.QryGARDShort() QryGARDShort ON " & "dbo.TblItemsParts.PartItemID = QryGARDShort.ItemID"
            StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & val(Me.DcboAssbliedItems.BoundText) & ""

            If Me.Chk.value = vbChecked Then
                If val(Me.DcboStores.BoundText) <> 0 Then
                    StrSQL = StrSQL + " AND QryGARDShort.StoreID=" & val(Me.DcboStores.BoundText) & ""
                End If
            End If

     If val(Me.DCboStoreName.BoundText) <> 0 Then
                    StrSQL = StrSQL + " AND QryGARDShort.StoreID=" & val(Me.DCboStoreName.BoundText) & ""
     End If
                
            StrSQL = StrSQL + " Order BY dbo.TblItemsParts.TableID"
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        
        End If
    End If

    Build_Sql = StrSQL
    Exit Function
ErrTrap:
End Function

Private Sub Retrive()
    Dim StrSQL As String
    Dim Num As Integer
    Dim RsData As ADODB.Recordset
    Dim RowNum As Long
    Dim ItemTransInfo As LastItemTransInfo
    Dim RsSumQty As ADODB.Recordset

    On Error GoTo ErrTrap
    fg.Clear flexClearScrollable, flexClearEverything
    FgSum.Clear flexClearScrollable, flexClearEverything

    GetItemData 0, Trim(Me.XPTxtCode.Text)

    If Not (rs.EOF Or rs.BOF) Then
        If True Then
            If False = True Then
                If Me.DCboItemsName.BoundText <> rs("ItemID").value Then
                    Me.DCboItemsName.BoundText = rs("ItemID").value
                End If

                LblHaveSerial.Visible = True
            Else
                LblHaveSerial.Visible = True
            End If
        End If

        fg.Rows = rs.RecordCount + 1
Dim LngUnitID As Long
        For Num = 1 To rs.RecordCount

            With fg
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num

                '    .TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(rs("ItemSerial").value), "·«ÌÊÃœ", (rs("ItemSerial").value))
                If Not (IsNull(rs("SUMQTY").value)) Then
                    .TextMatrix(Num, .ColIndex("Quantity")) = rs("SUMQTY").value
                Else
                    .TextMatrix(Num, .ColIndex("Quantity")) = 0
                End If
            
                .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", (rs("StoreName").value))
                .TextMatrix(Num, .ColIndex("ColorName")) = IIf(IsNull(rs("ColorName").value), "", (rs("ColorName").value))
                .TextMatrix(Num, .ColIndex("ItemSize")) = IIf(IsNull(rs("SizeName").value), "", (rs("SizeName").value))
                .TextMatrix(Num, .ColIndex("ClassName")) = IIf(IsNull(rs("ClassName").value), "", (rs("ClassName").value))
                .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(rs("UnitName").value), "", (rs("UnitName").value))
        .TextMatrix(Num, .ColIndex("lotNo")) = IIf(IsNull(rs("lotNo").value), "", (rs("lotNo").value))

                LngUnitID = IIf(IsNull(rs("UnitId").value), 0, (rs("UnitId").value))
                .TextMatrix(Num, .ColIndex("price")) = GetItemPrice(val(DCboItemsName.BoundText), 1, LngUnitID)
                
                
            
            End With

            rs.MoveNext
        Next Num

        fg.AutoSize 0, fg.Cols - 1, False

        '  Exit Sub
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(2).Caption = "≈Ã„«·Ï «·ﬂ„Ì«  «·„ÊÃÊœ… : " & fg.Aggregate(flexSTSum, fg.FixedRows, fg.ColIndex("Quantity"), fg.Rows - 1, fg.ColIndex("Quantity"))
        Else
            Me.lbl(2).Caption = "Total Item Stock: " & fg.Aggregate(flexSTSum, fg.FixedRows, fg.ColIndex("Quantity"), fg.Rows - 1, fg.ColIndex("Quantity"))
        End If
    
        Set RsSumQty = New ADODB.Recordset

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        
            'StrSQL = "Select Sum(Qty) as SumQty ,StoreName "
            'StrSQL = StrSQL + " From dbo.QryGardComplete(0) QryGardComplete"
            'StrSQL = StrSQL + " Where ItemCode='" & Trim(XPTxtCode.text) & "'"
            'If Me.DCboItemsName.BoundText <> "" Then
            '    StrSQL = StrSQL + " AND ItemID=" & Me.DCboItemsName.BoundText & ""
            'End If
            'StrSQL = StrSQL + " Group By StoreName"
            'StrSQL = StrSQL + " Order By StoreName"
            StrSQL = "SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.TblStore.StoreName"
            StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
            StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
            StrSQL = StrSQL + "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
            StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
            StrSQL = StrSQL + "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
            StrSQL = StrSQL + "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"

            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
 
            StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
            StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
            StrSQL = StrSQL + " and Item_ID =" & val(DCboItemsName.BoundText)
     If val(Me.DCboStoreName.BoundText) <> 0 Then
                    StrSQL = StrSQL + " AND Transactions.StoreID=" & val(Me.DCboStoreName.BoundText) & ""
     End If
'            StrSQL = StrSQL + " GROUP BY dbo.TblStore.StoreName, dbo.TblUnites.UnitName"
StrSQL = StrSQL + " GROUP BY dbo.TblStore.StoreName "

            StrSQL = StrSQL + " HAVING      (SUM(dbo.Transaction_Details.quantity * dbo.TransactionTypes.StockEffect) <> 0)"
        
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "Select Sum(Qty) as SumQty ,StoreName "
            StrSQL = StrSQL + " From QryGardComplete"
            StrSQL = StrSQL + " Where ItemCode='" & Trim(XPTxtCode.Text) & "'"

            If Me.DCboItemsName.BoundText <> "" Then
                StrSQL = StrSQL + " AND ItemID=" & Me.DCboItemsName.BoundText & ""
            End If
     
        End If

        RsSumQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsSumQty.BOF Or RsSumQty.EOF) Then

            With Me.FgSum
                RsSumQty.MoveFirst
                .Rows = .FixedRows + RsSumQty.RecordCount

                For Num = .FixedRows To .Rows - 1
                    .TextMatrix(Num, .ColIndex("NumIndex")) = Num

                    If Not (IsNull(RsSumQty("SumQty").value)) Then
                        .TextMatrix(Num, .ColIndex("Quantity")) = Round(RsSumQty("SumQty").value, SystemOptions.SysDefCurrencyForamt)
                    Else
                        .TextMatrix(Num, .ColIndex("Quantity")) = ""
                    End If

                    .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(RsSumQty("StoreName").value), "", (RsSumQty("StoreName").value))
                    RsSumQty.MoveNext
                Next Num

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        RsSumQty.Close
        Set RsSumQty = Nothing
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(2).Caption = "·« ÊÃœ «Ì… ﬂ„Ì«  „‰ «·’‰›"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(2).Caption = "There Is NO Item Stock"
        End If
    End If

    If Me.DCboItemsName.BoundText <> "" Then
        StrSQL = "Select * From TblItems Where ItemID=" & Me.DCboItemsName.BoundText & ""
        Set RsData = New ADODB.Recordset
        RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsData.BOF Or RsData.EOF) Then
            lbl(8).Caption = IIf(IsNull(RsData("SallingPrice").value), "", RsData("SallingPrice").value)
            lbl(9).Caption = IIf(IsNull(RsData("CustomerPrice").value), "", RsData("CustomerPrice").value)
            lbl(10).Caption = IIf(IsNull(RsData("DealerPrice").value), "", RsData("DealerPrice").value)
        End If
    
        Set RsData = New ADODB.Recordset
        StrSQL = "select * From ItemsPrice where Item_ID=" & Me.DCboItemsName.BoundText
        RsData.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsData.EOF Or RsData.BOF) Then
            FgItemPriceList.Rows = RsData.RecordCount + 1

            For RowNum = 1 To RsData.RecordCount

                With FgItemPriceList
                    .TextMatrix(RowNum, .ColIndex("NumIndex")) = RowNum
                    .TextMatrix(RowNum, .ColIndex("Form")) = IIf(IsNull(RsData("From").value), "", Trim(RsData("From").value))
                    .TextMatrix(RowNum, .ColIndex("To")) = IIf(IsNull(RsData("To").value), "", Trim(RsData("To").value))
                    .TextMatrix(RowNum, .ColIndex("Price")) = IIf(IsNull(RsData("Price").value), "", Trim(RsData("Price").value))
                End With

                RsData.MoveNext
            Next RowNum

            FgItemPriceList.AutoSize 0, FgItemPriceList.Cols - 1, False
        End If

        ItemTransInfo = GetLastItemTrans(val(Me.DCboItemsName.BoundText))
        Me.lbl(16).Caption = ItemTransInfo.TransactionSerial

        If ItemTransInfo.TransactionDate <> "" Then
            Me.lbl(17).Caption = DisplayDate(CDate(ItemTransInfo.TransactionDate))
        End If

        Me.lbl(18).Caption = ItemTransInfo.StrCustomerName
        Me.lbl(19).Caption = ItemTransInfo.SngItemPrice
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub GetItemData(Optional LngItemID As Long = 0, _
                        Optional StrItemCode As String = "")

    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If LngItemID = 0 And StrItemCode <> "" Then
        StrSQL = "select * From TblItems where ItemCode='" & StrItemCode & " ' or  barCodeNO='" & Trim(StrItemCode) & "'"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If Opt(0).value = True Then
                Me.LblHaveSerial.Caption = WriteSerialCaption(RsTemp("HaveSerial").value)
                DCboItemsName.BoundText = RsTemp("ItemID").value
            ElseIf Opt(3).value = True Then
                Me.Label1.Caption = WriteSerialCaption(RsTemp("HaveSerial").value)
                DataCombo1.BoundText = RsTemp("ItemID").value
            End If
            
            'Cmd_Click (0)
        Else
            DCboItemsName.BoundText = ""
        End If

        If Me.Tag <> "" Then
            'Cmd_Click (0)
            Me.Tag = ""
        End If

    ElseIf LngItemID <> 0 And StrItemCode = "" Then
        StrSQL = "select * From TblItems where ItemID=" & LngItemID
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If Opt(0).value = True Then
                Me.LblHaveSerial.Caption = WriteSerialCaption(RsTemp("HaveSerial").value)
                DCboItemsName.BoundText = RsTemp("ItemID").value
            ElseIf Opt(3).value = True Then
                Me.Label1.Caption = WriteSerialCaption(RsTemp("HaveSerial").value)
                DataCombo1.BoundText = RsTemp("ItemID").value
            End If
        End If

        'Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Cmd(7).Caption = "Print"
    Me.Caption = "Check for Item Stock"
    Opt(0).Caption = "Check For Item Stock"
    Opt(1).Caption = "Check For All Items Group Stock"
    Opt(2).Caption = "Check For All Complex Items "
    LblLink.Caption = "Show Items Report Screen"
    lbl(5).Caption = "Item Code"
    lbl(6).Caption = "Item Name"
    CmdHelp.Caption = "Help"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    With fg
    
    .TextMatrix(0, .ColIndex("price")) = "Price"
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
        .TextMatrix(0, .ColIndex("Serial")) = "Part Serial"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("ClassName")) = "Class"
    
        .TextMatrix(0, .ColIndex("ItemCase")) = "Item Case"
        .TextMatrix(0, .ColIndex("ColorName")) = "Color "
        .TextMatrix(0, .ColIndex("ItemSize")) = "Item Size"
        .AutoSize 0, .Cols - 1, False
    End With

    TabMain(0).TabCaption(0) = "Quantity"
    TabMain(0).TabCaption(1) = "Item Price"
    lbl(3).Caption = "User Price:"
    lbl(4).Caption = "Customer Price:"
    lbl(7).Caption = "Dlear Price:"
    lbl(11).Caption = "Price in Items Price list"
    Fra.Caption = "Last Invoice Price"

    With FgItemPriceList
        .TextMatrix(0, .ColIndex("NumIndex")) = "S"
        .TextMatrix(0, .ColIndex("Form")) = "Form"
        .TextMatrix(0, .ColIndex("To")) = "To"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
    End With

    With Me.FgItems
        .TextMatrix(0, .ColIndex("Serial")) = "S"
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("ItemQty")) = "Item Quantity"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgSum
        .TextMatrix(0, .ColIndex("NumIndex")) = "S"
        .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .AutoSize 0, .Cols - 1, False
    End With

    lbl(12).Caption = "Inv NO:"
    lbl(13).Caption = "Inv Date:"
    lbl(14).Caption = "Customer Name:"
    lbl(15).Caption = "Item Price:"
    lbl(20).Caption = "Group Name:"
    lbl(21).Caption = "Group Code:"
    lbl(23).Caption = "Total Stock:"
    
    Opt(3).Caption = "Inquiry about the expiration date of items"
    TabMain(1).TabCaption(3) = "Inquiry about the expiration date of items"
    
    CheckBox1.RightToLeft = False
    CheckBox1.Caption = "Display items that do not have an expiry date"
    lbl(31).Caption = "Before the expiry date"
    lbl(35).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "No"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("Qty")) = "Current Qty"
        .TextMatrix(0, .ColIndex("ExpiryDate")) = "Expiry Date"
    End With
    
    Opt(4).Caption = "Query items that have reached the demand limit"
    TabMain(1).TabCaption(4) = "Query items that have reached the demand limit"
    

    lbl(36).Caption = "Store"
    lbl(38).Caption = "No of Items"
    
    With VSFlexGrid2
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("Requst")) = "Requst"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
        .TextMatrix(0, .ColIndex("StoreID")) = "Store ID"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("DefalutPrice")) = "Defalut Price"

    End With
End Sub

Private Sub ClearData()
    'Clear the form for the new data
    fg.Clear flexClearScrollable, flexClearEverything
    FgItemPriceList.Clear flexClearScrollable, flexClearEverything
    FgItemPriceList.Rows = FgItemPriceList.FixedRows
    fg.Rows = 1
    FgSum.Rows = 1
    Me.LblHaveSerial.Caption = ""
    Me.lbl(2).Caption = ""

    Me.lbl(8).Caption = ""
    Me.lbl(9).Caption = ""
    Me.lbl(10).Caption = ""

    Me.lbl(16).Caption = ""
    Me.lbl(17).Caption = ""
    Me.lbl(18).Caption = ""
    Me.lbl(19).Caption = ""
End Sub

Public Property Get LngGridRow() As Long
    LngGridRow = m_LngGridRow
End Property

Public Property Let LngGridRow(ByVal vNewValue As Long)
    m_LngGridRow = vNewValue
End Property

Private Sub LblLink_Click()
    OpenScreen PopUpShowItemCardScreen, val(Me.DCboItemsName.BoundText), 1
End Sub

Private Sub Opt_Click(Index As Integer)
    If Opt(0).value = True Then
        Me.TabMain(1).CurrTab = 0
        Me.TabMain(1).TabVisible(0) = True
        Me.TabMain(1).TabVisible(1) = False
        Me.TabMain(1).TabVisible(2) = False
        Me.TabMain(1).TabVisible(3) = False
        Me.TabMain(1).TabVisible(4) = False
        Me.TabMain(1).TabCaption(0) = Me.Opt(Index).Caption
    ElseIf Opt(1).value = True Then
        Me.TabMain(1).CurrTab = 1
        Me.TabMain(1).TabVisible(0) = False
        Me.TabMain(1).TabVisible(1) = True
        Me.TabMain(1).TabVisible(2) = False
        Me.TabMain(1).TabVisible(3) = False
        Me.TabMain(1).TabVisible(4) = False
        Me.TabMain(1).TabCaption(1) = Me.Opt(Index).Caption
    ElseIf Opt(2).value = True Then
        Me.TabMain(1).CurrTab = 2
        Me.TabMain(1).TabVisible(0) = False
        Me.TabMain(1).TabVisible(1) = False
        Me.TabMain(1).TabVisible(2) = True
        Me.TabMain(1).TabVisible(3) = False
        Me.TabMain(1).TabVisible(4) = False
        Me.TabMain(1).TabCaption(2) = Me.Opt(Index).Caption
    ElseIf Opt(3).value = True Then
       DataCombo1_Change
        Me.TabMain(1).CurrTab = 3
        Me.TabMain(1).TabVisible(0) = False
        Me.TabMain(1).TabVisible(1) = False
        Me.TabMain(1).TabVisible(2) = False
        Me.TabMain(1).TabVisible(3) = True
        Me.TabMain(1).TabVisible(4) = False
        Me.TabMain(1).TabCaption(3) = Me.Opt(Index).Caption
        Cmd_Click (0)
    ElseIf Opt(4).value = True Then
        Me.TabMain(1).CurrTab = 4
        Me.TabMain(1).TabVisible(0) = False
        Me.TabMain(1).TabVisible(1) = False
        Me.TabMain(1).TabVisible(2) = False
        Me.TabMain(1).TabVisible(3) = False
        Me.TabMain(1).TabVisible(4) = True
        Me.TabMain(1).TabCaption(4) = Me.Opt(Index).Caption
    End If

End Sub

Private Sub XPTxtCode_Change()
    Cmd_Click (0)
End Sub

Private Sub XPTxtCode_KeyDown(KeyCode As Integer, _
                              Shift As Integer)
    Dim LngTempID As Long

    If KeyCode = vbKeyReturn Then
        If Trim(Me.XPTxtCode.Text) = "" Then Exit Sub
        LngTempID = GetItemID(Trim(Me.XPTxtCode.Text))

        If LngTempID = 0 Then
            Me.DCboItemsName.BoundText = ""
            Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        ElseIf val(Me.DCboItemsName.BoundText) <> LngTempID Then
            DCboItemsName.BoundText = LngTempID
        End If
    End If

End Sub

Private Sub XPTxtCode_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If XPTxtCode.Text <> "" Then
            Cmd_Click (0)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub RetriveGroup()
    Dim i As Integer

    If rs.State = adStateClosed Then
        Exit Sub
    End If

    With Me.FgItems
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)

                If Not IsNull(rs("Qty").value) Then
                    .TextMatrix(i, .ColIndex("ItemQty")) = Format(rs("Qty").value, SystemOptions.SysDefCurrencyForamt)
                End If

                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                rs.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False

        If FgItems.Rows > 1 Then
            Me.lbl(22).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemQty"), .Rows - 1, .ColIndex("ItemQty"))
        End If

    End With

End Sub

Private Function WriteSerialCaption(BolSerialType As Boolean)

    If SystemOptions.UserInterface = ArabicInterface Then
        If BolSerialType = True Then
            WriteSerialCaption = "«·’‰› ·Â ”Ì—Ì«·"
        Else
            WriteSerialCaption = "«·’‰› ·Ì” ·Â ”Ì—Ì«·"
        End If

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        If BolSerialType = True Then
            WriteSerialCaption = "Item Have Serial System"
        Else
            WriteSerialCaption = "Item Have NO Serial System"
        End If
    End If

End Function

Private Sub RetriveAssbliedItem()
    Dim i As Integer

    If rs.State = adStateClosed Then
        Exit Sub
    End If
        
    With Me.FgItemsA
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("PartItemID").value), "", rs("PartItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("PartItemQty")) = IIf(IsNull(rs("PartItemQty").value), "", rs("PartItemQty").value)

                If Not IsNull(rs("Qty").value) Then
                    .TextMatrix(i, .ColIndex("ItemQty")) = Format(rs("Qty").value, SystemOptions.SysDefCurrencyForamt)
                End If

                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
            
                If val(.TextMatrix(i, .ColIndex("ItemQty"))) = 0 Then
                    .TextMatrix(i, .ColIndex("InQty")) = 0
                Else
                    .TextMatrix(i, .ColIndex("InQty")) = val(.TextMatrix(i, .ColIndex("ItemQty"))) \ val(.TextMatrix(i, .ColIndex("PartItemQty")))
                End If

                rs.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False

        If FgItemsA.Rows > 1 Then
            Me.lbl(27).Caption = .Aggregate(flexSTMin, .FixedRows, .ColIndex("InQty"), .Rows - 1, .ColIndex("InQty"))
        End If

    End With

End Sub

Private Sub FillGrid()
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As ADODB.Recordset
    On Error GoTo hErr

    Set RsTemp = New ADODB.Recordset
     With VSFlexGrid2
            .ColHidden(.ColIndex("GroupName")) = False
             .Clear flexClearScrollable, flexClearEverything
            .Rows = 1
            .ExplorerBar = flexExSortShowAndMove
     End With
 ' My_SQL = " SELECT     TOP 100 PERCENT dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee, "
 ' My_SQL = My_SQL & "                    dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode AS GropFullcode,"
 ' My_SQL = My_SQL & "                    dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
 ' My_SQL = My_SQL & "                    dbo.GeQtyofStore(dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID) AS QNty, dbo.TblSettsRequestLimitDet.UnitFactor,"
 ' My_SQL = My_SQL & "                    dbo.TblStore.storename , dbo.TblStore.storenamee, dbo.TblStore.code"
 ' My_SQL = My_SQL & " FROM         dbo.Groups RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblItems RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblUnites RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblStore RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblSettsRequestLimitDet INNER JOIN"
'  My_SQL = My_SQL & "                    dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID ON"
'  My_SQL = My_SQL & "                    dbo.TblStore.StoreID = dbo.TblSettsRequestLimitDet.StoreID ON dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON"
'  My_SQL = My_SQL & "                    dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID"
'  My_SQL = My_SQL & " Where (dbo.TblSettsRequestLimitDet.typ = 0)"


'My_SQL = My_SQL & " GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.ItemName, "
'My_SQL = My_SQL & "                      dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode, dbo.TblSettsRequestLimitDet.StoreID,"
'My_SQL = My_SQL & "                      dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName,"
'My_SQL = My_SQL & "                      dbo.TblStore.storenamee , dbo.TblStore.code, dbo.GeQtyofStore(dbo.TblSettsRequestLimitDet.StoreId, dbo.TblSettsRequestLimitDet.ItemID)"
'My_SQL = My_SQL & " ORDER BY dbo.TblSettsRequestLimitDet.StoreID"
My_SQL = " SELECT     Xb.Qty, Xb.ConsuRateLowQty, Xb.GroupName, Xb.GroupNamee, Xb.ItemName, Xb.ItemNamee, Xb.Fullcode, Xb.UnitName, Xb.UnitNamee, Xb.GropFullcode,"
My_SQL = My_SQL & "                      Xb.StoreId , Xb.ItemID, Xb.ConsuRate, Xb.UnitFactor, Xb.storename, Xb.storenamee, Xb.code, BX.QNty"
My_SQL = My_SQL & " FROM         (SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
My_SQL = My_SQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
My_SQL = My_SQL & "                                              dbo.Groups.Fullcode AS GropFullcode, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
My_SQL = My_SQL & "                                              dbo.TblStore.code"
My_SQL = My_SQL & "                        FROM         dbo.TblStore INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblUnites INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblItems INNER JOIN"
My_SQL = My_SQL & "                                              dbo.Groups INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet INNER JOIN"
My_SQL = My_SQL & "                                              dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID ON"
My_SQL = My_SQL & "                                              dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON"
My_SQL = My_SQL & "                                              dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON dbo.TblStore.StoreID = dbo.TblSettsRequestLimitDet.StoreID"
My_SQL = My_SQL & "                        Where (dbo.TblSettsRequestLimitDet.Typ = 0)"
My_SQL = My_SQL & "                        GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
My_SQL = My_SQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code) Xb INNER JOIN"
My_SQL = My_SQL & "                          (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID"
My_SQL = My_SQL & "                             FROM         dbo.Transactions INNER JOIN"
My_SQL = My_SQL & "                                                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "                                                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
My_SQL = My_SQL & "                             GROUP BY Item_ID, StoreID)  BX ON BX.Item_ID = Xb.ItemID  AND BX.StoreID = Xb.StoreID and BX.Item_ID in (select ItemID from TblSettsRequestLimitDet)"
My_SQL = My_SQL & " where 1=1"
    If SystemOptions.usertype = UserAdminAll Then
  If val(DcbStore.BoundText) <> 0 Then
        My_SQL = My_SQL & " and   Xb.StoreID =" & val(DcbStore.BoundText) & ""
 End If
 Else
         My_SQL = My_SQL & " and   Xb.StoreID =" & val(DcbStore.BoundText) & ""
 End If
    RsTemp.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

     
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Me.lbl(37).Caption = RsTemp.RecordCount

        With VSFlexGrid2
            .ColHidden(.ColIndex("GroupName")) = False
             .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .ExplorerBar = flexExSortShowAndMove

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                .TextMatrix(RowNum, .ColIndex("UnitFactor")) = IIf(IsNull(RsTemp("UnitFactor").value), 0, RsTemp("UnitFactor").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "", RsTemp("GroupName").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "", RsTemp("StoreName").value)
                .TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(RsTemp("UnitName").value), "", RsTemp("UnitName").value)
                Else
                .TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(RsTemp("UnitNamee").value), "", RsTemp("UnitNamee").value)
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupNamee").value), "", RsTemp("GroupNamee").value)
                 .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreNamee").value), "", RsTemp("StoreNamee").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemNamee").value), "", RsTemp("ItemNamee").value)
                End If
                .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), "", RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                
                .TextMatrix(RowNum, .ColIndex("Requst")) = IIf(IsNull(RsTemp("Qty").value), "0", RsTemp("Qty").value)
                  .TextMatrix(RowNum, .ColIndex("qty")) = IIf(IsNull(RsTemp("QNty").value), "0", RsTemp("QNty").value)
If val(.TextMatrix(RowNum, .ColIndex("UnitFactor"))) <> 0 Then
   .TextMatrix(RowNum, .ColIndex("qty")) = Round(val(.TextMatrix(RowNum, .ColIndex("qty"))) / val(.TextMatrix(RowNum, .ColIndex("UnitFactor"))), 2)
End If
                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), 0, RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "", RsTemp("StoreName").value)
                
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "", RsTemp("GroupName").value)
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.lbl(37).Caption = ""
    End If

    Exit Sub
hErr:
End Sub
Function print_reportReq(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  
  MySQL = "SELECT     TOP 100 PERCENT dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee, "
  MySQL = MySQL & "                     dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode AS GropFullcode,"
  MySQL = MySQL & "                    dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
  MySQL = MySQL & "                    dbo.GeQtyofStore(dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID) AS QNty, dbo.TblSettsRequestLimitDet.UnitFactor,"
  MySQL = MySQL & "                    dbo.TblStore.storename , dbo.TblStore.StoreNamee, dbo.TblStore.Code"
  MySQL = MySQL & "   FROM         dbo.Groups RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblItems RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblUnites RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.Transaction_Details RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblSettsRequestLimitDet LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblStore ON dbo.TblSettsRequestLimitDet.StoreID = dbo.TblStore.StoreID ON dbo.Transaction_Details.Item_ID = dbo.TblSettsRequestLimitDet.ItemID ON"
  MySQL = MySQL & "                    dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON"
  MySQL = MySQL & "                    dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID"
  MySQL = MySQL & "  Where (dbo.TblSettsRequestLimitDet.typ = 0)"
  
    If SystemOptions.usertype = UserAdminAll Then
  If val(DcbStore.BoundText) <> 0 Then
        MySQL = MySQL & " and   dbo.TblSettsRequestLimitDet.StoreId=" & val(DcbStore.BoundText) & ""
 End If
 Else
         MySQL = MySQL & " and   dbo.TblSettsRequestLimitDet.StoreId=" & val(DcbStore.BoundText) & ""
 End If
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RequestItemsSerial.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RequestItemsSerialE.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
       If SystemOptions.UserInterface = ArabicInterface Then
         Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
         Msg = "No Data"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
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
