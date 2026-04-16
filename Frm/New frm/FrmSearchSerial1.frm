VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSearchSerial1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈” ⁄·«„ ⁄‰  «·«’‰«› «·»œÌ·…"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   HelpContextID   =   10
   Icon            =   "FrmSearchSerial1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   10785
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
      Height          =   8805
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10785
      _cx             =   19024
      _cy             =   15531
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
      _GridInfo       =   $"FrmSearchSerial1.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   720
         Index           =   2
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8070
         Width           =   10755
         _cx             =   18971
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
            Left            =   5730
            TabIndex        =   6
            Top             =   150
            Width           =   1785
            _ExtentX        =   3149
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
            Left            =   8730
            TabIndex        =   7
            Top             =   150
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   7200
            TabIndex        =   8
            Top             =   150
            Width           =   1950
            _ExtentX        =   3440
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
            Cancel          =   -1  'True
            Height          =   390
            Index           =   2
            Left            =   270
            TabIndex        =   9
            Top             =   150
            Width           =   1815
            _ExtentX        =   3201
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
            Left            =   2430
            TabIndex        =   11
            Top             =   120
            Width           =   2730
            _ExtentX        =   4815
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3825
         Index           =   1
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   10755
         _cx             =   18971
         _cy             =   6747
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
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            Height          =   1860
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   615
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
            Height          =   1725
            Index           =   0
            Left            =   5745
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   -9495
            Value           =   -1  'True
            Width           =   4965
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… „Ã„Ê⁄…"
            Height          =   1725
            Index           =   1
            Left            =   5745
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   -11715
            Width           =   4965
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰› „Ã„⁄"
            Height          =   1725
            Index           =   2
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   -9375
            Width           =   4980
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   8040
         Index           =   1
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   10755
         _cx             =   18971
         _cy             =   14182
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
         Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›|≈” ⁄·«„ ⁄‰ ﬂ„Ì… „Ã„Ê⁄…|’‰› „Ã„⁄"
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
         Flags(1)        =   2
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7665
            Index           =   5
            Left            =   45
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   330
            Width           =   10665
            _cx             =   18812
            _cy             =   13520
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
            Begin VB.TextBox TxtAssbliedItemCode 
               Alignment       =   1  'Right Justify
               Height          =   480
               Left            =   4485
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   135
               Width           =   4545
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
               Height          =   420
               Left            =   4605
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   1155
               Visible         =   0   'False
               Width           =   5910
            End
            Begin MSDataListLib.DataCombo DcboStores 
               Height          =   315
               Left            =   150
               TabIndex        =   14
               Top             =   1185
               Visible         =   0   'False
               Width           =   3030
               _ExtentX        =   5345
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItemsA 
               Height          =   6570
               Left            =   75
               TabIndex        =   17
               Top             =   1035
               Width           =   10515
               _cx             =   18547
               _cy             =   11589
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
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSearchSerial1.frx":040B
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
               Left            =   1050
               TabIndex        =   18
               Top             =   645
               Width           =   7980
               _ExtentX        =   14076
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdItemSearch 
               Height          =   405
               Index           =   1
               Left            =   150
               TabIndex        =   19
               Top             =   630
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   714
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
               ButtonImage     =   "FrmSearchSerial1.frx":05CC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·’‰›"
               Height          =   450
               Index           =   24
               Left            =   8640
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   570
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   450
               Index           =   25
               Left            =   8640
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   120
               Width           =   1845
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
               Height          =   285
               Index           =   26
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   6705
               Visible         =   0   'False
               Width           =   4875
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   27
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   6705
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õœœ «·„Œ“‰"
               Height          =   390
               Index           =   28
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   1140
               Width           =   1275
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7665
            Index           =   4
            Left            =   11700
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   330
            Width           =   10665
            _cx             =   18812
            _cy             =   13520
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
               Height          =   525
               Left            =   2385
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   45
               Width           =   6150
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   6075
               Left            =   75
               TabIndex        =   27
               Top             =   1200
               Width           =   10515
               _cx             =   18547
               _cy             =   10716
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
               FormatString    =   $"FrmSearchSerial1.frx":0B66
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
               Left            =   495
               TabIndex        =   28
               Top             =   630
               Width           =   8040
               _ExtentX        =   14182
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄…"
               Height          =   450
               Index           =   20
               Left            =   8385
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   615
               Width           =   2100
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
               Height          =   570
               Index           =   21
               Left            =   8385
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   75
               Width           =   2100
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
               Height          =   255
               Index           =   22
               Left            =   75
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   7350
               Width           =   2550
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
               Height          =   315
               Index           =   23
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   7290
               Width           =   2160
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7665
            Index           =   3
            Left            =   11400
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   330
            Width           =   10665
            _cx             =   18812
            _cy             =   13520
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
            Begin VB.TextBox XPTxtCode 
               Alignment       =   1  'Right Justify
               Height          =   495
               Left            =   4335
               MaxLength       =   40
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   60
               Width           =   4605
            End
            Begin C1SizerLibCtl.C1Tab TabMain 
               Height          =   5475
               Index           =   0
               Left            =   45
               TabIndex        =   35
               Top             =   1365
               Width           =   10470
               _cx             =   18468
               _cy             =   9657
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
               Picture(0)      =   "FrmSearchSerial1.frx":0C55
               Picture(1)      =   "FrmSearchSerial1.frx":0FEF
               Flags(1)        =   2
               Picture(2)      =   "FrmSearchSerial1.frx":1389
               Flags(2)        =   2
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   5010
                  Index           =   6
                  Left            =   11400
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   10365
                  _cx             =   18283
                  _cy             =   8837
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
                  Height          =   5010
                  Index           =   7
                  Left            =   11100
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   10365
                  _cx             =   18283
                  _cy             =   8837
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
                     TabIndex        =   38
                     Top             =   30
                     Width           =   2865
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   19
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   46
                        Top             =   1110
                        Width           =   1545
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   18
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   45
                        Top             =   840
                        Width           =   1545
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   17
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   44
                        Top             =   570
                        Width           =   1545
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Height          =   255
                        Index           =   16
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   43
                        Top             =   300
                        Width           =   1545
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
                        TabIndex        =   42
                        Top             =   1110
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
                        TabIndex        =   41
                        Top             =   840
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
                        TabIndex        =   40
                        Top             =   570
                        Width           =   1185
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "—ﬁ„ «·›« Ê—…:"
                        ForeColor       =   &H00000040&
                        Height          =   255
                        Index           =   12
                        Left            =   1620
                        RightToLeft     =   -1  'True
                        TabIndex        =   39
                        Top             =   300
                        Width           =   1185
                     End
                  End
                  Begin ImpulseButton.ISButton ISButton1 
                     Height          =   405
                     Left            =   270
                     TabIndex        =   47
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
                     Left            =   270
                     TabIndex        =   48
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
                     TabIndex        =   49
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
                     FormatString    =   $"FrmSearchSerial1.frx":1723
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
                     TabIndex        =   57
                     Top             =   960
                     Width           =   2055
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   10
                     Left            =   3390
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   660
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   9
                     Left            =   3390
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   360
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   8
                     Left            =   3420
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   60
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(œÌ·—)"
                     Height          =   285
                     Index           =   7
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   660
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
                     TabIndex        =   52
                     Top             =   360
                     Width           =   1395
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·»Ì⁄(„” Â·ﬂ)"
                     Height          =   285
                     Index           =   3
                     Left            =   4110
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   60
                     Width           =   1395
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
                     TabIndex        =   50
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   4395
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   5010
                  Index           =   8
                  Left            =   45
                  TabIndex        =   58
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   10365
                  _cx             =   18283
                  _cy             =   8837
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
                     Height          =   3375
                     Left            =   30
                     TabIndex        =   59
                     Top             =   30
                     Width           =   6345
                     _cx             =   11192
                     _cy             =   5953
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
                     Cols            =   9
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmSearchSerial1.frx":17A8
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
                     Left            =   8190
                     TabIndex        =   60
                     Top             =   2460
                     Width           =   6345
                     _cx             =   11192
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
                     FormatString    =   $"FrmSearchSerial1.frx":1901
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
               Left            =   1080
               TabIndex        =   61
               Top             =   540
               Width           =   7860
               _ExtentX        =   13864
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdItemSearch 
               Height          =   435
               Index           =   0
               Left            =   45
               TabIndex        =   62
               Top             =   525
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   767
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
               ButtonImage     =   "FrmSearchSerial1.frx":1989
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseAniLabel.ISAniLabel LblLink 
               Height          =   330
               Left            =   45
               TabIndex        =   63
               Top             =   105
               Width           =   4245
               _ExtentX        =   7488
               _ExtentY        =   582
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
               MouseIcon       =   "FrmSearchSerial1.frx":1F23
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   "⁄—÷ ‘«‘…  ﬁ«—Ì— «·’‰›"
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
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
               Height          =   375
               Index           =   2
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   1005
               Width           =   5145
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
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   6915
               Width           =   8010
            End
            Begin VB.Label LblHaveSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«’‰«› «·»œÌ·… ÂÌ"
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   6495
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   1005
               Width           =   3930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   480
               Index           =   5
               Left            =   8865
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   60
               Width           =   1560
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·’‰›"
               Height          =   420
               Index           =   6
               Left            =   8865
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   585
               Width           =   1560
            End
         End
      End
   End
End
Attribute VB_Name = "FrmSearchSerial1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rs As New ADODB.Recordset
Dim cSearchDcbo(3) As clsDCboSearch
Dim m_LngGridRow As Long
Dim FirstPeriodDateInthisYear  As Date

Private Sub Chk_Click()
    Me.lbl(28).Enabled = CBool(Me.Chk.value)
    Me.DcboStores.Enabled = CBool(Me.Chk.value)
End Sub

Public Sub Cmd_Click(Index As Integer)
Retrive
Exit Sub

    Dim Msg As String
    Dim StSQL As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
    
            If rs.State = adStateOpen Then
                rs.Close
            End If

          '  If Opt(0).value = True Then
                If DcboAssbliedItems.BoundText = "" Then
                    If Trim(XPTxtCode.Text) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "»—Ã«¡  ÕœÌœ «·’‰›..!!"
                            '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        ElseIf SystemOptions.UserInterface = EnglishInterface Then
                            Msg = "Please choose an Item Name....!"
                            '    MsgBox Msg, vbOKOnly + vbExclamation, App.Title
                        End If

                        DcboAssbliedItems.SetFocus
                        SendKeys "{F4}"
                        'Exit Sub
                    End If
                End If

                rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount < 1 Then
                    'Exit Sub
                End If

                Retrive
          '  ElseIf Opt(1).value = True Then
'
'                If val(Me.DcboGroupID.BoundText) = 0 Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·„Ã„Ê⁄…....!!!!"
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                    Exit Sub
'                End If
'
'                rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'                If rs.RecordCount < 1 Then
'
'                End If

'                RetriveGroup
'            ElseIf Opt(2).value = True Then
'
'                If val(Me.DcboAssbliedItems.BoundText) = 0 Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·’‰› «·„Ã„⁄....!!!!"
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                    Exit Sub
'                End If
'
'                If Me.Chk.value = vbChecked Then
'                    If Me.DcboStores.BoundText = "" Then
'                        Msg = "ÌÃ»  ÕœÌœ «”„ «·„Œ“‰...!!!"
'                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                        Exit Sub
'                    End If
'                End If
'
'                rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                RetriveAssbliedItem
'            End If
'
        Case 1
            clear_all Me
            ClearData

        Case 2
            Unload Me

        Case 7
            printing
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Function printing()
 
 '   Dim VReport As ClsGardReport
 '
   ' Set VReport = New ClsGardReport
 '
    'VReport.ShowGardData2 Build_Sql, FirstPeriodDateInthisYear, Date, DCboItemsName.Text
    
End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdItemSearch_Click(Index As Integer)

    Select Case Index

        Case 0
            PutFormOnTop Me.hWnd, False
            ModOpenScreen.ShowDialogItemsSearch Me.DCboItemsName
            PutFormOnTop Me.hWnd, True

        Case 1
            PutFormOnTop Me.hWnd, False
            ModOpenScreen.ShowDialogItemsSearch Me.DcboAssbliedItems
            PutFormOnTop Me.hWnd, True
    End Select

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
Retrive
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
        StrSQL = "Select * From TblItems Where ItemID=" & Me.DCboItemsName.BoundText & ""
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



Private Sub Fg_DblClick()
    Dim LngRow As Long
    On Error GoTo ErrTrap

    If FG.Row < FG.FixedRows Then Exit Sub
    If FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) = "·«ÌÊÃœ" Then Exit Sub
    If mdifrmmain.ActiveForm Is Nothing Then Exit Sub

    LngRow = Me.LngGridRow

    If Txt.Text = "Search" Then
        If Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("Serial")) <> "" Then
            mdifrmmain.ActiveForm.FG.TextMatrix(LngRow, mdifrmmain.ActiveForm.FG.ColIndex("Serial")) = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Retutn" Then

        If Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("Serial")) <> "" Then
            FrmReturnpurchases.FG.TextMatrix(FrmReturnpurchases.FG.Row, FrmReturnpurchases.FG.ColIndex("Serial")) = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Moving" Then

        If Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("Serial")) <> "" Then
            FrmMoving.FG.TextMatrix(FrmMoving.FG.Row, FrmMoving.FG.ColIndex("Serial")) = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Destruction" Then

        If Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("Serial")) <> "" Then
            FrmDestruction.FG.TextMatrix(FrmDestruction.FG.Row, FrmDestruction.FG.ColIndex("Serial")) = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
            Unload Me
        End If

    ElseIf Txt.Text = "Replace" Then

        If Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("Serial")) <> "" Then
            FrmReplace.TxtNewSerial.Text = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
            FrmReplace.DCboStoreName.BoundText = FG.TextMatrix(FG.Row, FG.ColIndex("StoreName"))
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
        Msg = Msg & Chr(13) & "You can display a item cart report "
        Msg = Msg & "which show all the item transactions "
        Msg = Msg & "you can display this report From the Report Screen"
        lbl(0).Caption = Msg
        Msg = "Press F7 to show item search..."
        lbl(1).Caption = Msg
    Else
        Msg = "„·ÕÊŸ…:-"
        Msg = Msg & Chr(13) & "Ì„ﬂ‰ﬂ ⁄—÷  ﬁ—Ì— »ﬂ«—  «·’‰› «·–Ï Ì⁄—÷ ·ﬂ "
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
'    StrSQL = "select * From TblStore"
'    RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

   ' With FG
   '     StrList = .BuildComboList(RsNote, "StoreName", "StoreID")
'
'        If StrList <> "" Then
'            .ColComboList(.ColIndex("StoreName")) = "|" & StrList
'        End If
'
'    End With

    CenterForm Me
'    FG.WallPaper = BG.SearchWallpaper
'    FgItems.WallPaper = BG.SearchWallpaper
'    FgItemsA.WallPaper = BG.SearchWallpaper
'    FgItemsA.AutoSize 0, FgItemsA.Cols - 1, False
'    Set FgItemPriceList.WallPaper = BG.Picture
'    Set FgSum.WallPaper = BG.Picture
'    FgItemPriceList.AutoSize 0, FgItemPriceList.Cols - 1, False
'
    FormPostion Me, GetPostion
    Set Dcombos = New ClsDataCombos
'    Dcombos.GetItemsNames Me.DCboItemsName, 0
'    Dcombos.GetItemSGroups Me.DcboGroupID
'    Set cSearchDcbo(0) = New clsDCboSearch
'    Set cSearchDcbo(0).Client = Me.DCboItemsName
'    Set cSearchDcbo(1) = New clsDCboSearch
'    Set cSearchDcbo(1).Client = Me.DcboGroupID
    Dcombos.GetItemsNames Me.DcboAssbliedItems
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DcboAssbliedItems
'    Dcombos.GetStores Me.DcboStores
'    Set cSearchDcbo(3) = New clsDCboSearch
'    Set cSearchDcbo(3).Client = Me.DcboStores

   ' Opt(0).value = True
   ' Opt_Click 0
   ' Me.Chk.value = vbUnchecked
   ' Chk_Click
 
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

'Private Function Build_Sql()
'    Dim StrSQL As String
'
'    'On Error GoTo ErrTrap
'   ' If Opt(0).value = True Then
'
'            StrSQL = "SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.TblStore.StoreName, dbo.TblUnites.UnitName, "
'            StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
'            StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'            StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'            StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'            StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
'            StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
'            StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
'            StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
'            StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
'
'            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
'
'            StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
'           ' StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
'           ' StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
'            StrSQL = StrSQL + " and Item_ID =" & val(DcboAssbliedItems.BoundText)
'
'            StrSQL = StrSQL & "  GROUP BY dbo.TblStore.StoreName, dbo.TblUnites.UnitName, dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName,"
'            StrSQL = StrSQL & "  dbo.TblItemsColors.ColorName"
'            StrSQL = StrSQL & "  HAVING      (SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) <> 0)"
'            ' StrSQL = "SELECT * From dbo.QryGardComplete(0)"
'            ' StrSQL = StrSQL + " where ItemCode='" & XPTxtCode.text & "'"
'            ' StrSQL = StrSQL + " Order By StoreName"
'
'     '   End If
'
'    'ElseIf Opt(1).value = True Then
''
''        If SystemOptions.SysDataBaseType = AccessDataBase Then
''            StrSQL = "select * From QryGardComplete"
''            StrSQL = StrSQL + " where GroupID=" & val(Me.DcboGroupID.BoundText) & ""
''        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
''            StrSQL = "SELECT * From dbo.QryGardComplete(0)"
''            StrSQL = StrSQL + " where GroupID=" & val(Me.DcboGroupID.BoundText) & ""
''            StrSQL = StrSQL + " Order By ItemName"
''        End If
''
''    ElseIf Opt(2).value = True Then
''
''        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
''            StrSQL = "SELECT dbo.TblItemsParts.TableID, dbo.TblItemsParts.ItemID," & "dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty," & "dbo.TblItemsParts.PartItemPrice,QryGARDShort.QTY, QryGARDShort.ItemCode," & "QryGARDShort.ItemName, QryGARDShort.StoreID, QryGARDShort.StoreName," & "QryGARDShort.GroupID"
''            StrSQL = StrSQL + " FROM         dbo.TblItemsParts LEFT OUTER JOIN "
''            StrSQL = StrSQL + " dbo.QryGARDShort() QryGARDShort ON " & "dbo.TblItemsParts.PartItemID = QryGARDShort.ItemID"
''            StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & val(Me.DcboAssbliedItems.BoundText) & ""
''
''            If Me.Chk.value = vbChecked Then
''                If val(Me.DcboStores.BoundText) <> 0 Then
''                    StrSQL = StrSQL + " AND QryGARDShort.StoreID=" & val(Me.DcboStores.BoundText) & ""
''                End If
''            End If
''
''            StrSQL = StrSQL + " Order BY dbo.TblItemsParts.TableID"
'        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
''
''        End If
''    End If
''
'    Build_Sql = StrSQL
'    Exit Function
'ErrTrap:
'End Function

Private Sub Retrive()
    Dim StrSQL As String
    Dim Num As Integer
    Dim Rs2 As ADODB.Recordset
    Dim Sql As String
    On Error GoTo ErrTrap
    FgItemsA.Clear flexClearScrollable, flexClearEverything
    FgItemsA.Rows = 1
   Sql = " SELECT     dbo.TblAotherItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee, dbo.TblAotherItems.IDItem,"
   Sql = Sql & "                    dbo.TblAotherItems.Remark, dbo.TblAotherItems.Valu, dbo.TblAotherItems.Quntity, dbo.TblAotherItems.UnitID, dbo.TblUnites.UnitName,"
   Sql = Sql & "                    dbo.TblUnites.UnitNamee"
   Sql = Sql & "     FROM         dbo.TblAotherItems LEFT OUTER JOIN"
   Sql = Sql & "                    dbo.TblUnites ON dbo.TblAotherItems.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   Sql = Sql & "                    dbo.TblItems ON dbo.TblAotherItems.ItemID = dbo.TblItems.ItemID"
   Sql = Sql & "  Where (dbo.TblAotherItems.IDitem = " & val(DcboAssbliedItems.BoundText) & ")"
   Set Rs2 = New ADODB.Recordset
   
   Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
   If Rs2.RecordCount > 0 Then
   With FgItemsA
   Rs2.MoveFirst
   .Rows = Rs2.RecordCount + 1
   For Num = 1 To .Rows - 1
   .TextMatrix(Num, .ColIndex("Serial")) = Num
   .TextMatrix(Num, .ColIndex("UnitID")) = IIf(IsNull(Rs2("UnitID").value), 0, Rs2("UnitID").value)
   .TextMatrix(Num, .ColIndex("Valu")) = IIf(IsNull(Rs2("Valu").value), 0, Rs2("Valu").value)
   .TextMatrix(Num, .ColIndex("Remark")) = IIf(IsNull(Rs2("Remark").value), "", Rs2("Remark").value)
   .TextMatrix(Num, .ColIndex("ItemID")) = IIf(IsNull(Rs2("IDItem").value), "", Rs2("IDItem").value)
   .TextMatrix(Num, .ColIndex("ItemCode")) = IIf(IsNull(Rs2("Fullcode").value), "", Rs2("Fullcode").value)
   .TextMatrix(Num, .ColIndex("ItemQty")) = IIf(IsNull(Rs2("Quntity").value), "", Rs2("Quntity").value)
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(Rs2("ItemName").value), "", Rs2("ItemName").value)
   .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(Rs2("UnitName").value), "", Rs2("UnitName").value)
   Else
   .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(Rs2("ItemNamee").value), "", Rs2("ItemNamee").value)
   .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(Rs2("UnitNamee").value), "", Rs2("UnitNamee").value)
   End If
   Rs2.MoveNext
   Next Num
   .AutoSize 0, .Cols - 1, False
   End With
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
        StrSQL = "select * From TblItems where ItemCode='" & StrItemCode & "'"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Me.LblHaveSerial.Caption = WriteSerialCaption(RsTemp("HaveSerial").value)
            DCboItemsName.BoundText = RsTemp("ItemID").value
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
            Me.LblHaveSerial.Caption = WriteSerialCaption(RsTemp("HaveSerial").value)
            XPTxtCode.Text = RsTemp("ItemCode").value
        End If

        'Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Cmd(7).Caption = "Print"
  Me.Caption = "Alternative Items"
'    Opt(0).Caption = "Check For Item Stock"
'    Opt(1).Caption = "Check For All Items Group Stock"
'    Opt(2).Caption = "Check For All Complex Items "
'    LblLink.Caption = "Show Items Report Screen"
    lbl(25).Caption = "Item Code"
    lbl(24).Caption = "Item Name"
    CmdHelp.Caption = "Help"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
TabMain(1).Caption = "'"
'    With FG
'        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
'        .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
'        .TextMatrix(0, .ColIndex("Serial")) = "Part Serial"
'        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
'        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
'        .TextMatrix(0, .ColIndex("ClassName")) = "Class"
'
'        .TextMatrix(0, .ColIndex("ItemCase")) = "Item Case"
'        .TextMatrix(0, .ColIndex("ColorName")) = "Color "
'        .TextMatrix(0, .ColIndex("ItemSize")) = "Item Size"
'        .AutoSize 0, .Cols - 1, False
'    End With
'
'    TabMain(0).TabCaption(0) = "Quantity"
'    TabMain(0).TabCaption(1) = "Item Price"
'    lbl(3).Caption = "User Price:"
'    lbl(4).Caption = "Customer Price:"
'    lbl(7).Caption = "Dlear Price:"
'    lbl(11).Caption = "Price in Items Price list"
'    Fra.Caption = "Last Invoice Price"
'
   ' With FgItemPriceList
   '     .TextMatrix(0, .ColIndex("NumIndex")) = "S"
   '     .TextMatrix(0, .ColIndex("Form")) = "Form"
   '     .TextMatrix(0, .ColIndex("To")) = "To"
      '  .TextMatrix(0, .ColIndex("Price")) = "Price"
   ' End With

    With Me.FgItemsA
        .TextMatrix(0, .ColIndex("Serial")) = "S"
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("ItemQty")) = "Item Quantity"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("Valu")) = "Value"
        .TextMatrix(0, .ColIndex("Remark")) = "Remarks"
        .AutoSize 0, .Cols - 1, False
    End With

    'With Me.FgSum
    '    .TextMatrix(0, .ColIndex("NumIndex")) = "S"
    '    .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
    '    .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
    '    .AutoSize 0, .Cols - 1, False
    'End With
'
'    lbl(12).Caption = "Inv NO:"
'    lbl(13).Caption = "Inv Date:"
'    lbl(14).Caption = "Customer Name:"
'    lbl(15).Caption = "Item Price:"
'    lbl(20).Caption = "Group Name:"
'    lbl(21).Caption = "Group Code:"
'    lbl(23).Caption = "Total Stock:"
End Sub

Private Sub ClearData()
    'Clear the form for the new data
    FG.Clear flexClearScrollable, flexClearEverything
    FgItemPriceList.Clear flexClearScrollable, flexClearEverything
    FgItemPriceList.Rows = FgItemPriceList.FixedRows
    FG.Rows = 1
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

'Private Sub LblLink_Click()
'    OpenScreen PopUpShowItemCardScreen, val(Me.DCboItemsName.BoundText), 1
'End Sub

'Private Sub Opt_Click(Index As Integer)
'
'    If Opt(0).value = True Then
'        Me.TabMain(1).CurrTab = 0
'        Me.TabMain(1).TabVisible(0) = True
'        Me.TabMain(1).TabVisible(1) = False
'        Me.TabMain(1).TabVisible(2) = False
'        Me.TabMain(1).TabCaption(0) = Me.Opt(Index).Caption
'    ElseIf Opt(1).value = True Then
'        Me.TabMain(1).CurrTab = 1
'        Me.TabMain(1).TabVisible(0) = False
'        Me.TabMain(1).TabVisible(1) = True
'        Me.TabMain(1).TabVisible(2) = False
'        Me.TabMain(1).TabCaption(1) = Me.Opt(Index).Caption
'    ElseIf Opt(2).value = True Then
'        Me.TabMain(1).CurrTab = 2
'        Me.TabMain(1).TabVisible(0) = False
'        Me.TabMain(1).TabVisible(1) = False
'        Me.TabMain(1).TabVisible(2) = True
'        Me.TabMain(1).TabCaption(2) = Me.Opt(Index).Caption
'    End If
'
'End Sub
'
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

'Private Sub RetriveGroup()
'    Dim i As Integer
'
'    If rs.State = adStateClosed Then
'        Exit Sub
'    End If
'
'    With Me.FgItems
'        .Rows = .FixedRows
'
'        If Not (rs.BOF Or rs.EOF) Then
'            .Rows = .FixedRows + rs.RecordCount
'
'            For i = .FixedRows To .Rows - 1
'                .TextMatrix(i, .ColIndex("Serial")) = i
'                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
'                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
'                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
'
'                If Not IsNull(rs("Qty").value) Then
'                    .TextMatrix(i, .ColIndex("ItemQty")) = Format(rs("Qty").value, SystemOptions.SysDefCurrencyForamt)
'                End If
'
'                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
'                rs.MoveNext
'            Next i
'
'        End If
'
'        .AutoSize 0, .Cols - 1, False
'
'        If FgItems.Rows > 1 Then
'            Me.lbl(22).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemQty"), .Rows - 1, .ColIndex("ItemQty"))
'        End If
'
'    End With
'
'End Sub

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

'Private Sub RetriveAssbliedItem()
'    Dim i As Integer
'
'    If rs.State = adStateClosed Then
'        Exit Sub
'    End If
'
'    With Me.FgItemsA
'        .Rows = .FixedRows
'
'        If Not (rs.BOF Or rs.EOF) Then
'            .Rows = .FixedRows + rs.RecordCount
'
'            For i = .FixedRows To .Rows - 1
'                .TextMatrix(i, .ColIndex("Serial")) = i
'                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("PartItemID").value), "", rs("PartItemID").value)
'                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
'                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
'                .TextMatrix(i, .ColIndex("PartItemQty")) = IIf(IsNull(rs("PartItemQty").value), "", rs("PartItemQty").value)

      '          If Not IsNull(rs("Qty").value) Then
      '              .TextMatrix(i, .ColIndex("ItemQty")) = Format(rs("Qty").value, SystemOptions.SysDefCurrencyForamt)
      '          End If

      '          .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
            
              '  If val(.TextMatrix(i, .ColIndex("ItemQty"))) = 0 Then
              '      .TextMatrix(i, .ColIndex("InQty")) = 0
              '  Else
              '      .TextMatrix(i, .ColIndex("InQty")) = val(.TextMatrix(i, .ColIndex("ItemQty"))) \ val(.TextMatrix(i, .ColIndex("PartItemQty")))
              '  End If

      '          rs.MoveNext
      '      Next i
'
'        End If

        

   '     If FgItemsA.Rows > 1 Then
   '         Me.lbl(27).Caption = .Aggregate(flexSTMin, .FixedRows, .ColIndex("InQty"), .Rows - 1, .ColIndex("InQty"))
   '     End If
'
'    End With
'
'End Sub

