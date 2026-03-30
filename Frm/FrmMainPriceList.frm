VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{85FD608E-54A8-11D4-8ED4-00E07D815373}#1.0#0"; "MBClrPkr.ocx"
Begin VB.Form FrmMainPriceList 
   Caption         =   "ﬁ«∆„… «·√”⁄«— "
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   HelpContextID   =   230
   Icon            =   "FrmMainPriceList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10800
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7845
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10800
      _cx             =   19050
      _cy             =   13838
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
      GridRows        =   5
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmMainPriceList.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1110
         Index           =   0
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6720
         Width           =   10770
         _cx             =   18997
         _cy             =   1958
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
         Begin MBColorPicker.ColorPicker CPicColColor 
            Height          =   345
            Left            =   1455
            TabIndex        =   3
            Top             =   105
            Width           =   1140
            _ExtentX        =   1773
            _ExtentY        =   556
            CustomButtonText=   " Œ’Ì’"
            BackColor       =   14871017
            Style           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumColors       =   64
            Color1          =   0
            Color2          =   128
            Color3          =   32768
            Color4          =   32896
            Color5          =   8388608
            Color6          =   8388736
            Color7          =   8421376
            Color8          =   12632256
            Color9          =   8421504
            Color10         =   255
            Color11         =   65280
            Color12         =   65535
            Color13         =   16711680
            Color14         =   16711935
            Color15         =   16776960
            Color18         =   12632319
            Color19         =   12640511
            Color20         =   12648447
            Color21         =   12648384
            Color22         =   16777152
            Color23         =   16761024
            Color24         =   16761087
            Color25         =   14737632
            Color26         =   8421631
            Color27         =   8438015
            Color28         =   8454143
            Color29         =   8454016
            Color30         =   16777088
            Color31         =   16744576
            Color32         =   16744703
            Color33         =   12632256
            Color34         =   255
            Color35         =   33023
            Color36         =   65535
            Color37         =   65280
            Color38         =   16776960
            Color39         =   16711680
            Color40         =   16711935
            Color41         =   8421504
            Color42         =   192
            Color43         =   16576
            Color44         =   49344
            Color45         =   49152
            Color46         =   12632064
            Color47         =   12582912
            Color48         =   12583104
            Color49         =   4210752
            Color50         =   128
            Color51         =   16512
            Color52         =   32896
            Color53         =   32768
            Color54         =   8421376
            Color55         =   8388608
            Color56         =   8388736
            Color57         =   0
            Color58         =   64
            Color59         =   4210816
            Color60         =   16448
            Color61         =   16384
            Color62         =   4210688
            Color63         =   4194304
            Color64         =   4194368
         End
         Begin MBColorPicker.ColorPicker CPicTextColor 
            Height          =   345
            Left            =   1455
            TabIndex        =   4
            Top             =   480
            Width           =   1140
            _ExtentX        =   1773
            _ExtentY        =   556
            CustomButtonText=   " Œ’Ì’"
            BackColor       =   14871017
            Style           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumColors       =   64
            Color1          =   0
            Color2          =   128
            Color3          =   32768
            Color4          =   32896
            Color5          =   8388608
            Color6          =   8388736
            Color7          =   8421376
            Color8          =   12632256
            Color9          =   8421504
            Color10         =   255
            Color11         =   65280
            Color12         =   65535
            Color13         =   16711680
            Color14         =   16711935
            Color15         =   16776960
            Color18         =   12632319
            Color19         =   12640511
            Color20         =   12648447
            Color21         =   12648384
            Color22         =   16777152
            Color23         =   16761024
            Color24         =   16761087
            Color25         =   14737632
            Color26         =   8421631
            Color27         =   8438015
            Color28         =   8454143
            Color29         =   8454016
            Color30         =   16777088
            Color31         =   16744576
            Color32         =   16744703
            Color33         =   12632256
            Color34         =   255
            Color35         =   33023
            Color36         =   65535
            Color37         =   65280
            Color38         =   16776960
            Color39         =   16711680
            Color40         =   16711935
            Color41         =   8421504
            Color42         =   192
            Color43         =   16576
            Color44         =   49344
            Color45         =   49152
            Color46         =   12632064
            Color47         =   12582912
            Color48         =   12583104
            Color49         =   4210752
            Color50         =   128
            Color51         =   16512
            Color52         =   32896
            Color53         =   32768
            Color54         =   8421376
            Color55         =   8388608
            Color56         =   8388736
            Color57         =   0
            Color58         =   64
            Color59         =   4210816
            Color60         =   16448
            Color61         =   16384
            Color62         =   4210688
            Color63         =   4194304
            Color64         =   4194368
         End
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   330
            Left            =   30
            TabIndex        =   7
            Top             =   720
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
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
         Begin ImpulseButton.ISButton XPBtnPrint 
            Height          =   330
            Left            =   30
            TabIndex        =   8
            Top             =   375
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   330
            Left            =   30
            TabIndex        =   9
            Top             =   30
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1035
            Index           =   2
            Left            =   3960
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   30
            Width           =   6750
            _cx             =   11906
            _cy             =   1826
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   192
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "»ÕÀ ⁄‰ ’‰›"
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
            Begin VB.ComboBox XPCboSearchType 
               Height          =   315
               Left            =   3240
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   255
               Width           =   3405
            End
            Begin VB.TextBox XPTxtSearchValue 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3270
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   600
               Width           =   3405
            End
            Begin VB.CheckBox XPChkFullMuch 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ·„… »«·ﬂ«„·"
               Height          =   285
               Left            =   1725
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   630
               Width           =   1470
            End
            Begin VB.CheckBox XPChkSearchType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ÿ«»ﬁ… Õ«·… «·√Õ—› "
               Height          =   285
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   240
               Width           =   2235
            End
            Begin ImpulseButton.ISButton XPBtnSearch 
               Height          =   285
               Left            =   90
               TabIndex        =   15
               Top             =   600
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Ê‰ «·Œ·›Ì…"
            Height          =   345
            Index           =   0
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   135
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Ê‰ «·‰’"
            Height          =   270
            Index           =   2
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   540
            Width           =   1245
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FgMain 
         Height          =   5760
         Left            =   15
         TabIndex        =   1
         Top             =   945
         Width           =   10770
         _cx             =   18997
         _cy             =   10160
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmMainPriceList.frx":03EF
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
         Height          =   915
         Index           =   1
         Left            =   15
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   15
         Width           =   10770
         _cx             =   18997
         _cy             =   1614
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
         Begin VB.Frame XPPnlSupplier 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   0
            Width           =   1650
            Begin ImpulseButton.ISButton XPBtnRemove 
               Height          =   315
               Left            =   90
               TabIndex        =   24
               Top             =   30
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   16777215
               ColorHighlight  =   4194304
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton XPBtnAdd 
               Height          =   315
               Left            =   930
               TabIndex        =   25
               Top             =   30
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   16777215
               ColorHighlight  =   4194304
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton XPBtnAddSupplier 
               Height          =   315
               Left            =   510
               TabIndex        =   26
               Top             =   30
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   16777215
               ColorHighlight  =   4194304
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
         End
         Begin VB.ComboBox XPCboMenuType 
            Height          =   315
            Left            =   7335
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   435
            Width           =   3330
         End
         Begin C1SizerLibCtl.C1Elastic XPPnlViewType 
            Height          =   930
            Left            =   3270
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   -60
            Width           =   4050
            _cx             =   7144
            _cy             =   1640
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
            Begin VB.OptionButton XPOptViewType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ‘Ã—… «·√’‰«›"
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   210
               Value           =   -1  'True
               Width           =   1680
            End
            Begin VB.OptionButton XPOptViewType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ÃœÊ·"
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   1
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   510
               Width           =   1680
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Left            =   120
               TabIndex        =   29
               Top             =   480
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
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
               DrawFocusRectangle=   0   'False
            End
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Œ›«¡ «·√—’œ… «·’›—Ì…"
               Height          =   315
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   150
               Width           =   2115
            End
         End
         Begin C1SizerLibCtl.C1Elastic XPPnlSuppliers 
            Height          =   795
            Left            =   1740
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   75
            Width           =   5580
            _cx             =   9843
            _cy             =   1402
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
            Begin MSDataListLib.DataCombo DBCboSupplierName 
               Height          =   315
               Left            =   135
               TabIndex        =   20
               ToolTipText     =   "«”„ «·„Ê—œ"
               Top             =   285
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ"
               ForeColor       =   &H000000FF&
               Height          =   315
               Index           =   1
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   330
               Width           =   960
            End
         End
         Begin ImpulseButton.ISButton XPBtnShowCol 
            Height          =   630
            Left            =   0
            TabIndex        =   27
            Top             =   120
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1111
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈Œ›«¡ Ê≈ŸÂ«— «·√⁄„œ…"
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
            ColorHighlight  =   4194304
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ — ‰Ê⁄ «·ﬁ«∆„… «·„—«œ ⁄—÷Â«"
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
            Height          =   300
            Index           =   3
            Left            =   7380
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   75
            Width           =   3285
         End
      End
   End
End
Attribute VB_Name = "FrmMainPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrintReport As ClsListPriceReport
Dim SupPriceReport As ClsSupplierPrice
Dim TTP As clstooltip
Dim cSearchDcbo As clsDCboSearch

Private Sub SetupGrid()
    Dim My_SQL As String
    Dim RsData As ADODB.Recordset
    Dim LngParentRow As Long
    Dim i As Integer
    Dim IntColName As Integer
    Dim BolRtl As Boolean
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    With Me.FgMain
        .Redraw = flexRDNone
        .Rows = 1

        If BolRtl = True Then
            IntColName = 1
            .AddItem "‘Ã—… «·√’‰«›"
        Else
            .AddItem "Items Tree"
            IntColName = 1
        End If

        .Rowdata(.Rows - 1) = "1G"
        .IsSubtotal(.Rows - 1) = True
        .Cell(flexcpFontBold, .Rows - 1, 1) = True
        .GridLines = flexGridFlat
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .NodeClosedPicture = mdifrmmain.ImgLstTree.ListImages("Closed_Node").Picture
        .NodeOpenPicture = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
        .RowHeightMin = 300
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        My_SQL = " SELECT Groups.GroupID, Groups.GroupName, Groups.ParentID " & "FROM Groups Where Groups.ParentID=1"
        Set RsData = New ADODB.Recordset
       
        RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call LoadGridTree("1G", RsData, FgMain, "Groups", "ParentID", "", , IntColName, vbBlue)

        Set RsData = New ADODB.Recordset
        RsData.Open "PriceListTree", Cn, adOpenStatic, adLockReadOnly, adCmdTable

        If Not (RsData.EOF Or RsData.BOF) Then
            RsData.MoveFirst
            Set RsTemp = New ADODB.Recordset
            StrSQL = "select * From ItemsPrice"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            Do While Not RsData.EOF
                LngParentRow = .FindRow(CStr(RsData("GroupID").value) & "G", 0, -1, False, True)

                If LngParentRow > 0 Then
                    .AddItem RsData(IntColName).value, (LngParentRow + 1)
                    .Rowdata((LngParentRow + 1)) = RsData("ItemID").value & "I"
                    .RowOutlineLevel((LngParentRow + 1)) = .RowOutlineLevel(LngParentRow) + 1
                    .Cell(flexcpPicture, LngParentRow + 1, 0) = mdifrmmain.ImgLstTree.ListImages("Item").Picture
                    'Put the Item Data In the grid
                
                    .TextMatrix(LngParentRow + 1, .ColIndex("ItemID")) = IIf(IsNull(RsData("ItemID").value), "", RsData("ItemID").value)
                    .TextMatrix(LngParentRow + 1, .ColIndex("ItemCode")) = IIf(IsNull(RsData("ItemCode").value), "", RsData("ItemCode").value)
                    .TextMatrix(LngParentRow + 1, .ColIndex("DefalutPrice")) = IIf(IsNull(RsData("SallingPrice").value), "0", RsData("SallingPrice").value)
                
                    .TextMatrix(LngParentRow + 1, .ColIndex("CustomerPrice")) = IIf(IsNull(RsData("CustomerPrice").value), "0", RsData("CustomerPrice").value)
                
                    .TextMatrix(LngParentRow + 1, .ColIndex("DealerPrice")) = IIf(IsNull(RsData("DealerPrice").value), "0", RsData("DealerPrice").value)
                
                    .TextMatrix(LngParentRow + 1, .ColIndex("qty")) = IIf(IsNull(RsData("qty").value), "0", RsData("qty").value)
                    .TextMatrix(LngParentRow + 1, .ColIndex("LastUpdate")) = IIf(IsNull(RsData("LastUpdate").value), "", Format(RsData("LastUpdate").value, "yyyy/m/d"))
                
                    ' „ÌÌ“ «·√’‰«› «· Ì ·Â« ‘—«∆Õ √”⁄«—
                    If RsData("ItemID").value <> "" Then
                        'StrSQL = "select * From ItemsPrice where Item_ID=" & RsData("ItemID").Value
                        'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        RsTemp.Filter = adFilterNone
                        RsTemp.Filter = "Item_ID=" & RsData("ItemID").value & ""

                        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                            .Cell(flexcpPicture, LngParentRow + 1, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
                        End If
                    End If
                End If

                RsData.MoveNext
            Loop

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignRightCenter
        Else
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignLeftCenter
        End If

        .AutoSize 0, .Cols - 1, False
        .Redraw = True
        .Outline 1
    End With

    RsData.Close
    RsTemp.Close
    Set RsData = Nothing
    Set RsTemp = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

End Sub

Private Sub Cmd_Click()
    On Error GoTo ErrTrap

    If Me.XPOptViewType(0).value = True Then
        SetupGrid
        FgMain.TextMatrix(0, FgMain.ColIndex("Tree")) = ""
    ElseIf Me.XPOptViewType(1).value = True Then
        TableShow
        FgMain.TextMatrix(0, FgMain.ColIndex("Tree")) = " «”„ «·’‰›"
    End If

    FgMain.AutoSize 0, FgMain.Cols - 1, False
    GetMeSetting
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CPicColColor_Change(ByVal NewColor As stdole.OLE_COLOR)
    Dim RowNum As Integer
    On Error GoTo ErrTrap
    FgMain.Cell(flexcpBackColor, 1, FgMain.Col, FgMain.Rows - 1, FgMain.Col) = CPicColColor.color

    If XPCboMenuType.ListIndex = 0 Then
        SaveMeSetting
    Else
        SaveSupPriceSetting
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CPicTextColor_Change(ByVal NewColor As stdole.OLE_COLOR)
    On Error GoTo ErrTrap
    FgMain.Cell(flexcpForeColor, 1, FgMain.Col, FgMain.Rows - 1, FgMain.Col) = CPicTextColor.color

    If XPCboMenuType.ListIndex = 0 Then
        SaveMeSetting
    Else
        SaveSupPriceSetting
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboSupplierName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim rs As ADODB.Recordset
    Dim RsPrice As ADODB.Recordset

    If DBCboSupplierName.text <> "" Then
        XPPnlSupplier.Enabled = True
        FgMain.Clear flexClearScrollable, flexClearEverything
        FgMain.Rows = 2
        StrSQL = "SELECT CusJuncItem.ID,CusJuncItem.LastUpdate, CusJuncItem.CusID, CusJuncItem.ItemID, " & "CusJuncItem.ItemPrice, TblItems.ItemCode, TblItems.ItemName FROM TblItems " & " INNER JOIN CusJuncItem ON TblItems.ItemID = CusJuncItem.ItemID where CusID = " & DBCboSupplierName.BoundText
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.EOF Or rs.BOF) Then
            FgMain.Rows = rs.RecordCount + 1

            With FgMain

                For RowNum = 1 To rs.RecordCount
                    .Rowdata(RowNum) = rs("ID").value
                    .TextMatrix(RowNum, .ColIndex("Tree")) = rs("ItemName").value
                    .TextMatrix(RowNum, .ColIndex("DefalutPrice")) = rs("ItemPrice").value
                    .TextMatrix(RowNum, .ColIndex("ItemID")) = rs("ItemID").value
                    .TextMatrix(RowNum, .ColIndex("ItemCode")) = rs("ItemCode").value
                    .TextMatrix(RowNum, .ColIndex("LastUpdate")) = IIf(IsNull(rs("LastUpdate").value), "", Format(rs("LastUpdate").value, "yyyy/mm/dd"))
                    ' „ÌÌ“ «·√’‰«› «· Ì ·Â« ‘—«∆Õ √”⁄«—
                    Set RsPrice = New ADODB.Recordset

                    If rs("ItemID").value <> "" Then
                        StrSQL = "select * From JuncPrice where juncID=" & rs("ItemID").value
                        RsPrice.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                        If Not (RsPrice.EOF Or RsPrice.BOF) Then
                            .Cell(flexcpPicture, RowNum, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
                        End If

                        RsPrice.Close
                    End If

                    rs.MoveNext
                Next RowNum

            End With

        End If

    Else
        XPPnlSupplier.Enabled = False
    End If

    GetSupPriceSetting
    Exit Sub
ErrTrap:
End Sub

Private Sub EleMain_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FgMain_Click()
    On Error GoTo ErrTrap

    With FgMain
        CPicColColor.color = .Cell(flexcpBackColor, .Rows - 1, .Col)
        CPicTextColor.color = .Cell(flexcpForeColor, .Rows - 1, .Col)
    End With

    Exit Sub
ErrTrap:
End Sub

Public Sub FgMain_DblClick()
    On Error GoTo ErrTrap
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim XGrdNode As VSFlexNode

    With Me.FgMain

        Select Case XPCboMenuType.ListIndex

            Case 0

                If .Row = -1 Then Exit Sub
                If .Col = -1 Then Exit Sub

                Select Case .Col

                    Case .ColIndex("Tree")

                        If .IsSubtotal(.Row) = True Then
                            Set XGrdNode = .GetNode(.Row)

                            If Not (XGrdNode Is Nothing) Then
                                XGrdNode.Expanded = Not XGrdNode.Expanded
                            End If
                        End If

                    Case .ColIndex("DefalutPrice")

                        If XPOptViewType(0).value = True Then
                            If .Rowdata(.Row) <> "" Then
                                If right(.Rowdata(.Row), 1) = "I" Then

                                    With FgMain
                                        FrmItemsPrice.XPLblItemName.Caption = .TextMatrix(.Row, .ColIndex("tree"))
                                        FrmItemsPrice.TxtQty.text = .TextMatrix(.Row, .ColIndex("Qty"))
                                        FrmItemsPrice.XPLblItemCode.Caption = .TextMatrix(.Row, .ColIndex("ItemCode"))
                                        FrmItemsPrice.XPTxtPrice.text = .TextMatrix(.Row, .ColIndex("DefalutPrice"))
                                        FrmItemsPrice.TxtCustomerPrice.text = .TextMatrix(.Row, .ColIndex("CustomerPrice"))
                                        FrmItemsPrice.TxtDealerPrice.text = .TextMatrix(.Row, .ColIndex("DealerPrice"))
                                        FrmItemsPrice.TxtCompareValue.text = .TextMatrix(.Row, .ColIndex("DefalutPrice"))
                                        FrmItemsPrice.XPLblItemID.Caption = left(.Rowdata(.Row), (Len(.Rowdata(.Row)) - 1))
                                    End With

                                    FrmItemsPrice.show vbModal
                                End If
                            End If

                        ElseIf XPOptViewType(1).value = True Then

                            If .Rowdata(.Row) <> "" Then
                                FrmItemsPrice.XPLblItemName.Caption = .TextMatrix(.Row, .ColIndex("tree"))
                                FrmItemsPrice.TxtQty.text = .TextMatrix(.Row, .ColIndex("Qty"))
                                FrmItemsPrice.XPLblItemCode.Caption = .TextMatrix(.Row, .ColIndex("ItemCode"))
                                FrmItemsPrice.XPTxtPrice.text = .TextMatrix(.Row, .ColIndex("DefalutPrice"))
                                FrmItemsPrice.TxtCustomerPrice.text = .TextMatrix(.Row, .ColIndex("CustomerPrice"))
                                FrmItemsPrice.TxtDealerPrice.text = .TextMatrix(.Row, .ColIndex("DealerPrice"))
                                FrmItemsPrice.TxtCompareValue.text = .TextMatrix(.Row, .ColIndex("DefalutPrice"))
                                FrmItemsPrice.XPLblItemID.Caption = .Rowdata(.Row)
                                FrmItemsPrice.show vbModal
                            End If
                        End If

                    Case .ColIndex("Qty")

                        If XPOptViewType(0).value = True Then
                            If right(.Rowdata(.Row), 1) = "I" Then
                                If .TextMatrix(.Row, .ColIndex("ItemID")) <> "" Then
                                    StrSQL = "select * From TblItems where ItemID=" & .TextMatrix(.Row, .ColIndex("ItemID"))
                                    Set RsTemp = New ADODB.Recordset
                                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
                                    FrmSearchSerial.Tag = RsTemp("ItemCode").value
                                    FrmSearchSerial.Txt.text = "PriceList"
                                    FrmSearchSerial.show vbModal
                                    RsTemp.Close
                                End If
                            End If

                        ElseIf XPOptViewType(1).value = True Then

                            If .Rowdata(.Row) <> "" Then
                                StrSQL = "select * From TblItems where ItemID=" & .TextMatrix(.Row, .ColIndex("ItemID"))
                                Set RsTemp = New ADODB.Recordset
                                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
                                FrmSearchSerial.Tag = RsTemp("ItemCode").value
                                FrmSearchSerial.Txt.text = "PriceList"
                                FrmSearchSerial.show vbModal
                                RsTemp.Close
                            End If
                        End If

                End Select

            Case 1

                If .Row = 0 Or .Col = -1 Then Exit Sub
                If .TextMatrix(.Row, .ColIndex("Tree")) = "" Then Exit Sub
                If .TextMatrix(.Row, .Col) <> "" Then
               '     FrmSup_ItemsPrice.XPLblSupName.Caption = DBCboSupplierName.text
               '     FrmSup_ItemsPrice.XPLblItemName.Caption = .Cell(flexcpTextDisplay, .Row, .ColIndex("tree"))
               '     FrmSup_ItemsPrice.XPTxtPrice.text = .TextMatrix(.Row, .ColIndex("DefalutPrice"))
               ''     FrmSup_ItemsPrice.TxtCompareValue.text = .TextMatrix(.Row, .ColIndex("DefalutPrice"))
                '    FrmSup_ItemsPrice.XPLblPriceID.Caption = .Rowdata(.Row)
                '    FrmSup_ItemsPrice.show vbModal
                End If

        End Select

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub FgMain_MouseUp(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    On Error GoTo ErrTrap
    Dim tp            As POINTAPI
    Dim lX            As Single
    Dim lY            As Single
    Dim tr            As RECT

    If XPCboMenuType.ListIndex = 0 Then
        If XPOptViewType(0).value = True Then
            If right(FgMain.Rowdata(FgMain.Row), 1) = "I" Then
                '            XPPopUp.Menus(1).MenuItems(1).Enabled = True
                '            XPPopUp.Menus(1).MenuItems(2).Enabled = True
                mdifrmmain.ShowItems.Enabled = True
                mdifrmmain.ItemsPrice.Enabled = True
            Else
                mdifrmmain.ShowItems.Enabled = False
                mdifrmmain.ItemsPrice.Enabled = False
            End If

        ElseIf XPOptViewType(1).value = True Then

            If FgMain.Rowdata(FgMain.Row) <> "" Then
                mdifrmmain.ShowItems.Enabled = True
                mdifrmmain.ItemsPrice.Enabled = True
            Else
                mdifrmmain.ShowItems.Enabled = False
                mdifrmmain.ItemsPrice.Enabled = False
            End If
        End If

        If Button = vbRightButton Then
            GetCursorPos tp
            lX = (tp.X) * Screen.TwipsPerPixelX
            lY = tp.Y * Screen.TwipsPerPixelY
            '        XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
            PopupMenu mdifrmmain.PriceListPop, vbPopupMenuRightAlign, X, Y + 200
        End If

    Else

        If FgMain.Rowdata(FgMain.Row) <> "" Then
            mdifrmmain.AddItem.Enabled = True
            mdifrmmain.DelItem.Enabled = True
            mdifrmmain.PriceChips.Enabled = True
        '    mdifrmmain.PriceOffer.Enabled = True
        Else
            mdifrmmain.AddItem.Enabled = False
            mdifrmmain.DelItem.Enabled = False
            mdifrmmain.PriceChips.Enabled = False
           ' mdifrmmain.PriceOffer.Enabled = False
        End If

        If Button = vbRightButton Then
            GetCursorPos tp
            lX = (tp.X) * Screen.TwipsPerPixelX
            lY = tp.Y * Screen.TwipsPerPixelY
            PopupMenu mdifrmmain.SupList, vbPopupMenuRightAlign, X, Y + 200
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF6 Then
        XPBtnPrint_Click
    End If

End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap

    Set Me.CmdHelp.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Help2").Picture
    Set XPBtnPrint.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Print").Picture
    Set CmdExit.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Exit").Picture
    Set Me.XPBtnSearch.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Find").Picture

    Set Me.XPBtnAdd.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("AddUser").Picture
    Set Me.XPBtnAddSupplier.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Find").Picture
    Set Me.XPBtnRemove.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Find").Picture

    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboSupplierName, False
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DBCboSupplierName
    Set Me.Cmd.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Refresh").Picture

    With FgMain
        .Cell(flexcpPicture, 0, .ColIndex("ItemID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("ItemCode")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Qty")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Tree")) = mdifrmmain.ImgLstTree.ListImages("Item").Picture
        .Cell(flexcpPicture, 0, .ColIndex("LastUpdate")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
        Else
            .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignLeftCenter
        End If

    End With

    XPCboSearchType.Clear

    If SystemOptions.UserInterface = ArabicInterface Then
        XPCboSearchType.AddItem "«”„ «·’‰› "
        XPCboSearchType.AddItem "ﬂÊœ «·’‰› "
    Else
        XPCboSearchType.AddItem "Item Name"
        XPCboSearchType.AddItem "Item Code"
    End If

    XPCboSearchType.ListIndex = 0

    XPCboMenuType.Clear

    If SystemOptions.UserInterface = ArabicInterface Then
        XPCboMenuType.AddItem "ﬁ«∆„… «·√”⁄«— «·⁄«„…"
        XPCboMenuType.AddItem "ﬁ«∆„… √”⁄«— „Ê—œ"
    Else
        XPCboMenuType.AddItem "Main Items Price List"
        XPCboMenuType.AddItem "Supplier Items Price List"
    End If

    XPCboMenuType.ListIndex = 0
    'GetViewMode
    AddTip
    Exit Sub
ErrTrap:
End Sub

Public Sub SaveMeSetting()
    'ViewType As Integer
    On Error GoTo ErrTrap
    FgMain.Row = 1
    FgMain.Col = FgMain.ColIndex("Tree")
    SaveSetting StrAppRegPath, "ListColor", "TreeColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "TreeForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("ItemID")
    SaveSetting StrAppRegPath, "ListColor", "IDColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "IDForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("ItemCode")
    SaveSetting StrAppRegPath, "ListColor", "CodeColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "CodeForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("Qty")
    SaveSetting StrAppRegPath, "ListColor", "QtyColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "QtyForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("DefalutPrice")
    SaveSetting StrAppRegPath, "ListColor", "PriceColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "PriceForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("LastUpdate")
    SaveSetting StrAppRegPath, "ListColor", "LastUpdateColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "LastUpdateForeColor", FgMain.CellForeColor
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMeSetting
    SaveFontSetting
    Set cSearchDcbo = Nothing
    SaveViewMode
End Sub

Public Sub GetMeSetting()
    On Error GoTo ErrTrap
    Dim ColColor As String
    Dim RowNum As Integer
    Dim LngPriceColColor As Long
    Dim LngQtyColColor As Long
    Dim ShowCol As Boolean

    '√·Ê«‰ «·√⁄„œ…
    With FgMain
        ColColor = GetSetting(StrAppRegPath, "ListColor", "TreeColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "TreeForeColor", "0 ")
        .Cell(flexcpForeColor, 1, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "IDColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("ItemID"), .Rows - 1, .ColIndex("ItemID")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "IDForeColor", "0 ")
        .Cell(flexcpForeColor, 1, .ColIndex("ItemID"), .Rows - 1, .ColIndex("ItemID")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "CodeColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("ItemCode"), .Rows - 1, .ColIndex("ItemCode")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "CodeForeColor", "0 ")
        .Cell(flexcpForeColor, 1, .ColIndex("ItemCode"), .Rows - 1, .ColIndex("ItemCode")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "QtyColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("Qty"), .Rows - 1, .ColIndex("Qty")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "QtyForeColor", "0 ")
        .Cell(flexcpForeColor, 1, .ColIndex("Qty"), .Rows - 1, .ColIndex("Qty")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "PriceColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("DefalutPrice"), .Rows - 1, .ColIndex("DefalutPrice")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "PriceForeColor", "0 ")
        .Cell(flexcpForeColor, 1, .ColIndex("DefalutPrice"), .Rows - 1, .ColIndex("DefalutPrice")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "LastUpdateColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("LastUpdate"), .Rows - 1, .ColIndex("LastUpdate")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "LastUpdateForeColor", "0 ")
        .Cell(flexcpForeColor, 1, .ColIndex("LastUpdate"), .Rows - 1, .ColIndex("LastUpdate")) = val(ColColor)
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
    End With

    '«ŸÂ«— «·√⁄„œ…
    With FgMain
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowItemID", True)
        .ColHidden(FgMain.ColIndex("ItemID")) = Not (ShowCol)
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowItemCode", True)
        .ColHidden(FgMain.ColIndex("ItemCode")) = Not (ShowCol)
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowQty", True)
        .ColHidden(FgMain.ColIndex("Qty")) = Not (ShowCol)
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowDefalutPrice", True)
        .ColHidden(FgMain.ColIndex("DefalutPrice")) = Not (ShowCol)
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowLastUpdate", True)
        .ColHidden(FgMain.ColIndex("LastUpdate")) = Not (ShowCol)
    
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowCustomerPrice", True)
        .ColHidden(FgMain.ColIndex("CustomerPrice")) = Not (ShowCol)
    
        ShowCol = GetSetting(StrAppRegPath, "ShowCol", "ShowDealerPrice", True)
        .ColHidden(FgMain.ColIndex("DealerPrice")) = Not (ShowCol)
    
    End With

    ' ‰”Ìﬁ  «·Œÿ
    FgMain.FontBold = GetSetting(StrAppRegPath, "GridSetting", "FontBold", FgMain.FontBold)
    FgMain.FontItalic = GetSetting(StrAppRegPath, "GridSetting", "FontItalic", FgMain.FontItalic)
    FgMain.FontName = GetSetting(StrAppRegPath, "GridSetting", "FontName", FgMain.FontName)
    FgMain.fontsize = GetSetting(StrAppRegPath, "GridSetting", "FontSize", FgMain.fontsize)
    Exit Sub
ErrTrap:
End Sub

Public Sub XPBtnAdd_Click()
    Dim Msg As String

    On Error GoTo ErrTrap

    If DBCboSupplierName.text <> "" Then
        FrmAddItemsPriceList.TxtSupID.text = DBCboSupplierName.BoundText
        FrmAddItemsPriceList.XPLblSupName.Caption = DBCboSupplierName.text
        FrmAddItemsPriceList.show vbModal
    Else
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnAddSupplier_Click()
    On Error GoTo ErrTrap

    'With FrmAddNewCustemer
        '    .Tag = "xxxx"
    '    .DealingForm = PriceList
    '    .show vbModal
    '    .Caption = "≈÷«›… „Ê—œ ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·„Ê—œ"
    '    .lbl(0).Caption = "«”„ «·„Ê—œ"
    'End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnPrint_Click()

    Dim Msg As String
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    On Error GoTo ErrTrap

    Select Case XPCboMenuType.ListIndex

        Case 0

            If XPOptViewType(0).value = True Then
                Set PrintReport = New ClsListPriceReport
                PrintReport.ShowListPrice 1
            ElseIf XPOptViewType(1).value = True Then
                Set PrintReport = New ClsListPriceReport
                PrintReport.ShowListPrice 0
            End If

        Case 1

            If DBCboSupplierName.BoundText = "" Then
                Msg = "Õœœ «”„ «·„Ê—œ «·–Ì  —€» ›Ì ÿ»«⁄… " & Chr(13)
                Msg = Msg + "ﬁ«∆„… √”⁄«— «·√’‰«› «·Œ«’… »Â"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            Else
                Set SupPriceReport = New ClsSupplierPrice
                SupPriceReport.SupplierPrice DBCboSupplierName.BoundText
            End If

    End Select

    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«ı ·«Ì„ﬂ‰ «·ÿ»«⁄… ....!!!"
    Msg = Msg & Chr(13) & "Err.Description:" & Err.description
    Msg = Msg & Chr(13) & "Err.Number:" & Err.Number
    Msg = Msg & Chr(13) & "Err.Source:" & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap
    Dim StrSQL As String

    If FgMain.Rowdata(FgMain.Row) = "" Then
        Exit Sub
    End If

    StrSQL = "Delete  From CusJuncItem where ID=" & FgMain.Rowdata(FgMain.Row)
    Cn.Execute StrSQL

    With FgMain

        If .Rows > 1 Then
            If .Rows = 2 Then
                .Clear flexClearScrollable, flexClearEverything
            Else

                If .Rows > 1 Then
                    If .Row <> .FixedRows - 1 Then
                        .RemoveItem (.Row)
                    End If
                End If
            End If
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnSearch_Click()
    On Error GoTo ErrTrap
    Dim XGrdNode As VSFlexNode
    Dim LngFindRow As Long
    Dim Msg As String

    If XPCboSearchType.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «·Õﬁ· " & Chr(13)
        Msg = Msg + "«·–Ì ”Ì „ «·»ÕÀ ›ÌÂ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPCboSearchType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If XPTxtSearchValue.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «· Ì  —€» ›Ì «·»ÕÀ ⁄‰Â« "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtSearchValue.SetFocus
        Exit Sub
    End If

    If XPChkFullMuch.value = Checked Then
        If XPChkSearchType.value = Checked Then

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("Tree"), True, True)

                Case 1
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("ItemCode"), True, True)
            End Select

        Else

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("Tree"), False, True)

                Case 1
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("ItemCode"), False, True)
            End Select

        End If

    Else

        If XPChkSearchType.value = Checked Then

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("Tree"), True, False)

                Case 1
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("ItemCode"), True, False)
            End Select

        Else

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("Tree"), False, False)

                Case 1
                    LngFindRow = FgMain.FindRow(XPTxtSearchValue.text, FgMain.FixedRows, FgMain.ColIndex("ItemCode"), False, False)
            End Select

        End If
    End If

    If LngFindRow > 0 Then
        If XPCboMenuType.ListIndex = 0 Then
            If XPOptViewType(0).value = True Then
                Set XGrdNode = FgMain.GetNode(LngFindRow)

                If Not XGrdNode Is Nothing Then
                    XGrdNode.Expanded = True
                End If
            End If
        End If

        FgMain.Row = LngFindRow
        FgMain.ShowCell LngFindRow, FgMain.ColIndex("Tree")
    Else
        FgMain.Row = 1
        Msg = "·„ Ì „ «·⁄ÀÊ— ⁄·Ï Â–« «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnShowCol_Click()
    On Error GoTo ErrTrap

    With FrmShowCol.FG
        .TextMatrix(0, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("ItemID")))
        .TextMatrix(1, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("ItemCode")))
        .TextMatrix(2, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("Qty")))
        .TextMatrix(3, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("DefalutPrice")))
        .TextMatrix(4, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("LastUpdate")))
        .TextMatrix(5, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("CustomerPrice")))
        .TextMatrix(6, .ColIndex("show")) = Not (Me.FgMain.ColHidden(FgMain.ColIndex("DealerPrice")))
    End With

    FrmShowCol.show vbModal
    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboMenuType_Change()
    On Error GoTo ErrTrap

    With FgMain

        If XPCboMenuType.ListIndex = 0 Then
            XPPnlSuppliers.Visible = False
            XPPnlViewType.Visible = True
            .ColHidden(.ColIndex("ItemID")) = False
            .ColHidden(.ColIndex("ItemCode")) = False
            .ColHidden(.ColIndex("Qty")) = False
            '    .ColHidden(.ColIndex("LastUpdate")) = False
            XPPnlSupplier.Visible = False
            XPBtnShowCol.Visible = True
            GetMeSetting
            GetViewMode
            'XPFrmSearch.Enabled = True
        Else
            SaveViewMode
            XPPnlSuppliers.Visible = True
            XPPnlViewType.Visible = False
            .ColHidden(.ColIndex("ItemID")) = False
            .ColHidden(.ColIndex("ItemCode")) = False
            .ColHidden(.ColIndex("Qty")) = True
            '    .ColHidden(.ColIndex("LastUpdate")) = True
            .ColHidden(.ColIndex("DefalutPrice")) = False
            XPPnlSupplier.Visible = True
            DBCboSupplierName.text = ""
            XPBtnShowCol.Visible = False
            FgMain.Clear flexClearScrollable, flexClearEverything
            FgMain.Rows = 1
            FgMain.Rows = 2
            GetSupPriceSetting
            XPCboSearchType.ListIndex = 0
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboMenuType_Click()
    XPCboMenuType_Change
End Sub

Private Sub XPOptViewType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            XPOptViewType(0).value = True
            XPOptViewType(1).value = False
            SetupGrid
            FgMain.TextMatrix(0, FgMain.ColIndex("Tree")) = " "

        Case 1
            XPOptViewType(1).value = True
            XPOptViewType(0).value = False
            TableShow
            FgMain.TextMatrix(0, FgMain.ColIndex("Tree")) = " «”„ «·’‰›"
    End Select

    FgMain.AutoSize 0, FgMain.Cols - 1, False
    GetMeSetting
    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtSearchValue_KeyDown(KeyCode As Integer, _
                                     Shift As Integer)

    If KeyCode = vbKeyReturn Then
        XPBtnSearch_Click
    End If

End Sub

Private Sub TableShow()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim RsPrice As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Set RsTemp = New ADODB.Recordset

    StrSQL = "Select * From PriceListTree"

    If Me.Chk.value = vbChecked Then
        StrSQL = StrSQL + " Where Qty > 0 "
    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsTemp.EOF Or RsTemp.BOF Then
        RsTemp.Close
        Set RsTemp = Nothing
        Exit Sub
    End If

    Set RsPrice = New ADODB.Recordset
    StrSQL = "Select * From ItemsPrice"
    RsPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Screen.MousePointer = vbArrowHourglass

    With FgMain
        .Redraw = flexRDNone
        FgMain.Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        .OutlineBar = flexOutlineBarNone
        .ScrollTips = False
        FgMain.Rows = FgMain.FixedRows + RsTemp.RecordCount

        For RowNum = 1 To RsTemp.RecordCount
            .Rowdata(RowNum) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
            .TextMatrix(RowNum, .ColIndex("Tree")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
            .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
            .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("ItemCode").value), "", RsTemp("ItemCode").value)
            .TextMatrix(RowNum, .ColIndex("DefalutPrice")) = IIf(IsNull(RsTemp("SallingPrice").value), "0", RsTemp("SallingPrice").value)
        
            .TextMatrix(RowNum, .ColIndex("CustomerPrice")) = IIf(IsNull(RsTemp("CustomerPrice").value), "0", RsTemp("CustomerPrice").value)
            .TextMatrix(RowNum, .ColIndex("DealerPrice")) = IIf(IsNull(RsTemp("DealerPrice").value), "0", RsTemp("DealerPrice").value)
        
            .TextMatrix(RowNum, .ColIndex("qty")) = IIf(IsNull(RsTemp("qty").value), "0", RsTemp("qty").value)
            .TextMatrix(RowNum, .ColIndex("LastUpdate")) = IIf(IsNull(RsTemp("LastUpdate").value), "", Format(RsTemp("LastUpdate").value, "yyyy/m/d"))

            If SystemOptions.UserInterface = ArabicInterface Then
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
            Else
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignLeftCenter
            End If

            ' „ÌÌ“ «·√’‰«› «· Ì ·Â« ‘—«∆Õ √”⁄«—
            If RsTemp("ItemID").value <> "" Then
                RsPrice.Filter = adFilterNone
                RsPrice.Filter = "Item_ID=" & RsTemp("ItemID").value

                If Not (RsPrice.EOF Or RsPrice.BOF) Then
                    .Cell(flexcpPicture, RowNum, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
                End If
            End If

            RsTemp.MoveNext
        Next RowNum

        .AutoSize 0, .Cols - 1, False
        .Redraw = True
    End With

    RsPrice.Close
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveViewMode()
    SaveSetting StrAppRegPath, "ViewType", "ViewTree", XPOptViewType(0).value
End Sub

Private Sub GetViewMode()
    On Error GoTo ErrTrap
    Dim ViewMode As Boolean
    ViewMode = GetSetting(StrAppRegPath, "ViewType", "ViewTree", True)

    If ViewMode = True Then
        XPOptViewType_Click (0)
    Else
        XPOptViewType_Click (1)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl FgMain, "ﬁ«∆„… «·√”⁄«— «·Œ«’… »«·√’‰«› ...", True
    End With

    'With TTP
    '    .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
    '    .MaxWidth = 4000
    '    .VisibleTime = 9000
    '    .DelayTime = 600
    '    .AddControl CPicColColor, _
    '    " €ÌÌ— √·Ê«‰ «·√⁄„œ…..." & Wrap & _
    '    "Õœœ Œ·Ì… ›Ì «·⁄„Êœ «·Ì  —€» ›Ì  €ÌÌ— ·Ê‰Â" & Wrap & _
    '    " À„ ﬁ„ »«Œ Ì«— «··Ê‰ «·–Ì  —€» ›Ì «” Œœ«„Â", True
    'End With
    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnPrint, "ÿ»«⁄…..." & Wrap & "·ÿ»«⁄…  ﬁ—Ì— »ﬁ«∆„… «·√”⁄«—" & Wrap & " «÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPCboSearchType, "Õﬁ· «·»ÕÀ ..." & Wrap & " Õœœ «·Õﬁ· «·–Ì  —€» ›Ì «·»ÕÀ »œ«Œ·Â  ", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPTxtSearchValue, "ﬁÌ„… «·»ÕÀ ..." & Wrap & "»⁄œ  ÕœÌœ  «·Õﬁ· «·–Ì  —€» ›Ì «·»ÕÀ »œ«Œ·Â" & Wrap & " «ﬂ » Â‰« «·ﬁÌ„… «· Ì  —€» ›Ì «·»ÕÀ ⁄‰Â«" & Wrap & "Enter À„ «÷€ÿ „› «Õ" & Wrap & "√Ê «÷€ÿ „› «Õ »ÕÀ", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnSearch, "»ÕÀ  ..." & Wrap & "»⁄œ  ÕœÌœ «·Õﬁ· «·–Ì  —€» ›Ì «·»ÕÀ »œ«Œ·Â" & Wrap & "Êﬂ «»… «·ﬁÌ„… «· Ì  —€» ›Ì «·»ÕÀ ⁄‰Â«" & Wrap & " ≈÷€ÿ Â‰« · ‰›Ì– ⁄„·Ì… «·»ÕÀ", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPChkSearchType, "„ÿ«»ﬁ… Õ«·… «·√Õ—›  ..." & Wrap & "·· „ÌÌ“ »Ì‰ «·Õ—Ê› «·ﬂ»Ì—… Ê«·Õ—Ê› «·’€Ì—…", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnShowCol, "≈Œ›«¡ Ê≈ŸÂ«— «·√⁄„œ…   ..." & Wrap & "·· Õﬂ„ ›Ì ≈ŸÂ«— Ê≈Œ›«¡ «·√⁄„œ… «·„⁄—Ê÷…", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPCboMenuType, "··«Œ Ì«— »Ì‰ ﬁ«∆„… «·√”⁄«— «·⁄«„…" & Wrap & "Êﬁ«∆„… «·√”⁄«— «·Œ«’… »„Ê—œ „⁄Ì‰", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl DBCboSupplierName, "«Œ Ì«— «”„ «·„Ê—œ " & Wrap & "«·–Ì ‰—€» ›Ì ⁄—÷ ﬁ«∆„… «·√”⁄«— «·Œ«’… »Â", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPOptViewType(0), "⁄—÷ ‘Ã—… «·√’‰«› " & Wrap & "⁄—÷ ﬁ«∆„… √”⁄«— «·√’‰«› ›Ì ‘ﬂ· ‘Ã—Ì", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPOptViewType(1), "⁄—÷ ÃœÊ· " & Wrap & "⁄—÷ ﬁ«∆„… √”⁄«— «·√’‰«› ›Ì ‘ﬂ· ÃœÊ·", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnAddSupplier, "≈÷«›… »Ì«‰«  „Ê—œ ÃœÌœ ", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnAdd, "≈÷«›… √”⁄«— «·√’‰«› «·Œ«’… »«·„Ê—œ «·„Õœœ ", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnRemove, "Õ–› ’‰› „‰ ﬁ«∆„… «·„Ê—œ " & Wrap & "Õœœ «·’‰› ›Ì «·ﬁ«∆„… Ê«÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "ﬁ«∆„… «·√”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdExit, "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    Exit Sub
ErrTrap:
End Sub

Public Sub SaveSupPriceSetting()
    On Error GoTo ErrTrap
    FgMain.Row = 1
    FgMain.Col = FgMain.ColIndex("Tree")
    SaveSetting StrAppRegPath, "ListColor", "ItemColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "ItemForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("DefalutPrice")
    SaveSetting StrAppRegPath, "ListColor", "SupPriceColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "SupPriceForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("ItemID")
    SaveSetting StrAppRegPath, "ListColor", "SupIDColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "SupIDForeColor", FgMain.CellForeColor
    FgMain.Col = FgMain.ColIndex("ItemCode")
    SaveSetting StrAppRegPath, "ListColor", "SupCODEColor", FgMain.CellBackColor
    SaveSetting StrAppRegPath, "ListColor", "SupCODEForeColor", FgMain.CellForeColor
    Exit Sub
ErrTrap:
End Sub

Public Sub GetSupPriceSetting()
    On Error GoTo ErrTrap
    Dim ColColor As String

    With FgMain
        ColColor = GetSetting(StrAppRegPath, "ListColor", "ItemColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "ItemForeColor", "0")
        .Cell(flexcpForeColor, 1, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "SupPriceColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("DefalutPrice"), .Rows - 1, .ColIndex("DefalutPrice")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "SupPriceForeColor", "0")
        .Cell(flexcpForeColor, 1, .ColIndex("DefalutPrice"), .Rows - 1, .ColIndex("DefalutPrice")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "SupIDColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("ItemID"), .Rows - 1, .ColIndex("ItemID")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "SupIDForeColor", "0")
        .Cell(flexcpForeColor, 1, .ColIndex("ItemID"), .Rows - 1, .ColIndex("ItemID")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "SupCODEColor", "16777215 ")
        .Cell(flexcpBackColor, 1, .ColIndex("ItemCode"), .Rows - 1, .ColIndex("ItemCode")) = val(ColColor)
        ColColor = GetSetting(StrAppRegPath, "ListColor", "SupCODEForeColor", "0")
        .Cell(flexcpForeColor, 1, .ColIndex("ItemCode"), .Rows - 1, .ColIndex("ItemCode")) = val(ColColor)
    End With

    Exit Sub
ErrTrap:
End Sub

Public Sub SaveFontSetting()
    On Error GoTo ErrTrap:
    'Font properties
    SaveSetting StrAppRegPath, "GridSetting", "FontBold", FgMain.FontBold
    SaveSetting StrAppRegPath, "GridSetting", "FontItalic", FgMain.FontItalic
    SaveSetting StrAppRegPath, "GridSetting", "FontName", FgMain.FontName
    SaveSetting StrAppRegPath, "GridSetting", "FontSize", FgMain.fontsize
    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Items Price List"
    Me.lbl(1).Caption = "Supplier Name"
    Me.lbl(3).Caption = "Choose your List"
    XPOptViewType(0).Caption = "Tree View"
    XPOptViewType(1).Caption = "Table View"
    Chk.Caption = "Hide Zero Quantity"
    Cmd.Caption = "Refresh"
    Ele(2).Caption = "Search For Item"
    XPChkSearchType.Caption = "Case Sensitive"
    XPChkFullMuch.Caption = "Full Match"
    XPBtnSearch.Caption = "&Search"
    lbl(0).Caption = "Background Color"
    lbl(2).Caption = "Fore COlor"
    CmdHelp.Caption = "&Help"
    XPBtnPrint.Caption = "&Print"
    CmdExit.Caption = "E&xit"

    With Me.FgMain
        .Cell(flexcpText, 0, .ColIndex("ItemID")) = "Item ID"
        .Cell(flexcpText, 0, .ColIndex("ItemCode")) = "Item Code"
        .Cell(flexcpText, 0, .ColIndex("Qty")) = "Quantity"
        .Cell(flexcpText, 0, .ColIndex("DefalutPrice")) = "Saling Price"
        .Cell(flexcpText, 0, .ColIndex("DealerPrice")) = "Dealer Price"
        .Cell(flexcpText, 0, .ColIndex("CustomerPrice")) = "Customer Price"
        .Cell(flexcpText, 0, .ColIndex("Tree")) = ""
        .Cell(flexcpText, 0, .ColIndex("LastUpdate")) = "Last Update Date"
    End With

    XPBtnShowCol.Caption = "Show && Hide Columns"
    XPBtnShowCol.Font.Bold = True
End Sub
