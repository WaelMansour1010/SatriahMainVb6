VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAccEditJournal2 
   Caption         =   " Õ—Ì— ﬁÌœ ÌÊ„Ì…"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15885
   HelpContextID   =   450
   Icon            =   "FrmAccEditJournal2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   15885
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8685
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15885
      _cx             =   28019
      _cy             =   15319
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
      BackColor       =   14737632
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
      _GridInfo       =   $"FrmAccEditJournal2.frx":030A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic EleTop 
         Height          =   660
         Left            =   15
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   15
         Width           =   15855
         _cx             =   27966
         _cy             =   1164
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   20.25
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
         BackColor       =   65280
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   " Õ—Ì— ﬁÌœ ÌÊ„Ì…"
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
         Style           =   0
         TagSplit        =   2
         PicturePos      =   7
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1245
            TabIndex        =   44
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal2.frx":038B
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
            Left            =   180
            TabIndex        =   45
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal2.frx":0725
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
            Left            =   1770
            TabIndex        =   46
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal2.frx":0ABF
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
            Left            =   705
            TabIndex        =   47
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal2.frx":0E59
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSAdodcLib.Adodc numbering 
            Height          =   585
            Left            =   2880
            Top             =   0
            Visible         =   0   'False
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   1032
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   " Õ—Ìﬂ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc detect_no 
            Height          =   585
            Left            =   1680
            Top             =   0
            Visible         =   0   'False
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   1032
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   " Õ—Ìﬂ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   5865
         Left            =   15
         TabIndex        =   4
         Top             =   1635
         Width           =   15855
         _cx             =   27966
         _cy             =   10345
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
         Caption         =   "«·ﬁÌÊœ|«·‘—Õ «·⁄«„"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   6
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5775
            Index           =   0
            Left            =   45
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   45
            Width           =   14895
            _cx             =   26273
            _cy             =   10186
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
            BackColor       =   16777215
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   2
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
            GridRows        =   2
            GridCols        =   4
            Frame           =   1
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmAccEditJournal2.frx":11F3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic EleOpt 
               Height          =   945
               Left            =   3750
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   4800
               Visible         =   0   'False
               Width           =   3690
               _cx             =   6509
               _cy             =   1667
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
               ForeColorDisabled=   -2147483630
               Caption         =   "⁄—÷ «·œ·Ì· «·„Õ«”»Ï"
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
               Begin VB.Frame Frame1 
                  Height          =   975
                  Left            =   -120
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   0
                  Width           =   3450
                  Begin VB.CommandButton Command6 
                     Caption         =   "Command6"
                     Height          =   375
                     Left            =   2040
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   600
                     Width           =   975
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ÃœÊ·Ï"
                     Height          =   285
                     Index           =   2
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   600
                     Width           =   1455
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ÿ«„ «·‘Ã—Ï"
                     Height          =   270
                     Index           =   0
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   390
                     Width           =   1455
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ÿ«„ «·„”«—"
                     Height          =   270
                     Index           =   1
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   120
                     Value           =   -1  'True
                     Width           =   1575
                  End
               End
               Begin C1SizerLibCtl.C1Elastic EleSortOpt 
                  Height          =   540
                  Left            =   2955
                  TabIndex        =   19
                  TabStop         =   0   'False
                  Top             =   285
                  Width           =   7410
                  _cx             =   13070
                  _cy             =   953
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
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " — Ì» »«·œ·Ì· «·„Õ«”»Ï"
                     Height          =   195
                     Index           =   11
                     Left            =   -735
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   -90
                     Value           =   -1  'True
                     Width           =   9180
                  End
               End
               Begin VB.Image ImgNote 
                  Height          =   240
                  Left            =   30
                  Picture         =   "FrmAccEditJournal2.frx":125E
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   240
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   4740
               Left            =   30
               TabIndex        =   5
               Top             =   30
               Width           =   14835
               _cx             =   26167
               _cy             =   8361
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
               Rows            =   10
               Cols            =   19
               FixedRows       =   2
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmAccEditJournal2.frx":17E8
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
               Begin VB.Frame Frame3 
                  Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
                  Height          =   1215
                  Left            =   -120
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   3720
                  Visible         =   0   'False
                  Width           =   4215
                  Begin VB.CommandButton Command5 
                     Caption         =   "‰”Œ"
                     Height          =   255
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   720
                     Width           =   1215
                  End
                  Begin VB.TextBox Text4 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   240
                     Width           =   2175
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—ﬁ„ «·ﬁÌœ"
                     Height          =   255
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin VB.PictureBox PicDes 
                  BorderStyle     =   0  'None
                  Height          =   3915
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   3915
                  ScaleWidth      =   9405
                  TabIndex        =   18
                  Top             =   810
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.TextBox TxtDese 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1485
                     Left            =   120
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   93
                     Top             =   2040
                     Width           =   8955
                  End
                  Begin VB.TextBox txtcodesub 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   3600
                     Width           =   855
                  End
                  Begin VB.CommandButton Command4 
                     Caption         =   "Add des"
                     Height          =   255
                     Left            =   7440
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   3600
                     Width           =   1350
                  End
                  Begin VB.CommandButton Command3 
                     Caption         =   "Call des"
                     Height          =   255
                     Left            =   6240
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   3600
                     Width           =   1095
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                     Height          =   3900
                     Left            =   120
                     TabIndex        =   86
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   10905
                     _cx             =   19235
                     _cy             =   6879
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial (Arabic)"
                        Size            =   20.25
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
                     BackColor       =   16777215
                     ForeColor       =   4210688
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
                     CaptionPos      =   7
                     WordWrap        =   -1  'True
                     MaxChildSize    =   0
                     MinChildSize    =   0
                     TagWidth        =   0
                     TagPosition     =   0
                     Style           =   0
                     TagSplit        =   2
                     PicturePos      =   7
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
                     Begin VB.TextBox TxtDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000018&
                        BorderStyle     =   0  'None
                        Height          =   1605
                        Left            =   0
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   87
                        Top             =   360
                        Visible         =   0   'False
                        Width           =   8955
                     End
                     Begin VB.Label LblDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000C&
                        Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                        ForeColor       =   &H0000C8FF&
                        Height          =   315
                        Left            =   6840
                        RightToLeft     =   -1  'True
                        TabIndex        =   88
                        Top             =   0
                        Width           =   2445
                     End
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Code"
                     Height          =   495
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   3480
                     Width           =   735
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     Height          =   495
                     Left            =   1560
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   1200
                     Width           =   975
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Code"
                     Height          =   255
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   1320
                     Width           =   735
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   17
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   2475
                  _cx             =   1973752078
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmAccEditJournal2.frx":1B08
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   113
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   945
               Left            =   0
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   4830
               Width           =   14895
               _cx             =   26273
               _cy             =   1667
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   20.25
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
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   2
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
               Style           =   0
               TagSplit        =   2
               PicturePos      =   7
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
               Begin VB.Frame Frame2 
                  Height          =   855
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   0
                  Width           =   3015
                  Begin ALLButtonS.ALLButton CmdRemove 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   72
                     Tag             =   "Delete Row"
                     Top             =   120
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   "Õ–› ”ÿ—"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   0
                     BCOLO           =   0
                     FCOL            =   255
                     FCOLO           =   255
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "FrmAccEditJournal2.frx":20A2
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "  — Ì» »«·œ·Ì· «·„Õ«”»Ì"
                     Height          =   270
                     Index           =   1
                     Left            =   840
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   480
                     Width           =   1995
                  End
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "  — Ì» «»ÃœÏ"
                     Height          =   270
                     Index           =   0
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   240
                     Width           =   1515
                  End
               End
               Begin DBPIXLib.DBPix20 DBPix202 
                  Height          =   855
                  Left            =   3240
                  TabIndex        =   57
                  Top             =   0
                  Width           =   3135
                  _Version        =   131072
                  _ExtentX        =   5530
                  _ExtentY        =   1508
                  _StockProps     =   1
                  BackColor       =   16777215
                  _Image          =   "FrmAccEditJournal2.frx":20BE
                  ImageResampleWidth=   100
                  ImageResampleHeight=   100
                  ImageResampleMode=   1
                  ImageSaveFormat =   0
                  JPEGQuality     =   75
                  JPEGEncoding    =   0
                  JPEGColorMode   =   0
                  JPEGNoRecompress=   -1  'True
                  JPEGRotateWarning=   0
                  PNGColorDepth   =   0
                  PNGCompression  =   0
                  PNGFilter       =   0
                  PNGInterlace    =   1
                  ImageDitherMethod=   3
                  ImagePaletteMethod=   4
                  ImagePreviewMode=   0   'False
                  ImageKeepMetaData=   -1  'True
                  UseAmbientBackcolor=   -1  'True
                  ViewAsyncDecoding=   -1  'True
                  ViewEnableMouseZoom=   -1  'True
                  ViewInitialZoom =   0
                  ViewHAlign      =   1
                  ViewVAlign      =   1
                  ViewMenuMode    =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ÊﬁÌ⁄"
                  Height          =   240
                  Index           =   5
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Tag             =   "51"
                  Top             =   0
                  Width           =   1410
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5775
            Index           =   1
            Left            =   16500
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   45
            Width           =   14895
            _cx             =   26273
            _cy             =   10186
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
            Begin VB.TextBox Txtcode 
               Alignment       =   1  'Right Justify
               Height          =   420
               Left            =   19230
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   4440
               Width           =   2220
            End
            Begin VB.CommandButton Command2 
               Caption         =   "«” œ⁄«¡ ﬁ«·» ‘—Õ"
               Height          =   540
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   4695
               Width           =   4560
            End
            Begin VB.CommandButton Command1 
               Caption         =   "«÷«›… ﬁ«·» ‘—Õ"
               Height          =   540
               Left            =   16455
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   5295
               Width           =   4635
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   3870
               Left            =   4080
               MaxLength       =   1000
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Top             =   390
               Width           =   10575
            End
            Begin VB.Label Lb_note_value_by_characters 
               Alignment       =   1  'Right Justify
               Height          =   480
               Left            =   9855
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   6315
               Width           =   11940
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Code"
               Height          =   540
               Left            =   16815
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   4440
               Width           =   2010
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄·Ìﬁ:"
               Height          =   150
               Index           =   6
               Left            =   17070
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Tag             =   "22"
               Top             =   510
               Width           =   4605
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   1155
         Left            =   15
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   7515
         Width           =   15855
         _cx             =   27966
         _cy             =   2037
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
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Left            =   30
            TabIndex        =   20
            Top             =   120
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   12648447
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtTotalCredit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   5700
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   127
            Width           =   1995
         End
         Begin VB.TextBox TxtTotalDebit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   9915
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   127
            Width           =   1845
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   0
            Left            =   14430
            TabIndex        =   35
            Top             =   480
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   635
            ButtonStyle     =   1
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   1
            Left            =   13020
            TabIndex        =   36
            Top             =   450
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   635
            ButtonStyle     =   1
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
            Height          =   360
            Index           =   2
            Left            =   11235
            TabIndex        =   37
            Top             =   480
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   635
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
            Height          =   360
            Index           =   3
            Left            =   9840
            TabIndex        =   38
            Top             =   450
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
            ButtonStyle     =   1
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
            Height          =   360
            Index           =   4
            Left            =   6225
            TabIndex        =   39
            Top             =   450
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«÷«›…"
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
            Height          =   360
            Index           =   5
            Left            =   4800
            TabIndex        =   40
            Top             =   450
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   6
            Left            =   0
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   450
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
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
            Height          =   360
            Index           =   7
            Left            =   3075
            TabIndex        =   42
            Top             =   450
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   635
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   360
            Left            =   1530
            TabIndex        =   43
            Top             =   450
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   635
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
            Height          =   360
            Index           =   8
            Left            =   8070
            TabIndex        =   92
            Top             =   480
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            ButtonStyle     =   1
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ALLButtonS.ALLButton ALLButton20 
            Height          =   255
            Left            =   10530
            TabIndex        =   98
            Top             =   840
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "«⁄ „«œ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":20D6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton6 
            Height          =   255
            Left            =   9450
            TabIndex        =   99
            Top             =   840
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "ﬁÌœ œÊ—Ì"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":20F2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton7 
            Height          =   255
            Left            =   6090
            TabIndex        =   100
            Top             =   840
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   " ÕÊÌ· «·Ï ﬁ«·»"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":210E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton8 
            Height          =   255
            Left            =   3465
            TabIndex        =   101
            Top             =   840
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "«·€«¡ «· √ÀÌ—"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":212A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton9 
            Height          =   255
            Left            =   1695
            TabIndex        =   102
            Top             =   840
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "«‰‘«¡ ﬁÌœ ⁄ﬂ”Ì"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65535
            BCOLO           =   65535
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":2146
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton10 
            Height          =   255
            Left            =   4800
            TabIndex        =   103
            Top             =   840
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "«” œ⁄«¡ ﬁ«·»"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":2162
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   255
            Left            =   11670
            TabIndex        =   104
            Top             =   840
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "„—«ﬂ“ «· ﬂ·›…"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":217E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   255
            Left            =   0
            TabIndex        =   105
            Top             =   840
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "«·„—›ﬁ« "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":219A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton3 
            Height          =   255
            Left            =   7725
            TabIndex        =   106
            Top             =   840
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "«‰‘«¡ ﬁÌœ œÊ—Ì"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAccEditJournal2.frx":21B6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ…"
            Height          =   240
            Index           =   8
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Tag             =   "51"
            Top             =   150
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ÿ—› «·œ«∆‰"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   2
            Left            =   7815
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Tag             =   "56"
            Top             =   150
            Width           =   2040
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ÿ—› «·„œÌ‰"
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   1
            Left            =   11925
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Tag             =   "55"
            Top             =   120
            Width           =   2190
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   930
         Left            =   15
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   690
         Width           =   15855
         _cx             =   27966
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
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   12600
            Top             =   960
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin VB.Frame Frame17 
            Height          =   855
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   30
            Width           =   12720
            Begin VB.CheckBox chkAll 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬂ·"
               Height          =   285
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   540
               Width           =   675
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               Caption         =   "œ«∆‰"
               Height          =   195
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   180
               Width           =   615
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "„œÌ‰"
               Height          =   195
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   180
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   9600
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   120
               Width           =   1335
            End
            Begin VB.CheckBox ChkLastAccount 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ «·Õ”«» «·‰Â«∆Ï ›ﬁÿ"
               Height          =   270
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   390
               Value           =   1  'Checked
               Width           =   2955
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Text            =   "Text1"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œÌ„ «· √ÀÌ—"
               Height          =   195
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   90
               Width           =   1455
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «⁄ „«œÂ"
               Height          =   195
               Left            =   900
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   150
               Width           =   1335
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁ«·»"
               Height          =   195
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   150
               Width           =   1095
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌœ œÊ—Ì"
               Height          =   195
               Left            =   -240
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   600
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               Caption         =   "„·€Ì"
               Height          =   195
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   300
               Width           =   1335
            End
            Begin MSDataListLib.DataCombo DcCostCenter 
               Bindings        =   "FrmAccEditJournal2.frx":21D2
               Height          =   315
               Left            =   5880
               TabIndex        =   91
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
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
            Begin MSDataListLib.DataCombo dcprojects 
               Bindings        =   "FrmAccEditJournal2.frx":21E7
               Height          =   315
               Left            =   6420
               TabIndex        =   94
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
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
            Begin MSComCtl2.DTPicker txtDueDate 
               Height          =   300
               Left            =   3990
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   510
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   143720449
               CurrentDate     =   37140
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
               Height          =   180
               Index           =   16
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Tag             =   "53"
               Top             =   555
               Width           =   1170
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„‘—Ê⁄ «·⁄«„"
               Height          =   255
               Left            =   8520
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
               Height          =   255
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "»‰«¡ ⁄·Ï"
               Height          =   255
               Left            =   11400
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "„’œ— «·ﬁÌœ"
               Height          =   255
               Left            =   11880
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.PictureBox DtHijriTrans 
            BackColor       =   &H000000FF&
            Height          =   1005
            Left            =   0
            ScaleHeight     =   945
            ScaleWidth      =   1110
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   330
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   75
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   12825
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   60
            Width           =   1650
         End
         Begin VB.TextBox TxtValue 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   6945
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   1005
            Visible         =   0   'False
            Width           =   3435
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11220
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   60
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox TxtDEVID 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   405
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtDEV_NO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   12075
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   780
            Visible         =   0   'False
            Width           =   2370
         End
         Begin C1SizerLibCtl.C1Elastic ElePost 
            Height          =   450
            Left            =   615
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   900
            Visible         =   0   'False
            Width           =   3555
            _cx             =   6271
            _cy             =   794
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
            ForeColorDisabled=   -2147483630
            Caption         =   "Õ«·… «·”‰œ"
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   2
            ChildSpacing    =   1
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
            Frame           =   4
            FrameStyle      =   3
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.CheckBox ChkPost 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·”‰œ"
               Height          =   225
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   45
               Width           =   1485
            End
            Begin VB.Image Img 
               Height          =   180
               Index           =   1
               Left            =   1635
               Top             =   285
               Width           =   285
            End
            Begin VB.Image Img 
               Height          =   225
               Index           =   0
               Left            =   90
               Top             =   90
               Width           =   270
            End
         End
         Begin MSComCtl2.DTPicker DTP_Date 
            Height          =   330
            Left            =   12750
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   435
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   143720451
            CurrentDate     =   37140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   315
            Index           =   0
            Left            =   14445
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Tag             =   "52"
            Top             =   495
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   270
            Index           =   3
            Left            =   14475
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Tag             =   "53"
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·”‰œ"
            Height          =   270
            Index           =   4
            Left            =   10260
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Tag             =   "54"
            Top             =   1020
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   330
            Index           =   7
            Left            =   14505
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Tag             =   "57"
            Top             =   600
            Visible         =   0   'False
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "FrmAccEditJournal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim line_no1 As Integer
Dim last_line_id As Integer
Dim numbering_type As Integer
Dim TTP As New clstooltip
Dim BolEditOnMainAccounts As Boolean
Dim PicHeight As Long
Dim PicWidth As Long
Dim Dcombos As ClsDataCombos
Dim DCboSearch As New clsDCboSearch
Public LngRow As Long
Private Enum PrintTarget
    WindowTarget
    PrinterTarget
End Enum

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    auto_sanad_no = ""
    departement_name = 1
    branch_no = 1
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=0"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=200 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=200 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(DTP_Date.value, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=200 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = mId(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4) & mId(Format$(DTP_Date.value, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Private Sub ALLButton1_Click()
    'On Error GoTo ErrTrap
    On Error Resume Next

    If DcCostCenter.BoundText <> "" Then

        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.Text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = Text1.Text
    'Else
    'opr_id = TxtDEV_NO.text
    'End If

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) = "" And Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) = "0" Then
            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            
            marakes_taklefa_tawze3.txtAccountSerial = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("Account_Serial"))
            
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        Else
    
            If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) = "" And Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) = "0" Then
                marakes_taklefa_tawze3.show
            
                marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) 'Text5.Text
                marakes_taklefa_tawze3.depit_or_credit.Caption = "œ«∆‰"
                marakes_taklefa_tawze3.kedno = opr_id
                    
                marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
                marakes_taklefa_tawze3.txtAccountSerial = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("Account_Serial"))
                marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
                marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
             
            End If
        End If

        marakes_taklefa_tawze3.opr_type = "”‰œ ﬁÌœ"
        marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
        marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
        marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
        marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        marakes_taklefa_tawze3.Adodc3.Refresh
        Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub ALLButton10_Click()

    If Me.TxtModFlg.Text <> "N" Then MsgBox "·«»œ „‰ «·÷€ÿ ⁄·Ï ÃœÌœ «Ê·« ·«” œ⁄«¡ «·ﬁ«·» ": Exit Sub
  
    'If Fg_Journal.Rows > 4 Then MsgBox "ÌÊÃœ «”ÿ— ›Ì Â–« «·ﬁÌœ ·–·ﬂ ·«Ì„ﬂ‰ «” œ⁄«¡ ﬁ«·» «·ﬁÌœ": Exit Sub

    KALEB.show
End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next
 
    If TxtSerial.Text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— ﬁÌœ «Ê·«": Exit Sub

    imaged.show

    If SystemOptions.UserInterface = EnglishInterface Then

        imaged.Label9.Caption = "Voucher #"
        imaged.Caption = "Voucher Attachment"
        imaged.txtopeation_type = "„—›ﬁ«  «·ﬁÌœ"
        imaged.SUBJECT_NO = TxtSerial.Text
        imaged.Label6.Caption = "Voucher #"
    Else

        imaged.Label9.Caption = "„—›ﬁ«  ”‰œ ﬁÌœ  —ﬁ„"
        imaged.Caption = "„—›ﬁ«  «·ﬁÌœ  "
        imaged.txtopeation_type = "„—›ﬁ«  «·ﬁÌœ"
        imaged.SUBJECT_NO = TxtSerial.Text
        imaged.Label6.Caption = "—ﬁ„  «·ﬁÌœ"

    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·ﬁÌœ' and subject_no='" & TxtSerial.Text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub ALLButton20_Click()

    If Dir(App.path & "\images\sign" & user_id & ".JPG") <> "" Then
        DBPix202.ImageLoadFile (App.path & "\images\sign" & user_id & ".JPG")
   
        Check2.value = 1

    Else
        MsgBox "·« ÌÕﬁ ·Â–« «·„” Œœ„ «⁄ „«œ «·”‰œ« "
    End If

End Sub

Private Sub ALLButton3_Click()

    If Me.TxtModFlg.Text <> "N" Then MsgBox "·«»œ „‰ «·÷€ÿ ⁄·Ï ÃœÌœ «Ê·« ·«’œ«— «·ﬁÌœ «·œÊ—Ì": Exit Sub
    keddawrym.show

End Sub

Private Sub ALLButton6_Click()

    'If Me.TxtModFlg.text <> "E" And Me.TxtModFlg.text <> "N" Then MsgBox "«÷€ÿ  ⁄œÌ·  «Ê ÃœÌœ «Ê·«", vbCritical: Exit Sub
    If TxtDEV_NO.Text = "" Then MsgBox "«Œ — ﬁÌœ «Ê·«", vbCritical: Exit Sub
    ked_dawry.show
    ked_dawry.ID = Me.TxtNoteID ' TxtDEV_NO.text
    ked_dawry.desc = Txt.Text
    ked_dawry.TxtSerial = Me.TxtSerial
    Check4.value = vbChecked
End Sub

Private Sub ALLButton7_Click()

    If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then MsgBox "«÷€ÿ  ⁄œÌ·  «Ê ÃœÌœ «Ê·«", vbCritical: Exit Sub
    X = MsgBox(" √ﬂÌœ «· ÕÊÌ· «·Ï ﬁ«·»", vbInformation + vbYesNo)

    If X = vbYes Then
        Check3.value = 1
    End If

End Sub

Private Sub ALLButton8_Click()

    If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then MsgBox "«÷€ÿ  ⁄œÌ·  «Ê ÃœÌœ «Ê·«", vbCritical: Exit Sub
    If Check1.value = vbChecked Then
        Check1.value = 1
        Check1.value = Unchecked
    Else
        Check1.value = vbChecked
    End If

End Sub

Private Sub ALLButton9_Click()
    Me.Txt.Text = ""
    Check1.value = Unchecked
    Check2.value = Unchecked
    Check3.value = Unchecked
    Check4.value = Unchecked
    Check5.value = Unchecked

    Me.TxtNoteID.Text = ""
    Me.TxtDEVID.Text = ""
    Me.DTP_Date.value = Date
    Me.TxtSerial.Text = ""
    Me.TxtValue.Text = ""

    Me.ChkPost.value = vbUnchecked

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.ChkPost.Caption = "€Ì— „—Õ·"
    Else
        Me.ChkPost.Caption = "Not Poasted"
    End If

    Me.ChkPost.ForeColor = vbBlack
    'Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Me.DcboUsers.BoundText = user_id
 
    Me.TxtModFlg.Text = "N"
    setfoxy
    DcCostCenter.Text = ""
    Option1.value = True
    Dim temp_value As Double

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Single

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) > 0 Then
                
                    .TextMatrix(i, .ColIndex("CreditValue")) = val(.TextMatrix(i, .ColIndex("DebitValue")))
                    .TextMatrix(i, .ColIndex("DebitValue")) = 0
                Else
                    .TextMatrix(i, .ColIndex("DebitValue")) = val(.TextMatrix(i, .ColIndex("CreditValue")))
                    .TextMatrix(i, .ColIndex("CreditValue")) = 0
                End If
            
            End If

        Next i

    End With

End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

    If ButtonID = vdsDownArrow Then
        If CboDes.IsDropped = False Then
            If PicHeight > 0 Then
                '    PicDes.Height = PicHeight
                '    PicDes.Width = PicWidth
            Else
                '    PicDes.Width = CboDes.Width - 10
                '    PicDes.Height = CboDes.Height * 8
            End If

            '  Debug.Print PicHeight
            '  Debug.Print PicWidth
            TxtDes.Visible = True
            TxtDes.Text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            TxtDese.Visible = True
            TxtDese.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("dese")) ' Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Dese"))
        
            CboDes.DropDown PicDes.hwnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            '  Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{F4}"
    End If

End Sub

Private Sub ChkPost_Click()

    'Stop
    If ChkPost.value = vbChecked Then
        Img(1).Visible = True
        Img(0).Visible = False
        ChkPost.ForeColor = vbRed
    ElseIf ChkPost.value = vbUnchecked Then
        Img(0).Visible = True
        Img(1).Visible = False
        ChkPost.ForeColor = vbBlack
    End If

End Sub

Function setfoxy_Line() As Integer
    
    last_line_id = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = last_line_id
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = last_line_id
 
    rs.update
    
End Function

Function setfoxy()
    Text1.Text = CStr(new_id("foxy", "id", "", True))
    'last_line_id = CStr(new_id("foxy", "id1", "", True))
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.Text
 
    rs.update
    
End Function

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            SetForNew
            Me.TxtModFlg.Text = "N"
            setfoxy
            DcCostCenter.Text = ""
    
            Option1.value = True

        Case 1
    
            Me.TxtModFlg.Text = "E"
  
            Fg_Journal.Rows = Fg_Journal.Rows + 1
   
        Case 2

            '  Me.DcboUsers.BoundText = user_id
            If Me.TxtModFlg.Text = "N" Then
                If Notes_coding(val(my_branch), DTP_Date.value) = "error" Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        MsgBox "can't Add new voucher because you exceed the numbering  ": Exit Sub
                    Else
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁÌœ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    End If
 
                Else

                    If Notes_coding(val(my_branch), DTP_Date.value) = "" Then

                        If TxtSerial.Text = "" Then
                            If SystemOptions.UserInterface = EnglishInterface Then
                                MsgBox "Enter Voucher code ": Exit Sub
                            Else
                                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·ﬁÌœ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                 
                            End If
                        End If

                    Else
  
                        TxtSerial.Text = Notes_coding(val(my_branch), DTP_Date.value)
          
                    End If
 
                End If
            End If

            SaveData

        Case 3
            Undo
        
        Case 4
            Frame3.Visible = True
      
        Case 5
            Voucher_search.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc TxtSerial.Text, , 200

        Case 8
            Del_Trans
    End Select

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    'On Error GoTo ErrTrap
    If TxtNoteID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·ﬁÌœ —ﬁ„ " & CHR(13)
        Msg = Msg + (Me.TxtSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
   
            StrSQL = "Delete  Notes  where NoteID =" & val(TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
  
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            Dim rs As New ADODB.Recordset

            StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & "From notes where (((notes.NoteType)=200)) " & "    ORDER BY NOTES.NoteID "
    
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
           
            If rs.RecordCount < 1 Then
                clear_all Me
                '  Fg_Journal.Clear flexClearScrollable, flexClearEverything
                
                TxtModFlg_Change
               
                Fg_Journal.Clear flexClearScrollable, flexClearEverything
                Me.TxtTotalCredit.Text = 0
                Me.TxtTotalDebit.Text = 0

            Else

                If Not (IsNull(rs("NoteID").value)) Then
                    Me.Retrive rs("NoteID").value
                    StrOldTransID = rs("NoteID").value
                End If

            End If
        
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub Undo()

    'On Error GoTo ErrTrap
    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            '   Rs.find "id='" & Val(Me.TXTid.text) & "'", , adSearchForward, adBookmarkFirst
            '         If Rs.EOF Or Rs.BOF Then
            '            Me.TxtModFlg.text = "R"
            '            Exit Sub
            '         End If
            Retrive (val(TxtDEV_NO.Text))
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    Dim sql As String

    sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
    Cn.Execute sgl, , adExecuteNoRecords
    
    If Fg_Journal.Rows > 1 Then
        If Fg_Journal.Rows = 2 Then
            Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Fg_Journal.Rows > 1 Then
                If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                    Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

    With Fg_Journal
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
    End With
            
End Sub

Private Sub Command1_Click()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[ked_desc]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("ked").value = Txt.Text
    rs("code").value = Txtcode.Text
        
    rs.update
    '    Cn.CommitTrans
    rs.Close
End Sub

Private Sub Command2_Click()
    Unload KEDDES
    KEDDES.show
End Sub

Private Sub Command3_Click()
    Unload KEDDES
    KEDDES.show
    KEDDES.case_id = 1
    KEDDES.rowno = Fg_Journal.Row
    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)

End Sub

Private Sub Command4_Click()

    If Len(TxtDes.Text) = 0 Then Exit Sub
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[ked_desc]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("ked").value = TxtDes.Text
    rs("code").value = txtcodesub.Text
        
    rs.update
    '    Cn.CommitTrans
    rs.Close
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Dim X As Long

    If Len(Text4.Text) = 0 Then Exit Sub
    X = get_Notes_id(Text4.Text)

    If X <> 0 Then
        Me.Retrive2 (X)
        Frame3.Visible = False
        ReLineGrid
        Fg_Journal.Rows = Fg_Journal.Rows + 1
        Text4.Text = ""
    End If

End Sub

Private Sub Command6_Click()
    ' .Cell(flexcpData, .Row, .ColIndex("Des")) = "Hiiiiiii"
    '                   .TextMatrix(I, .ColIndex("des")) = IIf(IsNull(Rs("Double_Entry_Vouchers_Description").value), _
                        "", Rs("Double_Entry_Vouchers_Description").value)
            
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
    check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "DebitValue", "CreditValue"

                'remove destribution
     
                sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If
                
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
            
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        If LastAccount(rs("Account_Code").value) = False Then
                            .TextMatrix(Row, Col) = ""
                            .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                            Exit Sub
                        End If
                    End If

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
        
                sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                If BolEditOnMainAccounts = False Then
                    If LastAccount(StrAccountCode) = False Then
                        .TextMatrix(Row, Col) = ""
                        .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Else

                        .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                        .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    End If

                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If

                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ReLineGrid

    End With

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
                Cancel = True
            End If
        End If

        Select Case .ColKey(Col)

            Case "LineNo"
                .ComboList = ""
                Cancel = True
                Exit Sub

            Case "DebitValue", "CreditValue", "Account_Serial"
                .ComboList = ""

            Case "DebitValueE", "CreditValuEe", "Account_Serial"
                .ComboList = ""
            
            Case "DebitCode", "CreditCode"
                .ComboList = ""

            Case "Des"
                Cancel = True
        End Select

    End With

End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
  With Me.Fg_Journal

        Select Case .ColKey(Col)

            
            Case "DueDate"
                Dim Frm As New FrmDateOpProject
                
                Frm.Index = 542
                Me.LngRow = Row
                Frm.show 1
                
        End Select

    End With
End Sub

Private Sub Fg_Journal_Click()
    On Error Resume Next

    If user_id = 1 Or Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")) = CStr(user_id) Or Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")) = "" Then

    Else

        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Can't Edit this Record because it created by user : " & get_user_name(val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")))), vbCritical: Exit Sub
        Else
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ›Ì Â–« «·”ÿ— ·«‰Â  „ «÷«› … »Ê«”ÿ… „” Œœ„ «Œ— ÊÂÊ   : " & get_user_name(val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")))), vbCritical: Exit Sub
        End If
    End If

    check_cost_center
End Sub

Function check_cost_center()

    If Fg_Journal.Row = 2 Then Exit Function

    If Fg_Journal.TextMatrix(Fg_Journal.Row - 1, Fg_Journal.ColIndex("cost_center")) <> "True" Then
        Exit Function
    Else

        If Fg_Journal.TextMatrix(Fg_Journal.Row - 1, Fg_Journal.ColIndex("cost_center")) = "True" And Fg_Journal.TextMatrix(Fg_Journal.Row - 1, Fg_Journal.ColIndex("distributed")) = "" Then

            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Must select Cost Center For this Account ", vbCritical
            Else
                MsgBox "·«»œ „‰ «œŒ«· „—ﬂ“ «· ﬂ·›… ", vbCritical
            End If

            Exit Function
        End If
    End If

End Function

Private Sub Fg_Journal_DblClick()

    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
 
        update_accounts
    End If

    If KeyCode = 46 Then
        CmdRemove_Click
    End If

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 200

    End If

End Sub

Private Sub Fg_Journal_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    With Fg_Journal

        If Button = vbRightButton Then
            '    If .FixedRows <= .Row And .Row < .Rows - 1 Then
            '        If .TextMatrix(.Row, .ColIndex("AccountCode")) <> "" Then
            '            MDIFrmamin.MnuPopJournal_Parent.Tag = .Row
            '            MDIFrmamin.MnuPopJournal(0).Enabled = True
            '            Me.PopupMenu MDIFrmamin.MnuPopJournal_Parent
            '        Else
            '            MDIFrmamin.MnuPopJournal_Parent.Tag = .Row
            '            MDIFrmamin.MnuPopJournal(0).Enabled = False
            '            Me.PopupMenu MDIFrmamin.MnuPopJournal_Parent
            '        End If
            '    End If
        End If

    End With

End Sub

Function update_accounts()
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal
    
        If Opt(0).value = True Then
            'Tree display
            StrSQL = "SELECT ACCOUNTS.Account_Code, Space(2*(Len(Account_Code)))" & "+ ACCOUNTS.Account_Name   As DisName , ACCOUNTS.Parent_Account_Code," & "ACCOUNTS.last_account, ACCOUNTS.cannot_del" & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <> 'r' "

            If ChkLastAccount.value = vbChecked Then
                'StrSQL = StrSQL + " And(((ACCOUNTS.last_account) = True)) "
            End If

            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = Fg_Journal.BuildComboList(rs, "DisName", "Account_Code")
                
        ElseIf Opt(1).value = True Then

            'Full Path Display
            If SystemOptions.UserInterface = EnglishInterface Then
                
                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                End If

                If OptSort(1).value = True Then
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                End If
                
            Else
                
                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                End If

                If OptSort(1).value = True Then
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                End If
                
            End If
                
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
            Debug.Print StrSQL
        ElseIf Opt(2).value = True Then 'the normal Display
            StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del " & "From ACCOUNTS Where  ACCOUNTS.Account_Code <>'r' "

            If ChkLastAccount.value = vbChecked Then
                If SystemOptions.SysDataBaseType = AccessDataBase Then
                    StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                Else
                    StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                End If
            End If

            If OptSort(1).value = True Then
                StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
            Else
                StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
            End If

            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
        End If

        If StrComboList <> "" Then
            StrComboList = "|" & StrComboList
        End If

        .ComboList = StrComboList
   
    End With

End Function

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "AccountName"

                If Opt(0).value = True Then
                    'Tree display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, Space(2*(Len(Account_Code)))" & "+ ACCOUNTS.Account_Name   As DisName , ACCOUNTS.Parent_Account_Code," & "ACCOUNTS.last_account, ACCOUNTS.cannot_del" & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <> 'r' "

                    If ChkLastAccount.value = vbChecked Then
                        'StrSQL = StrSQL + " And(((ACCOUNTS.last_account) = True)) "
                    End If

                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "DisName", "Account_Code")
                
                ElseIf Opt(1).value = True Then

                    'Full Path Display
                    If SystemOptions.UserInterface = EnglishInterface Then
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                            End If
                        End If

                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                        End If
                
                    Else
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                            End If
                        End If

                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                        End If
                
                    End If
                
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                    Debug.Print StrSQL
                ElseIf Opt(2).value = True Then 'the normal Display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del " & "From ACCOUNTS Where  ACCOUNTS.Account_Code <>'r' "

                    If ChkLastAccount.value = vbChecked Then
                        If SystemOptions.SysDataBaseType = AccessDataBase Then
                            StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                        Else
                            StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                        End If
                    End If

                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If

                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Form_Activate()
    'Application_Mode Me.TxtModFlg.text
End Sub

Private Sub Form_Load()
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim GrdBck As New ClsBackGroundPic

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL
    StrSQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
    fill_combo Me.dcprojects, StrSQL

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(8).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Me.TxtModFlg.Text = "R"
    SetDtpickerDate Me.DTP_Date
    Me.TabMain.CurrTab = 0

    ' adjust the grid
    With Fg_Journal
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeCol(.ColIndex("LineNo")) = True
        .Cell(flexcpText, 0, .ColIndex("LineNo"), 1, .ColIndex("LineNo")) = "—ﬁ„ «·”ÿ—"

        .MergeCol(.ColIndex("DebitValue")) = True
        .MergeCol(.ColIndex("CreditValue")) = True
        .MergeCol(.ColIndex("Account_Serial")) = True
        .MergeCol(.ColIndex("AccountName")) = True
    
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "ﬂÊœ «·Õ”«»"
        .ColWidth(.ColIndex("Account_Serial")) = 1500

        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "«”„ «·Õ”«»"
        .ColWidth(.ColIndex("AccountName")) = 2500
    
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = " ﬁÌ„… «·ﬁÌœ »«·⁄„·… «·„Õ·Ì… "
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter

        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "„œÌ‰"
        .ColWidth(.ColIndex("DebitValue")) = 1590
        .ColFormat(.ColIndex("DebitValue")) = SystemOptions.SysDefCurrencyForamt '"#,###.00"
     
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "œ«∆‰"
        .ColWidth(.ColIndex("CreditValue")) = 1590
        .ColFormat(.ColIndex("CreditValue")) = SystemOptions.SysDefCurrencyForamt '"#,###.00"
    
        .Cell(flexcpText, 0, .ColIndex("DebitValueE"), 0, .ColIndex("CreditValueE")) = " ﬁÌ„… «·ﬁÌœ »«·⁄„·… «·«Ã‰»Ì… "
    
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter
        
        .Cell(flexcpText, 1, .ColIndex("DebitValueE"), 1, .ColIndex("DebitValueE")) = "„œÌ‰"
        .Cell(flexcpText, 1, .ColIndex("CreditValueE"), 1, .ColIndex("CreditValueE")) = "œ«∆‰"
        .ColFormat(.ColIndex("DebitValueE")) = SystemOptions.SysDefCurrencyForamt ' "#,###.00"
        .ColFormat(.ColIndex("CreditValueE")) = SystemOptions.SysDefCurrencyForamt ' "#,###.00"

        '.MergeCol(.ColIndex("Des")) = True
        '.Cell(flexcpText, 0, .ColIndex("Des"), 1, .ColIndex("Des")) = "«·‘—Õ"
        '.ColWidth(.ColIndex("Des")) = 2200
        Set .WallPaper = GrdBck.Picture
    End With

    'If SystemOptions.UserInterface = EnglishInterface Then
    '    SetInterface Me
    '    ChangeLang
    'End If
    'Me.Img(0).Picture = MDIFrmamin.ImgLstMenuIcons.ListImages("Unlock").Picture
    'Img(0).Visible = True
    'Me.Img(1).Picture = MDIFrmamin.ImgLstMenuIcons.ListImages("Lock").Picture
    'Img(1).Visible = False
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DcboUsers
    AddTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    'Resize_Form Me,    TransactionSize
    XPBtnMove_Click 2
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
                Cmd_Click (2)

                ' SaveData
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    '    Select Case QueryCloseMsg(Me.TxtModFlg.text, Me.Caption)
    '        Case vbYes
    '            Cancel = True
    '            Do_Action Do_save
    '        Case vbNo
    '            Cancel = False
    '            Application_Mode "R"
    '        Case vbCancel
    '            Cancel = True
    '    End Select
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Dcombos = Nothing
    Set DCboSearch = Nothing
    Set TTP = Nothing
End Sub

Private Sub Opt_Click(Index As Integer)

    Select Case Index

        Case 0
            ChkLastAccount.Enabled = False

        Case 1
            ChkLastAccount.Enabled = True

        Case 2
            ChkLastAccount.Enabled = True
    End Select

End Sub

Private Function LastAccount(StrAccountCode As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    If StrAccountCode = "" Then
        LastAccount = False
        Exit Function
    End If

    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account,ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Code='" & StrAccountCode & "'"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs("last_account").value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«·Õ”«» " & rs("Account_Name").value & CHR(13)
            Msg = Msg & "Õ”«» €Ì— ‰Â«∆Ï Ê·«Ì„ﬂ‰ ﬂ «»… ﬁÌœ ⁄·ÌÂ " & CHR(13)
            Msg = Msg & "»—Ã«¡  ÕœÌœ √Ï Õ”«» ›—⁄Ï  Õ  Â–« «·Õ”«»" & CHR(13)
            Msg = Msg & "√Ê ﬁ„ » ⁄—Ì› Õ”«»«  ›—⁄Ì… ÃœÌœ  Õ  Â–« «·Õ”«»"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            Msg = "The " & IIf(IsNull(rs("Account_NameEng").value), rs("Account_Name").value, rs("Account_NameEng").value) & " Account " & CHR(13)
            Msg = Msg & "is not a last account..!" & CHR(13)
            Msg = Msg & "and it is not accepted."
            MsgBox Msg, vbExclamation, App.title
        End If

        LastAccount = False
    Else
        LastAccount = True
    End If

Exit_Function:
    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    LastAccount = False
    Resume Exit_Function
End Function

Private Sub SetForNew()
    Me.Txt.Text = ""
    Check1.value = Unchecked
    Check2.value = Unchecked
    Check3.value = Unchecked
    Check4.value = Unchecked
    Check5.value = Unchecked

    Me.TxtNoteID.Text = ""
    Me.TxtDEVID.Text = ""
    Me.DTP_Date.value = Date
        Me.txtDueDate.value = Date


    Me.TxtSerial.Text = ""
    Me.TxtValue.Text = ""

    Me.ChkPost.value = vbUnchecked

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.ChkPost.Caption = "€Ì— „—Õ·"
    Else
        Me.ChkPost.Caption = "Not Poasted"
    End If

    Me.ChkPost.ForeColor = vbBlack
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Me.TxtTotalCredit.Text = 0
    Me.TxtTotalDebit.Text = 0
    Me.DcboUsers.BoundText = user_id
End Sub

Public Property Let Cmd_New(ByVal vNewValue As Boolean)
    m_Cmd_New = vNewValue
End Property

Public Property Get Cmd_Undo() As Boolean
    'Dim Msg As String
    'Dim BolTemp  As Boolean
    'Cmd_Undo = m_Cmd_Undo
    'On Error GoTo ErrTrap
    'Select Case TxtModFlg.text
    '    Case "N"
    '        If QueryUndoMsg(Me.TxtModFlg.text, Me.Caption) = vbYes Then
    '            BolTemp = Cmd_New
    '        Else
    '            Cmd_Undo = False
    '            Exit Property
    '        End If
    '    Case "E"
    '        If QueryUndoMsg(Me.TxtModFlg.text, Me.Caption) = vbYes Then
    '           Me.Retrive Me.TxtNoteID
    '            Cmd_Undo = True
    '        Else
    '            Cmd_Undo = False
    '            Exit Property
    '        End If
    'End Select
    'Cmd_Undo = True
    'Exit Property
    'ErrTrap:
End Property

Public Property Let Cmd_Undo(ByVal vNewValue As Boolean)
    m_Cmd_Undo = vNewValue
End Property

Private Sub PicDes_Resize()

    With PicDes
        '  LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        '  TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub TxtDes_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
    'TxtDes.RightToLeft = True
    TxtDes.Alignment = 1

End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        '    PutData
        '    CboDes.CloseUp
    End If

End Sub

Private Sub TxtDes_LostFocus()
    'PicHeight = PicDes.Height
    'PicWidth = PicDes.Width
    'CboDes.CloseUp
    'CboDes.Visible = False
End Sub

Private Sub TxtDese_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtModFlg_Change()

    Select Case TxtModFlg.Text

        Case "N"
            Me.EleHeader.Enabled = True
            Me.Fg_Journal.Editable = flexEDKbdMouse
            EleOpt.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = True
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False
            CmdRemove.Enabled = True
            Fg_Journal.Enabled = True
            ALLButton1.Enabled = True

        Case "E"
            Me.EleHeader.Enabled = True
            Me.Fg_Journal.Editable = flexEDKbdMouse
            EleOpt.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = True
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False
            CmdRemove.Enabled = True
            Fg_Journal.Enabled = True
            ALLButton1.Enabled = True

        Case "R"
            Me.EleHeader.Enabled = False
            Me.Fg_Journal.Editable = flexEDNone
            EleOpt.Enabled = False
            CboDes.CloseUp
            CboDes.Visible = False
        
            Cmd(0).Enabled = True
            Cmd(1).Enabled = True
            Cmd(2).Enabled = False
            Cmd(3).Enabled = False
            Cmd(4).Enabled = False
            Cmd(5).Enabled = True
            Cmd(7).Enabled = True
            CmdRemove.Enabled = False
            Fg_Journal.Enabled = False
            ALLButton1.Enabled = False
    End Select

End Sub

Public Function ReLineGridP()
    ReLineGrid
End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
            
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter

                If .TextMatrix(i, .ColIndex("LineNo1")) = "" Then
                    ' setfoxy_Line
                    .TextMatrix(i, .ColIndex("LineNo1")) = setfoxy_Line  'last_line_id

                End If
            
            End If

        Next i

    End With

    line_no1 = IntCounter

End Sub

Public Property Get Cmd_Search() As Boolean
    Cmd_Search = m_Cmd_Search
    Frm_SandSearch.show vbModal
    Cmd_Search = True
End Property

Public Property Let Cmd_Search(ByVal vNewValue As Boolean)
    m_Cmd_Search = vNewValue
End Property

Public Sub Retrive(LngNoteID As Long)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    StrSQL = "SELECT NOTES.project_id, NOTES.project_depit_or_credit,  NOTES.foxy_no,NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & _
       "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & _
       "NOTES.Remark,NOTES.general_cost_center, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & _
       "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,DOUBLE_ENTREY_VOUCHERS.USERID," & _
       "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS.Account_Code," & _
       "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,DOUBLE_ENTREY_VOUCHERS.Valuee,DOUBLE_ENTREY_VOUCHERS.currency,DOUBLE_ENTREY_VOUCHERS.rate," & _
       "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,DOUBLE_ENTREY_VOUCHERS.DueDate,DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name  " & _
       ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & _
       " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
       " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & _
       "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

    StrSQL = StrSQL + " Where NOTES.NoteID=" & LngNoteID & ""
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    If rs("DAWRY").value = 0 Then
        Check4.value = vbUnchecked
    Else
        Check4.value = vbChecked
    End If
  
    If rs("KALEB").value = 0 Then
        Check3.value = vbUnchecked
    Else
        Check3.value = vbChecked
    End If
  
    ' Check3.value = RsNetes("KALEB").value
    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If
 
    Me.TxtNoteID.Text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.dcprojects.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)

    If Not IsNull(rs("project_depit_or_credit").value) Then
        If rs("project_depit_or_credit").value = 0 Then
            Option1.value = True
        ElseIf rs("project_depit_or_credit").value = 1 Then
            Option2.value = True
        End If
    End If

    If rs("Notetype").value = 200 Then
        Text2.Text = "Manual"

    Else
        Text2.Text = "Auto"

    End If

    Text3.Text = get_note_type_name(rs("Notetype").value)

    Me.TxtDEVID.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)
    Me.TxtDEV_NO.Text = ""
    Me.TxtValue.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    Me.TxtDEV_NO.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)

    Me.DTP_Date.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Me.txtDueDate.value = IIf(IsNull(rs("DueDate").value), Date, rs("DueDate").value)

    
    Me.TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), Date, rs("NoteSerial").value)

    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    Me.DcboUsers.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)

    If Not (IsNull(rs("NoteType").value)) Then
        If rs("NoteType").value = "2" Then
            'Me.OptType(0).Value = True
        ElseIf rs("NoteType").value = 1 Then
            'Me.OptType(1).Value = True
        End If
    End If

    If rs("NotePosted").value = True Then
        ChkPost.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            ChkPost.Caption = "„—Õ·"
        Else
            ChkPost.Caption = "Posted"
        End If

        ChkPost.ForeColor = vbRed
    Else
        ChkPost.value = vbUnchecked

        If SystemOptions.UserInterface = ArabicInterface Then
            ChkPost.Caption = "€Ì— „—Õ·"
        Else
            ChkPost.Caption = "Not Posted"
        End If

        ChkPost.ForeColor = vbBlack
    End If

    rs.MoveFirst

    With Me.Fg_Journal
        .Rows = .FixedRows + rs.RecordCount

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(rs("DEV_ID_Line_No").value), "", rs("DEV_ID_Line_No").value)
            
            .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(rs("DEV_ID_Line_No1").value), "", rs("DEV_ID_Line_No1").value)
            
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
             .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(rs("DueDate").value), "", rs("DueDate").value)
            If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Nameeng").value), "", rs("Account_Nameeng").value)
                 
            Else
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            End If
            
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            
            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If

            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
            
                .TextMatrix(i, .ColIndex("DebitValuee")) = IIf(IsNull(rs("Valuee").value), "", rs("Valuee").value)
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = "0"
            
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = IIf(IsNull(rs("Valuee").value), "", rs("Valuee").value)
                .TextMatrix(i, .ColIndex("DebitValuee")) = "0"
                
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("userid")) = IIf(IsNull(rs("userid").value), "", rs("userid").value)
            
            .TextMatrix(i, .ColIndex("currenct_code")) = IIf(IsNull(rs("currency").value), "", rs("currency").value)
            
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(rs("rate").value), "", rs("rate").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
             
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
    
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
    
        Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
    
    End With

End Sub

Public Sub Retrive2(LngNoteID As Long)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    StrSQL = "SELECT  NOTES.foxy_no,NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & "NOTES.Remark, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,DOUBLE_ENTREY_VOUCHERS.USERID," & "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS.Account_Code," & "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,DOUBLE_ENTREY_VOUCHERS.Valuee,DOUBLE_ENTREY_VOUCHERS.currency,DOUBLE_ENTREY_VOUCHERS.rate," & "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,ACCOUNTS.Account_Name  " & ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

    StrSQL = StrSQL + " Where NOTES.NoteID=" & LngNoteID & ""
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    'If Rs("DAWRY").value = 0 Then
    'Check4.value = vbUnchecked
    'Else
    ' Check4.value = vbChecked
    'End If
  
    '  If Rs("KALEB").value = 0 Then
    'Check3.value = vbUnchecked
    'Else
    ' Check3.value = vbChecked
    'End If
  
    ' Check3.value = RsNetes("KALEB").value
    
    'Me.TxtNoteID.text = IIf(IsNull(Rs("NoteID").value), "", Rs("NoteID").value)
    'Me.Text1.text = IIf(IsNull(Rs("foxy_no").value), "", Rs("foxy_no").value)

    'If Rs("Notetype").value = 200 Then
    'Text2.text = "Manual"

    'Else
    'Text2.text = "Auto"

    'End If

    'Text3.text = get_note_type_name(Rs("Notetype").value)

    'Me.TxtDEVID.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)
    'Me.TxtDEV_NO.text = ""
    'Me.TxtValue.text = IIf(IsNull(Rs("Note_Value").value), "", Rs("Note_Value").value)
    'Me.TxtDEV_NO.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)

    'Me.DTP_Date.value = IIf(IsNull(Rs("NoteDate").value), Date, Rs("NoteDate").value)
    'Me.TxtSerial.text = IIf(IsNull(Rs("NoteSerial").value), Date, Rs("NoteSerial").value)

    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    'Me.DcboUsers.BoundText = IIf(IsNull(Rs("UserID").value), "", Rs("UserID").value)
    'Me.Txt.text = IIf(IsNull(Rs("Remark").value), "", Rs("Remark").value)
    'If Not (IsNull(Rs("NoteType").value)) Then
    '    If Rs("NoteType").value = "2" Then
    '        'Me.OptType(0).Value = True
    '    ElseIf Rs("NoteType").value = 1 Then
    '        'Me.OptType(1).Value = True
    '    End If
    'End If
    'If Rs("NotePosted").value = True Then
    '    ChkPost.value = vbChecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "„—Õ·"
    '    Else
    '        ChkPost.Caption = "Posted"
    '    End If
    '    ChkPost.ForeColor = vbRed
    'Else
    '    ChkPost.value = vbUnchecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "€Ì— „—Õ·"
    '    Else
    '        ChkPost.Caption = "Not Posted"
    '    End If
    '    ChkPost.ForeColor = vbBlack
    'End If
    Dim last_row As Integer
    rs.MoveFirst

    With Me.Fg_Journal
        last_row = .Rows
        .Rows = .Rows + rs.RecordCount - 1

        For i = last_row - 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("LineNo")) = i ' IIf(IsNull(Rs("DEV_ID_Line_No").value), "", Rs("DEV_ID_Line_No").value)
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            
            If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Nameeng").value), "", rs("Account_Nameeng").value)
                 
            Else
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            End If
            
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            
            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If

            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
            
                .TextMatrix(i, .ColIndex("DebitValuee")) = IIf(IsNull(rs("Valuee").value), "", rs("Valuee").value)
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = "0"
            
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = IIf(IsNull(rs("Valuee").value), "", rs("Valuee").value)
                .TextMatrix(i, .ColIndex("DebitValuee")) = "0"
                
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("userid")) = IIf(IsNull(rs("userid").value), "", rs("userid").value)
            
            .TextMatrix(i, .ColIndex("currenct_code")) = IIf(IsNull(rs("currency").value), "", rs("currency").value)
            
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(rs("rate").value), "", rs("rate").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
    End With

End Sub

Public Sub retrive1(LngNoteID As Long)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    StrSQL = "SELECT  NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & "NOTES.Remark, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID," & "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS.Account_Code," & "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit," & "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name  " & ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

    StrSQL = StrSQL + " Where NOTES.NoteID=" & LngNoteID & ""
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    ' If Rs("DAWRY").value = 0 Then
    ' ' Check3.value = vbUnchecked
    '' Else
    ' Check3.value = vbChecked
    'End If
  
    '    If Rs("KALEB").value = 0 Then
    '  Check4.value = vbUnchecked
    '  Else
    '   Check4.value = vbChecked
    '  End If
    '
    ' Check3.value = RsNetes("KALEB").value
    
    'Me.TxtNoteID.text = IIf(IsNull(Rs("NoteID").value), "", Rs("NoteID").value)

    'Me.TxtDEVID.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)
    'Me.TxtDEV_NO.text = ""
    'Me.TxtValue.text = IIf(IsNull(Rs("Note_Value").value), "", Rs("Note_Value").value)
    'Me.TxtDEV_NO.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)

    'Me.DTP_Date.value = IIf(IsNull(Rs("NoteDate").value), Date, Rs("NoteDate").value)
    'Me.TxtSerial.text = IIf(IsNull(Rs("NoteSerial").value), Date, Rs("NoteSerial").value)

    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    'Me.DcboUsers.BoundText = IIf(IsNull(Rs("UserID").value), "", Rs("UserID").value)
    'Me.Txt.text = IIf(IsNull(Rs("Remark").value), "", Rs("Remark").value)
    'If Not (IsNull(Rs("NoteType").value)) Then
    '    If Rs("NoteType").value = "2" Then
    '        'Me.OptType(0).Value = True
    '    ElseIf Rs("NoteType").value = 1 Then
    '        'Me.OptType(1).Value = True
    '    End If
    'End If
    'If Rs("NotePosted").value = True Then
    '    ChkPost.value = vbChecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "„—Õ·"
    '    Else
    '        ChkPost.Caption = "Posted"
    '    End If
    '    ChkPost.ForeColor = vbRed
    'Else
    '    ChkPost.value = vbUnchecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "€Ì— „—Õ·"
    '    Else
    '        ChkPost.Caption = "Not Posted"
    '    End If
    '    ChkPost.ForeColor = vbBlack
    'End If

    rs.MoveFirst

    With Me.Fg_Journal
        .Rows = .FixedRows + rs.RecordCount

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(rs("DEV_ID_Line_No").value), "", rs("DEV_ID_Line_No").value)
            
            .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(rs("DEV_ID_Line_No1").value), "", rs("DEV_ID_Line_No1").value)
            
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)

            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If
        
            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("USERID")) = IIf(IsNull(rs("USERID").value), "", rs("USERID").value)
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        Me.TxtTotalCredit.Text = Format(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
     
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TxtTotalDebit.Text = Format(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
    
    End With

End Sub

Public Property Get Cmd_Edit() As Boolean
    Dim Msg As String
    Cmd_Edit = m_Cmd_Edit

    If Trim(Me.TxtNoteID.Text) = "" Then
        'Msg = "·«ÌÊÃœ ”Ã· Õ«÷— ·· ⁄œÌ·"
        GetMsgs 72, vbExclamation
        Cmd_Edit = False
        Exit Property
    ElseIf Me.ChkPost.value = vbChecked Then
        'Msg = "Â–« «·”‰œ „—Õ· ...!!" & Chr(13)
        'Msg = Msg & "Ê·« Ì„ﬂ‰  ⁄œÌ· «·ﬁÌœ"
        GetMsgs 73, vbExclamation
        Cmd_Edit = False
        Exit Property
    Else
        Me.DcboUsers.BoundText = user_id 'LngUserID
        Cmd_Edit = True
        Exit Property
    End If

End Property

Public Property Let Cmd_Edit(ByVal vNewValue As Boolean)
    m_Cmd_Edit = vNewValue
End Property

Public Property Get Cmd_Delete() As Boolean
    Dim StrSQL  As String
    Dim Msg As String
    Dim BolTemp As Boolean
    Dim TransBegine As Boolean
    Dim rs As New ADODB.Recordset
    Dim IntRes As Integer
    On Error GoTo ErrTrap
    Cmd_Delete = m_Cmd_Delete

    If Me.TxtNoteID.Text = "" Then
        Cmd_Delete = True
        Exit Property
    End If

    If Me.ChkPost.value = vbChecked Then
        'Msg = "Â–« «·”‰œ „—Õ· ...!!" & Chr(13)
        'Msg = Msg & "Ê·« Ì„ﬂ‰ Õ–› «·ﬁÌœ...!!"
        GetMsgs 74, vbExclamation
        Cmd_Delete = True
        Exit Property
    End If

    StrSQL = "Delete * From Notes Where Notes.Note_ID='" & Trim(Me.TxtNoteID.Text) & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ê› Ì „ Õ–› Â–« «·”‰œ —ﬁ„ " & Trim(Me.TxtSerial.Text) & CHR(13)
        Msg = Msg & "›Â· √‰  „ √ﬂœ „‰ «·√” „—«— ...!!"
        IntRes = MsgBox(Msg, vbQuestion + vbOKCancel + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
    Else
        Msg = "This voucher " & Trim(Me.TxtSerial.Text) & CHR(13)
        Msg = Msg & "will be deleted " & CHR(13)
        Msg = Msg & "are you sure to continue ..?"
        IntRes = MsgBox(Msg, vbQuestion + vbOKCancel, App.title)
    End If

    If IntRes = vbOK Then
        Cn.BeginTrans
        TransBegine = True
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.CommitTrans
        TransBegine = False
    
        'Msg = " „ Õ–› «·”Ã·."
        GetMsgs 75, vbInformation
    End If

    Cmd_Delete = True
    Exit Property
ErrTrap:

    If TransBegine = True Then
        Cn.RollbackTrans
    End If

    'Msg = "ÕœÀ Œÿ√ √À‰«¡ Õ–› «·”Ã·"
    GetMsgs 76, vbExclamation
    Cmd_Delete = True
End Property

Public Property Let Cmd_Delete(ByVal vNewValue As Boolean)
    m_Cmd_Delete = vNewValue
End Property

Private Sub PutData()
    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)

    With Fg_Journal

        If Len(TxtDes.Text) > 0 And Len(TxtDese.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
            .TextMatrix(.Row, .ColIndex("des")) = TxtDes.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = TxtDes.Text
        
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Dese")) = flexAlignLeftCenter
            .TextMatrix(.Row, .ColIndex("dese")) = TxtDese.Text
        ElseIf Len(TxtDes.Text) > 0 And Len(TxtDese.Text) = 0 Then
    
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
            .TextMatrix(.Row, .ColIndex("des")) = TxtDes.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Dese")) = flexAlignLeftCenter
            .TextMatrix(.Row, .ColIndex("dese")) = ""
        ElseIf Len(TxtDes.Text) = 0 And Len(TxtDese.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
            .TextMatrix(.Row, .ColIndex("des")) = ""
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = TxtDes.Text
            .TextMatrix(.Row, .ColIndex("dese")) = TxtDese.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Dese")) = flexAlignLeftCenter
        ElseIf Len(TxtDes.Text) = 0 And Len(TxtDese.Text) = 0 Then
            .TextMatrix(.Row, .ColIndex("des")) = ""
            .TextMatrix(.Row, .ColIndex("dese")) = ""
    
            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        
        End If

    End With

End Sub

Public Property Get Cmd_Print() As Boolean

    If Me.TxtNoteID.Text = "" Then
        GetMsgs 140, vbExclamation
        Cmd_Print = False
    Else
        Cmd_Print = FireReport(PrinterTarget)
    End If

End Property

Public Property Let Cmd_Print(ByVal vNewValue As Boolean)
    m_Cmd_Print = vNewValue
End Property

Private Function FireReport(m_Destination As PrintTarget) As Boolean
    'Dim RsData As New ADODB.Recordset
    'Dim Rs As New ADODB.Recordset
    'Dim xApp As New CRAXDRT.Application
    'Dim xReport As CRAXDRT.Report
    'Dim Msg As String
    'Dim StrSQL As String
    'Dim StrPrinterName As String
    'Dim XPrinter As Object
    'Dim Frm As FrmPrint
    'Dim I As Integer
    'Dim StrFileName As String
    'On Error GoTo FireReportErrTrap
    'If Me.TxtNoteID.text = "" Then
    '    FireReport = False
    '    Exit Function
    'End If
    'StrSQL = "SELECT NOTES.NoteID, NOTES.Employee_ID, NOTES.NoteType, NOTES.NoteDate," & _
    '    "NOTES.Value, NOTES.Remark, NOTES.Chique_Serial_No, NOTES.Transaction_Header_ID," & _
    '    "NOTES.Dealer_Code, NOTES.NotePosted, NOTES.PostedBy, NOTES.PostDate," & _
    '    "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No," & _
    '    "DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Value as DEV_Value, DOUBLE_ENTREY_VOUCHERS." & _
    '    "Credit_Or_Debit, DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Remark," & _
    '    "DOUBLE_ENTREY_VOUCHERS.Notes_Id,ACCOUNTS.Account_Name, EMPLOYEES.Employee_Name," & _
    '    "USERS.UserName AS UserIssued, USERS_1.UserName AS UserPosted ,ACCOUNTS.Account_Serial "
    'StrSQL = StrSQL + " FROM (EMPLOYEES RIGHT JOIN ((USERS INNER JOIN NOTES ON USERS.User_ID = " & _
    '    "NOTES.Issued_BY) LEFT JOIN USERS AS USERS_1 ON NOTES.PostedBy = USERS_1.User_ID) " & _
    '    "ON EMPLOYEES.Employee_Code = NOTES.Employee_ID) INNER JOIN  " & _
    '    "(ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code =  " & _
    '    "DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id" & _
    '    " where NOTES.Note_ID='" & Me.TxtNoteID.text & "'" & _
    '    " ORDER BY Val(DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No);"
    'If SystemOptions.UserInterface = ArabicInterface Then
    '    StrFileName = App.Path & "\Reports\Journal.rpt"
    'Else
    '    StrFileName = App.Path & "\Reports\Journal_Eng.rpt"
    'End If
    'If Dir(StrFileName) = "" Then
    '    GetMsgs 139, vbExclamation
    '    FireReport = False
    '    Exit Function
    'End If
    'RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    'If RsData.BOF Or RsData.EOF Then
    '    GetMsgs 138, vbExclamation
    '    FireReport = False
    '    RsData.Close
    '    Set RsData = Nothing
    '    Exit Function
    'End If
    'Screen.MousePointer = vbArrowHourglass
    'Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData
    'Rs.Open "Options", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    'xReport.ParameterFields(1).AddCurrentValue Rs("Company_Name_Arabic").Value
    'xReport.ParameterFields(2).AddCurrentValue Rs("Comment_Arabic").Value
    'xReport.ParameterFields(3).AddCurrentValue Rs("Company_Name_Eng").Value
    'xReport.ParameterFields(4).AddCurrentValue Rs("Comment_Eng").Value
    'xReport.ParameterFields(5).AddCurrentValue StrUserName
    'If SystemOptions.UserInterface = ArabicInterface Then
    '     xReport.ReportTitle = "ÿ»«⁄… ﬁÌœ «·ÌÊ„Ì… —ﬁ„ " & Me.TxtSerial.text
    'Else
    '     xReport.ReportTitle = "Journal Voucher NO." & Me.TxtSerial.text
    'End If
    'xReport.EnableParameterPrompting = False
    'xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.Title
    '
    ''xReport.PaperSize=
    'If Not (IsNull(Rs("DefaultPrinter").Value)) Then
    '    StrPrinterName = Rs("DefaultPrinter").Value
    '    For I = 0 To Printers.count - 1
    '        If StrPrinterName = Printers(I).DeviceName Then
    '            Set XPrinter = Printers.Item(I)
    '            Exit For
    '        End If
    '    Next I
    '    If Not XPrinter Is Nothing Then
    '        xReport.SelectPrinter XPrinter.DriverName, XPrinter.DeviceName, XPrinter.Port
    '    End If
    'End If
    '
    'Set Frm = New FrmPrint
    'With Frm
    '    .CRViewerMain.ReportSource = xReport
    '    Do While .CRViewerMain.IsBusy
    '        DoEvents
    '    Loop
    '    .CRViewerMain.Zoom IIf(IsNull(Rs("RptZoom").Value), 100, Rs("RptZoom").Value)
    '    If m_Destination = WindowTarget Then
    '        .CRViewerMain.ViewReport
    '        .WindowState = vbMaximized
    '    Else
    '        'xReport.PrintOut "⁄œœ «·‰”Œ", 12
    '        xReport.PrintOut
    '        .CRViewerMain.PrintReport
    '    End If
    '
    '    If m_Destination = WindowTarget Then
    '        .Show
    '    Else
    '        Unload Frm
    '    End If
    'End With
    'Set xApp = Nothing
    'Set xReport = Nothing
    ''SendCrystalSetting cr, "ﬁÌÊœ «·ÌÊ„Ì…"
    'FireReport = True
    'Screen.MousePointer = vbDefault
    'Exit Function
    'FireReportErrTrap:
    'FireReport = False
    'Screen.MousePointer = vbDefault
End Function

Private Sub ChangeLang()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Me.Caption = "Edit Journal"
    Me.EleTop.Caption = Me.Caption

    Frame3.Caption = "Enter Voucher No. To copy it"
    Label7.Caption = "Voucher #"
    Command5.Caption = "Copy"

    'Rs.Open "Lang", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    'Rs.MoveFirst
    'For I = Me.lbl.LBound To Me.lbl.UBound
    '    If Trim(lbl(I).Tag) <> "" Then
    '        Rs.MoveFirst
    '        Rs.find "ID=" & Val(Me.lbl(I).Tag) & "", , adSearchForward, 1
    '        If Not (Rs.BOF Or Rs.EOF) Then
    '            Me.lbl(I).Caption = IIf(IsNull(Rs("Eng").value), "", Rs("Eng").value) & ":"
    '        End If
    '    End If
    'Next I
    'Rs.Close
    'Set Rs = Nothing
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Label1.Caption = "Source"
    Label2.Caption = "Based ON"

    lbl(7).Caption = "ID"
    lbl(0).Caption = "Date"
    lbl(3).Caption = "Serial"
    lbl(4).Caption = "Value"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Modify"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Insert"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    TabMain.TabCaption(0) = "Journal"
    TabMain.TabCaption(1) = "Comment"
    ElePost.Caption = "Posting State"
    ChkPost.Caption = "Voucher State"
    Check3.Caption = "Template"
    Check2.Caption = "Approved"
    Check1.Caption = "Cancel Action"
    Check5.Caption = "Deleted"
    Check4.Caption = "periodic"
    lbl(1).Caption = "Depit Sum"
    lbl(2).Caption = "Credit Sum"
    lbl(8).Caption = "By"
    lbl(5).Caption = "Signature"
    ALLButton1.Caption = "Cost Center"
    ALLButton20.Caption = "Approved"
    ALLButton3.Caption = "Repeat Voucher"
    ALLButton6.Caption = "periodic"
    ALLButton7.Caption = "template"
    ALLButton10.Caption = "Insert template"
    ALLButton8.Caption = "Cancel Action"
    ALLButton9.Caption = "Perview"
    ALLButton2.Caption = "Attachments"

    Command1.Caption = "Add to Explain Template"
    Command2.Caption = "Call Explain Template"

    EleOpt.Caption = "Show Of Accounts"
    Opt(0).Caption = "Hierarchy View"
    Opt(1).Caption = "Parent Path View"
    Opt(2).Caption = "Tabular View"
    ChkLastAccount.Caption = "Show Last Accounts Only"
    OptSort(0).Caption = "alphabetically"
    OptSort(1).Caption = "Charts sequence"

    With Fg_Journal
        .Cell(flexcpText, 0, .ColIndex("LineNo"), 1, .ColIndex("LineNo")) = "Line NO."
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = "Current Currency value"
        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "Credit"
    
        .Cell(flexcpText, 0, .ColIndex("DebitValueE"), 0, .ColIndex("CreditValueE")) = "Forign Currency value"
        .Cell(flexcpText, 1, .ColIndex("DebitValueE"), 1, .ColIndex("DebitValueE")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValueE"), 1, .ColIndex("CreditValueE")) = "Credit"
    
        '  .Cell(flexcpText, 0, .ColIndex("DebitValuee"), 0, .ColIndex("CreditValueE")) = "ValueE"
        '   .Cell(flexcpText, 1, .ColIndex("DebitValuee"), 1, .ColIndex("DebitValueE")) = "Debit"
        '   .Cell(flexcpText, 1, .ColIndex("CreditValuee"), 1, .ColIndex("CreditValueE")) = "Credit"
    
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Serial"
        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "Account Name"
        .Cell(flexcpText, 0, .ColIndex("Des"), 1, .ColIndex("Des")) = "Comment"
    
        .Cell(flexcpText, 0, .ColIndex("currenct_code"), 1, .ColIndex("currenct_code")) = "currency"
     
        .Cell(flexcpText, 0, .ColIndex("rate"), 1, .ColIndex("rate")) = "rate"
       
    End With

    LblDes.Caption = "Write your comment."
End Sub

Private Sub AddTip()

    Dim Wrap As String
    Dim Msg As String

    Wrap = CHR(13) + CHR(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "—ﬁ„ «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "—ﬁ„ «·ﬁÌœ «·Œ«’ »«·„” ‰œ"
            .AddControl TxtDEV_NO, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "„”·”·", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "„”·”· Â–« «·„” ‰œ ›Ï  Õ—Ì— «·ﬁÌÊœ"
            .AddControl TxtSerial, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "ﬁÌ„… «·”‰œ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "«·ﬁÌ„… «·√Ã„«·Ì… ··ﬁÌœ"
            .AddControl TxtValue, Msg, True
        End With

        With TTP
            .Create Me.hwnd, " «—ÌŒ «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = " «—ÌŒ  Õ—Ì— «·ﬁÌœ." & Wrap & "≈› —«÷Ì« ÌﬂÊ‰  «—ÌŒ «·ÌÊ„."
            .AddControl DTP_Date, Msg, True
        End With

        With TTP
            .Create Me.hwnd, " ⁄·Ìﬁ ⁄·Ï «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Ì„ﬂ‰ﬂ Â‰« ﬂ «»…  ⁄·Ìﬁ „‰«”»" & Wrap & "⁄·Ï Â–« «·Õ”«» ·ÌŸÂ— »ÃÊ«—Â" & Wrap & "›Ï ⁄„·Ì… „—«Ã⁄… «·ﬁÌÊœ √Ê " & Wrap & "«·ÿ»«⁄…."
            .AddControl TxtDes, Msg, True
        End With

        '
        With TTP
            .Create Me.hwnd, " ⁄·Ìﬁ ⁄·Ï «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "≈÷€ÿ Â‰« · ŸÂ— ·ﬂ ‰«›–…" & Wrap & " Õ—Ì— «· ⁄·Ìﬁ · ﬂ »  ⁄·Ìﬁ" & Wrap & "„‰«”» ⁄·Ï Â–« «·Õ”«»."
            .AddControl CboDes, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "⁄—÷ «·Õ”«» «·‰Â«∆Ï ›ﬁÿ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì„ﬂ‰ﬂ ÕÃ»" & Wrap & " «·Õ”«» «·—∆Ì”Ì… Ê≈ŸÂ«— «·Õ”«»« " & Wrap & "«·‰Â«∆Ì… Ê«· Ï Ì„ﬂ‰ﬂ  ”ÃÌ· " & Wrap & "«·ﬁÌÊœ ·Â«."
            .AddControl ChkLastAccount, Msg, True
        End With

        'OptSort
        With TTP
            .Create Me.hwnd, Opt(1).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ⁄—÷ «”„«¡ «·Õ”«»«  «· Ï " & Wrap & "Ì„ﬂ‰ﬂ ﬂ «»… Ê ”ÃÌ· «·ﬁÌœ ·Â«  ŸÂ— ›Ï " & Wrap & "‘ﬂ· ÃœÊ·Ï Ì⁄—÷ «”„ «·Õ”«» «·‰Â«∆Ï Ê«”„" & Wrap & "«·Õ”«» «·„ ›—⁄ „‰Â Ê«Ì÷« «”„ «·Õ”«» " & Wrap & "«·√⁄·Ï „‰Â( À·«À… „” ‰ÊÌ« )."
            .AddControl Opt(1), Msg, True
        End With

        With TTP
            .Create Me.hwnd, Opt(2).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ⁄—÷ «”„«¡ «·Õ”«»«  «· Ï " & Wrap & "Ì„ﬂ‰ﬂ ﬂ «»… Ê ”ÃÌ· «·ﬁÌœ ·Â«  ŸÂ— ›Ï " & Wrap & "‘ﬂ· ÃœÊ·Ï Ì⁄—÷ «”„ «·Õ”«» ›ﬁÿ."
            .AddControl Opt(2), Msg, True
        End With

        With TTP
            .Create Me.hwnd, Opt(0).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ⁄—÷ «”„«¡ «·Õ”«»«  «· Ï " & Wrap & "Ì„ﬂ‰ﬂ ﬂ «»… Ê ”ÃÌ· «·ﬁÌœ ·Â«  ŸÂ— ›Ï " & Wrap & "‘ﬂ· ‘Ã—Ï »«·Ÿ»ÿ „À· «·œ·Ì· «·„Õ«”»Ï."
            .AddControl Opt(0), Msg, True
        End With

        With TTP
            .Create Me.hwnd, OptSort(1).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· «”„«¡ «·Õ”«»« " & Wrap & " „— »… Õ”» „Êﬁ⁄Â« Ê — Ì»Â« " & Wrap & "««·œ·Ì· «·„Õ«”»Ï »«·Ÿ»ÿ. "
            .AddControl OptSort(1), Msg, True
        End With

        With TTP
            .Create Me.hwnd, OptSort(0).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· «”„«¡ «·Õ”«»« " & Wrap & " „— »…  —ÌÌ»« √»ÃœÌ« »€÷ " & Wrap & "«·‰Ÿ— ⁄‰ „Êﬁ⁄Â« ›Ï «·œ·Ì·" & Wrap & "«·„Õ«”»Ï."
            .AddControl OptSort(0), Msg, True
        End With

    Else

        With TTP
            .Create Me.hwnd, "DEV NO.", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "The serial of double entery voucher "
            .AddControl TxtDEV_NO, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Serial", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "The Serial of the voucher in the " & Wrap & "editing journals transactions"
            .AddControl TxtSerial, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Voucher Value", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "The total talue which will be" & Wrap & "recorded"
            .AddControl TxtValue, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Date", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Data of editing the voucher" & Wrap & "by default it is current ." & Wrap & "system date."
            .AddControl DTP_Date, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Comment", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Write your comment here to" & Wrap & " appear in auditing journal" & Wrap & "screen or in auditing report "
            .AddControl TxtDes, Msg, False
        End With

        '
        With TTP
            .Create Me.hwnd, "Write comment", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Click here to show the " & Wrap & "editing window to write" & Wrap & "your comment."
            .AddControl CboDes, Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkLastAccount.Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option when enabled shows" & Wrap & "the last accounts only."
            .AddControl ChkLastAccount, Msg, False
        End With

        'OptSort
        With TTP
            .Create Me.hwnd, Opt(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts" & Wrap & "in tabluar form !! and display " & Wrap & "the last three levels of chart" & Wrap & "of accounts."
            .AddControl Opt(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, Opt(2).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts" & Wrap & "in tabluar form !! and display" & Wrap & "just only the last account."
            .AddControl Opt(2), Msg, False
        End With

        With TTP
            .Create Me.hwnd, Opt(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts" & Wrap & "in hierarchy view exactly like" & Wrap & "the view of chart of accounts."
            .AddControl Opt(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, OptSort(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts " & Wrap & "sorted by it is index in the" & Wrap & "chart of accounts "
            .AddControl OptSort(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, OptSort(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This Option shows the accounts" & Wrap & "sorted alphabetically regardless " & Wrap & "it is index in the chart of " & Wrap & "accounts."
            .AddControl OptSort(0), Msg, False
        End With

    End If

End Sub

Public Function RefreshData() As Boolean

End Function

Public Property Get Cmd_Preview() As Boolean

    If Me.TxtNoteID.Text = "" Then
        GetMsgs 140, vbExclamation
        Cmd_Print = False
    Else
        Cmd_Print = FireReport(WindowTarget)
    End If

End Property

Public Property Let Cmd_Preview(ByVal vNewValue As Boolean)
    m_Cmd_Preview = vNewValue
End Property

Private Sub SaveData()
    Dim TransBegine As Boolean
    Dim Msg As String
    Dim i As Integer
    Dim StrSQL As String
    Dim RsTemp  As New ADODB.Recordset
    Dim RsNetes As New ADODB.Recordset
    Dim RsDev As New ADODB.Recordset
    Dim IntNoteType As Integer
    Dim StrInsertSQL  As String
    Dim IntAutoAccPost As Integer
    Dim StrPost As String
    Dim StrUnPost As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrPost = "„—Õ·"
        StrUnPost = "€Ì— „—Õ·"
    Else
        StrPost = "Posted"
        StrUnPost = "Not Posted"
    End If

    'On Error GoTo ErrTrap

    If val(TxtValue.Text) = 0 Then
        TxtValue.Text = 0
        '  Msg = "„‰ ›÷·ﬂ ﬁ„ »≈œŒ«· ﬁÌ„… «·”‰œ"
        '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '  'GetMsgs 59, vbExclamation
        '  TxtValue.SetFocus
        '  Exit Sub
    End If

    With Fg_Journal

        i = .FixedRows

        Do While i <= .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                .RemoveItem i
                i = i
            Else
                i = i + 1
            End If

        Loop

        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) = 0 And val(.TextMatrix(i, .ColIndex("CreditValue"))) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                
                        Msg = "«·Õ”«» " & .TextMatrix(i, .ColIndex("AccountName")) & CHR(13)
                        Msg = Msg & "·„  Õœœ ·Â Â· ÂÊ ÿ—› œ«∆‰ √Ê „œÌ‰.øø!!" & CHR(13)
                        Msg = Msg & "»—Ã«¡ ﬂ «»… ﬁÌ„… –·ﬂ «·Õ”«»"
                
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Else
                        Msg = "The Account " & .TextMatrix(i, .ColIndex("AccountName")) & CHR(13)
                        Msg = Msg & "not set as a Credit Or as Debit.??" & CHR(13)
                        Msg = Msg & "Please Write this account value.!"
                        MsgBox Msg, vbExclamation, App.title
                    End If

                    Exit Sub
                End If
            End If

        Next i

    End With

    If Me.TxtTotalCredit.Text <> Me.TxtTotalDebit.Text Then

        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Depit And Credit not matched ..!!" & CHR(13)
            Msg = Msg & "please correct this error."
        Else
            Msg = "ÿ—›Ï «·ﬁÌœ €Ì— „ “‰Ì‰ ..!!" & CHR(13)
            Msg = Msg & "„‰ ›÷·ﬂ ﬁ„ »„—«Ã⁄… ÿ—›Ï «·ﬁÌœ."
        End If

        'GetMsgs 60, vbExclamation
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    'If Val(Me.TxtValue.text) <> Val(Me.TxtTotalDebit.text) Then
    '    Msg = "ﬁÌ„… «·”‰œ €Ì— „ﬁ»Ê·… ..!!" & Chr(13)
    '    Msg = Msg & "„‰ ›÷·ﬂ ﬁ„ »„—«Ã⁄… ÿ—›Ï «·ﬁÌœ."
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    'GetMsgs 61, vbExclamation
    '    Exit Sub
    'End If
    '---------------------------Get the serial--------------
    If Me.TxtModFlg.Text = "N" Then
        ' Me.TxtSerial.text = ModAccounts.GetNewDEV_Serial(Me.DTP_Date.value)
    End If

    IntNoteType = 20

    Cn.BeginTrans
    TransBegine = True

    If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete   Notes Where Notes.NoteID='" & Trim(TxtNoteID.Text) & "'"
        Cn.Execute StrSQL, , adExecuteNoRecords
    
        If DcCostCenter.BoundText <> "" Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
    
    ElseIf Me.TxtModFlg.Text = "N" Then
        '---------------------------Get The Note ID ------------
        Me.TxtNoteID.Text = CStr(new_id("notes", "NoteID", ""))
        Me.TxtDEVID.Text = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
        Me.TxtDEV_NO.Text = Me.TxtDEVID.Text
        '---------------------------Begine of Saving------------
    End If

    Set RsNetes = New ADODB.Recordset
    RsNetes.Open "NOTES", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsNetes.AddNew
    RsNetes("NoteID").value = val(Me.TxtNoteID.Text)
    RsNetes("NoteType").value = 200
    RsNetes("NoteSerial").value = val(Me.TxtSerial.Text)
    RsNetes("numbering_type").value = sand_numbering_type(0) ' numbering_type
    RsNetes("sanad_year").value = year(DTP_Date.value)
    

    RsNetes("DueDate").value = Me.txtDueDate.value
    
    RsNetes("sanad_month").value = Month(DTP_Date.value)
    RsNetes("foxy_no").value = val(Text1.Text)
    RsNetes("NoteDate").value = Me.DTP_Date.value
    RsNetes("Note_Value").value = val(Me.TxtValue.Text)
    RsNetes("Double_Entry_Vouchers_ID").value = val(Me.TxtDEVID.Text)
    RsNetes("DAWRY").value = Check4.value
    RsNetes("KALEB").value = Check3.value
    
    RsNetes("Remark").value = Trim$(Me.Txt.Text)
    RsNetes("UserID").value = val(Me.DcboUsers.BoundText)
    Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.TxtTotalDebit.Text, "0.00"), 0, True, ".")
    RsNetes("note_value_by_characters").value = Trim$(Me.Lb_note_value_by_characters.Caption)
    RsNetes("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    
    If Me.dcprojects.BoundText <> "" Then
        Dim project_id As Integer
        project_id = IIf(Me.dcprojects.BoundText = "", 0, Me.dcprojects.BoundText)
        RsNetes("project_id").value = project_id
        Dim project_depit_or_credit As Integer
    
        If Option1.value = True Then
            project_depit_or_credit = 0
        Else
            project_depit_or_credit = 1
        End If
    
        RsNetes("project_depit_or_credit").value = project_depit_or_credit
    
    End If
    
    RsNetes.update
    Dim valuee As Long

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Single

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) > 0 Then
                    IntDEV_Type = 0
                    SngDEV_Value = val(.TextMatrix(i, .ColIndex("DebitValue")))
                Else
                    IntDEV_Type = 1
                    SngDEV_Value = val(.TextMatrix(i, .ColIndex("CreditValue")))
                End If
            
                project_id = IIf(Me.dcprojects.BoundText = "", 0, Me.dcprojects.BoundText)
            
                If IntDEV_Type = 0 And Option1.value = True Then
               
                ElseIf IntDEV_Type = 1 And Option2.value = True Then
            
                Else
                    project_id = 0
                End If
            
                If val(.TextMatrix(i, .ColIndex("DebitValuee"))) > 0 Then
               
                    valuee = val(.TextMatrix(i, .ColIndex("DebitValuee")))
                Else
                 
                    valuee = val(.TextMatrix(i, .ColIndex("CreditValuee")))
                End If
            
                If ModAccounts.AddNewDev(val(Me.TxtDEVID.Text), .TextMatrix(i, .ColIndex("LineNo")), .TextMatrix(i, .ColIndex("AccountCode")), SngDEV_Value, IntDEV_Type, CStr(.Cell(flexcpData, i, .ColIndex("Des"))), val(Me.TxtNoteID.Text), , , SystemOptions.SysCurrentAccountIntervalID, Me.DTP_Date.value, val(.TextMatrix(i, .ColIndex("userid"))), , Me.TxtSerial.Text, , valuee, .TextMatrix(i, .ColIndex("currenct_code")), val(.TextMatrix(i, .ColIndex("rate"))), , .TextMatrix(i, .ColIndex("dese")), IIf(.TextMatrix(i, .ColIndex("LineNo1")) <> "", .TextMatrix(i, .ColIndex("LineNo1")), setfoxy_Line), , project_id, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , (.TextMatrix(i, .ColIndex("DueDate")))) = False Then
                    GoTo ErrTrap
                End If
            End If

        Next i

    End With

    Cn.CommitTrans
    TransBegine = False

    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Saved"
    Else
        Msg = " „  ⁄„·Ì… «·Õ›Ÿ"
    End If

    save_cost_center

    'Õ›Ÿ „—ﬂ“ «· ﬂ·›… «·⁄«„
    If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "”‰œ ﬁÌœ", Me.DTP_Date.value
    End If

    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Me.TxtModFlg.Text = "R"
    '------------------------End of Saving--------------
    Exit Sub
ErrTrap:

    If TransBegine = True Then
        Cn.RollbackTrans
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "error During Saving"
    Else
        Msg = "⁄›Ê« ... ÕœÀ Œÿ« «À‰«¡ ⁄„·Ì… «·Õ›Ÿ."
    End If

    'Msg = Msg & Chr(13) & Err.Remark
    MsgBox Msg, vbExclamation, App.title
End Sub

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.Text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String
    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.Text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center

                If val(.TextMatrix(i, .ColIndex("DebitValue"))) = 0 Then
                    rs("value").value = .TextMatrix(i, .ColIndex("CreditValue"))
                    rs("depit_or_credit").value = "œ«∆‰"
            
                Else
                    rs("value").value = .TextMatrix(i, .ColIndex("DebitValue"))
                    rs("depit_or_credit").value = "„œÌ‰"
            
                End If
        
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
        
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs.update
        
            End If

        Next i

    End With

    rs.Close
End Function

Private Sub XPBtnMove_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Static StrOldTransID As String
    Dim StrSQL As String
    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     " From notes where (((notes.NoteType) =200)) " & _
     " ORDER BY NOTES.NoteID "
    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     "From notes where (((notes.NoteType)=200)) " & _
     "    ORDER BY NOTES.NoteID "
    
    StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & "From notes      ORDER BY NOTES.NoteID  "
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    If StrOldTransID <> "" Then
        rs.Find "NoteID=" & StrOldTransID & "", , adSearchForward, 1

        If rs.BOF Or rs.EOF Then
            rs.MoveFirst
        End If

    Else
        rs.MoveFirst
    End If

    Select Case Index

        Case 1 'First

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst
            End If

        Case 0 'Previous

            If Not (rs.BOF Or rs.EOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveNext
            End If

        Case 3 'NEXT

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MovePrevious
            End If

        Case 2 'Last

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

    End Select

    If Not (IsNull(rs("NoteID").value)) Then
        Me.Retrive rs("NoteID").value
        StrOldTransID = rs("NoteID").value
    End If

    rs.Close
    Set rs = Nothing
End Sub


Private Sub chkAll_Click()
    
    If chkAll.value = vbChecked Then
    
        With Fg_Journal
            Dim i As Long
            For i = 2 To .Rows - 1
                If .TextMatrix(i, .ColIndex("AccountName")) <> "" Then
                    .TextMatrix(i, .ColIndex("DueDate")) = txtDueDate.value
                End If
            Next
        End With
    End If
End Sub


