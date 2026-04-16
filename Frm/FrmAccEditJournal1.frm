VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
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
Begin VB.Form FrmAccEditJournal1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ ﬁÌœ «›  «ÕÌ"
   ClientHeight    =   8700
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11520
   HelpContextID   =   450
   Icon            =   "FrmAccEditJournal1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8700
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11520
      _cx             =   20320
      _cy             =   15346
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
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmAccEditJournal1.frx":030A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic EleTop 
         Height          =   660
         Left            =   15
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   15
         Width           =   11490
         _cx             =   20267
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
         BackColor       =   12648447
         ForeColor       =   8421376
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "”‰œ ﬁÌœ «›  «ÕÌ"
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
         Begin VB.TextBox TxtSerial 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   0
            Visible         =   0   'False
            Width           =   2220
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1245
            TabIndex        =   13
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal1.frx":0363
            ColorButton     =   12648447
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
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal1.frx":06FD
            ColorButton     =   12648447
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
            TabIndex        =   15
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal1.frx":0A97
            ColorButton     =   12648447
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
            TabIndex        =   16
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAccEditJournal1.frx":0E31
            ColorButton     =   12648447
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
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   5040
            Picture         =   "FrmAccEditJournal1.frx":11CB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   5880
         Left            =   15
         TabIndex        =   1
         Top             =   1635
         Width           =   11490
         _cx             =   20267
         _cy             =   10372
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
         Caption         =   "«·ﬁÌÊœ|«·‘—Õ «·⁄«„|Õ«·… «·«⁄ „«œ"
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
            Height          =   5790
            Index           =   0
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   45
            Width           =   10470
            _cx             =   18468
            _cy             =   10213
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
            _GridInfo       =   $"FrmAccEditJournal1.frx":4E33
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic EleOpt 
               Height          =   945
               Left            =   2640
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   4815
               Visible         =   0   'False
               Width           =   2580
               _cx             =   4551
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
                  Left            =   -390
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   0
                  Width           =   17205
                  Begin VB.CommandButton Command6 
                     Caption         =   "Command6"
                     Height          =   375
                     Left            =   2040
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   600
                     Width           =   975
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ÃœÊ·Ï"
                     Height          =   285
                     Index           =   2
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   600
                     Value           =   -1  'True
                     Width           =   1455
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ÿ«„ «·‘Ã—Ï"
                     Height          =   270
                     Index           =   0
                     Left            =   600
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     Top             =   390
                     Width           =   1455
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ÿ«„ «·„”«—"
                     Height          =   270
                     Index           =   1
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   120
                     Width           =   1575
                  End
               End
               Begin C1SizerLibCtl.C1Elastic EleSortOpt 
                  Height          =   540
                  Left            =   14730
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   285
                  Width           =   38220
                  _cx             =   67416
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
                     Left            =   -1740
                     RightToLeft     =   -1  'True
                     TabIndex        =   3
                     Top             =   -90
                     Value           =   -1  'True
                     Width           =   21105
                  End
               End
               Begin VB.Image ImgNote 
                  Height          =   240
                  Left            =   120
                  Picture         =   "FrmAccEditJournal1.frx":4E9E
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   240
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   4755
               Left            =   30
               TabIndex        =   2
               Top             =   30
               Width           =   10410
               _cx             =   18362
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
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   10
               Rows            =   10
               Cols            =   26
               FixedRows       =   2
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmAccEditJournal1.frx":5428
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
                  TabIndex        =   37
                  Top             =   3720
                  Visible         =   0   'False
                  Width           =   4215
                  Begin VB.CommandButton Command5 
                     Caption         =   "‰”Œ"
                     Height          =   255
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   720
                     Width           =   1215
                  End
                  Begin VB.TextBox Text4 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   240
                     Width           =   2175
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—ﬁ„ «·ﬁÌœ"
                     Height          =   255
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
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
                  TabIndex        =   10
                  Top             =   750
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.TextBox TxtDese 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1485
                     Left            =   0
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   42
                     Top             =   2040
                     Width           =   8955
                  End
                  Begin VB.TextBox txtcodesub 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   3600
                     Width           =   855
                  End
                  Begin VB.CommandButton Command4 
                     Caption         =   "«÷«›… ‘—Õ"
                     Height          =   255
                     Left            =   7440
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   3600
                     Width           =   1350
                  End
                  Begin VB.CommandButton Command3 
                     Caption         =   "«” œ⁄«¡ ‘—Õ"
                     Height          =   255
                     Left            =   6240
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   3600
                     Width           =   1095
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                     Height          =   3900
                     Left            =   120
                     TabIndex        =   43
                     TabStop         =   0   'False
                     Top             =   150
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
                        Height          =   1245
                        Left            =   0
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   44
                        Top             =   480
                        Visible         =   0   'False
                        Width           =   8955
                     End
                     Begin VB.Label Label10 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "X"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   12
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H000000FF&
                        Height          =   420
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   49
                        Top             =   0
                        Width           =   255
                     End
                     Begin VB.Label LblDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000C&
                        Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                        ForeColor       =   &H0000C8FF&
                        Height          =   315
                        Left            =   6840
                        RightToLeft     =   -1  'True
                        TabIndex        =   45
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
                     TabIndex        =   35
                     Top             =   3480
                     Width           =   735
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     Height          =   495
                     Left            =   1560
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   1200
                     Width           =   975
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Code"
                     Height          =   255
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   1320
                     Width           =   735
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   9
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
                  Picture         =   "FrmAccEditJournal1.frx":5873
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
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   4845
               Width           =   10470
               _cx             =   18468
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
               AutoSizeChildren=   7
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
                  Left            =   75
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   0
                  Width           =   2700
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "  — Ì» »«·œ·Ì· «·„Õ«”»Ì"
                     Height          =   270
                     Index           =   1
                     Left            =   600
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   480
                     Width           =   1995
                  End
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "  — Ì» «»ÃœÏ"
                     Height          =   270
                     Index           =   0
                     Left            =   1080
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   240
                     Width           =   1515
                  End
                  Begin ALLButtonS.ALLButton CmdRemove 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   48
                     Tag             =   "Delete Row"
                     Top             =   120
                     Width           =   855
                     _ExtentX        =   1508
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
                     MICON           =   "FrmAccEditJournal1.frx":5E0D
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
               End
               Begin DBPIXLib.DBPix20 DBPix202 
                  Height          =   30
                  Left            =   1890
                  TabIndex        =   20
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   4275
                  _Version        =   131072
                  _ExtentX        =   7541
                  _ExtentY        =   53
                  _StockProps     =   1
                  BackColor       =   16777215
                  _Image          =   "FrmAccEditJournal1.frx":5E29
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
               Begin VB.Label lblAccountBalance 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   855
                  Left            =   5250
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   0
                  Width           =   5130
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ÊﬁÌ⁄"
                  Height          =   240
                  Index           =   5
                  Left            =   3285
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Tag             =   "51"
                  Top             =   0
                  Width           =   810
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5790
            Index           =   1
            Left            =   12135
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   10470
            _cx             =   18468
            _cy             =   10213
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
               Height          =   555
               Left            =   26340
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   6210
               Width           =   3075
            End
            Begin VB.CommandButton Command2 
               Caption         =   "«” œ⁄«¡ ﬁ«·» ‘—Õ"
               Height          =   735
               Left            =   15315
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   7410
               Width           =   6240
            End
            Begin VB.CommandButton Command1 
               Caption         =   "«÷«›… ﬁ«·» ‘—Õ"
               Height          =   735
               Left            =   22560
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   7410
               Width           =   6360
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   4965
               Left            =   75
               MaxLength       =   1000
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   690
               Width           =   10230
            End
            Begin VB.Label Lb_note_value_by_characters 
               Alignment       =   1  'Right Justify
               Height          =   645
               Left            =   13545
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   8820
               Width           =   16335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Code"
               Height          =   765
               Left            =   23055
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   6210
               Width           =   2730
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄·Ìﬁ:"
               Height          =   255
               Index           =   6
               Left            =   23400
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Tag             =   "22"
               Top             =   690
               Width           =   6315
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   5790
            Left            =   12435
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   45
            Width           =   10470
            _cx             =   18468
            _cy             =   10213
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   3750
               Left            =   120
               TabIndex        =   81
               Tag             =   "1"
               Top             =   120
               Width           =   10695
               _cx             =   18865
               _cy             =   6615
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
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmAccEditJournal1.frx":5E41
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
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   4440
               Width           =   3375
            End
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   5880
               Width           =   3375
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   1155
         Left            =   15
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   7530
         Width           =   11490
         _cx             =   20267
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
         Begin VB.TextBox TxtTotalCredit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   435
            Left            =   4110
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   75
            Width           =   1485
         End
         Begin VB.TextBox TxtTotalDebit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   435
            Left            =   7200
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   75
            Width           =   1305
         End
         Begin VB.TextBox TXTResults 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   420
            Left            =   2145
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   75
            Width           =   1395
         End
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Left            =   15
            TabIndex        =   55
            Top             =   75
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            BackColor       =   12648447
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   0
            Left            =   10515
            TabIndex        =   56
            Top             =   645
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
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
            Height          =   300
            Index           =   1
            Left            =   9435
            TabIndex        =   57
            Top             =   645
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
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
            Height          =   300
            Index           =   2
            Left            =   8385
            TabIndex        =   58
            Top             =   645
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   529
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
            Height          =   300
            Index           =   3
            Left            =   7500
            TabIndex        =   59
            Top             =   645
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
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
            Height          =   300
            Index           =   4
            Left            =   5835
            TabIndex        =   60
            Top             =   645
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   529
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
            Height          =   300
            Index           =   5
            Left            =   5040
            TabIndex        =   61
            Top             =   645
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
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
            Height          =   300
            Index           =   6
            Left            =   2160
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   645
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
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
            Height          =   300
            Index           =   7
            Left            =   4140
            TabIndex        =   63
            Top             =   645
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   529
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
            Height          =   300
            Left            =   3135
            TabIndex        =   64
            Top             =   645
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
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
            Height          =   300
            Index           =   8
            Left            =   6675
            TabIndex        =   65
            Top             =   645
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
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
            Height          =   285
            Left            =   8610
            TabIndex        =   66
            Top             =   795
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":5F84
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
            Height          =   285
            Left            =   8115
            TabIndex        =   67
            Top             =   795
            Visible         =   0   'False
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":5FA0
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
            Height          =   285
            Left            =   5070
            TabIndex        =   68
            Top             =   795
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":5FBC
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
            Height          =   285
            Left            =   3000
            TabIndex        =   69
            Top             =   795
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":5FD8
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
            Height          =   285
            Left            =   1575
            TabIndex        =   70
            Top             =   795
            Visible         =   0   'False
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":5FF4
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
            Height          =   285
            Left            =   4080
            TabIndex        =   71
            Top             =   795
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":6010
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
            Height          =   285
            Left            =   9900
            TabIndex        =   72
            Top             =   795
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "„—«ﬂ“ «· ﬂ·›…"
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
            MICON           =   "FrmAccEditJournal1.frx":602C
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
            Height          =   285
            Left            =   195
            TabIndex        =   73
            Top             =   795
            Visible         =   0   'False
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
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
            MICON           =   "FrmAccEditJournal1.frx":6048
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
            Height          =   285
            Left            =   6495
            TabIndex        =   74
            Top             =   795
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "«” œ⁄«¡ ﬁÌœ œÊ—Ï"
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
            MICON           =   "FrmAccEditJournal1.frx":6064
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   345
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
            BackColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   -2147483635
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«÷€ÿ »«·“— «·«Ì„‰ ·⁄—÷ ﬂ‘› «·Õ”«»"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   0
            Width           =   2040
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ…"
            Height          =   150
            Index           =   8
            Left            =   1290
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Tag             =   "51"
            Top             =   210
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ÿ—› «·œ«∆‰"
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   5670
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Tag             =   "56"
            Top             =   210
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ÿ—› «·„œÌ‰"
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Tag             =   "55"
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—ﬁ"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   13
            Left            =   3615
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Tag             =   "56"
            Top             =   180
            Width           =   390
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   930
         Left            =   15
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   690
         Width           =   11490
         _cx             =   20267
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
         Begin VB.CheckBox chkAll 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
            Height          =   285
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   630
            Width           =   675
         End
         Begin VB.TextBox TxtDEV_NO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   8760
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   780
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.TextBox TxtDEVID 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   405
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8130
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   60
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox TxtValue 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   1005
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   330
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   75
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Frame Frame17 
            Height          =   855
            Left            =   -6660
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Visible         =   0   'False
            Width           =   8625
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               Caption         =   "„·€Ì"
               Height          =   195
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   480
               Width           =   1335
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌœ œÊ—Ì"
               Height          =   195
               Left            =   -240
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   600
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁ«·»"
               Height          =   195
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «⁄ „«œÂ"
               Height          =   195
               Left            =   900
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œÌ„ «· √ÀÌ—"
               Height          =   195
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Text            =   "Text1"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CheckBox ChkLastAccount 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ «·Õ”«» «·‰Â«∆Ï ›ﬁÿ"
               Height          =   270
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   480
               Value           =   1  'Checked
               Width           =   2955
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   120
               Width           =   1575
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4200
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   88
               Top             =   480
               Width           =   5295
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "„’œ— «·ﬁÌœ"
               Height          =   255
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "»‰«¡ ⁄·Ï"
               Height          =   255
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   480
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   8745
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   0
            Width           =   1770
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   12600
            Top             =   960
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic ElePost 
            Height          =   450
            Left            =   405
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   900
            Visible         =   0   'False
            Width           =   2625
            _cx             =   4630
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
               TabIndex        =   105
               Top             =   45
               Width           =   1485
            End
            Begin VB.Image Img 
               Height          =   225
               Index           =   0
               Left            =   90
               Top             =   90
               Width           =   270
            End
            Begin VB.Image Img 
               Height          =   180
               Index           =   1
               Left            =   1635
               Top             =   285
               Width           =   285
            End
         End
         Begin MSComCtl2.DTPicker DTP_Date 
            Height          =   330
            Left            =   8775
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   435
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   143589377
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   4845
            TabIndex        =   107
            Top             =   0
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmAccEditJournal1.frx":6080
            Height          =   315
            Left            =   4845
            TabIndex        =   108
            Top             =   360
            Width           =   2070
            _ExtentX        =   3651
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
         Begin Dynamic_Byte.NourHijriCal DtHijriTrans 
            Height          =   255
            Left            =   2040
            TabIndex        =   109
            Top             =   360
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker txtDueDate 
            Height          =   300
            Left            =   3000
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   143589377
            CurrentDate     =   37140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   180
            Index           =   16
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Tag             =   "53"
            Top             =   675
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   330
            Index           =   7
            Left            =   10515
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Tag             =   "57"
            Top             =   600
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·”‰œ"
            Height          =   270
            Index           =   4
            Left            =   7425
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Tag             =   "54"
            Top             =   1020
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   270
            Index           =   3
            Left            =   10500
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Tag             =   "53"
            Top             =   120
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   315
            Index           =   0
            Left            =   10500
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Tag             =   "52"
            Top             =   495
            Width           =   945
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄ «·⁄«„"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6885
            TabIndex        =   112
            Top             =   0
            Width           =   885
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌœ «·Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   3375
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   0
            Width           =   750
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
            Height          =   255
            Left            =   7005
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   360
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "FrmAccEditJournal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim line_no1 As Double
Dim last_line_id As Double
Dim numbering_type As Integer
Dim TTP As New clstooltip
Dim BolEditOnMainAccounts As Boolean
Dim PicHeight As Long
Dim PicWidth As Long
Dim Dcombos As ClsDataCombos
Dim DCboSearch As New clsDCboSearch
  Dim Rs1 As New ADODB.Recordset
  Dim ScreenNameArabic As String
Public LngRow As Long
Dim ScreenNameEnglish As String

Private Enum PrintTarget
    WindowTarget
    PrinterTarget
End Enum
Dim FirstPeriodDateInthisYear  As Date

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·ﬁÌœ  " & TxtSerial1.Text & CHR(13) & "   «· «—ÌŒ  " & DTP_Date & CHR(13) & "   «·›—⁄ «·⁄«„   " & dcBranch & CHR(13) & "     «·«Ã„«·Ì    " & TxtTotalDebit
       '
                     
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr No     " & TxtSerial1.Text & CHR(13) & "   Date  " & DTP_Date & CHR(13) & "   General Branch  " & dcBranch & CHR(13) & "     Total    " & TxtTotalDebit
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtSerial)
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(TxtSerial)
    End If
    
End Function

Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
        
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 20) = &HFFFFC0
            Else
                .Cell(flexcpBackColor, i, 1, i, 20) = vbWhite
            End If

        Next i

    End With

    line_no1 = IntCounter

End Sub

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
 
If val(TxtNoteID.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
     
    SendTopost Me.Name, "Notes", "NoteID", 0, val(dcBranch.BoundText), val(TxtNoteID.Text), TxtSerial.Text
  '' RsNetes.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
fillapprovData
End Sub
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.TxtNoteID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label24.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label24.Caption = "Approved"
                                 End If
                            Label24.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label24.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label24.Caption = "Currently required Approve"
                            End If
                 Label24.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.Rows = 1
    End If
RsDetails.Close

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
    Unload marakes_taklefa_tawze3
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
marakes_taklefa_tawze3.DTP_Date.value = DTP_Date.value
        marakes_taklefa_tawze3.opr_type = "”‰œ ﬁÌœ «›  «ÕÌ "
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
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial, "1608201801"

Exit Sub
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
    ked_dawry.ID = TxtDEV_NO.Text
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
    'On Error Resume Next
    'Form3.Show
 
    'Form3.case_id = 16
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
            TxtDes.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("des"))
            TxtDese.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("dese"))
    
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

Function setfoxy_Line() As Double
    
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
        
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            SetForNew
            Label9.Visible = False
            Me.TxtModFlg.Text = "N"
            setfoxy
            DcCostCenter.Text = ""
            Accredit.Caption = ""
            Me.dcBranch.BoundText = branch_id
             Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.Rows = 1
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '    DcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            Me.Fg_Journal.Editable = flexEDKbdMouse

            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.DTP_Date.value = FirstPeriodDateInthisYear

        Case 1
         
             If ScreenAproved(val(TxtNoteID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
  


           If ChekClodePeriod(DTP_Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.TxtNoteID.Text) = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬁÌœ «·Ì «»œ«", vbCritical
                Else
                    MsgBox "Can't Edit", vbCritical
                End If

                Exit Sub
            End If
    
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.DTP_Date.value = FirstPeriodDateInthisYear
    
            Me.TxtModFlg.Text = "E"
  
            Fg_Journal.Rows = Fg_Journal.Rows + 1
 
            'TxtSerial.text = year(DTP_Date.value) & 1
            'TxtSerial1.text = TxtSerial.text
   
            CuurentLogdata

        Case 2
           If ChekClodePeriod(DTP_Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
            If val(TxtTotalDebit.Text) = 0 And val(TxtTotalCredit.Text) = 0 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = " There is no iAccounts in vouchers"
                Else
                    Msg = "·„ Ì „ «œŒ«· Õ”«»«  ›Ì «·ﬁÌœ"
                End If

                MsgBox Msg, vbCritical
                Exit Sub
            End If

            '  Me.DcboUsers.BoundText = user_id
            If Me.TxtModFlg.Text = "N" Then
                my_branch = val(Me.dcBranch.BoundText)
        
                If TxtSerial1.Text = "" Then
                    If OpeningVoucher_coding(val(my_branch), DTP_Date.value, 3, 101) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁÌœ «›  «ÕÌ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                        Else
                        MsgBox "Code Exceding   ": Exit Sub
                        End If
                    Else
                   
                        If OpeningVoucher_coding(val(my_branch), DTP_Date.value, 3, 101) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                          Else
                          MsgBox "Enter Voucher Code Manually ": Exit Sub
                          End If
                        Else
                            TxtSerial1.Text = OpeningVoucher_coding(val(my_branch), DTP_Date.value, 3, 101)
                            TxtSerial.Text = TxtSerial1.Text
                        End If
                    End If
                End If
                  
            End If

            SaveData

        Case 3
            Undo
        
        Case 4
            Frame3.Visible = True
      
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
        Unload Voucher_search1
            Voucher_search1.case_id = 3
            Voucher_search1.show
            'Voucher_search.Show

        Case 6
            Unload Me

        Case 7
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
            ShowGL_ccOpening TxtSerial.Text, , 200, val(Me.TxtNoteID.Text)

        Case 8
        
      If ScreenAproved(val(TxtNoteID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If



           If ChekClodePeriod(DTP_Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If Me.TxtNoteID.Text = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰ Õ–›  ﬁÌœ «·Ì «»œ«", vbCritical
                Else
                    MsgBox "Can't Delete", vbCritical
                End If

                Exit Sub
            End If
    
            Del_Trans
    End Select

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If TxtNoteID.Text <> "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·ﬁÌœ —ﬁ„ " & CHR(13)
        Msg = Msg + (Me.TxtSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

Else
Msg = Msg + " Confirm Deletion?"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            CuurentLogdata ("D")
                Deletepost Me.Name, "Notes", "NoteID", 0, val(dcBranch.BoundText), val(TxtNoteID.Text), TxtSerial.Text
                
   
            StrSQL = "Delete  Notes1  where NoteID =" & val(TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
  
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            Dim rs As New ADODB.Recordset

            StrSQL = "SELECT NOTES1.NoteID, NOTES1.NoteType " & "From notes1 where (((notes1.NoteType)=101)) " & "    ORDER BY NOTES1.NoteID "
    
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
           
            If rs.RecordCount < 1 Then
                clear_all Me
                '  Fg_Journal.Clear flexClearScrollable, flexClearEverything
                
                TxtModFlg_Change
               
                Fg_Journal.Clear flexClearScrollable, flexClearEverything
                Me.TxtTotalCredit.Text = 0
                Me.TxtTotalDebit.Text = 0
                Me.TXTResults.Text = 0
            Else

                If Not (IsNull(rs("NoteID").value)) Then
                    Me.Retrive rs("NoteID").value
                    StrOldTransID = rs("NoteID").value
                End If

            End If
        
        End If

    Else
        'clear_all Me
                                    If SystemOptions.UserInterface = ArabicInterface Then

        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
Else
        Msg = "No Record To Delete"
End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "error During Delete " & CHR(13)
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap
    Dim sgl As String

    Select Case TxtModFlg.Text

        Case "N"
            sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (2)
        SetForNew
        Case "E"
            sgl = "delete  marakes_taklefa_temp  where ok is null and  kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
        
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

    sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
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
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                 
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
    End With
            
End Sub

Private Sub Command1_Click()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
 '   rs.Open "[ked_desc]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    StrSQL = "SELECT  *  from ked_desc Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
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
  '  rs.Open "[ked_desc]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     StrSQL = "SELECT  *  from ked_desc Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
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

Private Sub Dcbranch_Click(Area As Integer)

TxtSerial.Text = ""


End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, _
                                       ByVal Col As Long)

    With Me.Fg_Journal

        Select Case .ColKey(Col)

            Case "CC"
                ALLButton1_Click
            Case "DueDate"
                Dim Frm As New FrmDateOpProject
                
                Frm.Index = 541
                Me.LngRow = Row
                Frm.show 1
        End Select

    End With

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
  Case "project"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("projectid")) = StrAccountCode
                '.TextMatrix(Row, .ColIndex("oper")) = ""
                '.TextMatrix(Row, .ColIndex("pand")) = ""
                If StrAccountCode <> "" Then
                StrSQL = "Select Fullcode from projects where id =" & val(StrAccountCode) & " "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("ProjectCode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                Else
                .TextMatrix(Row, .ColIndex("ProjectCode")) = ""
                End If
                End If
                Case "ProjectCode"
                '.TextMatrix(Row, .ColIndex("pand")) = ""
                '.TextMatrix(Row, .ColIndex("oper")) = ""
                If .TextMatrix(Row, .ColIndex("ProjectCode")) <> "" Then
                StrSQL = "Select  * from projects where Fullcode ='" & .TextMatrix(Row, .ColIndex("ProjectCode")) & "' "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("projectid")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
                Else
                .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(rs("Project_nameE").value), "", rs("Project_nameE").value)
                End If
                Else
                .TextMatrix(Row, .ColIndex("projectid")) = ""
                .TextMatrix(Row, .ColIndex("project")) = ""
                End If
                End If
            Case "BranchName"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("BranchId")) = StrAccountCode
        
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
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
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
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
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
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
            
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("BranchId")) = IIf(val(Me.dcBranch.BoundText) = 0, 1, val(Me.dcBranch.BoundText))
                .TextMatrix(Row, .ColIndex("BranchName")) = Me.dcBranch.Text

                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where  ( ACCOUNTS.Block=0 or  ACCOUNTS.Block is null)  and  ACCOUNTS.AccountTypes<>2 and  ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                
                StrSQL = StrSQL & GetAccountByBarnchUser
                StrSQL = StrSQL & GetAccountCodeHiding
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        If LastAccount(rs("Account_Code").value) = False Then
                            .TextMatrix(Row, Col) = ""
                            .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                            .TextMatrix(Row, .ColIndex("AccountName")) = ""
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
 
  If rs2.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
  Else
     .TextMatrix(Row, .ColIndex("currenct_code")) = 1
                    
                    .TextMatrix(Row, .ColIndex("rate")) = 1
  
  End If
  
  
 '                   .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
 '                   .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                   ' GetMsgs 130, vbExclamation
                    
                  If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "—ﬁ„ Õ”«» €Ì— ’ÕÌÕ", vbCritical
                  Else
                        MsgBox "Account Code  not Exist ", vbCritical
                  End If
                  
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    .TextMatrix(Row, .ColIndex("AccountName")) = ""
                    
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
        
                sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("BranchId")) = val(Me.dcBranch.BoundText)
                .TextMatrix(Row, .ColIndex("BranchName")) = Me.dcBranch.Text

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
                StrSQL = StrSQL & GetAccountByBarnchUser
                StrSQL = StrSQL & GetAccountCodeHiding
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
 
        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·Õ”«» «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("DebitValue") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„… «·„œÌ‰…   «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("DebitValue")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change  debit value" & .Cell(flexcpTextDisplay, Row, .ColIndex("DebitValue")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
        ElseIf Col = .ColIndex("CreditValue") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„… «·œ«∆‰…   «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("CreditValue")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change  Credit value" & .Cell(flexcpTextDisplay, Row, .ColIndex("CreditValue")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
 
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
        ElseIf Col = .ColIndex("BranchName") Then
            LogTextA = "   ⁄œÌ· «·›—⁄  «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("BranchName")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change Branch Name " & .Cell(flexcpTextDisplay, Row, .ColIndex("BranchName")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
        
        End If

        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), TxtSerial1

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
                .ComboList = ""
            Case "ProjectCode"
                .ComboList = ""
                ' Cancel = True
        End Select

    End With

End Sub

Private Sub Fg_Journal_Click()
    On Error Resume Next
With Fg_Journal
lblAccountBalance.Caption = GetbalanceBar(.TextMatrix(.Row, .ColIndex("AccountCode")))
End With

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

        If Fg_Journal.ColKey(c) <> "Des" And Fg_Journal.ColKey(c) <> "Dese" Then
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

        TxtDes.Text = Fg_Journal.TextMatrix(r, Fg_Journal.ColIndex("des"))
        TxtDese.Text = Fg_Journal.TextMatrix(r, Fg_Journal.ColIndex("dese"))
        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus

        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c

        If SystemOptions.UserInterface = ArabicInterface Then
            '    TxtDes.SetFocus
        Else
            '    TxtDese.SetFocus
        End If
    
    End With

End Sub

Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
Exit Sub

  '  SendKeys "{F4}"
If Me.TxtModFlg = "R" Then
Exit Sub
End If

    SendKeys "{F4}"
SendKeys "{BACKSPACE}"
SendKeys CHR(KeyAscii)

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyF5 Then
 
        update_accounts
    End If

    If KeyCode = vbKeyF9 Then
          With Fg_Journal
            
                    If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
             
                   .TextMatrix(.Row, .ColIndex("Des")) = .TextMatrix(.Row - 1, .ColIndex("Des"))
                    End If
            
                End With
   End If
    
    
    If KeyCode = 46 Then
        CmdRemove_Click
    End If

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 2001

    End If

    If KeyCode = vbKeyReturn Then

        With Fg_Journal

            If .Col = 7 And val(.TextMatrix(.Row, 7)) = 0 Then
                .Col = .Col + 2
            ElseIf .Col = 7 And val(.TextMatrix(.Row, 7)) <> 0 Then
                .Row = .Row + 1
                .Col = 5
           
            ElseIf .Col = 9 Then
                .Row = .Row + 1
                .Col = 5
            Else
                .Col = .Col + 1
            End If

            .ShowCell .Row, .Col + 1
            
            .SetFocus
        End With

    End If

End Sub

Private Sub Fg_Journal_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    With Fg_Journal

        If Button = vbRightButton Then
        
                   Dim FirstPeriod As Date
     Dim AccountName As String
      Dim AccountCode As String

                    If .FixedRows <= .Row And .Row < .Rows - 1 Then
                       If .TextMatrix(.Row, .ColIndex("AccountCode")) <> "" Then
                             AccountCode = .TextMatrix(.Row, .ColIndex("AccountCode"))
      AccountName = .TextMatrix(.Row, .ColIndex("AccountName"))
      'AccountName
      
            getFirstPeriodDateInthisYear FirstPeriod
            Get_Account_name
             ShowReport AccountCode, AccountName, FirstPeriod, Date
             
             
                       
                       
                        End If
                        
               End If
                         
            '        End If
            
            
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
            StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*ParentName", "Account_Code")
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
Opt(1).value = True
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
         Case "project"

                StrSQL = " SELECT     Project_name,Project_nameE , id From dbo.Projects "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Project_name", "id")
         Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Project_nameE", "id")
End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
                
            Case "BranchName"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "  select branch_id,branch_name from TblBranchesData   "
                Else
                    StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
                End If

                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "branch_name", "branch_id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            Case "AccountName"

                If Opt(0).value = True Then
                    'Tree display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, Space(2*(Len(Account_Code)))" & "+ ACCOUNTS.Account_Name   As DisName , ACCOUNTS.Parent_Account_Code," & "ACCOUNTS.last_account, ACCOUNTS.cannot_del" & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <> 'r' "
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
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
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1 and ACCOUNTS.AccountTypes<>2)"
                            End If
                        End If
                       StrSQL = StrSQL & GetAccountByBarnchUser
                       StrSQL = StrSQL & GetAccountCodeHiding
                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                        End If
                
                    Else
                
                        '    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & _
                             "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & _
                             " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & _
                             "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & _
                             "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & _
                             "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code,  REPLACE(REPLACE(REPLACE(ACCOUNTS.Account_Name, CHAR(10), ''), CHAR(13), ''), CHAR(9), '')  As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                        StrSQL = StrSQL & GetAccountByBarnchUser
                        StrSQL = StrSQL & GetAccountCodeHiding
                        
                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(         ( ACCOUNTS.Block=0 or  ACCOUNTS.Block is null)     and ACCOUNTS.last_account=1)"
                            End If
                        End If
 StrSQL = StrSQL + " And(ACCOUNTS.last_account=1 and ACCOUNTS.AccountTypes<>2)"
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
                     StrSQL = StrSQL & GetAccountByBarnchUser
                     StrSQL = StrSQL & GetAccountCodeHiding
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
    
    ScreenNameArabic = "”‰œ ﬁÌœ «›  «ÕÌ"
    ScreenNameEnglish = "Opening Balance Ge"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim GrdBck As New ClsBackGroundPic

'    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
'    fill_combo Me.DcCostCenter, StrSQL


    Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
Dcombos.GetCostCenter DcCostCenter



     If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select branch_id,branch_name from   TblBranchesData where branch_id in(" & Current_branchSql & ")    "
    Else
        StrSQL = "  select branch_id,branch_namee from TblBranchesData   where branch_id in(" & Current_branchSql & ")    "
    End If


    fill_combo dcBranch, StrSQL

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
'    SetDtpickerDate Me.DTP_Date
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
        .ColWidth(.ColIndex("AccountName")) = 4500
    
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = " ﬁÌ„… «·ﬁÌœ »«·⁄„·… «·„Õ·Ì… "
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter

        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "„œÌ‰"
        .ColWidth(.ColIndex("DebitValue")) = 1590
        .ColFormat(.ColIndex("DebitValue")) = "#,###.00"
     
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "œ«∆‰"
        .ColWidth(.ColIndex("CreditValue")) = 1590
        .ColFormat(.ColIndex("CreditValue")) = "#,###.00"
    
        .Cell(flexcpText, 0, .ColIndex("DebitValueE"), 0, .ColIndex("CreditValueE")) = " ﬁÌ„… «·ﬁÌœ »«·⁄„·… «·«Ã‰»Ì… "
    
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter
        
        .Cell(flexcpText, 1, .ColIndex("DebitValueE"), 1, .ColIndex("DebitValueE")) = "„œÌ‰"
        .Cell(flexcpText, 1, .ColIndex("CreditValueE"), 1, .ColIndex("CreditValueE")) = "œ«∆‰"
        .ColFormat(.ColIndex("DebitValueE")) = "#,###.00"
        .ColFormat(.ColIndex("CreditValueE")) = "#,###.00"

        '.MergeCol(.ColIndex("Des")) = True
        '.Cell(flexcpText, 0, .ColIndex("Des"), 1, .ColIndex("Des")) = "«·‘—Õ"
        '.ColWidth(.ColIndex("Des")) = 2200
        Set .WallPaper = GrdBck.Picture
        .ColComboList(.ColIndex("CC")) = "..."
 
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
    
        StrSQL = "SELECT NOTES1.NoteID, NOTES1.NoteType " & "From notes1 where   notes1.NoteType =-1 "
    
 
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    
    'Resize_Form Me,    TransactionSize
   ' XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

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
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Label10_Click()
    PicDes.Visible = False
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
    TxtSerial1.Text = ""
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
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Me.TxtTotalCredit.Text = 0
    Me.TxtTotalDebit.Text = 0
    Me.TXTResults.Text = 0
    Me.DcboUsers.BoundText = user_id
    Opt(2).value = True
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
        EleHeader.Enabled = True
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
Cmd(8).Enabled = False
        Case "E"
        Cmd(8).Enabled = False
        EleHeader.Enabled = True
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

            'Fg_Journal.Enabled = True
        Case "R"
        EleHeader.Enabled = False
            Me.EleHeader.Enabled = False
            Me.Fg_Journal.Editable = flexEDNone
            EleOpt.Enabled = False
            CboDes.CloseUp
            CboDes.Visible = False
        
            Cmd(0).Enabled = True
            Cmd(1).Enabled = True
            Cmd(2).Enabled = False
            Cmd(3).Enabled = False
            Cmd(8).Enabled = False
            Cmd(5).Enabled = True
            Cmd(7).Enabled = True
                        Cmd(8).Enabled = True
                        
            CmdRemove.Enabled = False
            ' Fg_Journal.Enabled = False
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
    Coloring
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

    'StrSQL = "SELECT  NOTES1.foxy_no,NOTES1.KALEB, NOTES1.DAWRY, NOTES1.NoteID,  NOTES1.NoteType," & _
     "NOTES1.NoteDate, NOTES1.Note_Value,NOTES1.NoteHijriDate," & _
     "NOTES1.Remark,NOTES1.general_cost_center, NOTES1.NotePosted,NOTES1.UserID,NoteSerial ,NoteSerial1," & _
     "DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_ID,DOUBLE_ENTREY_VOUCHERS1.USERID," & _
     "DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No,DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS1.Account_Code," & _
     "DOUBLE_ENTREY_VOUCHERS1.Value, DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit,DOUBLE_ENTREY_VOUCHERS1.Valuee,DOUBLE_ENTREY_VOUCHERS1.currency,DOUBLE_ENTREY_VOUCHERS1.rate," & _
     "DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description,DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name, DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id  " & _
     ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & _
     " FROM ACCOUNTS INNER JOIN (NOTES1 INNER JOIN DOUBLE_ENTREY_VOUCHERS1 " & _
     " ON NOTES1.NoteID = DOUBLE_ENTREY_VOUCHERS1.Notes_Id) ON " & _
     "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS1.Account_Code "
    
    
    'StrSQL = "SELECT     TOP 100 PERCENT dbo.Notes1.foxy_no, dbo.Notes1.KALEB, dbo.Notes1.DAWRY, dbo.Notes1.NoteID, dbo.Notes1.NoteType, dbo.Notes1.NoteDate, "
    'StrSQL = StrSQL & "   dbo.Notes1.Note_Value, dbo.Notes1.NoteHijriDate, dbo.Notes1.Remark, dbo.Notes1.general_cost_center, dbo.Notes1.NotePosted, dbo.Notes1.UserID,"
    'StrSQL = StrSQL & " dbo.Notes1.NoteSerial, dbo.Notes1.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_ID,"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.UserID AS Expr1, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No,"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS1.[Value],"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS1.valuee, dbo.DOUBLE_ENTREY_VOUCHERS1.currency,"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.rate, dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description,"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione, dbo.ACCOUNTS.Account_Name,"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial, dbo.Notes1.branch_no,"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id,branch_name,branch_namee"
    'StrSQL = StrSQL & "  FROM         dbo.ACCOUNTS INNER JOIN"
    'StrSQL = StrSQL & " dbo.Notes1 INNER JOIN"
    'StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS1 ON dbo.Notes1.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID ON"
    'StrSQL = StrSQL & " dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code LEFT OUTER JOIN"
    'StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id"
StrSQL = "SELECT     TOP 100 PERCENT dbo.Notes1.foxy_no, dbo.Notes1.KALEB, dbo.Notes1.DAWRY, dbo.Notes1.NoteID, dbo.Notes1.NoteType, dbo.Notes1.NoteDate, "
 StrSQL = StrSQL + "  dbo.Notes1.Note_Value, dbo.Notes1.NoteHijriDate, dbo.Notes1.Remark, dbo.Notes1.general_cost_center, dbo.Notes1.NotePosted, dbo.Notes1.UserID,"
 StrSQL = StrSQL + " dbo.Notes1.NoteSerial, dbo.Notes1.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS1.UserID,"
 StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code,"
 StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS1.[Value], dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS1.valuee,"
 StrSQL = StrSQL + "  dbo.DOUBLE_ENTREY_VOUCHERS1.currency,DOUBLE_ENTREY_VOUCHERS1.DueDate, dbo.DOUBLE_ENTREY_VOUCHERS1.rate, dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description,"
   StrSQL = StrSQL + "   dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione, dbo.ACCOUNTS.Account_Name,"
   StrSQL = StrSQL + "   dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial, dbo.Notes1.branch_no,"
   StrSQL = StrSQL + "   dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
   StrSQL = StrSQL + "  dbo.DOUBLE_ENTREY_VOUCHERS1.project_id , dbo.Projects.Project_name , dbo.projects.Project_nameE ,dbo.Notes1.LockedInterval ,dbo.Projects.Fullcode as ProjectCode"
   StrSQL = StrSQL + " FROM         dbo.ACCOUNTS INNER JOIN"
   StrSQL = StrSQL + "   dbo.Notes1 INNER JOIN"
   StrSQL = StrSQL + "    dbo.DOUBLE_ENTREY_VOUCHERS1 ON dbo.Notes1.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID ON"
   StrSQL = StrSQL + "    dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code LEFT OUTER JOIN"
   StrSQL = StrSQL + "   dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS1.project_id = dbo.projects.id LEFT OUTER JOIN"
   StrSQL = StrSQL + "   dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id"
   StrSQL = StrSQL + " Where NOTES1.NoteID=" & LngNoteID & ""
   StrSQL = StrSQL + GetAccountCodeHiding

    If LngNoteID = 1 Then
        StrSQL = StrSQL + " Order By  Credit_Or_Debit , value"
        'strsql = strsql + " Order By DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No "
    Else

        StrSQL = StrSQL + " Order By DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No "

    End If

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

    If Me.TxtNoteID.Text = 1 Then
        Me.Label9.Visible = True
    Else
        Me.Label9.Visible = False
    End If

    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)

    If rs("Notetype").value = 101 Then
        Text2.Text = "ÌœÊÌ"

    Else
        Text2.Text = "«·Ì"

    End If
If Not (IsNull(rs("LockedInterval").value)) Then
If rs("LockedInterval").value = True Then
Cmd(1).Enabled = False
Cmd(8).Enabled = False
Else
Cmd(1).Enabled = True
Cmd(8).Enabled = True
End If
Else
Cmd(1).Enabled = True
Cmd(8).Enabled = True
End If


    Text3.Text = get_note_type_name(rs("Notetype").value)

    Me.TxtDEVID.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)
    Me.TxtDEV_NO.Text = ""
    Me.TxtValue.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    Me.TxtDEV_NO.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)

    Me.DTP_Date.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Me.txtDueDate.value = IIf(IsNull(rs("DueDate").value), Date, rs("DueDate").value)

    Me.TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), Date, rs("NoteSerial").value)
    Me.TxtSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), Date, rs("NoteSerial1").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

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
            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("branch_id").value), "", rs("branch_id").value)
            .TextMatrix(i, .ColIndex("ProjectCode")) = IIf(IsNull(rs("ProjectCode").value), "", rs("ProjectCode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            
            End If
    
            .TextMatrix(i, .ColIndex("opening_balance_voucher_id")) = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
    
            .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(rs("DEV_ID_Line_No").value), "", rs("DEV_ID_Line_No").value)
            
            .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(rs("DEV_ID_Line_No1").value), "", rs("DEV_ID_Line_No1").value)
             .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(rs("DueDate").value), "", rs("DueDate").value)
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
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", Round(rs("Value").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("DebitValuee")) = IIf(IsNull(rs("Valuee").value), "", Round(rs("Valuee").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = "0"
            
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", Round(rs("Value").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = IIf(IsNull(rs("Valuee").value), "", Round(rs("Valuee").value, SystemOptions.SysDefCurrencyForamt))
                .TextMatrix(i, .ColIndex("DebitValuee")) = "0"
                
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If
              
            .TextMatrix(i, .ColIndex("userid")) = IIf(IsNull(rs("userid").value), "", rs("userid").value)
            
            .TextMatrix(i, .ColIndex("currenct_code")) = IIf(IsNull(rs("currency").value), "", rs("currency").value)
            
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(rs("rate").value), "", rs("rate").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
             
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
            
            
            .TextMatrix(i, .ColIndex("projectid")) = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
           If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs("Project_namee").value), "", rs("Project_namee").value)
                 
            Else
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
            End If
            
            rs.MoveNext
        Next i
        
        
        Dim s As String
        
        s = " SELECT SUM(DOUBLE_ENTREY_VOUCHERS1.[Value]) as value"
        s = s & " From dbo.Notes1"
        s = s & "        INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS1"
        s = s & "                         ON  dbo.Notes1.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID"
        s = s + " Where NOTES1.NoteID=" & LngNoteID & ""
        s = s & "                    AND DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit = 1"
        Dim rsDummy As New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If Not rsDummy.EOF Then
            Me.TxtTotalCredit.Text = rsDummy!value & ""
        Else
            Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        End If
    
        s = " SELECT SUM(DOUBLE_ENTREY_VOUCHERS1.[Value]) as value"
        s = s & " From dbo.Notes1"
        s = s & "        INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS1"
        s = s & "                         ON  dbo.Notes1.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID"
        s = s + " Where NOTES1.NoteID=" & LngNoteID & ""
        s = s & "                    AND DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit = 0"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If Not rsDummy.EOF Then
            Me.TxtTotalDebit.Text = rsDummy!value & ""
        Else
            Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        End If
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
        
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
    
        '  Me.TxtTotalCredit.text =Round(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
        '  Me.TxtTotalDebit.text =Round(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
        Coloring
fillapprovData
        If val(Me.TxtNoteID.Text) = 1 Then
            ReLineGrid
        End If

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
    
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
    
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
    End With

End Sub

Public Sub retrive1(LngNoteID As Long)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    StrSQL = "SELECT  NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & "NOTES.Remark, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID," & "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, DOUBLE_ENTREY_VOUCHERS.Account_Code," & "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit," & "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name  " & ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

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
     
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
    
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
    
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
        Msg = "”Ê› Ì „ Õ–› Â–« «·”‰œ —ﬁ„ " & Trim(Me.TxtSerial1.Text) & CHR(13)
        Msg = Msg & "›Â· √‰  „ √ﬂœ „‰ «·√” „—«— ...!!"
        IntRes = MsgBox(Msg, vbQuestion + vbOKCancel + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
    Else
        Msg = "This voucher " & Trim(Me.TxtSerial1.Text) & CHR(13)
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
    Me.Caption = "Opening Balance"
    Me.EleTop.Caption = Me.Caption
    Command4.Caption = "Add Des"
    Command3.Caption = "Call Des"
    Frame3.Caption = "Enter Voucher No. To copy it"
    Label7.Caption = "Voucher #"
    Command5.Caption = "Copy"
    Label8.Caption = "General C.C."
    Label17.Caption = "Right Click On Acc. to Show Statement"
    
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
    Label9.Caption = "Auto Voucher"
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
    lbl(3).Caption = "Code"
    Label11.Caption = "General Branch"
    lbl(4).Caption = "Value"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Modify"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Insert"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    Cmd(8).Caption = "Delete"

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
    lbl(13).Caption = "Result"
    lbl(8).Caption = "By"
    lbl(5).Caption = "Signature"
    ALLButton1.Caption = "Cost Center"
    ALLButton20.Caption = "Approved"
    ALLButton3.Caption = "Repeat Voucher"
    ALLButton6.Caption = "periodic"
    ALLButton7.Caption = "Template"
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
    OptSort(0).Caption = "Alphabetically"
    OptSort(1).Caption = "Charts sequence"

    With Fg_Journal
        .Cell(flexcpText, 0, .ColIndex("LineNo"), 1, .ColIndex("LineNo")) = "Line NO."
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = "Current Currency Value"
        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "Credit"
    
        .Cell(flexcpText, 0, .ColIndex("DebitValueE"), 0, .ColIndex("CreditValueE")) = "Forign Currency Value"
        .Cell(flexcpText, 1, .ColIndex("DebitValueE"), 1, .ColIndex("DebitValueE")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValueE"), 1, .ColIndex("CreditValueE")) = "Credit"
    
        '  .Cell(flexcpText, 0, .ColIndex("DebitValuee"), 0, .ColIndex("CreditValueE")) = "ValueE"
        '   .Cell(flexcpText, 1, .ColIndex("DebitValuee"), 1, .ColIndex("DebitValueE")) = "Debit"
        '   .Cell(flexcpText, 1, .ColIndex("CreditValuee"), 1, .ColIndex("CreditValueE")) = "Credit"
    
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Serial"
        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "Account Name"
        .Cell(flexcpText, 0, .ColIndex("Des"), 1, .ColIndex("Des")) = "Comment"
    
        .Cell(flexcpText, 0, .ColIndex("currenct_code"), 1, .ColIndex("currenct_code")) = "Currency"
     
        .Cell(flexcpText, 0, .ColIndex("rate"), 1, .ColIndex("rate")) = "Rate"
        .Cell(flexcpText, 0, .ColIndex("BranchName"), 1, .ColIndex("BranchName")) = "BranchName"
        .Cell(flexcpText, 0, .ColIndex("CC"), 1, .ColIndex("CC")) = "CC"
        .Cell(flexcpText, 0, .ColIndex("project"), 1, .ColIndex("project")) = "Project"
        .Cell(flexcpText, 0, .ColIndex("ProjectCode"), 1, .ColIndex("ProjectCode")) = "Project Code"
       
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
3    Dim TransBegine As Boolean
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
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
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
             .Col = .ColIndex("Account_Serial")
                             .Row = i
                             .ShowCell i, .ColIndex("Account_Serial")
                             
                             .SetFocus
                             
                    Exit Sub
                End If
            End If

        Next i

    End With

    If val(Me.TXTResults.Text) <> 0 Then

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
    If CheckSusAccounts1() = False Then
Exit Sub
End If

    If Me.TxtModFlg.Text = "N" Then
        ' Me.TxtSerial.text = ModAccounts.GetNewDEV_Serial(Me.DTP_Date.value)
    End If

    IntNoteType = 20

    Cn.BeginTrans
    TransBegine = True

    If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete   Notes1 Where Notes1.NoteID='" & Trim(TxtNoteID.Text) & "'"
        Cn.Execute StrSQL, , adExecuteNoRecords
     
        If DcCostCenter.BoundText <> "" Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
     
        If DcCostCenter.BoundText <> "" Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
    
    ElseIf Me.TxtModFlg.Text = "N" Then
        '---------------------------Get The Note ID ------------
        Me.TxtNoteID.Text = CStr(new_id("notes1", "NoteID", ""))
        Me.TxtDEVID.Text = CStr(new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", ""))
        Me.TxtDEV_NO.Text = Me.TxtDEVID.Text
        '---------------------------Begine of Saving------------
    End If

    Set RsNetes = New ADODB.Recordset
   ' RsNetes.Open "NOTES1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT  * from dbo.Notes1 Where (1 = -1)"
   RsNetes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
  
    RsNetes.AddNew
    RsNetes("branch_no").value = val(Me.dcBranch.BoundText)
    RsNetes("NoteID").value = val(Me.TxtNoteID.Text)
    RsNetes("NoteType").value = 101
    RsNetes("NoteSerial").value = val(Me.TxtSerial.Text)
    RsNetes("NoteSerial1").value = val(Me.TxtSerial1.Text)
    
    RsNetes("numbering_type").value = sand_numbering_type(0) ' „”·”· «·ﬁÌœ
    RsNetes("numbering_type1").value = sand_numbering_type(3) ' „”·”· «·”‰œ
    
    RsNetes("sanad_year").value = year(DTP_Date.value)
    RsNetes("sanad_month").value = Month(DTP_Date.value)
    RsNetes("foxy_no").value = val(Text1.Text)
    RsNetes("NoteDate").value = Me.DTP_Date.value

    RsNetes("DueDate").value = Me.txtDueDate.value

    RsNetes("Note_Value").value = val(Me.TxtValue.Text)
    RsNetes("Double_Entry_Vouchers_ID").value = val(Me.TxtDEVID.Text)
    RsNetes("DAWRY").value = Check4.value
    RsNetes("KALEB").value = Check3.value
    
    RsNetes("Remark").value = Trim$(Me.Txt.Text)
    RsNetes("UserID").value = val(Me.DcboUsers.BoundText)
    Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.TxtTotalDebit.Text, "0.00"), 0, True, ".")
    RsNetes("note_value_by_characters").value = Trim$(Me.Lb_note_value_by_characters.Caption)
    RsNetes("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    
    RsNetes.update
    Dim valuee As Variant
    Dim opening_balance_voucher_id As Double

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Variant

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) > 0 Then
                    IntDEV_Type = 0
                    SngDEV_Value = val(.TextMatrix(i, .ColIndex("DebitValue")))
                Else
                    IntDEV_Type = 1
                    SngDEV_Value = val(.TextMatrix(i, .ColIndex("CreditValue")))
                End If
            
                If val(.TextMatrix(i, .ColIndex("DebitValuee"))) > 0 Then
               
                    valuee = val(.TextMatrix(i, .ColIndex("DebitValuee")))
                Else
                 
                    valuee = val(.TextMatrix(i, .ColIndex("CreditValuee")))
                End If
            
                If val(.TextMatrix(i, .ColIndex("BranchId"))) = 0 Then
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(val(Me.dcBranch.BoundText) = 0, 1, val(Me.dcBranch.BoundText))
                End If

                opening_balance_voucher_id = val(.TextMatrix(i, .ColIndex("opening_balance_voucher_id")))

                If opening_balance_voucher_id = 0 Then opening_balance_voucher_id = -1
                If ModAccounts.AddNewDev(val(Me.TxtDEVID.Text), .TextMatrix(i, .ColIndex("LineNo")), .TextMatrix(i, .ColIndex("AccountCode")), SngDEV_Value, IntDEV_Type, .TextMatrix(i, .ColIndex("des")), val(Me.TxtNoteID.Text), , , SystemOptions.SysCurrentAccountIntervalID, Me.DTP_Date.value, val(.TextMatrix(i, .ColIndex("userid"))), , Me.TxtSerial.Text, , valuee, .TextMatrix(i, .ColIndex("currenct_code")), val(.TextMatrix(i, .ColIndex("rate"))), , .TextMatrix(i, .ColIndex("dese")), IIf(.TextMatrix(i, .ColIndex("LineNo1")) <> "", .TextMatrix(i, .ColIndex("LineNo1")), setfoxy_Line), , val(.TextMatrix(i, .ColIndex("projectid"))), , True, opening_balance_voucher_id, , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , , (.TextMatrix(i, .ColIndex("DueDate")))) = False Then
                    GoTo ErrTrap
                End If
            End If

        Next i

    End With

    Cn.CommitTrans
    TransBegine = False

    ' ÕœÌÀ «·—’Ìœ «·«›   «ÕÌ
    With Fg_Journal

        For i = .FixedRows To .Rows - 1
      
            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                '    update_account_opening_balance .TextMatrix(I, .ColIndex("AccountCode"))
 
            End If

        Next i

    End With

    CuurentLogdata

    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Saved"
    Else
        Msg = " „  ⁄„·Ì… «·Õ›Ÿ"
    End If

    'Õ›Ÿ „—ﬂ“ «· ﬂ·›… «·⁄«„
    '        If Me.DcCostCenter.BoundText <> "" Then
    save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "”‰œ ﬁÌœ «›  «ÕÌ", Me.DTP_Date.value
    '        End If
    save_cost_center

    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Me.TxtModFlg.Text = "R"
    fillapprovData
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
        rs("NoteDate").value = DTP_Date.value
        rs("NoteSerial").value = TxtSerial.Text
        rs("Remark").value = "”‰œ ﬁÌœ «›  «ÕÌ »—ﬁ„ " & TxtSerial1.Text & "    " & Me.TxtDes
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

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
 
   ' rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
 
    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("general_des").value = 1
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
                rs("NoteDate").value = DTP_Date.value
                rs("NoteSerial").value = TxtSerial.Text
                rs("Remark").value = Txt.Text
                rs.update
        
            End If

        Next i

    End With

    rs.Close
End Function

Private Sub TXTResults_Change()
    Me.TXTResults.Text = Round(val(Me.TXTResults.Text), 2)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
  
    Static StrOldTransID As String
    Dim StrSQL As String
On Error Resume Next
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        SetForNew
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (2)
    End If

    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     " From notes where (((notes.NoteType) =200)) " & _
     " ORDER BY NOTES.NoteID "
    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     "From notes where (((notes.NoteType)=200)) " & _
     "    ORDER BY NOTES.NoteID "
    

If Index = 2 Then GoTo ll
    If Rs1.BOF Or Rs1.EOF Then
        Exit Sub
    End If

    If StrOldTransID <> "" Then
        Rs1.Find "NoteID=" & StrOldTransID & "", , adSearchForward, 1

        If Rs1.BOF Or Rs1.EOF Then
            Rs1.MoveFirst
        End If

    Else
        Rs1.MoveFirst
    End If
ll:
    Select Case Index

        Case 1 'First

            If Not (Rs1.BOF Or Rs1.EOF) Then
                Rs1.MoveFirst
            End If

        Case 0 'Previous

            If Not (Rs1.BOF Or Rs1.EOF) Then
                Rs1.MovePrevious

                If Rs1.BOF Then Rs1.MoveNext
            End If

        Case 3 'NEXT

            If Not (Rs1.BOF Or Rs1.EOF) Then
                Rs1.MoveNext

                If Rs1.EOF Then Rs1.MovePrevious
            End If

        Case 2 'Last
        Rs1.Close
        
    StrSQL = "SELECT NOTES1.NoteID, NOTES1.NoteType " & "From notes1 where   notes1.NoteType =101      ORDER BY NOTES1.NoteID  "
    
'    If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = "SELECT  NOTES1.NoteID, NOTES1.NoteType   From notes1    where branch_no=" & Current_branch & " and  notetype =101   ORDER BY NOTES1.NoteID "
'     StrSQL = "SELECT  NOTES1.NoteID, NOTES1.NoteType   From notes1    where branch_no in(" & Current_branchSql & ") and  notetype =101   ORDER BY NOTES1.NoteID "
     


'    End If
    
  StrSQL = "SELECT  NOTES1.NoteID, NOTES1.NoteType   From notes1    where  branch_no=0 or  branch_no in(" & Current_branchSql & ") and  notetype =101   ORDER BY NOTES1.NoteID "
  
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
            If Not (Rs1.BOF Or Rs1.EOF) Then
                Rs1.MoveLast
                Me.TxtModFlg.Text = ""
                Me.TxtModFlg.Text = "R"
            End If

    End Select

    If Not (IsNull(Rs1("NoteID").value)) Then
        Me.Retrive Rs1("NoteID").value
        StrOldTransID = Rs1("NoteID").value
    
    End If
'Print Rs1.RecordCount
        Me.TxtModFlg.Text = ""
        Me.TxtModFlg.Text = "R"
        
   ' rs1.Close
   ' Set rs = Nothing
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


