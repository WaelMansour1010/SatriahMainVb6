VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmYearDurations 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "FrmYearDurations.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   10065
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9180
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10065
      _cx             =   17754
      _cy             =   16193
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1455
         Left            =   60
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   9945
         _cx             =   17542
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
         Begin VB.ComboBox cbDiff 
            Height          =   288
            Left            =   1428
            TabIndex        =   4
            Top             =   600
            Width           =   2844
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   5925
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            Top             =   240
            Width           =   2892
         End
         Begin VB.ComboBox cbType 
            Height          =   288
            Left            =   1428
            TabIndex        =   2
            Top             =   240
            Width           =   2844
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   5925
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   2892
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   324
            Left            =   7344
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   936
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98959363
            CurrentDate     =   37140
            MinDate         =   -182619
         End
         Begin Dynamic_Byte.NourHijriCal FromdateH 
            Height          =   330
            Left            =   5925
            TabIndex        =   6
            Top             =   930
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   476
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   336
            Left            =   2892
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   936
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98959363
            CurrentDate     =   37140
            MinDate         =   -182619
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   336
            Left            =   1428
            TabIndex        =   8
            Top             =   936
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   582
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   570
            Index           =   15
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1005
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
            ButtonImage     =   "FrmYearDurations.frx":038A
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›—ﬁ «·«Ì«„ «·ÂÃ—Ì…"
            Height          =   285
            Index           =   2
            Left            =   4380
            TabIndex        =   30
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   285
            Index           =   1
            Left            =   8625
            TabIndex        =   29
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì»œ√ „‰ "
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   8745
            TabIndex        =   28
            Top             =   930
            Width           =   1050
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì‰ ÂÏ ›Ï"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   4590
            TabIndex        =   27
            Top             =   930
            Width           =   1050
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ﬁÊÌ„"
            Height          =   285
            Index           =   0
            Left            =   4380
            TabIndex        =   26
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄«„ «·œ—«”Ï"
            Height          =   285
            Index           =   16
            Left            =   8625
            TabIndex        =   25
            Top             =   600
            Width           =   1200
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   720
         Left            =   -45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   -45
         Width           =   10290
         _cx             =   18150
         _cy             =   1270
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "    ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…   "
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
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   20
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
            ButtonImage     =   "FrmYearDurations.frx":6BEC
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
            TabIndex        =   21
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
            ButtonImage     =   "FrmYearDurations.frx":6F86
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
            TabIndex        =   22
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
            ButtonImage     =   "FrmYearDurations.frx":7320
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
            TabIndex        =   23
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
            ButtonImage     =   "FrmYearDurations.frx":76BA
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   855
         Left            =   75
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   8265
         Width           =   9930
         _cx             =   17515
         _cy             =   1508
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   0
            Left            =   8760
            TabIndex        =   32
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":7A54
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
            Height          =   540
            Index           =   1
            Left            =   7740
            TabIndex        =   33
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":E2B6
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
            Height          =   540
            Index           =   2
            Left            =   6696
            TabIndex        =   34
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":14B18
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
            Height          =   540
            Index           =   3
            Left            =   5628
            TabIndex        =   35
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":1B37A
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
            Height          =   540
            Index           =   4
            Left            =   4524
            TabIndex        =   36
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":21BDC
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
            Height          =   540
            Index           =   6
            Left            =   1308
            TabIndex        =   37
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":2843E
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
            Height          =   540
            Left            =   240
            TabIndex        =   38
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":52060
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
            Height          =   540
            Index           =   7
            Left            =   3492
            TabIndex        =   39
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":588C2
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
            Height          =   540
            Index           =   9
            Left            =   2400
            TabIndex        =   67
            Top             =   156
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   953
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
            ButtonImage     =   "FrmYearDurations.frx":5F124
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   75
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   7545
         Width           =   5745
         _cx             =   10134
         _cy             =   1085
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   396
            Index           =   4
            Left            =   816
            TabIndex        =   44
            Top             =   144
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   390
            Index           =   3
            Left            =   3810
            TabIndex        =   43
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   396
            Left            =   120
            TabIndex        =   42
            Top             =   144
            Width           =   660
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   390
            Left            =   2925
            TabIndex        =   41
            Top             =   150
            Width           =   825
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5160
         Left            =   0
         TabIndex        =   45
         Top             =   2280
         Width           =   9975
         _cx             =   17595
         _cy             =   9102
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
         FrontTabColor   =   -2147483633
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·› —« | ”ÃÌ· «Ì«„ «·⁄ÿ·«  «·—”„Ì…"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   4785
            Left            =   45
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   45
            Width           =   9885
            _cx             =   17436
            _cy             =   8440
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
            Begin VB.TextBox txtTotDayVacation 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   4410
               Width           =   960
            End
            Begin VB.TextBox txttotDaywork 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   71
               Top             =   4410
               Width           =   960
            End
            Begin VB.TextBox txttot 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   4440
               Locked          =   -1  'True
               TabIndex        =   69
               Top             =   4410
               Width           =   975
            End
            Begin VSFlex8UCtl.VSFlexGrid fg 
               Height          =   3300
               Left            =   0
               TabIndex        =   17
               Top             =   1065
               Width           =   9975
               _cx             =   17595
               _cy             =   5821
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   10.5
                  Charset         =   178
                  Weight          =   700
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
               BackColorAlternate=   16776960
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
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmYearDurations.frx":65986
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
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   1080
               Left            =   0
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   0
               Width           =   9915
               _cx             =   17489
               _cy             =   1905
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
               Caption         =   "«·⁄ÿ·… «·«”»Ê⁄Ì…"
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
               Begin VB.CheckBox opt_Fr 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ã„⁄…"
                  Height          =   675
                  Left            =   1080
                  TabIndex        =   16
                  Top             =   285
                  Width           =   840
               End
               Begin VB.CheckBox opt_sa 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”» "
                  Height          =   675
                  Left            =   8805
                  TabIndex        =   10
                  Top             =   285
                  Width           =   810
               End
               Begin VB.CheckBox opt_su 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Õœ"
                  Height          =   675
                  Left            =   7875
                  TabIndex        =   11
                  Top             =   285
                  Width           =   630
               End
               Begin VB.CheckBox opt_mo 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«À‰Ì‰"
                  Height          =   675
                  Left            =   6435
                  TabIndex        =   12
                  Top             =   285
                  Width           =   840
               End
               Begin VB.CheckBox opt_tu 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·À·«À«¡"
                  Height          =   675
                  Left            =   5115
                  TabIndex        =   13
                  Top             =   285
                  Width           =   810
               End
               Begin VB.CheckBox opt_We 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«—»⁄«¡"
                  Height          =   675
                  Left            =   3750
                  TabIndex        =   14
                  Top             =   285
                  Width           =   840
               End
               Begin VB.CheckBox opt_Th 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Œ„Ì”"
                  Height          =   675
                  Left            =   2430
                  TabIndex        =   15
                  Top             =   285
                  Width           =   870
               End
            End
            Begin VB.Label Label12 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «·⁄ÿ·« "
               Height          =   255
               Left            =   1080
               TabIndex        =   72
               Top             =   4410
               Width           =   1800
            End
            Begin VB.Label Label11 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «Ì«„ «·⁄„·"
               Height          =   255
               Left            =   3240
               TabIndex        =   70
               Top             =   4410
               Width           =   1080
            End
            Begin VB.Label Label10 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «Ì«„ «·⁄«„ «·œ—«”Ï"
               Height          =   255
               Left            =   5625
               TabIndex        =   68
               Top             =   4410
               Width           =   1590
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   4785
            Left            =   10620
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   45
            Width           =   9885
            _cx             =   17436
            _cy             =   8440
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
            Caption         =   "⁄ÿ·«  «Œ—Ï"
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
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   732
               Left            =   1380
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Top             =   1260
               Width           =   7395
            End
            Begin MSDataListLib.DataCombo dcVacType 
               Height          =   315
               Left            =   1380
               TabIndex        =   50
               Top             =   480
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcDur 
               Height          =   315
               Left            =   5895
               TabIndex        =   51
               Top             =   480
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   9885
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   900
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   98959363
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal ToDateH2 
               Height          =   315
               Left            =   1380
               TabIndex        =   53
               Top             =   885
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker ToDate2 
               Height          =   315
               Left            =   2760
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   885
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   98959363
               CurrentDate     =   41640
            End
            Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
               Height          =   2625
               Left            =   0
               TabIndex        =   55
               Top             =   2115
               Width           =   9885
               _cx             =   17436
               _cy             =   4630
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   10.5
                  Charset         =   178
                  Weight          =   700
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
               BackColorAlternate=   16776960
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
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmYearDurations.frx":65B47
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
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   405
               Index           =   5
               Left            =   240
               TabIndex        =   56
               Top             =   480
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   714
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
               ButtonImage     =   "FrmYearDurations.frx":65D32
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
               Height          =   375
               Index           =   8
               Left            =   240
               TabIndex        =   57
               Top             =   960
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   661
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
               ButtonImage     =   "FrmYearDurations.frx":6C594
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   -2147483637
            End
            Begin Dynamic_Byte.NourHijriCal FromdateH2 
               Height          =   315
               Left            =   5895
               TabIndex        =   65
               Top             =   900
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker FromDate2 
               Height          =   315
               Left            =   7275
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   900
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   98959363
               CurrentDate     =   41640
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄«„ «·œ—«”Ï"
               Height          =   270
               Left            =   8895
               TabIndex        =   64
               Top             =   480
               Width           =   840
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   240
               Left            =   8775
               TabIndex        =   63
               Top             =   900
               Width           =   960
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘—Õ"
               Height          =   330
               Left            =   9135
               TabIndex        =   62
               Top             =   1305
               Width           =   600
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï  «—ÌŒ"
               Height          =   360
               Left            =   4920
               TabIndex        =   61
               Top             =   900
               Width           =   735
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   360
               Left            =   11235
               TabIndex        =   60
               Top             =   900
               Width           =   990
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄«„ «·œ—«”Ï"
               Height          =   405
               Left            =   11235
               TabIndex        =   59
               Top             =   510
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·⁄ÿ·…"
               Height          =   405
               Left            =   4920
               TabIndex        =   58
               Top             =   480
               Width           =   735
            End
         End
      End
   End
End
Attribute VB_Name = "FrmYearDurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp2 As ADODB.Recordset
Dim rs_Dur  As ADODB.Recordset
Dim rs_vac As ADODB.Recordset
Dim rs_hol As ADODB.Recordset
Dim rsHolidays  As ADODB.Recordset
Dim rsDuration As ADODB.Recordset

Dim TTP As clstooltip
Dim FromDate_ As Date
Dim ToDate_ As Date
Dim FromDateH_ As String
Dim ToDateH_ As String


Private Sub cbType_Click()

'Exit Sub

If cbType.ListIndex = 0 Then
    
    FromdateH.Enabled = False
    ToDateH.Enabled = False
    FromDate.Enabled = True
    ToDate.Enabled = True
    
    fg.ColWidth(fg.ColIndex("FromDateH")) = 0
    fg.ColWidth(fg.ColIndex("ToDateH")) = 0
    fg.ColWidth(fg.ColIndex("FromDate")) = 1152
    fg.ColWidth(fg.ColIndex("ToDate")) = 1152
ElseIf cbType.ListIndex = 1 Then
    FromDate.Enabled = False
    ToDate.Enabled = False
     FromdateH.Enabled = True
    ToDateH.Enabled = True
    
     fg.ColWidth(fg.ColIndex("FromDate")) = 0
     fg.ColWidth(fg.ColIndex("ToDate")) = 0
     fg.ColWidth(fg.ColIndex("FromDateH")) = 1152
     fg.ColWidth(fg.ColIndex("ToDateH")) = 1152
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
            TxtId.Text = CStr(new_id("tbldurations", "ID", "", True))
            TxtName.SetFocus
        Case 1
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "E"
        Case 2
            SaveData
        Case 3
            Undo
        Case 8
             
            Del_row
        Case 4
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_Company
                    
        Case 5
                save_Vac
         Case 6
            Unload Me
         Case 7
   '      print_report2
   
  Case 15
 Calc
           Case 9
            '    FrmSearch_Duration.SendForm = "Dur"
            '    FrmSearch_Duration.show
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Save_WeekVac(DurID As Integer)
Dim StrSQL As String

On Error GoTo errortrap
   StrSQL = " delete From tblholidays  where durationID =" & DurID
   Cn.Execute StrSQL, , adExecuteNoRecords

    Set rsHolidays = New ADODB.Recordset
    StrSQL = "SELECT  *  From tblholidays "
    rsHolidays.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rsHolidays.AddNew
    
    rsHolidays("id") = CStr(new_id("tblholidays", "id", "", True))
    rsHolidays("Sa") = opt_sa.value
    rsHolidays("Su") = opt_su.value
    rsHolidays("Mo") = opt_mo.value
    rsHolidays("Tu") = opt_tu.value
    rsHolidays("We") = opt_We.value
    rsHolidays("Th") = opt_Th.value
    rsHolidays("Fr") = opt_Fr.value
    rsHolidays("DurationID") = DurID
    rsHolidays.update
   ' MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
   
errortrap:
  'MsgBox ("ÕœÀ Œÿ¡ «À‰«¡ Õ›Ÿ «·»Ì«‰« ")
   
End Sub


Private Sub Retrive_Holidays(DurID As Integer)

    Set rsHolidays = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From tblholidays where durationID = " & DurID
    rsHolidays.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    If rsHolidays.RecordCount > 0 Then
         rsHolidays.MoveFirst
         
         If rsHolidays("Sa").value = True Then
         opt_sa.value = 1
         ElseIf rsHolidays("Sa").value = False Then
         opt_sa.value = 0
         End If
         
         If rsHolidays("Su").value = True Then
                opt_su.value = 1
         ElseIf rsHolidays("Su").value = False Then
                opt_su.value = 0
         End If
         
         If rsHolidays("mo").value = True Then
         opt_mo.value = 1
         ElseIf rsHolidays("mo").value = False Then
         opt_mo.value = 0
         End If
         
         If rsHolidays("Tu").value = True Then
         opt_tu.value = 1
         ElseIf rsHolidays("Tu").value = False Then
         opt_tu.value = 0
         End If
         
         If rsHolidays("We").value = True Then
         opt_We.value = 1
         ElseIf rsHolidays("We").value = False Then
         opt_We.value = 0
         End If
         
        If rsHolidays("Th").value = True Then
         opt_Th.value = 1
         ElseIf rsHolidays("Th").value = False Then
         opt_Th.value = 0
         End If
         
         If rsHolidays("Fr").value = True Then
         opt_Fr.value = 1
         ElseIf rsHolidays("Fr").value = False Then
         opt_Fr.value = 0
         End If
    End If

End Sub

Private Sub Calc()
Dim BeginFirstPer As String, EndFirstPer As String, str As String, EndLastPer As String, BeginLastPer As String

If cbType.ListIndex = -1 Then
MsgBox ("«Œ — ‰Ê⁄ «·› —… «Ê·«")
cbType.SetFocus
SendKeys ("{F4}")
Exit Sub
End If

Dim j As Integer
With fg
    For j = 1 To fg.Rows - 1
        .TextMatrix(j, .ColIndex("FromDate")) = ""
        .TextMatrix(j, .ColIndex("ToDate")) = ""
        .TextMatrix(j, .ColIndex("FromDateH")) = ""
        .TextMatrix(j, .ColIndex("ToDateH")) = ""
    Next
End With

If cbType.ListIndex = 0 Then

'///////////////////////////////////////////////////
BeginFirstPer = FromDate.value
EndFirstPer = dhLastDayInMonth(FromDate)
EndLastPer = ToDate.value
BeginLastPer = dhFirstDayInMonth(ToDate.value)

Dim count As Integer

VBA.Calendar = vbCalGreg
count = Date_MonthsBetweenDates(CDate(DateAdd("d", 1, EndFirstPer)), CDate(DateAdd("d", -1, BeginLastPer)))

With fg
.Rows = count + 4
If count = -2 Then
    .TextMatrix(1, .ColIndex("FromDate")) = FromDate.value
    .TextMatrix(1, .ColIndex("ToDate")) = ToDate.value
    .TextMatrix(1, .ColIndex("FromDateh")) = ToHijriDate(FromDate.value)
     .TextMatrix(1, .ColIndex("ToDateh")) = ToHijriDate(ToDate.value)
     .TextMatrix(1, .ColIndex("Month")) = Format(FromDate, "MMMM")
     Exit Sub
 ElseIf count < -3 Then
    Exit Sub
End If
End With


Dim i As Integer
i = 1
With fg
Dim n  As Integer

      .TextMatrix(i, .ColIndex("FromDate")) = Format(BeginFirstPer, "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("ToDate")) = Format(EndFirstPer, "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("FromDateh")) = Format(ToHijriDate(BeginFirstPer), "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("ToDateh")) = Format(ToHijriDate(EndFirstPer), "yyyy/MM/dd")
       
       .TextMatrix(i, .ColIndex("Month")) = Format(BeginFirstPer, "MMMM")
       
      .TextMatrix(i, .ColIndex("Serial")) = 1
    For i = 1 To count + 1
        .TextMatrix(i + 1, .ColIndex("FromDate")) = Format(DateAdd("d", 1, .TextMatrix(i, .ColIndex("ToDate"))), "yyyy/MM/dd")
        .TextMatrix(i + 1, .ColIndex("ToDate")) = Format(dhLastDayInMonth(DateAdd("d", 1, .TextMatrix(i, .ColIndex("ToDate")))), "yyyy/MM/dd")                                              'DateAdd("M", 1, .TextMatrix(i, .ColIndex("ToDate")))
        .TextMatrix(i + 1, .ColIndex("FromDateH")) = Format(ToHijriDate(DateAdd("d", 1, .TextMatrix(i, .ColIndex("ToDate")))), "yyyy/MM/dd")
        .TextMatrix(i + 1, .ColIndex("ToDateH")) = Format(ToHijriDate(DateAdd("M", 1, .TextMatrix(i, .ColIndex("ToDate")))), "yyyy/MM/dd")
        
        .TextMatrix(i + 1, .ColIndex("Month")) = Format(.TextMatrix(i + 1, .ColIndex("FromDate")), "MMMM")
        
        .TextMatrix(i + 1, .ColIndex("Serial")) = i + 1
                
     Next
      .TextMatrix(3 + count, .ColIndex("FromDate")) = Format(BeginLastPer, "yyyy/MM/dd")
      .TextMatrix(3 + count, .ColIndex("ToDate")) = Format(EndLastPer, "yyyy/MM/dd")
      .TextMatrix(3 + count, .ColIndex("FromDateH")) = Format(ToHijriDate(BeginLastPer), "yyyy/MM/dd")
      .TextMatrix(3 + count, .ColIndex("ToDateH")) = Format(ToHijriDate(EndLastPer), "yyyy/MM/dd")
      .TextMatrix(3 + count, .ColIndex("Serial")) = i + 1
      .TextMatrix(3 + count, .ColIndex("Month")) = Format(.TextMatrix(3 + count, .ColIndex("FromDate")), "MMMM")
      
End With
'//////////////////////////////

ElseIf cbType.ListIndex = 1 Then

'///////////////////////////////////////////////////
Dim EndFirstPerG As String
Dim FirstMonth As String

BeginFirstPer = FromdateH.value

VBA.Calendar = vbCalHijri
EndFirstPer = dhLastDayInMonth(FromdateH.value)
VBA.Calendar = vbCalGreg

EndLastPer = ToDateH.value
VBA.Calendar = vbCalHijri
BeginLastPer = dhFirstDayInMonth(ToDateH.value)
FirstMonth = dhFirstDayInMonth(FromdateH.value)
count = 0
'
Dim s1 As String, s2 As String, s3 As String, s4 As String

s1 = Format(EndFirstPer, "dd/MM/yyyy")
s3 = Format(BeginLastPer, "dd/MM/yyyy")

VBA.Calendar = vbCalHijri
s2 = DateAdd("d", 1, s1)
s4 = DateAdd("d", -1, s3)
VBA.Calendar = vbCalGreg

VBA.Calendar = vbCalHijri
'count = Date_MonthsBetweenDates(CDate(s1), CDate(s4))
'count = Date_MonthsBetweenDates(CDate("1437/01/30"), CDate("1437/11/01"))
'VBA.Calendar = vbCalGreg

count = Date_MonthsBetweenDates(CDate(s1), CDate(s4))
VBA.Calendar = vbCalGreg



With fg
  If count = -2 Then
    .TextMatrix(1, .ColIndex("FromDate")) = Format(ToGregorianDate(FromdateH.value), "yyyy/MM/dd")
    .TextMatrix(1, .ColIndex("ToDate")) = Format(ToGregorianDate(ToDateH.value), "yyyy/MM/dd")
    .TextMatrix(1, .ColIndex("FromDateh")) = Format(FromdateH.value, "yyyy/MM/dd")
    .TextMatrix(1, .ColIndex("ToDateh")) = Format(ToDateH.value, "yyyy/MM/dd")
     Exit Sub
 ElseIf count < -3 Then
    Exit Sub
End If
End With

i = 1
With fg
.Rows = count + 3
      .TextMatrix(i, .ColIndex("FromDateh")) = Format(BeginFirstPer, "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("ToDateh")) = Format(EndFirstPer, "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("FromDate")) = Format(ToGregorianDate(BeginFirstPer), "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("ToDate")) = Format(ToGregorianDate(EndFirstPer), "yyyy/MM/dd")
       
       VBA.Calendar = vbCalHijri
      .TextMatrix(1, .ColIndex("Month")) = Format(BeginFirstPer, "MMMM")
       VBA.Calendar = vbCalGreg
      
      .TextMatrix(i, .ColIndex("ToDate")) = Format(ToGregorianDate(EndFirstPer), "yyyy/MM/dd")
      .TextMatrix(i, .ColIndex("Serial")) = 1
    For i = 1 To count
    
    VBA.Calendar = vbCalHijri
     .TextMatrix(i + 1, .ColIndex("FromDateh")) = Format(DateAdd("M", 1 * i, FirstMonth), "yyyy/MM/dd")
     .TextMatrix(i + 1, .ColIndex("ToDateh")) = Format(DateAdd("d", -1, DateAdd("M", 1 * (i + 1), FirstMonth)), "yyyy/MM/dd")
     VBA.Calendar = vbCalGreg
    ' .TextMatrix(i + 1, .ColIndex("FromDate")) = Format(ToGregorianDate(.TextMatrix(i, .ColIndex("ToDateH"))), "yyyy/MM/dd")
    ' .TextMatrix(i + 1, .ColIndex("ToDate")) = Format(ToGregorianDate(.TextMatrix(i, .ColIndex("ToDateH"))), "yyyy/MM/dd")
   
  .TextMatrix(i + 1, .ColIndex("FromDate")) = Format(ToGregorianDate(.TextMatrix(i + 1, .ColIndex("FromDateh"))), "yyyy/MM/dd")
   .TextMatrix(i + 1, .ColIndex("ToDate")) = Format(ToGregorianDate(.TextMatrix(i + 1, .ColIndex("ToDateh"))), "yyyy/MM/dd")
       VBA.Calendar = vbCalHijri
      .TextMatrix(i + 1, .ColIndex("Month")) = Format(.TextMatrix(i + 1, .ColIndex("FromDateh")), "MMMM")
       VBA.Calendar = vbCalGreg
        
        .TextMatrix(i + 1, .ColIndex("Serial")) = i + 1
     Next
     
     
      .TextMatrix(2 + count, .ColIndex("FromDateH")) = Format(BeginLastPer, "yyyy/MM/dd")
      .TextMatrix(2 + count, .ColIndex("ToDateH")) = Format(EndLastPer, "yyyy/MM/dd")
      .TextMatrix(2 + count, .ColIndex("FromDate")) = Format(ToGregorianDate(BeginLastPer), "yyyy/MM/dd")
      .TextMatrix(2 + count, .ColIndex("ToDate")) = Format(ToGregorianDate(EndLastPer), "yyyy/MM/dd")
      
       VBA.Calendar = vbCalHijri
      .TextMatrix(2 + count, .ColIndex("Month")) = Format(BeginLastPer, "MMMM")
       VBA.Calendar = vbCalGreg
       .TextMatrix(2 + count, .ColIndex("Serial")) = i + 1
End With
'//////////////////////////////

End If
End Sub

Private Sub Add_Schedule(FromDate As Date, ToDate As Date, dur As Integer, DDID As Integer)
   Dim str As String, str1 As String, day As String
   
   Do While FromDate <= ToDate
        str = Weekday(FromDate, vbSaturday)
        day = WeekdayName(str, False, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        If IsHoliday(str, dur) Then
                 AddRowToSchedule dur, FromDate, ToHijriDate(FromDate), True, day, DDID
        Else
                 AddRowToSchedule dur, FromDate, ToHijriDate(FromDate), False, day, DDID
        End If
        
       VBA.Calendar = vbCalGreg
       FromDate = DateAdd("d", 1, FromDate)
   Loop
End Sub


Private Sub Add_ScheduleH(FromDate As String, ToDate As String, dur As Integer, DDID As Integer)
  
  Dim str As String, str1 As String, day As String
       VBA.Calendar = vbCalHijri
  FromDate = Format(FromDate, "yyyy/MM/dd")
  ToDate = Format(ToDate, "yyyy/MM/dd")
   
   Do While DateDiff("d", FromDate, ToDate) >= 0
        VBA.Calendar = vbCalHijri
        str = Weekday(FromDate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        day = WeekdayName(Weekday(FromDate, vbSaturday), False, vbSaturday)
        VBA.Calendar = vbCalGreg
        
        If IsHoliday(str, dur) Then
                 AddRowToSchedule dur, ToGregorianDate(FromDate), FromDate, True, day, DDID
        Else
                 AddRowToSchedule dur, ToGregorianDate(FromDate), FromDate, False, day, DDID
        End If
         
       VBA.Calendar = vbCalHijri
         FromDate = DateAdd("d", 1, FromDate)
        VBA.Calendar = vbCalGreg
             
        FromDate = Format(FromDate, "yyyy/MM/dd")
   
'**********************************
      VBA.Calendar = vbCalHijri
  FromDate = Format(FromDate, "yyyy/MM/dd")
  ToDate = Format(ToDate, "yyyy/MM/dd")
  '**********************************
   
   Loop
   VBA.Calendar = vbCalGreg
End Sub


Private Function IsHoliday(day As String, dur As Integer) As Boolean
    Dim str As String
    str = " select * from  tblholidays  where DurationID =" & dur
    Set rs_hol = New ADODB.Recordset
    rs_hol.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs_hol.RecordCount > 0 Then
         '   If rs_hol("sa").value = True And day = "Sat" Then
         '       IsHoliday = True
         ''   ElseIf rs_hol("su").value = True And day = "Sun" Then
         '       IsHoliday = True
         ''   ElseIf rs_hol("Mo").value = True And day = "Mon" Then
         '       IsHoliday = True
         '''   ElseIf rs_hol("Tu").value = True And day = "Tue" Then
          '       IsHoliday = True
          '  ElseIf rs_hol("We").value = True And day = "Wed" Then
          '      IsHoliday = True
          '  ElseIf rs_hol("Th").value = True And day = "Thu" Then
          '      IsHoliday = True
          '  ElseIf rs_hol("Fr").value = True And day = "Fri" Then
          '      IsHoliday = True
          '  End If
          
            If rs_hol("sa").value = True And day = "1" Then
                IsHoliday = True
            ElseIf rs_hol("su").value = True And day = "2" Then
                IsHoliday = True
            ElseIf rs_hol("Mo").value = True And day = "3" Then
                IsHoliday = True
            ElseIf rs_hol("Tu").value = True And day = "4" Then
                 IsHoliday = True
            ElseIf rs_hol("We").value = True And day = "5" Then
                IsHoliday = True
            ElseIf rs_hol("Th").value = True And day = "6" Then
                IsHoliday = True
            ElseIf rs_hol("Fr").value = True And day = "7" Then
                IsHoliday = True
            End If
    End If

End Function

'Private Function ISOfficialVacation(dt As Date, dur As Integer) As Boolean
'
'    Dim str As String
'    str = " select * from  TblVacationDays  where DurationID =   " & dur & "  and   " & dt & " Between FromDate And ToDate  "
'    Set rs_vac = New ADODB.Recordset
'    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If rs_vac.RecordCount > 0 Then
'        ISOfficialVacation = True
'    End If
'
''End Function


Private Sub AddRowToSchedule(dur As Integer, dt As Date, dth As String, isvac As Boolean, day As String, DDID As Integer)
        
       Set rs_Dur = New ADODB.Recordset
       rs_Dur.Open " TblVacationschedule ", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
       rs_Dur.AddNew
       rs_Dur("ID") = CStr(new_id("TblVacationschedule", "ID", "", True))
       rs_Dur("DurationID") = dur
       rs_Dur("Date") = dt
       rs_Dur("DateH") = Format(dth, "yyyy/MM/dd")
       rs_Dur("isvac") = isvac
       rs_Dur("day") = day
       rs_Dur("DDID") = DDID
       If isvac = True Then
            rs_Dur("color") = "255"
       End If
       rs_Dur.update

End Sub


Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    
 dhLastDayInMonth = MonthLastDay(CDate(dtmDate))
    Exit Function
    dhLastDayInMonth = DateSerial(year(dtmDate), _
     Month(dtmDate) + 1, 0)
End Function

Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhFirstDayInMonth = DateSerial(year(dtmDate), _
     Month(dtmDate), 1)
End Function


Private Function Date_MonthsBetweenDates(dt As Date, dT2 As Date) As Integer
      Dim dBeginDate As Date
      Dim dEndDate As Date
      Dim intMonths As Integer
      
      ' Beginning date.
      dBeginDate = dt
      ' Ending Date.
      dEndDate = dT2
      ' Calculate number of months between dates.
      intMonths = ((year(dEndDate) - year(dBeginDate)) * 12) + Month(dEndDate) - Month(dBeginDate)
      ' Display number of months.
      
'       MsgBox str$(intMonths) & " month(s)"
     
      Date_MonthsBetweenDates = intMonths
      
End Function

Private Sub Calculations(Optional WithMsg As Boolean = True)
'    On Error GoTo ErrTrap
 
End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub txtboxValue_KeyPress(KeyAscii As Integer)
   'KeyAscii = KeyAscii_Num(KeyAscii, Me.txtboxValue.text, 0)
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
End Sub

 

Private Sub DcDur_Change()


Dim str  As String
Set rsDuration = New ADODB.Recordset
str = " select * from tbldurations where id =  " & val(dcDur.BoundText)
rsDuration.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

Retrive_Vacation (val(dcDur.BoundText))

If rsDuration.RecordCount > 0 Then
    rsDuration.MoveFirst
    FromDate_ = rsDuration("FromDate").value
    ToDate_ = rsDuration("ToDate").value
    FromDateH_ = rsDuration("FromDateH").value
    ToDateH_ = rsDuration("ToDateH").value


    FromDate2.value = FromDate_
    FromdateH2.value = FromDateH_
    ToDate2.value = FromDate_
    toDateH2.value = FromDateH_

    If rsDuration("Type").value = 0 Then
        FromDate2.Enabled = True
        ToDate2.Enabled = True
        FromdateH2.Enabled = False
        toDateH2.Enabled = False
        
        With FgInstallments
            .ColWidth(.ColIndex("FromDate")) = 1200
            .ColWidth(.ColIndex("ToDate")) = 1200
            .ColWidth(.ColIndex("FromDateH")) = 0
            .ColWidth(.ColIndex("ToDateH")) = 0
        End With
        
    ElseIf rsDuration("Type").value = 1 Then
        
        FromDate2.Enabled = False
        ToDate2.Enabled = False
        FromdateH2.Enabled = True
        toDateH2.Enabled = True
         
         With FgInstallments
            .ColWidth(.ColIndex("FromDate")) = 0
            .ColWidth(.ColIndex("ToDate")) = 0
            .ColWidth(.ColIndex("FromDateH")) = 1200
            .ColWidth(.ColIndex("ToDateH")) = 1200
        End With
         
    End If
    
End If


End Sub

Private Sub fg_Click()

Dim i As Integer

i = val(fg.TextMatrix(fg.Row, fg.ColIndex("id")))

'frmDaysHistory.DurID = i
'frmDaysHistory.show

 Load frmDaysHistory
 frmDaysHistory.show
 frmDaysHistory.Retrive_Det i, cbType.ListIndex

End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos
    Dim str As String

    
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetCustomersSuppliers 2, dcVendor
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    
    With cbType
        If SystemOptions.UserInterface = ArabicInterface Then
               .Clear
               .AddItem ("„Ì·«œÏ")
               .AddItem ("ÂÃ—Ï")
        Else
        .Clear
               .AddItem ("Gregorian")
               .AddItem ("hijri ")
        End If
    End With



    With cbDiff
        .Clear
        .AddItem ("-2")
        .AddItem ("-1")
        .AddItem ("0")
        .AddItem ("1")
        .AddItem ("2")
        
    End With


    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…   "
    LogTexte = " Open Window " & " Confirm  Violation "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
       
    
    Resize_Form Me
    
    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From tbldurations order by ID "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Me.TxtModFlg.Text = "R"
    
Load_Combos
FromDate.value = Date
ToDate.value = Date
FromdateH.value = ToHijriDate(Date)
ToDateH.value = ToHijriDate(Date)
cbType.ListIndex = 0
 XPBtnMove_Click 2
 If OPEN_NEW_SCREEN = True Then
       Cmd_Click (0)
 End If
 

 C1Tab1.CurrTab = 0
 'Retrive_Vacation (txtID.text)
    Exit Sub

ErrTrap:
End Sub

Private Sub Load_Combos()

    Dim sql As String
    sql = " select ID , Name  from tbldurations "
    fill_combo dcDur, sql
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
    sql = "  select id , name  from TblVacationTypes  "
    Else
    sql = " select id , namee  from TblVacationTypes "
    End If
    fill_combo dcVacType, sql


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

 
   Lbl(0).Caption = "No."
   Lbl(3).Caption = " Name Ar"
   Lbl(7).Caption = " Name En"
   Label3.Caption = "City"
   
  Lbl(2).Caption = "Current Record"
  Lbl(4).Caption = "Recors Count"
   
    Me.Caption = "Managerial Area"
    EleHeader.Caption = Me.Caption
   
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "    ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…  "
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


Private Sub FromDate_Change()
     
     FromdateH.value = ToHijriDate(FromDate.value)
     
     
End Sub

Private Sub Fromdateh_LostFocus()

     VBA.Calendar = vbCalGreg
     FromDate.value = ToGregorianDate(FromdateH.value)
   
End Sub

Private Sub ToDate_Change()
   
        ToDateH.value = ToHijriDate(ToDate.value)
      
End Sub

Private Sub ToDateH_LostFocus()
       
       VBA.Calendar = vbCalGreg
       ToDate.value = ToGregorianDate(ToDateH.value)
      
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… "
            Else
                Me.Caption = "Durations"
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
        
            'Me.txtID.locked = True
            'Me.txtName.locked = True
          '  Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
           C1Elastic2.Enabled = False
           C1Elastic3.Enabled = False
           C1Elastic4.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ( ÃœÌœ )"
            Else
                Me.Caption = " Durations (New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
          Me.Cmd(9).Enabled = False
            
            C1Elastic2.Enabled = True
           C1Elastic3.Enabled = True
           C1Elastic4.Enabled = True
          
       
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… (  ⁄œÌ· )"
            Else
                Me.Caption = "Durations (Edit)"
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
        
            C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
            C1Elastic4.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

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
    

    TxtId.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    TxtName.Text = IIf(IsNull(rs("Name").value), "", Trim(rs("name").value))
    cbType.ListIndex = IIf(IsNull(rs("Type").value), -1, Trim(rs("Type").value))
    cbDiff.ListIndex = IIf(IsNull(rs("DayDiff").value), -1, Trim(rs("DayDiff").value))
    FromDate.value = IIf(IsNull(rs("FromDate").value), Date, Trim(rs("FromDate").value))
    FromdateH.value = IIf(IsNull(rs("FromdateH").value), ToHijriDate(Date), Trim(rs("FromdateH").value))
    ToDate.value = IIf(IsNull(rs("ToDate").value), Date, Trim(rs("ToDate").value))
    ToDateH.value = IIf(IsNull(rs("ToDateH").value), ToHijriDate(Date), Trim(rs("ToDateH").value))
    dcDur.BoundText = val(TxtId.Text)
    Retrive_Holidays (val(TxtId.Text))
    
    fg.Rows = 1
    Dim str As String
        str = " select * from TblDurations_details where DID  = " & val(TxtId.Text) & "    order by id    "
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        If Rs_Temp.RecordCount > 0 Then
        fg.Rows = Rs_Temp.RecordCount + 1
        
            Rs_Temp.MoveFirst
            Dim H As Integer
            With fg
                For H = 1 To Rs_Temp.RecordCount
                    .TextMatrix(H, .ColIndex("Serial")) = IIf(IsNull(Rs_Temp("Serial").value), "", Rs_Temp("Serial").value)
                    .TextMatrix(H, .ColIndex("ID")) = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
                    .TextMatrix(H, .ColIndex("FromDate")) = IIf(IsNull(Rs_Temp("FromDate").value), "", Rs_Temp("FromDate").value)
                    .TextMatrix(H, .ColIndex("TODate")) = IIf(IsNull(Rs_Temp("ToDate").value), "", Rs_Temp("ToDate").value)
                    .TextMatrix(H, .ColIndex("FromDateH")) = IIf(IsNull(Rs_Temp("FromDateH").value), "", Rs_Temp("FromDateH").value)
                    .TextMatrix(H, .ColIndex("ToDateH")) = IIf(IsNull(Rs_Temp("ToDateH").value), "", Rs_Temp("ToDateH").value)
                    .TextMatrix(H, .ColIndex("Month")) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
                    
                    summation val(TxtId.Text), IIf(IsNull(Rs_Temp("ID").value), 0, Rs_Temp("ID").value), H
                    Rs_Temp.MoveNext
                Next
            End With
        End If
              Final_Summation
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub Final_Summation()

Dim i As Integer, tot As Integer, daywork As Integer, vacDay As Integer

With fg
For i = 1 To fg.Rows - 1
        If .TextMatrix(i, .ColIndex("ID")) <> "" Then
            tot = tot + val(.TextMatrix(i, .ColIndex("tot")))
            daywork = daywork + val(.TextMatrix(i, .ColIndex("WorkDay")))
            vacDay = vacDay + val(.TextMatrix(i, .ColIndex("VacDays")))
        End If
Next
End With

txttot.Text = tot
txtTotDayVacation.Text = vacDay
txttotDaywork.Text = daywork

End Sub


Private Sub summation(DurID, DDID As Integer, Row As Integer)
    
    Dim str As String
    With fg
     
    str = " select count ( *) as tot  from TblVacationSchedule  where  DDID =  " & DDID
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
        If Rs_Temp2.RecordCount > 0 Then
            .TextMatrix(Row, .ColIndex("tot")) = IIf(IsNull(Rs_Temp2("tot").value), 0, Rs_Temp2("tot").value)
        End If
   
    str = " select count(*) as WorkDay  from TblVacationSchedule  where isvac = 0 and  DDID =  " & DDID
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
        If Rs_Temp2.RecordCount > 0 Then
            .TextMatrix(Row, .ColIndex("WorkDay")) = IIf(IsNull(Rs_Temp2("WorkDay").value), 0, Rs_Temp2("WorkDay").value)
        End If
    
       str = " select count(*) as VacDays  from TblVacationSchedule  where isvac = 1 and  DDID =  " & DDID & " and durationid = " & DurID
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
        If Rs_Temp2.RecordCount > 0 Then
            .TextMatrix(Row, .ColIndex("VacDays")) = IIf(IsNull(Rs_Temp2("VacDays").value), 0, Rs_Temp2("VacDays").value)
        End If
        
        
    End With
End Sub

Private Sub TxtName_GotFocus()
On Error Resume Next
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
 SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
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

 
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If cbType.ListIndex = -1 Then
            MsgBox "„‰ ›÷·ﬂ ‰Ê⁄ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            cbType.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If

        If TxtName.Text = "" Then
            MsgBox "„‰ ›÷·ﬂ  «”„ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtName.SetFocus
            
            Exit Sub
        End If

        Select Case Me.TxtModFlg.Text
            Case "N"
                           
            rs.AddNew
            TxtId.Text = CStr(new_id("tbldurations", "ID", "", True))
            Case "E"
            
              '  StrSQL = "select * From  TblViolationTypes where Name='" & Trim(txtName.text) & "'"
              StrSQL = "delete From TblDurations_details where  DID =" & val(TxtId.Text)
              Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = " delete from TblVacationSchedule where durationID =  " & val(TxtId.Text)
              Cn.Execute StrSQL, , adExecuteNoRecords
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(TxtId.Text)
        rs("type").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("DayDiff").value = IIf(cbDiff.ListIndex = -1, Null, cbDiff.ListIndex)
        rs("Name").value = IIf(TxtName.Text = "", Null, TxtName.Text)
        rs("FromDate") = IIf(IsNull(FromDate.value), Date, FromDate.value)
        rs("FromDateH") = IIf(IsNull(FromdateH.value), ToHijriDate(Date), FromdateH.value)
        rs("ToDate") = IIf(IsNull(ToDate.value), Date, ToDate.value)
        rs("ToDateH") = IIf(IsNull(ToDateH.value), ToHijriDate(Date), ToDateH.value)
        rs("CreationDate") = Date
        rs("UserID") = user_id
        rs.update
        
        Save_WeekVac (val(TxtId.Text))
        
       Dim StrID As String
       Dim j As Integer
       Set Rs_Temp = New ADODB.Recordset
       Rs_Temp.Open " TblDurations_details ", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       With fg
        For j = 1 To fg.Rows - 1
            
            If .TextMatrix(j, .ColIndex("FromDate")) <> "" Then
                Rs_Temp.AddNew
                StrID = CStr(new_id("TblDurations_details", "ID", "", True))
                Rs_Temp("ID") = val(StrID)
                Rs_Temp("DID") = val(TxtId.Text)
                Rs_Temp("Serial") = val(.TextMatrix(j, .ColIndex("Serial")))
                Rs_Temp("FromDate") = Format(.TextMatrix(j, .ColIndex("FromDate")), "yyyy/MM/dd")
                Rs_Temp("FromDateH") = Format(.TextMatrix(j, .ColIndex("FromDateH")), "yyyy/MM/dd")
                Rs_Temp("ToDate") = Format(.TextMatrix(j, .ColIndex("ToDate")), "yyyy/MM/dd")
                Rs_Temp("ToDateH") = .TextMatrix(j, .ColIndex("ToDateH"))
                Rs_Temp("Name") = .TextMatrix(j, .ColIndex("Month"))
                Rs_Temp.update
                
                If cbType.ListIndex = 0 Then
                        Add_Schedule CDate(.TextMatrix(j, .ColIndex("FromDate"))), CDate(.TextMatrix(j, .ColIndex("ToDate"))), val(TxtId.Text), val(StrID)
                ElseIf cbType.ListIndex = 1 Then
                        Add_ScheduleH (.TextMatrix(j, .ColIndex("FromDateH"))), (.TextMatrix(j, .ColIndex("ToDateH"))), val(TxtId.Text), val(StrID)
                End If
            End If
        Next
        End With
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata
        Load_Combos


        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Retrive (val(TxtId.Text))
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                Retrive (val(TxtId.Text))
        End Select
        TxtModFlg.Text = "R"
    End If
save_VacAll
     Retrive (val(TxtId.Text))
     
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(TxtId.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If TxtId.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·› —… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtId.Text) & CHR(13)
        Msg = Msg + "Êﬂ· «·»Ì«‰«  «·„ ⁄·ﬁ… »Â« " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblVacationDays  where  durationid =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "delete From TblVacationSchedule where  durationid =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From TblDurations_details where  DID =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
               StrSQL = "delete From TblDurations where  ID =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "SELECT  *  From TblDurations "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub Del_row()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("id"))
        sr = FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs1.RecordCount < 1 Then
                StrSQL = "delete From TblVacationDays  where  ID =" & val(str)
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblViolationTypes"
                   Rs1.Close
                   Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If rs.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_Vacation (val(dcDur.BoundText))
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_Vacation (val(dcDur.BoundText))
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
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
        .Create Me.hwnd, " ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›…  ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«    ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰  ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«    ⁄—Ì› «·”‰Ê«  «·œ—«”Ì… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub


Private Sub save_VacAll()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  'On Error GoTo errortrap
  
 
 
 
 Cn.Execute "delete  from TblVacationDays where DurationID=" & val(TxtId.Text)
  
    Cn.BeginTrans
    BeginTrans = True
    Set Rs1 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblVacationDays where id=-1 "
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '**********************************************************************************
     Dim j As Integer
      With FgInstallments
        For j = 1 To FgInstallments.Rows - 1
            
            If .TextMatrix(j, .ColIndex("VacationTypeID")) <> "" Then
              
                
    
    Rs1.AddNew
    Rs1("id") = CStr(new_id("TblVacationDays", "id", "", True))
    Rs1("VacationTypeID") = val(.TextMatrix(j, .ColIndex("VacationTypeID")))
    Rs1("VacationType") = (.TextMatrix(j, .ColIndex("VacationType")))
    Rs1("DurationID") = val(TxtId.Text)
    Rs1("Duration") = TxtName.Text
    Rs1("FromDate") = (.TextMatrix(j, .ColIndex("FromDate")))
    Rs1("ToDate") = (.TextMatrix(j, .ColIndex("ToDate")))
    Rs1("FromDateH") = (.TextMatrix(j, .ColIndex("FromDateH")))
    Rs1("ToDateH") = (.TextMatrix(j, .ColIndex("ToDateH")))
    Rs1("Description") = (.TextMatrix(j, .ColIndex("VacationType")))
    
    Dim str2 As String
     If cbType.ListIndex = 0 Then
        str2 = DateDiff("d", Rs1("FromDate"), Rs1("ToDate"), vbSaturday) + 1
    Else
        str2 = DateDiff("d", Rs1("FromDateH"), Rs1("ToDateH"), vbSaturday) + 1
    End If
    Rs1("DaysCount") = str2
    
    Rs1.update
    
     If cbType.ListIndex = 0 Then
              Update_Schedule Rs1("FromDate"), Rs1("ToDate").value, val(TxtId.Text), val(.TextMatrix(j, .ColIndex("VacationTypeID")))
     Else
              Update_ScheduleH Rs1("FromDateH"), Rs1("ToDateH"), val(TxtId.Text), val(.TextMatrix(j, .ColIndex("VacationTypeID")))
     End If
     End If
   Next j
   End With
    Cn.CommitTrans
    BeginTrans = False
    'MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Vacation (val(dcDur.BoundText))
Exit Sub
errortrap:


    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
End Sub


Private Sub save_Vac()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  
 If dcDur.BoundText = "" Then
 MsgBox ("«Œ — «·⁄«„ «·œ—«”Ï «Ê·«")
 dcDur.SetFocus
 SendKeys ("{F4}")
 Exit Sub
 End If
 
 If dcVacType.BoundText = "" Then
 MsgBox ("«Œ —‰Ê⁄ «·⁄ÿ·…")
 dcVacType.SetFocus
 SendKeys ("{F4}")
 Exit Sub
 End If
 
  
    Cn.BeginTrans
    BeginTrans = True
    Set Rs1 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblVacationDays "
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs1.AddNew
    Rs1("id") = CStr(new_id("TblVacationDays", "id", "", True))
    Rs1("VacationTypeID") = IIf(dcVacType.BoundText = "", Null, val(dcVacType.BoundText))
    Rs1("VacationType") = IIf(dcVacType.Text = "", Null, (dcVacType.Text))
    Rs1("DurationID") = IIf(dcDur.BoundText = "", Null, val(dcDur.BoundText))
    Rs1("Duration") = IIf(dcDur.Text = "", Null, (dcDur.Text))
    Rs1("FromDate") = IIf(IsNull(FromDate2.value), Date, FromDate2.value)
    Rs1("ToDate") = IIf(IsNull(ToDate2.value), Date, ToDate2.value)
    Rs1("FromDateH") = IIf(IsNull(FromdateH2.value), ToHijriDate(Date), FromdateH2.value)
    Rs1("ToDateH") = IIf(IsNull(toDateH2.value), ToHijriDate(Date), toDateH2.value)
    Rs1("Description") = Text1.Text
    
    Dim str2 As String
    If FromDate2.Enabled = True Then
        str2 = DateDiff("d", FromDate2.value, ToDate2.value, vbSaturday) + 1
    Else
        str2 = DateDiff("d", FromdateH2.value, toDateH2.value, vbSaturday) + 1
    End If
    Rs1("DaysCount") = str2
    
    Rs1.update
    
     If FromDate2.Enabled = True Then
              Update_Schedule FromDate2.value, ToDate2.value, val(dcDur.BoundText)
     Else
              Update_ScheduleH FromdateH2.value, toDateH2.value, val(dcDur.BoundText)
     End If
   
    Cn.CommitTrans
    BeginTrans = False
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Vacation (val(dcDur.BoundText))
Exit Sub
errortrap:


    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
End Sub
   
Private Sub Retrive_Vacation(DurID As Integer)

Dim i As Integer
     Set Rs1 = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblVacationDays where DurationID = " & DurID
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    FgInstallments.Rows = 1
      
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
        dcDur.BoundText = DurID
         
        With FgInstallments
        .Rows = Rs1.RecordCount + 1
         For i = 1 To FgInstallments.Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("id").value), "", Rs1("id").value)
         .TextMatrix(i, .ColIndex("VacationType")) = IIf(IsNull(Rs1("VacationType").value), "", Rs1("VacationType").value)
         .TextMatrix(i, .ColIndex("VacationTypeID")) = IIf(IsNull(Rs1("VacationTypeID").value), "", Rs1("VacationTypeID").value)
         .TextMatrix(i, .ColIndex("Duration")) = IIf(IsNull(Rs1("Duration").value), "", Rs1("Duration").value)
         .TextMatrix(i, .ColIndex("DurationID")) = IIf(IsNull(Rs1("DurationID").value), "", Rs1("DurationID").value)
         .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(Rs1("FromDate").value), "", Rs1("FromDate").value)
         .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(Rs1("ToDate").value), "", Rs1("ToDate").value)
         .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(Rs1("FromDateH").value), "", Rs1("FromDateH").value)
         .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(Rs1("ToDateH").value), "", Rs1("ToDateH").value)
         .TextMatrix(i, .ColIndex("Description")) = IIf(IsNull(Rs1("Description").value), "", Rs1("Description").value)
          .TextMatrix(i, .ColIndex("DaysCount")) = IIf(IsNull(Rs1("DaysCount").value), "", Rs1("DaysCount").value)
         Rs1.MoveNext
         Next
         End With
    End If
End Sub

Private Sub UpdateRowSchedule(dur As Integer, dt As Date, dth As String, isvac As Boolean, Optional VACID As Integer = 0)
       
       Dim str As String, ss As String
       ss = Format(dt, "yyyy/MM/dd")
       If FromDate.Enabled = True Then
                str = " Select   *  from  TblVacationschedule where Date = '" & ss & "' and  DurationID = " & dur
       Else
                str = " Select   *  from  TblVacationschedule where DateH = '" & Format(dth, "yyyy/MM/dd") & "' and DurationID = " & dur
       End If
       Set rs_Dur = New ADODB.Recordset
       rs_Dur.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
       Dim Color As String
       
       If rs_Dur.RecordCount > 0 Then
       
           If 1 = 1 Then
                   If VACID = 0 Then
                str = " select color from TblVacationTypes where id =  " & dcVacType.BoundText
           Else
           str = " select color from TblVacationTypes where id =  " & VACID
           End If
           
          '      str = " select color from TblVacationTypes where id =  " & dcVacType.BoundText
                Set Rs_Temp = New ADODB.Recordset
                Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs_Temp.RecordCount > 0 Then
                    Color = Rs_Temp("color").value
                End If
           End If
            
            rs_Dur.MoveFirst
            rs_Dur("isvac") = isvac
            If VACID = 0 Then
            rs_Dur("VacationTypeID") = dcVacType.BoundText
            Else
          rs_Dur("VacationTypeID") = VACID
            End If
            rs_Dur("color") = Color
            rs_Dur.update
       End If

End Sub

Private Sub Update_Schedule(FromDate As Date, ToDate As Date, dur As Integer, Optional VACID As Integer = 0)
  
   Dim str As String, str1 As String
   Do While FromDate <= ToDate
        str = Weekday(FromDate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        
        If ISOfficialVacation(FromDate, dur) Then
                 UpdateRowSchedule dur, FromDate, ToHijriDate(FromDate), True, VACID
        Else
                 UpdateRowSchedule dur, FromDate, ToHijriDate(FromDate), False, VACID
        End If
        VBA.Calendar = vbCalGreg
       FromDate = DateAdd("d", 1, FromDate)
   Loop
End Sub


Private Sub Update_ScheduleH(FromDate As String, ToDate As String, dur As Integer, Optional VACID As Integer = 0)
   Dim str As String, str1 As String
   VBA.Calendar = vbCalHijri
   
   
  FromDate = Format(FromDate, "yyyy/MM/dd")
  ToDate = Format(ToDate, "yyyy/MM/dd")
  
   
   Do While FromDate <= ToDate
        str = Weekday(FromDate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        
        If ISOfficialVacationH(FromDate, dur) Then
                 UpdateRowSchedule dur, ToGregorianDate(FromDate), FromDate, True, VACID
        Else
                 UpdateRowSchedule dur, ToGregorianDate(FromDate), FromDate, False, VACID
        End If
        VBA.Calendar = vbCalHijri
        FromDate = DateAdd("d", 1, FromDate)
         FromDate = Format(FromDate, "yyyy/MM/dd")
        VBA.Calendar = vbCalGreg
   Loop
   VBA.Calendar = vbCalGreg
End Sub


Private Function ISOfficialVacation(dt As Date, dur As Integer) As Boolean
    
    Dim str As String, ss As String
    ss = Format(dt, "yyyy/MM/dd")
    
    str = " select * from  TblVacationDays  where DurationID =   " & dur & "  and   '" & ss & "' Between FromDate And ToDate  "
    Set rs_vac = New ADODB.Recordset
    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs_vac.RecordCount > 0 Then
        ISOfficialVacation = True
    End If

End Function

Private Function ISOfficialVacationH(dt As String, dur As Integer) As Boolean
    
    Dim str As String
    str = " select * from  TblVacationDays  where DurationID =   " & dur & "  and    '" & dt & "'   >=  FromDateH  And   '" & dt & "'  <=   ToDateH  "
    Set rs_vac = New ADODB.Recordset
    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs_vac.RecordCount > 0 Then
        ISOfficialVacationH = True
    End If

End Function


