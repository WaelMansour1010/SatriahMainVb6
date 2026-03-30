VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmYearDurations2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«⁄œ«œ«  «·‰ ÌÃ…   "
   ClientHeight    =   9180
   ClientLeft      =   7845
   ClientTop       =   3345
   ClientWidth     =   16170
   Icon            =   "FrmYearDurations2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   16170
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9180
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16170
      _cx             =   28522
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
         Left            =   75
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   16020
         _cx             =   28258
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
            Height          =   315
            Left            =   2205
            TabIndex        =   4
            Top             =   600
            Width           =   4590
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   9630
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            Top             =   240
            Width           =   4560
         End
         Begin VB.ComboBox cbType 
            Height          =   315
            Left            =   2205
            TabIndex        =   2
            Top             =   240
            Width           =   4590
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9630
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   4560
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   330
            Left            =   11805
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   930
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95617027
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal FromdateH 
            Height          =   330
            Left            =   9630
            TabIndex        =   6
            Top             =   930
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   330
            Left            =   4650
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   930
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95617027
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   330
            Left            =   2205
            TabIndex        =   8
            Top             =   930
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   582
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   570
            Index           =   15
            Left            =   150
            TabIndex        =   9
            Top             =   600
            Width           =   1680
            _ExtentX        =   2963
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
            ButtonImage     =   "FrmYearDurations2.frx":038A
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
            Left            =   6915
            TabIndex        =   30
            Top             =   600
            Width           =   2325
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   285
            Index           =   1
            Left            =   13905
            TabIndex        =   29
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì»œ√ „‰ "
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   14070
            TabIndex        =   28
            Top             =   930
            Width           =   1755
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì‰ ÂÏ ›Ï"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   7200
            TabIndex        =   27
            Top             =   930
            Width           =   1965
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ﬁÊÌ„"
            Height          =   285
            Index           =   0
            Left            =   6915
            TabIndex        =   26
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰…"
            Height          =   285
            Index           =   16
            Left            =   14385
            TabIndex        =   25
            Top             =   600
            Width           =   1485
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   720
         Left            =   -60
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   -45
         Width           =   16215
         _cx             =   28601
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
         Caption         =   "«⁄œ«œ«  «·‰ ÌÃ…   "
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
            ButtonImage     =   "FrmYearDurations2.frx":6BEC
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
            ButtonImage     =   "FrmYearDurations2.frx":6F86
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
            ButtonImage     =   "FrmYearDurations2.frx":7320
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
            ButtonImage     =   "FrmYearDurations2.frx":76BA
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
         Left            =   0
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   8325
         Width           =   16170
         _cx             =   28522
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
            Height          =   540
            Index           =   0
            Left            =   14235
            TabIndex        =   32
            Top             =   150
            Width           =   1635
            _ExtentX        =   2884
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
            ButtonImage     =   "FrmYearDurations2.frx":7A54
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
            Left            =   12780
            TabIndex        =   33
            Top             =   150
            Width           =   1425
            _ExtentX        =   2514
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
            ButtonImage     =   "FrmYearDurations2.frx":E2B6
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
            Left            =   10935
            TabIndex        =   34
            Top             =   150
            Width           =   1725
            _ExtentX        =   3043
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
            ButtonImage     =   "FrmYearDurations2.frx":14B18
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
            Left            =   9225
            TabIndex        =   35
            Top             =   150
            Width           =   1605
            _ExtentX        =   2831
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
            ButtonImage     =   "FrmYearDurations2.frx":1B37A
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
            Left            =   7200
            TabIndex        =   36
            Top             =   150
            Width           =   1920
            _ExtentX        =   3387
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
            ButtonImage     =   "FrmYearDurations2.frx":21BDC
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
            Left            =   2085
            TabIndex        =   37
            Top             =   150
            Width           =   1530
            _ExtentX        =   2699
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
            ButtonImage     =   "FrmYearDurations2.frx":2843E
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
            Left            =   390
            TabIndex        =   38
            Top             =   150
            Width           =   1575
            _ExtentX        =   2778
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
            ButtonImage     =   "FrmYearDurations2.frx":52060
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
            Left            =   5625
            TabIndex        =   39
            Top             =   150
            Width           =   1515
            _ExtentX        =   2672
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
            ButtonImage     =   "FrmYearDurations2.frx":588C2
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
            Left            =   3765
            TabIndex        =   48
            Top             =   150
            Width           =   1695
            _ExtentX        =   2990
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
            ButtonImage     =   "FrmYearDurations2.frx":5F124
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
         Left            =   255
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   7560
         Width           =   9330
         _cx             =   16457
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
            Height          =   390
            Index           =   4
            Left            =   1215
            TabIndex        =   44
            Top             =   150
            Width           =   1785
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   390
            Index           =   3
            Left            =   6270
            TabIndex        =   43
            Top             =   150
            Width           =   1785
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   390
            Left            =   150
            TabIndex        =   42
            Top             =   150
            Width           =   1035
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   390
            Left            =   5085
            TabIndex        =   41
            Top             =   150
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5160
         Left            =   0
         TabIndex        =   45
         Top             =   2280
         Width           =   16065
         _cx             =   28337
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
         Caption         =   "«·› —« | ”ÃÌ· «Ì«„ «·⁄ÿ·«  «·—”„Ì…|› —«  «·—« »"
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
            Width           =   15975
            _cx             =   28178
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
               TabIndex        =   54
               Top             =   4410
               Width           =   1530
            End
            Begin VB.TextBox txttotDaywork 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   3345
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   4410
               Width           =   1620
            End
            Begin VB.TextBox txttot 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   7035
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   4410
               Width           =   1905
            End
            Begin VSFlex8UCtl.VSFlexGrid fg 
               Height          =   3300
               Left            =   0
               TabIndex        =   17
               Top             =   1065
               Width           =   16110
               _cx             =   28416
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
               FormatString    =   $"FrmYearDurations2.frx":65986
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
               Height          =   960
               Left            =   0
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   0
               Width           =   16020
               _cx             =   28258
               _cy             =   1693
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
                  Height          =   600
                  Left            =   2490
                  TabIndex        =   16
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox opt_sa 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”» "
                  Height          =   600
                  Left            =   10440
                  TabIndex        =   10
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox opt_su 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Õœ"
                  Height          =   600
                  Left            =   9315
                  TabIndex        =   11
                  Top             =   240
                  Width           =   705
               End
               Begin VB.CheckBox opt_mo 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«À‰Ì‰"
                  Height          =   600
                  Left            =   8040
                  TabIndex        =   12
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox opt_tu 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·À·«À«¡"
                  Height          =   600
                  Left            =   6750
                  TabIndex        =   13
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox opt_We 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«—»⁄«¡"
                  Height          =   600
                  Left            =   5190
                  TabIndex        =   14
                  Top             =   240
                  Width           =   1125
               End
               Begin VB.CheckBox opt_Th 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Œ„Ì”"
                  Height          =   600
                  Left            =   3900
                  TabIndex        =   15
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Label Label12 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «·⁄ÿ·« "
               Height          =   255
               Left            =   1665
               TabIndex        =   53
               Top             =   4410
               Width           =   1590
            End
            Begin VB.Label Label11 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «Ì«„ «·⁄„·"
               Height          =   255
               Left            =   5145
               TabIndex        =   51
               Top             =   4410
               Width           =   1710
            End
            Begin VB.Label Label10 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «Ì«„ «·⁄«„ «·œ—«”Ï"
               Height          =   255
               Left            =   9165
               TabIndex        =   49
               Top             =   4410
               Width           =   2565
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   4785
            Left            =   16710
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   45
            Width           =   15975
            _cx             =   28178
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
               Left            =   2295
               ScrollBars      =   2  'Vertical
               TabIndex        =   56
               Top             =   1260
               Width           =   11925
            End
            Begin MSDataListLib.DataCombo dcVacType 
               Height          =   315
               Left            =   2295
               TabIndex        =   57
               Top             =   480
               Width           =   4560
               _ExtentX        =   8043
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcDur 
               Height          =   315
               Left            =   9570
               TabIndex        =   58
               Top             =   480
               Width           =   4605
               _ExtentX        =   8123
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   15975
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   900
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   95617027
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal ToDateH2 
               Height          =   315
               Left            =   2295
               TabIndex        =   60
               Top             =   885
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker ToDate2 
               Height          =   315
               Left            =   4470
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   885
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   95617027
               CurrentDate     =   41640
            End
            Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
               Height          =   2625
               Left            =   0
               TabIndex        =   62
               Top             =   2115
               Width           =   15975
               _cx             =   28178
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
               FormatString    =   $"FrmYearDurations2.frx":65B4C
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
               Left            =   405
               TabIndex        =   63
               Top             =   480
               Width           =   1395
               _ExtentX        =   2461
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
               ButtonImage     =   "FrmYearDurations2.frx":65D37
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
               Left            =   405
               TabIndex        =   64
               Top             =   960
               Width           =   1395
               _ExtentX        =   2461
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
               ButtonImage     =   "FrmYearDurations2.frx":6C599
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
               Left            =   9570
               TabIndex        =   65
               Top             =   900
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker FromDate2 
               Height          =   315
               Left            =   11820
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   900
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   95617027
               CurrentDate     =   41640
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄«„ «·œ—«”Ï"
               Height          =   270
               Left            =   14400
               TabIndex        =   73
               Top             =   480
               Width           =   1350
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   240
               Left            =   14220
               TabIndex        =   72
               Top             =   900
               Width           =   1530
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘—Õ"
               Height          =   330
               Left            =   14760
               TabIndex        =   71
               Top             =   1305
               Width           =   990
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï  «—ÌŒ"
               Height          =   360
               Left            =   7935
               TabIndex        =   70
               Top             =   900
               Width           =   1230
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   360
               Left            =   18180
               TabIndex        =   69
               Top             =   900
               Width           =   1590
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄«„ «·œ—«”Ï"
               Height          =   405
               Left            =   18180
               TabIndex        =   68
               Top             =   510
               Width           =   1590
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·⁄ÿ·…"
               Height          =   405
               Left            =   7935
               TabIndex        =   67
               Top             =   480
               Width           =   1230
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   4785
            Left            =   17010
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   45
            Width           =   15975
            _cx             =   28178
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
            Begin VB.ComboBox DcbYear 
               Height          =   315
               Left            =   9555
               TabIndex        =   92
               Top             =   1560
               Width           =   4590
            End
            Begin VB.ComboBox DcbMonth 
               Height          =   315
               ItemData        =   "FrmYearDurations2.frx":72DFB
               Left            =   9555
               List            =   "FrmYearDurations2.frx":72DFD
               TabIndex        =   90
               Top             =   1200
               Width           =   4590
            End
            Begin VB.TextBox TxtDes 
               Alignment       =   1  'Right Justify
               Height          =   732
               Left            =   2265
               ScrollBars      =   2  'Vertical
               TabIndex        =   75
               Top             =   1260
               Width           =   4725
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   15975
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   900
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   95617027
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal TDateH 
               Height          =   315
               Left            =   2280
               TabIndex        =   77
               Top             =   780
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker TDate 
               Height          =   315
               Left            =   4590
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   780
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   95617027
               CurrentDate     =   41640
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   2625
               Left            =   0
               TabIndex        =   79
               Top             =   2160
               Width           =   15975
               _cx             =   28178
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmYearDurations2.frx":72DFF
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
            Begin ImpulseButton.ISButton BtnAdd 
               Height          =   405
               Left            =   420
               TabIndex        =   80
               Top             =   480
               Width           =   1380
               _ExtentX        =   2434
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
               ButtonImage     =   "FrmYearDurations2.frx":72F5B
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
            Begin ImpulseButton.ISButton BtnDelete 
               Height          =   375
               Left            =   420
               TabIndex        =   81
               Top             =   960
               Width           =   1380
               _ExtentX        =   2434
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
               ButtonImage     =   "FrmYearDurations2.frx":797BD
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   -2147483637
            End
            Begin Dynamic_Byte.NourHijriCal FrmDateH 
               Height          =   315
               Left            =   9555
               TabIndex        =   82
               Top             =   780
               Width           =   2130
               _ExtentX        =   3969
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker FrmDate 
               Height          =   315
               Left            =   11865
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   780
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   95617027
               CurrentDate     =   41640
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì »⁄ ”‰…"
               Height          =   240
               Left            =   14205
               TabIndex        =   93
               Top             =   1560
               Width           =   1530
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì »⁄ ‘Â—"
               Height          =   240
               Left            =   14205
               TabIndex        =   91
               Top             =   1200
               Width           =   1530
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   240
               Left            =   14205
               TabIndex        =   89
               Top             =   780
               Width           =   1530
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘—Õ"
               Height          =   330
               Left            =   8205
               TabIndex        =   88
               Top             =   1425
               Width           =   960
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï  «—ÌŒ"
               Height          =   360
               Left            =   7950
               TabIndex        =   87
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰  «—ÌŒ"
               Height          =   360
               Left            =   18165
               TabIndex        =   86
               Top             =   900
               Width           =   1590
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄«„ «·œ—«”Ï"
               Height          =   405
               Left            =   18165
               TabIndex        =   85
               Top             =   510
               Width           =   1590
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "› —«  «·—« »"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   405
               Left            =   7200
               TabIndex        =   84
               Top             =   120
               Width           =   1815
            End
         End
      End
   End
End
Attribute VB_Name = "FrmYearDurations2"
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


Private Sub btnAdd_Click()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbMonth.ListIndex) = -1 Or DcbMonth.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·‘Â—"
Else
MsgBox "Please Select Month"
End If
Exit Sub
End If
If val(DcbYear.ListIndex) = -1 Or DcbYear.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·”‰…"
Else
MsgBox "Please Select Year"
End If
Exit Sub
End If
If CheGrid() = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Â–Â «·› —… „ÊÃÊœ… „”»ﬁ«"
Else
MsgBox "This is  period already exists"
End If
Exit Sub
End If
FillGridS
End If
End Sub

Private Sub btnDelete_Click()
If Me.TxtModFlg.Text <> "R" Then
With VSFlexGrid1
If .Rows < 2 Then Exit Sub
.RemoveItem .Row
End With
End If
End Sub

Private Sub cbType_Click()

'Exit Sub

If cbType.ListIndex = 0 Then
    
    FromDateH.Enabled = False
    todateH.Enabled = False
    Fromdate.Enabled = True
    ToDate.Enabled = True
    
    fg.ColWidth(fg.ColIndex("FromDateH")) = 0
    fg.ColWidth(fg.ColIndex("ToDateH")) = 0
    fg.ColWidth(fg.ColIndex("FromDate")) = 1152
    fg.ColWidth(fg.ColIndex("ToDate")) = 1152
ElseIf cbType.ListIndex = 1 Then
    Fromdate.Enabled = False
    ToDate.Enabled = False
     FromDateH.Enabled = True
    todateH.Enabled = True
    
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
                VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
              VSFlexGrid1.Rows = 1
            TxtId.Text = CStr(new_id("Tbldurations2", "ID", "", True))
            TxtName.SetFocus
        Case 1
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "E"
        Case 2
            SaveData
            save_VacAll
            
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
               'salah FrmSearch_Duration.SendForm = "Dur"
              'salah  FrmSearch_Duration.show
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Save_WeekVac(DurID As Integer)
Dim StrSQL As String

On Error GoTo errortrap
   StrSQL = " delete From Tblholidays2  where durationID =" & DurID
   Cn.Execute StrSQL, , adExecuteNoRecords

    Set rsHolidays = New ADODB.Recordset
    StrSQL = "SELECT  *  From Tblholidays2 "
    rsHolidays.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rsHolidays.AddNew
    
    rsHolidays("id") = CStr(new_id("Tblholidays2", "id", "", True))
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
    StrSQL = "SELECT  *  From Tblholidays2 where durationID = " & DurID
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
BeginFirstPer = Fromdate.value
EndFirstPer = dhLastDayInMonth(Fromdate)
EndLastPer = ToDate.value
BeginLastPer = dhFirstDayInMonth(ToDate.value)

Dim count As Integer

VBA.Calendar = vbCalGreg
count = Date_MonthsBetweenDates(CDate(DateAdd("d", 1, EndFirstPer)), CDate(DateAdd("d", -1, BeginLastPer)))

With fg
.Rows = count + 4
If count = -2 Then
    .TextMatrix(1, .ColIndex("FromDate")) = Fromdate.value
    .TextMatrix(1, .ColIndex("ToDate")) = ToDate.value
    .TextMatrix(1, .ColIndex("FromDateh")) = ToHijriDate(Fromdate.value)
     .TextMatrix(1, .ColIndex("ToDateh")) = ToHijriDate(ToDate.value)
     .TextMatrix(1, .ColIndex("Month")) = Format(Fromdate, "MMMM")
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

BeginFirstPer = FromDateH.value

VBA.Calendar = vbCalHijri
EndFirstPer = dhLastDayInMonth(FromDateH.value)
VBA.Calendar = vbCalGreg

EndLastPer = todateH.value
BeginLastPer = dhFirstDayInMonth(todateH.value)
FirstMonth = dhFirstDayInMonth(FromDateH.value)
count = 0
'
Dim s1 As String, s2 As String, s3 As String, s4 As String

s1 = Format(EndFirstPer, "dd/MM/yyyy")
s3 = Format(BeginLastPer, "dd/MM/yyyy")

VBA.Calendar = vbCalHijri
s2 = DateAdd("d", 1, s1)
s4 = DateAdd("d", -1, s3)
VBA.Calendar = vbCalGreg

'VBA.Calendar = vbCalHijri
'count = Date_MonthsBetweenDates(CDate(s1), CDate(s4))
'count = Date_MonthsBetweenDates(CDate("1437/01/30"), CDate("1437/11/01"))
'VBA.Calendar = vbCalGreg

count = Date_MonthsBetweenDates(CDate(s1), CDate(s4))
VBA.Calendar = vbCalGreg



With fg
  If count = -2 Then
    .TextMatrix(1, .ColIndex("FromDate")) = Format(ToGregorianDate(FromDateH.value), "yyyy/MM/dd")
    .TextMatrix(1, .ColIndex("ToDate")) = Format(ToGregorianDate(todateH.value), "yyyy/MM/dd")
    .TextMatrix(1, .ColIndex("FromDateh")) = Format(FromDateH.value, "yyyy/MM/dd")
    .TextMatrix(1, .ColIndex("ToDateh")) = Format(todateH.value, "yyyy/MM/dd")
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

Private Sub Add_Schedule(Fromdate As Date, ToDate As Date, dur As Integer, DDID As Integer)
   Dim str As String, str1 As String, day As String
   
   Do While Fromdate <= ToDate
        str = Weekday(Fromdate, vbSaturday)
        day = WeekdayName(str, False, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        If IsHoliday(str, dur) Then
                 AddRowToSchedule dur, Fromdate, ToHijriDate(Fromdate), True, day, DDID
        Else
                 AddRowToSchedule dur, Fromdate, ToHijriDate(Fromdate), False, day, DDID
        End If
        
       VBA.Calendar = vbCalGreg
       Fromdate = DateAdd("d", 1, Fromdate)
   Loop
End Sub


Private Sub Add_ScheduleH(Fromdate As String, ToDate As String, dur As Integer, DDID As Integer)
  
  Dim str As String, str1 As String, day As String
  Fromdate = Format(Fromdate, "yyyy/MM/dd")
  ToDate = Format(ToDate, "yyyy/MM/dd")
   
   Do While Fromdate <= ToDate
        VBA.Calendar = vbCalHijri
        str = Weekday(Fromdate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        day = WeekdayName(Weekday(Fromdate, vbSaturday), False, vbSaturday)
        VBA.Calendar = vbCalGreg
        
        If IsHoliday(str, dur) Then
                 AddRowToSchedule dur, ToGregorianDate(Fromdate), Fromdate, True, day, DDID
        Else
                 AddRowToSchedule dur, ToGregorianDate(Fromdate), Fromdate, False, day, DDID
        End If
         
         VBA.Calendar = vbCalHijri
         Fromdate = DateAdd("d", 1, Fromdate)
        VBA.Calendar = vbCalGreg
             
        Fromdate = Format(Fromdate, "yyyy/MM/dd")
   Loop

End Sub


Private Function IsHoliday(day As String, dur As Integer) As Boolean
    Dim str As String
    str = " select * from  Tblholidays2  where DurationID =" & dur
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
'    str = " select * from  TblVacationDays2  where DurationID =   " & dur & "  and   " & dt & " Between FromDate And ToDate  "
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
       rs_Dur.Open " TblVacationschedule22 ", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
       rs_Dur.AddNew
       rs_Dur("ID") = CStr(new_id("TblVacationschedule22", "ID", "", True))
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


 

Private Sub DcDur_Change()


Dim str  As String
Set rsDuration = New ADODB.Recordset
str = " select * from Tbldurations2 where id =  " & val(dcDur.BoundText)
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

 i = val(fg.TextMatrix(fg.Row, fg.ColIndex("id")))

'frmDaysHistory.DurID = i
'frmDaysHistory.show

 Load frmDaysHistory
 frmDaysHistory.show
 frmDaysHistory.Retrive_Det2 i, cbType.ListIndex


End Sub


Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    DcbMonth.Clear

    For i = 1 To 12
        DcbMonth.AddItem MonthName(i)
    Next

    DcbMonth.ListIndex = Month(Date) - 1
    DcbYear.Clear

    For i = 2015 To 4050
        DcbYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = DcbYear.NewIndex
        End If

    Next

    DcbYear.ListIndex = IntDefIndex

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

    YearMonth
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
    StrSQL = "SELECT  *  From Tbldurations2 order by ID "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Me.TxtModFlg.Text = "R"
    
Load_Combos
Fromdate.value = Date
ToDate.value = Date
FromDateH.value = ToHijriDate(Date)
todateH.value = ToHijriDate(Date)
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
    sql = " select ID , Name  from Tbldurations2 "
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
    
    EleHeader.Caption = "Calendar settings"
    lbl(1).Caption = "Serial"
    lbl(16).Caption = "Academic year"
    Label2.Caption = "Starts on"
    lbl(0).Caption = "Calendar type"
    lbl(2).Caption = "Hijri Days Difference"
    Label4.Caption = "Ends on"
    Cmd(15).Caption = "Add"
    
   '########################################### First Tab ##########################################
    
    C1Elastic4.Caption = "Weekends"
    opt_sa.Caption = "Sat"
    opt_su.Caption = "Sun"
    opt_mo.Caption = "Mon"
    opt_tu.Caption = "Tue"
    opt_We.Caption = "Wed"
    opt_Th.Caption = "Thu"
    opt_Fr.Caption = "Fri"
    
    With Me.fg
        .TextMatrix(0, .ColIndex("Serial")) = "S"
        .TextMatrix(0, .ColIndex("Month")) = "Period Name"
        .TextMatrix(0, .ColIndex("FromDate")) = "Starts in G"
        .TextMatrix(0, .ColIndex("FromDateH")) = "Starts in H"
        .TextMatrix(0, .ColIndex("ToDate")) = "Ends on G"
        .TextMatrix(0, .ColIndex("ToDateH")) = "Ends on H"
        .TextMatrix(0, .ColIndex("tot")) = "No. of days in the month"
        .TextMatrix(0, .ColIndex("VacDays")) = "Vacation days"
        .TextMatrix(0, .ColIndex("WorkDay")) = "Work days"
        .TextMatrix(0, .ColIndex("Notes")) = "Notes"
    End With
    
    Label10.Caption = "Total days in academic year"
    Label11.Caption = "Total work days"
    Label12.Caption = "Total vacation days"
    
    '###########################################################################################
    
    '###################################### Second Tab #########################################
    
    C1Elastic7.Caption = "Other vacations"
    Label9.Caption = "Academic year"
    Label8.Caption = "Starts in"
    Label7.Caption = "Description"
    Label1.Caption = "Vacations Type"
    Label6.Caption = "Ends on"
    Cmd(5).Caption = "Add"
    Cmd(8).Caption = "Remove"
    
        With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "S"
        .TextMatrix(0, .ColIndex("FromDate")) = "From Date"
        .TextMatrix(0, .ColIndex("FromDateH")) = "From Date"
        .TextMatrix(0, .ColIndex("ToDate")) = "To Date"
        .TextMatrix(0, .ColIndex("ToDateH")) = "To Date"
       .TextMatrix(0, .ColIndex("Monthn")) = "Month"
       .TextMatrix(0, .ColIndex("YearID")) = "Year"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
    With Me.FgInstallments
        .TextMatrix(0, .ColIndex("Serial")) = "S"
        .TextMatrix(0, .ColIndex("Duration")) = "Academic year"
        .TextMatrix(0, .ColIndex("VacationType")) = "Vacation Type"
        .TextMatrix(0, .ColIndex("FromDate")) = "Starts in G"
        .TextMatrix(0, .ColIndex("FromDateH")) = "Starts in H"
        .TextMatrix(0, .ColIndex("ToDate")) = "Ends on G"
        .TextMatrix(0, .ColIndex("ToDateH")) = "Ends on H"
        .TextMatrix(0, .ColIndex("DaysCount")) = "Total vacation days"
        .TextMatrix(0, .ColIndex("Description")) = "Description"
    End With
    Label18.Caption = "From Date"
    Label19.Caption = "Month"
    Label20.Caption = "Year"
    '###########################################################################################
    Label17.Caption = "Remarks"
    Label16.Caption = "To Date"
    btnDelete.Caption = "Delete"
    C1Tab1.Caption = "Periods | Registration of public holidays | Period Salary"
    Label13.Caption = "Period Salary"
    lbl(3).Caption = "Current record"
    lbl(4).Caption = "Number of records"
    btnAdd.Caption = "Add"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Cancel"
    Cmd(4).Caption = "Delete"
    Cmd(7).Caption = "Print"
    Cmd(9).Caption = "Search"
    Cmd(6).Caption = "Exit"
    CmdAttach.Caption = "Attachments"

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


Private Sub FrmDate_Change()
If Me.TxtModFlg.Text <> "R" Then
FrmDateH.value = ToHijriDate(FrmDate.value)
End If
End Sub

Private Sub FrmDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
VBA.Calendar = vbCalGreg
       FrmDate.value = ToGregorianDate(FrmDateH.value)
       End If
End Sub

Private Sub FromDate_Change()
     
     FromDateH.value = ToHijriDate(Fromdate.value)
     
     
End Sub

Private Sub Fromdateh_LostFocus()

     VBA.Calendar = vbCalGreg
     Fromdate.value = ToGregorianDate(FromDateH.value)
   
End Sub

Private Sub TDate_Change()
If Me.TxtModFlg.Text <> "R" Then
TDateH.value = ToHijriDate(TDate.value)
End If
End Sub

Private Sub TDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
VBA.Calendar = vbCalGreg
       TDate.value = ToGregorianDate(TDateH.value)
End If
End Sub

Private Sub ToDate_Change()
   
        todateH.value = ToHijriDate(ToDate.value)
      
End Sub

Private Sub ToDateH_LostFocus()
       
       VBA.Calendar = vbCalGreg
       ToDate.value = ToGregorianDate(todateH.value)
      
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
    

    TxtId.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    TxtName.Text = IIf(IsNull(rs("Name").value), "", Trim(rs("name").value))
    cbType.ListIndex = IIf(IsNull(rs("Type").value), -1, Trim(rs("Type").value))
    cbDiff.ListIndex = IIf(IsNull(rs("DayDiff").value), -1, Trim(rs("DayDiff").value))
    Fromdate.value = IIf(IsNull(rs("FromDate").value), Date, Trim(rs("FromDate").value))
    FromDateH.value = IIf(IsNull(rs("FromdateH").value), ToHijriDate(Date), Trim(rs("FromdateH").value))
    ToDate.value = IIf(IsNull(rs("ToDate").value), Date, Trim(rs("ToDate").value))
    todateH.value = IIf(IsNull(rs("ToDateH").value), ToHijriDate(Date), Trim(rs("ToDateH").value))
    
    FrmDate.value = IIf(IsNull(rs("FrmDate").value), Date, Trim(rs("FrmDate").value))
    TDate.value = IIf(IsNull(rs("TDate").value), Date, Trim(rs("TDate").value))
    FrmDateH.value = IIf(IsNull(rs("FrmDateH").value), ToHijriDate(FrmDate.value), Trim(rs("FrmDateH").value))
    TDateH.value = IIf(IsNull(rs("TDateH").value), ToHijriDate(TDate.value), Trim(rs("TDateH").value))
    Me.DcbYear.ListIndex = IIf(IsNull(rs("YearID").value), -1, (rs("YearID").value))
    Me.DcbMonth.ListIndex = IIf(IsNull(rs("MonthID").value), -1, (rs("MonthID").value))
    TxtDes.Text = IIf(IsNull(rs("Des").value), "", Trim(rs("Des").value))
    
    dcDur.BoundText = val(TxtId.Text)
    Retrive_Holidays (val(TxtId.Text))
    
    fg.Rows = 1
    Dim str As String
        str = " select * from Tbldurations_details2 where DID  = " & val(TxtId.Text) & "    order by id    "
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
              FillGridSalary
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Sub FillGridS()
Dim k As Integer
Dim i As Integer
With VSFlexGrid1
k = .Rows
.Rows = .Rows + 1
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("FromDate")) = FrmDate.value
.TextMatrix(i, .ColIndex("FromDateH")) = FrmDateH.value
.TextMatrix(i, .ColIndex("ToDate")) = TDate.value
.TextMatrix(i, .ColIndex("ToDateH")) = TDateH.value
.TextMatrix(i, .ColIndex("MonthID")) = val(DcbMonth.ListIndex) + 1
.TextMatrix(i, .ColIndex("YearID")) = val(DcbYear.Text)
.TextMatrix(i, .ColIndex("Remarks")) = TxtDes.Text
.TextMatrix(i, .ColIndex("Monthn")) = MonthName((val(.TextMatrix(i, .ColIndex("MonthID")))))
Next i
End With
End Sub
Sub FillGridSalary()
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
   VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
              VSFlexGrid1.Rows = 1
sql = "select * from TblDurations2Salary where DurID =" & val(TxtId.Text) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With VSFlexGrid1
.Rows = Rs3.RecordCount + 1
Rs3.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(Rs3("FromDate").value), "", Rs3("FromDate").value)
.TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(Rs3("ToDate").value), "", Rs3("ToDate").value)
.TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(Rs3("FromDateH").value), "", Rs3("FromDateH").value)
.TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(Rs3("ToDateH").value), "", Rs3("ToDateH").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs3("Remarks").value), "", Rs3("Remarks").value)
.TextMatrix(i, .ColIndex("YearID")) = IIf(IsNull(Rs3("YearID").value), "", Rs3("YearID").value)
.TextMatrix(i, .ColIndex("MonthID")) = IIf(IsNull(Rs3("MonthID").value), 0, Rs3("MonthID").value)
.TextMatrix(i, .ColIndex("Monthn")) = MonthName(val(.TextMatrix(i, .ColIndex("MonthID"))))
Rs3.MoveNext
Next i
End With
End If
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
     
    str = " select count ( *) as tot  from TblVacationschedule22  where  DDID =  " & DDID
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
        If Rs_Temp2.RecordCount > 0 Then
            .TextMatrix(Row, .ColIndex("tot")) = IIf(IsNull(Rs_Temp2("tot").value), 0, Rs_Temp2("tot").value)
        End If
   
    str = " select count(*) as WorkDay  from TblVacationschedule22  where isvac = 0 and  DDID =  " & DDID
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
        If Rs_Temp2.RecordCount > 0 Then
            .TextMatrix(Row, .ColIndex("WorkDay")) = IIf(IsNull(Rs_Temp2("WorkDay").value), 0, Rs_Temp2("WorkDay").value)
        End If
    
       str = " select count(*) as VacDays  from TblVacationschedule22  where isvac = 1 and  DDID =  " & DDID & " and durationid = " & DurID
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


Sub SaveSalary()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Dim sql As String
If Me.TxtModFlg.Text = "E" Then
Cn.Execute "Delete From TblDurations2Salary where DurID =" & val(TxtId.Text) & ""
End If
sql = "Select * from  TblDurations2Salary where 1=-1"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With VSFlexGrid1
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("FromDate")) <> "" Then
Rs3.AddNew
Rs3("DurID").value = val(TxtId.Text)
Rs3("FromDate").value = IIf(.TextMatrix(i, .ColIndex("FromDate")) = "", Null, .TextMatrix(i, .ColIndex("FromDate")))
Rs3("ToDate").value = IIf(.TextMatrix(i, .ColIndex("ToDate")) = "", Null, .TextMatrix(i, .ColIndex("ToDate")))
Rs3("FromDateH").value = IIf(.TextMatrix(i, .ColIndex("FromDateH")) = "", Null, .TextMatrix(i, .ColIndex("FromDateH")))
Rs3("ToDateH").value = IIf(.TextMatrix(i, .ColIndex("ToDateH")) = "", Null, .TextMatrix(i, .ColIndex("ToDateH")))
Rs3("Remarks").value = IIf(.TextMatrix(i, .ColIndex("Remarks")) = "", "", .TextMatrix(i, .ColIndex("Remarks")))
Rs3("YearID").value = IIf(.TextMatrix(i, .ColIndex("YearID")) = "", Null, val(.TextMatrix(i, .ColIndex("YearID"))))
Rs3("MonthID").value = IIf(.TextMatrix(i, .ColIndex("MonthID")) = "", Null, val(.TextMatrix(i, .ColIndex("MonthID"))))
Rs3.update
End If
Next i
End With
End Sub

Function CheGrid() As Boolean
Dim i As Integer
With VSFlexGrid1
CheGrid = True
For i = 1 To .Rows - 1
If i <> 1 Then
If .TextMatrix(i, .ColIndex("ToDate")) <> "" Then
If .TextMatrix(i, .ColIndex("ToDate")) >= FrmDate.value Then
CheGrid = False
Exit Function
End If
End If
End If
Next i
End With
End Function

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
          If SystemOptions.UserInterface = EnglishInterface Then
          MsgBox "please, Calendar type ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          Else
          MsgBox "„‰ ›÷·ﬂ ‰Ê⁄ «· ﬁÊÌ„ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          End If
            cbType.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If

        If TxtName.Text = "" Then
          If SystemOptions.UserInterface = EnglishInterface Then
          MsgBox "Please, Period name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          Else
          MsgBox "„‰ ›÷·ﬂ  «”„ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          End If
            TxtName.SetFocus
            
            Exit Sub
        End If

        Select Case Me.TxtModFlg.Text
            Case "N"
                           
            rs.AddNew
            TxtId.Text = CStr(new_id("Tbldurations2", "ID", "", True))
            Case "E"
            
              '  StrSQL = "select * From  TblViolationTypes where Name='" & Trim(txtName.text) & "'"
              StrSQL = "delete From Tbldurations_details2 where  DID =" & val(TxtId.Text)
              Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = " delete from TblVacationschedule22 where durationID =  " & val(TxtId.Text)
              Cn.Execute StrSQL, , adExecuteNoRecords
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(TxtId.Text)
        rs("type").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("DayDiff").value = IIf(cbDiff.ListIndex = -1, Null, cbDiff.ListIndex)
        rs("Name").value = IIf(TxtName.Text = "", Null, TxtName.Text)
        rs("FromDate") = IIf(IsNull(Fromdate.value), Date, Fromdate.value)
        rs("FromDateH") = IIf(IsNull(FromDateH.value), ToHijriDate(Date), FromDateH.value)
        rs("ToDate") = IIf(IsNull(ToDate.value), Date, ToDate.value)
        rs("ToDateH") = IIf(IsNull(todateH.value), ToHijriDate(Date), todateH.value)
        rs("CreationDate") = Date
        rs("UserID") = user_id
        rs("FrmDate").value = FrmDate.value
        rs("TDate").value = TDate.value
        rs("FrmDateH").value = FrmDateH.value
        rs("TDateH").value = TDateH.value
        rs("YearID").value = val(Me.DcbYear.ListIndex)
        rs("MonthID").value = val(Me.DcbMonth.ListIndex)
        rs("Des").value = TxtDes.Text
        rs.update
        
        Save_WeekVac (val(TxtId.Text))
        
       Dim StrID As String
       Dim j As Integer
       Set Rs_Temp = New ADODB.Recordset
       Rs_Temp.Open " Tbldurations_details2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       With fg
        For j = 1 To fg.Rows - 1
            
            If .TextMatrix(j, .ColIndex("FromDate")) <> "" Then
                Rs_Temp.AddNew
                StrID = CStr(new_id("Tbldurations_details2", "ID", "", True))
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
SaveSalary
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
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
      If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Data cannot be saved" & CHR(13)
        Msg = Msg + "invalid values was entered" & CHR(13)
        Msg = Msg + "please make sure you entered a valid data and try again"
      Else
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
      End If
       
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
    Msg = "Sorry, something went wrong while saving the data" & CHR(13)
    Else
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
        
          If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "The recored No." & (TxtId.Text) & CHR(13)
            Msg = Msg + "will be deleted" & CHR(13)
            Msg = Msg + "and all data associated with it" & CHR(13)
            Msg = Msg + "are you sure you want to delete this record"
          Else
            Msg = "”Ì „ Õ–› »Ì«‰«  «·› —… —ﬁ„ " & CHR(13)
            Msg = Msg + (TxtId.Text) & CHR(13)
            Msg = Msg + "Êﬂ· «·»Ì«‰«  «·„ ⁄·ﬁ… »Â« " & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
          End If

    

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblVacationDays2  where  durationid =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "delete From TblVacationschedule22 where  durationid =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From Tbldurations_details2 where  DID =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
               StrSQL = "delete From Tbldurations2 where  ID =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Cn.Execute "Delete From TblDurations2Salary where DurID =" & val(TxtId.Text) & ""
                   StrSQL = "SELECT  *  From Tbldurations2 "
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
        
        If SystemOptions.UserInterface = EnglishInterface Then
          Msg = "This operation is not available right now because there's no records"
        Else
          Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        End If
        
        
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    
    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = "Sorry this record cannot be deleted because it associated with other data"
    Else
      Msg = "⁄›Ê« ·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· ·«— »«ÿ… »»Ì«‰«  «Œ—Ï"
    End If
    
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
        
        If SystemOptions.UserInterface = EnglishInterface Then
          Msg = "The row No." & (sr) & CHR(13)
          Msg = Msg + "will be deleted" & CHR(13)
          Msg = Msg + "are you sure you want to delete this row"
        Else
          Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
          Msg = Msg + (sr) & CHR(13)
          Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        End If
 
        

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not Rs1.RecordCount < 1 Then
                StrSQL = "delete From TblVacationDays2  where  ID =" & val(str)
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
        If SystemOptions.UserInterface = EnglishInterface Then
          Msg = "This operation is not available because there's no records"
        Else
          Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        End If
        
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_Vacation (val(dcDur.BoundText))
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    
     If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Sorry this record cannot be deleted because it associated with other data"
       'Msg = "" & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
     Else
       'Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
        Msg = "⁄›Ê« ·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· ·«— »«ÿ… »»Ì«‰«  «Œ—Ï"
     End If
    
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
  
 
 
 
 Cn.Execute "delete  from TblVacationDays2"
  
    Cn.BeginTrans
    BeginTrans = True
    Set Rs1 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblVacationDays2 where id=-1 "
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '**********************************************************************************
     Dim j As Integer
      With FgInstallments
        For j = 1 To FgInstallments.Rows - 1
            
            If .TextMatrix(j, .ColIndex("VacationTypeID")) <> "" Then
              
                
    
    Rs1.AddNew
    Rs1("id") = CStr(new_id("TblVacationDays2", "id", "", True))
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
     If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Data cannot be saved" & CHR(13)
        Msg = Msg + "invalid values was entered" & CHR(13)
        Msg = Msg + "please make sure you entered a valid data and try again"
      Else
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
      End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = "Sorry, something went wrong while saving the data" & CHR(13)
    Else
      Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    End If
    
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
End Sub


Private Sub save_Vac()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  
 If dcDur.BoundText = "" Then
 
   If SystemOptions.UserInterface = EnglishInterface Then
     MsgBox ("Please enter the academic year first")
   Else
     MsgBox ("«Œ — «·⁄«„ «·œ—«”Ï «Ê·«")
   End If

 
 dcDur.SetFocus
 SendKeys ("{F4}")
 Exit Sub
 End If
 
 If dcVacType.BoundText = "" Then
   If SystemOptions.UserInterface = EnglishInterface Then
   MsgBox ("Please select vacation type")
   Else
   MsgBox ("«Œ —‰Ê⁄ «·⁄ÿ·…")
   End If
 dcVacType.SetFocus
 SendKeys ("{F4}")
 Exit Sub
 End If
 
  
    Cn.BeginTrans
    BeginTrans = True
    Set Rs1 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblVacationDays2 "
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs1.AddNew
    Rs1("id") = CStr(new_id("TblVacationDays2", "id", "", True))
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
    
     If cbType.ListIndex = 0 Then
              Update_Schedule FromDate2.value, ToDate2.value, val(dcDur.BoundText)
     Else
              Update_ScheduleH FromdateH2.value, toDateH2.value, val(dcDur.BoundText)
     End If
   
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = EnglishInterface Then
      MsgBox ("Data was saved successfully")
    Else
      MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    End If
  
    Retrive_Vacation (val(dcDur.BoundText))
Exit Sub
errortrap:


    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
     If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Data cannot be saved" & CHR(13)
        Msg = Msg + "invalid values was entered" & CHR(13)
        Msg = Msg + "please make sure you entered a valid data and try again"
      Else
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
      End If
      
      MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      Exit Sub
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = "Sorry, something went wrong while saving the data" & CHR(13)
    Else
      Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    End If
    
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
End Sub
   
Private Sub Retrive_Vacation(DurID As Integer)

Dim i As Integer
     Set Rs1 = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblVacationDays2 where DurationID = " & DurID
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
         If cbType.ListIndex = 0 Then
                str = " Select   *  from  TblVacationschedule22 where Date = '" & ss & "' and  DurationID = " & dur
       Else
                str = " Select   *  from  TblVacationschedule22 where DateH = '" & Format(dth, "yyyy/MM/dd") & "' and DurationID = " & dur
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
                Set Rs_Temp = New ADODB.Recordset
                Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                Color = vbRed
                If Rs_Temp.RecordCount > 0 Then
                    Color = Rs_Temp("color").value
                End If
           End If
            
            rs_Dur.MoveFirst
            
            
            
            
            rs_Dur("isvac") = isvac
            rs_Dur("VacationTypeID") = VACID
            rs_Dur("color") = Color
            rs_Dur.update
       End If

End Sub

Private Sub Update_Schedule(Fromdate As Date, ToDate As Date, dur As Integer, Optional VACID As Integer = 0)
  
   Dim str As String, str1 As String
   Do While Fromdate <= ToDate
        str = Weekday(Fromdate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        
        If ISOfficialVacation(Fromdate, dur) Then
                 'UpdateRowSchedule dur, FromDate, ToHijriDate(FromDate), True, VACID
        Else
                 'UpdateRowSchedule dur, FromDate, ToHijriDate(FromDate), False, VACID
        End If
        UpdateRowSchedule dur, Fromdate, ToHijriDate(Fromdate), True, VACID
        
        VBA.Calendar = vbCalGreg
       Fromdate = DateAdd("d", 1, Fromdate)
   Loop
End Sub


Private Sub Update_ScheduleH(Fromdate As String, ToDate As String, dur As Integer, Optional VACID As Integer = 0)
   Dim str As String, str1 As String
   VBA.Calendar = vbCalHijri
   
   
  Fromdate = Format(Fromdate, "yyyy/MM/dd")
  ToDate = Format(ToDate, "yyyy/MM/dd")
  
   
   Do While Fromdate <= ToDate
        str = Weekday(Fromdate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        
        If ISOfficialVacationH(Fromdate, dur) Then
                 UpdateRowSchedule dur, ToGregorianDate(Fromdate), Fromdate, True, VACID
        Else
                 UpdateRowSchedule dur, ToGregorianDate(Fromdate), Fromdate, False, VACID
        End If
        VBA.Calendar = vbCalHijri
        Fromdate = DateAdd("d", 1, Fromdate)
         Fromdate = Format(Fromdate, "yyyy/MM/dd")
        VBA.Calendar = vbCalGreg
   Loop
   VBA.Calendar = vbCalGreg
End Sub


Private Function ISOfficialVacation(dt As Date, dur As Integer) As Boolean
    
    Dim str As String, ss As String
    ss = Format(dt, "yyyy/MM/dd")
    
    str = " select * from  TblVacationDays2  where DurationID =   " & dur & "  and   '" & ss & "' Between FromDate And ToDate  "
    Set rs_vac = New ADODB.Recordset
    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs_vac.RecordCount > 0 Then
        ISOfficialVacation = True
    End If

End Function

Private Function ISOfficialVacationH(dt As String, dur As Integer) As Boolean
    
    Dim str As String
    str = " select * from  TblVacationDays2  where DurationID =   " & dur & "  and    '" & dt & "'   >=  FromDateH  And   '" & dt & "'  <=   ToDateH  "
    Set rs_vac = New ADODB.Recordset
    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs_vac.RecordCount > 0 Then
        ISOfficialVacationH = True
    End If

End Function


