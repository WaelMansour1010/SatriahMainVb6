VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmADVPaymentsAlloc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Œ’Ì’ «·„ﬁ»Ê÷«  ··›Ê« Ì— Ê «·„” Œ·’« "
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12270
   Icon            =   "FrmADVPaymentsAlloc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   12270
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   585
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   12225
      _cx             =   21564
      _cy             =   1032
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " Œ’Ì’ «·„ﬁ»Ê÷«  ··›Ê« Ì— Ê «·„” Œ·’« "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      Begin VB.TextBox oldtxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   5460
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1125
         TabIndex        =   8
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmADVPaymentsAlloc.frx":000C
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
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmADVPaymentsAlloc.frx":03A6
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
         Height          =   375
         Index           =   1
         Left            =   1650
         TabIndex        =   10
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmADVPaymentsAlloc.frx":0740
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
         Height          =   375
         Index           =   3
         Left            =   585
         TabIndex        =   11
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmADVPaymentsAlloc.frx":0ADA
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   2400
         TabIndex        =   12
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin MSAdodcLib.Adodc numbering 
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
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   -360
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Index           =   11
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   60
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   7800
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12210
      _cx             =   21537
      _cy             =   13758
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
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "«·„ﬁ»Ê÷« | »Ì«‰«   „” Œ·’«  «·„‘«—Ì⁄"
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
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmADVPaymentsAlloc.frx":0E74
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7335
         Index           =   12
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   12120
         _cx             =   21378
         _cy             =   12938
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
         Begin VB.TextBox TxtManulaNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   240
            Visible         =   0   'False
            Width           =   1515
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   315
            Left            =   -1560
            TabIndex        =   146
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
            _extentx        =   2778
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   6750
            TabIndex        =   75
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104988673
            CurrentDate     =   38784
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄·Ê„«  «·ÕÊ«·Â"
            Enabled         =   0   'False
            Height          =   1095
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   3000
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   240
               Width           =   2565
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   120
               TabIndex        =   133
               Top             =   570
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
               _Version        =   393216
               Format          =   104988673
               CurrentDate     =   39614
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÕÊ«·Â"
               Height          =   285
               Index           =   45
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒÂ«"
               Height          =   285
               Index           =   44
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   570
               Width           =   735
            End
         End
         Begin VB.TextBox TxtCustCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Text            =   " "
            Top             =   1350
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Text            =   " "
            Top             =   600
            Width           =   1395
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Height          =   1005
            Index           =   0
            Left            =   12150
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   270
            Width           =   3735
            Begin VB.TextBox TxtTransID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   120
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox TxtTransSerial 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   570
               Width           =   1005
            End
            Begin VB.ComboBox CboTrans 
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   240
               Width           =   1995
            End
            Begin ImpulseButton.ISButton CmdSearchTrans 
               Height          =   345
               Left            =   600
               TabIndex        =   82
               Top             =   570
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "FrmADVPaymentsAlloc.frx":120E
            End
            Begin ImpulseButton.ISButton CmdOpenTrans 
               Height          =   345
               Left            =   90
               TabIndex        =   83
               Top             =   570
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "FrmADVPaymentsAlloc.frx":15A8
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   315
               Index           =   10
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   630
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   255
               Index           =   12
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   300
               Width           =   1305
            End
         End
         Begin VB.ComboBox DCboCashType 
            Height          =   315
            Left            =   8340
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   975
            Width           =   2265
         End
         Begin VB.TextBox XPMTxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   585
            Left            =   15810
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   4650
            Width           =   2715
         End
         Begin VB.TextBox XPTxtVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13680
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   2415
            Width           =   2685
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ Õ”«» ›« Ê—…"
            Height          =   195
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   120
            Width           =   1575
         End
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   13680
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   2760
            Width           =   2685
         End
         Begin VB.Frame FraNote 
            BackColor       =   &H00E2E9E9&
            Height          =   1965
            Left            =   14220
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   3120
            Width           =   4155
            Begin VB.TextBox TXTBankName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   480
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   810
               Width           =   2685
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   30
               TabIndex        =   61
               Top             =   1140
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Format          =   104988673
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   30
               TabIndex        =   63
               Top             =   480
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   30
               TabIndex        =   64
               Top             =   150
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcChequeBox 
               Height          =   315
               Left            =   0
               TabIndex        =   130
               Top             =   1560
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
               Height          =   285
               Index           =   43
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   9
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   15
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   510
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   285
               Index           =   16
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
               Height          =   285
               Index           =   17
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   1140
               Width           =   1215
            End
         End
         Begin VB.Frame FraInfo 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄·Ê„«   Â„ﬂ"
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
            Height          =   2235
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   8640
            Width           =   3705
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   0
               Left            =   1830
               TabIndex        =   41
               Top             =   780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":1942
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   42
               Top             =   780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":1AA4
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   2
               Left            =   1830
               TabIndex        =   43
               Top             =   1350
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":1C06
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   44
               Top             =   1350
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":1D68
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   4
               Left            =   1830
               TabIndex        =   45
               Top             =   1920
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":1ECA
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   5
               Left            =   120
               TabIndex        =   46
               Top             =   1920
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":202C
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   6
               Left            =   120
               TabIndex        =   47
               Top             =   540
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":218E
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   7
               Left            =   120
               TabIndex        =   48
               Top             =   1110
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":22F0
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   8
               Left            =   120
               TabIndex        =   49
               Top             =   1680
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
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
               MouseIcon       =   "FrmADVPaymentsAlloc.frx":2452
               BackColor       =   14871017
               Alignment       =   1
               Caption         =   ""
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬁ»Ê÷«  ›Ï «·≈”»Ê⁄ «·Õ«·Ï:"
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
               Index           =   19
               Left            =   1380
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   1110
               Width           =   2235
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬁ»Ê÷«  ›Ï «·‘Â— «·Õ«·Ï :"
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
               Index           =   20
               Left            =   1380
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   1680
               Width           =   2235
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁœÌ"
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
               Index           =   21
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1350
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·≈”»Ê⁄ «·Õ«·Ï"
               Height          =   255
               Index           =   22
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Ã„«·Ï „ﬁ»Ê÷«  «·ÌÊ„:"
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
               Index           =   23
               Left            =   1380
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   540
               Width           =   2235
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Ìﬂ« "
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
               Index           =   24
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁœÌ"
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
               Index           =   25
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   1920
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Ìﬂ« "
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
               Index           =   26
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   1920
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁœÌ"
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
               Index           =   27
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   780
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Ìﬂ« "
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
               Index           =   28
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   780
               Width           =   675
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
            Height          =   885
            Index           =   1
            Left            =   13620
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   5160
            Width           =   8175
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   200
               Width           =   1875
            End
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   90
               TabIndex        =   32
               Top             =   180
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   90
               TabIndex        =   33
               Top             =   510
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   285
               Index           =   32
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   180
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   285
               Index           =   31
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   510
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   315
               Index           =   30
               Left            =   6930
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   210
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   315
               Index           =   29
               Left            =   6930
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   540
               Width           =   975
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   210
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   33
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   510
               Width           =   1485
            End
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11880
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   930
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ŒÌ«—« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   90
            Visible         =   0   'False
            Width           =   3735
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "œ›⁄Â „ﬁœ„Â"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "FIFO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   480
               Width           =   1335
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   " ÕœÌœ ›Ê« Ì—"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   720
               Width           =   2055
            End
            Begin VB.OptionButton Option6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   " ÕœÌœ „” Œ·’« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   1080
               Value           =   -1  'True
               Width           =   2055
            End
            Begin ALLButtonS.ALLButton ALLButton3 
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   " ÕœÌœ"
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmADVPaymentsAlloc.frx":25B4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton ALLButton4 
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   1320
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   " ÕœÌœ"
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmADVPaymentsAlloc.frx":25D0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
         End
         Begin VB.TextBox txtAdv_payment_value 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   13320
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2415
            Width           =   2685
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   690
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   210
            Width           =   1395
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì Õ«·… «·„‘«—Ì⁄"
            Enabled         =   0   'False
            Height          =   615
            Left            =   14280
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1650
            Width           =   4215
            Begin VB.OptionButton Option4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄„Ì· ‰Â«∆Ì"
               Height          =   195
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬁ«Ê· »«ÿ‰"
               Height          =   195
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.TextBox txtperson 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14280
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   4170
            Width           =   2685
         End
         Begin vbalIml6.vbalImageList vbalImageList1 
            Left            =   10680
            Top             =   450
            _ExtentX        =   953
            _ExtentY        =   953
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   375
            Left            =   12120
            TabIndex        =   29
            Top             =   2610
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«ŸÂ«— «·«ﬁ”«ÿ"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmADVPaymentsAlloc.frx":25EC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DcboRevenuesTypes 
            Height          =   315
            Left            =   6360
            TabIndex        =   70
            Top             =   1350
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   6360
            TabIndex        =   76
            Top             =   1350
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   540
            Index           =   2
            Left            =   120
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   7050
            Width           =   7995
            _cx             =   14102
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
            Appearance      =   0
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
         Begin ImpulseAniLabel.ISAniLabel LblLink 
            Height          =   315
            Left            =   210
            TabIndex        =   86
            Top             =   1320
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   556
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
            MouseIcon       =   "FrmADVPaymentsAlloc.frx":2608
            BackColor       =   14871017
            Alignment       =   1
            Caption         =   ""
            ColorHover      =   16711680
            RightToLeft     =   -1  'True
            ImageCount      =   0
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   375
            Left            =   13200
            TabIndex        =   87
            Top             =   2850
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«ŸÂ«— ”‰œ «·„œÌÊ‰Ì…"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmADVPaymentsAlloc.frx":276A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DCPROJECT 
            Height          =   315
            Left            =   12840
            TabIndex        =   88
            Top             =   4650
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmADVPaymentsAlloc.frx":2786
            Height          =   315
            Left            =   13320
            TabIndex        =   89
            Top             =   2850
            Width           =   2655
            _ExtentX        =   4683
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
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmADVPaymentsAlloc.frx":279B
            Height          =   315
            Left            =   6720
            TabIndex        =   125
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
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
         Begin MSDataListLib.DataCombo dcEmployee 
            Height          =   315
            Left            =   6600
            TabIndex        =   138
            Top             =   1350
            Visible         =   0   'False
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCAccounts 
            Height          =   315
            Left            =   6360
            TabIndex        =   140
            Top             =   1320
            Visible         =   0   'False
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcEmp 
            Bindings        =   "FrmADVPaymentsAlloc.frx":27B0
            Height          =   315
            Left            =   13560
            TabIndex        =   144
            Top             =   2400
            Width           =   2595
            _ExtentX        =   4577
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
         Begin MSDataListLib.DataCombo DCCar 
            Height          =   315
            Left            =   13560
            TabIndex        =   149
            Top             =   2760
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDriver 
            Height          =   315
            Left            =   13560
            TabIndex        =   150
            Top             =   3120
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   2115
            Left            =   3480
            TabIndex        =   153
            Top             =   1920
            Width           =   8355
            _cx             =   14737
            _cy             =   3731
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmADVPaymentsAlloc.frx":27C5
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2115
            Left            =   6360
            TabIndex        =   154
            Top             =   4320
            Width           =   5475
            _cx             =   9657
            _cy             =   3731
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmADVPaymentsAlloc.frx":2A09
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·”‰œ« "
            Height          =   315
            Index           =   52
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   4080
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·„” Œ·’« "
            Height          =   315
            Index           =   51
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·„⁄œÂ/«·”Ì«—…"
            Height          =   285
            Index           =   50
            Left            =   14880
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·”«∆ﬁ"
            Height          =   285
            Index           =   49
            Left            =   14280
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   285
            Index           =   48
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   240
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰œÊ»"
            Height          =   255
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
            Height          =   315
            Index           =   47
            Left            =   16320
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   11160
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   600
            Width           =   735
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   6810
            Width           =   825
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   300
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   6810
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   37
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   6810
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   7
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   6810
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· Œ’Ì’"
            Height          =   285
            Index           =   6
            Left            =   10530
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   8370
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   285
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
            Height          =   285
            Index           =   2
            Left            =   13410
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   2430
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· √Ê «·„Ê—œ"
            Height          =   315
            Index           =   3
            Left            =   10650
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   1290
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·”‰œ"
            Height          =   285
            Index           =   4
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "–·ﬂ „ﬁ«»·"
            Height          =   285
            Index           =   5
            Left            =   13770
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   4650
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·Õ«·Ï:"
            Height          =   315
            Index           =   13
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
            Height          =   315
            Index           =   14
            Left            =   13410
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   2760
            Width           =   1275
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
            Height          =   435
            Index           =   18
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   1680
            Width           =   4065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   285
            Index           =   34
            Left            =   11760
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   4890
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblsqlstring 
            Alignment       =   1  'Right Justify
            Height          =   855
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   2250
            Width           =   2895
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ›⁄Â „ﬁœ„Â"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   35
            Left            =   13170
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   2430
            Width           =   1035
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
            Height          =   255
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   2850
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·„ﬂ—„"
            Height          =   285
            Index           =   36
            Left            =   13800
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   4170
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7335
         Index           =   0
         Left            =   12855
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   12120
         _cx             =   21378
         _cy             =   12938
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
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   5235
            Left            =   3600
            TabIndex        =   113
            Top             =   1080
            Width           =   8355
            _cx             =   14737
            _cy             =   9234
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmADVPaymentsAlloc.frx":2C4D
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   0
            TabIndex        =   124
            Tag             =   "Delete Row"
            Top             =   6240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› „” Œ·’"
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
            MICON           =   "FrmADVPaymentsAlloc.frx":2EB1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   495
            Left            =   12480
            Top             =   360
            Width           =   8175
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "«·„„” Œ·’«  «· Ì  „ ”œ«œÂ« ··„‘—Ê⁄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   42
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   41
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   38
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "ﬁ„ » ÕœÌœ «·„” Œ·’«   «·„—«œ ”œ«œÂ« ··„‘—Ê⁄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   0
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   4335
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   495
            Left            =   3720
            Top             =   360
            Width           =   8175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   840
            Width           =   7575
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   9300
      TabIndex        =   114
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   115
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   7515
      TabIndex        =   116
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   6615
      TabIndex        =   117
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   5730
      TabIndex        =   118
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
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
      Height          =   375
      Index           =   6
      Left            =   2160
      TabIndex        =   119
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   3045
      TabIndex        =   120
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   4830
      TabIndex        =   121
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   3945
      TabIndex        =   122
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   9
      Left            =   8160
      TabIndex        =   123
      Top             =   9000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·ﬁÌœ"
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   9480
      TabIndex        =   136
      Top             =   9000
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
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
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   46
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   142
      Top             =   8760
      Width           =   7155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   8
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   137
      Top             =   9000
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   495
      Left            =   0
      Top             =   5760
      Width           =   8175
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "ﬁ„ » ÕœÌœ «·„” Œ·’«   «·„—«œ ”œ«œÂ« ··„‘—Ê⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Index           =   40
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   110
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Index           =   39
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   109
      Top             =   5760
      Width           =   3735
   End
End
Attribute VB_Name = "FrmADVPaymentsAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim Line1 As Double
Dim Line2 As Double
Dim Line3 As Double
Dim Line4 As Double

Dim departement_name As Integer
Dim numbering_type As Integer
Dim Balance As String
Dim balanceString As String

Private Sub ALLButton1_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
        'INSTALLMENT_DATA1.show
        'INSTALLMENT_DATA1.Adodc1.CommandType = adCmdText
        'INSTALLMENT_DATA1.Adodc1.RecordSource = "select *  FROM INSTALLMENT_DETAILS where payed=0 and cust_id =" & Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA1.Adodc1.Refresh '
 
        'INSTALLMENT_DATA1.id.text = Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA1.lblcustid = Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA1.TxtName.text = Me.DBCboClientName.text
    End If

End Sub

Private Sub ALLButton2_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
        'sanad_dean.show
        'sanad_dean.LblID = DBCboClientName.BoundText
        'sanad_dean.LblName = DBCboClientName.text
        ''sanad_dean.lblaccountcode.Caption = txtaccount.text
        'sanad_dean.Adodc1.CommandType = adCmdText
        'sanad_dean.Adodc1.RecordSource = "select*  FROM sanad_dean where cust_id=" & DBCboClientName.BoundText
        'sanad_dean.Adodc1.Refresh
        'sanad_dean.ALLButton1.Visible = False
        'sanad_dean.ALLButton1.Visible = False
'
'        sanad_dean.Adodc2.CommandType = adCmdText
'        sanad_dean.Adodc2.RecordSource = "select *  FROM member_child where cust_id=" & DBCboClientName.BoundText
'        sanad_dean.Adodc2.Refresh
    End If

End Sub

Private Sub ALLButton3_Click()
    lblsqlstring.Caption = ""
    FrmPaymentTime1.show
    FrmPaymentTime1.lblcusid = DBCboClientName.BoundText
    FrmPaymentTime1.LblValue = val(XPTxtVal.text)
End Sub

Public Sub FillGridWithData(project_no As Integer, _
                            Optional TxtNoteSerial As String)

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim result As Double
    Dim resultpercentage As Double
    Dim sql As String

    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Rows = 1

    If DCboCashType.ListIndex <> 5 Then Exit Sub
 
    lbl(38).Caption = DBCboClientName.text
    lbl(41).Caption = DBCboClientName.text
    sql = "SELECT  * FROM     project_billl     where project_no = " & project_no
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If

    i = 0

    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable
  
        rs.MoveFirst

        For X = 1 To rs.RecordCount
       
            ActualTotal = getBillPayedToproject(val(rs.Fields("id").value))
            result = val(rs.Fields("total").value) - ActualTotal
            resultpercentage = Round((ActualTotal / val(rs.Fields("total").value)) * 100, 2)
 
            If val(rs.Fields("total").value) > ActualTotal Then
                i = i + 1
                .Rows = .Rows + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                '                             .TextMatrix(I, .ColIndex("bill_id")) = IIf(IsNull(rs.Fields("bill_id").value), _
                                              "", rs.Fields("bill_id").value)
            
                .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("bill_date").value), "", rs.Fields("bill_date").value)
                .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), "", rs.Fields("project_no").value)
                .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), "", rs.Fields("project_name").value)
            
                .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs.Fields("End_user_name").value), "", rs.Fields("End_user_name").value)
            
                .TextMatrix(i, .ColIndex("Sub_user_name")) = IIf(IsNull(rs.Fields("Sub_user_name").value), "", rs.Fields("Sub_user_name").value)
            
                .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
 
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
            
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = result

            End If

            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

    If TxtNoteSerial = "" Then

        Exit Sub
    End If

    sql = "SELECT  * FROM     ProjectBillBuy     where TxtNoteSerial ='" & TxtNoteSerial & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If
 
    With Me.Grid1
        .Rows = 1
        .Rows = .Rows + rs.RecordCount
        .Clear flexClearScrollable
  
        rs.MoveFirst

        For i = 1 To .Rows - 1
 
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
            .TextMatrix(i, .ColIndex("bill_id")) = IIf(IsNull(rs.Fields("bill_id").value), "", rs.Fields("bill_id").value)
            
            .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("RecordDate").value), "", rs.Fields("RecordDate").value)
            '                                           .TextMatrix(I, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), _
                                                        "", rs.Fields("project_no").value)
            '                         .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), _
                                      "", rs.Fields("project_name").value)
            
            .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
 
            .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
            
            .TextMatrix(i, .ColIndex("ActualTotal")) = IIf(IsNull(rs.Fields("value").value), "", rs.Fields("value").value)
            result = val(.TextMatrix(i, .ColIndex("total"))) - val(rs.Fields("value").value)
            resultpercentage = val(rs.Fields("value").value) / val(.TextMatrix(i, .ColIndex("total"))) * 100
            .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
            .TextMatrix(i, .ColIndex("Result")) = result
      
            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub ALLButton4_Click()

    If DCboCashType.ListIndex <> 5 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–… «·⁄„·Ì… „ «Õ… „⁄ ›Ê« Ì— «·„‘«—Ì⁄ ›ﬁÿ", vbInformation
        Else
            MsgBox "This Process For Project Bill Only", vbInformation
    
        End If

        DCboCashType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If val(DBCboClientName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ — „‘—Ê⁄ «Ê·«", vbInformation
        Else
            MsgBox "select Project Firstly, vbInformation"
    
        End If

        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Exit Sub

    End If
 
    FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text

End Sub

Private Sub CboPayMentType_Change()
DBCboClientName_Change

    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        DcChequeBox.text = ""
        TXTBankName.text = ""
    End If

    DcChequeBox.Enabled = False

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ"
        lbl(17).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(16).Caption = "Cheque No"
        lbl(17).Caption = "Due Date"
    End If
    
    If Me.CboPaymentType.ListIndex = 0 Then
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Frame3.Enabled = False
    ElseIf Me.CboPaymentType.ListIndex = 1 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            DcChequeBox.Enabled = True
        Else
            TXTBankName.Visible = False
        End If

        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame3.Enabled = False
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
 
        TXTBankName.Visible = False
 
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TXTBankName.Visible = False
        Frame3.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(16).Caption = "—ﬁ„ «·ÕÊ«·Â"
            lbl(17).Caption = " «—ÌŒÂ«"
    
        Else
            lbl(16).Caption = "Transfer No"
            lbl(17).Caption = "Date"
        End If
 
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
 
        TXTBankName.Visible = False
 
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TXTBankName.Visible = False
        Frame3.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ"
            lbl(17).Caption = " «—ÌŒÂ"
    
        Else
            lbl(16).Caption = "Chequ No"
            lbl(17).Caption = "Date"
        End If
 
    Else

        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub ChkTrans_Click()
    Me.lbl(10).Enabled = ChkTrans.value
    Me.lbl(12).Enabled = ChkTrans.value
    Me.CboTrans.Enabled = ChkTrans.value
    Me.TxtTransID.Enabled = ChkTrans.value
    Me.TxtTransSerial.Enabled = ChkTrans.value
    Me.CmdSearchTrans.Enabled = ChkTrans.value
    Me.CmdOpenTrans.Enabled = ChkTrans.value
End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As Integer
    auto_sanad_no = ""
    departement_name = 1
 
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=2"
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
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
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
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

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
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

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
                    auto_sanad_no = Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & start_at

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
                    NO = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Private Sub Cmd_Click(Index As Integer)

    Dim cNoteReport As ClsNotesReports
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = False
        ' XPDtbBill.Enabled = False
    End If

    Select Case Index

        Case 0

            If SystemOptions.SysRegisterState = DemoRun Then
                If Not rs Is Nothing Then
                    If Not (rs.BOF Or rs.EOF) Then
                        If rs.RecordCount >= 25 Then
                            Msg = "›Ï «·‰”Œ… «· Ã—Ì»Ì… ·«Ì„ﬂ‰  ”ÃÌ· «ﬂÀ— „‰ 25 ⁄„·Ì… ﬁ»÷ «Ê œ›⁄"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
        
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
          
            Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 1
            TxtModFlg.text = "N"
            '       XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
            Text1.text = setfoxy
            Option4.value = True
            Me.dcBranch.BoundText = Current_branch
            Txt_DateHigri.value = ToHijriDate(Date)

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If
              
            If SystemOptions.ChequeBox = True And CboPaymentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
    
            End If
    
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            CuurentLogdata

        Case 2

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText
 
            If Option2.value = True And lblsqlstring.Caption = "" Then MsgBox "·«»œ „‰  ÕœÌœ ›Ê« Ì—": Exit Sub
 
            'TxtNoteSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)
            SaveData
        
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.ChequeBox = True And CboPaymentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 4
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            ' If Val(Me.XPTxtID.text) <> 0 Then
            '     Set cNoteReport = New ClsNotesReports
            '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
            '     Set cNoteReport = Nothing
            ' End If
            If TxtNoteSerial <> "" Then
                print_report TxtNoteSerial, Me.TxtNoteSerial1.text, TXTBankName.text, CboPaymentType.text, DcboBox.text, TxtCustCode.text
            End If

        Case 8

            'ViewDataList
        Case 9
    
            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.text, , 200
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_report(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From payment_voucher  where noteserial='" & NoteSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
   
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    '
    xReport.ParameterFields(6).AddCurrentValue NoteSerial1

    xReport.ParameterFields(7).AddCurrentValue BankName
    xReport.ParameterFields(8).AddCurrentValue PaymentType
    xReport.ParameterFields(9).AddCurrentValue Box
    xReport.ParameterFields(10).AddCurrentValue Custcode
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

Private Sub ViewDataList()
    Dim FrmView As FrmViewList
    Dim Fg As VSFlex8UCtl.vsFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set Fg = FrmView.vsfGroup1.vsFlexGrid

    With Fg
        .Cols = 18
        .RowHeightMin = 320
        .ExplorerBar = flexExSortShowAndMove
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = " ‰Ê⁄ «·„ﬁ»Ê÷« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„ﬁ»Ê÷« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
        StrSQL = StrSQL + " Order By NoteID"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        'Rs.Close
        'Set Rs = Nothing
        .AutoSize 0, .Cols - 1, False
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ Õ–› «·„Õœœ ", vbCritical + vbYesNo)
    End If

    Dim sql As String

    If X = vbNo Then Exit Sub
    sql = "delete from ProjectBillBuy where id=" & val(Grid1.TextMatrix(Grid1.Row, Grid1.ColIndex("id")))
    Cn.Execute sql

    If Grid1.Rows > 1 Then
        If Grid1.Rows = 2 Then
            Me.Grid1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid1.Rows > 1 Then
                If Me.Grid1.Row <> Me.Grid1.FixedRows - 1 Then
                    Me.Grid1.RemoveItem (Me.Grid1.Row)
                End If
            End If
        End If
    End If

    If DCboCashType.ListIndex = 5 Then
        FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
    End If
  
End Sub

Private Sub CmdSearchTrans_Click()
    Dim Msg As String

    If Me.CboTrans.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·Õ—ﬂ… «·„—«œ «·»ÕÀ ⁄‰Â«..."
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboTrans.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.CboTrans.ListIndex = 0 Then
        ' ›« Ê—… „»Ì⁄« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = InvoiceTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPaymentType.ListIndex = 1
        FrmBuySearch.CboPaymentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 1 Then
        '›« Ê—… „— Ã⁄ „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = ReturnTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPaymentType.ListIndex = 1
        FrmBuySearch.CboPaymentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal
    ElseIf Me.CboTrans.ListIndex = 2 Then
        '›« Ê—… ’Ì«‰…
        Load FrmMaintanenceSearch
        Set FrmMaintanenceSearch.ExtraRetrunObject = Me.TxtTransID
        FrmMaintanenceSearch.CboPaymentType.ListIndex = 1
        FrmMaintanenceSearch.SearchType = 4
        FrmMaintanenceSearch.CboPaymentType.Enabled = False
        FrmMaintanenceSearch.show vbModal
    End If

End Sub

Private Sub Command1_Click()
 
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub
 Private Sub DBCboClientName_Change()
    TxtCustCode.text = ""

    If Me.DCboCashType.ListIndex = 5 And Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
       ' Option4.value = True
    End If

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

    TxtCustCode.text = fullcode

    If DBCboClientName.BoundText = "" Then Exit Sub
 
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        'Dim fullcode As String
      If Me.DCboCashType.ListIndex = 0 Then
        fullcode = ""
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode
'        TxtCustCode.text = fullcode

        DcEmp.BoundText = DefaultSalesPersonId
        ElseIf Me.DCboCashType.ListIndex = 5 Then
        
            fullcode = ""
        GetProjectsDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode
       TxtCustCode.text = fullcode

        DcEmp.BoundText = DefaultSalesPersonId
        
        
        End If
        
        
        
        
        
        If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPaymentType.ListIndex = 0 Then '‰ﬁœÌ
                                               If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                               
                            ElseIf CboPaymentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                    Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
                                             End If
                                     
                             ElseIf CboPaymentType.ListIndex = 2 Then 'ÕÊ«·… '
                                                If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                              ElseIf CboPaymentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                                    If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                              End If
                             
'
        Else '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
                Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))

        End If
        

        If DCboCashType.ListIndex = 5 Then 'Õ«·… «·„‘«—Ì⁄
                                        
       If Option4.value = True Then ' ⁄„Ì· ‰Â«∆Ì
                                        
        If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPaymentType.ListIndex = 0 Then '‰ﬁœÌ
                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                               
                            ElseIf CboPaymentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1") '  Õ  «· Õ’Ì·
                                                                      End If
                                     
                             ElseIf CboPaymentType.ListIndex = 2 Then 'ÕÊ«·… '
                                               If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                              ElseIf CboPaymentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                      If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                  Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                              End If
                             
'
        Else '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
                Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì

        End If
                                                
                                                
                                                
          Else '⁄„Ì· «·»«ÿ‰55555555555555555555555555555555555555555
          
                  If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPaymentType.ListIndex = 0 Then '‰ﬁœÌ
                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                               
                            ElseIf CboPaymentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1", 1) '  Õ  «· Õ’Ì·
                                                                      End If
                                     
                             ElseIf CboPaymentType.ListIndex = 2 Then 'ÕÊ«·… '
                                               If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                              ElseIf CboPaymentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                      If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                  Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                              End If
                             
'
        Else '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
                Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì

        End If
        
          
          
          
          End If
                                        
                          '(((((((((((((((((((((((((((((((((((((((
                                        
            
                        '(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
       End If
    End If

End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    'WriteCustomerBalPublic
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If DCboCashType.ListIndex = 0 Then
        If KeyCode = vbKeyF3 Then
         FrmCustemerSearch.SearchType = 3
            FrmCustemerSearch.show vbModal
           
        End If

    ElseIf DCboCashType.ListIndex = 1 Then

        If KeyCode = vbKeyF3 Then
          FrmCompanySearch.lblSearchtype.Caption = 2
            FrmCompanySearch.show vbModal
          
        End If

   ElseIf DCboCashType.ListIndex = 5 Then

        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 1
             FrmProjectSearch.show vbModal
           
        End If
  
    End If

End Sub

Private Sub DCAccounts_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        Me.DcboCreditSide.BoundText = DCAccounts.BoundText
  
    End If

End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 190
            
    End If

End Sub

Private Sub DcboBankName_Click(Area As Integer)

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        'Me.DcboDebitSide.BoundText =   "a1a2a4"
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.ChequeBox = True Then
            Me.DcboDebitSide.BoundText = ""
        Else

            If SystemOptions.banks_Accounts3 = True Then
                Me.DcboDebitSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code1")
            Else
                Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                     
            End If
        End If

        If CboPaymentType.ListIndex = 2 Or CboPaymentType.ListIndex = 3 Then
                     
            Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

    End If

End Sub

Private Sub DcboBox_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DCboCashType_Change()
    On Error GoTo ErrTrap
    Frame2.Enabled = False
    Dim StrSQL As String
    Dim intDef As Integer

    Select Case DCboCashType.ListIndex

        Case 0
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
        
            DcEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 1
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DcEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„Ê—œ"
            Else
                Me.lbl(3).Caption = "Vendor Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 2
            Dcombos.GetPersons Me.DBCboClientName
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DcEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = False
            Fra(0).Visible = False

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(3).Caption = "name"
            Else
                Me.lbl(3).Caption = "„ﬁ«Ê· «·»«ÿ‰"
            End If
                
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 3
            '≈Ì—«œ«  ≈Œ—Ï
            Me.DBCboClientName.Visible = False
            Me.DcboRevenuesTypes.Visible = True
            Me.ChkTrans.Visible = False
            DBCboClientName.Visible = False
            DcEmployee.Visible = False
            DCAccounts.Visible = False
            Fra(0).Visible = False
        
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "‰Ê⁄ «·«Ì—«œ"
            Else
                Me.lbl(3).Caption = "RVN Type"
            End If
                
            Me.lbl(13).Visible = False
            Me.LblLink.Visible = False
        
        Case 4
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DcEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
        
        Case 5
            Dim My_SQL As String
            My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null) order by Project_name" '
            fill_combo Me.DBCboClientName, My_SQL
         
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DcEmployee.Visible = False
            DCAccounts.Visible = False

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„‘—Ê⁄"
            Else
                Me.lbl(3).Caption = "project Name"
            End If
        
            Frame2.Enabled = True
        
        Case 6
            Dcombos.GetEmployees Me.DcEmployee
            Me.DcEmployee.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DBCboClientName.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True

            '   Fra(0).Visible = True
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„ÊŸ›"
            Else
                Me.lbl(3).Caption = "Employee  Name"
            End If

        Case 7
            Dcombos.GetAccountingCodes Me.DCAccounts, True
            DCAccounts.Visible = True
            Me.DcEmployee.Visible = False
            Me.DcboRevenuesTypes.Visible = False
            DBCboClientName.Visible = False
        
            ChkTrans.Visible = True

            '   Fra(0).Visible = True
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·Õ”«»"
            Else
                Me.lbl(3).Caption = "Accounts Nam  "
            End If
        
            '  Me.lbl(13).Visible = True
            '      Me.LblLink.Visible = True
    End Select

    cSearchDcbo.Refresh
    Exit Sub
ErrTrap:
End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub DcboCreditSide_Change()

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
End Sub

Private Sub DcboRevenuesTypes_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", val(Me.DcboRevenuesTypes.BoundText))
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub DcChequeBox_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxID", val(Me.DcChequeBox.BoundText), "Account_Code1")
    End If

End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 6
    End If

End Sub

Private Sub dcEmployee_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        Me.DcboCreditSide.BoundText = get_EMPLOYEE_Account(val(DcEmployee.BoundText), "Account_Code")
        TxtCustCode.text = val(DcEmployee.BoundText)
    End If

End Sub

Private Sub DCCar_Change()

    GetDriverInformation (val(DCCar.BoundText))

End Sub

Private Sub DCCar_Click(Area As Integer)
    GetDriverInformation (val(DCCar.BoundText))

End Sub

Function GetDriverInformation(id As Integer)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Dim sql As String
        Dim rs As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TblCarsData"
        sql = sql & " Where (id = " & id & ") "

        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            DCDriver.BoundText = IIf(IsNull(rs("Emp_id").value), 0, rs("Emp_id").value)
                  
        Else
            DcEmp = 0
               
        End If

    End If

End Function

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If

    If mdifrmmain.TransporterMain.Visible = False Then
        lbl(49).Visible = False
        lbl(50).Visible = False
        DCCar.Visible = False
        DCDriver.Visible = False

    End If

    ScreenNameArabic = "«·„ﬁ»Ê÷« "
    ScreenNameEnglish = "Cashing"
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1", 4
 
    Dim StrSQL As String
    Dim Msg As String
    Set Dcombos = New ClsDataCombos
    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL

    Dcombos.GetSalesRepData Me.DcEmp
    Dcombos.GetCars Me.DCCar
    Dcombos.GetEmployees Me.DCDriver, , True

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(8).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    'Resize_Form Me
    AddTip
    DCboCashType.AddItem "„‰ ⁄„Ì·"
    DCboCashType.AddItem "„‰ „Ê—œ"
    DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
    DCboCashType.AddItem "≈Ì—«œ«  ≈Œ—Ï"
    DCboCashType.AddItem "„œ›Ê⁄«  „ﬁœ„Â"
    DCboCashType.AddItem "„‘—Ê⁄"
    DCboCashType.AddItem "„‰ „ÊŸ›"
    DCboCashType.AddItem "„‰ Õ”«»"

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "  ‘Ìﬂ „Õ’· "
    
    End With

    Dcombos.GetUsers Me.DCboUserName

    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetChequeBox Me.DcChequeBox

    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, False
    Dcombos.GetRevenuesTypes Me.DcboRevenuesTypes
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = Me.DBCboClientName

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.dcBranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = False
    End If

    Set rs = New ADODB.Recordset
    'StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
    StrSQL = "select * From Notes where NoteType=4    "

    If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
            
    StrSQL = StrSQL & "and  displayed is null Order By NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    SetDtpickerDate Me.XPDtbTrans
    SetDtpickerDate Me.DtpChequeDueDate

    With Me.CboTrans
        .Clear
        .AddItem "›« Ê—… „»Ì⁄« "
        .AddItem "„— Ã⁄ „‘ —Ì« "
        .AddItem " ”·Ì„ ’Ì«‰… ·⁄„Ì·"
        .AddItem "Œœ„« "
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Msg = "„·ÕÊŸ…:-"
    Msg = Msg & Chr(13) & "≈–« ﬂ«‰  Â–Â «·„ﬁ»Ê÷«   Õ’Ì· ·›« Ê—… „⁄Ì‰…"
    Msg = Msg & "›ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ » ÕœÌœ Â–Â «·›« Ê—… "
    Msg = Msg & "Õ Ï Ì „ —»ÿ ⁄„·Ì… «· Õ’Ì· Â–Â „⁄ «·›« Ê—…"
    Me.lbl(11).Caption = Msg
    SetDtpickerDate Me.XPDtbTrans
    ChkTrans.value = Unchecked
    ChkTrans_Click
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    WriteInfo
      
    Dim My_SQL As String

    'My_SQL = "  select expanses_account,account_name from projects  where not (account_no is null)"
    My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
    fill_combo DcProject, My_SQL

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, 4

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If Index = 18 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(18).ToolTipText = "ﬁÌ„… „»·€ «·„ﬁ»Ê÷« :" & lbl(18).Caption
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(18).ToolTipText = "Notes Recivable Value:" & lbl(18).Caption
        End If
    End If

End Sub

Private Sub LblLink_Click()
 
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.text, FirstPeriod, Date

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
 
    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·ÿ—› «·œ«∆‰:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Credit Balance:" & WriteNo(Balance, 0, True)
    End If
 
End Sub

Private Sub Option1_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If
DBCboClientName_Change
End Sub

Private Sub Option2_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If
DBCboClientName_Change
End Sub

Private Sub Option3_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If
DBCboClientName_Change
End Sub

Private Sub Option4_Click()

    If DCboCashType.ListIndex <> 5 Then Exit Sub
 DBCboClientName_Change

End Sub

Private Sub Option5_Click()

    If DCboCashType.ListIndex <> 5 Then Exit Sub
 DBCboClientName_Change

End Sub

Private Sub Option6_Click()

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If

End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtCustCode.text, DCboCashType.ListIndex + 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
    
            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Receipts"
            Else
                '        Me.Caption = "«·„ﬁ»Ê÷« "
            End If

            Ele(0).Enabled = False
            Grid.Enabled = False
            Grid1.Enabled = False
            CmdRemove.Enabled = False
            ' Frame1.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPTxtVal.locked = True
            XPDtbTrans.Enabled = False
            XPMTxtRemarks.locked = True
            DBCboClientName.locked = True
            DCboCashType.locked = True
            Me.CboPaymentType.locked = True
            Me.DcboBox.locked = True
            Me.DcboBankName.locked = True
            Me.TxtChequeNumber.locked = True
            Me.DtpChequeDueDate.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            Fra(0).Enabled = False
            ChkTrans.Enabled = False

        Case "N"
            '        Me.Caption = "«·„ﬁ»Ê÷« ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Grid.Enabled = True
            Grid1.Enabled = False
            CmdRemove.Enabled = False
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            '    Me.XPBtnMove(0).Enabled = False
            '    Me.XPBtnMove(1).Enabled = False
            '    Me.XPBtnMove(2).Enabled = False
            '    Me.XPBtnMove(3).Enabled = False
            XPDtbTrans.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DCboCashType.ListIndex = 0
        
            Me.CboPaymentType.locked = False
            Me.DcboBox.locked = False
            Me.DcboBankName.locked = False
            Me.TxtChequeNumber.locked = False
            Me.DtpChequeDueDate.Enabled = True
        
            Fra(0).Enabled = True
            ChkTrans.Enabled = True

        Case "E"
            '        Me.Caption = "«·„ﬁ»Ê÷« (  ⁄œÌ· )"

            Grid.Enabled = True
            Grid1.Enabled = True
            CmdRemove.Enabled = True
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            XPDtbTrans.Enabled = True
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DCboCashType.locked = False
            Fra(0).Enabled = True
            ChkTrans.Enabled = True
            Me.CboPaymentType.locked = False
            Me.DcboBox.locked = False
            Me.DcboBankName.locked = False
            Me.TxtChequeNumber.locked = False
            Me.DtpChequeDueDate.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.TxtTransID.text <> "" Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                Me.TxtTransSerial.text = GetTransIDSerial(1, val(Me.TxtTransID.text))
            Else
                Me.TxtTransSerial.text = Me.TxtTransID.text
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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

Public Sub Retrive(Optional Lngid As Long = 0)
          
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If

    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.DcEmp.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    TxtManulaNO.text = IIf(IsNull(rs("ManulaNO").value), "", rs("ManulaNO").value)

    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(46).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    txtperson.text = IIf(IsNull(rs("person").value), "", rs("person").value)
Option1.value = False
Option2.value = False
Option3.value = False
If IsNull(rs("NCashingType").value) Then

Else
        If rs("NCashingType").value = 1 Then
               Option1.value = True
        ElseIf rs("NCashingType").value = 2 Then
              Option2.value = True
        ElseIf rs("NCashingType").value = 3 Then
             Option3.value = True
        End If
End If



    If Option1.value = True Then
       rs("NCashingType").value = 1
   ElseIf Option2.value = True Then
        rs("NCashingType").value = 2
   ElseIf Option3.value = True Then
        rs("NCashingType").value = 3
    Else
    
         rs("NCashingType").value = 0
   End If
   
   
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
    TXTBankName.text = IIf(IsNull(rs("BankName").value), "", Trim(rs("BankName").value))
 
    txtAdv_payment_value.text = IIf(IsNull(rs("Adv_payment_value").value), "", Trim(rs("Adv_payment_value").value))

    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    'dcproject.BoundText = IIf(IsNull(Rs("Remark").value), "", Trim(Rs("Remark").value))

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(XPDtbTrans.value), rs("NoteDateH").value)
    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)

    Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)
    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)

    '-----------------------------------------------------------------------------
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    
        'project_Expensen_account
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        Me.DcChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        Me.DcChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPaymentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
    
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
        If SystemOptions.ChequeBox = True Then
            Me.DcChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            Me.DcChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

    ElseIf rs("NoteCashingType").value = 2 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            'Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            TXTBankName.Visible = False
            Me.DcChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

        Me.CboPaymentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DcChequeBox.BoundText = ""

    ElseIf rs("NoteCashingType").value = 3 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            'Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            TXTBankName.Visible = False
            Me.DcChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

        Me.CboPaymentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DcChequeBox.BoundText = ""
    
    End If

    CboPayMentType_Change

    '-----------------------------------------------------------------------------
    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkTrans.value = vbChecked
        'Me.ChkTrans.Enabled = True
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Me.TxtTransID.text = RsTemp("Transaction_ID").value
            Me.TxtTransSerial.text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)

            If Not (IsNull(RsTemp("Transaction_Type").value)) Then
                If RsTemp("Transaction_Type").value = 5 Then
                    Me.CboTrans.ListIndex = 1
                ElseIf RsTemp("Transaction_Type").value = 2 Then
                    Me.CboTrans.ListIndex = 0
                End If
            End If
        End If

    ElseIf Not IsNull(rs("MaintananceID").value) Then
        Me.ChkTrans.value = vbChecked
        Me.CboTrans.ListIndex = 2
        Me.TxtTransID.text = rs("MaintananceID").value
        Me.TxtTransSerial.text = rs("MaintananceID").value
    ElseIf Not IsNull(rs("RevenuesID").value) Then
        Me.DcboRevenuesTypes.BoundText = rs("RevenuesID").value
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
    Else
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
    End If

    If DCboCashType.ListIndex = 5 Then
        Dim My_SQL As String
        My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
        fill_combo Me.DBCboClientName, My_SQL
      
        DBCboClientName.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
        Dim cus_or_sub As Integer
        cus_or_sub = IIf(IsNull(rs("cus_or_sub").value), 0, rs("cus_or_sub").value)

        If cus_or_sub = 0 Then
            Option4.value = True
        Else
            Option5.value = True
        End If

    End If

    If DCboCashType.ListIndex = 6 Then
        DcEmployee.BoundText = IIf(IsNull(rs("EmployeeID").value), "", rs("EmployeeID").value)
    End If
  
    If DCboCashType.ListIndex = 7 Then
        Me.DCAccounts.BoundText = IIf(IsNull(rs("AccountsCode").value), "", rs("AccountsCode").value)
    End If

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(33).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To 2 ' RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i

        End If
    End If

    '-----------------------------------------------------------------------------
    ChkTrans_Click
    '⁄—÷ «·„” Œ·’« 
    'If DCboCashType.ListIndex = 5 Then
    FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
    '  End If
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim LngDevID As Long
    Dim RsDev As ADODB.Recordset

    Dim BeginTrans As Boolean
     On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If DCboCashType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„ﬁ»Ê÷«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboCashType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Me.DCboCashType.ListIndex = 3 Then
            If val(Me.DcboRevenuesTypes.BoundText) = 0 Then
                Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·≈Ì—«œ«  «·√Œ—Ï...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                If Me.DcboRevenuesTypes.Visible = True Then
                    DcboRevenuesTypes.SetFocus
                    SendKeys "{F4}"
                End If

                Exit Sub
            End If
        End If

        If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Then
            If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DBCboClientName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If
    
        If Me.DCboCashType.ListIndex = 5 Then
            If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ ««·„‘—Ê⁄"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DBCboClientName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If
    
        If Me.DCboCashType.ListIndex = 6 Then
            If DcEmployee.BoundText = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·„ÊŸ›"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcEmployee.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If
    
        If Me.DCboCashType.ListIndex = 7 Then
            If Me.DCAccounts.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·Õ”«»"
                Else
                    Msg = "Select Account Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DCAccounts.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If
    
        If XPTxtVal.text = "" Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„ﬁ»Ê÷«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '        XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.text) Then
            Msg = "ﬁÌ„… «·„ﬁ»Ê÷«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = -1 Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                CboTrans.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim(Me.TxtTransSerial.text) = "" Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Sub
            Else

                If Me.CboTrans.ListIndex = 0 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 1 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 5)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 2 Then

                    If CheckDebitMaintaince(val(Me.TxtTransSerial.text)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 3 Then
                    Msg = "⁄›Ê« .. Ã«—Ï  ÿÊÌ— «·»—‰«„Ã .. ·⁄„· «·„ﬁ»Ê÷«  „‰ «·Œœ„« "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If
        End If

        If Me.CboPaymentType.ListIndex = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPaymentType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Me.CboPaymentType.ListIndex = 0 Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPaymentType.ListIndex = 1 Then
      
            '  If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '      Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '      MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '      DtpChequeDueDate.SetFocus
            '      SendKeys "{F4}"
            '      Exit Sub
            '  End If
            If SystemOptions.ChequeBox = True Then
         
                If DcChequeBox.BoundText = "" Then
                           
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Õœœ Õ«›Ÿ… «·‘Ìﬂ«  ...!!"
                    Else
                        Msg = "Select Cheque Box ...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DcChequeBox.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                   
                End If
    
                If TXTBankName.text = "" Then
                           
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«ﬂ » «”„ »‰ﬂ «·‘Ìﬂ    « ...!!"
                    Else
                        Msg = " Enter Bank Name For Cheque  ...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    TXTBankName.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                    
                End If
        
                If Trim$(Me.TxtChequeNumber.text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If

            Else
       
                If Me.DcboBankName.BoundText = "" Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DcboBankName.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                End If

                If Trim$(Me.TxtChequeNumber.text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If
            End If
    
        ElseIf Me.CboPaymentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·Â...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
     
        ElseIf Me.CboPaymentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
     
        End If

        Dim notes_result As String
        Dim Vchr_result As String
        
        If TxtNoteSerial1.text = "" Then
            Vchr_result = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4)

            If Vchr_result = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁ»÷ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                
                If Vchr_result = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    ' txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4)
                End If
            End If
        End If
    
        If TxtNoteSerial.text = "" Then
            notes_result = Notes_coding(val(my_branch), XPDtbTrans.value)

            If notes_result = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If notes_result = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    '     TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rs.AddNew
       
            rs("NoteID").value = val(XPTxtID.text)
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
         
        ElseIf TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If

        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("EmpId").value = IIf(Me.DcEmp.BoundText = "", Null, (Me.DcEmp.BoundText))
        rs("foxy_no").value = val(Text1.text)
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    
        rs("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
        rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    
        If TxtNoteSerial1.text = "" Then
            TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4)
        End If
    
        If TxtNoteSerial.text = "" Then
            TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
        End If
    If Option1.value = True Then
       rs("NCashingType").value = 1
   ElseIf Option2.value = True Then
        rs("NCashingType").value = 2
   ElseIf Option3.value = True Then
        rs("NCashingType").value = 3
    Else
    
         rs("NCashingType").value = 0
   End If
   
    
        rs("ManulaNO").value = IIf(Trim(Me.TxtManulaNO.text) = "", Null, Trim(Me.TxtManulaNO.text))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    
        rs("person").value = IIf(txtperson.text = "", "", Trim(txtperson.text))
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
        rs("Adv_payment_value").value = IIf(txtAdv_payment_value.text = "", Null, val(txtAdv_payment_value.text))
    
        '    Rs("Remark").value = IIf(dcproject.BoundText = "", "", Trim(dcproject.BoundText))
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("BankName").value = IIf(TXTBankName.text = "", "", Trim(TXTBankName.text))

        rs("NoteType").value = 4
           rs("NoteDate").value = XPDtbTrans.value
        'rs("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        rs("NoteDateH").value = Me.Txt_DateHigri.value

        Select Case DCboCashType.ListIndex

            Case 0, 1

                If Me.ChkTrans.value = vbChecked Then
                    If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                        rs("Transaction_ID").value = val(Me.TxtTransID.text)
                        rs("MaintananceID").value = Null
                    ElseIf Me.CboTrans.ListIndex = 2 Then
                        rs("Transaction_ID").value = Null
                        rs("MaintananceID").value = val(Me.TxtTransID.text)
                    End If

                Else
                    rs("Transaction_ID").value = Null
                    rs("MaintananceID").value = Null
                End If

                rs("RevenuesID").value = Null

            Case 2
                rs("Transaction_ID").value = Null
                rs("MaintananceID").value = Null
                rs("RevenuesID").value = Null

            Case 3
                rs("RevenuesID").value = val(Me.DcboRevenuesTypes.BoundText)
                rs("Transaction_ID").value = Null
                rs("MaintananceID").value = Null

            Case 4
                '       Set rs1 = New ADODB.Recordset
                '       StrSQL = "select * From Transactions"
                '       rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
                '       rs1.AddNew
                '       rs1("Transaction_ID").value = Val(XPTxtBillID.text)
                '       rs1("Transaction_Date").value = XPDtbTrans.value
                '       rs1("Transaction_Type").value = 23
                '       rs1.update
                '
                '        Rs("Transaction_ID").value = Val(XPTxtBillID.text)
                '
        End Select

        rs("CashingType").value = DCboCashType.ListIndex
    
        If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Or Me.DCboCashType.ListIndex = 4 Then
            rs("CusID").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
     
        ElseIf Me.DCboCashType.ListIndex = 5 Then
            Dim X As Double
                    If IsNull(rs("note_count").value) Then
                         rs("note_count").value = CStr(new_id("Notes", "note_count", " ", True, " project_id=" & val(DBCboClientName.BoundText) & ""))
                    End If
            
            If Option4.value = True Then
                X = get_project_customer_id(DBCboClientName.BoundText, "End_user_Account")
            Else
                X = get_project_customer_id(DBCboClientName.BoundText, "sub_contractor_Account")
            End If

            rs("CusID").value = X
     
        Else
            rs("CusID").value = Null
        End If

        '--------------------------------------------------------------------------
        'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
        If Me.CboPaymentType.ListIndex = 0 Then
            rs("NoteCashingType").value = 0
            rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
        
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            rs("NoteCashingType").value = 1
            rs("BoxID").value = Null

            If SystemOptions.ChequeBox = False Then
        
                rs("BankID").value = val(Me.DcboBankName.BoundText)
            Else
                rs("BankID").value = Null
            End If
        
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value

            If SystemOptions.ChequeBox = True Then
                rs("ChequeBoxID").value = IIf(DcChequeBox.BoundText = "", Null, DcChequeBox.BoundText)
            Else
                rs("ChequeBoxID").value = Null
                
            End If
                
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
                
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            rs("NoteCashingType").value = 3
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
                
        End If

        '--------------------------------------------------------------------------
        rs("UserID").value = user_id
        rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(2) '”‰œ «·ﬁ»÷
    
        If DCboCashType.ListIndex = 5 Then
            rs("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
        End If
    
        If DCboCashType.ListIndex = 6 Then
            rs("EmployeeID").value = IIf(DcEmployee.BoundText = "", 0, DcEmployee.BoundText)
        End If
    
        If DCboCashType.ListIndex = 7 Then
            rs("AccountsCode").value = IIf(Me.DCAccounts.BoundText = "", Null, DCAccounts.BoundText)
        End If
    
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
    
        If DCboCashType.ListIndex = 5 Then
            rs("note_value_by_characters").value = WriteNo(val(Me.XPTxtVal.text) * 2, 0, True)
        Else
            rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        End If

        If Option4.value = True Then
            rs("cus_or_sub").value = 0 '⁄„Ì· ‰Â«∆Ì
        Else
            rs("cus_or_sub").value = 1 '⁄„Ì· »«ÿ‰
        End If
    
        rs.update

        saveChequeBoxContents (XPTxtID.text)
        '==========================================================================
    
        Line1 = setfoxy_Line
        Line2 = setfoxy_Line
        Line3 = setfoxy_Line
        Line4 = setfoxy_Line

        ' ”ÃÌ· ﬁÌÊœ
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
            RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            '«·ÿ—› «·„œÌ‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("DEV_ID_Line_No1").value = Line1
            
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            '    If DCboCashType.ListIndex = 5 Then
            'RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
            '  End If

            RsDev.update
            '«·ÿ—› «·œ«∆‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 2
            RsDev("DEV_ID_Line_No1").value = Line2
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            '  If DCboCashType.ListIndex = 5 Then
            '      RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
            '  End If
            RsDev("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
    
            RsDev.update

            If DCboCashType.ListIndex = 5 And Option3.value = False Then
                '«·„‘«—Ì⁄
                Dim account_codeLegal As String
                Dim account_codeREVENUE_account As String
       
                account_codeLegal = get_project_Account(val(DBCboClientName.BoundText), "legal")
                account_codeREVENUE_account = get_project_Account(val(DBCboClientName.BoundText), "REVENUE_account")

                If account_codeLegal = "" Or account_codeREVENUE_account = "" Then GoTo ll
       
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 3
                RsDev("DEV_ID_Line_No1").value = Line3
            
                RsDev("Account_Code").value = account_codeLegal
                RsDev("Value").value = val(Me.XPTxtVal.text)
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID

                If DCboCashType.ListIndex = 5 Then
                    RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                End If

                RsDev.update
                '«·ÿ—› «·œ«∆‰
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 4
                RsDev("DEV_ID_Line_No1").value = Line4
                RsDev("Account_Code").value = account_codeREVENUE_account
                RsDev("Value").value = val(Me.XPTxtVal.text)
                RsDev("Credit_Or_Debit").value = 1
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID

                If DCboCashType.ListIndex = 5 Then
                    RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                End If
    
                RsDev.update
ll:
            End If

            LblDevID.Caption = LngDevID
            lbl(33).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If

        '==========================================================================
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        'Õ›Ÿ «·„” Œ·’« 
        If DCboCashType.ListIndex = 5 Then
            saveprojectBillPayment TxtNoteSerial.text, val(XPTxtVal.text)
  
        End If
    
        If DCboCashType.ListIndex = 5 Then
            FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
        End If
    
        If Me.ChkTrans.value = vbUnchecked Then
            Me.CboTrans.ListIndex = -1
            Me.TxtTransSerial.text = ""
            Me.TxtTransID.text = ""
        End If
    
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
        
            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                lbl(46).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
        
        End Select
    
        '   If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "„ﬁ»Ê÷« ", Me.XPDtbTrans.value
        save_cost_center
        '   End If
        
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„œ›Ê⁄«  Ê «·„ﬁ»Ê÷« 
     
        If SavePaymentAndReciveDetails(1, TxtNoteSerial.text, TxtNoteSerial1.text, "", XPDtbTrans.value) = True Then
        End If

        TxtModFlg.text = "R"
    End If

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
    WriteInfo

    If Option1.value = True Then
        FIFO_FUNCTION val(DBCboClientName.BoundText)
    End If
   
    If Option2.value Then
        Distribute_to_bills Me.lblsqlstring, val(DBCboClientName.BoundText)
    End If
   
    TxtModFlg.text = "R"
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    
    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function saveChequeBoxContents(NoteID As Double)

    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    If val(DcChequeBox.BoundText) = 0 Then Exit Function
 
'    rs.Open "TblChecqueBoxContent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     * from dbo.TblChecqueBoxContent Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    rs.AddNew
    rs("noteid").value = NoteID
    rs("ChequeBoxID").value = val(DcChequeBox.BoundText)
            
    rs("RecordDate").value = XPDtbTrans.value
    rs("DueDate").value = DtpChequeDueDate.value
    rs("BankName").value = TXTBankName.text
    rs("ChequeNo").value = TxtChequeNumber.text
    rs("ChequeValue").value = val(XPTxtVal.text)
    
    rs("Remarks").value = DcboCreditSide.text
    rs("Deposited").value = 0
    rs("Collected").value = 0
    rs("CreditAccount").value = (DcboCreditSide.BoundText)
    
            If DCboCashType.ListIndex = 0 Then
                        rs("customeraccount").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
                        rs("customeraccount1").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
                        rs("customeraccount2").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                        
             ElseIf DCboCashType.ListIndex = 5 Then
                       rs("customeraccount").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code")
                        rs("customeraccount1").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1")
                        rs("customeraccount2").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2")
                        
              
              
            End If
    
    rs.update
  
    rs.Close
End Function

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtNoteSerial.text
        rs("Remark").value = "”‰œ „ﬁ»Ê÷«     —ﬁ„ " & TxtNoteSerial1.text & "    " & Me.TxtCustCode
 
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

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
 
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    'ÿ—› „œÌ‰
    '       rs.AddNew
    '       rs("cost_center_id").value = cost_center_id
    '       rs("cost_center").value = cost_center
    '       rs("value").value = XPTxtVal.text
    '       rs("depit_or_credit").value = "„œÌ‰"
    '       rs("opr_id").value = Me.Text1.text
    '       rs("kedno").value = Me.Text1.text
    '
    '       rs("opr_type").value = opr_type
    '       rs("account_name").value = DcboDebitSide.text
    '       rs("account_no").value = DcboDebitSide.BoundText
    '       rs("line_no").value = Line1
    '       rs("record_date").value = record_date
    '       rs.update
    'ÿ—› œ«∆‰
    rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = XPTxtVal.text
    rs("depit_or_credit").value = "œ«∆‰"
    rs("opr_id").value = Me.Text1.text
    rs("kedno").value = Me.Text1.text

    rs("opr_type").value = opr_type
    rs("account_name").value = DcboCreditSide.text
    rs("account_no").value = DcboCreditSide.BoundText
    rs("line_no").value = Line2
    rs("record_date").value = record_date
    rs.update
 
    rs.Close
End Function

Function change_adv_payment_value(note_id As Double, value As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT * from notes   where  NoteID=" & note_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Rs3("Adv_payment_value").value = value
    Rs3.update
  
End Function

Function Distribute_to_bills(Sql1 As String, CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where  requiredvalue>0 and " & Sql1
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(txtAdv_payment_value.text)
  
    For i = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2000, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, val(DCboUserName.BoundText)
        Rs3.MoveNext
    Next i

    txtAdv_payment_value.text = total_value
    change_adv_payment_value XPTxtID.text, total_value

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close
 
End Function

Function FIFO_FUNCTION(CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where   (cusid=" & CusID & " and requiredvalue>0)"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(txtAdv_payment_value.text)
  
    For i = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2000, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, val(DCboUserName.BoundText)
        Rs3.MoveNext
    Next i

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
    txtAdv_payment_value.text = total_value
    change_adv_payment_value XPTxtID.text, total_value
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close

End Function

Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Integer, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
    Dim RsDev As New ADODB.Recordset
    RsDev.Open "notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '
        
    RsDev.AddNew
      
    RsDev("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    RsDev("NoteSerial").value = TxtNoteSerial.text ' CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=2000"))
              
    RsDev("NoteDate").value = NoteDate
    RsDev("NoteType").value = NoteType
           
    RsDev("Note_Value").value = Note_Value
    RsDev("Transaction_ID").value = Transaction_ID
    RsDev("CusID").value = CusID
    RsDev("BoxID").value = BoxID
    RsDev("UserID").value = UserID
    RsDev("displayed").value = 0
           
    RsDev.update

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        If Me.CboPaymentType.ListIndex = 0 Then
            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), Date, False) = False Then
                Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
        End If
    
        '      If Me.DCChequeBox.BoundText <> "" Then
        '      If ChequeBoxOperations(Val(Me.XPTxtID)) = False Then
        '          Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
        '          Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
        '          MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '          Exit Sub
        '      End If
        '  End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (TxtNoteSerial.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                Dim StrSQL As String
                StrSQL = "Delete From notes  Where  (NoteType=2000 OR NoteType=4 ) AND  NoteSerial=" & val(TxtNoteSerial.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
        
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
       
                StrSQL = "Delete From ReciveDetails Where NoteSerial1='" & val(TxtNoteSerial1.text) & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete From ProjectBillBuy Where TxtNoteSerial='" & TxtNoteSerial.text & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    clear_all Me
                    Retrive
                End If

                '--------
                WriteInfo
                '-------
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub ChangeLang()
    lbl(43).Caption = "Cheque Box"
    lbl(50).Caption = "Car"
    lbl(49).Caption = "Driver"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(35).Caption = "Adv. Payment"
    Frame1.Caption = "Options"
    Option3.Caption = "Adv. Payment"
    Option2.Caption = "Select Invoice"
    ALLButton3.Caption = "Select"
    lbl(22).Caption = "Current Week"
    Label8.Caption = "General C.C."
    lbl(36).Caption = "From"
 
    Cmd(9).Caption = "GL Print"
 
    Label2.Caption = "Branch"
    lbl(47).Caption = "Value"

    Frame2.Caption = "Project"
    Option4.Caption = "End User"
    Option5.Caption = "Sub-contractor"

    LblLink.Visible = False
    lbl(18).Visible = False
    ALLButton1.Caption = "Installment view"
    ALLButton2.Caption = "debt Voucher"
    Me.Caption = "Receipts"
    Me.XPTab301.TabCaption(0) = "Receipts"
    Me.XPTab301.TabCaption(1) = "Invoices"
    lbl(37).Caption = "Total Rec."""
    lbl(0).Caption = "Select bills"
    lbl(42).Caption = "Payed  bills"
    CmdRemove.Caption = "Remove Row"

    With Grid

        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("id")) = "Invoice No."
        .TextMatrix(0, .ColIndex("bill_date")) = "Invoice Date"
        .TextMatrix(0, .ColIndex("total")) = "Invoice Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed Totalt"
        .TextMatrix(0, .ColIndex("result")) = "Not Payed"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Not Payed%"
 
    End With

    With Grid1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("bill_id")) = "Invoice Id"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("id")) = "Invoice No."
        .TextMatrix(0, .ColIndex("bill_date")) = "Invoice Date"
        .TextMatrix(0, .ColIndex("total")) = "Invoice Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed Totalt"
        .TextMatrix(0, .ColIndex("result")) = "Not Payed"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Not Payed%"
 
    End With

    Ele(1).Caption = Me.Caption
    lbl(4).Caption = "Opr Code"
    lbl(1).Caption = "Date"
    'lbl(0).Caption = "Type"
    lbl(3).Caption = "Name"
    lbl(2).Caption = "Value"
    lbl(14).Caption = "Cash/Cheque"
    lbl(9).Caption = "Box Name"
    lbl(15).Caption = "Bank Name"
    lbl(16).Caption = "Cheque #"
    lbl(17).Caption = "Cheque Name"
    lbl(5).Caption = "Note"
    ChkTrans.Caption = "From bill"
    lbl(12).Caption = "Bill type"
    lbl(10).Caption = "Bill #"
    lbl(13).Caption = "Current Balance"
    FraInfo.Caption = "Information"
    lbl(22).Caption = "Current Week"

    lbl(23).Caption = "Today Receipts "
    lbl(27).Caption = "Cash"
    lbl(28).Caption = "Cheque"

    lbl(19).Caption = "Week Receipts "

    lbl(21).Caption = "Cash"
    lbl(24).Caption = "Cheque"

    lbl(20).Caption = "Month Receipts "

    lbl(25).Caption = "Cash"
    lbl(26).Caption = "Cheque"
    Fra(1).Caption = "GL"

    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"

    lbl(32).Caption = "Depit"
    lbl(31).Caption = "Credit"
    Cmd(8).Caption = "Table view"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Current "
    lbl(6).Caption = "Records Count "

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    DCboCashType.Clear
    DCboCashType.AddItem "To Customer"
    DCboCashType.AddItem "To Vendor"
    DCboCashType.AddItem "Sub-contractor"
    DCboCashType.AddItem "Another Revenues"
    DCboCashType.AddItem "Advanced Payment"
    DCboCashType.AddItem "Projects"
    DCboCashType.AddItem "From Employee"
    DCboCashType.AddItem "From  Account"

    With Me.CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Bank Transfer"
        .AddItem "Coll. Cheque"
    
    End With

    With Me.CboTrans
        .Clear
        .AddItem "Sales invoice"
        .AddItem "Returned purchases"
        .AddItem "Delivery of maintenance for a client"
        .AddItem "Services"
    End With
 
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·”‰œ " & TxtNoteSerial1.text & Chr(13) & "   «· «—ÌŒ " & XPDtbTrans & Chr(13) & "   ‰Ê⁄ «·„ﬁ»Ê÷«  " & DCboCashType & Chr(13) & "   «·›—⁄  " & dcBranch & Chr(13) & "   «·«”„  " & DBCboClientName & Chr(13) & "   ﬁÌ„Â «·„ﬁ»Ê÷«   " & XPTxtVal & Chr(13) & "   ÿ—Ìﬁ… «·ﬁ»÷ " & CboPaymentType & Chr(13) & "   «·Œ“Ì‰…  " & DcboBox & Chr(13) & "   «·»‰ﬂ  " & DcboBankName & Chr(13) & "   —ﬁ„ «·‘Ìﬂ  " & TxtChequeNumber & Chr(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & Chr(13) & "     »‰«¡ ⁄·Ï   " & XPMTxtRemarks & Chr(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & Chr(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & Chr(13) & "ÿ—› „œÌ‰  " & DcboDebitSide & Chr(13) & " ÿ—› œ«∆‰ " & DcboCreditSide & Chr(13) & " «·„‰œÊ» " & DcEmp
                        
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Vchr. NO.  " & TxtNoteSerial1.text & Chr(13) & "   Date " & XPDtbTrans & Chr(13) & "  Payment Type " & DCboCashType & Chr(13) & "   Branch  " & dcBranch & Chr(13) & "   Name  " & DBCboClientName & Chr(13) & "  Value" & XPTxtVal & Chr(13) & "   Cash/   Cheque " & CboPaymentType & Chr(13) & "   Box  " & DcboBox & Chr(13) & "   Bank  " & DcboBankName & Chr(13) & "   Cheque No" & TxtChequeNumber & Chr(13) & "  Due Date  " & DtpChequeDueDate & Chr(13) & " Ge NO.  " & TxtNoteSerial & Chr(13) & "Debit " & DcboDebitSide & Chr(13) & "Credit " & DcboCreditSide & Chr(13) & " UserName " & DCboUserName & Chr(13) & " Sales Person " & DcEmp
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 4, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, , , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 4, Date, Time, LogTextA, LogTextE, Me.name, "D", , , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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
            'Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
End Sub

Private Sub Txt_DateHigri_LostFocus()
    XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
 
End Sub

Private Sub XPTxtVal_Change()
    'Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)
    'txtAdv_payment_value.text = Format(Val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1)

    End If

    'If TxtModFlg.text = "N" Or TxtModFlg.text = "E" And Option3.value = True Then
    If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        txtAdv_payment_value.text = XPTxtVal.text
    End If

End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Function CheckDebitTrans(LngTransID As Long) As Boolean
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitTrans = False

    If LngTransID = 0 Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From Transactions where Transaction_ID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & Chr(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.text) Then
                Me.TxtTransID.text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType, " & "Notes.Note_Value, Notes.NoteID "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID WHERE (Notes.NoteType=1) AND Transactions.Transaction_ID= " & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                        Msg = Msg & Chr(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.text) Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & Chr(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID " & " Where ((Notes.NoteType = 4 OR Notes.NoteType = 9) And Transactions.Transaction_ID = " & LngTransID & ")"

            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
            End If

            StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType "
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    Msg = "⁄›Ê« ...!!!!!" & Chr(13)
                    Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & Chr(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & Chr(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Function
                End If
            End If

        Else
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitTrans = True
    Exit Function
ErrTrap:
End Function

Private Function CheckDebitMaintaince(LngTransID As Long) As Boolean
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitMaintaince = False

    If LngTransID = 0 Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From TblMaintenece where MaintananceID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & Chr(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.text) Then
                Me.TxtTransID.text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Notes.Note_Value, Notes.NoteID, TblMaintenece.MaintananceID," & "TblMaintenece.PaymentType, TblMaintenece.MType "
            StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON " & "TblMaintenece.MaintananceID = Notes.MaintananceID " & " WHERE (((Notes.NoteType)=1)) AND TblMaintenece.MaintananceID=" & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                        Msg = Msg & Chr(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.text) Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & Chr(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
        
            StrSQL = "SELECT  TblMaintenece.MaintananceID," & "TblMaintenece.MType, TblMaintenece.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON TblMaintenece.MaintananceID =" & "Notes.MaintananceID " & " Where ((Notes.NoteType = 4) And TblMaintenece.MaintananceID = " & LngTransID & ")"

            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
            End If

            StrSQL = StrSQL + " GROUP BY TblMaintenece.MaintananceID," & "TblMaintenece.MType, TblMaintenece.PaymentType"
        
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    Msg = "⁄›Ê« ...!!!!!"
                    Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & Chr(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & Chr(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Function
                End If
            End If

        Else
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitMaintaince = True
    Exit Function
ErrTrap:
End Function

Public Function CheckDebitService()

End Function

Private Sub WriteInfo()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StartWeekDate As Date
    Dim EndWeekDate As Date
    Dim StrTemp As String
    Dim i As Integer

    StartWeekDate = GetWeekStartEND(Date, 0)
    EndWeekDate = DateAdd("d", 7, StartWeekDate)

    If SystemOptions.UserInterface = ArabicInterface Then
        StrTemp = "«·≈”»Ê⁄ «·Õ«·Ï „‰ " & DisplayDate(StartWeekDate)
        StrTemp = StrTemp & " ≈·Ï " & DisplayDate(EndWeekDate)
    Else
        StrTemp = "«Current Week From " & DisplayDate(StartWeekDate)
        StrTemp = StrTemp & " To " & DisplayDate(EndWeekDate)

    End If

    Me.lbl(22).Caption = StrTemp

    For i = LblLinkInfo.LBound To LblLinkInfo.UBound
        LblLinkInfo(i).Caption = "0"
    Next i

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·ÌÊ„
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(0).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(1).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(6).Caption = val(Me.LblLinkInfo(0).Caption) + val(Me.LblLinkInfo(1).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(6).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·√”»Ê⁄ «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(StartWeekDate, True)
    StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(EndWeekDate, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(2).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(3).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(7).Caption = val(Me.LblLinkInfo(2).Caption) + val(Me.LblLinkInfo(3).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(7).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·‘Â— «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND Month(NoteDate)=" & Month(Date) & ""
    StrSQL = StrSQL + " AND Year(NoteDate)=" & year(Date) & ""
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(4).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(5).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(8).Caption = val(Me.LblLinkInfo(4).Caption) + val(Me.LblLinkInfo(5).Caption)
    Else
        Me.LblLinkInfo(4).Caption = 0
        Me.LblLinkInfo(5).Caption = 0
        Me.LblLinkInfo(8).Caption = 0
    End If

End Sub

