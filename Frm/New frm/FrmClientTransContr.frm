VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmClientTransContr 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "« ›«ﬁÌ«  ⁄„·«¡ «·‰ﬁ·"
   ClientHeight    =   7560
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmClientTransContr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   12135
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7545
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12165
      _cx             =   21458
      _cy             =   13309
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
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
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6555
         Width           =   12105
         _cx             =   21352
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   12240
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   0
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
            FontSize        =   9.75
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmClientTransContr.frx":038A
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmClientTransContr.frx":0724
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmClientTransContr.frx":0ABE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   10680
            TabIndex        =   5
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
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
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9240
            TabIndex        =   6
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
            Height          =   375
            Index           =   2
            Left            =   7800
            TabIndex        =   7
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            CausesValidation=   0   'False
            Height          =   375
            Index           =   3
            Left            =   6240
            TabIndex        =   8
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
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
            Height          =   375
            Index           =   4
            Left            =   4680
            TabIndex        =   9
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
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
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   11040
            TabIndex        =   12
            Tag             =   "Delete Row"
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   12632319
            BCOLO           =   12632319
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmClientTransContr.frx":0E58
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   2400
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   480
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmClientTransContr.frx":0E74
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   9
            Left            =   960
            TabIndex        =   14
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·Â"
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
         Begin MSDataListLib.DataCombo DcbUser 
            Height          =   315
            Left            =   6120
            TabIndex        =   15
            Top             =   120
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   20
            Left            =   9720
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   120
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4305
         Index           =   2
         Left            =   0
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2280
         Width           =   12075
         _cx             =   21299
         _cy             =   7594
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5715
            Index           =   1
            Left            =   0
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   15345
            _cx             =   27067
            _cy             =   10081
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
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   3060
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   -3960
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   7800
               Width           =   2190
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   12240
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   2520
               Width           =   2235
            End
            Begin VB.TextBox txtType 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Text            =   "0"
               Top             =   3120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CheckBox ChKauto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ì"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   2760
               Width           =   1500
            End
            Begin MSDataListLib.DataCombo dcopr 
               Height          =   315
               Left            =   13200
               TabIndex        =   26
               Top             =   1800
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   556
               _Version        =   393216
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
            Begin MSDataListLib.DataCombo dcproject 
               Height          =   315
               Left            =   13200
               TabIndex        =   27
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
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
            Begin MSDataListLib.DataCombo Dcterm 
               Height          =   315
               Left            =   13440
               TabIndex        =   28
               Top             =   720
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   556
               _Version        =   393216
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
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   4140
               Left            =   0
               TabIndex        =   29
               Top             =   360
               Width           =   12060
               _cx             =   21272
               _cy             =   7302
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   23
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmClientTransContr.frx":76D6
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
               Height          =   390
               Index           =   21
               Left            =   0
               TabIndex        =   30
               Top             =   120
               Visible         =   0   'False
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   688
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
               ButtonImage     =   "FrmClientTransContr.frx":7A1C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   13905
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   840
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ«Ì… «· Œ’Ì’"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   8
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   2190
               Width           =   1800
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   8400
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   810
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   2385
         Left            =   0
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Width           =   12090
         _cx             =   21325
         _cy             =   4207
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   1200
            Width           =   4410
         End
         Begin VB.CheckBox ChkLocked 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ìﬁ«› «· ⁄«„·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox TxtTblCustomerContractD 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9360
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2685
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1950
            Width           =   1215
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   0
            Left            =   9720
            TabIndex        =   40
            Top             =   1920
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ„"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   765
            Index           =   5
            Left            =   0
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   0
            Width           =   12195
            _cx             =   21511
            _cy             =   1349
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   24
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
            Picture         =   "FrmClientTransContr.frx":7FB6
            Caption         =   "   « ›«ﬁÌ«  ⁄„·«¡ «·‰ﬁ·  "
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   0
            ChildSpacing    =   0
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
            PicturePos      =   0
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
               Height          =   375
               Index           =   0
               Left            =   1695
               TabIndex        =   42
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
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
               ButtonImage     =   "FrmClientTransContr.frx":8C90
               ColorButton     =   16777215
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
               Left            =   630
               TabIndex        =   43
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
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
               ButtonImage     =   "FrmClientTransContr.frx":902A
               ColorButton     =   16777215
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
               Left            =   2220
               TabIndex        =   44
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
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
               ButtonImage     =   "FrmClientTransContr.frx":93C4
               ColorButton     =   16777215
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
               Left            =   1155
               TabIndex        =   45
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
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
               ButtonImage     =   "FrmClientTransContr.frx":975E
               ColorButton     =   16777215
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
         Begin MSComCtl2.DTPicker dbFromDate 
            Height          =   270
            Left            =   3090
            TabIndex        =   46
            Top             =   840
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   476
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   253689857
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker dbTodate 
            Height          =   270
            Left            =   240
            TabIndex        =   47
            Top             =   840
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   476
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   253689857
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   5595
            TabIndex        =   48
            Top             =   1200
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo VehicleType 
            Height          =   315
            Left            =   5595
            TabIndex        =   49
            Top             =   1560
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   556
            _Version        =   393216
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
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   1
            Left            =   5670
            TabIndex        =   50
            Top             =   1920
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·—œ"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   2
            Left            =   4710
            TabIndex        =   51
            Top             =   1920
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·Ê“‰"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   20
            Left            =   120
            TabIndex        =   60
            Top             =   1920
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈÷«›…"
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
            ButtonImage     =   "FrmClientTransContr.frx":9AF8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   3
            Left            =   8220
            TabIndex        =   61
            Top             =   1920
            Width           =   1275
            _Version        =   786432
            _ExtentX        =   2249
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "»«·· —/«·Õ„Ê·…"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   4
            Left            =   6900
            TabIndex        =   62
            Top             =   1920
            Width           =   1185
            _Version        =   786432
            _ExtentX        =   2090
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«Ã—… «·‘Õ‰"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1395
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1155
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„œ Â« „‰"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   840
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„—ﬂ»… "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   50
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÊÕœ… «·ﬁÌ«”"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1920
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   3570
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   1920
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "FrmClientTransContr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim RecId As String
Dim rs As ADODB.Recordset
Private Sub Del_Trans()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If TxtTblCustomerContractD.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtTblCustomerContractD.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Cn.Execute "delete TblClientTransContrDet where ClintTransID=" & val(Me.TxtTblCustomerContractD.text)
                
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                        Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
                    '   XPTxtCurrent.Caption = 0
                    '   XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        If Trim(Me.DBCboClientName.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·⁄„Ì·..!!"
            Else
            MsgBox "Please Select Customer"
         End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "delete TblClientTransContrDet where ClintTransID=" & val(Me.TxtTblCustomerContractD.text)
   
    End If
    
    rs("ID").value = TxtTblCustomerContractD.text
    rs("CompID").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
    rs("FromDate").value = dbFromDate.value
    rs("Todate").value = dbTodate.value
    rs("Remarks").value = IIf(Me.TxtRemarks.text = "", "", Me.TxtRemarks.text)
    rs("UserID").value = IIf(Me.DcbUser.BoundText = "", Null, Me.DcbUser.BoundText)
    rs("VehicleType").value = val(VehicleType.BoundText)
    rs("Price").value = val(TxtPrice.text)
    If ChkLocked.value = vbChecked Then
        rs("LockedID").value = 1
    Else
        rs("LockedID").value = 0
    End If
    If RdTyped(1).value = True Then
    rs("Typed").value = 1
    ElseIf RdTyped(2).value Then
    rs("Typed").value = 2
    ElseIf RdTyped(3).value Then
    rs("Typed").value = 3
    ElseIf RdTyped(4).value Then
    rs("Typed").value = 4
    Else
    rs("Typed").value = 0
    End If
    
    rs.update
    Set RsDev = New ADODB.Recordset
    RsDev.Open "TblClientTransContrDet", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Dim i As Integer

    With Me.Grid

        For i = 1 To .rows - 1

            'If val(.TextMatrix(i, .ColIndex("VehicleType"))) <> 0 Then
                RsDev.AddNew
                RsDev("ClintTransID").value = val(Me.TxtTblCustomerContractD.text)
                RsDev("VehicleType").value = val(.TextMatrix(i, .ColIndex("VehicleType")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                
               RsDev("FromPrice").value = val(.TextMatrix(i, .ColIndex("FromPrice")))
               RsDev("ToPrice").value = val(.TextMatrix(i, .ColIndex("ToPrice")))
               RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
               RsDev("FromCityID").value = val(.TextMatrix(i, .ColIndex("FromCityID")))
               RsDev("ToCityID").value = val(.TextMatrix(i, .ColIndex("ToCityID")))
               
               RsDev("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
               RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
               
            
                
                RsDev("Remarks").value = (.TextMatrix(i, .ColIndex("Remarks")))
                RsDev("Typed").value = val(.TextMatrix(i, .ColIndex("Typed")))
                RsDev.update
                    
         '   End If
            
            '
        Next i

    End With
 
    RsDev.Close
    
    Cn.CommitTrans
    BeginTrans = False

    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.text = "R"
    'End If

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Cmd_Click(index As Integer)
    On Error GoTo ErrTrap

    Select Case index
Case 9
  TxtModFlg.text = "N"
      Me.TxtTblCustomerContractD.text = CStr(new_id("TblClientTransContr", "ID", "", True))
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
           
            Me.TxtTblCustomerContractD.text = CStr(new_id("TblClientTransContr", "ID", "", True))
            DcbUser.BoundText = user_id
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
            Grid.Enabled = True
            RdTyped(0).value = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            '  Load FrmCustomerContractSearch
            '  FrmCustomerContractSearch.show vbModal

        Case 6
            Unload Me
        Case 20
        Dim Msg As String
'        If Trim(Me.VehicleType.BoundText) = "" Or val(VehicleType.BoundText) = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·„—ﬂ»…..!!"
'            Else
'            MsgBox "Please Select Vehicle Type"
'         End If
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            VehicleType.SetFocus
'            Sendkeys "{F4}"
'            Exit Sub
'        End If
            addrow
        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

End Sub
Sub addrow()
Dim sql As String
Dim k As Integer
Dim i As Integer
With Me.Grid
k = .rows
.rows = .rows + 1
For i = k To .rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Price")) = val(TxtPrice.text)
.TextMatrix(i, .ColIndex("Vehname")) = VehicleType.text
.TextMatrix(i, .ColIndex("VehicleType")) = val(VehicleType.BoundText)
.TextMatrix(i, .ColIndex("Remarks")) = TxtRemarks.text
If RdTyped(0).value = True Then
.TextMatrix(i, .ColIndex("Typed")) = 1
ElseIf RdTyped(1).value = True Then
.TextMatrix(i, .ColIndex("Typed")) = 2
ElseIf RdTyped(2).value = True Then
.TextMatrix(i, .ColIndex("Typed")) = 3
ElseIf RdTyped(3).value = True Then
.TextMatrix(i, .ColIndex("Typed")) = 4
ElseIf RdTyped(4).value = True Then
.TextMatrix(i, .ColIndex("Typed")) = 1
End If
Next i
End With

End Sub


Private Sub RemoveGridRow()
    With Me.Grid
        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With
    ReLineGrid
End Sub



Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
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
    
    If Grid.rows > 1 Then
        If Grid.rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.rows > 1 Then
                If Me.Grid.row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.row)
                End If
            End If
        End If
    End If


End Sub

Private Sub DBCboClientName_Change()
  On Error Resume Next
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim Fullcode  As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.text = Fullcode
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 20916
        FrmCustemerSearch.show vbModal

    End If
End Sub


Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    With Grid
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰|#4;· —|#3;Õ„Ê·… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("Typed")) = "#1;Kg |#2;RD |#3;Weight|#4;Litr|#5;Weight"
            End If
          Set .WallPaper = GrdBack.Picture
    End With
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
    Dcombos.GetUsers Me.DcbUser
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
    Dcombos.GetTblCarsDataGroup VehicleType
    With Me.Grid
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblClientTransContr  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    ISButton2.Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Company Contarcts"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = "Start Date"
    lbl(2).Caption = "End Date"
 
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

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

Private Sub Grid_AfterEdit(ByVal row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    With Grid

        Select Case .ColKey(Col)
        
           Case "Vehname"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("VehicleType"), False, True)
             .TextMatrix(row, .ColIndex("VehicleType")) = StrAccountCode
         Case "FromCity"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("VehicleType"), False, True)
             .TextMatrix(row, .ColIndex("FromCityID")) = StrAccountCode
        Case "ToCity"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("VehicleType"), False, True)
             .TextMatrix(row, .ColIndex("ToCityID")) = StrAccountCode
        Case "ItemCode"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemCode"), False, True)
             .TextMatrix(row, .ColIndex("ItemID")) = StrAccountCode
        Case "ItemName"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemName"), False, True)
             .TextMatrix(row, .ColIndex("ItemID")) = StrAccountCode
        Case "UnitName"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitName"), False, True)
             .TextMatrix(row, .ColIndex("UnitID")) = StrAccountCode


        End Select
   
        If row = .rows - 1 Then
          .rows = .rows + 1
        End If

        ReLineGrid
    End With
    
    
   StrSQL = " Select UnitID,UnitName,UnitNamee from TblUnits Inner join TblItemsUnits On TblItemsUnits.UnitID = TblItemsUnits.UnitID "
               StrSQL = StrSQL & " Where TblItemsUnits.ItemId = " & val(Grid.TextMatrix(row, Grid.ColIndex("ItemID")))
               
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With Me.Grid
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("VehicleType"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    With Grid
        If .ColKey(Col) <> "Name" Then
            .ComboList = ""
        End If
Select Case .ColKey(Col)
'Case "Price"
'Cancel = True
'Case "Remarks"
'Cancel = True
End Select
    End With

End Sub

Private Sub Grid_StartEdit(ByVal row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim MyStrList As String
    With Me.Grid
        Select Case .ColKey(Col)
            Case "Vehname"
               StrSQL = "  SELECT  * From TBLCarTypes "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "ID")
              Else
                    StrComboList = .BuildComboList(rs, "NameE", "ID")
             End If
             .ComboList = StrComboList
             
            Case "FromCity"
               StrSQL = " Select TblCountriesGovernments.GovernmentID ID,TblCountriesGovernments.GovernmentName Name,TblCountriesGovernments.GovernmentName NameE from TblCountriesGovernments "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "ID")
              Else
                    StrComboList = .BuildComboList(rs, "NameE", "ID")
             End If
             .ComboList = StrComboList


            Case "ItemCode"
               StrSQL = " Select ItemCode,ItemName,ItemNamee,ItemID from TblItems "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "ItemCode", "ItemID")
              Else
                    StrComboList = .BuildComboList(rs, "ItemCode", "ItemID")
             End If
             .ComboList = StrComboList
             
        Case "ItemName"
               StrSQL = " Select ItemCode,ItemName,ItemNamee,ItemID from TblItems "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "ItemName", "ItemID")
              Else
                    StrComboList = .BuildComboList(rs, "ItemNamee", "ItemID")
             End If
             .ComboList = StrComboList
             
        Case "UnitName"
               StrSQL = " Select TblUnites.UnitID,UnitName,UnitNamee from TblUnites Inner join TblItemsUnits On TblItemsUnits.UnitID = TblUnites.UnitID "
               StrSQL = StrSQL & " Where TblItemsUnits.ItemId = " & val(Grid.TextMatrix(row, Grid.ColIndex("ItemID")))
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "UnitName", "UnitID")
              Else
                    StrComboList = .BuildComboList(rs, "UnitNamee", "UnitID")
             End If
             .ComboList = StrComboList
            Case "ToCity"
               StrSQL = " Select TblCountriesGovernments.GovernmentID ID,TblCountriesGovernments.GovernmentName Name,TblCountriesGovernments.GovernmentName NameE from TblCountriesGovernments "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "ID")
              Else
                    StrComboList = .BuildComboList(rs, "NameE", "ID")
             End If
             .ComboList = StrComboList


        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1

    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
   VehicleType.BoundText = IIf(IsNull(rs("VehicleType").value), "", rs("VehicleType").value)
    Me.TxtTblCustomerContractD.text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
    dbFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dbTodate.value = IIf(IsNull(rs("Todate").value), Date, rs("Todate").value)
    DBCboClientName.BoundText = IIf(IsNull(rs("CompID").value), "", rs("CompID").value)
    TxtRemarks.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    Me.DcbUser.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    TxtPrice.text = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
    If Not IsNull(rs("LockedID").value) Then
    If (rs("LockedID").value) = 1 Then
     ChkLocked.value = vbChecked
     Else
     ChkLocked.value = vbUnchecked
     End If
    Else
    ChkLocked.value = vbUnchecked
    End If
     If Not IsNull((rs("Typed").value)) Then
        If ((rs("Typed").value)) = 2 Then
            RdTyped(2).value = True
        ElseIf ((rs("Typed").value)) = 1 Then
            RdTyped(1).value = True
        ElseIf ((rs("Typed").value)) = 3 Then
            RdTyped(3).value = True
        ElseIf ((rs("Typed").value)) = 4 Then
            RdTyped(4).value = True
        Else
            RdTyped(0).value = True
        End If
     Else
        RdTyped(0).value = True
     End If
     
    StrSQL = " SELECT     dbo.TblClientTransContrDet.ClintTransID, dbo.TblClientTransContrDet.Price,dbo.TblClientTransContrDet.FromPrice,dbo.TblClientTransContrDet.ToPrice, dbo.TblClientTransContrDet.Remarks, dbo.TblClientTransContrDet.Typed, "
    StrSQL = StrSQL & "                  TblClientTransContrDet.FromCityID,TblClientTransContrDet.ToCityID ,"
    StrSQL = StrSQL & "                   dbo.TBLCarTypes.Name , dbo.TBLCarTypes.NameE, dbo.TblClientTransContrDet.VehicleType,CC2.GovernmentName as ToCity,TblCountriesGovernments.GovernmentName as FromCity,"
    StrSQL = StrSQL & "                   TblItems.ItemName,TblItems.ItemNamee,TblItems.ItemCode,TblItems.ItemID,TblUnites.UnitID,TblUnites.UnitName,TblUnites.UnitNamee"
    StrSQL = StrSQL & "     FROM         dbo.TblClientTransContrDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TBLCarTypes ON dbo.TblClientTransContrDet.VehicleType = dbo.TBLCarTypes.id"
    StrSQL = StrSQL & "                  LEFT OUTER JOIN TblCountriesGovernments On TblCountriesGovernments.GovernmentID =TblClientTransContrDet.FromCityID "
    StrSQL = StrSQL & "                  LEFT OUTER JOIN TblCountriesGovernments CC2 On CC2.GovernmentID =TblClientTransContrDet.ToCityID "
    StrSQL = StrSQL & "                  LEFT OUTER JOIN TblItems  On TblItems.ItemID =TblClientTransContrDet.ItemID"
    StrSQL = StrSQL & "                  LEFT OUTER JOIN TblUnites  On TblUnites.UnitID =TblClientTransContrDet.UnitID"
    
    StrSQL = StrSQL & "  Where (dbo.TblClientTransContrDet.ClintTransID = " & TxtTblCustomerContractD.text & ")"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
        With Me.Grid
            .rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .rows - 1
            .TextMatrix(i, .ColIndex("VehicleType")) = IIf(IsNull(RsDev("VehicleType").value), "", RsDev("VehicleType").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
                .TextMatrix(i, .ColIndex("Typed")) = IIf(IsNull(RsDev("Typed").value), "", RsDev("Typed").value)
                .TextMatrix(i, .ColIndex("FromCityID")) = IIf(IsNull(RsDev("FromCityID").value), "", RsDev("FromCityID").value)
                .TextMatrix(i, .ColIndex("ToCityID")) = IIf(IsNull(RsDev("ToCityID").value), "", RsDev("ToCityID").value)
                .TextMatrix(i, .ColIndex("FromPrice")) = IIf(IsNull(RsDev("FromPrice").value), "", RsDev("FromPrice").value)
                .TextMatrix(i, .ColIndex("ToPrice")) = IIf(IsNull(RsDev("ToPrice").value), "", RsDev("ToPrice").value)
                
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDev("ItemID").value), "", RsDev("ItemID").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(RsDev("UnitID").value), "", RsDev("UnitID").value)
                
                .TextMatrix(i, .ColIndex("FromCity")) = IIf(IsNull(RsDev("FromCity").value), "", RsDev("FromCity").value)
                .TextMatrix(i, .ColIndex("ToCity")) = IIf(IsNull(RsDev("ToCity").value), "", RsDev("ToCity").value)
                
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Vehname")) = IIf(IsNull(RsDev("name").value), "", (RsDev("name").value))
               Else
               .TextMatrix(i, .ColIndex("Vehname")) = IIf(IsNull(RsDev("namee").value), "", (RsDev("namee").value))
            End If
            
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RsDev.Close
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
Private Sub ISButton2_Click()
On Error GoTo ErrTrap
   If val(Me.TxtTblCustomerContractD.text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
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
    sql = " SELECT     dbo.TblClientTransContrDet.ClintTransID, dbo.TblClientTransContrDet.Price, dbo.TblClientTransContrDet.Remarks, dbo.TblClientTransContrDet.Typed, "
    sql = sql & "                   dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblClientTransContrDet.VehicleType, dbo.TblClientTransContr.LockedID, dbo.TblClientTransContr.FromDate,"
    sql = sql & "                   dbo.TblClientTransContr.Todate, dbo.TblClientTransContr.Remarks AS HaedRemarks, dbo.TblClientTransContr.Typed AS HeadTyped,"
    sql = sql & "                   dbo.TblClientTransContr.Price AS HeadPrice, dbo.TblClientTransContr.CompID, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    sql = sql & "                   dbo.TblCustemers.Fullcode, dbo.TblClientTransContr.VehicleType AS HeadVehicleType, TBLCarTypes_1.name AS HeadName,"
    sql = sql & "                   TBLCarTypes_1.namee AS HeadNameE"
    sql = sql & "    FROM         dbo.TBLCarTypes TBLCarTypes_1 RIGHT OUTER JOIN"
    sql = sql & "                   dbo.TblClientTransContr ON TBLCarTypes_1.id = dbo.TblClientTransContr.VehicleType LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblCustemers ON dbo.TblClientTransContr.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblClientTransContrDet ON dbo.TblClientTransContr.ID = dbo.TblClientTransContrDet.ClintTransID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TBLCarTypes ON dbo.TblClientTransContrDet.VehicleType = dbo.TBLCarTypes.id"
    sql = sql & "  Where (dbo.TblClientTransContrDet.ClintTransID = " & TxtTblCustomerContractD.text & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ClintTransContruct.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ClintTransContructE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function


Private Sub RdTyped_Click(index As Integer)

Grid.ColHidden(Grid.ColIndex("UnitName")) = True
    Grid.ColHidden(Grid.ColIndex("ItemName")) = True
        Grid.ColHidden(Grid.ColIndex("ItemCode")) = True
Select Case index
Case 0
    Grid.ColHidden(Grid.ColIndex("FromCity")) = False
    Grid.ColHidden(Grid.ColIndex("ToCity")) = False
    Grid.ColHidden(Grid.ColIndex("Vehname")) = False
    Grid.ColHidden(Grid.ColIndex("FromPrice")) = True
    Grid.ColHidden(Grid.ColIndex("ToPrice")) = True
    Grid.TextMatrix(0, Grid.ColIndex("FromPrice")) = "„‰ "
    Grid.TextMatrix(0, Grid.ColIndex("ToPrice")) = "«·Ï"
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
            Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;ﬂÃ„"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;Kg"
            End If
         
    
Case 1
    Grid.ColHidden(Grid.ColIndex("FromCity")) = False
    Grid.ColHidden(Grid.ColIndex("ToCity")) = False
    Grid.ColHidden(Grid.ColIndex("Vehname")) = False
    Grid.ColHidden(Grid.ColIndex("FromPrice")) = True
    Grid.ColHidden(Grid.ColIndex("ToPrice")) = True
    Grid.TextMatrix(0, Grid.ColIndex("FromPrice")) = "„‰ "
    Grid.TextMatrix(0, Grid.ColIndex("ToPrice")) = "«·Ï"
    
    If SystemOptions.UserInterface = ArabicInterface Then
            Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰|#4;· —|#3;Õ„Ê·… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;Kg |#2;RD |#3;Weight|#4;Litr|#5;Weight"
            End If
          
    
Case 2
    
    Grid.ColHidden(Grid.ColIndex("UnitName")) = False
    Grid.ColHidden(Grid.ColIndex("ItemName")) = False
        Grid.ColHidden(Grid.ColIndex("ItemCode")) = False
    
    
    Grid.ColHidden(Grid.ColIndex("FromCity")) = False
    Grid.ColHidden(Grid.ColIndex("ToCity")) = False
    Grid.ColHidden(Grid.ColIndex("Vehname")) = False
    Grid.ColHidden(Grid.ColIndex("FromPrice")) = True
    Grid.ColHidden(Grid.ColIndex("ToPrice")) = True
    Grid.TextMatrix(0, Grid.ColIndex("FromPrice")) = "„‰ "
    Grid.TextMatrix(0, Grid.ColIndex("ToPrice")) = "«·Ï"
    If SystemOptions.UserInterface = ArabicInterface Then
            Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰|#4;· —|#5;Õ„Ê·… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;Kg |#2;RD |#3;Weight|#4;Litr|#5;Weight"
            End If
          
    
Case 3
    Grid.ColHidden(Grid.ColIndex("FromCity")) = True
    Grid.ColHidden(Grid.ColIndex("ToCity")) = True
    Grid.ColHidden(Grid.ColIndex("Vehname")) = True
    Grid.ColHidden(Grid.ColIndex("Typed")) = True
    
    Grid.ColHidden(Grid.ColIndex("FromPrice")) = True
    Grid.ColHidden(Grid.ColIndex("ToPrice")) = True
    
      Grid.ColHidden(Grid.ColIndex("UnitName")) = False
    Grid.ColHidden(Grid.ColIndex("ItemName")) = False
        Grid.ColHidden(Grid.ColIndex("ItemCode")) = False
    
    Grid.ColHidden(Grid.ColIndex("ItemName")) = False
    Grid.ColHidden(Grid.ColIndex("ItemCode")) = False

    Grid.TextMatrix(0, Grid.ColIndex("FromPrice")) = "„‰"
    Grid.TextMatrix(0, Grid.ColIndex("ToPrice")) = "«·Ï"
    If SystemOptions.UserInterface = ArabicInterface Then
            Grid.ColComboList(Grid.ColIndex("Typed")) = "#4;· —"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           Grid.ColComboList(Grid.ColIndex("Typed")) = "#4;Litr"
            End If
          
    

Case 4
    Grid.ColHidden(Grid.ColIndex("FromCity")) = True
    Grid.ColHidden(Grid.ColIndex("ToCity")) = True
    Grid.ColHidden(Grid.ColIndex("Vehname")) = True
    Grid.ColHidden(Grid.ColIndex("FromPrice")) = False
    Grid.ColHidden(Grid.ColIndex("ToPrice")) = False
    Grid.TextMatrix(0, Grid.ColIndex("FromPrice")) = "„‰ "
    Grid.TextMatrix(0, Grid.ColIndex("ToPrice")) = "«·Ï"
    If SystemOptions.UserInterface = ArabicInterface Then
            Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;Kg"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           Grid.ColComboList(Grid.ColIndex("Typed")) = "#1;Kg"
            End If
Case 5
Case 6
End Select

End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False
        ISButton2.Enabled = False
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        ISButton2.Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False
       ISButton2.Enabled = True
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub XPBtnMove_Click(index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

    Select Case index

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
' aladein add
'''''''''''''''''''''''''''''''''''''''''
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    rs.Find "ID=" & RecId, , adSearchForward, 1
    If Not (rs.EOF) Then
        Retrive
        End If
    Exit Function
ErrTrap:
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
     '   BtnUndo_Click
    End If
  End Function
Private Sub DBCboClientName_Click(Area As Integer)
  On Error Resume Next
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim Fullcode  As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.text = Fullcode
End Sub
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If
 End Sub
'TxtItemCode
