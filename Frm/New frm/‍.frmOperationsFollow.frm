VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmOperationsFollow 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄Â «·⁄„·Ì«   "
   ClientHeight    =   8895
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14475
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
   Icon            =   "û.frmOperationsFollow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   14475
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8865
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14565
      _cx             =   25691
      _cy             =   15637
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
      AutoSizeChildren=   8
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"û.frmOperationsFollow.frx":000C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   7875
         Width           =   14505
         _cx             =   25585
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
            Left            =   11880
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   90
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
            ButtonImage     =   "û.frmOperationsFollow.frx":0091
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   4
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
            ButtonImage     =   "û.frmOperationsFollow.frx":042B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   5
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
            ButtonImage     =   "û.frmOperationsFollow.frx":07C5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   9660
            TabIndex        =   8
            Top             =   510
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            Height          =   495
            Index           =   1
            Left            =   8760
            TabIndex        =   9
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
            Height          =   495
            Index           =   2
            Left            =   7920
            TabIndex        =   10
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   3
            Left            =   6915
            TabIndex        =   11
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   4
            Left            =   5880
            TabIndex        =   12
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   6
            Left            =   3840
            TabIndex        =   13
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   5
            Left            =   4950
            TabIndex        =   14
            Top             =   510
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            TabIndex        =   7
            Top             =   225
            Width           =   1740
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
            TabIndex        =   6
            Top             =   240
            Width           =   1515
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7830
         Index           =   2
         Left            =   30
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         Width           =   14505
         _cx             =   25585
         _cy             =   13811
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   765
            Index           =   5
            Left            =   0
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   14355
            _cx             =   25321
            _cy             =   1349
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
            BackColor       =   16777215
            ForeColor       =   4210688
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "„ «»⁄Â «·⁄„·Ì«   "
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
               TabIndex        =   17
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
               ButtonImage     =   "û.frmOperationsFollow.frx":0B5F
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
               TabIndex        =   18
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
               ButtonImage     =   "û.frmOperationsFollow.frx":0EF9
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
               TabIndex        =   19
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
               ButtonImage     =   "û.frmOperationsFollow.frx":1293
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
               TabIndex        =   20
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
               ButtonImage     =   "û.frmOperationsFollow.frx":162D
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Left            =   6480
               Picture         =   "û.frmOperationsFollow.frx":19C7
               Stretch         =   -1  'True
               Top             =   120
               Width           =   525
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   2955
            Index           =   1
            Left            =   0
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   14385
            _cx             =   25374
            _cy             =   5212
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
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Height          =   225
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   495
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               Height          =   465
               Index           =   0
               Left            =   -3720
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   8430
               Width           =   2070
            End
            Begin VB.TextBox xptxtid 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   10395
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   750
               Width           =   2070
            End
            Begin VB.TextBox TxtChalenges 
               Alignment       =   1  'Right Justify
               Height          =   600
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   1245
               Width           =   4125
            End
            Begin VB.TextBox TxtPorposal 
               Alignment       =   1  'Right Justify
               Height          =   915
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Top             =   1920
               Width           =   4125
            End
            Begin VB.TextBox TxtGoal 
               Alignment       =   1  'Right Justify
               Height          =   585
               Left            =   6645
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   1560
               Width           =   5820
            End
            Begin VB.TextBox TxtNotes 
               Alignment       =   1  'Right Justify
               Height          =   600
               Left            =   6645
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   2250
               Width           =   5820
            End
            Begin MSDataListLib.DataCombo dcproject 
               Height          =   330
               Left            =   6675
               TabIndex        =   26
               Top             =   1245
               Width           =   5790
               _ExtentX        =   10213
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker XPDtbTrans 
               Height          =   330
               Left            =   6645
               TabIndex        =   30
               Top             =   750
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   582
               _Version        =   393216
               Format          =   94437377
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DcEmp 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   765
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   210
               Left            =   13305
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   1080
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„ «»⁄Â"
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
               Index           =   7
               Left            =   13080
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·„ «»⁄Â"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   8460
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   720
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Âœ› „‰ «·„ «»⁄Â"
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
               Index           =   0
               Left            =   13065
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   1590
               Width           =   1200
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁ«∆„ »«·„ «»⁄Â"
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
               Index           =   2
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„·«ÕŸ« "
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
               Index           =   3
               Left            =   13305
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   2160
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — «·„‘—Ê⁄"
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
               Index           =   4
               Left            =   13305
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   1245
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„⁄Êﬁ« "
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
               Index           =   6
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   1335
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«ﬁ —«Õ«  Ê «· Ê’Ì« "
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
               Index           =   10
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   2010
               Width           =   1575
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   2100
            Left            =   120
            TabIndex        =   41
            Top             =   5025
            Width           =   12945
            _cx             =   22834
            _cy             =   3704
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
            Rows            =   50
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"û.frmOperationsFollow.frx":562F
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
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   1620
            Left            =   120
            TabIndex        =   42
            Top             =   3240
            Width           =   12945
            _cx             =   22834
            _cy             =   2857
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
            Rows            =   50
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"û.frmOperationsFollow.frx":5931
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
            Left            =   11040
            TabIndex        =   43
            Tag             =   "Delete Row"
            Top             =   7320
            Width           =   1935
            _ExtentX        =   3413
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "û.frmOperationsFollow.frx":5C3C
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
            TabIndex        =   46
            Top             =   90
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«‰Ã«“« "
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
            Index           =   9
            Left            =   13185
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   5040
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„·Ì«  ·Õ«·Ì…"
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
            Index           =   11
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   3240
            Width           =   1080
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
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
      ButtonImage     =   "û.frmOperationsFollow.frx":5C58
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmOperationsFollow"
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
Dim rs As ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Public Sub YearMonth()

End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "

    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = Null
    rs("Remark").value = Msg

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = user_id
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Function Create_dev1()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    'StrAccountCode = Account_Code_dynamic
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With

    Set rs = New ADODB.Recordset
    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
 
    rs("voucher_id").value = LngDevID
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    dcproject.Text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub



Private Sub CboPayMentType_Click()
     
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

End Function

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub

Private Sub Combo1_Click()
 
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
        If Trim(Me.Dcemp.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·ﬁ«∆„ »«·„ «»⁄Â..!!"
            Else
                Msg = "Select  Follow Up Employee"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Dcemp.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        If Trim(Me.dcproject.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·„‘—Ê⁄..!!"
            Else
                Msg = "Select Project Frstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcproject.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete OperationsFollowDetails where OperationsFollowId=" & val(Me.xptxtid.Text)
   
    End If
    
    rs("ID").value = xptxtid.Text
   
    rs("RecordDate").value = XPDtbTrans.value
    rs("EmpID").value = IIf(Me.Dcemp.BoundText = "", Null, Me.Dcemp.BoundText)
    rs("ProjectID").value = IIf(Me.dcproject.BoundText = "", Null, Me.dcproject.BoundText)
 
    rs("Goal").value = IIf(Me.TxtGoal.Text = "", Null, Me.TxtGoal.Text)
    rs("Remarks").value = IIf(Me.TxtNotes.Text = "", Null, Me.TxtNotes.Text)
    rs("challenges").value = IIf(Me.TxtChalenges.Text = "", Null, Me.TxtChalenges.Text)
    rs("Proposals").value = IIf(Me.TxtPorposal.Text = "", Null, Me.TxtPorposal.Text)
     
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "OperationsFollowDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("FullCode")) <> "" Then
         
                RsDev.AddNew
                RsDev("OperationsFollowId").value = val(Me.xptxtid.Text)
        
                RsDev("opr_Fullcode").value = .TextMatrix(i, .ColIndex("FullCode"))
                RsDev("opr_Name").value = .TextMatrix(i, .ColIndex("Name"))
              
                RsDev("Achievements").value = .TextMatrix(i, .ColIndex("Achievements"))
                RsDev("variance").value = val(.TextMatrix(i, .ColIndex("variance")))
                RsDev("varianceReasons").value = .TextMatrix(i, .ColIndex("varianceReason"))
 
                RsDev("StatusID").value = val(.TextMatrix(i, .ColIndex("StatusID")))
                RsDev("Status").value = .TextMatrix(i, .ColIndex("Status"))
             
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Saved " & CHR(13)
                Msg = Msg + "Do you want to add other data"
         
            End If

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Amendments have been saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.Text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Me.xptxtid.Text = CStr(new_id("OperationsFollow", "ID", "", True))
       
            XPDtbTrans.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
            Grid.Enabled = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            ' Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If Grid.Rows > 1 Then
        If Grid.Rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.Rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub

Private Sub dcopr_Click(Area As Integer)

End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
                  
        My_SQL = " select  Emp_ID,Emp_Name from TblEmployee"
        fill_combo Dcemp, My_SQL
    End If

End Sub

Private Sub dcproject_Change()
    dcproject_Click (0)
End Sub

Private Sub dcproject_Click(Area As Integer)
Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 2
            Grid1.Enabled = True
    If dcproject.BoundText = "" Then Exit Sub
    getInformation val(dcproject.BoundText)
    

End Sub

Function getInformation(project_id As Integer)
    Dim i As Integer
    Dim My_SQL As String
    Dim RsDev As ADODB.Recordset
    My_SQL = " select  *  from terms_operations where  ended =0 and Project_ID=" & project_id
     
    Set RsDev = New ADODB.Recordset
    RsDev.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsDev.RecordCount > 0 Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Symbol")) = IIf(IsNull(RsDev("Symbol").value), "", RsDev("Symbol").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), 0, RsDev("name").value)
                .TextMatrix(i, .ColIndex("EarlyStartDate")) = IIf(IsNull(RsDev("EarlyStartDate").value), "", RsDev("EarlyStartDate").value)
             
                .TextMatrix(i, .ColIndex("EarlyEndDate")) = IIf(IsNull(RsDev("EarlyEndDate").value), "", RsDev("EarlyEndDate").value)
                .TextMatrix(i, .ColIndex("start_date")) = IIf(IsNull(RsDev("start_date").value), "", RsDev("start_date").value)
                .TextMatrix(i, .ColIndex("end_date")) = IIf(IsNull(RsDev("end_date").value), "", RsDev("end_date").value)
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    ReLineGrid

End Function

Private Sub dcproject_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
            
        My_SQL = " select id,Project_name from projects"
        fill_combo dcproject, My_SQL
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
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
 
    Dim My_SQL As String

'    My_SQL = " select id,Project_name from projects"
'    fill_combo dcproject, My_SQL
'
    My_SQL = " select  Emp_ID,Emp_Name from TblEmployee"
    fill_combo Dcemp, My_SQL

    'My_SQL = " select  Emp_ID,Emp_Name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Dcombos.GetProjects dcproject
    Set BKGrndPic = New ClsBackGroundPic

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From OperationsFollow  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    lbl(11).Caption = "Current Proesses"
    CmdRemove.Caption = "Remove Line"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Processes Follow up "
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Start Date"
    'Ele(3).Caption = "Select Interval"
    lbl(2).Caption = "By"
    lbl(4).Caption = "Project"

    'lbl(5).Caption = "process"
    lbl(6).Caption = "challenges"
    lbl(0).Caption = "Goal"
    lbl(10).Caption = "Proposals"
    lbl(3).Caption = "Notes"
    lbl(9).Caption = "Achievements"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"

        .TextMatrix(0, .ColIndex("Name")) = "Process"
        .TextMatrix(0, .ColIndex("Achievements")) = "Achievements"
        .TextMatrix(0, .ColIndex("variance")) = "Variance"
        .TextMatrix(0, .ColIndex("varianceReason")) = "Variance Reason"
        .TextMatrix(0, .ColIndex("Status")) = "Status"
        .TextMatrix(0, .ColIndex("Name")) = "Process"

    End With

    With Me.Grid1
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Symbol")) = "Symbol"
        .TextMatrix(0, .ColIndex("name")) = "name"
        .TextMatrix(0, .ColIndex("EarlyStartDate")) = "Early Start Date"
        .TextMatrix(0, .ColIndex("EarlyEndDate")) = "EarlyEndDate"
        .TextMatrix(0, .ColIndex("start_date")) = "start_date"
        .TextMatrix(0, .ColIndex("end_date")) = "end_date"
    End With

End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,ProjectID
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("ProjectID").value), "", rs.Fields("ProjectID").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

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

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
 
            Case "Name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("FullCode"), False, True)
                .TextMatrix(Row, .ColIndex("FullCode")) = StrAccountCode
             
            Case "Status"
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("StatusID"), False, True)
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("StatusID")) = StrAccountCode
 
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

    With Me.Grid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Symbol")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid
     
        If .ColKey(Col) <> "Name" Or .ColKey(Col) <> "Status" Then
            .ComboList = ""
        End If

        '  Cancel = True
     
    End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
    With Me.Grid

        Select Case .ColKey(Col)

            Case "Name"
        
                'Full Path Display
                StrSQL = "SELECT *  FROM terms_operations  where project_id=" & val(Me.dcproject.BoundText)

                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "fullcode,Symbol,*name", "fullcode")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            
            Case "Status"
        
                'Full Path Display
                 
                StrComboList = "#1;1  Running|#2;2  Finished"

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            
        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.xptxtid.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Me.Dcemp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    dcproject.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
 
    Me.TxtChalenges.Text = IIf(IsNull(rs("Goal").value), "", rs("Goal").value)
    Me.TxtNotes.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    Me.TxtChalenges.Text = IIf(IsNull(rs("challenges").value), "", rs("challenges").value)
    Me.TxtPorposal.Text = IIf(IsNull(rs("Proposals").value), "", rs("Proposals").value)
 
    StrSQL = "select * from OperationsFollowDetails where OperationsFollowId=" & val(Me.xptxtid.Text)
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
            
                .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(RsDev("opr_Fullcode").value), "", RsDev("opr_Fullcode").value)
            
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsDev("opr_Name").value), "", RsDev("opr_Name").value)
            
                .TextMatrix(i, .ColIndex("Achievements")) = IIf(IsNull(RsDev("Achievements").value), "", RsDev("Achievements").value)
                .TextMatrix(i, .ColIndex("variance")) = IIf(IsNull(RsDev("variance").value), 0, RsDev("variance").value)
                .TextMatrix(i, .ColIndex("varianceReason")) = IIf(IsNull(RsDev("varianceReasons").value), "", RsDev("varianceReasons").value)
            
                .TextMatrix(i, .ColIndex("StatusID")) = IIf(IsNull(RsDev("StatusID").value), 0, RsDev("StatusID").value)
                .TextMatrix(i, .ColIndex("Status")) = IIf(IsNull(RsDev("Status").value), "", RsDev("Status").value)
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
 
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

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
