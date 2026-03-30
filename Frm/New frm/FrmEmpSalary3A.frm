VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmpSalary3A 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘…  Œ’Ì’  „⁄œ«  ⁄·Ì „‘—Ê⁄ „⁄Ì‰"
   ClientHeight    =   9795
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14760
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
   Icon            =   "FrmEmpSalary3A.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   14760
   WindowState     =   2  'Maximized
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      ButtonImage     =   "FrmEmpSalary3A.frx":038A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9795
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14760
      _cx             =   26035
      _cy             =   17277
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
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   840
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8925
         Width           =   14700
         _cx             =   25929
         _cy             =   1482
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
            ButtonImage     =   "FrmEmpSalary3A.frx":0724
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
            ButtonImage     =   "FrmEmpSalary3A.frx":0ABE
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
            ButtonImage     =   "FrmEmpSalary3A.frx":0E58
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   10860
            TabIndex        =   6
            Top             =   390
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
            Left            =   9960
            TabIndex        =   7
            Top             =   390
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
            Left            =   9120
            TabIndex        =   8
            Top             =   390
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
            Left            =   8115
            TabIndex        =   9
            Top             =   390
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
            Left            =   7080
            TabIndex        =   10
            Top             =   390
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
            Left            =   3120
            TabIndex        =   11
            Top             =   360
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
            Left            =   6150
            TabIndex        =   12
            Top             =   390
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   9120
            TabIndex        =   13
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmEmpSalary3A.frx":11F2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmdpint 
            Height          =   495
            Left            =   5160
            TabIndex        =   14
            Top             =   360
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   495
            Left            =   3960
            TabIndex        =   15
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·…"
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   225
            Width           =   1740
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   9240
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   14700
         _cx             =   25929
         _cy             =   16298
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "."
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8820
            Index           =   2
            Left            =   45
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   45
            Width           =   14610
            _cx             =   25770
            _cy             =   15558
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   5
               Left            =   0
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   0
               Width           =   14715
               _cx             =   25956
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
               Picture         =   "FrmEmpSalary3A.frx":120E
               Caption         =   "‘«‘…  Œ’Ì’  „⁄œ«  ⁄·Ï „‘—Ê⁄ „⁄Ì‰ "
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
                  TabIndex        =   21
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
                  ButtonImage     =   "FrmEmpSalary3A.frx":1EE8
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
                  TabIndex        =   22
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
                  ButtonImage     =   "FrmEmpSalary3A.frx":2282
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
                  TabIndex        =   23
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
                  ButtonImage     =   "FrmEmpSalary3A.frx":261C
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
                  TabIndex        =   24
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
                  ButtonImage     =   "FrmEmpSalary3A.frx":29B6
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   10275
               Index           =   1
               Left            =   0
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   0
               Width           =   15225
               _cx             =   26855
               _cy             =   18124
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
                  Height          =   330
                  Left            =   5835
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   495
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   9705
                  Width           =   2175
               End
               Begin VB.TextBox xptxtid 
                  Alignment       =   1  'Right Justify
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
                  Left            =   10995
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   870
                  Width           =   2175
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Text            =   "0"
                  Top             =   750
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   26
                  Top             =   870
                  Width           =   6255
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   5865
                  Left            =   0
                  TabIndex        =   31
                  Top             =   2640
                  Width           =   14625
                  _cx             =   25797
                  _cy             =   10345
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
                  Cols            =   27
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEmpSalary3A.frx":2D50
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
               Begin MSComCtl2.DTPicker XPDtbTrans 
                  Height          =   330
                  Left            =   7395
                  TabIndex        =   32
                  Top             =   870
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   582
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   61734913
                  CurrentDate     =   38784
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic6 
                  Height          =   1380
                  Left            =   240
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   1290
                  Width           =   14445
                  _cx             =   25479
                  _cy             =   2434
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
                  Begin VB.TextBox TxtPrice 
                     Alignment       =   1  'Right Justify
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
                     Left            =   10710
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   570
                     Width           =   2175
                  End
                  Begin VB.TextBox TxtPeriod 
                     Alignment       =   1  'Right Justify
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
                     Height          =   330
                     Left            =   8040
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   570
                     Width           =   2175
                  End
                  Begin ImpulseButton.ISButton ISButton2 
                     Height          =   1080
                     Left            =   120
                     TabIndex        =   36
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   180
                     Width           =   1080
                     _ExtentX        =   1905
                     _ExtentY        =   1905
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
                     ButtonImage     =   "FrmEmpSalary3A.frx":313A
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin MSDataListLib.DataCombo DCPROJECT1 
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   37
                     Top             =   570
                     Width           =   4845
                     _ExtentX        =   8546
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
                  Begin MSDataListLib.DataCombo dcopr 
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   38
                     Top             =   975
                     Width           =   4845
                     _ExtentX        =   8546
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
                  Begin MSComCtl2.DTPicker FromDate 
                     Height          =   330
                     Left            =   4110
                     TabIndex        =   39
                     Top             =   120
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   582
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   61734913
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DcbEqpID 
                     Height          =   315
                     Left            =   8040
                     TabIndex        =   40
                     Top             =   120
                     Width           =   4845
                     _ExtentX        =   8546
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
                  Begin MSDataListLib.DataCombo Dcterm1 
                     Height          =   315
                     Left            =   8040
                     TabIndex        =   41
                     Top             =   975
                     Width           =   4845
                     _ExtentX        =   8546
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
                  Begin MSComCtl2.DTPicker ToDate 
                     Height          =   330
                     Left            =   1440
                     TabIndex        =   42
                     Top             =   120
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   582
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   61734913
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «·„‘—Ê⁄"
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
                     Left            =   6450
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   570
                     Width           =   1305
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·»‰œ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   0
                     Left            =   13005
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   975
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄„·Ì…"
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
                     Index           =   4
                     Left            =   6450
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   975
                     Width           =   1305
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰  «—ÌŒ"
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
                     Left            =   6450
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   105
                     Width           =   1305
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «·„⁄œ…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   3
                     Left            =   13005
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   120
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
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
                     Index           =   6
                     Left            =   3660
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   105
                     Width           =   345
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
                     Height          =   345
                     Index           =   9
                     Left            =   13005
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   570
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„œ…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   10
                     Left            =   9930
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   570
                     Width           =   720
                  End
               End
               Begin ImpulseButton.ISButton CMdDeleted 
                  Height          =   225
                  Left            =   13305
                  TabIndex        =   51
                  Top             =   8520
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   397
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ— "
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
                  ButtonImage     =   "FrmEmpSalary3A.frx":999C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdDeleteAll 
                  Height          =   225
                  Left            =   12000
                  TabIndex        =   52
                  Top             =   8520
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   397
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
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
                  ButtonImage     =   "FrmEmpSalary3A.frx":9F36
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   375
                  Left            =   13800
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   1005
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· Œ’Ì’"
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
                  Index           =   7
                  Left            =   12780
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   870
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
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
                  Index           =   8
                  Left            =   9045
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   870
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
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
                  Height          =   285
                  Index           =   11
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   870
                  Width           =   825
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
               TabIndex        =   57
               Top             =   90
               Width           =   1125
            End
         End
      End
   End
End
Attribute VB_Name = "FrmEmpSalary3A"
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



Private Sub CmdDeleteAll_Click()
If Me.TxtModFlg.Text <> "R" Then
Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
End If
End Sub
Private Sub RemoveGridRow()
    With Me.Grid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub

Private Sub CMdDeleted_Click()
If Me.TxtModFlg.Text <> "R" Then
RemoveGridRow
End If
End Sub

Private Sub Cmdpint_Click()
print_report
End Sub

Private Sub DcbEqpID_Click(Area As Integer)
DiffDtaeVlue
End Sub

Private Sub dcproject1_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(DCPROJECT1.BoundText) <> 0 Then
fillterms val(DCPROJECT1.BoundText)
End If
End If
End Sub
Function fillterms(project_id As Integer)
If Me.TxtModFlg.Text <> "R" Then
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

    fill_combo Me.Dcterm1, My_SQL
     
    Dcterm1.ReFill
    End If
End Function

Private Sub dcproject1_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
    If KeyCode = vbKeyF5 Then
      '  Dim My_SQL As String
            
      '  My_SQL = " select id,Project_name from projects"
      '  fill_combo dcproject1, My_SQL
        Dim Dcombos As ClsDataCombos
        Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Dcombos.GetProjects DCPROJECT1
    End If
    End If
End Sub

Private Sub Dcterm1_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
 Dim Dcombos As ClsDataCombos

       Set Dcombos = New ClsDataCombos
  If DCPROJECT1.BoundText <> "" Then
        
         If Me.Dcterm1.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt dcopr, val(DCPROJECT1.BoundText), , val(Dcterm1.BoundText), 2
         End If
       
    End If
    End If
End Sub

'Private Sub ChkDetails_Click()
'    FillGridWithData
'End Sub

'Private Sub ALLButton1_Click()
'    FrmShowCol1.show
'End Sub

'Function check_previous_dev(year As String, Month As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select * from notes where salary=" & year & Month
'
'    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If rs.RecordCount = 0 Then
'        check_previous_dev = False
'    Else
'        check_previous_dev = True
'    End If
'
'End Function

'Function check_previous_dev1(year As String, Month As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
'
'    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If rs.RecordCount = 0 Then
'        check_previous_dev1 = False
'    Else
'        check_previous_dev1 = True
'    End If
'
'End Function
'
'Function Create_dev()
'    Dim i As Integer
'    Dim LngDevID As Long
'    Dim Msg As String
''    Dim Account_Code_dynamic As String
'    Dim Account_Code_dynamic1 As String
'
'    Dim Employee_account As String
'    Dim StrAccountCode As String
'    Dim X As Integer
'    Dim rs As ADODB.Recordset
'    Dim notes_serial As String
'    Dim notes_id As String
'
'    Account_Code_dynamic = get_account_code_branch(16, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            GoTo ErrTrap
'
'        End If
'    End If
'
'    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "
'
'    Dim StrSQL As String
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From Notes where NoteType=66 order by NoteID"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    notes_id = CStr(new_id("Notes", "NoteID", "", True))
'    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
'
'    rs.AddNew
'    rs("NoteID").value = notes_id
'    rs("NoteSerial").value = notes_serial '
''    rs("Note_Value").value = Null
 '   rs("Remark").value = Msg
'
''    rs("NoteType").value = 66
 '   rs("NoteDate").value = Date
 '   rs("UserID").value = user_id
 '   rs.update
 '
 '   LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
 ''
  '  Dim line_no As Integer
  '  line_no = 1
'
'    With Grid
'
'        For i = .FixedRows To .Rows - 2
'
'            If .TextMatrix(i, .ColIndex("project")) = "0" Then
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'
'            Else
'                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
'            StrAccountCode = Employee_account
'
'            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
'                GoTo ErrTrap
'            End If
'
'            line_no = line_no + 2
'
'        Next i
'
'    End With
'
'    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
'    create_report_data
'
'    DoEvents
'
'    Exit Function
'ErrTrap:
'    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
'
'End Function
'
'Function Create_dev1()
'    Dim i As Integer
'    Dim LngDevID As Long
'    Dim Msg As String
'    Dim Account_Code_dynamic As String
'    Dim Account_Code_dynamic1 As String
'
'    Dim Employee_account As String
'    Dim StrAccountCode As String
'    Dim X As Integer
'    Dim rs As ADODB.Recordset
'
'    Account_Code_dynamic = get_account_code_branch(16, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            GoTo ErrTrap
'
'        End If
'    End If
'
'    'StrAccountCode = Account_Code_dynamic
'
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'    Dim line_no As Integer
'    line_no = 1
'
'    With Grid
'
'        For i = .FixedRows To .Rows - 2
'
'            If .TextMatrix(i, .ColIndex("project")) = "0" Then
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'
''            Else
 '               Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
'            StrAccountCode = Employee_account
'
'            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
'                GoTo ErrTrap
'            End If
'
'            line_no = line_no + 2
'
'        Next i
'
'    End With
'
'    Set rs = New ADODB.Recordset
'    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'    rs.AddNew
'
'    rs("voucher_id").value = LngDevID
'
'    rs.update
'
'    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
'    create_report_data
'
'    DoEvents
'
'    Exit Function
'ErrTrap:
'    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
'
'End Function

'Private Sub ALLButton2_Click()
'    'Dcemp.text = ""
'
'    dcproject.text = ""
'    FillGridWithData
'
'    DoEvents
'    Create_dev
'    CmdOk_Click
'End Sub



'Private Sub CboPayMentType_Click()
'    CboPayMentType_Change
'End Sub

'Private Sub CboYear_Click()
'    CmdOk_Click
'End Sub

'Private Sub Check1_Click()
'Exit Sub
'    If Check1.value = vbChecked Then
'        get_all_employee
'    Else
'
''        With Me.Grid
 '           .Rows = 2
 '           .Clear flexClearScrollable
 '       End With
'
'    End If
''
'End Sub

'Private Sub CmbMonth_Click()
'    CmdOk_Click
    'FillGridWithData
'End Sub

'Private Sub CmdExit_Click()
'    Unload Me
'End Sub



'Private Sub CmdPrint_Click()
'    On Error Resume Next
'    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
'    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
'    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

'    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
'End Sub



Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 
        If Trim(Me.DCPROJECT1.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·„‘—Ê⁄..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCPROJECT1.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
Me.xptxtid.Text = CStr(new_id("TblSpecificFixed", "ID", "", True))
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblSpecificFixedDeti where SPFixID=" & val(Me.xptxtid.Text)
   
    End If
    
    rs("ID").value = xptxtid.Text
    rs("RecordDtae").value = XPDtbTrans.value
    rs("ProjectID").value = IIf(Me.DCPROJECT1.BoundText = "", Null, val(Me.DCPROJECT1.BoundText))
    rs("EqpID").value = IIf(Me.DcbEqpID.BoundText = "", Null, val(Me.DcbEqpID.BoundText))
  '  rs("opr_type").value = IIf(Me.txtType.text = "", 0, Me.txtType.text)
    rs("Fromdate").value = FromDate.value
    rs("ToDate").value = ToDate.value
    rs("Period").value = val(TxtPeriod.Text)
    rs("Price").value = val(TxtPrice.Text)
    rs("Remarks").value = TxtRemarks.Text
    If Me.Dcterm1.BoundText <> "" Then
        rs("PandID").value = IIf(Me.Dcterm1.BoundText = "", Null, Me.Dcterm1.BoundText)
    End If
     
    If Me.dcopr.BoundText <> "" Then
        rs("OperID").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
    End If

    rs.Update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblSpecificFixedDeti", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim I As Integer

    With Me.Grid

        For I = .FixedRows To .Rows - 1

            If val(.TextMatrix(I, .ColIndex("Emp_id"))) <> 0 Then
         
                RsDev.AddNew
                RsDev("SPFixID").value = Me.xptxtid.Text
                RsDev("FixedID").value = val(.TextMatrix(I, .ColIndex("Emp_id")))
                RsDev("LngT").value = val(.TextMatrix(I, .ColIndex("interval")))
                RsDev("Price").value = val(.TextMatrix(I, .ColIndex("price")))
                RsDev("Total").value = val(.TextMatrix(I, .ColIndex("Total")))
                RsDev("ProjectID").value = val(.TextMatrix(I, .ColIndex("ProjectID")))
                RsDev("PandID").value = val(.TextMatrix(I, .ColIndex("PandID")))
                RsDev("OperID").value = val(.TextMatrix(I, .ColIndex("OperID")))
                RsDev("FromDate").value = IIf(.TextMatrix(I, .ColIndex("FromDate")) = "", Null, .TextMatrix(I, .ColIndex("FromDate")))
                RsDev("ToDate").value = IIf(.TextMatrix(I, .ColIndex("ToDate")) = "", Null, .TextMatrix(I, .ColIndex("ToDate")))
              '  RsDev("Project_id").value = IIf(DCPROJECT1.BoundText = "", Null, Me.DCPROJECT1.BoundText)
'                RsDev("opr_type").value = IIf(Me.txtType.text = "", 0, Me.txtType.text)
'
               ' If Me.Dcterm1.BoundText <> "" Then
               '     RsDev("term_Fullcode").value = IIf(Me.Dcterm1.BoundText = "", Null, Me.Dcterm1.BoundText)
               ' End If
            '
            '    If Me.dcopr.BoundText <> "" Then
            '        RsDev("opr_Fullcode").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
            '    End If
       '
       '         save_employee_current_status DCPROJECT1.BoundText, Me.Dcterm1.BoundText, Me.dcopr.BoundText, val(.TextMatrix(i, .ColIndex("Emp_id")))
                RsDev.Update
                    
            End If
            
            '
        Next I

    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.Text = "N"
            clear_all Me
            
       
            XPDtbTrans.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Grid.Enabled = True

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

           ' If ChKauto.value = vbChecked Then
           '     If SystemOptions.UserInterface = ArabicInterface Then
           '         MsgBox " ·« Ì„ﬂ‰  ⁄œÌ·  Œ’Ì’ «·Ì ", vbCritical
           '     Else
           '         MsgBox " Can't Delete Auto Employee Allocation ", vbCritical
           '     End If
'
'                Exit Sub
'            End If

            TxtModFlg.Text = "E"
           ' Grid.Rows = Grid.Rows + 1
           Grid.Enabled = True

        Case 2
 
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

             Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

'            Load FrmNotesSearch
'            FrmNotesSearch.SearchType = 3
'            FrmNotesSearch.show vbModal

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
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If xptxtid.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblSpecificFixed Where id=" & val(Me.xptxtid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                 Cn.Execute "delete TblSpecificFixedDeti where SPFixID=" & val(Me.xptxtid)

                If rs.RecordCount < 1 Then
                Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
                    clear_all Me
                    TxtModFlg_Change
                   ' XPTxtCurrent.Caption = 0
                   ' XPTxtCount.Caption = 0
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub
'Private Sub Dcdep_Click(Area As Integer)
'    CmdOk_Click
'End Sub

'Private Sub Dcedara_Click(Area As Integer)
'    CmdOk_Click
'End Sub

'Private Sub Dcemp_Click(Area As Integer)
'    CmdOk_Click
'End Sub

'Private Sub DCmboEmp_Click(Area As Integer)
'    FillGridWithData
'End Sub

'Function SHow_grig_col()
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    With Grid
''
 '       If rs2("s1").value = True Then
 '           .ColHidden(.ColIndex("Emp_Code")) = False
 '       Else
 '           .ColHidden(.ColIndex("Emp_Code")) = True
 '       End If
 '
 '       If rs2("s2").value = True Then
 '           .ColHidden(.ColIndex("Emp_Name")) = False
 '       Else
 '           .ColHidden(.ColIndex("Emp_Name")) = True
 '       End If
 '
 '       If rs2("s3").value = True Then
 '           .ColHidden(.ColIndex("Emp_Salary")) = False
 '       Else
 ''           .ColHidden(.ColIndex("Emp_Salary")) = True
  '      End If
  '
  '      If rs2("s4").value = True Then
  '          .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
  '      Else
  '          .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
  '      End If
  ''
   '     If rs2("s5").value = True Then
   '         .ColHidden(.ColIndex("Emp_Salary_bus")) = False
   '     Else
   ''         .ColHidden(.ColIndex("Emp_Salary_bus")) = True
    '    End If
    '
    '    If rs2("s6").value = True Then
    '        .ColHidden(.ColIndex("Emp_Salary_food")) = False
    '    Else
    '        .ColHidden(.ColIndex("Emp_Salary_food")) = True
    '    End If
    '
    '    If rs2("s7").value = True Then
    '        .ColHidden(.ColIndex("Emp_Salary_mob")) = False
    '    Else
    '        .ColHidden(.ColIndex("Emp_Salary_mob")) = True
    '    End If
    '
    '    If rs2("s8").value = True Then
    '        .ColHidden(.ColIndex("Emp_Salary_mang")) = False
    '    Else
    ''        .ColHidden(.ColIndex("Emp_Salary_mang")) = True
     '   End If
     '
     '   If rs2("s9").value = True Then
     '       .ColHidden(.ColIndex("Emp_Salary_others")) = False
     '   Else
     '       .ColHidden(.ColIndex("Emp_Salary_others")) = True
     '   End If
     '
     '   If rs2("s10").value = True Then
     '       .ColHidden(.ColIndex("OverTimePrice")) = False
     '   Else
     '       .ColHidden(.ColIndex("OverTimePrice")) = True
     '   End If
     ''
      '  If rs2("s11").value = True Then
      '      .ColHidden(.ColIndex("Mokafea")) = False
      '  Else
      '      .ColHidden(.ColIndex("Mokafea")) = True
      ''  End If
       '
       ' If rs2("s12").value = True Then
       '     .ColHidden(.ColIndex("SalesCom")) = False
       ' Else
       '     .ColHidden(.ColIndex("SalesCom")) = True
       ' End If
       ''
        'If rs2("s13").value = True Then
        '    .ColHidden(.ColIndex("total1")) = False
        'Else
        '    .ColHidden(.ColIndex("total1")) = True
        'End If
        '
        'If rs2("s14").value = True Then
        ''    .ColHidden(.ColIndex("TotalAdvance")) = False
        'Else
         '   .ColHidden(.ColIndex("TotalAdvance")) = True
        'End If
         '
        'if rs2("s15").value = True Then
         '   .ColHidden(.ColIndex("TotalDiscount")) = False
        'Else
         '   .ColHidden(.ColIndex("TotalDiscount")) = True
        'End If
         '
        'If rs2("s16").value = True Then
        '    .ColHidden(.ColIndex("total2")) = False
        'Else
        '    .ColHidden(.ColIndex("total2")) = True
        'End If
                 
        'If rs2("s17").value = True Then
        '    .ColHidden(.ColIndex("EmpTotalNet")) = False
        'Else
        '    .ColHidden(.ColIndex("EmpTotalNet")) = True
        'End If
                  
        'If rs2("s18").value = True Then
        '    .ColHidden(.ColIndex("sgn")) = False
        'Else
        '    .ColHidden(.ColIndex("sgn")) = True
        'End If
     '
    'End With

'End Function

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
Sub DiffDtaeVlue()
If Me.TxtModFlg.Text <> "R" Then
Me.TxtPeriod.Text = DateDiff("d", FromDate.value, ToDate.value) + 1
End If
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim StrSQL As String
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture

    Dim My_SQL2 As String
                 If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
        fill_combo Me.DcbEqpID, StrSQL
   If SystemOptions.UserInterface = ArabicInterface Then
     My_SQL2 = " select id,Project_name from projects "
     My_SQL2 = My_SQL2 & " where Not(Project_name is null)and Project_name <>N'""' "
     My_SQL2 = My_SQL2 & " order by Project_name"
     Else
     My_SQL2 = " select id,Project_nameE from projects"
     My_SQL2 = My_SQL2 & " where Not(Project_nameE is null) and Project_nameE <>N'""'"
     My_SQL2 = My_SQL2 & " order by Project_nameE"
  End If
    fill_combo DCPROJECT1, My_SQL2

    My_SQL2 = " select  oprid,des from projects_des"
    fill_combo Dcterm1, My_SQL2
    If SystemOptions.UserInterface = ArabicInterface Then
     My_SQL2 = " SELECT     dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessName"
     My_SQL2 = My_SQL2 & "          FROM         dbo.terms_operations LEFT OUTER JOIN"
     My_SQL2 = My_SQL2 & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
Else
     My_SQL2 = " SELECT     dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessNameE"
     My_SQL2 = My_SQL2 & "          FROM         dbo.terms_operations LEFT OUTER JOIN"
     My_SQL2 = My_SQL2 & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
End If
    'My_SQL2 = " select  id,name from terms_operations"
    fill_combo dcopr, My_SQL2

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
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
    StrSQL = "select * From TblSpecificFixed  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
'    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
CMdDeleted.Caption = "Delete"
CmdDeleteAll.Caption = "Delete All"
ISButton2.Caption = "Add"
lbl(11).Caption = "Remarks"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Projects Equipments Allocate"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Date"
    lbl(2).Caption = "From Date"
    lbl(6).Caption = "To"
    lbl(9).Caption = "Price"
    lbl(3).Caption = "Equipments"
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
    lbl(0).Caption = "Terms"
    lbl(4).Caption = "Process"
    lbl(5).Caption = "Project"
    lbl(10).Caption = "Period"

'    Check1.Caption = "Show All Equipments"

    CmdRemove.Caption = "Remove Line"
Cmdpint.Caption = "Print"
ISButton3.Caption = "Same Copy"
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Emp_code")) = "Equipments code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Equipments Name"
        .TextMatrix(0, .ColIndex("interval")) = "Period"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        .TextMatrix(0, .ColIndex("FromDate")) = "From Date"
        .TextMatrix(0, .ColIndex("ToDate")) = "To Date"
        .TextMatrix(0, .ColIndex("Project_name")) = "Project"
        .TextMatrix(0, .ColIndex("des")) = "Terms"
        .TextMatrix(0, .ColIndex("ProcessName")) = "Process"

    End With

End Sub
'Public Sub get_all_employee()
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
''    Set rs2 = New ADODB.Recordset
'    Dim J As Integer
'
'    Dim sql As String
'    Dim i As Integer
'
'    sql = "Select * from emp_all_details "
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then Exit Sub
'
'    With Grid
'
'        .Rows = 2
'        .Clear flexClearScrollable
'
'        If Rs3.RecordCount > 0 Then
'            .Rows = Rs3.RecordCount + 1
'            Rs3.MoveFirst
'
'            For i = 1 To Rs3.RecordCount
'                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
'
'                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
'                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
'                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
'                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
'                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
'
'                Rs3.MoveNext
'            Next i
'
'            .AutoSize 0, .Cols - 1, False
'        End If
'
'    End With
'
'    Rs3.Close
'
'End Sub
''Public Sub get_all_employee()
''    Dim Rs3 As ADODB.Recordset
''    Set Rs3 = New ADODB.Recordset
''    Dim rs2 As ADODB.Recordset
''    Set rs2 = New ADODB.Recordset
'    Dim J As Integer
''
 '   Dim sql As String
 '   Dim i As Integer
'
''    sql = "Select * from emp_all_details "
 '
 '   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '
 '   If Rs3.RecordCount = 0 Then Exit Sub
 ''
  '  With Grid
'
'        .Rows = 2
''        .Clear flexClearScrollable
'
'        If Rs3.RecordCount > 0 Then
'            .Rows = Rs3.RecordCount + 1
''            Rs3.MoveFirst
 '
 '           For i = 1 To Rs3.RecordCount
 '               .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
 ''
  '              .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
  '              .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
  '              .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
  '              .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
  '              .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
  '
  '              Rs3.MoveNext
  '          Next i
 '
 '           .AutoSize 0, .Cols - 1, False
 '       End If
'
'    End With
 
'    Rs3.Close

'End Sub

'Public Sub FillGridWithData()
'    Exit Sub
'
''    Dim i As Integer
 '   Dim rs As ADODB.Recordset
 '   Dim rs2 As ADODB.Recordset
 '   Dim LstDay As Date
 '   Dim FrstDay As Date
 '   Dim StrTxt As String
 '   Dim My_SQL As String
 ''   Dim StrWhere As String
  '  Dim StrGrp As String
  '  Dim IntMonth As Integer
  '  Dim IntYear As Integer
  ''  Dim Msg As String
'
'    On Error GoTo ErrTrap
'
'    Set rs = New ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'
'    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
 '   With Me.Grid
 '       .Rows = 2
 '       .Clear flexClearScrollable
'
'        If rs.RecordCount > 0 Then
'            .Rows = rs.RecordCount + 1
'            rs.MoveFirst
'
'            For i = 1 To .Rows - 1
'
'                .TextMatrix(i, .ColIndex("Ser")) = i
'                ',DepartmentID,project_id
''
 '               .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
 '
 '               .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
 '
 '               .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
 ''
  '              .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
  '
  '              .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
  '
  '              .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
  '
  '              .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
  ''
   '             .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
   '
   '             '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
   '              "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
   '
   '             '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
   '             '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
   '
   '             rs.MoveNext
   '
   '         Next
'
'            rs.Close
'        End If
'
'        .Rows = .Rows + 1
'        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
'        .IsSubtotal(.Rows - 1) = True
'        Dim SngTotal As Single
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
'        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
'        net_value = SngTotal
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
'        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
'        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
'        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
'        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
'        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
'        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
'        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
'        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
'
'        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
'        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
'        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
'        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
'        .AutoSize 0, .Cols - 1, False
'    End With
''
'ErrTrap:
'End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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

Private Sub FromDate_Change()
DiffDtaeVlue
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
     
    Dim Rs1 As ADODB.Recordset
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim Sql As String
Set Rs1 = New ADODB.Recordset
    With Grid

        Select Case .ColKey(Col)
               Case "Project_name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ProjectID"), False, True)
                .TextMatrix(Row, .ColIndex("ProjectID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("OperID")) = 0
                .TextMatrix(Row, .ColIndex("PandID")) = 0
                .TextMatrix(Row, .ColIndex("des")) = ""
                .TextMatrix(Row, .ColIndex("ProcessName")) = ""
               Case "des"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PandID"), False, True)
                .TextMatrix(Row, .ColIndex("PandID")) = StrAccountCode
              Case "ProcessName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("OperID"), False, True)
                .TextMatrix(Row, .ColIndex("OperID")) = StrAccountCode
               Case "Emp_id"
               Set rs = New ADODB.Recordset
                StrSQL = "SELECT  * from FixedAssets Where id=" & val(.TextMatrix(Row, .ColIndex("Emp_id")))
            
               
                   rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
                    If Not (rs.BOF Or rs.EOF) Then
                      .TextMatrix(Row, .ColIndex("Emp_Code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
  
                            
                   End If
               
               
             Case "Emp_Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_id"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_id")) = StrAccountCode
             
                StrSQL = "SELECT  * from FixedAssets Where id=" & val(StrAccountCode)
                Set rs = Nothing
            
               If StrAccountCode <> "" Then
                   rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
                    If Not (rs.BOF Or rs.EOF) Then
                      .TextMatrix(Row, .ColIndex("Emp_Code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
  
                            
                   End If
               End If
            Sql = " select (UsedPowerPriceH+Hourdipp+UsedElectricPriceH) as HourVal from TblEquipments where fixedAssetid=" & val(.TextMatrix(Row, .ColIndex("Emp_id"))) & ""
Rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
 .TextMatrix(Row, .ColIndex("price")) = IIf(IsNull(Rs1("HourVal").value), "", Rs1("HourVal").value)
 End If
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "Emp_Code"
                  
'
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_id"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_id")) = StrAccountCode
             
                StrSQL = "SELECT  * from FixedAssets Where id=" & val(StrAccountCode)
                Set rs = Nothing
            
               If StrAccountCode <> "" Then
                   rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
                    If Not (rs.BOF Or rs.EOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                      .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                      Else
                       .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                       End If
  
                            
                   End If
               End If
Sql = " select (UsedPowerPriceH+Hourdipp+UsedElectricPriceH) as HourVal from TblEquipments where fixedAssetid=" & val(.TextMatrix(Row, .ColIndex("Emp_id"))) & ""
Rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
 .TextMatrix(Row, .ColIndex("price")) = IIf(IsNull(Rs1("HourVal").value), "", Rs1("HourVal").value)
 End If
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
    Dim I As Integer

    With Me.Grid

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("Emp_ID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
                .TextMatrix(I, .ColIndex("Total")) = val(.TextMatrix(I, .ColIndex("price"))) * val(.TextMatrix(I, .ColIndex("interval")))
  
            End If

        Next I
   
    End With

End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = "SELECT     dbo.TblSpecificFixed.ID, dbo.TblSpecificFixedDeti.LngT, dbo.TblSpecificFixedDeti.Price, dbo.TblSpecificFixedDeti.total, dbo.TblSpecificFixedDeti.FixedID, "
MySQL = MySQL & "                      dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.FixedAssets.Fullcode, dbo.TblSpecificFixed.ProjectID, dbo.projects.Project_name,"
MySQL = MySQL & "                      dbo.projects.Project_nameE, dbo.TblSpecificFixed.PandID, dbo.projects_des.des, dbo.TblSpecificFixed.OperID, dbo.TblSpecificFixed.RecordDtae,"
MySQL = MySQL & "                      dbo.TblProcessDEF.ProcessName , dbo.TblProcessDEF.ProcessNameE"
MySQL = MySQL & " FROM         dbo.TblProcessDEF RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblSpecificFixed ON dbo.TblProcessDEF.TblProcessDEFID = dbo.TblSpecificFixed.OperID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects_des ON dbo.TblSpecificFixed.PandID = dbo.projects_des.oprid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects ON dbo.TblSpecificFixed.ProjectID = dbo.projects.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblSpecificFixedDeti ON dbo.FixedAssets.id = dbo.TblSpecificFixedDeti.FixedID ON dbo.TblSpecificFixed.ID = dbo.TblSpecificFixedDeti.SPFixID"
MySQL = MySQL & " Where (dbo.TblSpecificFixed.id = " & val(xptxtid.Text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSpecificEquep.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSpecificEquep.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
    
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
Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        Select Case .ColKey(Col)

            Case "Emp_Code"
                .ComboList = ""

            Case "interval"
                Cancel = True
         Case "FromDate"
                Cancel = True
           Case "ToDate"
                Cancel = True
                
            Case "price"
                .ComboList = ""
        
            Case "Total"
                .ComboList = ""

            
        End Select

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
      Case "Project_name"
               StrSQL = "  SELECT     id, Project_name, Project_nameE"
               StrSQL = StrSQL & "     From dbo.Projects"
             If SystemOptions.UserInterface = ArabicInterface Then
              StrSQL = StrSQL & " Where (Not (Project_name Is Null)) and Project_name<>N'""'"
              Else
              StrSQL = StrSQL & " Where (Not (Project_nameE Is Null)) and Project_nameE <>N'""'"
             End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Project_name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "Project_nameE", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
         ''//////////////
              Case "des"
               StrSQL = " SELECT     oprid, des"
               StrSQL = StrSQL & "          From dbo.projects_des"
               StrSQL = StrSQL & "          Where (project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & " and project_id<>0)"
        
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               StrComboList = .BuildComboList(rs, "des", "oprid")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                

          Case "ProcessName"
               StrSQL = "       SELECT     dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
               StrSQL = StrSQL & "       FROM         dbo.terms_operations LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "   Where (dbo.terms_operations.project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & ") And (dbo.terms_operations.ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("PandID"))) & ")"
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                   StrComboList = .BuildComboList(rs, "ProcessName", "OPRIDD")
                Else
                   StrComboList = .BuildComboList(rs, "ProcessNameE", "OPRIDD")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
Case "Emp_Name"
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing

 
           Case "Emp_Code"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, code"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, code"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
                      Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

              
                    StrComboList = .BuildComboList(rs, "code", "id")
             

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

               .ComboList = StrComboList
                rs.Close
                Set rs = Nothing

'
        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim I As Integer

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
    XPDtbTrans.value = IIf(IsNull(rs("RecordDtae").value), Date, rs("RecordDtae").value)
    Me.DCPROJECT1.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
    Me.Dcterm1.BoundText = IIf(IsNull(rs("PandID").value), "", rs("PandID").value)
    Me.dcopr.BoundText = IIf(IsNull(rs("OperID").value), "", rs("OperID").value)
    FromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    ToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
    Me.TxtPeriod.Text = IIf(IsNull(rs("Period").value), "", rs("Period").value)
    Me.TxtPrice.Text = IIf(IsNull(rs("Price").value), "", rs("Price").value)
    Me.TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    Me.DcbEqpID.BoundText = IIf(IsNull(rs("EqpID").value), "", rs("EqpID").value)
   ' dcopr.BoundText = IIf(IsNull(rs("OperID").value), "", rs("OperID").value)
   ' MsgBox dcopr.BoundText

  '  txtType.text = IIf(IsNull(rs("opr_type").value), 0, rs("opr_type").value)

   ' If IsNull(rs("auto").value) Then
   '     ChKauto.value = vbUnchecked
   ' Else
   '     ChKauto.value = vbChecked
   ' End If
StrSQL = " SELECT     dbo.TblSpecificFixedDeti.ID, dbo.TblSpecificFixedDeti.SPFixID, dbo.TblSpecificFixedDeti.LngT, dbo.TblSpecificFixedDeti.Price, dbo.TblSpecificFixedDeti.total, "
StrSQL = StrSQL & "                       dbo.TblSpecificFixedDeti.FixedID, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.TblSpecificFixedDeti.ToDate,"
StrSQL = StrSQL & "                       dbo.TblSpecificFixedDeti.FromDate, dbo.TblSpecificFixedDeti.ProjectID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblSpecificFixedDeti.PandID,"
StrSQL = StrSQL & "                       dbo.projects_des.des , dbo.TblSpecificFixedDeti.OperID, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
StrSQL = StrSQL & "  FROM         dbo.TblSpecificFixedDeti LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblProcessDEF ON dbo.TblSpecificFixedDeti.OperID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.projects_des ON dbo.TblSpecificFixedDeti.PandID = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.projects ON dbo.TblSpecificFixedDeti.ProjectID = dbo.projects.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.FixedAssets ON dbo.TblSpecificFixedDeti.FixedID = dbo.FixedAssets.id"
StrSQL = StrSQL & "  Where (dbo.TblSpecificFixedDeti.SPFixID = " & val(Me.xptxtid.Text) & ")"
   ' StrSQL = "select * from opr_employee_details where pk_id=" & Me.xptxtid.text
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For I = .FixedRows To .Rows - 1
 
                .TextMatrix(I, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("FixedID").value), "", RsDev("FixedID").value)
                .TextMatrix(I, .ColIndex("OperID")) = IIf(IsNull(RsDev("OperID").value), "", RsDev("OperID").value)
                .TextMatrix(I, .ColIndex("ToDate")) = IIf(IsNull(RsDev("ToDate").value), "", RsDev("ToDate").value)
                .TextMatrix(I, .ColIndex("FromDate")) = IIf(IsNull(RsDev("FromDate").value), "", RsDev("FromDate").value)
                .TextMatrix(I, .ColIndex("ProjectID")) = IIf(IsNull(RsDev("ProjectID").value), "", RsDev("ProjectID").value)
                .TextMatrix(I, .ColIndex("PandID")) = IIf(IsNull(RsDev("PandID").value), "", RsDev("PandID").value)
                .TextMatrix(I, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                .TextMatrix(I, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("code").value), "", RsDev("code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                .TextMatrix(I, .ColIndex("ProcessName")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                Else
                .TextMatrix(I, .ColIndex("ProcessName")) = IIf(IsNull(RsDev("ProcessNameE").value), "", RsDev("ProcessNameE").value)
                .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(RsDev("Project_nameE").value), "", RsDev("Project_nameE").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                End If
                .TextMatrix(I, .ColIndex("interval")) = IIf(IsNull(RsDev("LngT").value), "", RsDev("LngT").value)
                .TextMatrix(I, .ColIndex("price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                .TextMatrix(I, .ColIndex("Total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
            
                RsDev.MoveNext
            Next I
 
        End With

    End If
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Sub FillGrid()
Dim Sql As String
Dim CrrRow As Integer
With Grid
If .Rows = 1 Then
.Rows = .Rows + 1
CrrRow = 1
Else
CrrRow = .Rows - 1
.Rows = .Rows + 1
End If
.TextMatrix(CrrRow, .ColIndex("Ser")) = CrrRow
.TextMatrix(CrrRow, .ColIndex("ProjectID")) = val(DCPROJECT1.BoundText)
.TextMatrix(CrrRow, .ColIndex("Project_name")) = DCPROJECT1.Text
.TextMatrix(CrrRow, .ColIndex("PandID")) = val(Dcterm1.BoundText)
.TextMatrix(CrrRow, .ColIndex("des")) = Dcterm1.Text
.TextMatrix(CrrRow, .ColIndex("OperID")) = val(dcopr.BoundText)
.TextMatrix(CrrRow, .ColIndex("ProcessName")) = dcopr.Text
'.TextMatrix(CrrRow, .ColIndex("Emp_Code")) = 1
.TextMatrix(CrrRow, .ColIndex("Emp_Name")) = DcbEqpID.Text
.TextMatrix(CrrRow, .ColIndex("Emp_id")) = val(DcbEqpID.BoundText)
Grid_AfterEdit CrrRow, .ColIndex("Emp_id")
.TextMatrix(CrrRow, .ColIndex("FromDate")) = FromDate.value
.TextMatrix(CrrRow, .ColIndex("ToDate")) = ToDate.value
.TextMatrix(CrrRow, .ColIndex("interval")) = val(Me.TxtPeriod.Text)
.TextMatrix(CrrRow, .ColIndex("price")) = val(Me.TxtPrice.Text)
.TextMatrix(CrrRow, .ColIndex("Total")) = val(Me.TxtPrice.Text) * val(Me.TxtPeriod.Text)

End With
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
If DcbEqpID.Text = "" Or val(DcbEqpID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„⁄œÂ"
Else
MsgBox "Please Select Equipment"
End If
DcbEqpID.SetFocus
Exit Sub
End If
If DCPROJECT1.Text = "" Or val(DCPROJECT1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
Else
MsgBox "Please Select Project"
End If
DCPROJECT1.SetFocus
Exit Sub
End If
If val(TxtPeriod.Text) <= 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «· «—ÌŒ  "
Else
MsgBox "Please Select Date"
End If
'TxtPeriod.SetFocus
Exit Sub
End If

FillGrid
End If
End Sub

Private Sub ISButton3_Click()
   xptxtid.Text = ""
    TxtModFlg.Text = "N"
            Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
            Cmd(1).Enabled = True
End Sub

Private Sub ToDate_Change()
DiffDtaeVlue
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
