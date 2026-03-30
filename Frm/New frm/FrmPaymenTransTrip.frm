VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmPaymenTransTrip 
   BackColor       =   &H00E2E9E9&
   Caption         =   "   ‘«‘… ›Ê« Ì— «·⁄„·«¡"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   HelpContextID   =   580
   Icon            =   "FrmPaymenTransTrip.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   18960
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
      _cx             =   33443
      _cy             =   16801
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
         Height          =   885
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8640
         Width           =   18960
         _cx             =   33443
         _cy             =   1561
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
         Align           =   2
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
         Begin VB.CheckBox vbcheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Check1"
            Height          =   195
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   480
            Width           =   135
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   315
            Left            =   11250
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   75
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":038A
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   300
            Left            =   12090
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   210
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":0724
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   255
            Left            =   13245
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":0ABE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   465
            Index           =   0
            Left            =   15075
            TabIndex        =   8
            Top             =   345
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
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
            Height          =   465
            Index           =   1
            Left            =   14115
            TabIndex        =   9
            Top             =   345
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   820
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
            Height          =   465
            Index           =   2
            Left            =   13320
            TabIndex        =   10
            Top             =   360
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   820
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
            Height          =   465
            Index           =   3
            Left            =   12360
            TabIndex        =   11
            Top             =   345
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   820
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
            Height          =   465
            Index           =   4
            Left            =   11625
            TabIndex        =   12
            Top             =   345
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
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
            Height          =   345
            Index           =   6
            Left            =   1890
            TabIndex        =   13
            Top             =   465
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
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
            Height          =   465
            Index           =   5
            Left            =   10995
            TabIndex        =   14
            Top             =   345
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
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
            Height          =   345
            Left            =   15675
            TabIndex        =   15
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   609
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
            MICON           =   "FrmPaymenTransTrip.frx":0E58
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   465
            Index           =   7
            Left            =   10110
            TabIndex        =   25
            Top             =   360
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   820
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
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   465
            Index           =   8
            Left            =   9030
            TabIndex        =   26
            Top             =   360
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «Ã„«·Ì"
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
            Height          =   465
            Index           =   10
            Left            =   7335
            TabIndex        =   27
            Top             =   360
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… ”Ì«—«  «·€Ì—"
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
            Height          =   465
            Index           =   11
            Left            =   14970
            TabIndex        =   146
            Top             =   8100
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «Ã„«·Ì"
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
            Index           =   12
            Left            =   5880
            TabIndex        =   147
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·‰„Ê–Ã "
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
         Begin ImpulseButton.ISButton ISButton2 
            CausesValidation=   0   'False
            Height          =   420
            Left            =   3195
            TabIndex        =   154
            Top             =   420
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   741
            ButtonStyle     =   1
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
            Index           =   13
            Left            =   4320
            TabIndex        =   159
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·‰„Ê–Ã «Ã„«·Ì  "
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   225
            Left            =   1470
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   225
            Width           =   930
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì"
            Height          =   225
            Left            =   3975
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   225
            Width           =   930
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
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
            Left            =   450
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   210
            Width           =   975
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
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
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   225
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Index           =   5
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   18915
         _cx             =   33364
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
         Picture         =   "FrmPaymenTransTrip.frx":0E74
         Caption         =   "    ‘«‘… ›Ê« Ì— «·⁄„·«¡  "
         Align           =   0
         AutoSizeChildren=   7
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
         Begin VB.TextBox TxtRowNumber 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Text            =   "Text4"
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7515
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1680
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":1B4E
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":1EE8
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
            Left            =   2205
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":2282
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
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "FrmPaymenTransTrip.frx":261C
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7905
         Left            =   0
         TabIndex        =   28
         Top             =   720
         Width           =   18900
         _cx             =   33338
         _cy             =   13944
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
         Caption         =   "›Ê« Ì— «·⁄„·«¡|‘—Õ «·„Ê«“‰…|”Ì«—  «·€Ì—"
         Align           =   0
         CurrTab         =   1
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
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7485
            Left            =   -19455
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   45
            Width           =   18810
            _cx             =   33179
            _cy             =   13203
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   6405
               Left            =   0
               TabIndex        =   30
               Top             =   600
               Width           =   18390
               _cx             =   32438
               _cy             =   11298
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
               Rows            =   12
               Cols            =   33
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPaymenTransTrip.frx":29B6
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
               Caption         =   "”Ì«—«  «·€Ì—"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   450
               Index           =   17
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   120
               Width           =   3390
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì ﬂ„Ì… «· Õ„Ì·"
               Height          =   360
               Index           =   18
               Left            =   15120
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   7080
               Width           =   1560
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì ﬂ„Ì… «· ›—Ì€"
               Height          =   360
               Index           =   19
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   7080
               Width           =   1560
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   360
               Index           =   20
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   7080
               Width           =   1560
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   360
               Index           =   21
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   7080
               Width           =   1560
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7485
            Index           =   1
            Left            =   45
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   45
            Width           =   18810
            _cx             =   33179
            _cy             =   13203
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
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
            Begin VB.TextBox txtContractNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   11280
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   2115
               Width           =   1230
            End
            Begin VB.TextBox txtVendor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   14220
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   2115
               Width           =   1230
            End
            Begin VB.TextBox txtContainerNo 
               BackColor       =   &H0000FFFF&
               Height          =   345
               Left            =   6540
               TabIndex        =   161
               Top             =   60
               Width           =   1740
            End
            Begin VB.PictureBox Picture1 
               Height          =   495
               Left            =   2280
               ScaleHeight     =   435
               ScaleWidth      =   210
               TabIndex        =   160
               Top             =   6360
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.CheckBox chkoWithoutVat 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œÊ‰"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   6840
               Width           =   1020
            End
            Begin VB.CommandButton Command1 
               Caption         =   "«·€«¡«·—»ÿ"
               Height          =   375
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   1920
               Width           =   1410
            End
            Begin VB.CheckBox ChkDate 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " — Ì» »«· «—ÌŒ"
               Height          =   255
               Left            =   4830
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   2010
               Width           =   1470
            End
            Begin VB.TextBox TxtNetValue 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   375
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   6600
               Width           =   1590
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   6600
               Width           =   1230
            End
            Begin VB.TextBox TxtRefNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9690
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   120
               Width           =   1230
            End
            Begin VB.CheckBox chkItem 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’‰›"
               Height          =   195
               Left            =   17490
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   1200
               Width           =   1020
            End
            Begin VB.CheckBox chkTypeTransport 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‰Ê«⁄ «·‰ﬁ·"
               Height          =   195
               Left            =   17490
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   810
               Width           =   1020
            End
            Begin VB.TextBox txtNoteSerial1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   16170
               Locked          =   -1  'True
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   120
               Width           =   1230
            End
            Begin VB.Frame Frame10 
               Caption         =   "»Ì«‰«  „Õ«”»Ì…"
               Height          =   825
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   6570
               Width           =   3990
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1320
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.CommandButton Command9 
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                  Height          =   375
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   195
                  Index           =   35
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   360
                  Width           =   990
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "«Ã„«·Ì« "
               Height          =   1065
               Left            =   -1050
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   7635
               Width           =   12270
               Begin VB.TextBox TxtTotalTo 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   720
                  Width           =   1065
               End
               Begin VB.TextBox TxtPhone 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   360
                  Width           =   945
               End
               Begin VB.TextBox TxtCommiValue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   8280
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtElectricity 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   2160
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   360
                  Width           =   945
               End
               Begin VB.TextBox TxtWater 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtInsuranceValue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   6240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox TxtTotalContract 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„” Õﬁ ··€Ì—"
                  Height          =   195
                  Index           =   1
                  Left            =   11520
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   720
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œœ„« "
                  Height          =   195
                  Index           =   27
                  Left            =   1035
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   480
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄Ì/—”Ê„"
                  Height          =   405
                  Index           =   25
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬂÂ—»«¡"
                  Height          =   195
                  Index           =   21
                  Left            =   2985
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   480
                  Width           =   750
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„Ì«Â"
                  Height          =   195
                  Index           =   20
                  Left            =   5385
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   480
                  Width           =   750
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " √„Ì‰"
                  Height          =   195
                  Index           =   19
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   360
                  Width           =   510
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·«ÌÃ«—"
                  Height          =   195
                  Index           =   6
                  Left            =   11505
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   360
                  Width           =   990
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Õœœ «· «—ÌŒ"
               Height          =   615
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   1245
               Width           =   6600
               Begin MSComCtl2.DTPicker Fromdate 
                  Height          =   330
                  Left            =   4560
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   270925825
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker todate 
                  Height          =   330
                  Left            =   1440
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   270925825
                  CurrentDate     =   41640
               End
               Begin Dynamic_Byte.NourHijriCal Fromdate√H 
                  Height          =   330
                  Left            =   3240
                  TabIndex        =   88
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   582
               End
               Begin Dynamic_Byte.NourHijriCal todateH 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   89
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   582
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·› —… „‰"
                  Height          =   315
                  Index           =   0
                  Left            =   5535
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   240
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈«·Ï"
                  Height          =   435
                  Index           =   14
                  Left            =   2580
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   240
                  Width           =   540
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "«œŒ«· «·”‰Ê«  «·„«÷Ì…"
               Height          =   780
               Left            =   -5550
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1680
               Width           =   3930
               Begin VB.OptionButton OptActual 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Ì"
                  Height          =   195
                  Index           =   1
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   240
                  Width           =   1305
               End
               Begin VB.OptionButton OptActual 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÌœÊÌ"
                  Height          =   195
                  Index           =   0
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1305
               End
            End
            Begin VB.OptionButton OptAlarms 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «·Õ”«»"
               Height          =   315
               Index           =   1
               Left            =   5430
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   -4410
               Width           =   1860
            End
            Begin VB.OptionButton OptAlarms 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ–Ì— ›ﬁÿ"
               Height          =   315
               Index           =   0
               Left            =   7260
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   -4410
               Width           =   1410
            End
            Begin VB.ComboBox OperatorsID 
               Height          =   315
               ItemData        =   "FrmPaymenTransTrip.frx":2E70
               Left            =   19230
               List            =   "FrmPaymenTransTrip.frx":2E80
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Text            =   " "
               Top             =   2070
               Width           =   1350
            End
            Begin VB.TextBox Percentage 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   19740
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Text            =   "0"
               Top             =   2070
               Width           =   1080
            End
            Begin VB.Frame Frame6 
               Caption         =   "Õœœ «·„Ê«“‰«  «·”«»ﬁ…"
               Height          =   1560
               Left            =   20040
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   405
               Width           =   8460
               Begin VSFlex8Ctl.VSFlexGrid GridOldEstimation 
                  Height          =   915
                  Left            =   120
                  TabIndex        =   77
                  Top             =   240
                  Width           =   8265
                  _cx             =   14579
                  _cy             =   1614
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
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmPaymenTransTrip.frx":2E9C
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
            End
            Begin VB.Frame Frame5 
               Caption         =   "Õœœ ”‰Ê«  «·„ﬁ«—‰…"
               Height          =   1560
               Left            =   19560
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   405
               Width           =   5280
               Begin VSFlex8Ctl.VSFlexGrid GridIntervals1 
                  Height          =   915
                  Left            =   120
                  TabIndex        =   75
                  Top             =   240
                  Width           =   4545
                  _cx             =   8017
                  _cy             =   1614
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmPaymenTransTrip.frx":2F3A
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
            End
            Begin VB.Frame Frame1 
               Caption         =   "«· Ê“Ì⁄ ⁄·Ï «Õ”«»« "
               Height          =   1065
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   9075
               Width           =   17310
               Begin VB.TextBox TxtRemarks1 
                  Alignment       =   1  'Right Justify
                  Height          =   615
                  Left            =   2160
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   67
                  Top             =   120
                  Width           =   3615
               End
               Begin VB.TextBox TxtPercentage 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   240
                  Width           =   1215
               End
               Begin MSDataListLib.DataCombo DCAccountDist 
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   68
                  Top             =   240
                  Width           =   3855
                  _ExtentX        =   6800
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   960
                  TabIndex        =   69
                  Top             =   240
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmPaymenTransTrip.frx":301F
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   240
                  TabIndex        =   70
                  Top             =   240
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmPaymenTransTrip.frx":33B9
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   9
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   240
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·‰”»Â"
                  Height          =   315
                  Index           =   6
                  Left            =   8040
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Õ”«»"
                  Height          =   315
                  Index           =   5
                  Left            =   12720
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1080
               End
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   1020
               Left            =   150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   64
               Top             =   600
               Width           =   3570
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   930
               Left            =   19110
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1035
               Width           =   2670
               Begin VB.OptionButton PercentagType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”» ÌœÊÌÂ"
                  Height          =   210
                  Index           =   1
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.OptionButton PercentagType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”» «·ÌÂ"
                  Height          =   210
                  Index           =   0
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.TextBox TxtTransID 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   16170
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   120
               Width           =   1230
            End
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   -4860
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   11790
               Width           =   2670
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   510
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   -405
               Visible         =   0   'False
               Width           =   2640
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1290
               Left            =   20670
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   405
               Width           =   2910
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï  «·›—Ê⁄"
                  Height          =   210
                  Index           =   2
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï Õ”«»« "
                  Height          =   210
                  Index           =   0
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.OptionButton DistType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Ê“Ì⁄ ⁄·Ï „—«ﬂ“  ﬂ·›…"
                  Height          =   210
                  Index           =   1
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   480
                  Width           =   2055
               End
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   375
               Left            =   4080
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   7080
               Width           =   4110
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   16470
               TabIndex        =   52
               Top             =   480
               Width           =   930
            End
            Begin VB.TextBox TxtItemCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   16470
               TabIndex        =   51
               Top             =   1200
               Width           =   930
            End
            Begin VB.CheckBox Check17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   16800
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   2160
               Width           =   1710
            End
            Begin VB.TextBox TxtQtyDownload 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   15600
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   6600
               Width           =   1710
            End
            Begin VB.TextBox TxtQtyDischarge 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   15600
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   7080
               Width           =   1710
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   13440
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   6600
               Width           =   1350
            End
            Begin VB.TextBox TxtTotalValue 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   375
               Left            =   11040
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   6600
               Width           =   1590
            End
            Begin VB.TextBox TxtVAT 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   375
               Left            =   4080
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   6600
               Width           =   1350
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·« Ã«Â „‰ "
               Height          =   195
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   480
               Width           =   1110
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·« Ã«Â «·Ï"
               Height          =   315
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   840
               Width           =   1110
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”›Ì‰…"
               Height          =   195
               Left            =   17400
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   1560
               Width           =   1110
            End
            Begin VB.TextBox TxtDesc 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Top             =   1800
               Width           =   3570
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   495
               Left            =   450
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   0
               Width           =   3570
               _cx             =   6297
               _cy             =   873
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
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   42
                  Top             =   120
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÌœÊÌ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   43
                  Top             =   120
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«·Ì"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin XtremeSuiteControls.RadioButton RdQty 
               Height          =   255
               Index           =   0
               Left            =   17400
               TabIndex        =   49
               Top             =   6600
               Width           =   1230
               _Version        =   786432
               _ExtentX        =   2170
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "ﬂ„Ì… «· Õ„Ì·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCAccountMaster 
               Height          =   315
               Left            =   22890
               TabIndex        =   111
               Top             =   630
               Width           =   6390
               _ExtentX        =   11271
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
               Height          =   4905
               Left            =   20700
               TabIndex        =   112
               Top             =   2550
               Width           =   18510
               _cx             =   32650
               _cy             =   8652
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
               Cols            =   28
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPaymenTransTrip.frx":3953
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
               Height          =   315
               Left            =   12120
               TabIndex        =   113
               Top             =   120
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Format          =   266141697
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo DcBranch 
               Height          =   315
               Left            =   3930
               TabIndex        =   114
               Top             =   120
               Width           =   1920
               _ExtentX        =   3387
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
            Begin Dynamic_Byte.NourHijriCal recordDateH 
               Height          =   315
               Left            =   13770
               TabIndex        =   115
               Top             =   120
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
            End
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
               Height          =   4125
               Left            =   330
               TabIndex        =   116
               Top             =   2520
               Width           =   18390
               _cx             =   32438
               _cy             =   7276
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
               Rows            =   12
               Cols            =   42
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPaymenTransTrip.frx":3D8A
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
               ExplorerBar     =   3
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
            Begin MSDataListLib.DataCombo DBCboClientName2 
               Height          =   315
               Left            =   12150
               TabIndex        =   117
               Top             =   480
               Width           =   4260
               _ExtentX        =   7514
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbTypeTransport 
               Height          =   315
               Left            =   12150
               TabIndex        =   118
               Top             =   840
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboItems 
               Height          =   315
               Left            =   12150
               TabIndex        =   119
               Top             =   1200
               Width           =   4260
               _ExtentX        =   7514
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   9
               Left            =   4200
               TabIndex        =   120
               Top             =   1440
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmPaymenTransTrip.frx":438E
               DrawFocusRectangle=   0   'False
            End
            Begin XtremeSuiteControls.RadioButton RdQty 
               Height          =   255
               Index           =   1
               Left            =   17400
               TabIndex        =   121
               Top             =   7080
               Width           =   1230
               _Version        =   786432
               _ExtentX        =   2170
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "ﬂ„Ì… «· ›—Ì€"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbShip 
               Height          =   315
               Left            =   12150
               TabIndex        =   122
               Top             =   1560
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcCityFromId 
               Height          =   315
               Left            =   5250
               TabIndex        =   123
               Top             =   480
               Width           =   5670
               _ExtentX        =   10001
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
            Begin MSDataListLib.DataCombo DcCityToId 
               Height          =   315
               Left            =   5250
               TabIndex        =   124
               Top             =   840
               Width           =   5670
               _ExtentX        =   10001
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
            Begin MSComCtl2.DTPicker txtDuDate 
               Height          =   315
               Left            =   8460
               TabIndex        =   167
               Top             =   2100
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Format          =   266141697
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
               Height          =   315
               Index           =   28
               Left            =   9540
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   2160
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·⁄ﬁœ"
               Height          =   315
               Index           =   27
               Left            =   12180
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   2145
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·›‰œ— «·„ﬁ«Ê·"
               Height          =   315
               Index           =   26
               Left            =   15240
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   2145
               Width           =   1530
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «—«„ﬂÊ"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   0
               Left            =   8250
               TabIndex        =   162
               Top             =   150
               Width           =   1590
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ì"
               Height          =   315
               Index           =   25
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   6600
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Œ’Ê„« "
               Height          =   315
               Index           =   24
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   6600
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„—Ã⁄ "
               Height          =   315
               Index           =   23
               Left            =   10350
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   120
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ „Œ«·›… «· ﬁœÌ—Ï"
               ForeColor       =   &H000000FF&
               Height          =   480
               Index           =   16
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   -4410
               Width           =   2370
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   390
               Left            =   11280
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   -2610
               Width           =   900
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»…"
               Height          =   390
               Index           =   0
               Left            =   21030
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   2070
               Width           =   900
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· ﬁœÌ— „ Ê”ÿ „«”»ﬁ"
               Height          =   465
               Index           =   15
               Left            =   19200
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   2070
               Width           =   2370
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   210
               Index           =   13
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   120
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ… Â«„…:-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Index           =   37
               Left            =   -1410
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   1950
               Visible         =   0   'False
               Width           =   1680
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   360
               Index           =   2
               Left            =   3900
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   750
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —… „‰ "
               Height          =   495
               Index           =   4
               Left            =   18990
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   630
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· Ê“Ì⁄"
               Height          =   330
               Index           =   3
               Left            =   19980
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   1560
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂ…"
               Height          =   315
               Index           =   8
               Left            =   15150
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   120
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               Height          =   315
               Index           =   7
               Left            =   16920
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   120
               Width           =   1530
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   17010
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   1185
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   315
               Index           =   1
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   7080
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   315
               Index           =   64
               Left            =   16920
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   480
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   315
               Index           =   10
               Left            =   14400
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   6600
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… «·„÷«›…"
               Height          =   195
               Index           =   11
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   6600
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   315
               Index           =   12
               Left            =   12360
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   6600
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ê’› «·›« Ê—…"
               Height          =   360
               Index           =   22
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   1830
               Width           =   960
            End
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
      ButtonImage     =   "FrmPaymenTransTrip.frx":4728
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmPaymenTransTrip"
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
 
Private Sub Del_Trans()
    Dim Msg As String
    On Error GoTo ErrTrap
   Dim StrSQL As String
    If TxtTransID.text <> "" Then
     Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtTransID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
             StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
         StrSQL = "Delete From Notes Where  NoteType= 9080 and NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
       StrSQL = " update  TblTripTypesTransport set  allocations=null where ID in( " & " select NoteID from TblTravDueKDet where TravID=" & TxtTransID & ")"
     Cn.Execute StrSQL
     Cn.Execute "delete TblTravDueKDet where TravID=" & val(Me.TxtTransID.text)
                rs.delete
             
        
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                
                    TxtModFlg_Change
                    LabCurrRec.Caption = 0
                    LabCountRec.Caption = 0
                Else
                    clear_all Me
                    Retrive
                End If

                '--------
            
                '-------
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
    StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID, dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID,"
    StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
    StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,TblTravDueKDet.ContItem,TblTravDueKDet.PurchaseOrderNo,TblTravDueKDet.LocationName,TblTravDueKDet.RentType,"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueK.recordDate, dbo.TblTravDueK.recordDateH,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.Fromdate, dbo.TblTravDueK.FromdateH, dbo.TblTravDueK.todate, dbo.TblTravDueK.todateH, dbo.TblTravDueK.Remarks AS HRemarks,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.QtyDischarge AS HQtyDischarge,TblTravDueK.Descrp, dbo.TblTravDueK.QtyDownload AS HQtyDownload, dbo.TblTravDueK.TotalValue, dbo.TblTravDueK.Price,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.VAT, dbo.TblTravDueK.RdQty, dbo.TblTravDueK.TypeTransportID, dbo.TblTypesTransport.Name AS TypeTransName,"
    StrSQL = StrSQL & "                      dbo.TblTypesTransport.NameE AS TypeTransNameE, dbo.TblTravDueK.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblTravDueK.ID AS HID,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.BranchId AS HBranchId, TblBranchesData_1.branch_name AS Hbranch_name, TblBranchesData_1.branch_namee AS Hbranch_nameE ,dbo.TblTravDueK.NoteSerial1,"
    StrSQL = StrSQL & "                      TblTravDueK.QrCodeImage,TblCustemers.VATNO,TblCustemers.Address as AddressSupp,TblTravDueK.Vendor,TblTravDueK.DuDate,TblTravDueK.ContractNo ,"
    StrSQL = StrSQL & "                      TblBranchesData.Company_Arabic_Name,TblBranchesData.Company_Name_Eng,"
    StrSQL = StrSQL & "                         TblBranchesData.CityName as CompanyCityName,TblBranchesData.CitySubdivisionName,TblBranchesData.PostalZone CompanyPostalZone,TblBranchesData.StreetName CompanyStreetName,TblBranchesData.VATRegNo CompanyVATRegNo,TblBranchesData.Company_Comment CompanyRegNo"
    StrSQL = StrSQL & "  ,dbo.TblCustemers.Address ,TblCustemers.CustGID, TblCustemers.VATNO,TblCustemers.CusName,TblCustemers.CusNamee,TblCustemers.AddressE "
    StrSQL = StrSQL & ",TblCountriesData.CountryName,TblCountriesData.ECountryName,TblCountriesGovernments.GovernmentName"
    StrSQL = StrSQL & " FROM         dbo.TblCarsData RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTypesTransport RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTravDueK ON TblBranchesData_1.branch_id = dbo.TblTravDueK.BranchId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblTravDueK.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblTravDueK.ItemID = dbo.TblItems.ItemID ON dbo.TblTypesTransport.ID = dbo.TblTravDueK.TypeTransportID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet ON dbo.TblBranchesData.branch_id = dbo.TblTravDueKDet.BranchID ON dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID ON dbo.TblCarsData.id = dbo.TblTravDueKDet.CarID"
    
       
        StrSQL = StrSQL & "                       Left Outer join TblCountriesData On TblCustemers.CountryID =TblCountriesData.CountryID "
        StrSQL = StrSQL & "                       Left Outer join TblCountriesGovernments On TblCustemers.GovernmentID =TblCountriesGovernments.GovernmentID "
 
       
        
    StrSQL = StrSQL & "   Where (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") and (dbo.TblTravDueKDet.TypeTrans =0 or dbo.TblTravDueKDet.TypeTrans is null) "

 
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(txtTotal.text), "#.##"), 0, True)
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(5).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(5).AddCurrentValue GetRegVATNo(val(Dcbranch.BoundText))
    End If
    'xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , StrSQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_report4(Optional NoteSerial As String)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    StrSQL = " SELECT     TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TblTravDueK.recordDate, dbo.TblTravDueK.recordDateH, dbo.TblTravDueK.Fromdate, "
    StrSQL = StrSQL & "                      dbo.TblTravDueK.FromdateH, dbo.TblTravDueK.todate, dbo.TblTravDueK.todateH, dbo.TblTravDueK.Remarks AS HRemarks,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.QtyDischarge AS HQtyDischarge, dbo.TblTravDueK.QtyDownload AS HQtyDownload, dbo.TblTravDueK.TotalValue, dbo.TblTravDueK.Price,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.VAT, dbo.TblTravDueK.RdQty, dbo.TblTravDueK.TypeTransportID, dbo.TblTypesTransport.Name AS TypeTransName,"
    StrSQL = StrSQL & "                      dbo.TblTypesTransport.NameE AS TypeTransNameE, dbo.TblTravDueK.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblTravDueK.ID,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.BranchId AS HBranchId, TblBranchesData_1.branch_name AS Hbranch_name, TblBranchesData_1.branch_namee AS Hbranch_nameE,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.Descrp , dbo.TblTravDueK.NoteSerial1,TblTravDueK.QrCodeImage,TblCustemers.VATNO,TblCustemers.Address as AddressSupp"
    StrSQL = StrSQL & " FROM         dbo.TblTypesTransport RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTravDueK ON TblBranchesData_1.branch_id = dbo.TblTravDueK.BranchId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblTravDueK.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblTravDueK.ItemID = dbo.TblItems.ItemID ON dbo.TblTypesTransport.ID = dbo.TblTravDueK.TypeTransportID"
    StrSQL = StrSQL & "   Where (dbo.TblTravDueK.ID = " & val(Me.TxtTransID.text) & ") "

 
 
 
 
  StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID, dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID,"
    StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
    StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,TblTravDueKDet.ContItem,TblTravDueKDet.PurchaseOrderNo,TblTravDueKDet.LocationName,TblTravDueKDet.RentType,"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueK.recordDate, dbo.TblTravDueK.recordDateH,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.Fromdate, dbo.TblTravDueK.FromdateH, dbo.TblTravDueK.todate, dbo.TblTravDueK.todateH, dbo.TblTravDueK.Remarks AS HRemarks,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.QtyDischarge AS HQtyDischarge,TblTravDueK.Descrp, dbo.TblTravDueK.QtyDownload AS HQtyDownload, dbo.TblTravDueK.TotalValue, dbo.TblTravDueK.Price,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.VAT, dbo.TblTravDueK.RdQty, dbo.TblTravDueK.TypeTransportID, dbo.TblTypesTransport.Name AS TypeTransName,"
    StrSQL = StrSQL & "                      dbo.TblTypesTransport.NameE AS TypeTransNameE, dbo.TblTravDueK.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblTravDueK.ID AS HID,"
    StrSQL = StrSQL & "                      dbo.TblTravDueK.BranchId AS HBranchId, TblBranchesData_1.branch_name AS Hbranch_name, TblBranchesData_1.branch_namee AS Hbranch_nameE ,dbo.TblTravDueK.NoteSerial1,"
    StrSQL = StrSQL & "                      TblTravDueK.QrCodeImage,TblCustemers.VATNO,TblCustemers.Address as AddressSupp,TblTravDueK.Vendor,TblTravDueK.DuDate,TblTravDueK.ContractNo ,"
    StrSQL = StrSQL & "                      TblBranchesData.Company_Arabic_Name,TblBranchesData.Company_Name_Eng,"
    StrSQL = StrSQL & "                         TblBranchesData.CityName as CompanyCityName,TblBranchesData.CitySubdivisionName,TblBranchesData.PostalZone CompanyPostalZone,TblBranchesData.StreetName CompanyStreetName,TblBranchesData.VATRegNo CompanyVATRegNo,TblBranchesData.Company_Comment CompanyRegNo"
    StrSQL = StrSQL & "  ,dbo.TblCustemers.Address ,TblCustemers.CustGID, TblCustemers.VATNO,TblCustemers.CusName,TblCustemers.CusNamee,TblCustemers.AddressE "
    StrSQL = StrSQL & ",TblCountriesData.CountryName,TblCountriesData.ECountryName,TblCountriesGovernments.GovernmentName"
    StrSQL = StrSQL & " FROM         dbo.TblCarsData RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTypesTransport RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTravDueK ON TblBranchesData_1.branch_id = dbo.TblTravDueK.BranchId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblTravDueK.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblTravDueK.ItemID = dbo.TblItems.ItemID ON dbo.TblTypesTransport.ID = dbo.TblTravDueK.TypeTransportID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblTravDueKDet ON dbo.TblBranchesData.branch_id = dbo.TblTravDueKDet.BranchID ON dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID ON dbo.TblCarsData.id = dbo.TblTravDueKDet.CarID"
       
       
        StrSQL = StrSQL & "                       Left Outer join TblCountriesData On TblCustemers.CountryID =TblCountriesData.CountryID "
        StrSQL = StrSQL & "                       Left Outer join TblCountriesGovernments On TblCustemers.GovernmentID =TblCountriesGovernments.GovernmentID "
 
       
    StrSQL = StrSQL & "   Where (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") and (dbo.TblTravDueKDet.TypeTrans =0 or dbo.TblTravDueKDet.TypeTrans is null) "
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip4.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip4.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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
        StrReportTitle = ""
 
    End If

      xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(txtTotal.text), "#.##"), 0, True)
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(5).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(5).AddCurrentValue GetRegVATNo(val(Dcbranch.BoundText))
    End If
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_report3(Optional NoteSerial As String)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID, dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
StrSQL = StrSQL & "                      TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueK.recordDate, dbo.TblTravDueK.recordDateH,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.Fromdate, dbo.TblTravDueK.FromdateH, dbo.TblTravDueK.todate, dbo.TblTravDueK.todateH, dbo.TblTravDueK.Remarks AS HRemarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.QtyDischarge AS HQtyDischarge, dbo.TblTravDueK.QtyDownload AS HQtyDownload, dbo.TblTravDueK.TotalValue, dbo.TblTravDueK.Price,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.VAT, dbo.TblTravDueK.RdQty, dbo.TblTravDueK.TypeTransportID, dbo.TblTypesTransport.Name AS TypeTransName,"
StrSQL = StrSQL & "                      dbo.TblTypesTransport.NameE AS TypeTransNameE, dbo.TblTravDueK.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblTravDueK.ID AS HID,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.BranchId AS HBranchId, TblBranchesData_1.branch_name AS Hbranch_name, TblBranchesData_1.branch_namee AS Hbranch_nameE,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.LeaderName, dbo.TblVendorCars.CustomerID, TblCustemers_1.CusName AS VendorName, TblCustemers_1.CusNamee AS VendorNameE,"
StrSQL = StrSQL & "                      TblCustemers_1.Fullcode AS VendorFullcode , dbo.TblTravDueK.NoteSerial1,TblTravDueK.QrCodeImage,TblCustemers.VATNO,TblCustemers.Address as AddressSupp"""
StrSQL = StrSQL & " FROM         dbo.TblTypesTransport RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTravDueK ON TblBranchesData_1.branch_id = dbo.TblTravDueK.BranchId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblTravDueK.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblTravDueK.ItemID = dbo.TblItems.ItemID ON dbo.TblTypesTransport.ID = dbo.TblTravDueK.TypeTransportID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_2 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet ON TblBranchesData_2.branch_id = dbo.TblTravDueKDet.BranchID ON dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON TblCustemers_1.CusID = dbo.TblVendorCars.CustomerID ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id"
 StrSQL = StrSQL & "   Where (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") and (dbo.TblTravDueKDet.TypeTrans =1) "

 
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip3.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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
        StrReportTitle = ""
 
    End If

      xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue WriteNo(val(txtTotal.text), 0, True)
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(5).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(5).AddCurrentValue GetRegVATNo(val(Dcbranch.BoundText))
    End If
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_report2(Optional NoteSerial As String)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID, dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
StrSQL = StrSQL & "                      TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueK.recordDate, dbo.TblTravDueK.recordDateH,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.Fromdate, dbo.TblTravDueK.FromdateH, dbo.TblTravDueK.todate, dbo.TblTravDueK.todateH, dbo.TblTravDueK.Remarks AS HRemarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.QtyDischarge AS HQtyDischarge, dbo.TblTravDueK.QtyDownload AS HQtyDownload, dbo.TblTravDueK.TotalValue, dbo.TblTravDueK.Price,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.VAT, dbo.TblTravDueK.RdQty, dbo.TblTravDueK.TypeTransportID, dbo.TblTypesTransport.Name AS TypeTransName,"
StrSQL = StrSQL & "                      dbo.TblTypesTransport.NameE AS TypeTransNameE, dbo.TblTravDueK.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblTravDueK.ID AS HID,"
StrSQL = StrSQL & "                      dbo.TblTravDueK.BranchId AS HBranchId, TblBranchesData_1.branch_name AS Hbranch_name, TblBranchesData_1.branch_namee AS Hbranch_nameE,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.LeaderName, dbo.TblVendorCars.CustomerID, TblCustemers_1.CusName AS VendorName, TblCustemers_1.CusNamee AS VendorNameE,"
StrSQL = StrSQL & "                      TblCustemers_1.Fullcode AS VendorFullcode , dbo.TblTravDueK.NoteSerial1,TblTravDueK.QrCodeImage,TblCustemers.VATNO,TblCustemers.Address as AddressSupp"""
StrSQL = StrSQL & " FROM         dbo.TblTypesTransport RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTravDueK ON TblBranchesData_1.branch_id = dbo.TblTravDueK.BranchId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblTravDueK.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblTravDueK.ItemID = dbo.TblItems.ItemID ON dbo.TblTypesTransport.ID = dbo.TblTravDueK.TypeTransportID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_2 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet ON TblBranchesData_2.branch_id = dbo.TblTravDueKDet.BranchID ON dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON TblCustemers_1.CusID = dbo.TblVendorCars.CustomerID ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id"
StrSQL = StrSQL & "   Where (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") and (dbo.TblTravDueKDet.TypeTrans =0 or dbo.TblTravDueKDet.TypeTrans is null) "

 
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPayemntTransTrip2.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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
        StrReportTitle = ""
 
    End If

      xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue WriteNo(val(txtTotal.text), 0, True)
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(5).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(5).AddCurrentValue GetRegVATNo(val(Dcbranch.BoundText))
    End If
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function


Function print_report5(Optional NoteSerial As String, Optional rptNo As Double = 0)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
  StrSQL = " SELECT    dbo.TblTravDueK.Fromdate, dbo.TblTravDueK.FromdateH, dbo.TblTravDueK.todate, dbo.TblTravDueK.todateH , TblTravDueK.discount,   RefNo,   dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID,dbo.TblTravDueK.RdQty ,dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
StrSQL = StrSQL & "                          TblTravDueK.RecordDate ,dbo.TblTravDueK.TotalValue , dbo.TblTravDueK.Vat,dbo.TblTravDueK.TotalValue + TblTravDueK.Vat as NetValue,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.Price,TblTravDueKDet.RecNo,TblTravDueKDet.Weight,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueKDet.TypeTrans, dbo.TblTravDueKDet.ShipID,"
StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.TblTravDueKDet.LeaderName,"
StrSQL = StrSQL & "                      tc.CusName , tc.VATNO, tc.Address,TblTravDueK.noteserial1"
StrSQL = StrSQL & " FROM         dbo.TblTravDueKDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.TblTravDueKDet.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblTravDueKDet.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "                                  LEFT OUTER JOIN dbo.TblTravDueK"
StrSQL = StrSQL & "                                              ON  dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID"
StrSQL = StrSQL & "                                              LEFT OUTER JOIN dbo.TblCustemers AS tc"
StrSQL = StrSQL & "                                              ON  tc.CusId = dbo.TblTravDueK.CusId"
StrSQL = StrSQL & "   Where 1= 1 and (dbo.TblTravDueKDet.TypeTrans is null or dbo.TblTravDueKDet.TypeTrans=0)  "
StrSQL = StrSQL & "   and (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") "

 
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymenTransTrip.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymenTransTrip.rpt"
        End If

If rptNo = 0 Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymenTransTrip.rpt"

ElseIf rptNo = 1 Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymenTransTrip1.rpt"

End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
    
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
       ' xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
 
    End If

    '  xReport.ParameterFields(3).AddCurrentValue user_name
    ' xReport.ParameterFields(4).AddCurrentValue WriteNo(val(TxtTotal.Text), 0, True)
    If SystemOptions.VATNoAccordActivity = False Then
   ' xReport.ParameterFields(5).AddCurrentValue cCompanyInfo.VATRegNo
    Else
   ' xReport.ParameterFields(5).AddCurrentValue GetRegVATNo(val(DcBranch.BoundText))
    End If
    Dim i As Long
    For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@NetValueString}"
            xReport.FormulaFields.Item(i).text = "'" & WriteNo(Format(Me.txtTotal.text, "0.00"), 0, True, ".", , 0) & "'"
        End Select
    Next i
    
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , StrSQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub Check1_Click()
If Me.Check1.value = vbUnchecked Then
DcCityFromId.BoundText = ""
End If
End Sub

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.GridInstallments
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
            Next i

        End With

    Else

        With Me.GridInstallments

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If

ReLineGrid
End Sub


Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim RsDev2 As ADODB.Recordset
    Dim LngDevID As Long
    'On Error GoTo ErrTrap
    '----------------------------------------------------------------------------------------
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
           Me.TxtTransID.text = CStr(new_id("TblTravDueK", "ID", "", True))
    ElseIf Me.TxtModFlg.text = "E" Then
            Cn.Execute "delete TblTravDueKDet where TravID=" & val(Me.TxtTransID.text)
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
           
          '          StrSQL = "Delete From Notes Where  NoteType= 9080 and NoteID=" & val(Me.TXTNoteID.Text)
       ' Cn.Execute StrSQL, , adExecuteNoRecords
        
    End If
    If TxtNoteserial1.text = "" Then
              TxtNoteserial1.text = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 76, 76)
        End If
        rs("NoteSerial1").value = IIf(Me.TxtNoteserial1 <> "", val(TxtNoteserial1.text), Null)
    rs("ID").value = TxtTransID.text
    rs("recordDate").value = XPDtbTrans.value
    rs("RecorddateH").value = RecorddateH.value
    rs("Fromdate").value = FromDate.value
    rs("todate").value = ToDate.value
    rs("Fromdateh").value = ToHijriDate(FromDate.value)
    rs("todateh").value = ToHijriDate(ToDate.value)
    rs("BranchId").value = IIf(Me.Dcbranch.BoundText = "", Null, val(Me.Dcbranch.BoundText))
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", "", Me.txtRemarks.text)
    rs("CusID").value = val(DBCboClientName2.BoundText)
    rs("TypeTransportID").value = val(DcbTypeTransport.BoundText)
    rs("ItemID").value = val(DcboItems.BoundText)
    rs("QtyDischarge").value = val(TxtQtyDischarge.text)
    rs("QtyDownload").value = val(TxtQtyDownload.text)
    rs("QtyDischarge2").value = val(lbl(21).Caption)
    rs("ContainerNo").value = IIf(txtContainerNo.text = "", Null, Trim(txtContainerNo.text))
    rs("QtyDownload2").value = val(lbl(20).Caption)
    rs("VAT").value = val(TxtVAT.text)
    rs("Price").value = val(TxtPrice.text)
    rs("TotalValue").value = val(TxtTotalValue.text)
    rs("CityFromId").value = val(Me.DcCityFromId.BoundText)
    rs("CityToId").value = val(Me.DcCityToId.BoundText)
    rs("ShipID").value = val(Me.DcbShip.BoundText)
    rs("Descrp").value = TxtDesc.text
    rs("Discount").value = val(txtDiscount.text)
    rs("NetValue").value = val(TxtNetValue.text)
    rs("total").value = val(txtTotal.text)
    
    rs("RefNo").value = TxtRefNo.text
    If RdQty(1).value = True Then
       rs("RdQty").value = 1
    Else
       rs("RdQty").value = 0
    End If
     If RdAuto_Manual(1).value = True Then
       rs("RdAuto_Manual").value = 1
    Else
       rs("RdAuto_Manual").value = 0
    End If
    If Me.Check1.value = vbChecked Then
       rs("Ch1").value = 1
    Else
       rs("Ch1").value = 0
    End If
    If Me.Check2.value = vbChecked Then
       rs("Ch2").value = 1
    Else
       rs("Ch2").value = 0
    End If
    If Me.Check3.value = vbChecked Then
       rs("Ch3").value = 1
    Else
       rs("Ch3").value = 0
    End If
        If chkTypeTransport.value = vbChecked Then
       rs("chkTypeTransport").value = 1
    Else
       rs("chkTypeTransport").value = 0
    End If
    
       If chkoWithoutVat.value = vbChecked Then
       rs("chkoWithoutVat").value = 1
    Else
       rs("chkoWithoutVat").value = 0
    End If
       
       
     If chkItem.value = vbChecked Then
       rs("chkItem").value = 1
    Else
       rs("chkItem").value = 0
    End If
        rs("Vendor").value = Trim(txtVendor.text)
        rs("ContractNo").value = Trim(txtContractNo.text)
        rs("DuDate").value = (txtDuDate.value)
    
    rs.update
    

    Set RsDetails1 = New ADODB.Recordset
 'DB_CreateField "TblTravDueKDet", "Price", adCurrency, adColNullable, , , , False, True
   StrSQL = "SELECT  *  from dbo.TblTravDueKDet Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Dim i As Integer
      
    With Me.GridInstallments
'Selected
        For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         RsDetails1.AddNew
         RsDetails1("TypeTrans").value = 0
         RsDetails1("TravID").value = val(Me.TxtTransID.text)
         RsDetails1("ShipID").value = val(.TextMatrix(i, .ColIndex("ShipID")))
         RsDetails1("ContainerNo").value = (.TextMatrix(i, .ColIndex("ContainerNo")))
        
         RsDetails1("notesallid").value = val(.TextMatrix(i, .ColIndex("NoteIDA")))
         RsDetails1("TxtRent").value = val(.TextMatrix(i, .ColIndex("TxtRent")))
         
         
         RsDetails1("ContItem").value = (.TextMatrix(i, .ColIndex("ContItem")))
         RsDetails1("PurchaseOrderNo").value = (.TextMatrix(i, .ColIndex("PurchaseOrderNo")))
         RsDetails1("LocationName").value = (.TextMatrix(i, .ColIndex("LocationName")))
         RsDetails1("RentType").value = (.TextMatrix(i, .ColIndex("RentType")))
         
         RsDetails1("TripNo").value = (.TextMatrix(i, .ColIndex("TripNo")))
         RsDetails1("TripDate").value = (.TextMatrix(i, .ColIndex("TripDate")))
         RsDetails1("BranchID").value = val(.TextMatrix(i, .ColIndex("BranchID")))
         RsDetails1("CardNO").value = (.TextMatrix(i, .ColIndex("CardNO")))
         RsDetails1("QtyDownload").value = val(.TextMatrix(i, .ColIndex("QtyDownload")))
         RsDetails1("CardNO2").value = (.TextMatrix(i, .ColIndex("CardNO2")))
         RsDetails1("QtyDischarge").value = val(.TextMatrix(i, .ColIndex("QtyDischarge")))
         RsDetails1("FromID").value = val(.TextMatrix(i, .ColIndex("FromID")))
         RsDetails1("ToID").value = val(.TextMatrix(i, .ColIndex("ToID")))
         RsDetails1("CarTypeID").value = val(.TextMatrix(i, .ColIndex("CarTypeID")))
         RsDetails1("CarID").value = val(.TextMatrix(i, .ColIndex("CarID")))
         RsDetails1("CarType1").value = val(.TextMatrix(i, .ColIndex("CarType1")))
         RsDetails1("Price").value = val(.TextMatrix(i, .ColIndex("Value")))
          RsDetails1("RecNo").value = val(.TextMatrix(i, .ColIndex("RecNo")))
         RsDetails1("Weight").value = val(.TextMatrix(i, .ColIndex("Weight")))
         RsDetails1("TotalValue").value = val(.TextMatrix(i, .ColIndex("TotalValue")))
         RsDetails1("Remarks").value = (.TextMatrix(i, .ColIndex("Remarks")))
         RsDetails1("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
         RsDetails1("LeaderName").value = (.TextMatrix(i, .ColIndex("EmpName")))
         RsDetails1.update
         Cn.Execute " update  TblTripTypesTransport set  allocations=1 where ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
         Else
       
        Cn.Execute " update  TblTripTypesTransport set  allocations=null where ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
        End If
           Next i
        RsDetails1.Close
    End With
        Set RsDetails1 = New ADODB.Recordset
   StrSQL = "SELECT  *  from dbo.TblTravDueKDet Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.VSFlexGrid2
'Selected
        For i = 1 To .rows - 1
''        If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         RsDetails1.AddNew
         RsDetails1("TypeTrans").value = 1
         RsDetails1("TravID").value = val(Me.TxtTransID.text)
         RsDetails1("ShipID").value = val(.TextMatrix(i, .ColIndex("ShipID")))
         RsDetails1("TripNo").value = val(.TextMatrix(i, .ColIndex("TripNo")))
         RsDetails1("TripDate").value = IIf(.TextMatrix(i, .ColIndex("TripDate")) = "", Null, .TextMatrix(i, .ColIndex("TripDate")))
         RsDetails1("BranchID").value = val(.TextMatrix(i, .ColIndex("BranchID")))
         RsDetails1("CardNO").value = (.TextMatrix(i, .ColIndex("CardNO")))
         RsDetails1("QtyDownload").value = val(.TextMatrix(i, .ColIndex("QtyDownload")))
         RsDetails1("CardNO2").value = (.TextMatrix(i, .ColIndex("CardNO2")))
         RsDetails1("QtyDischarge").value = val(.TextMatrix(i, .ColIndex("QtyDischarge")))
         RsDetails1("FromID").value = val(.TextMatrix(i, .ColIndex("FromID")))
         RsDetails1("ToID").value = val(.TextMatrix(i, .ColIndex("ToID")))
         RsDetails1("CarTypeID").value = val(.TextMatrix(i, .ColIndex("CarTypeID")))
         RsDetails1("CarID").value = val(.TextMatrix(i, .ColIndex("CarID")))
         RsDetails1("CarType1").value = val(.TextMatrix(i, .ColIndex("CarType1")))
         RsDetails1("Remarks").value = (.TextMatrix(i, .ColIndex("Remarks")))
         RsDetails1("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
         RsDetails1("LeaderName").value = (.TextMatrix(i, .ColIndex("EmpName")))
         RsDetails1.update
 ''        End If
           Next i
        RsDetails1.Close
    End With

    Cn.CommitTrans
                 
SaveQRCode "TblTravDueK", "ID", val(TxtTransID.text), TxtNoteserial1.text, (XPDtbTrans.value), _
        (TxtNetValue.text), Picture1, 0, (TxtVAT.text), (txtTotal.text)

    BeginTrans = False
 createVoucher
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
Retrive
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

Private Sub Check2_Click()
If Me.Check2.value = vbUnchecked Then
DcCityToId.BoundText = ""
End If
End Sub

Private Sub Check3_Click()
If Me.Check3.value = vbUnchecked Then
DcbShip.BoundText = ""
End If
End Sub

Private Sub chkItem_Click()
If Me.chkItem.value = vbUnchecked Then
DcboItems.BoundText = ""
TxtItemCode.text = "'"
End If
End Sub

Private Sub chkoWithoutVat_Click()
ReLineGrid
End Sub

Private Sub chkTypeTransport_Click()
If Me.chkTypeTransport.value = vbUnchecked Then
DcbTypeTransport.BoundText = ""
End If
End Sub

 Private Sub Cmd_Click(index As Integer)

    'On Error GoTo ErrTrap
    
    Select Case index
    Dim X As Integer
        
        Case 9
            If Me.TxtModFlg.text = "E" Then
                X = MsgBox("”Ì „ «·€«¡ «· Œ’Ì’ «·Õ«·Ì", vbCritical + vbOKCancel)
                If X = vbOK Then
If Me.TxtModFlg.text = "E" Then
VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 1

        '    Cn.Execute "delete TblTravDueKDet where TravID=" & val(Me.TxtTransID.Text)
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    
                    Cn.Execute " update  TblTripTypesTransport set  allocations=0 where id  in( " & " select noteid from TblTravDueKDet where TravID=" & TxtTransID & ")"
        
        Else

        Exit Sub
            End If
End If
If Check1.value = vbChecked Then
If val(DcCityFromId.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·« Ã«Â „‰"
 Else
MsgBox "Please Select Direction"
End If
DcCityFromId.SetFocus
Exit Sub
End If
End If
If Check2.value = vbChecked Then
If val(DcCityToId.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·« Ã«Â «·Ì"
Else
MsgBox "Please Select Direction"
End If
DcCityToId.SetFocus
Exit Sub
End If
End If
If Check3.value = vbChecked Then
                If val(DcbShip.BoundText) = 0 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ «Œ Ì«— «·”›Ì‰…"
                            Else
                            MsgBox "Please Select Ship"
                            End If
                DcbShip.SetFocus
                Exit Sub
                End If
End If

If chkTypeTransport.value = vbChecked Then
                If val(DcbTypeTransport.BoundText) = 0 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·‰ﬁ·"
                            Else
                            MsgBox "Please Select Ship"
                            End If
                DcbTypeTransport.SetFocus
                Exit Sub
                End If
End If

If chkItem.value = vbChecked Then
                If val(DcboItems.BoundText) = 0 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
                            Else
                            MsgBox "Please Select item"
                            End If
                DcboItems.SetFocus
                Exit Sub
                End If
End If


FillGrid

        Case 0
 
            TxtModFlg.text = "N"
            clear_all Me
        OperatorsID.ListIndex = 0
       OptAlarms(0).value = True
       OptActual(1).value = True
       RdQty(0).value = True
       RdAuto_Manual(1).value = True
       RdAuto_Manual_Click (1)
       Me.XPDtbTrans.value = Date
       RecorddateH.value = ToHijriDate(Date)
            Me.FromDate.value = Date
            Me.ToDate.value = Date
            Check17.value = vbChecked
            Me.Fromdate√H.value = ToHijriDate(Date)
            todateH.value = ToHijriDate(Date)
            Me.Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
            GridInstallments.Enabled = True
            
            If SystemOptions.TripRevenueAuto = True Then
           DcbTypeTransport.BoundText = 1
            End If
            
        Case 1
             If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            TxtModFlg.text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
ReLineGrid
        Case 2
        If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
       If val(Me.Dcbranch.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
            Else
                MsgBox "Select Branch Firstly    ", vbCritical
            End If

            Dcbranch.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
            If val(Me.DBCboClientName2.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·⁄„Ì· «Ê·«", vbCritical
            Else
                MsgBox "Select Customer Firstly    ", vbCritical
            End If

            DBCboClientName2.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        If RdAuto_Manual(1).value = True Then
        If val(Me.DcbTypeTransport.BoundText) = 0 And chkTypeTransport Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ ‰Ê⁄ «·‰ﬁ· «Ê·«", vbCritical
            Else
                MsgBox "Select Type Firstly    ", vbCritical
            End If

            DcbTypeTransport.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
            If val(Me.DcboItems.BoundText) = 0 And chkItem Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ«·’‰› «Ê·«", vbCritical
            Else
                MsgBox "Select Item Firstly    ", vbCritical
            End If

            DcboItems.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
    End If
If TxtNoteSerial.text = "" Then     'ÃœÌœ ›ﬁÿ
                        If Notes_coding(val(my_branch), Me.XPDtbTrans.value) = "error" Then
                            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                        Else
                                       
                                        If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                                        Else
                                             
                                        End If
                        End If
 End If
 Dim Account_Code_dynamic As String
 '  Account_Code_dynamic = get_account_code_branch(2, my_branch)
 '
 '           If Account_Code_dynamic = "NO branch" Then
 '               If SystemOptions.UserInterface = ArabicInterface Then
 '                   MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
 '               Else
 '                   MsgBox "Branch Not Created", vbCritical
 '               End If
'
'                GoTo ErrTrap
'            Else
'
'                If Account_Code_dynamic = "NO account" Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                    Else
'                        MsgBox "Sales Account Not Defined in this Branch", vbCritical
'                    End If
'
'                    GoTo ErrTrap
'
'                End If
'            End If
         If val(txtTotal.text) = 0 Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "Ì—ÃÏ «Œ Ì«— ﬁÌ„… Ê«Õœ… ⁄·Ï «·«ﬁ·"
         Else
         MsgBox "Please Select Value"
         End If
         Exit Sub
         End If
         'DcbTypeTransport.BoundText = 1
         GetAccountTypeTrans val(DcbTypeTransport.BoundText), Account_Code_dynamic
         If Account_Code_dynamic = "" Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "ÌƒÃÏ  ÕœÌœ Õ”«» «·«Ì—«œ«  ·‰Ê⁄ «·‰ﬁ·"
         Else
         MsgBox "Please Select Account"
         End If
         Exit Sub
         End If
         If val(Me.txtDiscount.text) <> 0 Then
         Account_Code_dynamic = get_account_code_branch(160, my_branch)
         If Account_Code_dynamic = "" Or Account_Code_dynamic = "NO account" Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "ÌƒÃÏ  ÕœÌœ Õ”«» «·Œ’Ê„« "
         Else
         MsgBox "Please Select Account"
         End If
         Exit Sub
         End If
         End If
         
 If val(Me.TxtVAT.text) > 0 Then
If GetValueAddedAccount(XPDtbTrans.value, , , 1, 21) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„»Ì⁄« "
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
               Dim TxtNoteSerial1str As String

    If TxtNoteserial1.text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 76, 76)
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    End If
                End If
    End If
    Dim StrTempAccountCode As String
    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName2.BoundText))
    If StrTempAccountCode = "" Then
    MsgBox " «ﬂœ „‰ Õ”«» «·⁄„Ì·"
    Exit Sub
    End If
            SaveData
           
        Case 3
            Undo

        Case 4
            If ChekClodePeriod(XPDtbTrans.value) = True Then
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

            Del_Trans
       
        Case 6
            Unload Me

        Case 7
        If RdAuto_Manual(0).value = True Then
           print_report4
        Else
           print_report
        End If
            '   ViewDataList
        Case 5
                 If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

             FrmProjectSearch.C1Tab1 = 5
             FrmProjectSearch.Caption = "»ÕÀ ›Ê« Ì— «·⁄„·«¡"
             FrmProjectSearch.show vbModal
        Case 21
            RemoveGridRow

        Case 8
          print_report2
           Case 10
          print_report3
        Case 12
              If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report5 , 0
            
            
            
                Case 13
              If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report5 , 1
            
            
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub RemoveGridRow()
    With Me.Grid
        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
End Sub

Private Sub Undo()
   ' On Error GoTo ErrTrap

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
 
Private Sub Command1_Click()
  Dim X As Double
  Dim My_SQL As String
'  Exit Sub

  
    X = MsgBox("”Ì „ «·€«¡ «· Œ’Ì’ «·Õ«·Ì ·Â–« «·⁄„Ì· ›Ì Â–… «·› —… ", vbCritical + vbOKCancel)
                If X = vbOK Then
If Me.TxtModFlg.text = "E" Then
          '  Cn.Execute "delete TblTravDueKDet where TravID=" & val(Me.TxtTransID.Text)
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    
VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 1
    
    
My_SQL = " update TblTripTypesTransport set allocations=0  "
           'My_SQL = My_SQL & "  Where (dbo.notes_all.notetype = 370) AND IsNull(dbo.TblTripTypesTransport.allocations,0) = 0"
           My_SQL = My_SQL & " FROM         dbo.notes_all LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblShipsData ON dbo.notes_all.ShipID = dbo.TblShipsData.id RIGHT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport ON dbo.notes_all.NoteID = dbo.TblTripTypesTransport.NotesallID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
My_SQL = My_SQL & "  Where (dbo.notes_all.notetype = 370) "
My_SQL = My_SQL + "  and (dbo.TblTripTypesTransport.BillDate >=" & SQLDate(Me.FromDate, True) & ""
My_SQL = My_SQL + "  and dbo.TblTripTypesTransport.BillDate <=" & SQLDate(ToDate, True) & " )"
My_SQL = My_SQL & "  and (dbo.notes_all.CusID = " & val(DBCboClientName2.BoundText) & ") "
Cn.Execute My_SQL
MsgBox " „ «·€«¡ «·—»ÿ"

                Else
        
                        Exit Sub
               End If
            
End Sub

Private Sub Command9_Click()
       ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "«À»«  «” Õﬁ«ﬁ «·—Õ·«  ⁄‰ «·› —… „‰  " '& Fromdate√H.value & "  Õ Ï  " & TodateH.value & Chr(13)
des = des & " „‰ " & FromDate.value & "  Õ Ï  " & ToDate.value & CHR(13)
des = des & " ··⁄„Ì· " & DBCboClientName2.text & CHR(13)
des = des & " ‰Ê⁄ «·‰ﬁ· " & DcbTypeTransport.text & CHR(13)
des = des & " «·’‰› " & DcboItems.text & CHR(13)
des = des & " ··›—⁄ " & Dcbranch.text & "     " & txtRemarks

Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "TblTravDueK"
Filedname = "ID"
ContNo = TxtTransID
Notevalue = 0

Notevalue = Format(val(txtTotal.text), "#.##")
 Cn.Execute "Delete Notes where NoteId = " & val(TXTNoteID)
 TxtNoteSerial.text = ""
 TXTNoteID.text = ""
If Me.TxtModFlg = "N" Or TxtNoteSerial.text = "" Then
CreateNotes NoteID, (XPDtbTrans.value), val(Dcbranch.BoundText), 9080, Notevalue, NoteSerial, Me.TxtNoteserial1, tablename, Filedname, ContNo, des, RecorddateH.value
 TXTNoteID.text = NoteID
TxtNoteSerial.text = NoteSerial
Else
sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
sql = sql & ",NoteSerial1='" & Me.TxtNoteserial1 & "',remark='" & des & "'"
  sql = sql & " where NoteID=" & val(TXTNoteID.text)
   Cn.Execute sql
End If
If RdAuto_Manual(0).value = True Then
CREATE_VOUCHER_GE val(TXTNoteID.text), val(Dcbranch.BoundText), user_id, XPDtbTrans.value, des
Else
If SystemOptions.InvoiceTransferJLTotal = False Then
CREATE_VOUCHER_GE2 val(TXTNoteID.text), val(Dcbranch.BoundText), user_id, XPDtbTrans.value, des
Else
CREATE_VOUCHER_GE val(TXTNoteID.text), val(Dcbranch.BoundText), user_id, XPDtbTrans.value, des
End If
End If
rs.Resync adAffectCurrent
End Function
'Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
', NoteDate As Date)
' Dim Notevalue As Variant
'    Dim LngDevID As Long
'    Dim LngDevNO  As Integer
'    Dim StrTempAccountCode As String
'    Dim StrTempDes As String
'    Dim actiondesdes As String
'    Dim SngTemp  As Variant
'    Dim Account_Code_dynamic As String
'    Dim i As Integer
'
'' Dim StrSQL As String
 '
 '        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
 '       Cn.Execute StrSQL, , adExecuteNoRecords
 '
 'LngDevNO = 0
'
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
''
''
'
'    my_branch = BranchID
'
'                                   '«·ÿ—› «·„Ì‰
'                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName2.BoundText))
'
'                              Notevalue = val(TxtTotal.Text)
'
'                                  If Notevalue > 0 Then
'                                    LngDevNO = LngDevNO + 1
'                                      '   actiondesdes = "ﬁÌ„… «·—Õ·… "
'                                           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                                                        GoTo ErrTrap
'                                         End If
'                                         End If
'
'                                   If val(TxtVAT.Text) > 0 Then
'
'                                           Notevalue = val(TxtVAT.Text)
'                                           GetValueAddedAccount XPDtbTrans.value, , StrTempAccountCode, 1, 21
'                                           LngDevNO = LngDevNO + 1
'
'                                           actiondesdes = "Õ”«»  «·ﬁÌ„… «·„÷«›… „»Ì⁄«  "
'                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                                                                 GoTo ErrTrap
'                                                     End If
'                                        End If
'
'                                        If val(TxtTotalValue.Text) > 0 Then
'                                                    LngDevNO = LngDevNO + 1
'                                               Notevalue = val(TxtTotalValue.Text)
'
'                                          ' StrTempAccountCode = get_account_code_branch(2, my_branch)
'                                           GetAccountTypeTrans val(DcbTypeTransport.BoundText), StrTempAccountCode
'                                           actiondesdes = "Õ”«» «Ì—«œ«  «·‰ﬁ· "
'                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                                                                 GoTo ErrTrap
'                                                     End If
'                                        End If
'
'
'
'    updateNotesValueAndNobytext (general_noteid)
'ErrTrap:
'End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date, Optional des As String)
 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim actiondesdes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    
    
   
    my_branch = BranchID

                                   '«·ÿ—› «·„Ì‰
                                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName2.BoundText))

                              Notevalue = Round(txtTotal.text, 3)
                            
                                  If Notevalue > 0 Then
                                    LngDevNO = LngDevNO + 1
                                         actiondesdes = " " & CHR(13) & txtRemarks.text
                                           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, des & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         End If
                                         
                                   If val(TxtVAT.text) > 0 Then
                                    
                                           Notevalue = Round(TxtVAT.text, 3)
                                           GetValueAddedAccount XPDtbTrans.value, , StrTempAccountCode, 1, 21
                                           LngDevNO = LngDevNO + 1
                                           
                                           actiondesdes = "Õ”«»  «·ﬁÌ„… «·„÷«›… „»Ì⁄«  " & CHR(13) & txtRemarks.text
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                        End If
                                        
                                        If Round(TxtNetValue.text, 3) > 0 Then
                                                    LngDevNO = LngDevNO + 1
                                               Notevalue = val(TxtNetValue.text)
                            
                                          ' StrTempAccountCode = get_account_code_branch(2, my_branch)
                                           GetAccountTypeTrans val(DcbTypeTransport.BoundText), StrTempAccountCode
                                           actiondesdes = "Õ”«» «Ì—«œ«  «·‰ﬁ· " & CHR(13) & txtRemarks.text
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                        End If
                                  
                                       If Round(val(txtDiscount.text), 3) > 0 Then
                                                    LngDevNO = LngDevNO + 1
                                               Notevalue = Round(Me.txtDiscount.text, 2)
                            
                                           StrTempAccountCode = get_account_code_branch(160, my_branch)
                                    
                                           actiondesdes = "Õ”«» «·Œ’Ê„«  " & CHR(13) & txtRemarks.text
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                        End If
                                        If Round(val(txtDiscount.text), 3) > 0 Then
                                                    LngDevNO = LngDevNO + 1
                                               Notevalue = val(Me.txtDiscount.text)
                                               StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName2.BoundText))
                                         '  StrTempAccountCode = get_account_code_branch(160, my_branch)
                                           actiondesdes = "Õ”«» «·⁄„Ì· " & CHR(13) & txtRemarks.text
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                        End If

    updateNotesValueAndNobytext (general_noteid)
ErrTrap:
End Function
Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date, Optional des As String)
 Dim Notevalue As Variant
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim actiondesdes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim CarID1 As Double
    Dim CarID2 As Integer
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    my_branch = BranchID

                                   '«·ÿ—› «·„Ì‰
   StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName2.BoundText))

  Notevalue = Round(txtTotal.text, 3)
                               
                                  If Notevalue > 0 Then
                                    LngDevNO = LngDevNO + 1
                                         actiondesdes = " Õ”«» «·⁄„Ì· " & CHR(13) & txtRemarks.text
                                           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , CarID2, , , BranchID, CarID1) = False Then
                                                        GoTo ErrTrap
                                         End If
                                         End If

   
 '«·Œ’Ê„« 
If SystemOptions.traveDiscountFromCustomerDirect = True Then
                     If val(txtDiscount.text) > 0 Then
                                                    LngDevNO = LngDevNO + 1
                                               Notevalue = Round(Me.txtDiscount.text, 3)
                            
                                           StrTempAccountCode = get_account_code_branch(160, my_branch)
                                    
                                           actiondesdes = "Õ”«» «·Œ’Ê„«  " & CHR(13) & txtRemarks.text
                                                    
                                                    
                                                     
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                        End If
End If

                                   If Round(TxtVAT.text, 3) > 0 Then
   
                                          Notevalue = val(TxtVAT.text)
                                           GetValueAddedAccount XPDtbTrans.value, , StrTempAccountCode, 1, 21
                                           LngDevNO = LngDevNO + 1
                                           
                                           actiondesdes = "Õ”«»  «·ﬁÌ„… «·„÷«›… „»Ì⁄«  " & CHR(13) & txtRemarks.text
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , CarID2, , , BranchID, CarID1) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                        End If

   
   With GridInstallments
  For i = 1 To .rows - 1
   If val(.TextMatrix(i, .ColIndex("CarType1"))) = 1 Then
   CarID1 = val(.TextMatrix(i, .ColIndex("CarID")))
   CarID2 = GetCarsFixedAssetID(val(.TextMatrix(i, .ColIndex("CarID"))))
   Else
   CarID1 = 0
   CarID2 = 0
End If
Dim Car As String
Car = (.TextMatrix(i, .ColIndex("Car")))
                                        If Round(txtTotal.text, 3) > 0 Then
                                                    LngDevNO = LngDevNO + 1
                                               Notevalue = val(Me.txtTotal.text)
   If RdQty(0).value = True Then
        If SystemOptions.TransBillPriceByGrid = True Then
                              Notevalue = (.TextMatrix(i, .ColIndex("TotalValue"))) ' * (TXTPrice.Text)   ''-  (.TextMatrix(i, .ColIndex("QtyDownload"))) *  (TXTPrice.Text) * 5 / 100
         Else
                              Notevalue = (TxtPrice.text) * Round(.TextMatrix(i, .ColIndex("QtyDownload")), 3) '- ((.TextMatrix(i, .ColIndex("QtyDownload"))) * (TxtPrice.Text) * 5 / 100)
        End If
        
                        Notevalue = Round(Notevalue, 3)
                        
    Else
            If SystemOptions.TransBillPriceByGrid = True Then
                              Notevalue = Round(.TextMatrix(i, .ColIndex("TotalValue")), 3) ' * (TXTPrice.Text)   ''-  (.TextMatrix(i, .ColIndex("QtyDownload"))) *  (TXTPrice.Text) * 5 / 100
         Else
                              Notevalue = Round(TxtPrice.text, 3) * Round(.TextMatrix(i, .ColIndex("QtyDischarge")), 3) ' - ((.TextMatrix(i, .ColIndex("QtyDischarge"))) * (TxtPrice.Text) * 5 / 100)
        End If
                          '     Notevalue = val(.TextMatrix(i, .ColIndex("TotalValue"))) ' * (TXTPrice.Text)  ''- val(.TextMatrix(i, .ColIndex("QtyDischarge"))) * val(TXTPrice.Text) * 5 / 100
                               Notevalue = Round(Notevalue, 3)
    End If
    
If SystemOptions.traveDiscountFromCustomerDirect = False Then
    If val(Me.txtDiscount.text) <> 0 Then
    Notevalue = Notevalue - (Notevalue / (Round(TxtTotalValue.text, 3))) * Round(Me.txtDiscount.text, 3)
    End If
    
  End If
  
    
                                           GetAccountTypeTrans val(DcbTypeTransport.BoundText), StrTempAccountCode
                                           actiondesdes = "Õ”«» «Ì—«œ«  «·‰ﬁ· " & CHR(13) & txtRemarks.text & Car
      Dim carsrevenacc As String
      Dim carsÒRentacc As String
      Dim RentValue As Double
      Dim RemainValue As Double
     If SystemOptions.CarsRevenuePerOwner = True Then '«·Õ’Ê· ⁄·Ì Õ”«» «·«Ì—«œ  „‰ ‘«‘Â «·„⁄œ« /«·”Ì«—« 
     carsrevenacc = GetCarsREbenueAcountCode(CarID1)
     carsÒRentacc = GetCarsREbenueAcountCode2(CarID1)
     RentValue = Round(.TextMatrix(i, .ColIndex("TxtRent")), 3)
     RemainValue = Notevalue - RentValue
     
     
            If carsrevenacc <> "" Then
            StrTempAccountCode = carsrevenacc
            End If
            
     End If
     RemainValue = Notevalue - RentValue
     
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, RemainValue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , CarID2, , , BranchID, CarID1) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                         If RentValue > 0 Then
                                         LngDevNO = LngDevNO + 1
            If carsÒRentacc = "" Then
            carsÒRentacc = StrTempAccountCode
            End If
            
                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, carsÒRentacc, RentValue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , CarID2, , , BranchID, CarID1) = False Then
                                                                 GoTo ErrTrap
                                                     End If
                                          End If
                                          
                                                     
                                        End If
                                  
Next i
End With
'

                                        'If Round(val(TxtDiscount.Text), 3) > 0 Then
                                        '            LngDevNO = LngDevNO + 1
                                        '       Notevalue = Round(Me.TxtDiscount.Text)
                                        '       StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName2.BoundText))
                                         '  StrTempAccountCode = get_account_code_branch(160, my_branch)
                                        '   actiondesdes = "Õ”«» «·⁄„Ì· " & CHR(13) & txtRemarks.Text
                                        '            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        '                         GoTo ErrTrap
                                        '             End If
                                        'End If
    updateNotesValueAndNobytext (general_noteid)
ErrTrap:
End Function

Private Sub DBCboClientName2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 102
        FrmCustemerSearch.show vbModal

    End If
End Sub

Private Sub DcboItems_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 101
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg.text <> "R" Then
TxtNoteserial1.text = ""
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
   If SystemOptions.IsQrCodePrint Then
    
   End If
  
    
    If SystemOptions.TransBillPriceByGrid = True Then
    TxtPrice.Enabled = False
    Else
    TxtPrice.Enabled = True
    End If
    Dim My_SQL As String
    With GridInstallments
            .ColComboList(.ColIndex("CarType1")) = "#1;„„·Êﬂ… |#2;„„·Êﬂ… ··€Ì— "
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰ "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("Typed")) = "#1;Kg |#2;Kg |#3;Weight"
      End If
           If SystemOptions.UserInterface = ArabicInterface Then
         .ColComboList(.ColIndex("Show")) = "⁄—÷"
        Else
        .ColComboList(.ColIndex("Show")) = "View"
        End If
        
              
                
    End With
        With VSFlexGrid2
            .ColComboList(.ColIndex("CarType1")) = "#1;„„·Êﬂ… |#2;„„·Êﬂ… ··€Ì— "
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰ "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("Typed")) = "#1;Kg |#2;Kg |#3;Weight"
      End If
          
    End With
    
  
                
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
      Dcombos.GetBranches Dcbranch
      Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName2
      Dcombos.GetTypesTransport Me.DcbTypeTransport
      Dcombos.GetShips Me.DcbShip
      Dcombos.GetCitiesDistance Me.DcCityFromId, 0
      Dcombos.GetCitiesDistance Me.DcCityToId, 1
     ' Dcombos.GetItemsNames Me.DcboItems
     If SystemOptions.UserInterface = ArabicInterface Then
           StrSQL = "select ItemID,ItemName from tblitems  where GroupID in ( "
     Else
           StrSQL = "select ItemID,ItemNamee from tblitems  where GroupID in ( "
     End If
                StrSQL = StrSQL & " SELECT     GroupID "
                StrSQL = StrSQL & " From dbo.Groups"
                StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"
   fill_combo DcboItems, StrSQL
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset

StrSQL = "select * From TblTravDueK  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"


    
    If SystemOptions.IsHiddenTransportInv Then
       With GridInstallments
            .ColHidden(.ColIndex("CardNO")) = True
            .ColHidden(.ColIndex("QtyDownload")) = True
            .ColHidden(.ColIndex("CardNO2")) = True
            .ColHidden(.ColIndex("Weight")) = True
            .ColHidden(.ColIndex("Ship")) = True
            .ColHidden(.ColIndex("From")) = True
            .ColHidden(.ColIndex("To")) = True
            .ColHidden(.ColIndex("CarType")) = True
            .ColHidden(.ColIndex("CarType1")) = True
            .ColHidden(.ColIndex("Car")) = True
            '.ColHidden(.ColIndex("CaEmpNamer")) = True
            .ColHidden(.ColIndex("EmpName")) = True
            
       End With
       
           
    End If
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub
Private Sub ChangeLang()
Label4.Caption = "Record No."
Label3.Caption = "Curr.Record"
 Command9.Caption = "Print GL"
 Label1(35).Caption = "GL. No."
 Frame10.Caption = "Data of Accounting"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    Ele(5).Caption = " Customer Invoices"
    Me.Caption = Ele(5).Caption
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Date"
    lbl(13).Caption = "Branch"
    Check1.RightToLeft = False
    Check1.Caption = "From"
    Check2.RightToLeft = False
    Check2.Caption = "To"
    lbl(64).Caption = "Customer"
    Frame8.Caption = "Select Date"
    lbl(0).Caption = "From"
    lbl(14).Caption = "To"
    lbl(72).Caption = "Type"
    lbl(78).Caption = "Item"
    Check3.RightToLeft = False
    Check3.Caption = "Ship"
    Check17.Caption = "Select All"
    Check17.RightToLeft = False
    lbl(2).Caption = "Remarks"
    RdAuto_Manual(0).RightToLeft = False
    RdAuto_Manual(1).RightToLeft = False
    RdAuto_Manual(0).Caption = "Auto"
    RdAuto_Manual(1).Caption = "Manual"
    Cmd(9).Caption = "Add"
    lbl(22).Caption = "Description"
    lbl(10).Caption = "Price"
    lbl(12).Caption = "Total"
    lbl(1).Caption = "Net"
    lbl(11).Caption = "VAT"
    CmdRemove.Caption = "Delete"
    RdQty(0).RightToLeft = False
    RdQty(1).RightToLeft = False
    RdQty(0).Caption = "Loading Qty"
    RdQty(1).Caption = "UnLoading Qty"
    lbl(18).Caption = "Loading Qty"
    lbl(19).Caption = "UnLoading Qty"
    lbl(17).Caption = "Other Cars"
    Cmd(7).Caption = "Print"
    Cmd(8).Caption = "Print Total"
    Cmd(10).Caption = "Print Other"
    With Me.VSFlexGrid2
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("TripNo")) = "Trip No."
        .TextMatrix(0, .ColIndex("TripDate")) = "Date"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("CardNO")) = "Card No."
        .TextMatrix(0, .ColIndex("QtyDownload")) = "Loading Qty"
        .TextMatrix(0, .ColIndex("CardNO2")) = "Card No."
        .TextMatrix(0, .ColIndex("QtyDischarge")) = "UnLoading Qty"
        .TextMatrix(0, .ColIndex("Ship")) = "Ship"
        .TextMatrix(0, .ColIndex("From")) = "From"
        .TextMatrix(0, .ColIndex("To")) = "To"
        .TextMatrix(0, .ColIndex("CarType")) = "Model"
        .TextMatrix(0, .ColIndex("CarType1")) = "Car Type"
        .TextMatrix(0, .ColIndex("Car")) = "Car "
        .TextMatrix(0, .ColIndex("EmpName")) = "Driver "
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
        .TextMatrix(0, .ColIndex("Owner")) = "Owner "
    End With
       With GridInstallments
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("TripNo")) = "Trip No."
        .TextMatrix(0, .ColIndex("TripDate")) = "Date"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("CardNO")) = "Card No."
        .TextMatrix(0, .ColIndex("QtyDownload")) = "Loading Qty"
        .TextMatrix(0, .ColIndex("CardNO2")) = "Card No."
        .TextMatrix(0, .ColIndex("QtyDischarge")) = "UnLoading Qty"
        .TextMatrix(0, .ColIndex("Ship")) = "Ship"
        .TextMatrix(0, .ColIndex("From")) = "From"
        .TextMatrix(0, .ColIndex("To")) = "To"
        .TextMatrix(0, .ColIndex("CarType")) = "Model"
        .TextMatrix(0, .ColIndex("CarType1")) = "Car Type"
        .TextMatrix(0, .ColIndex("Car")) = "Car "
        .TextMatrix(0, .ColIndex("EmpName")) = "Driver "
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
    End With
    Me.C1Tab1.TabCaption(0) = Ele(5).Caption
    Me.C1Tab1.TabCaption(1) = "Other Cars"
    'Me.C1Tab1.TabCaption(2) = "Other Cars"

 'Exit Sub
    
End Sub

Public Sub FillGrid(Optional ByVal ContainerNo As String = "")

    
    Dim i As Double
    Dim Rs3 As ADODB.Recordset
    Dim My_SQL As String
    Set Rs3 = New ADODB.Recordset
 
'My_SQL = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.Note_Value, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, "
'My_SQL = My_SQL & "                       dbo.TblBranchesData.branch_namee, dbo.notes_all.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
'My_SQL = My_SQL & "                       dbo.notes_all.TotalQty, dbo.notes_all.Typed, dbo.notes_all.Total, dbo.notes_all.Price, dbo.notes_all.VehicleType, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
'My_SQL = My_SQL & "                       dbo.notes_all.CarId, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Name AS CarName, dbo.notes_all.general_des, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_ID,"
'My_SQL = My_SQL & "                       dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.notes_all.CityFromId,"
'My_SQL = My_SQL & "                       TblCountriesGovernments_1.GovernmentName, dbo.notes_all.CityToId, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
'My_SQL = My_SQL & "                       dbo.notes_all.allocations ,dbo.notes_all.NoteSerial1"
'My_SQL = My_SQL & "  FROM         dbo.TblCountriesGovernments TblCountriesGovernments_1 RIGHT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.notes_all ON TblCountriesGovernments_1.GovernmentID = dbo.notes_all.CityToId LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
'My_SQL = My_SQL & "  Where (dbo.notes_all.notetype = 370) and (dbo.notes_all.allocations=0  or dbo.notes_all.allocations is null)"
'My_SQL = My_SQL + " and (dbo.notes_all.NoteDate >=" & SQLDate(Me.Fromdate, True) & ""
'My_SQL = My_SQL + " and dbo.notes_all.NoteDate <=" & SQLDate(todate, True) & " )"
'My_SQL = My_SQL + "   order by dbo.notes_all.NoteSerial1 "
My_SQL = " SELECT  notes_all.TxtRent ,   dbo.TblTripTypesTransport.BillDate,dbo.TblTripTypesTransport.ID, dbo.TblTripTypesTransport.NotesallID, dbo.TblTripTypesTransport.CardNO, dbo.TblTripTypesTransport.QtyDownload, "
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport.CardNO2, dbo.TblTripTypesTransport.QtyDischarge, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
My_SQL = My_SQL & "                      dbo.notes_all.NoteDate,dbo.notes_all.RecNo,dbo.notes_all.Weight, dbo.notes_all.NoteSerial1, dbo.notes_all.general_des, dbo.notes_all.CityFromId, TblCountriesGovernments_2.GovernmentName,"
My_SQL = My_SQL & "                      dbo.notes_all.CityToId, TblCountriesGovernments_1.GovernmentName AS GovernmentNameTO, dbo.notes_all.VehicleType, dbo.TBLCarTypes.name,"
My_SQL = My_SQL & "                      dbo.TBLCarTypes.namee, dbo.notes_all.CarId,dbo.notes_all.Price, dbo.TblCarsData.BoardNO, dbo.notes_all.CarID2, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.notes_all.CusID,"
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport.ItemID, dbo.notes_all.TypeTransportID, dbo.notes_all.NoteID, dbo.notes_all.NoteType, dbo.notes_all.branch_no, dbo.notes_all.CarType,"
My_SQL = My_SQL & "                      dbo.notes_all.ShipID,notes_all.ContainerNo, dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_Name,"
My_SQL = My_SQL & "                      dbo.TblEmployee.fullcode , dbo.TblEmployee.Emp_Namee, dbo.notes_all.LeaderName,TblTripTypesTransport.HOverVoucher"
My_SQL = My_SQL & " FROM         dbo.notes_all LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblShipsData ON dbo.notes_all.ShipID = dbo.TblShipsData.id RIGHT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport ON dbo.notes_all.NoteID = dbo.TblTripTypesTransport.NotesallID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
My_SQL = My_SQL & "  Where (dbo.notes_all.notetype = 370) AND IsNull(dbo.TblTripTypesTransport.allocations,0) = 0"

If Trim(ContainerNo) = "" Then
My_SQL = My_SQL + "  and (IsNull(dbo.TblTripTypesTransport.BillDate,notes_all.NoteDate) >=" & SQLDate(Me.FromDate, True) & ""
My_SQL = My_SQL + "  and IsNull(dbo.TblTripTypesTransport.BillDate,notes_all.NoteDate) <=" & SQLDate(ToDate, True) & " )"
My_SQL = My_SQL & "  and (dbo.notes_all.CusID = " & val(DBCboClientName2.BoundText) & ") "
End If
If val(DcbTypeTransport.BoundText) <> 0 Then
My_SQL = My_SQL & "  and (dbo.notes_all.TypeTransportID = " & val(DcbTypeTransport.BoundText) & ") "
End If
If val(DcboItems.BoundText) <> 0 Then
My_SQL = My_SQL & "  and (dbo.TblTripTypesTransport.ItemID = " & val(DcboItems.BoundText) & ") "
End If

  If Trim(ContainerNo) <> "" Then
          My_SQL = My_SQL & " and dbo.notes_all.ContainerNo  = '" & Trim(ContainerNo) & "'"
    End If


If val(DcCityFromId.BoundText) <> 0 Then
My_SQL = My_SQL & "  and (dbo.notes_all.CityFromId= " & val(DcCityFromId.BoundText) & ") "
End If
If val(DcCityToId.BoundText) <> 0 Then
My_SQL = My_SQL & "  and (dbo.notes_all.CityToId= " & val(DcCityToId.BoundText) & ") "
End If
If val(DcbShip.BoundText) <> 0 Then
My_SQL = My_SQL & "  and (dbo.notes_all.ShipID= " & val(DcbShip.BoundText) & ") "
End If

'salimbreak
 'My_SQL = My_SQL & "  and  NoteSerial1  not in (  "
'My_SQL = My_SQL & " SELECT     TripNo FROM         dbo.TblTravDueKDet  "
'If val(TxtTransID) <> 0 Then
'    My_SQL = My_SQL & "  WHERE     (TravID <> " & TxtTransID & ")"
'End If
'  My_SQL = My_SQL & " )"

'salimbreak
If ChkDate.value = vbChecked Then

My_SQL = My_SQL + "   order by dbo.notes_all.NoteDate "
Else
My_SQL = My_SQL + "   order by dbo.notes_all.NoteSerial1 "
End If

 
    Rs3.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    

      With Me.GridInstallments
      Dim Xb As Integer
       .rows = 1
        .Clear flexClearScrollable
        If Rs3.RecordCount > 0 Then
           .rows = Rs3.RecordCount + 1
           Rs3.MoveFirst
            For i = 1 To .rows - 1
            
                       If SystemOptions.UserInterface = ArabicInterface Then
         .ColComboList(.ColIndex("Show")) = "⁄—÷"
        Else
        .ColComboList(.ColIndex("Show")) = "View"
        End If
        
        
        .TextMatrix(i, .ColIndex("Ser")) = i
           .TextMatrix(i, .ColIndex("ShipID")) = (IIf(IsNull(Rs3.Fields("ShipID").value), 0, Rs3.Fields("ShipID").value))
            .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(Rs3.Fields("ID").value), 0, Rs3.Fields("ID").value))
            .TextMatrix(i, .ColIndex("NoteIDA")) = (IIf(IsNull(Rs3.Fields("NoteID").value), 0, Rs3.Fields("NoteID").value))
            .TextMatrix(i, .ColIndex("ContainerNo")) = (IIf(IsNull(Rs3.Fields("ContainerNo").value), "", Rs3.Fields("ContainerNo").value))
            
            .TextMatrix(i, .ColIndex("TxtRent")) = (IIf(IsNull(Rs3.Fields("TxtRent").value), 0, Rs3.Fields("TxtRent").value))

            .TextMatrix(i, .ColIndex("TripNo")) = (IIf(IsNull(Rs3.Fields("NoteSerial1").value), "", Rs3.Fields("NoteSerial1").value))
            .TextMatrix(i, .ColIndex("TripDate")) = (IIf(IsNull(Rs3.Fields("BillDate").value), Rs3.Fields("NoteDate").value, Rs3.Fields("BillDate").value))
            .TextMatrix(i, .ColIndex("BranchID")) = (IIf(IsNull(Rs3.Fields("branch_no").value), 0, Rs3.Fields("branch_no").value))
            .TextMatrix(i, .ColIndex("CardNO")) = (IIf(IsNull(Rs3.Fields("CardNO").value), "", Rs3.Fields("CardNO").value))
            .TextMatrix(i, .ColIndex("QtyDownload")) = (IIf(IsNull(Rs3.Fields("QtyDownload").value), "", Rs3.Fields("QtyDownload").value))
           ' Xb = (IIf(IsNull(Rs3.Fields("Typed").value), 0, Rs3.Fields("Typed").value))
           ' .TextMatrix(i, .ColIndex("Typed")) = Xb + 1
            .TextMatrix(i, .ColIndex("CarType1")) = (IIf(IsNull(Rs3.Fields("CarType").value), 0, Rs3.Fields("CarType").value)) + 1
            .TextMatrix(i, .ColIndex("CardNO2")) = (IIf(IsNull(Rs3.Fields("CardNO2").value), "", Rs3.Fields("CardNO2").value))
            .TextMatrix(i, .ColIndex("FromID")) = (IIf(IsNull(Rs3.Fields("CityFromId").value), 0, Rs3.Fields("CityFromId").value))
            .TextMatrix(i, .ColIndex("ToID")) = (IIf(IsNull(Rs3.Fields("CityToId").value), 0, Rs3.Fields("CityToId").value))
            .TextMatrix(i, .ColIndex("From")) = (IIf(IsNull(Rs3.Fields("GovernmentName").value), "", Rs3.Fields("GovernmentName").value))
            .TextMatrix(i, .ColIndex("To")) = (IIf(IsNull(Rs3.Fields("GovernmentNameTO").value), "", Rs3.Fields("GovernmentNameTO").value))
            .TextMatrix(i, .ColIndex("CarTypeID")) = (IIf(IsNull(Rs3.Fields("VehicleType").value), 0, Rs3.Fields("VehicleType").value))
            If val(.TextMatrix(i, .ColIndex("CarType1"))) = 1 Then
            .TextMatrix(i, .ColIndex("CarType1")) = 1
            End If
            If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 Then
            .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(Rs3.Fields("CarID2").value), 0, Rs3.Fields("CarID2").value))
            .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(Rs3.Fields("BoardNo2").value), "", Rs3.Fields("BoardNo2").value))
            Else
             .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(Rs3.Fields("CarId").value), 0, Rs3.Fields("CarId").value))
            .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(Rs3.Fields("BoardNO").value), "", Rs3.Fields("BoardNO").value))
            End If
            .TextMatrix(i, .ColIndex("QtyDischarge")) = (IIf(IsNull(Rs3.Fields("QtyDischarge").value), 0, Rs3.Fields("QtyDischarge").value))
            .TextMatrix(i, .ColIndex("Remarks")) = (IIf(IsNull(Rs3.Fields("general_des").value), "", Rs3.Fields("general_des").value))
            .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(Rs3.Fields("Price").value), "", Rs3.Fields("Price").value))
            .TextMatrix(i, .ColIndex("RecNo")) = (IIf(IsNull(Rs3.Fields("HOverVoucher").value), IIf(IsNull(Rs3.Fields("RecNo").value), "", Rs3.Fields("RecNo").value), Rs3.Fields("HOverVoucher").value))
            .TextMatrix(i, .ColIndex("Weight")) = (IIf(IsNull(Rs3.Fields("Weight").value), "", Rs3.Fields("Weight").value))
            
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(Rs3.Fields("ShipName").value), "", Rs3.Fields("ShipName").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value))
            .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(Rs3.Fields("Emp_Name").value), (IIf(IsNull(Rs3.Fields("LeaderName").value), "", Rs3.Fields("LeaderName").value)), Rs3.Fields("Emp_Name").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(Rs3.Fields("branch_name").value), "", Rs3.Fields("branch_name").value))
            Else
            
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(Rs3.Fields("NameE").value), "", Rs3.Fields("NameE").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(Rs3.Fields("namee").value), "", Rs3.Fields("namee").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(Rs3.Fields("branch_namee").value), "", Rs3.Fields("branch_namee").value))
            .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(Rs3.Fields("Emp_Namee").value), (IIf(IsNull(Rs3.Fields("LeaderName").value), "", Rs3.Fields("LeaderName").value)), Rs3.Fields("Emp_Namee").value))
            End If
        Rs3.MoveNext

            Next i
 End If
            Rs3.Close
        .RowHeight(-1) = 300
    End With
ReLineGrid
End Sub
Function GetOwnerName(Optional ID As Double) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblVendorCars.ID, dbo.TblVendorCars.CustomerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
sql = sql & " FROM         dbo.TblVendorCars LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblVendorCars.CustomerID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.TblVendorCars.ID = " & ID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
GetOwnerName = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
Else
GetOwnerName = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
End If
Else
GetOwnerName = ""
End If
End Function
Sub FillGrid2()
If Me.TxtModFlg.text <> "E" And Me.TxtModFlg.text <> "N" Then Exit Sub
Dim k As Integer
Dim i As Integer
k = 0
 VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 1
      With Me.GridInstallments
            For i = 1 To .rows - 1
            If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
            VSFlexGrid2.rows = VSFlexGrid2.rows + 1
            
            k = k + 1
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("EmpName")) = .TextMatrix(i, .ColIndex("EmpName"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("ShipID")) = .TextMatrix(i, .ColIndex("ShipID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("ContainerNo")) = .TextMatrix(i, .ColIndex("ContainerNo"))
            
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("NoteID")) = .TextMatrix(i, .ColIndex("NoteID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("TripNo")) = .TextMatrix(i, .ColIndex("TripNo"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("TripDate")) = .TextMatrix(i, .ColIndex("TripDate"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("BranchID")) = .TextMatrix(i, .ColIndex("BranchID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CardNO")) = .TextMatrix(i, .ColIndex("CardNO"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("QtyDownload")) = .TextMatrix(i, .ColIndex("QtyDownload"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CarType1")) = .TextMatrix(i, .ColIndex("CarType1"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CardNO2")) = .TextMatrix(i, .ColIndex("CardNO2"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("FromID")) = .TextMatrix(i, .ColIndex("FromID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("ToID")) = .TextMatrix(i, .ColIndex("ToID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("From")) = .TextMatrix(i, .ColIndex("From"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("To")) = .TextMatrix(i, .ColIndex("To"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CarTypeID")) = .TextMatrix(i, .ColIndex("CarTypeID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CarType1")) = .TextMatrix(i, .ColIndex("CarType1"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CarID")) = .TextMatrix(i, .ColIndex("CarID"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("Car")) = .TextMatrix(i, .ColIndex("Car"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("QtyDischarge")) = .TextMatrix(i, .ColIndex("QtyDischarge"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("Remarks")) = .TextMatrix(i, .ColIndex("Remarks"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("Ship")) = .TextMatrix(i, .ColIndex("Ship"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("CarType")) = .TextMatrix(i, .ColIndex("CarType"))
            VSFlexGrid2.TextMatrix(k, VSFlexGrid2.ColIndex("Branch")) = .TextMatrix(i, .ColIndex("Branch"))
              End If
            Next i
    End With
    ReLineGrid2
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

Private Sub FromDate_Change()
If Me.TxtModFlg.text <> "R" Then
     
         Me.Fromdate√H.value = ToHijriDate(FromDate.value)
       
End If
End Sub

Private Sub Fromdate√H_LostFocus()
     If Me.TxtModFlg.text <> "R" Then
             
            FromDate.value = ToGregorianDate(Fromdate√H.value)
               
        End If

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
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(row, .ColIndex("UnitID")) = code
                .TextMatrix(row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If row = .rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub
Sub Calculte()
If SystemOptions.TransBillPriceByGrid = False Then
If val(Me.TxtPrice.text) = 0 Then Me.TxtPrice.text = 0
If RdQty(1).value = True Then
TxtTotalValue.text = val(Me.TxtPrice.text) * val(TxtQtyDischarge.text)
Else
TxtTotalValue.text = Round(Me.TxtPrice.text, 3) * Round(val(TxtQtyDownload.text), 3)
TxtTotalValue.text = Round(TxtTotalValue.text, 3)
End If
End If
TxtNetValue.text = val(TxtTotalValue.text) - val(txtDiscount.text)


''''''''''''''''''''''''''''''''''''''''''
Dim Percetage2 As Double

 PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, , Percetage2
 'If Percetage2 = 0 Then Percetage2 = 5

'*******************************
            If Percetage2 = 0 Then
                    TxtVAT.text = 0
            Else
                    TxtVAT.text = Round((TxtNetValue.text) * Percetage2 / 100, 3)
            End If
            
If chkoWithoutVat.value = vbChecked Then
TxtVAT.text = 0
End If

txtTotal.text = val(TxtNetValue.text) + val(TxtVAT.text)
txtTotal.text = Round(txtTotal.text, 3)
End Sub
Private Sub ReLineGrid2()
If Me.TxtModFlg.text <> "R" Then
Dim SumVal As Double
Dim i As Integer
Dim sumQtyDischarge As Double
sumQtyDischarge = 0
SumVal = 0
With VSFlexGrid2
For i = 1 To .rows - 1
SumVal = SumVal + Round(val(.TextMatrix(i, .ColIndex("QtyDownload"))), 3)
sumQtyDischarge = sumQtyDischarge + Round(val(.TextMatrix(i, .ColIndex("QtyDischarge"))), 3)
.TextMatrix(i, .ColIndex("Owner")) = GetOwnerName(val(.TextMatrix(i, .ColIndex("CarID"))))
Next i
End With
lbl(20).Caption = SumVal
lbl(21).Caption = sumQtyDischarge
End If
End Sub
Private Sub ReLineGrid()
ReLineGrid2

Dim SumVal As Double
Dim SumPrice As Double
Dim i As Integer
Dim sumQtyDischarge As Double
sumQtyDischarge = 0
SumVal = 0
If Me.TxtModFlg.text <> "R" And RdAuto_Manual(1).value = True Then
With GridInstallments
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And SystemOptions.TransBillPriceByGrid = True Then
            If val(.TextMatrix(i, .ColIndex("QtyDownload"))) <> 0 Then
            .TextMatrix(i, .ColIndex("TotalValue")) = Round(.TextMatrix(i, .ColIndex("QtyDownload")), 3) * Round(.TextMatrix(i, .ColIndex("Value")), 3)
            Else
            .TextMatrix(i, .ColIndex("TotalValue")) = Round(val(.TextMatrix(i, .ColIndex("QtyDischarge"))), 3) * Round(val(.TextMatrix(i, .ColIndex("Value"))), 3)
            End If
            .TextMatrix(i, .ColIndex("TotalValue")) = Round(.TextMatrix(i, .ColIndex("TotalValue")), 3)
Else
.TextMatrix(i, .ColIndex("TotalValue")) = 0
End If
            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
            SumVal = SumVal + Round(val(.TextMatrix(i, .ColIndex("QtyDownload"))), 3)
            SumPrice = SumPrice + Round(.TextMatrix(i, .ColIndex("TotalValue")), 3)
            sumQtyDischarge = sumQtyDischarge + Round((val(.TextMatrix(i, .ColIndex("QtyDischarge")))), 3)
            End If
Next i
End With
TxtQtyDownload.text = Round(SumVal, 3)
TxtQtyDischarge.text = Round(sumQtyDischarge, 3)
If SystemOptions.TransBillPriceByGrid = True Then
TxtTotalValue.text = SumPrice
End If
Calculte
FillGrid2
Else

End If
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
Dim StrSQL  As String
'    StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID,dbo.TblTravDueK.RdQty ,dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
'StrSQL = StrSQL & "                          TblTravDueK.RecordDate ,dbo.TblTravDueK.TotalValue , dbo.TblTravDueK.Vat,dbo.TblTravDueK.TotalValue + TblTravDueK.Vat as NetValue,"
'StrSQL = StrSQL & "                      dbo.TblTravDueKDet.Price,TblTravDueKDet.RecNo,TblTravDueKDet.Weight,"
'StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
'StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
'StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
'StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
'StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueKDet.TypeTrans, dbo.TblTravDueKDet.ShipID,"
'StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.TblTravDueKDet.LeaderName,"
'StrSQL = StrSQL & "                      tc.CusName , tc.VATNO, tc.Address,TblTravDueK.noteserial1"
'StrSQL = StrSQL & " FROM         dbo.TblTravDueKDet LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.TblTravDueKDet.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblTravDueKDet.BranchID = dbo.TblBranchesData.branch_id"
'StrSQL = StrSQL & "                                  LEFT OUTER JOIN dbo.TblTravDueK"
'StrSQL = StrSQL & "                                              ON  dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID"
'StrSQL = StrSQL & "                                              LEFT OUTER JOIN dbo.TblCustemers AS tc"
'StrSQL = StrSQL & "                                              ON  tc.CusId = dbo.TblTravDueK.CusId"
'StrSQL = StrSQL & "   Where 1= 1 and (dbo.TblTravDueKDet.TypeTrans is null or dbo.TblTravDueKDet.TypeTrans=0)  "
'db_createOrUpdateviewSQL "View_TblTravDueKDet", StrSQL


    Dim RsDev As ADODB.Recordset
    Dim RsDev1 As ADODB.Recordset
   ' Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.rows = 1
    If rs.RecordCount < 1 Then
       ' XPTxtCurrent.Caption = 0
        'XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        'Lngid
        If Lngid <> 0 Then
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
    Me.TxtNoteserial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.TxtDesc.text = IIf(IsNull(rs("Descrp").value), "", rs("Descrp").value)
    Me.TxtTransID.text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    RecorddateH.value = IIf(IsNull(rs("recordDateH").value), ToHijriDate(Date), rs("recordDateH").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    FromDate.value = IIf(IsNull(rs("Fromdate").value), Date, rs("Fromdate").value)
    Me.Fromdate√H.value = IIf(IsNull(rs("FromDateh").value), ToHijriDate(Date), rs("FromDateh").value)
    txtContainerNo = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
    ToDate.value = IIf(IsNull(rs("todate").value), Date, rs("todate").value)
    todateH.value = IIf(IsNull(rs("todateH").value), ToHijriDate(Date), rs("todateH").value)
    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    txtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
    TxtQtyDownload.text = IIf(IsNull(rs("QtyDownload").value), "", rs("QtyDownload").value)
    TxtQtyDischarge.text = IIf(IsNull(rs("QtyDischarge").value), "", rs("QtyDischarge").value)
    lbl(20).Caption = IIf(IsNull(rs("QtyDownload2").value), "", rs("QtyDownload2").value)
    lbl(21).Caption = IIf(IsNull(rs("QtyDischarge2").value), "", rs("QtyDischarge2").value)
    TxtVAT.text = IIf(IsNull(rs("VAT").value), "", rs("VAT").value)
    TxtPrice.text = IIf(IsNull(rs("Price").value), "", rs("Price").value)
    TxtTotalValue.text = IIf(IsNull(rs("TotalValue").value), "", rs("TotalValue").value)
    Me.DBCboClientName2.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DcboItems.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
    Me.DcbTypeTransport.BoundText = IIf(IsNull(rs("TypeTransportID").value), "", rs("TypeTransportID").value)
    Me.DcbShip.BoundText = IIf(IsNull(rs("ShipID").value), "", rs("ShipID").value)
    Me.DcCityFromId.BoundText = IIf(IsNull(rs("CityFromId").value), "", rs("CityFromId").value)
    Me.DcCityToId.BoundText = IIf(IsNull(rs("CityToId").value), "", rs("CityToId").value)
    TxtRefNo.text = IIf(IsNull(rs("RefNo").value), "", rs("RefNo").value)
    TxtNetValue.text = IIf(IsNull(rs("NetValue").value), TxtTotalValue.text, rs("NetValue").value)
    txtDiscount.text = IIf(IsNull(rs("Discount").value), 0, rs("Discount").value)
    txtTotal.text = IIf(IsNull(rs("total").value), 0, rs("total").value)
    
        txtVendor.text = IIf(IsNull(rs("Vendor").value), "", rs("Vendor").value)
        txtContractNo.text = IIf(IsNull(rs("ContractNo").value), "", rs("ContractNo").value)
        txtDuDate.value = IIf(IsNull(rs("DuDate").value), Date, rs("DuDate").value)
       
        


         Dim mmm As String
    
    If Not (IsNull(rs("QrCodeImage").value)) Then
        LoadPictureFromDB Picture1, rs, "QrCodeImage", mmm
    Else
     Set Picture1.Picture = Nothing
    End If
    
    If Not IsNull(rs("RdQty").value) Then
    If (rs("RdQty").value) = 1 Then
    RdQty(1).value = True
    Else
    RdQty(0).value = True
    End If
    Else
    RdQty(0).value = True
    End If
    If Not IsNull(rs("RdAuto_Manual").value) Then
    If (rs("RdAuto_Manual").value) = 1 Then
    RdAuto_Manual(1).value = True
    Else
    RdAuto_Manual(0).value = True
    End If
    Else
    RdAuto_Manual(0).value = True
    End If
    
    If Not IsNull(rs("Ch1").value) Then
    If (rs("Ch1").value) = 1 Then
    Me.Check1.value = vbChecked
    Else
    Me.Check1.value = vbUnchecked
    End If
    Else
    Me.Check1.value = vbUnchecked
    End If
    If Not IsNull(rs("Ch2").value) Then
    If (rs("Ch2").value) = 1 Then
    Me.Check2.value = vbChecked
    Else
    Me.Check2.value = vbUnchecked
    End If
    Else
    Me.Check2.value = vbUnchecked
    End If
    
    If Not IsNull(rs("Ch3").value) Then
    If (rs("Ch3").value) = 1 Then
    Me.Check3.value = vbChecked
    Else
    Me.Check3.value = vbUnchecked
    End If
    Else
    Me.Check3.value = vbUnchecked
    End If
     
     If Not IsNull(rs("chkTypeTransport").value) Then
    
    If (rs("chkTypeTransport").value) = True Then
    Me.chkTypeTransport.value = vbChecked
    Else
    Me.chkTypeTransport.value = vbUnchecked
    End If
    Else
    Me.chkTypeTransport.value = vbUnchecked
    End If
    
    
     If Not IsNull(rs("chkoWithoutVat").value) Then
    
                    If (rs("chkoWithoutVat").value) = True Then
                    Me.chkoWithoutVat.value = vbChecked
                    Else
                    Me.chkoWithoutVat.value = vbUnchecked
                    End If
    Else
    Me.chkoWithoutVat.value = vbUnchecked
    End If
    
    
    
    If Not IsNull(rs("chkItem").value) Then
    If (rs("chkItem").value) = True Then
    Me.chkItem.value = vbChecked
    Else
    Me.chkItem.value = vbUnchecked
    End If
    Else
    Me.chkItem.value = vbUnchecked
    End If
    
   ' chkTypeTransport = IIf(Not IsNull(rs!chkTypeTransport), vbChecked, vbUnchecked)
   ' chkItem = IIf(Not IsNull(rs!chkItem), vbChecked, vbUnchecked)
    
  Me.TxtTransID.text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
StrSQL = " SELECT   TblTravDueKDet.TxtRent, notesallid,  dbo.TblTravDueKDet.notesallid,  dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID, dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.Price,dbo.TblTravDueKDet.TotalValue , TblTravDueKDet.RecNo,TblTravDueKDet.Weight,TblTravDueKDet.ContItem,TblTravDueKDet.PurchaseOrderNo,TblTravDueKDet.LocationName,TblTravDueKDet.RentType,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID,TblTravDueKDet.ContainerNo, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueKDet.TypeTrans, dbo.TblTravDueKDet.ShipID,"
StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.TblTravDueKDet.LeaderName"
StrSQL = StrSQL & " FROM         dbo.TblTravDueKDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.TblTravDueKDet.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblTravDueKDet.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "   Where (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") and (dbo.TblTravDueKDet.TypeTrans is null or dbo.TblTravDueKDet.TypeTrans=0)  "
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
        With Me.GridInstallments
            .rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .rows - 1
                       If SystemOptions.UserInterface = ArabicInterface Then
         .ColComboList(.ColIndex("Show")) = "⁄—÷"
        Else
        .ColComboList(.ColIndex("Show")) = "View"
        End If
        
        .TextMatrix(i, .ColIndex("Ser")) = i
           .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
           'RsDetails1("notesallid").value = val(.TextMatrix(i, .ColIndex("NoteIDA")))
           
           .TextMatrix(i, .ColIndex("NoteIDA")) = (IIf(IsNull(RsDev.Fields("notesallid").value), 0, RsDev.Fields("notesallid").value))
           .TextMatrix(i, .ColIndex("TxtRent")) = (IIf(IsNull(RsDev.Fields("TxtRent").value), 0, RsDev.Fields("TxtRent").value))
           
    
           
           .TextMatrix(i, .ColIndex("ContItem")) = (IIf(IsNull(RsDev.Fields("ContItem").value), "", RsDev.Fields("ContItem").value))
           .TextMatrix(i, .ColIndex("PurchaseOrderNo")) = (IIf(IsNull(RsDev.Fields("PurchaseOrderNo").value), "", RsDev.Fields("PurchaseOrderNo").value))
           .TextMatrix(i, .ColIndex("LocationName")) = (IIf(IsNull(RsDev.Fields("LocationName").value), "", RsDev.Fields("LocationName").value))
           .TextMatrix(i, .ColIndex("RentType")) = (IIf(IsNull(RsDev.Fields("RentType").value), "", RsDev.Fields("RentType").value))
           
           .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(RsDev.Fields("LeaderName").value), "", RsDev.Fields("LeaderName").value))
           .TextMatrix(i, .ColIndex("ContainerNo")) = (IIf(IsNull(RsDev.Fields("ContainerNo").value), "", RsDev.Fields("ContainerNo").value))
           
           .TextMatrix(i, .ColIndex("ShipID")) = (IIf(IsNull(RsDev.Fields("ShipID").value), 0, RsDev.Fields("ShipID").value))
           .TextMatrix(i, .ColIndex("TripNo")) = (IIf(IsNull(RsDev.Fields("TripNo").value), "", RsDev.Fields("TripNo").value))
           .TextMatrix(i, .ColIndex("TripDate")) = (IIf(IsNull(RsDev.Fields("TripDate").value), "", RsDev.Fields("TripDate").value))
           .TextMatrix(i, .ColIndex("BranchID")) = (IIf(IsNull(RsDev.Fields("BranchID").value), 0, RsDev.Fields("BranchID").value))
           .TextMatrix(i, .ColIndex("QtyDownload")) = (IIf(IsNull(RsDev.Fields("QtyDownload").value), "", RsDev.Fields("QtyDownload").value))
           .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(RsDev.Fields("Price").value), "", RsDev.Fields("Price").value))
           .TextMatrix(i, .ColIndex("TotalValue")) = (IIf(IsNull(RsDev.Fields("TotalValue").value), "", RsDev.Fields("TotalValue").value))
           .TextMatrix(i, .ColIndex("Weight")) = (IIf(IsNull(RsDev.Fields("Weight").value), "", RsDev.Fields("Weight").value))
           .TextMatrix(i, .ColIndex("RecNo")) = (IIf(IsNull(RsDev.Fields("RecNo").value), "", RsDev.Fields("RecNo").value))
           .TextMatrix(i, .ColIndex("QtyDischarge")) = (IIf(IsNull(RsDev.Fields("QtyDischarge").value), "", RsDev.Fields("QtyDischarge").value))
           .TextMatrix(i, .ColIndex("CarType1")) = (IIf(IsNull(RsDev.Fields("CarType1").value), 1, RsDev.Fields("CarType1").value))
           .TextMatrix(i, .ColIndex("CardNO")) = (IIf(IsNull(RsDev.Fields("CardNO").value), "", RsDev.Fields("CardNO").value))
           .TextMatrix(i, .ColIndex("CardNO2")) = (IIf(IsNull(RsDev.Fields("CardNO2").value), "", RsDev.Fields("CardNO2").value))
           .TextMatrix(i, .ColIndex("Remarks")) = (IIf(IsNull(RsDev.Fields("Remarks").value), "", RsDev.Fields("Remarks").value))
           .TextMatrix(i, .ColIndex("FromID")) = (IIf(IsNull(RsDev.Fields("FromID").value), 0, RsDev.Fields("FromID").value))
          .TextMatrix(i, .ColIndex("From")) = (IIf(IsNull(RsDev.Fields("GovernmentName").value), "", RsDev.Fields("GovernmentName").value))
          .TextMatrix(i, .ColIndex("ToID")) = (IIf(IsNull(RsDev.Fields("ToID").value), 0, RsDev.Fields("ToID").value))
          .TextMatrix(i, .ColIndex("To")) = (IIf(IsNull(RsDev.Fields("ToGovernmentName").value), "", RsDev.Fields("ToGovernmentName").value))
          .TextMatrix(i, .ColIndex("CarTypeID")) = (IIf(IsNull(RsDev.Fields("CarTypeID").value), 0, RsDev.Fields("CarTypeID").value))
          .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(RsDev.Fields("CarID").value), 0, RsDev.Fields("CarID").value))
          If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 Then
          .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(RsDev.Fields("BoardNo2").value), "", RsDev.Fields("BoardNo2").value))
          Else
          .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(RsDev.Fields("BoardNO").value), "", RsDev.Fields("BoardNO").value))
          End If
            .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(RsDev.Fields("NoteID").value), 0, RsDev.Fields("NoteID").value))
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(RsDev.Fields("ShipName").value), "", RsDev.Fields("ShipName").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(RsDev.Fields("name").value), "", RsDev.Fields("name").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(RsDev.Fields("branch_name").value), "", RsDev.Fields("branch_name").value))
         Else
         .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(RsDev.Fields("ShipNameE").value), "", RsDev.Fields("ShipNameE").value))
         .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(RsDev.Fields("namee").value), "", RsDev.Fields("namee").value))
         .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(RsDev.Fields("branch_namee").value), "", RsDev.Fields("branch_namee").value))
        End If
        RsDev.MoveNext
        Next i
        End With
    End If
 RsDev.Close
 ''//////////////////////////////////////
 StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID, dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
 StrSQL = StrSQL & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
 StrSQL = StrSQL & "                     dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
 StrSQL = StrSQL & "                     dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
 StrSQL = StrSQL & "                     TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
 StrSQL = StrSQL & "                     dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueKDet.TypeTrans, dbo.TblTravDueKDet.ShipID,"
 StrSQL = StrSQL & "                     dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.TblTravDueKDet.LeaderName"
 StrSQL = StrSQL & "    FROM         dbo.TblTravDueKDet LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblShipsData ON dbo.TblTravDueKDet.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblTravDueKDet.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "   Where (dbo.TblTravDueKDet.TravID = " & val(Me.TxtTransID.text) & ") and (dbo.TblTravDueKDet.TypeTrans =1)"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
        With Me.VSFlexGrid2
            .rows = .FixedRows + RsDev.RecordCount
          For i = .FixedRows To .rows - 1
           .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
           .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(RsDev.Fields("LeaderName").value), "", RsDev.Fields("LeaderName").value))
           .TextMatrix(i, .ColIndex("ShipID")) = (IIf(IsNull(RsDev.Fields("ShipID").value), 0, RsDev.Fields("ShipID").value))
           .TextMatrix(i, .ColIndex("TripNo")) = (IIf(IsNull(RsDev.Fields("TripNo").value), "", RsDev.Fields("TripNo").value))
           .TextMatrix(i, .ColIndex("TripDate")) = (IIf(IsNull(RsDev.Fields("TripDate").value), "", RsDev.Fields("TripDate").value))
           .TextMatrix(i, .ColIndex("BranchID")) = (IIf(IsNull(RsDev.Fields("BranchID").value), 0, RsDev.Fields("BranchID").value))
           .TextMatrix(i, .ColIndex("QtyDownload")) = (IIf(IsNull(RsDev.Fields("QtyDownload").value), "", RsDev.Fields("QtyDownload").value))
           .TextMatrix(i, .ColIndex("QtyDischarge")) = (IIf(IsNull(RsDev.Fields("QtyDischarge").value), "", RsDev.Fields("QtyDischarge").value))
           .TextMatrix(i, .ColIndex("CarType1")) = (IIf(IsNull(RsDev.Fields("CarType1").value), 1, RsDev.Fields("CarType1").value))
           .TextMatrix(i, .ColIndex("CardNO")) = (IIf(IsNull(RsDev.Fields("CardNO").value), "", RsDev.Fields("CardNO").value))
           .TextMatrix(i, .ColIndex("CardNO2")) = (IIf(IsNull(RsDev.Fields("CardNO2").value), "", RsDev.Fields("CardNO2").value))
           .TextMatrix(i, .ColIndex("Remarks")) = (IIf(IsNull(RsDev.Fields("Remarks").value), "", RsDev.Fields("Remarks").value))
           .TextMatrix(i, .ColIndex("FromID")) = (IIf(IsNull(RsDev.Fields("FromID").value), 0, RsDev.Fields("FromID").value))
           .TextMatrix(i, .ColIndex("From")) = (IIf(IsNull(RsDev.Fields("GovernmentName").value), "", RsDev.Fields("GovernmentName").value))
           .TextMatrix(i, .ColIndex("ToID")) = (IIf(IsNull(RsDev.Fields("ToID").value), 0, RsDev.Fields("ToID").value))
           .TextMatrix(i, .ColIndex("To")) = (IIf(IsNull(RsDev.Fields("ToGovernmentName").value), "", RsDev.Fields("ToGovernmentName").value))
           .TextMatrix(i, .ColIndex("CarTypeID")) = (IIf(IsNull(RsDev.Fields("CarTypeID").value), 0, RsDev.Fields("CarTypeID").value))
           .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(RsDev.Fields("CarID").value), 0, RsDev.Fields("CarID").value))
          If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 Then
           .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(RsDev.Fields("BoardNo2").value), "", RsDev.Fields("BoardNo2").value))
          Else
          .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(RsDev.Fields("BoardNO").value), "", RsDev.Fields("BoardNO").value))
          End If
            .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(RsDev.Fields("NoteID").value), 0, RsDev.Fields("NoteID").value))
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(RsDev.Fields("ShipName").value), "", RsDev.Fields("ShipName").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(RsDev.Fields("name").value), "", RsDev.Fields("name").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(RsDev.Fields("branch_name").value), "", RsDev.Fields("branch_name").value))
         Else
         .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(RsDev.Fields("ShipNameE").value), "", RsDev.Fields("ShipNameE").value))
         .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(RsDev.Fields("namee").value), "", RsDev.Fields("namee").value))
         .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(RsDev.Fields("branch_namee").value), "", RsDev.Fields("branch_namee").value))
        End If
        RsDev.MoveNext
        Next i
        End With
    End If
 RsDev.Close
 ReLineGrid
 ReLineGrid2
    LabCurrRec.Caption = rs.AbsolutePosition
    LabCountRec.Caption = rs.RecordCount
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub GridInstallments_AfterEdit(ByVal row As Long, ByVal Col As Long)
ReLineGrid
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    With GridInstallments
 
         If .ColKey(Col) <> "Select" And .ColKey(Col) <> "Value" And .ColKey(Col) <> "RecNo" And .ColKey(Col) <> "Weight" And .ColKey(Col) <> "Remarks" And .ColKey(Col) <> "ContItem" And .ColKey(Col) <> "PurchaseOrderNo" And .ColKey(Col) <> "LocationName" And .ColKey(Col) <> "RentType" Then
        Cancel = True
        ElseIf .ColKey(Col) = "Value" Then
        If SystemOptions.TransBillPriceByGrid = True Then
        .ComboList = ""
        Else
        Cancel = True
        End If
        ElseIf .ColKey(Col) <> "Value" Then
        .ComboList = ""
        End If
        
 
        
    End With
End Sub

Private Sub GridInstallments_Click()
With GridInstallments
If .ColKey(GridInstallments.Col) = "Show" Then
 FrmTravelTransactions.show
 FrmTravelTransactions.TxtModFlg = "R"
 FrmTravelTransactions.FindRecbyNoteserial1 (GridInstallments.TextMatrix(GridInstallments.row, 5))

End If
End With
End Sub
Public Function FindRecbyNoteserial1(ByVal NoteSerial1 As String)
    On Error GoTo ErrTrap
     
    rs.Find "NoteSerial1='" & NoteSerial1 & "'", , adSearchForward, 1
    
    Retrive
    Exit Function
ErrTrap:

  End Function

Private Sub ISButton2_Click()
 On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            If vbcheck.value = vbUnchecked Then
ShowAttachments TxtNoteserial1 & "-" & TxtTransID.text, "0411201803"
Else
ShowAttachments TxtNoteserial1, "0411201803"
End If


End Sub

Private Sub RdAuto_Manual_Click(index As Integer)
If SystemOptions.TransBillPriceByGrid = False Then
If Me.RdAuto_Manual(0).value = True Then
TxtQtyDownload.Enabled = True
TxtQtyDischarge.Enabled = True
TxtPrice.Enabled = True
Else
TxtQtyDownload.Enabled = False
TxtQtyDischarge.Enabled = False
TxtPrice.Enabled = True
End If
End If
End Sub

Private Sub RdQty_Click(index As Integer)
Calculte
End Sub

Private Sub RecordDateH_LostFocus()
     If Me.TxtModFlg.text <> "R" Then
             
            XPDtbTrans.value = ToGregorianDate(RecorddateH.value)
                     TxtNoteserial1.text = ""
           TxtNoteSerial.text = ""
           
        End If
End Sub

Private Sub ToDate_Change()
If Me.TxtModFlg.text <> "R" Then
     
         todateH.value = ToHijriDate(ToDate.value)
       
End If
End Sub

Private Sub ToDateH_LostFocus()
     If Me.TxtModFlg.text <> "R" Then
             
            ToDate.value = ToGregorianDate(todateH.value)
               
        End If
End Sub

Private Sub DcboItems_Change()
DcboItems_Click (0)
End Sub

Private Sub DcboItems_Click(Area As Integer)
  Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
End Sub

Private Sub txtContainerNo_Change()
If Me.TxtModFlg.text <> "R" Then
    FillGrid txtContainerNo
    ReLineGrid
End If
End Sub

Private Sub txtDiscount_Change()
Calculte
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If
End Sub

Private Sub TxtItemCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 100
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub TxtModFlg_Change()
 
    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False
RecorddateH.Enabled = True
XPDtbTrans.Enabled = True
Dcbranch.Enabled = True

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
    ElseIf Me.TxtModFlg.text = "E" Then
    RecorddateH.Enabled = True
XPDtbTrans.Enabled = True
Dcbranch.Enabled = True
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
RecorddateH.Enabled = False
XPDtbTrans.Enabled = False
Dcbranch.Enabled = False
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub DBCboClientName2_Change()
DBCboClientName2_Click (0)
End Sub

Private Sub DBCboClientName2_Click(Area As Integer)
    Dim Fullcode As String
     Dim Dcombos As New ClsDataCombos
    GetCustomersDetail val(DBCboClientName2.BoundText), , Fullcode, 1
    TxtSearchCode.text = Fullcode
End Sub

Private Sub txtPrice_Change()
Calculte
End Sub

Private Sub TxtQtyDischarge_Change()
Calculte
End Sub

Private Sub TxtQtyDownload_Change()
Calculte
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName2.BoundText = CUSTID
    End If
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
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

 
Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.text <> "R" Then
     
         RecorddateH.value = ToHijriDate(XPDtbTrans.value)
      TxtNoteserial1.text = ""
           TxtNoteSerial.text = ""
           
End If
End Sub
' search for select noteserial
Public Function FindRecbyNoteserial(ByVal NoteSerial1 As Long)
    On Error GoTo ErrTrap
    'Dim rsFinding As ADODB.Recordset
    
   'Set rsFinding = New ADODB.Recordset
    
    'StrSQL = "select * from Notes where notetype = 9080 and noteserial1 = " & noteserial1
    'rsFinding.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs.Find "NoteSerial=" & NoteSerial1, , adSearchForward, 1
    
    Retrive ' (NoteSerial1)
    Exit Function
ErrTrap:

  End Function

