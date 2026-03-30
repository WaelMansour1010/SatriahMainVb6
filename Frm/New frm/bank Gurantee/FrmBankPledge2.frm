VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBankPledge2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»  „œÌœ ÷„«‰ »‰ﬂÌ"
   ClientHeight    =   8250
   ClientLeft      =   6705
   ClientTop       =   1620
   ClientWidth     =   15330
   Icon            =   "FrmBankPledge2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   15330
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8250
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   15330
      _cx             =   27040
      _cy             =   14552
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   645
         Left            =   360
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   6600
         Width           =   14805
         _cx             =   26114
         _cy             =   1138
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   64
            Top             =   150
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   255
            Index           =   0
            Left            =   13200
            TabIndex        =   65
            Top             =   150
            Width           =   1395
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   4185
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   180
            Width           =   1725
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   150
            Width           =   1785
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   2
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   4
            Left            =   2220
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   180
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   690
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   15555
         _cx             =   27437
         _cy             =   1217
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
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
         Caption         =   ""
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
            Left            =   3810
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1755
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
            ButtonImage     =   "FrmBankPledge2.frx":038A
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
            Left            =   690
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
            ButtonImage     =   "FrmBankPledge2.frx":0724
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
            Left            =   2280
            TabIndex        =   24
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
            ButtonImage     =   "FrmBankPledge2.frx":0ABE
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
            Left            =   1215
            TabIndex        =   25
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
            ButtonImage     =   "FrmBankPledge2.frx":0E58
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ·»  „œÌœ ÷„«‰ »‰ﬂÌ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   9720
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   120
            Width           =   5760
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   705
         Left            =   150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   7365
         Width           =   15030
         _cx             =   26511
         _cy             =   1244
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
            Height          =   465
            Index           =   0
            Left            =   13695
            TabIndex        =   31
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":11F2
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
            Height          =   465
            Index           =   1
            Left            =   12285
            TabIndex        =   32
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":7A54
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
            Height          =   465
            Index           =   2
            Left            =   10935
            TabIndex        =   33
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":E2B6
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
            Height          =   465
            Index           =   3
            Left            =   9630
            TabIndex        =   34
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":14B18
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
            Height          =   465
            Index           =   4
            Left            =   8280
            TabIndex        =   35
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":1B37A
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
            Height          =   465
            Index           =   6
            Left            =   150
            TabIndex        =   36
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":21BDC
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
            Height          =   465
            Left            =   3735
            TabIndex        =   37
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":4B7FE
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
            Height          =   465
            Index           =   7
            Left            =   6930
            TabIndex        =   38
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":52060
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
            Height          =   465
            Index           =   9
            Left            =   5370
            TabIndex        =   39
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   820
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
            ButtonImage     =   "FrmBankPledge2.frx":588C2
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   465
            Left            =   1800
            TabIndex        =   72
            Top             =   120
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   820
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
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   4890
         Left            =   0
         TabIndex        =   41
         Top             =   1560
         Width           =   15285
         _cx             =   26961
         _cy             =   8625
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
         Caption         =   "«·»Ì«‰«  «·«”«”Ì… |Õ«·… «·«⁄ „«œ"
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
         Begin C1SizerLibCtl.C1Elastic pnlHeader 
            Height          =   4470
            Left            =   45
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   45
            Width           =   15195
            _cx             =   26802
            _cy             =   7885
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
            Begin VB.TextBox ThirdPartyNameTxt 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   600
               Width           =   8250
            End
            Begin VB.TextBox NotesTxt 
               Alignment       =   1  'Right Justify
               Height          =   2205
               Left            =   270
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   2160
               Width           =   13425
            End
            Begin VB.TextBox Code3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   1320
               Width           =   915
            End
            Begin VB.TextBox Code2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7575
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   1320
               Width           =   915
            End
            Begin VB.TextBox Code1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   12780
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   1320
               Width           =   915
            End
            Begin VB.TextBox NumberTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10035
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   600
               Width           =   3660
            End
            Begin VB.TextBox beneficiaryTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4950
               RightToLeft     =   -1  'True
               TabIndex        =   2
               Top             =   270
               Width           =   3540
            End
            Begin VB.TextBox PledgeValueTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5025
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   960
               Width           =   3465
            End
            Begin VB.TextBox CompetValueTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10035
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   960
               Width           =   3660
            End
            Begin VB.TextBox PledgeMarginTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10035
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   1680
               Width           =   3660
            End
            Begin VB.TextBox CompetTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   270
               Width           =   3465
            End
            Begin MSComCtl2.DTPicker ReqDate 
               Height          =   285
               Left            =   10035
               TabIndex        =   1
               TabStop         =   0   'False
               Top             =   270
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   96403459
               CurrentDate     =   37140
            End
            Begin MSComCtl2.DTPicker ReqTime 
               Height          =   285
               Left            =   12360
               TabIndex        =   0
               Top             =   270
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               CustomFormat    =   "hh:mm:ss"
               Format          =   96403458
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic100 
               Height          =   750
               Left            =   120
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   4680
               Visible         =   0   'False
               Width           =   14820
               _cx             =   26141
               _cy             =   1323
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
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·ÿ—› «·À«·À"
                  Height          =   255
                  Left            =   6330
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   2145
               End
            End
            Begin MSDataListLib.DataCombo applicantDC 
               Bindings        =   "FrmBankPledge2.frx":5F124
               Height          =   315
               Left            =   10035
               TabIndex        =   10
               Top             =   1320
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
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
            Begin MSDataListLib.DataCombo MangerDC 
               Bindings        =   "FrmBankPledge2.frx":5F139
               Height          =   315
               Left            =   5025
               TabIndex        =   12
               Top             =   1320
               Width           =   2580
               _ExtentX        =   4551
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
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
            Begin MSDataListLib.DataCombo GMangerDC 
               Bindings        =   "FrmBankPledge2.frx":5F14E
               Height          =   315
               Left            =   240
               TabIndex        =   14
               Top             =   1320
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
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
            Begin MSComCtl2.DTPicker ReciveDate 
               Height          =   285
               Left            =   7215
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1680
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   96403459
               CurrentDate     =   37140
            End
            Begin XtremeSuiteControls.CheckBox ThirdPartyChk 
               Height          =   240
               Left            =   8640
               TabIndex        =   5
               Top             =   600
               Width           =   1350
               _Version        =   786432
               _ExtentX        =   2381
               _ExtentY        =   423
               _StockProps     =   79
               Caption         =   "÷„«‰ ÿ—› À«·À"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Shape Shape2 
               BorderWidth     =   2
               Height          =   375
               Left            =   240
               Top             =   1680
               Width           =   6855
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "„·«ÕŸ… / Â–Â «·⁄„·Ì…  Õ «Ã «·Ì ÌÊ„Ì ⁄„· „‰ Êﬁ   ﬁœÌ„ «·„” ‰œ ··Õ”«»«   "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   1680
               Width           =   6780
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·«” ·«„"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8625
               TabIndex        =   66
               Top             =   1680
               Width           =   1320
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   255
               Left            =   13890
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   3000
               Width           =   1155
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·ÿ·»"
               Height          =   285
               Index           =   0
               Left            =   11280
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   270
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·ÿ·»"
               Height          =   285
               Left            =   13890
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   270
               Width           =   1170
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” ›Ìœ"
               Height          =   285
               Left            =   8685
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„"
               Height          =   255
               Left            =   13890
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   600
               Width           =   1170
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·÷„«‰"
               Height          =   255
               Left            =   8760
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   990
               Width           =   1035
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·„‰«›”…"
               Height          =   255
               Left            =   13665
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   960
               Width           =   1395
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬁœ„ «·ÿ·»"
               Height          =   255
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1320
               Width           =   1275
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ— ﬁ”„ «·⁄ﬁÊœ"
               Height          =   255
               Left            =   8670
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1320
               Width           =   1275
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œÌ— «·⁄«„"
               Height          =   255
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1320
               Width           =   1035
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Â«„‘ «·÷„«‰"
               Height          =   255
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1680
               Width           =   1275
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‰«›”…"
               Height          =   285
               Left            =   3855
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   270
               Width           =   1050
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   4470
            Left            =   15930
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   45
            Width           =   15195
            _cx             =   26802
            _cy             =   7885
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
               Height          =   3975
               Left            =   120
               TabIndex        =   69
               Tag             =   "1"
               Top             =   120
               Width           =   15015
               _cx             =   26485
               _cy             =   7011
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
               FormatString    =   $"FrmBankPledge2.frx":5F163
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
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   4920
               Width           =   3375
            End
            Begin VB.Label Label110 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   11055
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   4080
               Width           =   3360
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   630
         Left            =   120
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   840
         Width           =   15060
         _cx             =   26564
         _cy             =   1111
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
         Begin VB.TextBox ID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   12180
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   120
            Width           =   1380
         End
         Begin MSComCtl2.DTPicker RecoredDate 
            Height          =   285
            Left            =   9975
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   96403459
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo BranchDC 
            Height          =   315
            Left            =   240
            TabIndex        =   59
            Top             =   120
            Width           =   8265
            _ExtentX        =   14579
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   240
            Index           =   8
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   120
            Width           =   1380
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   11340
            TabIndex        =   61
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   225
            Index           =   24
            Left            =   8700
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "FrmBankPledge2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
 
    SendTopost Me.Name, "TblBankPledge2", "ID", 0, val(Me.BranchDC.BoundText), val(ID.Text), ID.Text
   If Me.TxtModFlg.Text <> "N" And Me.TxtModFlg.Text <> "E" Then
    rs.Resync
  End If
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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.ID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
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
        GRID2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label110.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label110.Caption = "Approved"
                                 End If
                            Label110.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label110.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label110.Caption = "Currently required Approve"
                            End If
                 Label110.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function
Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap

    Select Case Index
        Case 0
Unload FrmInsurancesSearch
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            ThirdPartyChk.value = xtpUnchecked
            ThirdPartyChk_Click
            ID.Text = CStr(new_id("TblBankPledge2", "ID", "", True))
            Me.DCboUserName.BoundText = user_id
            BranchDC.BoundText = branch_id
            GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 1
            Accredit.Caption = ""
            
        Case 1
            Unload FrmInsurancesSearch
            
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            
            End If
         
            If ScreenAproved(val(ID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not modify.This process associated with approvals"
                End If
                Exit Sub
            End If
            TxtModFlg.Text = "E"
            CuurentLogdata
        Case 2
            SaveData
        Case 3
            Undo
        Case 4
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
       If ScreenAproved(val(ID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If
            Del_Action
        Case 5

        Case 6
            Unload Me
        Case 7
            print_report2
        Case 9
            Unload FrmInsurancesSearch
            FrmInsurancesSearch.SendForm = 5
            FrmInsurancesSearch.BankInx = 2
            FrmInsurancesSearch.show
    End Select
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
    
    If DoPremis(Do_Attach, Me.Name, True) = False Then
        Exit Sub
    End If
    ShowAttachments ID.Text, "2809201702"
End Sub

Private Sub CompetValueTxt_KeyPress(KeyAscii As Integer)
   ' KeyAscii = KeyAscii_Num(KeyAscii, CompetValueTxt.Text, 1)
End Sub
Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    'On Error GoTo ErrTrap
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dcombos.GetBranches BranchDC
    Dcombos.GetEmployees applicantDC
    Dcombos.GetEmployees MangerDC
    Dcombos.GetEmployees GMangerDC
    Dcombos.GetUsers DCboUserName
        
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    
    Resize_Form Me
    
    Set rs = New ADODB.Recordset
    
    Dim StrSQL As String
    StrSQL = ""
    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = "SELECT  *  From TblBankPledge2"
    Else
        StrSQL = "SELECT  *  From TblBankPledge2"
    End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    RecoredDate.value = Date
    ReqDate.value = Date
    ReciveDate.value = Date
    
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub
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
Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Accredit.Caption = "Send For Approval"
    Label110.Caption = "Approval Requested By"
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With
    
    Lbl(2).Caption = "Current Record"
    Lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(9).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"
    
    Me.Caption = "Bank Pledge Request Renewal"
    Label1(2).Caption = Me.Caption
    
    Lbl(8).Caption = "Serial"
    Label3.Caption = "Date"
    Lbl(24).Caption = "Branch"
    Label1(0).Caption = "Request Date"
    Label2.Caption = "Request Time"
    Label19.Caption = "Beneficiary"
    Label5.Caption = "Number"
    Label7.Caption = "Competition Value"
    Label6.Caption = "Pledge Value"
    Label11.Caption = "Applicant"
    Label12.Caption = "Competition Dept Manager"
    Label13.Caption = "General Manager"
    Label14.Caption = "Pledge Margin"
    Label15.Caption = "Receiving Date"
    ThirdPartyChk.RightToLeft = False
    ThirdPartyChk.Caption = "Third Party Pledge"
    Label9.Caption = "Third party name"
    Label10.Caption = "Note / This process takes 2 working days from the submission time to the accounting dept"
    Label4.Caption = "Competition"
    Label8.Caption = "Notes"
    Lbl(0).Caption = "By"
    C1Tab1.TabCaption(0) = "Basic Data"
    C1Tab1.TabCaption(1) = "Confirmation Status"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰   "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

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
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 
End Sub
Private Sub PledgeMarginTxt_KeyPress(KeyAscii As Integer)
'    KeyAscii = KeyAscii_Num(KeyAscii, PledgeMarginTxt.Text, 1)
End Sub

Private Sub PledgeValueTxt_KeyPress(KeyAscii As Integer)
'    KeyAscii = KeyAscii_Num(KeyAscii, PledgeValueTxt.Text, 1)
End Sub

Private Sub ThirdPartyChk_Click()
    If ThirdPartyChk.value = xtpChecked Then
        ThirdPartyNameTxt.Enabled = True
    Else
        ThirdPartyNameTxt.Enabled = False
    End If
End Sub
Private Sub TxtModFlg_Change()
    
    On Error GoTo ErrTrap
    
    Select Case Me.TxtModFlg.Text
        Case "R"
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
            ID.locked = True
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            pnlHeader.Enabled = False
        Case "N"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
            pnlHeader.Enabled = True
        Case "E"
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
            ID.locked = True
           pnlHeader.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    MySQL = "SELECT dbo.TblBankPledge2.ID, dbo.TblBankPledge2.RecordDate, dbo.TblBankPledge2.BranchID, dbo.TblBankPledge2.ReqDate, dbo.TblBankPledge2.ReqTime, dbo.TblBankPledge2.beneficiary, dbo.TblBankPledge2.Number, "
    MySQL = MySQL & " dbo.TblBankPledge2.CompetValue, dbo.TblBankPledge2.PledgeValue, dbo.TblBankPledge2.GMangerID, dbo.TblBankPledge2.PledgeMargin, dbo.TblBankPledge2.ReciveDate, dbo.TblBankPledge2.ThirdPartyFlg,"
    MySQL = MySQL & " dbo.TblBankPledge2.ThirdPartyName, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBankPledge2.AppliedID, dbo.TblBankPledge2.MangerID, TblEmployee_1.Emp_Name AS MangerName,"
    MySQL = MySQL & " TblEmployee_1.Emp_Namee AS MangerNamee, TblEmployee_2.Emp_Name AS GMangerName, TblEmployee_2.Emp_Namee AS GMangerNamee, dbo.TblEmployee.Emp_Name AS AppliedName,"
    MySQL = MySQL & " dbo.TblEmployee.Emp_Namee AS AppliedNamee, dbo.TblBankPledge2.Compet"
    MySQL = MySQL & " FROM dbo.TblBranchesData RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblBankPledge2 LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee AS TblEmployee_2 ON dbo.TblBankPledge2.GMangerID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee AS TblEmployee_1 ON dbo.TblBankPledge2.MangerID = TblEmployee_1.Emp_ID ON dbo.TblEmployee.Emp_ID = dbo.TblBankPledge2.AppliedID ON"
    MySQL = MySQL & " dbo.TblBranchesData.branch_id = dbo.TblBankPledge2.BranchID"
    MySQL = MySQL & " Where dbo.TblBankPledge2.ID = " & ID.Text & ""


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repBankPledge2.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repBankPledge2E.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
        'xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        'xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    'xReport.ParameterFields(3).AddCurrentValue user_name

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
   
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    RecoredDate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    BranchDC.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    ReqDate.value = IIf(IsNull(rs("ReqDate").value), Date, rs("ReqDate").value)
    If Not IsNull(rs.Fields("ReqTime").value) Then
        ReqTime.value = FormatDateTime(rs.Fields("ReqTime").value, vbShortTime)
    End If
    beneficiaryTxt.Text = IIf(IsNull(rs("beneficiary").value), "", Trim(rs("beneficiary").value))
    NumberTxt.Text = IIf(IsNull(rs("Number").value), "", Trim(rs("Number").value))
    CompetValueTxt.Text = IIf(IsNull(rs("CompetValue").value), "", Trim(rs("CompetValue").value))
    PledgeValueTxt.Text = IIf(IsNull(rs("PledgeValue").value), "", Trim(rs("PledgeValue").value))
    applicantDC.BoundText = IIf(IsNull(rs("AppliedID").value), "", rs("AppliedID").value)
    MangerDC.BoundText = IIf(IsNull(rs("MangerID").value), "", rs("MangerID").value)
    GMangerDC.BoundText = IIf(IsNull(rs("GMangerID").value), "", rs("GMangerID").value)
    PledgeMarginTxt.Text = IIf(IsNull(rs("PledgeMargin").value), "", Trim(rs("PledgeMargin").value))
    'ReciveDate.value = IIf(IsNull(rs("CompetDate").value), Date, Trim(rs("CompetDate").value))
    CompetTxt.Text = IIf(IsNull(rs("Compet").value), "", Trim(rs("Compet").value))
    If Not IsNull(rs("ThirdPartyFlg").value) Then
        If rs("ThirdPartyFlg").value = True Then
            ThirdPartyChk.value = xtpChecked
        Else
            ThirdPartyChk.value = xtpUnchecked
        End If
    Else
        ThirdPartyChk.value = xtpUnchecked
    End If
    
    ThirdPartyNameTxt.Text = IIf(IsNull(rs("ThirdPartyName").value), "", Trim(rs("ThirdPartyName").value))
    NotesTxt.Text = IIf(IsNull(rs("Notes").value), "", Trim(rs("Notes").value))
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fillapprovData
    Exit Sub
ErrTrap:
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
        
        Select Case Me.TxtModFlg.Text
           Case "N"
                rs.AddNew
                ID.Text = CStr(new_id("TblBankPledge2", "ID", "", True))
        End Select

        rs("ID").value = val(ID.Text)
        rs("RecordDate").value = RecoredDate.value
        rs("BranchID").value = IIf(BranchDC.BoundText = "", Null, BranchDC.BoundText)
        rs("ReqDate").value = ReqDate.value
        rs("ReqTime").value = FormatDateTime(ReqTime.value, vbShortTime)
        rs("beneficiary").value = IIf(beneficiaryTxt.Text = "", Null, beneficiaryTxt.Text)
        rs("Number").value = IIf(NumberTxt.Text = "", Null, NumberTxt.Text)
        rs("CompetValue").value = IIf(CompetValueTxt.Text = "", Null, val(CompetValueTxt.Text))
        rs("PledgeValue").value = IIf(PledgeValueTxt.Text = "", Null, val(PledgeValueTxt.Text))
        rs("AppliedID").value = IIf(applicantDC.BoundText = "", Null, applicantDC.BoundText)
        rs("MangerID").value = IIf(MangerDC.BoundText = "", Null, MangerDC.BoundText)
        rs("GMangerID").value = IIf(GMangerDC.BoundText = "", Null, GMangerDC.BoundText)
        rs("PledgeMargin").value = IIf(PledgeMarginTxt.Text = "", Null, val(PledgeMarginTxt.Text))
        rs("ReciveDate").value = ReciveDate.value
        rs("Compet").value = IIf(CompetTxt.Text = "", Null, CompetTxt.Text)
        If ThirdPartyChk.value = xtpChecked Then
            rs("ThirdPartyFlg").value = True
            rs("ThirdPartyName").value = IIf(ThirdPartyNameTxt.Text = "", Null, ThirdPartyNameTxt.Text)
        Else
            rs("ThirdPartyFlg").value = False
        End If
        
        rs("Notes").value = IIf(NotesTxt.Text = "", Null, NotesTxt.Text)
        rs("UserID").value = IIf(DCboUserName.BoundText = "", Null, DCboUserName.BoundText)
        rs.update

        Dim StrDes As String

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        
        CuurentLogdata
        
        Select Case Me.TxtModFlg.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰ " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg.Text = "R"
        Retrive
    End If
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
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
            rs.find " ID='" & val(ID.Text) & "'", , adSearchForward, adBookmarkFirst
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
Private Sub Del_Action()
  
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
 
    'On Error GoTo ErrTrap
            
        If ID.Text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  «·”Ã· —ﬁ„ " & CHR(13)
                Msg = Msg + (ID.Text) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Delete Recored File No. ?" & CHR(13)
                Msg = Msg + (ID.Text) & CHR(13)
                Msg = Msg + "  Are you sure you want to delete ?"
            End If
        
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
                If Not rs.RecordCount < 1 Then
                    Deletepost Me.Name, "TblBankPledge2", "ID", 0, val(Me.BranchDC.BoundText), val(ID.Text), ID.Text

                    StrSQL = "delete From TblBankPledge2 where  ID =" & val(ID.Text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                 
                    rs.MoveFirst
                    
                    StrSQL = "SELECT  *  From TblBankPledge2 "
                    rs.Close
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    
                    CuurentLogdata "D"
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
       
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "this process Not Aailable"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtModFlg_Change
        Exit Sub
    End If
    TxtModFlg_Change
    Exit Sub
ErrTrap:
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ…  ﬁÌÌ„ «·„ÊŸ›Ì‰ "
        Else
            Msg = "Sorry can't delete data"
        End If
        Msg = Msg & CHR(13) & Err.description
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
End Sub
Private Sub Code1_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Code1.Text, EmpID
        Me.applicantDC.BoundText = EmpID
    End If
End Sub
Private Sub Code2_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Code2.Text, EmpID
        Me.MangerDC.BoundText = EmpID
    End If
End Sub
Private Sub Code3_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Code3.Text, EmpID
        Me.GMangerDC.BoundText = EmpID
    End If
End Sub
Private Sub applicantDC_Change()
    applicantDC_Click (0)
End Sub
Private Sub applicantDC_Click(Area As Integer)
    If val(applicantDC.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , applicantDC.BoundText, EmpCode
    Code1.Text = EmpCode
End Sub
Private Sub MangerDC_Change()
MangerDC_Click (0)
End Sub
Private Sub MangerDC_Click(Area As Integer)
    If val(MangerDC.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , MangerDC.BoundText, EmpCode
    Code2.Text = EmpCode
End Sub
Private Sub GMangerDC_Change()
    GMangerDC_Click (0)
End Sub
Private Sub GMangerDC_Click(Area As Integer)
    If val(GMangerDC.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , GMangerDC.BoundText, EmpCode
    Code3.Text = EmpCode
End Sub
Function CuurentLogdata(Optional Currentmode As String)
    ScreenNameArabic = "ÿ·»  „œÌœ ÷„«‰ »‰ﬂÌ"
    ScreenNameEnglish = "Bank Gurantee Request Extend"
    LogTextA = " ‘«‘… " & ScreenNameArabic & CHR(13) & "„”·”· " & ID.Text & CHR(13) & "   «· «—ÌŒ " & RecoredDate.value & CHR(13) & "   «·›—⁄ " & BranchDC.Text & CHR(13) & " Êﬁ  «·ÿ·»" & ReqTime & CHR(13) & " «—ÌŒ «·ÿ·»" & ReqDate
    LogTextA = LogTextA & CHR(13) & " «”„ «·„” ›Ìœ " & beneficiaryTxt.Text
    
    
    
    'LogTextA = LogTextA & CHR(13) & " «”„ «·„‘—Ê⁄ " & ProjectTxt.Text
    LogTextA = LogTextA & CHR(13) & " —ﬁ„     " & NumberTxt.Text
    LogTextA = LogTextA & CHR(13) & "÷„«‰ «·ÿ—› «·À«·À " & ThirdPartyNameTxt.Text
    LogTextA = LogTextA & CHR(13) & "ﬁÌ„Â «·„‰«›”Â " & CompetValueTxt.Text
    LogTextA = LogTextA & CHR(13) & "ﬁÌ„Â «·÷„«‰ " & PledgeValueTxt.Text
'    LogTextA = LogTextA & CHR(13) & " «—ÌŒ  ﬁœÌ„ «·„‰«›”Â " & CompeDate.value
  '  LogTextA = LogTextA & CHR(13) & "› Õ «·„Ÿ«—Ì› " & OpenEnvelopeTxt.Text

    LogTextA = LogTextA & CHR(13) & "„ﬁœ„ «·ÿ·» " & applicantDC.Text
    LogTextA = LogTextA & CHR(13) & "„œÌ— ﬁ”„ «·⁄ﬁÊœ  " & MangerDC.Text
    LogTextA = LogTextA & CHR(13) & " «·„œÌ— «·⁄«„ " & GMangerDC.Text
    LogTextA = LogTextA & CHR(13) & " «·„·«ÕŸ«    " & NotesTxt.Text
    LogTextA = LogTextA & CHR(13) & "Â«„‘ «·÷„«‰ " & PledgeMarginTxt.Text
    LogTextA = LogTextA & CHR(13) & "   «—ÌŒ «·«” ·«„  " & ReciveDate.value
    
         
    LogTextE = " Screen " & ScreenNameEnglish & CHR(13) & "No " & ID.Text & CHR(13) & "   Date " & RecoredDate.value & CHR(13) & "   Branch " & BranchDC.Text & CHR(13) & "  Order Time" & ReqTime & CHR(13) & " Order Date " & ReqDate
    LogTextE = LogTextE & CHR(13) & "To " & beneficiaryTxt.Text
    'LogTextE = LogTextE & CHR(13) & " Project Name" & ProjectTxt.Text
    LogTextE = LogTextE & CHR(13) & "Project No.  " & NumberTxt.Text
    LogTextE = LogTextE & CHR(13) & "Third party Gurantee " & ThirdPartyNameTxt.Text
    LogTextE = LogTextE & CHR(13) & "  competition  Value" & CompetValueTxt.Text
    LogTextE = LogTextE & CHR(13) & " Gurantee VAlue " & PledgeValueTxt.Text
    'LogTextE = LogTextE & CHR(13) & " competition Date " & CompeDate.value
    'LogTextE = LogTextE & CHR(13) & "envelops opening " & OpenEnvelopeTxt.Text
    LogTextE = LogTextE & CHR(13) & "Margin of guarantee" & PledgeMarginTxt.Text
    LogTextE = LogTextE & CHR(13) & " By " & applicantDC.Text
    LogTextE = LogTextE & CHR(13) & "Director Contracts Department  " & MangerDC.Text
    LogTextE = LogTextE & CHR(13) & "General Manger " & GMangerDC.Text
    LogTextE = LogTextE & CHR(13) & " Remarks   " & NotesTxt.Text
    LogTextE = LogTextE & CHR(13) & "  received date  " & ReciveDate.value
    
               
  ' LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "No " & ID.Text & CHR(13) & "   Date " & RecoredDate.value & CHR(13) & "   Remarks " & NotesTxt.Text
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
    
End Function
