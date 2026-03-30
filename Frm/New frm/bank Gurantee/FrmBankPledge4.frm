VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBankPledge4 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» ÷„«‰ »‰ﬂÌ"
   ClientHeight    =   8190
   ClientLeft      =   6705
   ClientTop       =   1620
   ClientWidth     =   17880
   Icon            =   "FrmBankPledge4.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8190
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   17880
      _cx             =   31538
      _cy             =   14446
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
         Height          =   750
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   6120
         Width           =   17880
         _cx             =   31538
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   9765
            TabIndex        =   50
            Top             =   180
            Width           =   5670
            _ExtentX        =   10001
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   225
            Index           =   0
            Left            =   15750
            TabIndex        =   51
            Top             =   180
            Width           =   1410
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   270
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   210
            Width           =   2220
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   270
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   270
            Index           =   2
            Left            =   7005
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   210
            Width           =   1770
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   450
            Index           =   4
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   210
            Width           =   1605
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   690
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   17850
         _cx             =   31485
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
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   17
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
            ButtonImage     =   "FrmBankPledge4.frx":038A
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
            TabIndex        =   18
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
            ButtonImage     =   "FrmBankPledge4.frx":0724
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
            TabIndex        =   19
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
            ButtonImage     =   "FrmBankPledge4.frx":0ABE
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
            ButtonImage     =   "FrmBankPledge4.frx":0E58
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
            Caption         =   "‰„Ê–Ã ‘—«¡ „‰«›”…"
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
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   120
            Width           =   4800
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   960
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   7110
         Width           =   17880
         _cx             =   31538
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
            Height          =   660
            Index           =   0
            Left            =   16065
            TabIndex        =   26
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":11F2
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
            Height          =   660
            Index           =   1
            Left            =   14295
            TabIndex        =   27
            Top             =   150
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":7A54
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
            Height          =   660
            Index           =   2
            Left            =   12495
            TabIndex        =   28
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":E2B6
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
            Height          =   660
            Index           =   3
            Left            =   10695
            TabIndex        =   29
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":14B18
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
            Height          =   660
            Index           =   4
            Left            =   8850
            TabIndex        =   30
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":1B37A
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
            Height          =   660
            Index           =   6
            Left            =   150
            TabIndex        =   31
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":21BDC
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
            Height          =   660
            Left            =   3735
            TabIndex        =   32
            Top             =   150
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":4B7FE
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
            Height          =   660
            Index           =   7
            Left            =   7140
            TabIndex        =   33
            Top             =   150
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":52060
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
            Height          =   660
            Index           =   9
            Left            =   5370
            TabIndex        =   34
            Top             =   150
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   1164
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
            ButtonImage     =   "FrmBankPledge4.frx":588C2
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
            Height          =   660
            Left            =   1920
            TabIndex        =   56
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1164
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
         Height          =   4260
         Left            =   0
         TabIndex        =   36
         Top             =   1680
         Width           =   17820
         _cx             =   31432
         _cy             =   7514
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
            Height          =   3840
            Left            =   45
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   45
            Width           =   17730
            _cx             =   31274
            _cy             =   6773
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic100 
               Height          =   555
               Left            =   120
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   4035
               Visible         =   0   'False
               Width           =   17475
               _cx             =   30824
               _cy             =   979
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
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   4005
               Left            =   0
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   -120
               Width           =   17835
               _cx             =   31459
               _cy             =   7064
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
               Begin VB.ComboBox PaymentTypeCb 
                  Height          =   315
                  ItemData        =   "FrmBankPledge4.frx":5F124
                  Left            =   11625
                  List            =   "FrmBankPledge4.frx":5F126
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   960
                  Width           =   2025
               End
               Begin VB.TextBox CompNoTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2835
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   600
                  Width           =   1410
               End
               Begin VB.OptionButton OtherRd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√Œ—Ï"
                  Height          =   195
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   240
                  Width           =   825
               End
               Begin VB.OptionButton SuppRd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ê—œ"
                  Height          =   195
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   255
                  Width           =   825
               End
               Begin VB.TextBox Code2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3300
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   1005
                  Width           =   930
               End
               Begin VB.TextBox CompNameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   5490
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   600
                  Width           =   4740
               End
               Begin VB.TextBox NotesTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   2070
                  Left            =   240
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   11
                  Top             =   1560
                  Width           =   15870
               End
               Begin VB.TextBox Code1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9075
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   960
                  Width           =   1155
               End
               Begin VB.TextBox CopyPriceTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   11625
                  RightToLeft     =   -1  'True
                  TabIndex        =   3
                  Top             =   615
                  Width           =   4485
               End
               Begin VB.TextBox OtherTxt 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   240
                  Width           =   4035
               End
               Begin MSComCtl2.DTPicker SumbitDate 
                  Height          =   330
                  Left            =   210
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   615
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   582
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   96403459
                  CurrentDate     =   37140
               End
               Begin MSDataListLib.DataCombo applicantDC 
                  Bindings        =   "FrmBankPledge4.frx":5F128
                  Height          =   315
                  Left            =   5505
                  TabIndex        =   8
                  Top             =   960
                  Width           =   3540
                  _ExtentX        =   6244
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
               Begin MSComCtl2.DTPicker OpenEnvDate 
                  Height          =   315
                  Left            =   14775
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   960
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   96403459
                  CurrentDate     =   37140
               End
               Begin MSDataListLib.DataCombo MangerDC 
                  Bindings        =   "FrmBankPledge4.frx":5F13D
                  Height          =   315
                  Left            =   210
                  TabIndex        =   10
                  Top             =   1005
                  Width           =   3060
                  _ExtentX        =   5398
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
               Begin MSDataListLib.DataCombo DepDC 
                  Bindings        =   "FrmBankPledge4.frx":5F152
                  Height          =   315
                  Left            =   14160
                  TabIndex        =   61
                  Top             =   240
                  Width           =   1965
                  _ExtentX        =   3466
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
               Begin MSComCtl2.DTPicker CompDate 
                  Height          =   315
                  Left            =   11640
                  TabIndex        =   62
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   96403459
                  CurrentDate     =   37140
               End
               Begin MSDataListLib.DataCombo SupplierDC 
                  Bindings        =   "FrmBankPledge4.frx":5F167
                  Height          =   315
                  Left            =   5490
                  TabIndex        =   63
                  Top             =   240
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
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
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ › Õ «·„Ÿ«—Ì›"
                  Height          =   405
                  Left            =   16185
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   975
                  Width           =   1425
               End
               Begin VB.Label Label17 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   210
                  Left            =   16185
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   2280
                  Width           =   1425
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·„‰«›”…"
                  Height          =   225
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   675
                  Width           =   1140
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   195
                  Left            =   16185
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   255
                  Width           =   1425
               End
               Begin VB.Label Label19 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·„‰«›”…"
                  Height          =   195
                  Left            =   13005
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÃÂ… «·ÿ«·»…"
                  Height          =   195
                  Left            =   10245
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   255
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”⁄— «·‰”Œ…"
                  Height          =   225
                  Left            =   16200
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   675
                  Width           =   1425
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„‰«›”…"
                  Height          =   225
                  Left            =   10245
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   675
                  Width           =   1215
               End
               Begin VB.Label Label10 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ﬁœÌ„"
                  Height          =   285
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   675
                  Width           =   1260
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ÊŸ› «·„”ƒÊ·"
                  Height          =   420
                  Left            =   10245
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   990
                  Width           =   1215
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ— ﬁ”„ «·⁄ﬁÊœ"
                  Height          =   420
                  Left            =   4290
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   990
                  Width           =   1140
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   13680
                  TabIndex        =   40
                  Top             =   975
                  Width           =   1080
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   3840
            Left            =   18465
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   45
            Width           =   17730
            _cx             =   31274
            _cy             =   6773
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
               Height          =   3450
               Left            =   120
               TabIndex        =   58
               Tag             =   "1"
               Top             =   135
               Width           =   17415
               _cx             =   30718
               _cy             =   6085
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
               FormatString    =   $"FrmBankPledge4.frx":5F17C
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
            Begin VB.Label Label110 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   120
               Left            =   11055
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   3585
               Width           =   3360
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   150
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   3945
               Width           =   3375
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   630
         Left            =   120
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   855
         Width           =   17595
         _cx             =   31036
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
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   300
            Left            =   13725
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   180
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker RecoredDate 
            Height          =   300
            Left            =   11670
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   180
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
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
            TabIndex        =   2
            Top             =   180
            Width           =   9540
            _ExtentX        =   16828
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
            Caption         =   "«·›—⁄"
            Height          =   255
            Index           =   24
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12735
            TabIndex        =   54
            Top             =   180
            Width           =   1245
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   270
            Index           =   8
            Left            =   15795
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   180
            Width           =   1440
         End
      End
   End
End
Attribute VB_Name = "FrmBankPledge4"
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
 
    SendTopost Me.Name, "TblBankPledge4", "ID", 0, val(Me.BranchDC.BoundText), val(ID.Text), ID.Text
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
            
            SuppRd.value = True
            SuppRd_Click
            ID.Text = CStr(new_id("TblBankPledge4", "ID", "", True))
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
            TxtModFlg.Text = "E"
            If ScreenAproved(val(ID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
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
            print_report
        Case 9
            Unload FrmInsurancesSearch
            FrmInsurancesSearch.SendForm = 6
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
    ShowAttachments ID.Text, "2809201701"
End Sub

Private Sub CopyPriceTxt_KeyPress(KeyAscii As Integer)
'    KeyAscii = KeyAscii_Num(KeyAscii, CopyPriceTxt.Text, 0)
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
    Dcombos.GetEmpDepartments DepDC
    Dcombos.GetCustomersSuppliers 2, SupplierDC
    Dcombos.GetUsers DCboUserName
    
    With PaymentTypeCb
        .Clear
        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem ("‰ﬁœ")
            .AddItem ("‘Ìﬂ „’œﬁ")
            .AddItem ("ÕÊ«·Â »‰ﬂÌ… Õ”» «·„—›ﬁ")
        Else
            .AddItem ("Cash")
            .AddItem ("Certified Check")
            .AddItem ("Bank Transfer")
        End If
    End With
    
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
        StrSQL = "SELECT  *  From TblBankPledge4"
    Else
        StrSQL = "SELECT  *  From TblBankPledge4"
    End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    RecoredDate.value = Date
    CompDate.value = Date
    SumbitDate.value = Date
    OpenEnvDate.value = Date
    
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
    
    Lbl(2).Caption = "Current Record"
    Lbl(4).Caption = "NO. Recordes"
    Accredit.Caption = "Send For Approval"
    Label110.Caption = "Approval Requested By"
    
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With
    
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(9).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"
    
    Me.Caption = "Tender Purchase Form "
    Label1(2).Caption = Me.Caption
    
    Lbl(8).Caption = "Serial"
    Label3.Caption = "Date"
    Lbl(24).Caption = "Branch"
    
    Label2.Caption = "Department"
    Label19.Caption = "Competition Date"
    Label4.Caption = "Requesting Party"
    SuppRd.Caption = "Supplier"
    OtherRd.Caption = "Other"
    Label5.Caption = "Copy Price"
    Label7.Caption = "Competition Name"
    Label14.Caption = "Competition Number"
    Label10.Caption = "Submitting Date"
    Label6.Caption = "Open Envelope Date"
    Label15.Caption = "Payment Type"
    Label11.Caption = "Responsible Employee"
    Label12.Caption = "Manager"
    Label17.Caption = "Notes"
    Lbl(0).Caption = "By"
    C1Tab1.TabCaption(0) = "Basic Data"
    C1Tab1.TabCaption(1) = "Confirmation Status"
    Label17.Caption = "Notes"
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

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
End Sub


Private Sub Code2_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Code2.Text, EmpID
        Me.MangerDC.BoundText = EmpID
    End If
End Sub
Private Sub Code1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Code1.Text, EmpID
        Me.applicantDC.BoundText = EmpID
    End If
End Sub

Private Sub OtherRd_Click()
    chkSuppOrOther
    SupplierDC.BoundText = 0
End Sub

Private Sub SuppRd_Click()
    chkSuppOrOther
    OtherTxt.Text = ""
End Sub
Private Sub chkSuppOrOther()
    If SuppRd.value = True Then
        SupplierDC.Enabled = True
        OtherTxt.Enabled = False
    ElseIf OtherRd.value = True Then
        SupplierDC.Enabled = False
        OtherTxt.Enabled = True
    Else
        SupplierDC.Enabled = False
        OtherTxt.Enabled = False
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
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    MySQL = "SELECT TblBankPledge4.ID, TblBankPledge4.RecordDate, TblBankPledge4.BranchID, TblBankPledge4.DepID, TblBankPledge4.CompetDate, TblBankPledge4.SuppOrOther, TblBankPledge4.SuppID, TblBankPledge4.OtherName,"
    MySQL = MySQL & " TblBankPledge4.CopyPrice, TblBankPledge4.CompetitionName, TblBankPledge4.CompetitionNumber, TblBankPledge4.SubmittingDate, TblBankPledge4.OpenEnvelopeDate, TblBranchesData.branch_name,"
    MySQL = MySQL & " TblBranchesData.branch_namee, TblBankPledge4.AppliedID, TblBankPledge4.MangerID, TblEmployee_1.Emp_Name AS MangerName, TblEmployee_1.Emp_Namee AS MangerNamee,"
    MySQL = MySQL & " TblEmployee.Emp_Name AS AppliedName, TblEmployee.Emp_Namee AS AppliedNamee, TblCustemers.CusName, TblCustemers.CusNamee, TblBankPledge4.Notes, TblBankPledge4.PaymentType, TblBankPledge4.UserID,"
    MySQL = MySQL & " TblEmpDepartments.DepartmentName , TblEmpDepartments.DepartmentNamee"
    MySQL = MySQL & " FROM TblEmployee RIGHT OUTER JOIN"
    MySQL = MySQL & " TblEmployee AS TblEmployee_1 RIGHT OUTER JOIN"
    MySQL = MySQL & " TblCustemers RIGHT OUTER JOIN"
    MySQL = MySQL & " TblBankPledge4 LEFT OUTER JOIN"
    MySQL = MySQL & " TblEmpDepartments ON TblBankPledge4.DepID = TblEmpDepartments.DeparmentID ON TblCustemers.CusID = TblBankPledge4.SuppID ON TblEmployee_1.Emp_ID = TblBankPledge4.MangerID ON"
    MySQL = MySQL & " TblEmployee.Emp_ID = TblBankPledge4.AppliedID LEFT OUTER JOIN"
    MySQL = MySQL & " TblBranchesData ON TblBankPledge4.BranchID = TblBranchesData.branch_id"
    MySQL = MySQL & " Where dbo.TblBankPledge4.ID = " & ID.Text & ""

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repBankPledge4.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repBankPledge4E.rpt"
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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
    DepDC.BoundText = IIf(IsNull(rs("DepID").value), "", rs("DepID").value)
    CompDate.value = IIf(IsNull(rs("CompetDate").value), Date, rs("CompetDate").value)
    If Not IsNull(rs("SuppOrOther").value) Then
        If rs("SuppOrOther").value = 0 Then
            SuppRd.value = True
            OtherRd.value = False
            SupplierDC.BoundText = IIf(IsNull(rs("SuppID").value), "", rs("SuppID").value)
        ElseIf rs("SuppOrOther").value = 1 Then
            SuppRd.value = False
            OtherRd.value = True
            OtherTxt.Text = IIf(IsNull(rs("OtherName").value), "", Trim(rs("OtherName").value))
        End If
    Else
        SuppRd.value = False
        OtherRd.value = False
    End If
    CopyPriceTxt.Text = IIf(IsNull(rs("CopyPrice").value), "", Trim(rs("CopyPrice").value))
    CompNameTxt.Text = IIf(IsNull(rs("CompetitionName").value), "", Trim(rs("CompetitionName").value))
    CompNoTxt.Text = IIf(IsNull(rs("CompetitionNumber").value), "", Trim(rs("CompetitionNumber").value))
    SumbitDate.value = IIf(IsNull(rs("SubmittingDate").value), Date, Trim(rs("SubmittingDate").value))
    OpenEnvDate.value = IIf(IsNull(rs("OpenEnvelopeDate").value), Date, Trim(rs("OpenEnvelopeDate").value))
    PaymentTypeCb.ListIndex = IIf(IsNull(rs("PaymentType").value), -1, Trim(rs("PaymentType").value))
    applicantDC.BoundText = IIf(IsNull(rs("AppliedID").value), "", rs("AppliedID").value)
    MangerDC.BoundText = IIf(IsNull(rs("MangerID").value), "", rs("MangerID").value)
    NotesTxt.Text = IIf(IsNull(rs("Notes").value), "", Trim(rs("Notes").value))
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", Trim(rs("UserID").value))
    
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
                ID.Text = CStr(new_id("TblBankPledge4", "ID", "", True))
        End Select

        rs("ID").value = val(ID.Text)
        rs("RecordDate").value = RecoredDate.value
        rs("BranchID").value = IIf(BranchDC.BoundText = "", Null, BranchDC.BoundText)
        rs("DepID").value = IIf(DepDC.BoundText = "", Null, DepDC.BoundText)
        rs("CompetDate").value = CompDate.value
        If SuppRd.value = True Then
            rs("SuppOrOther").value = 0
            rs("SuppID").value = IIf(SupplierDC.BoundText = "", Null, SupplierDC.BoundText)
            rs("OtherName").value = Null
        ElseIf OtherRd.value = True Then
            rs("SuppOrOther").value = 1
            rs("SuppID").value = Null
            rs("OtherName").value = IIf(OtherTxt.Text = "", Null, OtherTxt.Text)
        End If
        rs("CopyPrice").value = IIf(CopyPriceTxt.Text = "", Null, val(CopyPriceTxt.Text))
        rs("CompetitionName").value = IIf(CompNameTxt.Text = "", Null, CompNameTxt.Text)
        rs("CompetitionNumber").value = IIf(CompNoTxt.Text = "", Null, CompNoTxt.Text)
        rs("SubmittingDate").value = SumbitDate.value
        rs("OpenEnvelopeDate").value = OpenEnvDate.value
        rs("PaymentType").value = PaymentTypeCb.ListIndex
        rs("AppliedID").value = IIf(applicantDC.BoundText = "", Null, applicantDC.BoundText)
        rs("MangerID").value = IIf(MangerDC.BoundText = "", Null, MangerDC.BoundText)
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
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«  " & CHR(13)
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
               Deletepost Me.Name, "TblBankPledge4", "ID", 0, val(Me.BranchDC.BoundText), val(ID.Text), ID.Text
                
                    StrSQL = "delete From TblBankPledge4 where  ID =" & val(ID.Text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                 
                    rs.MoveFirst
                    
                    StrSQL = "SELECT  *  From TblBankPledge4 "
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

 Function CuurentLogdata(Optional Currentmode As String)
ScreenNameArabic = "‰„Ê–Ã ‘—«¡ „‰«›”… "
ScreenNameEnglish = "Bank Gurantee Request"
    LogTextA = " ‘«‘… " & ScreenNameArabic & CHR(13) & "„”·”· " & ID.Text & CHR(13) & "   «· «—ÌŒ " & RecoredDate.value & CHR(13) & "   «·›—⁄ " & BranchDC.Text & CHR(13) & " " & CHR(13) & " «—ÌŒ «·ÿ·»"
    LogTextA = LogTextA & CHR(13) & " «·ﬁ”„   " & DepDC.Text
    'LogTextA = LogTextA & CHR(13) & " «—ÌŒ «·„‰«›”Â" & CompDate.Text
    LogTextA = LogTextA & CHR(13) & " «·ÃÂÂ «·ÿ«·Ì…     " & SupplierDC.Text
    LogTextA = LogTextA & CHR(13) & " ”⁄— «·‰”Œ… " & CopyPriceTxt.Text
    LogTextA = LogTextA & CHR(13) & "«”„ «·„‰«›”Â " & CompNameTxt.Text
    LogTextA = LogTextA & CHR(13) & "—ﬁ„ «·„‰«›”Â " & CompNoTxt.Text
    LogTextA = LogTextA & CHR(13) & " «—ÌŒ › Õ «·„Ÿ«—Ì› " & OpenEnvDate.value
    LogTextA = LogTextA & CHR(13) & " ÿ—Ìﬁ… «·œ›⁄ " & PaymentTypeCb.Text
    LogTextA = LogTextA & CHR(13) & "«·„ÊŸ› «·„”∆Ê·" & applicantDC.Text
    
 
    LogTextA = LogTextA & CHR(13) & "„œÌ— ﬁ”„ «·⁄ﬁÊœ  " & MangerDC.Text
    'LogTextA = LogTextA & CHR(13) & " «·„œÌ— «·⁄«„ " & GMangerDC.Text
    LogTextA = LogTextA & CHR(13) & " «·„·«ÕŸ«    " & NotesTxt.Text
     
         
    LogTextE = " Screen " & ScreenNameEnglish & CHR(13) & "No " & ID.Text & CHR(13) & "   Date " & RecoredDate.value & CHR(13) & "   Branch " & BranchDC.Text & CHR(13) & "  Order Time" & CHR(13) & " Order Date "
    'LogTextE = LogTextE & CHR(13) & "To " & beneficiaryTxt.Text
    'LogTextE = LogTextE & CHR(13) & " Project Name" & ProjectTxt.Text
    'LogTextE = LogTextE & CHR(13) & "Project No.  " & NumberTxt.Text
    'LogTextE = LogTextE & CHR(13) & "Third party Gurantee " & ThirdPartyNameTxt.Text
    'LogTextE = LogTextE & CHR(13) & "  competition  Value" & CompetValueTxt.Text
    'LogTextE = LogTextE & CHR(13) & " Gurantee VAlue " & PledgeValueTxt.Text
    'LogTextE = LogTextE & CHR(13) & " competition Date " & CompeDate.value
    'LogTextE = LogTextE & CHR(13) & "envelops opening " & OpenEnvelopeTxt.Text
    'LogTextE = LogTextE & CHR(13) & "Margin of guarantee" & PledgeMarginTxt.Text
    LogTextE = LogTextE & CHR(13) & " By " & applicantDC.Text
    LogTextE = LogTextE & CHR(13) & "Director Contracts Department  " & MangerDC.Text
    'LogTextE = LogTextE & CHR(13) & "General Manger " & GMangerDC.Text
    LogTextE = LogTextE & CHR(13) & " Remarks   " & NotesTxt.Text
    'LogTextE = LogTextE & CHR(13) & "  received date  " & ReciveDate.value
    
               
  ' LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "No " & ID.Text & CHR(13) & "   Date " & RecoredDate.value & CHR(13) & "   Remarks " & NotesTxt.Text
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
    
End Function

