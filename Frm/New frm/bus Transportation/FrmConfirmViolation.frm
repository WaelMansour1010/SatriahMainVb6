VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmConfirmViolation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«À»«  «·„Œ«·›« "
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   Icon            =   "FrmConfirmViolation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   9810
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8064
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9804
      _cx             =   17304
      _cy             =   14235
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   636
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   9768
         _cx             =   17224
         _cy             =   1111
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
         Caption         =   "   «À»«  «·„Œ«·›«    "
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
            TabIndex        =   11
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   12
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
            ButtonImage     =   "FrmConfirmViolation.frx":038A
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
            TabIndex        =   13
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
            ButtonImage     =   "FrmConfirmViolation.frx":0724
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
            TabIndex        =   14
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
            ButtonImage     =   "FrmConfirmViolation.frx":0ABE
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
            TabIndex        =   15
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
            ButtonImage     =   "FrmConfirmViolation.frx":0E58
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
         Height          =   624
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7332
         Width           =   9636
         _cx             =   16986
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   480
            Index           =   0
            Left            =   8328
            TabIndex        =   2
            Top             =   72
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":11F2
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
            Height          =   480
            Index           =   1
            Left            =   7332
            TabIndex        =   3
            Top             =   72
            Width           =   948
            _ExtentX        =   1667
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":7A54
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
            Height          =   480
            Index           =   2
            Left            =   6324
            TabIndex        =   4
            Top             =   72
            Width           =   1008
            _ExtentX        =   1773
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":E2B6
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
            Height          =   480
            Index           =   3
            Left            =   5292
            TabIndex        =   5
            Top             =   72
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":14B18
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
            Height          =   480
            Index           =   4
            Left            =   4260
            TabIndex        =   6
            Top             =   72
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":1B37A
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
            Height          =   480
            Index           =   6
            Left            =   1248
            TabIndex        =   8
            Top             =   72
            Width           =   936
            _ExtentX        =   1640
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":21BDC
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
            Height          =   480
            Left            =   240
            TabIndex        =   9
            Top             =   72
            Width           =   1008
            _ExtentX        =   1773
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":4B7FE
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
            Height          =   480
            Index           =   7
            Left            =   2184
            TabIndex        =   7
            Top             =   72
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":52060
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
            Height          =   480
            Index           =   5
            Left            =   3240
            TabIndex        =   16
            Top             =   72
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   847
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
            ButtonImage     =   "FrmConfirmViolation.frx":588C2
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6456
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   720
         Width           =   9636
         _cx             =   16986
         _cy             =   11377
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
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   4944
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   120
            Width           =   3492
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   5784
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1920
            Width           =   2652
         End
         Begin VB.TextBox txtValue 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4944
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   5160
            Width           =   3492
         End
         Begin VB.ComboBox cbViolationType 
            Enabled         =   0   'False
            Height          =   288
            Left            =   5904
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   4080
            Width           =   2532
         End
         Begin VB.TextBox txtContractValue 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   4944
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1560
            Width           =   3492
         End
         Begin VB.TextBox contartvalue 
            Alignment       =   1  'Right Justify
            Height          =   372
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.TextBox studentcustom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   960
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.TextBox daycustom 
            Alignment       =   1  'Right Justify
            Height          =   492
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1440
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   492
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2160
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.TextBox t1 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   2760
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.TextBox t2 
            Alignment       =   1  'Right Justify
            Height          =   492
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   3120
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.TextBox t3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   3840
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.TextBox t4 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   4320
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4944
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   4080
            Width           =   972
         End
         Begin VB.TextBox txtAbsenceCount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   6624
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   4440
            Width           =   1812
         End
         Begin VB.TextBox txtDayRate 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4944
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   4440
            Width           =   852
         End
         Begin VB.TextBox txtcount 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4944
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   4800
            Width           =   3492
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   768
            Left            =   4920
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   5520
            Width           =   3492
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   288
            Left            =   6804
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   3348
            Width           =   1632
            _ExtentX        =   2884
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpDateH 
            Height          =   288
            Left            =   4944
            TabIndex        =   37
            Top             =   3348
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   476
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   4944
            TabIndex        =   38
            Top             =   840
            Width           =   3492
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcViolation 
            Height          =   288
            Left            =   4944
            TabIndex        =   39
            Top             =   3720
            Width           =   3492
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcVendor 
            Height          =   288
            Left            =   4944
            TabIndex        =   40
            Top             =   2280
            Width           =   3492
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcContract 
            Height          =   288
            Left            =   4920
            TabIndex        =   41
            Top             =   480
            Width           =   3492
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   612
            Left            =   0
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   5400
            Width           =   4716
            _cx             =   8308
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
            Begin VB.Label XPTxtCurrent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   396
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   144
               Width           =   612
            End
            Begin VB.Label XPTxtCount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   396
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   144
               Width           =   636
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·”Ã· «·Õ«·Ì:"
               Height          =   420
               Index           =   3
               Left            =   3612
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   144
               Width           =   1008
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ⁄œœ «·”Ã·« :"
               Height          =   396
               Index           =   8
               Left            =   912
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   144
               Width           =   1092
            End
         End
         Begin MSDataListLib.DataCombo dcMonth 
            Height          =   288
            Left            =   4944
            TabIndex        =   47
            Top             =   1200
            Width           =   3492
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCar 
            Height          =   288
            Left            =   4944
            TabIndex        =   48
            Top             =   2640
            Width           =   3492
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   312
            Left            =   120
            TabIndex        =   49
            Top             =   6120
            Width           =   2856
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcSchoolFile 
            Height          =   288
            Left            =   4920
            TabIndex        =   50
            Top             =   3000
            Width           =   3588
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   324
            Index           =   16
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   840
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Œ«·›…"
            Height          =   324
            Index           =   0
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   3720
            Width           =   1092
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·„Œ«·›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   8304
            TabIndex        =   67
            Top             =   3348
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   324
            Index           =   1
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   120
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   324
            Index           =   2
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   2280
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄ﬁœ"
            Height          =   324
            Index           =   4
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   480
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·⁄ﬁœ"
            Height          =   204
            Index           =   5
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1920
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„Œ«·›…"
            Height          =   204
            Index           =   6
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   5160
            Width           =   1092
         End
         Begin VB.Image Image2 
            Height          =   5400
            Left            =   -120
            Picture         =   "FrmConfirmViolation.frx":5F124
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4812
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„Œ«·›…"
            Height          =   252
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   4080
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁÌ„… «·⁄ﬁœ"
            Height          =   324
            Index           =   7
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1560
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   324
            Index           =   9
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1200
            Width           =   1092
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÊ„"
            Height          =   252
            Left            =   4944
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1920
            Width           =   732
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   324
            Index           =   10
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   2640
            Width           =   1092
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬁÌ„… «·ÌÊ„"
            Height          =   252
            Left            =   5904
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   4440
            Width           =   612
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «Ì«„ «·€Ì«»"
            Height          =   252
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   4440
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄œœ"
            Height          =   204
            Index           =   11
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   4800
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   204
            Index           =   12
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   5520
            Width           =   1092
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   288
            Index           =   13
            Left            =   3228
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   6132
            Width           =   900
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„œ—”…"
            Height          =   300
            Index           =   14
            Left            =   8592
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3000
            Width           =   828
         End
      End
   End
End
Attribute VB_Name = "FrmConfirmViolation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp2  As ADODB.Recordset
Dim TTP As clstooltip

Private Sub cbViolationType_Change()
 calc_total
End Sub

Private Sub calc_total()

   txtValue.Text = ""
    If cbViolationType.ListIndex = 1 Then
              txtValue.Text = val(txtcount.Text) * Round((val(txtDayRate.Text) / 100) * val(t1.Text), 2)
              ' txtvalue.text = Round((val(daycustom.text) / 100) * val(t1.text), 2)
    ElseIf cbViolationType.ListIndex = 2 Then
             txtValue.Text = val(txtcount.Text) * Round((val(studentcustom.Text) / 100) * val(t2.Text), 2)
    ElseIf cbViolationType.ListIndex = 3 Then
           ' lblTitle.Caption = "Œ’„ »«·ÿ«·»"
            txtValue.Text = val(studentcustom.Text) * val(txtPer.Text)     ' val(t3.text)
    ElseIf cbViolationType.ListIndex = 4 Then
           'lblTitle.Caption = "Œ’„ »«·ÌÊ„"
           txtValue.Text = val(daycustom.Text) * val(txtPer.Text)  ' val(t4.text)
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
            txtID.Text = CStr(new_id("TblConfirmViolation", "ID", "", True))
         '   txtName.SetFocus
         txtcount.Text = "1"
        Case 1
                If ChekClodePeriod(dtpDate.value) = True Then
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
              
            If ISAllowDeleteUpdateContract = False Then
                MsgBox ("·«Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï «·”‰œ »”»» ⁄„· ÿ·» ’—› ⁄·ÌÂ Ê „ «‰‘«¡ «·ﬁÌœ ")
                Exit Sub
            Else
                TxtModFlg.Text = "E"
            End If
            
        Case 2
                 If ISAllowDeleteUpdateContract = False Then
                MsgBox ("·«Ì„ﬂ‰ Õ›Ÿ Â–« «·”‰œ »”»» ⁄„· ÿ·» ’—› ⁄·ÌÂ Ê „ «‰‘«¡ «·ﬁÌœ ")
                Exit Sub
       
            End If
            
            

                                               If ChekClodePeriod(dtpDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            SaveData

        Case 3
            Undo

        Case 4
                                     If ChekClodePeriod(dtpDate.value) = True Then
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
            
           If ISAllowDeleteUpdateContract = False Then
                MsgBox ("·«Ì„ﬂ‰ Õ–› «·”‰œ »”»» ⁄„· ÿ·» ’—› ⁄·ÌÂ")
                Exit Sub
           Else
                Del_Company
           End If
           
        Case 5
                Unload FrmSearch_BasicData
                FrmSearch_BasicData.SendForm = "ConfirmViolation"
                FrmSearch_BasicData.show
        Case 6
            Unload Me
         Case 7
 print_report
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

 

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtID, "16062020001"
End Sub

Private Sub dcCar_Change()
On Error Resume Next
txtDayRate.Text = ""
   Dim str As String, Add As String, cnt As Integer
   If dcCar.BoundText = "" Then Exit Sub
    Set Rs_Temp = New ADODB.Recordset
    str = "  SELECT  D.dayrate FROM     TblAttributionContract H ,  TblVehicleAllocation_Details D"
    str = str & "  Where H.IDAC = d.IDVA And d.Type = 3 And H.DurationID = " & val(dcDuration.BoundText) & "  And d.CarID =  " & val(dcCar.BoundText) & "  And H.IDAC = " & val(dcContract.BoundText)
    str = str & " AND (D.SchoolFileID = " & val(dcSchoolFile.BoundText) & ")  "
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    If Rs_Temp.RecordCount > 0 Then
          txtDayRate.Text = IIf(IsNull(Rs_Temp("dayrate").value), 0, Rs_Temp("dayrate").value)
    End If
    
       Set Rs_Temp2 = New ADODB.Recordset
    Set dcSchoolFile.RowSource = Rs_Temp2
    str = " select  distinct schoolfileID , schoolfile  from  TblAttributionContract H ,TblVehicleAllocation_Details   D where h.IDAC = d.IDVA and Type = 3 and h.IDAC  =   " & val(dcContract.BoundText)
    str = str & " and carid=" & val(dcCar.BoundText)
    fill_combo dcSchoolFile, str
    dcSchoolFile.Refresh
    
    
    
End Sub

Private Sub dcContract_Change()
Dim str As String

Set Rs_Temp2 = New ADODB.Recordset
    Set dcCar.RowSource = Rs_Temp2
    str = " select distinct carid , BoardNo  from  TblAttributionContract H ,TblVehicleAllocation_Details   D where h.IDAC = d.IDVA and Type = 3 and h.IDAC  =   " & val(dcContract.BoundText)
    fill_combo dcCar, str
    dcCar.Refresh
    
    
   Set Rs_Temp2 = New ADODB.Recordset
    Set dcSchoolFile.RowSource = Rs_Temp2
    str = " select  distinct schoolfileID , schoolfile  from  TblAttributionContract H ,TblVehicleAllocation_Details   D where h.IDAC = d.IDVA and Type = 3 and h.IDAC  =   " & val(dcContract.BoundText)
    fill_combo dcSchoolFile, str
    dcSchoolFile.Refresh
    
    
FillMonth
End Sub

Private Sub dcContract_Click(Area As Integer)
Reset_Data
If dcContract.BoundText <> "" Then
   Dim str As String, Add As String, cnt As Integer
   
    Set Rs_Temp = New ADODB.Recordset
    str = "  select IDMC , durationid ,StudentCount ,StudentCustom , ActualDayValue ,DisCount ,AdditionalType  , VendorID , netValue , StartContractDate , EndContractDate   from TblAttributionContract  where IDAC  =  " & val(dcContract.BoundText)
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    If Rs_Temp.RecordCount > 0 Then
            dcDuration.BoundText = IIf(IsNull(Rs_Temp("durationid")), "", Rs_Temp("durationid"))
            dcVendor.BoundText = IIf(IsNull(Rs_Temp("VendorID")), "", Rs_Temp("VendorID"))
            txtContractValue.Text = IIf(IsNull(Rs_Temp("netValue")), 0, Rs_Temp("netValue"))
            
            cnt = DateDiff("d", IIf(IsNull(Rs_Temp("StartContractDate")), Date, Rs_Temp("StartContractDate")), IIf(IsNull(Rs_Temp("EndContractDate")), Date, Rs_Temp("EndContractDate")))
            Text2.Text = cnt
                      
            contartvalue.Text = cnt
            studentcustom.Text = IIf(IsNull(Rs_Temp("StudentCustom")), 0, Rs_Temp("StudentCustom"))
            daycustom.Text = IIf(IsNull(Rs_Temp("ActualDayValue")), 0, Rs_Temp("ActualDayValue"))
    End If
    
    
End If

    

End Sub

Private Sub Reset_Data()


dcDuration.BoundText = ""
dcMonth.BoundText = ""
txtContractValue.Text = ""
Text2.Text = ""
dcVendor.BoundText = ""
dcCar.BoundText = ""
dcViolation.BoundText = ""
cbViolationType.ListIndex = -1
txtPer.Text = ""
txtAbsenceCount.Text = ""
txtDayRate.Text = ""
txtValue.Text = ""




End Sub



Private Sub dcContract_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
    Unload FrmSearch_MinistryContract
    FrmSearch_MinistryContract.SendForm = "ConfirmViolation"
    FrmSearch_MinistryContract.show
End If

End Sub

Private Sub dcDuration_Change()
FillMonth
End Sub

Private Sub FillMonth()
    
Dim str As String
str = "   select ID , Name  from TblDurations_Details  where did =" & val(dcDuration.BoundText)
fill_combo dcMonth, str

End Sub

Private Sub dcSchoolFile_Click(Area As Integer)
On Error Resume Next
txtDayRate.Text = ""
   Dim str As String, Add As String, cnt As Integer
   If dcCar.BoundText = "" Then Exit Sub
    Set Rs_Temp = New ADODB.Recordset
    str = "  SELECT  D.dayrate FROM     TblAttributionContract H ,  TblVehicleAllocation_Details D"
    str = str & "  Where H.IDAC = d.IDVA And d.Type = 3 And H.DurationID = " & val(dcDuration.BoundText) & "  And d.CarID =  " & val(dcCar.BoundText) & "  And H.IDAC = " & val(dcContract.BoundText)
    str = str & " AND (D.SchoolFileID = " & val(dcSchoolFile.BoundText) & ")  "
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    If Rs_Temp.RecordCount > 0 Then
          txtDayRate.Text = IIf(IsNull(Rs_Temp("dayrate").value), 0, Rs_Temp("dayrate").value)
    End If
    
End Sub

Private Sub dcViolation_Change()
If dcViolation.BoundText <> "" Then
   Dim str As String
    Set Rs_Temp = New ADODB.Recordset
    str = " select * from TblViolationTypes where id =  " & val(dcViolation.BoundText)
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
      Dim o  As Integer
    If Rs_Temp.RecordCount > 0 Then
      
       
      t1.Text = IIf(IsNull(Rs_Temp("ps").value), 0, Rs_Temp("ps").value)
      t2.Text = IIf(IsNull(Rs_Temp("pc").value), 0, Rs_Temp("pc").value)
      t3.Text = IIf(IsNull(Rs_Temp("bystudent").value), -1, Rs_Temp("bystudent").value)
      t4.Text = IIf(IsNull(Rs_Temp("byDay").value), -1, Rs_Temp("byDay").value)
      cbViolationType.ListIndex = IIf(IsNull(Rs_Temp("Type").value), -1, Rs_Temp("Type").value)
      o = IIf(IsNull(Rs_Temp("absence").value), 0, Rs_Temp("absence").value)
        If o = 0 Then
             txtAbsenceCount.Enabled = False
             txtPer.Enabled = True
             txtcount.Enabled = True
        Else
             txtAbsenceCount.Enabled = True
             txtPer.Enabled = False
             txtcount.Enabled = False
        End If
    End If
    
   cbViolationType_Change
    
End If
End Sub

Private Sub dtpDate_Change()

        dtpDateH.value = ToHijriDate(dtpDate.value)
End Sub

Private Sub dtpDateH_LostFocus()
        VBA.Calendar = vbCalGreg
dtpDate.value = ToGregorianDate(dtpDateH.value)
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
    Dcombos.GetCustomersSuppliers 2, dcVendor
    
     Dcombos.GetUsers Me.DCboUserName
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    str = " Select id , name from TblDurations "
    fill_combo dcDuration, str

   ' str = " select IDMC , Name  from TblMinistryContract  "
   str = " select IDAC , IDAC from TblAttributionContract  "
    fill_combo dcContract, str
    
    str = " select ID , Name from TblViolationTypes   "
    fill_combo dcViolation, str
    
    With cbViolationType
        If SystemOptions.UserInterface = ArabicInterface Then
               .Clear
        .AddItem ("ﬁÌ„…")
       ' .AddItem ("‰”»… „‰ «Ã„«·Ï «·⁄ﬁœ")
       .AddItem ("‰”»… „‰ «·«Ã— «·ÌÊ„Ï")
       .AddItem ("‰”»… „‰ „Œ’’ «·ÿ«·»")
       .AddItem ("‰Ê⁄ «·Œ’„ »«·ÿ«·»")
       .AddItem ("‰Ê⁄ «·Œ’„ »«·ÌÊ„")
        Else
        .Clear
                 .AddItem ("Value")
     '  .AddItem ("Percent From Contract Total Value")
       .AddItem ("Percent From Day Salary")
       .AddItem ("Percent From Student Custom")
        .AddItem ("")
       .AddItem ("")
        End If
    End With



   If SystemOptions.UserInterface = ArabicInterface Then
    str = "Select ID , Name from tblSchooleFile "
   Else
   str = "Select ID , NameE from tblSchooleFile "
   End If
   fill_combo dcSchoolFile, str

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «À»«  «·„Œ«·›«   "
    LogTexte = " Open Window " & " Confirm  Violation "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
       
    
    Resize_Form Me
    
    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblConfirmViolation "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «À»«  «·„Œ«·›«   "
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

Private Sub Text1_Change()

End Sub

Private Sub txtAbsenceCount_Change()
txtValue.Text = val(txtDayRate.Text) * val(txtAbsenceCount.Text)
End Sub

Private Sub txtCount_Change()
calc_total
End Sub

Private Sub txtDayRate_Change()
'txtValue.text = txtDayRate.text
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «·„Œ«·›«  "
            Else
                Me.Caption = "Violation Types"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.txtID.locked = True
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
            C1Elastic3.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «·„Œ«·›« ( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «·„Œ«·›« ( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
              
            Me.txtID.locked = True
          '  Me.txtName.locked = False
       C1Elastic3.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «·„Œ«·›«  (  ⁄œÌ· )"
            Else
                Me.Caption = "Violation Types(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.txtID.locked = True
           ' Me.txtName.locked = False
       '     Me.XPMTxtRemark.locked = False
       C1Elastic3.Enabled = True
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
    
    
   Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    txtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
  
    txtContractValue.Text = IIf(IsNull(rs("MinistryContractValue").value), "", Trim(rs("MinistryContractValue").value))
  
    dcDuration.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    dcVendor.BoundText = IIf(IsNull(rs("VendorID").value), "", Trim(rs("VendorID").value))
    dcContract.BoundText = IIf(IsNull(rs("MinistryContractID").value), "", Trim(rs("MinistryContractID").value))
    dcViolation.BoundText = IIf(IsNull(rs("ViolationID").value), "", Trim(rs("ViolationID").value))
    cbViolationType.ListIndex = IIf(IsNull(rs("ViolationType").value), -1, Trim(rs("ViolationType").value))

   ' dtpDateH.value = IIf(IsNull(rs("DateH").value), ToHijriDate(Date), Trim(rs("DateH").value))
    dtpDate.value = IIf(IsNull(rs("Date").value), Date, Trim(rs("Date").value))
    dtpDateH.value = ToHijriDate(dtpDate.value)
    
    dcMonth.BoundText = IIf(IsNull(rs("MonthID").value), "", (rs("MonthID").value))
   
    txtAbsenceCount.Text = IIf(IsNull(rs("AbsenceCount").value), "", rs("AbsenceCount").value)
  
    dcCar.BoundText = IIf(IsNull(rs("CarID").value), "", (rs("CarID").value))
    txtcount.Text = IIf(IsNull(rs("Vcount").value), "", rs("Vcount").value)
    txtValue.Text = IIf(IsNull(rs("Value").value), "", Trim(rs("Value").value))
    txtRemarks.Text = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
      dcSchoolFile.BoundText = IIf(IsNull(rs("SchoolID").value), "", (rs("SchoolID").value))
     
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub




Private Sub TxtName_GotFocus()
On Error Resume Next
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
 SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtPer_Change()
'  txtValue.text = ""
'
'    If cbViolationType.ListIndex = 1 Then
'
'             txtValue.text = Round((val(txtDayRate.text) / 100) * val(txtPer.text), 2)
'    ElseIf cbViolationType.ListIndex = 2 Then
'            txtValue.text = Round((val(studentcustom.text) / 100) * val(txtPer.text), 2)
'    ElseIf cbViolationType.ListIndex = 3 Then
'            txtValue.text = Round(val(studentcustom.text) * val(txtPer.text), 2)      ' val(t3.text)
''    ElseIf cbViolationType.ListIndex = 4 Then
'            txtValue.text = Round(val(daycustom.text) * val(txtPer.text), 2)    ' val(t4.text)
'    End If
calc_total
    
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
    
        If dcDuration.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ «Œ — «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDuration.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If

        If dcVendor.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·„ ⁄Âœ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcVendor.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If
        
       If dcContract.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·⁄ﬁœ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcContract.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If
        
        
         If dcMonth.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcMonth.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If
        
        
         If dcCar.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·„⁄œÂ/«·”Ì«—… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcCar.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If
        
         If dcSchoolFile.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·„œ—”… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcSchoolFile.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If
        
        
        If val(txtValue.Text) = 0 Then
            MsgBox "„‰ ›÷·ﬂ  «œŒ· ﬁÌ„… «·„Œ«·›…  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        Select Case Me.TxtModFlg.Text
            Case "N"
                           
            rs.AddNew
            txtID.Text = CStr(new_id("TblConfirmViolation", "ID", "", True))
            Case "E"
              '  StrSQL = "select * From  TblViolationTypes where Name='" & Trim(txtName.text) & "'"
               
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(txtID.Text)
        rs("DurationID").value = IIf(dcDuration.BoundText = "", Null, dcDuration.BoundText)
        rs("VendorID").value = IIf(dcVendor.BoundText = "", Null, dcVendor.BoundText)
        rs("MinistryContractID").value = IIf(dcContract.BoundText = "", Null, dcContract.BoundText)
        rs("ViolationID").value = IIf(dcViolation.BoundText = "", Null, dcViolation.BoundText)
        rs("ViolationType").value = IIf(cbViolationType.ListIndex = -1, Null, cbViolationType.ListIndex)
        rs("Value") = IIf(IsNumeric(txtValue.Text), val(txtValue.Text), 0)
        rs("MinistryContractValue") = IIf(IsNumeric(txtContractValue.Text), val(txtContractValue.Text), 0)
        rs("Date") = IIf(IsNull(dtpDate.value), Date, dtpDate.value)
        rs("DateH") = IIf(IsNull(dtpDateH.value), ToHijriDate(Date), dtpDateH.value)
        rs("CreationDate") = Date
        rs("MonthID").value = IIf(dcMonth.BoundText = "", Null, dcMonth.BoundText)
        rs("UserID") = user_id
        rs("CarID") = IIf(dcCar.BoundText = "", Null, dcCar.BoundText)
        rs("AbsenceCount").value = val(txtAbsenceCount.Text)
        rs("Vcount").value = val(txtcount.Text)
        rs("remarks").value = IIf(txtRemarks.Text = "", Null, txtRemarks.Text)
        rs("SchoolID").value = IIf(dcSchoolFile.BoundText = "", Null, dcSchoolFile.BoundText)
        
        rs.update
        Cn.CommitTrans
        
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       CuurentLogdata

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
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

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
            rs.find "ID='" & val(txtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
            
        If txtID.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«  «À»«  «·„Œ«·›… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                  StrSQL = "delete From TblConfirmViolation where  ID =" & val(txtID.Text)
                  Cn.Execute StrSQL, , adExecuteNoRecords
                   
                   CuurentLogdata ("D")
                
                   StrSQL = "SELECT  *  From TblConfirmViolation "
                   Set rs = New ADODB.Recordset
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
Public Function CheckScoolId(ID As Double, SchoolFileID As Double) As Boolean
     Dim Rs_Temp2 As New ADODB.Recordset
        Dim str As String
        CheckScoolId = False
        str = " select schoolfileID  from "
        str = str & " dbo.TblVehicleAllocation_Details"
        str = str & " Where ID =" & ID
        Set Rs_Temp2 = New ADODB.Recordset
        Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic
        Dim innerschoolfileID As Double
        If Rs_Temp2.RecordCount > 0 Then
                innerschoolfileID = IIf(IsNull(Rs_Temp2("schoolfileID").value), 0, Rs_Temp2("schoolfileID").value)
         If SchoolFileID = innerschoolfileID Then
         CheckScoolId = True
         Else
         CheckScoolId = False
         End If
         
       End If
       Rs_Temp2.Close
       
   
End Function
Public Function ISAllowDeleteUpdateContract() As Boolean
        Dim EntryCreated As Boolean
        Dim str As String
        
        str = " SELECT  * from tblexchangerequest H , TblExchangeReques_Detailst D where H.id = d.HID and H.DurationID = " & val(dcDuration.BoundText) & "  and H.Month =" & val(dcMonth.BoundText) & " and d.CusID= " & val(dcVendor.BoundText) & " and d.boardno = '" & dcCar.Text & "'"
        str = str & "Order by HID"
        Set Rs_Temp2 = New ADODB.Recordset
        Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic
        If Rs_Temp2.RecordCount > 0 Then
        'EntryCreated
             Dim strLinked As String
        Dim i As Integer
        strLinked = ""
        Dim insid As Double
               For i = 1 To Rs_Temp2.RecordCount
               EntryCreated = IIf(IsNull(Rs_Temp2("EntryCreated").value), 0, Rs_Temp2("EntryCreated").value)
               insid = IIf(IsNull(Rs_Temp2("insid").value), 0, Rs_Temp2("insid").value)
  '    If CheckScoolId(insid, val(dcSchoolFile.BoundText)) = False Then
  '    ISAllowDeleteUpdateContract = True
  '    Exit Function
  '   End If
               
                  
       strLinked = strLinked & CHR(13) & " —ﬁ„ ÿ·» «·’—› : " & IIf(IsNull(Rs_Temp2("HID").value), 0, Rs_Temp2("HID").value)
       strLinked = strLinked & "   «·⁄ﬁœ —ﬁ„    " & IIf(IsNull(Rs_Temp2("IDAC").value), 0, Rs_Temp2("IDAC").value) & IIf(IsNull(Rs_Temp2("EntryCreated").value), " »œÊ‰ ﬁÌœ", "    ·Â ﬁÌœ   ")
       Rs_Temp2.MoveNext
       Next i
       If strLinked <> "" Then
       MsgBox " „— »ÿ »ÿ·»«  «·’—› «· «·Ì…  : " & strLinked
       End If
       
      
   
        

        If EntryCreated = 0 Then '·„ Ì „ «‰‘«¡ ·=«·ﬁÌœ
        ISAllowDeleteUpdateContract = True
        Exit Function
        Else
        ISAllowDeleteUpdateContract = False
        
        
        Exit Function
        End If
        End If
        
        ISAllowDeleteUpdateContract = True

End Function




Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  «À»«  «·„Œ«·›…  ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «À»«  «·„Œ«·›… " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «À»«  «·„Œ«·›…  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «À»«  «·„Œ«·›… " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰« «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ «À»«  «·„Œ«·›… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   «À»«  «·„Œ«·›… ", 1, 15204351, -2147483630
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


Function print_report(Optional NoteSerial As Integer)
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
            
 MySQL = MySQL & "           SELECT            dbo.TblConfirmViolation.VendorID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblViolationTypes.Name AS violationName,"
  MySQL = MySQL & "          dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.DurationID, dbo.TblDurations.Name AS DurName, dbo.TblConfirmViolation.MonthID,"
  MySQL = MySQL & "          dbo.TblConfirmViolation.CarID, dbo.TblConfirmViolation.MinistryContractID, dbo.TblConfirmViolation.ViolationType, dbo.TblConfirmViolation.MinistryContractValue,"
   MySQL = MySQL & "         dbo.TblConfirmViolation.Date, dbo.TblConfirmViolation.DateH, dbo.TblConfirmViolation.Value, dbo.TblConfirmViolation.AbsenceCount, dbo.TblConfirmViolation.ID,"
   MySQL = MySQL & "         dbo.TblDurations_Details.Name AS MonthName, dbo.TblVehicleAllocation_Details.Type, dbo.TblConfirmViolation.Remarks, dbo.TblVehicleAllocation_Details.BoardNo"
  MySQL = MySQL & "          FROM     dbo.TblVehicleAllocation_Details INNER JOIN"
  MySQL = MySQL & "          dbo.TblAttributionContract ON dbo.TblVehicleAllocation_Details.IDVA = dbo.TblAttributionContract.IDAC RIGHT OUTER JOIN"
  MySQL = MySQL & "          dbo.TblConfirmViolation ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblConfirmViolation.CarID AND"
   MySQL = MySQL & "         dbo.TblAttributionContract.IDAC = dbo.TblConfirmViolation.MinistryContractID LEFT OUTER JOIN"
   MySQL = MySQL & "         dbo.TblDurations_Details ON dbo.TblConfirmViolation.MonthID = dbo.TblDurations_Details.ID LEFT OUTER JOIN"
   MySQL = MySQL & "         dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID LEFT OUTER JOIN"
   MySQL = MySQL & "         dbo.TblCustemers ON dbo.TblConfirmViolation.VendorID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
   MySQL = MySQL & "         dbo.TblDurations ON dbo.TblConfirmViolation.DurationID = dbo.TblDurations.ID"


  MySQL = MySQL & "   where  TblConfirmViolation.id  = " & val(txtID.Text)
     MySQL = MySQL & "  order by TblConfirmViolation.id "
     
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_ViolationReceipt.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_ViolationReceipt.rpt"
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
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
    End If
    
    
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

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «À»«  „Œ«·›… " _
       & CHR(13) & " „”·”·  " & txtID.Text _
       & CHR(13) & "«·⁄ﬁœ  " & dcContract.Text _
       & CHR(13) & "   «·”‰… «·œ—«”Ì…   " & dcDuration.Text _
       & CHR(13) & " «·› —…     " & dcMonth.Text _
       & CHR(13) & " ﬁÌ„… «·⁄ﬁœ     " & txtContractValue.Text _
       & CHR(13) & " „œ… «·⁄ﬁœ     " & Text2.Text _
       & CHR(13) & "  «·„ ⁄Âœ       " & dcVendor.Text _
       & CHR(13) & " «·„⁄œÂ/«·”Ì«—…   " & dcCar.Text _
       & CHR(13) & "  «—ÌŒ «·„Œ«·›…   " & dtpDate.value & "   " & dtpDateH.value _
       & CHR(13) & "  «·„Œ«·›…  " & dcViolation.Text _
       & CHR(13) & "  ‰Ê⁄ «·„Œ«·›… " & cbViolationType.Text _
       & CHR(13) & " ⁄œœ «Ì«„ «·€Ì«»  " & txtAbsenceCount.Text _
       & CHR(13) & " ﬁÌ„… «·ÌÊ„  " & txtDayRate.Text _
       & CHR(13) & "  «·⁄œœ  " & txtcount.Text _
       & CHR(13) & "ﬁÌ„… «·„Œ«·›…  " & txtValue.Text _
       & CHR(13) & "   „·«ÕŸ«   " & txtRemarks.Text & " "
       

    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function


