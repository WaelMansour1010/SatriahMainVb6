VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmPO8 
   Caption         =   "ÿ·» «·‘—«¡"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18705
   HelpContextID   =   340
   Icon            =   "FrmPO8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10605
   ScaleWidth      =   18705
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   10605
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   18705
      _cx             =   32994
      _cy             =   18706
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   1
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   9960
         Width           =   18660
         _cx             =   32914
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
            Height          =   390
            Index           =   0
            Left            =   17160
            TabIndex        =   11
            Top             =   120
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   688
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
            Height          =   390
            Index           =   1
            Left            =   15360
            TabIndex        =   12
            Top             =   120
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   688
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
            Height          =   390
            Index           =   2
            Left            =   13560
            TabIndex        =   13
            Top             =   120
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
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
            Height          =   390
            Index           =   3
            Left            =   11880
            TabIndex        =   14
            Top             =   120
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   688
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
            Height          =   390
            Index           =   4
            Left            =   9960
            TabIndex        =   15
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
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
            Height          =   390
            Index           =   5
            Left            =   8160
            TabIndex        =   16
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   6
            Left            =   360
            TabIndex        =   17
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   688
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
            Height          =   390
            Index           =   7
            Left            =   6360
            TabIndex        =   18
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   390
            Left            =   1440
            TabIndex        =   19
            Top             =   120
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   688
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   390
            Left            =   2640
            TabIndex        =   98
            Top             =   120
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   688
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
            ButtonImage     =   "FrmPO8.frx":038A
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
            Height          =   390
            Index           =   8
            Left            =   -360
            TabIndex        =   149
            Top             =   90
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â ÿ·» ‘—«¡ "
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
            Height          =   390
            Index           =   10
            Left            =   8040
            TabIndex        =   167
            Top             =   120
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…2"
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
            Height          =   390
            Index           =   12
            Left            =   4680
            TabIndex        =   205
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… 2 "
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   510
         Index           =   3
         Left            =   15
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   9510
         Width           =   18495
         _cx             =   32623
         _cy             =   900
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   18390
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   -210
            Visible         =   0   'False
            Width           =   1200
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4725
            TabIndex        =   22
            Top             =   45
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   9480
            TabIndex        =   146
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   360
            Index           =   49
            Left            =   11310
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   75
            Width           =   735
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   30
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label LblDiscountsTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   12075
            TabIndex        =   143
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   12180
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   360
            Index           =   50
            Left            =   13650
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   75
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   375
            Index           =   63
            Left            =   18375
            TabIndex        =   82
            Top             =   135
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   17550
            TabIndex        =   81
            Top             =   0
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   480
            Left            =   14715
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Width           =   1950
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   360
            Index           =   25
            Left            =   16830
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   60
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   330
            Index           =   3
            Left            =   19830
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   75
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   0
            Left            =   2985
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   2
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   345
            Left            =   2295
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   105
            Width           =   705
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   135
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   390
            Index           =   1
            Left            =   7770
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   75
            Width           =   1110
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3420
         Index           =   0
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   18690
         _cx             =   32967
         _cy             =   6033
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
         Begin VB.ComboBox Dcbpayment2 
            Height          =   315
            Left            =   4110
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1800
            Width           =   1485
         End
         Begin VB.TextBox TxtDuring 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   810
            TabIndex        =   201
            Top             =   1800
            Width           =   1830
         End
         Begin VB.TextBox TxtPeriodReq 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6705
            TabIndex        =   198
            Top             =   1800
            Width           =   1275
         End
         Begin VB.ComboBox DcbReqStatus 
            Height          =   315
            ItemData        =   "FrmPO8.frx":6BEC
            Left            =   810
            List            =   "FrmPO8.frx":6BEE
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   1440
            Width           =   1830
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√„—"
            Height          =   255
            Index           =   1
            Left            =   11685
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   -120
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   915
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ·»"
            Height          =   255
            Index           =   0
            Left            =   12585
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   -120
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox TxtPO6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   165
            Width           =   1710
         End
         Begin VB.TextBox TxtPayment 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   990
            TabIndex        =   178
            Top             =   -315
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox TxtModeSupply 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   10665
            TabIndex        =   171
            Top             =   5355
            Visible         =   0   'False
            Width           =   6360
         End
         Begin VB.TextBox TxtModeRecept 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   360
            TabIndex        =   168
            Top             =   3555
            Visible         =   0   'False
            Width           =   7560
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   810
            TabIndex        =   162
            Top             =   990
            Width           =   7170
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   10650
            TabIndex        =   158
            Top             =   2460
            Width           =   2070
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   14910
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   2460
            Width           =   2115
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   255
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   -540
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.TextBox TxtPONo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   -315
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   6900
            TabIndex        =   141
            Top             =   540
            Width           =   1080
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   270
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   3675
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   555
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   4995
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15285
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   165
            Width           =   1740
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   1005
            Left            =   930
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   2325
            Width           =   7050
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   480
            Left            =   15885
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1440
            Width           =   1140
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15885
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   990
            Width           =   1140
         End
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   735
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3630
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  «·«⁄ „«œ"
            Height          =   840
            Left            =   -2115
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   -3855
            Visible         =   0   'False
            Width           =   5280
            Begin VB.TextBox TxtLcNo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4080
               TabIndex        =   66
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   70123521
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4560
               TabIndex        =   67
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   70123521
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   68
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   70123521
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4560
               TabIndex        =   69
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   70123521
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   70
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   70123521
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷"
               BackColor       =   12632256
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12632256
               ColorHighlight  =   16777215
               ColorHoverText  =   255
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledText=   16711680
               ColorToggledHoverText=   255
               ColorTextShadow =   -2147483637
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "„·«ÕŸ« "
               Height          =   375
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   255
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·«⁄ „«œ"
               Height          =   255
               Index           =   0
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3015
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3975
            Visible         =   0   'False
            Width           =   8445
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   55
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   70123521
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   56
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo11 
               Height          =   315
               Left            =   2640
               TabIndex        =   57
               Top             =   960
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   " «·«‰ Â«¡"
               Height          =   285
               Index           =   24
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   23
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   285
               Index           =   22
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·…"
               Height          =   285
               Index           =   21
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»‰ﬂ"
               Height          =   285
               Index           =   20
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "‰Ê⁄ «·«„—"
               Height          =   285
               Index           =   19
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2835
            Left            =   -2685
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   -6315
            Visible         =   0   'False
            Width           =   9270
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   600
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   3120
               TabIndex        =   42
               Top             =   960
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   3120
               TabIndex        =   43
               Top             =   1320
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Height          =   315
               Left            =   120
               TabIndex        =   44
               Top             =   960
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   3120
               TabIndex        =   45
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   3120
               TabIndex        =   86
               Top             =   600
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Height          =   315
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„‘—Ê⁄"
               Height          =   270
               Index           =   11
               Left            =   2130
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„—ﬂ“ «· ﬂ·›…"
               Height          =   285
               Index           =   10
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   17
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ’‰Ì›"
               Height          =   285
               Index           =   16
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   285
               Index           =   15
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
               Height          =   285
               Index           =   14
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»·œ"
               Height          =   285
               Index           =   13
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   960
               Width           =   1215
            End
         End
         Begin VB.ComboBox CboPriceType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5475
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   -330
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   15285
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -330
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   -525
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   540
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   -1890
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   30
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   -450
            Visible         =   0   'False
            Width           =   2655
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   10665
            TabIndex        =   2
            Top             =   990
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   10665
            TabIndex        =   3
            Top             =   1440
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   390
            Left            =   15285
            TabIndex        =   1
            Top             =   540
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   688
            _Version        =   393216
            Format          =   70123521
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   690
            Left            =   9630
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2340
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1217
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmPO8.frx":6BF0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   930
            Left            =   2100
            TabIndex        =   33
            Top             =   3915
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1640
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈œ—«Ã ⁄—÷ Ã«Â“"
            BackColor       =   12632256
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   4
            Left            =   21000
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   3150
            Width           =   5385
            _cx             =   9499
            _cy             =   1535
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
            Begin VB.CheckBox XPChkTAX 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   330
               Left            =   1860
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   210
               Width           =   1815
            End
            Begin VB.TextBox XPTxtTaxValue 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   150
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   240
               Index           =   4
               Left            =   990
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   285
               Width           =   720
            End
         End
         Begin MSDataListLib.DataCombo Dccurrency 
            Height          =   315
            Left            =   10665
            TabIndex        =   84
            Top             =   525
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   4110
            TabIndex        =   94
            Top             =   180
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   810
            TabIndex        =   142
            Top             =   540
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   690
            Left            =   0
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   1217
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmPO8.frx":6F8A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbDetpartment 
            Height          =   315
            Left            =   10665
            TabIndex        =   170
            Top             =   2040
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPayment 
            Height          =   315
            Left            =   810
            TabIndex        =   179
            Top             =   180
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbShiping 
            Height          =   315
            Left            =   4110
            TabIndex        =   180
            Top             =   1395
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker SippingDate 
            Height          =   315
            Left            =   14895
            TabIndex        =   182
            Top             =   2955
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70123521
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DeliverDate 
            Height          =   315
            Left            =   10665
            TabIndex        =   184
            Top             =   2955
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70123521
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   615
            Index           =   11
            Left            =   -240
            TabIndex        =   194
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   330
            Index           =   52
            Left            =   0
            TabIndex        =   203
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ·«·"
            Height          =   330
            Index           =   51
            Left            =   2820
            TabIndex        =   202
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   330
            Index           =   48
            Left            =   5565
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "› —… «·ÿ·»"
            Height          =   300
            Index           =   47
            Left            =   8250
            TabIndex        =   199
            Top             =   1800
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·ÿ·»"
            Height          =   330
            Index           =   46
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   285
            Index           =   45
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   915
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ÿ·» œ«Œ·Ì/⁄—÷ ”⁄— „‘ —Ì« "
            Height          =   285
            Index           =   44
            Left            =   12450
            RightToLeft     =   -1  'True
            TabIndex        =   187
            ToolTipText     =   "»‰«¡ ⁄·Ï ÿ·» œ«Œ·Ì/⁄—÷ ”⁄— „‘ —Ì« "
            Top             =   240
            Width           =   2805
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· Ê’Ì·"
            Height          =   270
            Index           =   43
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   2955
            Width           =   2085
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·‘Õ‰"
            Height          =   270
            Index           =   42
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   2955
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·‘Õ‰ Ê«· Ê—Ìœ"
            Height          =   540
            Index           =   41
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   1395
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ÿ—Ìﬁ… ««” ·«„ «·„Ê«œ"
            Height          =   375
            Index           =   38
            Left            =   8235
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   -240
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Index           =   28
            Left            =   7980
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   2685
            Width           =   2235
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   465
            Index           =   32
            Left            =   8100
            TabIndex        =   174
            Top             =   465
            Width           =   2355
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ «·‰ﬁœÌ"
            Height          =   240
            Index           =   36
            Left            =   7965
            TabIndex        =   173
            Top             =   990
            Width           =   2490
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «· Ê—Ìœ"
            Height          =   450
            Index           =   39
            Left            =   20400
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   2160
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «·ÿ«·»…"
            Height          =   375
            Index           =   37
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   390
            Index           =   35
            Left            =   13140
            TabIndex        =   161
            Top             =   2460
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   285
            Index           =   34
            Left            =   17160
            TabIndex        =   160
            Top             =   2520
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
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
            Index           =   55
            Left            =   10050
            TabIndex        =   159
            Top             =   2505
            Width           =   375
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   330
            Index           =   33
            Left            =   13605
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   -315
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·ÿ·»Ì…"
            Height          =   330
            Index           =   18
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   3675
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8370
            TabIndex        =   95
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·Â"
            Height          =   465
            Index           =   12
            Left            =   13260
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   525
            Width           =   1665
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   330
            Index           =   9
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   165
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   375
            Index           =   5
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   165
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ÿ·»"
            Height          =   270
            Index           =   6
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   405
            Index           =   7
            Left            =   17190
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   540
            Index           =   8
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1440
            Width           =   1440
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   6390
         Left            =   0
         TabIndex        =   101
         Top             =   3045
         Width           =   18705
         _cx             =   32994
         _cy             =   11271
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
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|Õ«·Â «·«⁄ „«œ|⁄—Ê÷ «·«”⁄«—|«·ÿ·»«  «·œ«Œ·Ì…"
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
         Picture(0)      =   "FrmPO8.frx":7324
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   5925
            Left            =   19350
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   45
            Width           =   18615
            _cx             =   32835
            _cy             =   10451
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   4230
               Left            =   2520
               TabIndex        =   134
               Tag             =   "1"
               Top             =   960
               Width           =   13230
               _cx             =   23336
               _cy             =   7461
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
               FormatString    =   $"FrmPO8.frx":76BE
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
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   4560
               Width           =   3375
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5925
            Index           =   15
            Left            =   45
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   45
            Width           =   18615
            _cx             =   32835
            _cy             =   10451
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
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
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   1
            ChildSpacing    =   1
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
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmPO8.frx":7801
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5895
               Index           =   16
               Left            =   15
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   15
               Width           =   18585
               _cx             =   32782
               _cy             =   10398
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
               Appearance      =   5
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
                  Height          =   8820
                  Index           =   5
                  Left            =   0
                  TabIndex        =   112
                  TabStop         =   0   'False
                  Top             =   -465
                  Width           =   18645
                  _cx             =   32888
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
                  Begin VB.Frame Frame4 
                     BorderStyle     =   0  'None
                     Height          =   945
                     Left            =   600
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   5565
                     Visible         =   0   'False
                     Width           =   2235
                     Begin DBPIXLib.DBPix20 DBPix202 
                        Height          =   855
                        Left            =   240
                        TabIndex        =   114
                        Top             =   -120
                        Width           =   2415
                        _Version        =   131072
                        _ExtentX        =   4260
                        _ExtentY        =   1508
                        _StockProps     =   1
                        _Image          =   "FrmPO8.frx":7837
                        ImageResampleWidth=   100
                        ImageResampleHeight=   100
                        ImageResampleMode=   1
                        ImageSaveFormat =   0
                        JPEGQuality     =   75
                        JPEGEncoding    =   0
                        JPEGColorMode   =   0
                        JPEGNoRecompress=   -1  'True
                        JPEGRotateWarning=   0
                        PNGColorDepth   =   0
                        PNGCompression  =   0
                        PNGFilter       =   0
                        PNGInterlace    =   1
                        ImageDitherMethod=   3
                        ImagePaletteMethod=   4
                        ImagePreviewMode=   0   'False
                        ImageKeepMetaData=   -1  'True
                        UseAmbientBackcolor=   -1  'True
                        ViewAsyncDecoding=   -1  'True
                        ViewEnableMouseZoom=   -1  'True
                        ViewInitialZoom =   1
                        ViewHAlign      =   1
                        ViewVAlign      =   1
                        ViewMenuMode    =   0
                     End
                     Begin VB.Label LblPostedPerson 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "."
                        Height          =   255
                        Left            =   3600
                        RightToLeft     =   -1  'True
                        TabIndex        =   117
                        Top             =   240
                        Width           =   1695
                     End
                     Begin VB.Label Label10 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·‰ÊﬁÌ⁄"
                        Height          =   255
                        Left            =   2640
                        RightToLeft     =   -1  'True
                        TabIndex        =   116
                        Top             =   240
                        Width           =   855
                     End
                     Begin VB.Label Label4 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ì⁄ „œ"
                        Height          =   255
                        Left            =   5160
                        RightToLeft     =   -1  'True
                        TabIndex        =   115
                        Top             =   240
                        Width           =   735
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   825
                     Index           =   2
                     Left            =   0
                     TabIndex        =   118
                     TabStop         =   0   'False
                     Top             =   1425
                     Width           =   18285
                     _cx             =   32253
                     _cy             =   1455
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
                     Begin VB.TextBox TxtSerial 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00FFFFFF&
                        Enabled         =   0   'False
                        Height          =   345
                        Left            =   -7095
                        MaxLength       =   20
                        RightToLeft     =   -1  'True
                        TabIndex        =   188
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   4320
                     End
                     Begin VB.ComboBox CboItemCase 
                        Height          =   315
                        Left            =   5610
                        RightToLeft     =   -1  'True
                        Style           =   2  'Dropdown List
                        TabIndex        =   121
                        Top             =   300
                        Width           =   2100
                     End
                     Begin VB.TextBox TxtQuantity 
                        Alignment       =   1  'Right Justify
                        Height          =   360
                        Left            =   3045
                        MaxLength       =   10
                        RightToLeft     =   -1  'True
                        TabIndex        =   120
                        Top             =   300
                        Width           =   2505
                     End
                     Begin VB.TextBox TxtPrice 
                        Alignment       =   1  'Right Justify
                        Height          =   360
                        Left            =   720
                        MaxLength       =   10
                        RightToLeft     =   -1  'True
                        TabIndex        =   119
                        Top             =   300
                        Width           =   2325
                     End
                     Begin MSDataListLib.DataCombo DCboItemsName 
                        Height          =   315
                        Left            =   7710
                        TabIndex        =   122
                        Top             =   300
                        Width           =   6870
                        _ExtentX        =   12118
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCboItemsCode 
                        Height          =   315
                        Left            =   14640
                        TabIndex        =   123
                        Top             =   300
                        Width           =   3465
                        _ExtentX        =   6112
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin ImpulseButton.ISButton CmdAdd 
                        Height          =   525
                        Left            =   60
                        TabIndex        =   124
                        Top             =   255
                        Width           =   600
                        _ExtentX        =   1058
                        _ExtentY        =   926
                        ButtonStyle     =   1
                        ButtonPositionImage=   4
                        Caption         =   ""
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
                        BackStyle       =   0
                        ButtonImage     =   "FrmPO8.frx":784F
                        ColorButton     =   14871017
                        ColorHighlight  =   16777215
                        ColorHoverText  =   16711680
                        ColorShadow     =   -2147483637
                        ColorOutline    =   0
                        DrawFocusRectangle=   0   'False
                        ColorToggledHoverText=   16711680
                        LowerToggledContent=   0   'False
                        ColorTextShadow =   -2147483637
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "ﬂÊœ «·’‰›"
                        Height          =   255
                        Index           =   31
                        Left            =   14775
                        RightToLeft     =   -1  'True
                        TabIndex        =   129
                        Top             =   0
                        Width           =   3390
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈”„ «·’‰›"
                        Height          =   255
                        Index           =   30
                        Left            =   10080
                        RightToLeft     =   -1  'True
                        TabIndex        =   128
                        Top             =   0
                        Width           =   3210
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Õ«·… «·’‰›"
                        Height          =   255
                        Index           =   29
                        Left            =   5910
                        RightToLeft     =   -1  'True
                        TabIndex        =   127
                        Top             =   0
                        Width           =   1800
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·ﬂ„Ì…"
                        Height          =   255
                        Index           =   27
                        Left            =   3300
                        RightToLeft     =   -1  'True
                        TabIndex        =   126
                        Top             =   0
                        Width           =   2250
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”⁄—"
                        Height          =   255
                        Index           =   26
                        Left            =   795
                        RightToLeft     =   -1  'True
                        TabIndex        =   125
                        Top             =   0
                        Width           =   2250
                     End
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   3510
                     Left            =   120
                     TabIndex        =   130
                     Top             =   2250
                     Width           =   18345
                     _cx             =   32359
                     _cy             =   6191
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
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   20
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmPO8.frx":7BE9
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
                     WallPaperAlignment=   0
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   24
                  End
                  Begin MSComctlLib.Toolbar TBar 
                     Height          =   630
                     Left            =   600
                     TabIndex        =   131
                     Top             =   5955
                     Width           =   4365
                     _ExtentX        =   7699
                     _ExtentY        =   1111
                     ButtonWidth     =   609
                     ButtonHeight    =   1005
                     Appearance      =   1
                     _Version        =   393216
                  End
                  Begin ImpulseButton.ISButton Accredit 
                     Height          =   510
                     Left            =   5025
                     TabIndex        =   151
                     Top             =   5895
                     Width           =   2460
                     _ExtentX        =   4339
                     _ExtentY        =   900
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "«—”«· ··«⁄ „«œ"
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     Height          =   375
                     Index           =   40
                     Left            =   8250
                     RightToLeft     =   -1  'True
                     TabIndex        =   177
                     Top             =   6495
                     Width           =   1635
                  End
                  Begin VB.Label LblItemsCount 
                     Alignment       =   2  'Center
                     BackColor       =   &H00404040&
                     ForeColor       =   &H0000FFFF&
                     Height          =   285
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   132
                     Top             =   5835
                     Width           =   600
                  End
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Label12"
                  Height          =   960
                  Left            =   3855
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   240
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3660
                  Index           =   62
                  Left            =   3705
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   1515
                  Width           =   585
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5895
               Index           =   9
               Left            =   15
               TabIndex        =   105
               TabStop         =   0   'False
               Top             =   15
               Width           =   18585
               _cx             =   32782
               _cy             =   10398
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
               Appearance      =   5
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
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
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   3180
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   1515
                  Width           =   1245
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   4785
                  Left            =   4935
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   930
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3720
                  Index           =   69
                  Left            =   4305
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   1515
                  Width           =   630
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   3165
                  Index           =   68
                  Left            =   5775
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1800
                  Width           =   345
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3180
                  Index           =   67
                  Left            =   3705
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   1515
                  Width           =   600
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5925
            Left            =   19650
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   45
            Width           =   18615
            _cx             =   32835
            _cy             =   10451
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
            Begin VSFlex8UCtl.VSFlexGrid Grid3 
               Height          =   3390
               Left            =   2040
               TabIndex        =   155
               Tag             =   "1"
               Top             =   1320
               Width           =   13230
               _cx             =   23336
               _cy             =   5980
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO8.frx":7F12
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   5925
            Left            =   19950
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   45
            Width           =   18615
            _cx             =   32835
            _cy             =   10451
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   3390
               Left            =   2280
               TabIndex        =   164
               Tag             =   "1"
               Top             =   1080
               Width           =   13230
               _cx             =   23336
               _cy             =   5980
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
               FormatString    =   $"FrmPO8.frx":802F
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   0
         Width           =   18645
         _cx             =   32888
         _cy             =   1085
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
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
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
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   6225
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   240
            Visible         =   0   'False
            Width           =   405
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1995
            TabIndex        =   136
            Top             =   105
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO8.frx":816F
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
            Height          =   345
            Index           =   3
            Left            =   1080
            TabIndex        =   137
            Top             =   105
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO8.frx":8509
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
            Height          =   345
            Index           =   1
            Left            =   2835
            TabIndex        =   138
            Top             =   105
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO8.frx":88A3
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
            Height          =   345
            Index           =   2
            Left            =   180
            TabIndex        =   139
            Top             =   105
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO8.frx":8C3D
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
         Begin VB.Image Image1 
            Height          =   495
            Left            =   17535
            Picture         =   "FrmPO8.frx":8FD7
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ·» «·‘—«¡"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   13005
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   120
            Visible         =   0   'False
            Width           =   3930
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   8505
            Picture         =   "FrmPO8.frx":E38C
            Stretch         =   -1  'True
            Top             =   120
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   64
            Left            =   3885
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   360
            Width           =   8580
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   9
      Left            =   0
      TabIndex        =   166
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   688
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
End
Attribute VB_Name = "FrmPO8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)   As clsDCboSearch
 Dim CurrentTransactionType As Integer

Function CheckAcconts() As Boolean
CheckAcconts = False
Dim Account_Code_dynamic101 As String
Dim Account_Code_dynamic102 As String

            Account_Code_dynamic101 = get_account_code_branch(101, my_branch)
            Account_Code_dynamic102 = get_account_code_branch(102, my_branch)
             If Account_Code_dynamic101 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·„œÌ‰ ·«Ê«„— «·‘—«¡  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
              
                  If Account_Code_dynamic102 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·œ«∆‰ ·«Ê«„— «·‘—«¡  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
  
   CheckAcconts = True
   Exit Function
ErrTrap:
      CheckAcconts = False
End Function

Private Sub Cmd_Click(Index As Integer)
    Dim intDef As Integer
 '   On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            Dccurrency.BoundText = 1
        '    FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Me.CboPriceType.ListIndex = 0
                   CboPaymentType.ListIndex = 0
                   
Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.dcBranch.Enabled = True
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If

                   
                   
            Me.dcBranch.BoundText = Current_branch
            DBPix202.ImageClear
Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
FillOrderGrid
FillOrderGrid2
Opt(1).value = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If


              If ChekClodePeriod(Me.XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "› —Â „€·ﬁ… "
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If


            If ScreenAproved(val(Me.XPTxtBillID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If


            TxtModFlg.Text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id

        Case 2
            Dim Msg  As String

If SystemOptions.POMustentryAndBillMustEntry = True And (TxtPO6.Text = "") Then
   MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ·⁄œ„ «Œ Ì«— »‰«¡ ⁄·Ì Ê ÕœÌœ «·—ﬁ„", vbCritical
   Exit Sub
End If

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
If SystemOptions.PoCreateVoucher = True Then
  'If CheckAcconts = False Then Exit Sub
End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If


              If ChekClodePeriod(Me.XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "› —Â „€·ﬁ… "
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If


            If ScreenAproved(val(Me.XPTxtBillID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If


            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = GridTransType.purchaserequest
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   «·‘—«¡ "
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport
         Case 10
         print_report
         
         Case 11
           ShowGL_cc TxtNoteSerial.Text, , 200
           
        
        Case 8
            On Error GoTo ErrTrap

            If XPTxtBillID.Text <> "" Then
                Set SaleReport = New ClsSaleReport
                SaleReport.ShowPrice XPTxtBillID.Text, 6, DcboEmp.Text, val(DBCboClientName.BoundText)
            End If

            '        PrintReport1 (Txt_order_no.text)
        Case 6
            Unload Me
            
         Case 12

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report2
    End Select

    Exit Sub
ErrTrap:
End Sub

Function PrintReport1(order_no As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From QRY_items_orders_data where order_no='" & order_no & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status"
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.TITLE
    xReport.ReportAuthor = App.TITLE
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.CboPriceType.ListIndex = 0 Then
        Set Frm = New frmsalebill
    ElseIf Me.CboPriceType.ListIndex = 1 Then
        Set Frm = New FrmBillBuy
    End If

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Dccurrency.BoundText = Me.Dccurrency.BoundText

        For RowNum = 1 To FG.Rows - 1

            If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.Rows = .FG.Rows + 1
            End If

            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
        
            StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))
        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdTemplate_Click()
    Dim Frm  As FrmBuySearch
    On Error GoTo ErrTrap
    Set Frm = New FrmBuySearch

    With Frm
        .DealingForm = InsertTemplate
        .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
        '    .MDIChild = True
        .BorderStyle = 0
        '  .MinButton = True
        .show vbModeless, mdifrmmain
        .Visible = True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Change()
    TxtSearchCode.Text = ""

    Dim DefaultSalesPersonId As Integer
    Dim Fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode

    TxtSearchCode.Text = Fullcode

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
 
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

        If Not DefaultSalesPersonId = 0 Then

            Me.DcboEmp.BoundText = DefaultSalesPersonId
        End If
        FillOrderGrid
        FillOrderGrid2
    End If
 
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
      FrmCompanySearch.lblSearchtype.Caption = 8
       FrmCompanySearch.show vbModal
        
        
    End If
          
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos

       
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
      
    End If

End Sub
 
Private Sub DcboEmp_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetSalesRepData Me.DcboEmp

    End If

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.Text = DCboItemsCode.Text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 22
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.Text = DCboItemsCode.Text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 22
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
   If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

                 If CheckStoreCoding(val(dcBranch.BoundText), 29) = True Then
                ' TxtNoteSerial.text = ""
                TxtNoteSerial1.Text = ""
            
                 End If
     
    End If
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
'DCboStoreName_Change
End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName

    End If
        
End Sub

Private Sub Dcbranch_Change()
  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtNoteSerial1.Text = ""
TxtNoteSerial.Text = ""
End If



End Sub

Private Sub Dcbranch_Click(Area As Integer)
'TxtNoteSerial.text = ""
'TxtNoteSerial1.text = ""
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = " select id,code from currency"
 
        fill_combo Me.Dccurrency, My_SQL

    End If

End Sub

Private Sub Ele_Click(Index As Integer)

    Select Case Index

        Case 6
            On Error GoTo ErrTrap
            '        If Me.WindowState = vbNormal Then
            '            Me.WindowState = vbMaximized
            '        Else
            '            Me.WindowState = vbNormal
            '        End If
    End Select

    Exit Sub
ErrTrap:
End Sub


Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngRow  As Double
Dim StrAccountCode As String
' With FG
'        Select Case .ColKey(Col)
'            Case "countris"
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("countrisid"), False, True)
'                .TextMatrix(Row, .ColIndex("countrisid")) = StrAccountCode
'         End Select
' End With
 
    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , , Me.Txt_order_no
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), , Me.Txt_order_no

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
ReLineGrid
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
     '  FrmAddNewItem.DealingForm = ShowPrice
     '   FrmAddNewItem.show vbModal
    End If

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String

'With FG
'  Select Case .ColKey(Col)
'
'            Case "countris"
'
'                StrSQL = "select * from TblCountriesData"
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                    StrComboList = FG.BuildComboList(rs, "CountryName", "CountryID")
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'                 .ComboList = StrComboList
'            End Select
'
'End With
ReLineGrid
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
   ''''///
   Dim SUM As Double
   IntCounter = 0
   SUM = 0
     With FG

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                            End If

        Next i

    End With
    End Sub
Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

 

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.Text)
    'Frame3.Visible = True
End Sub

Private Sub ISButton2_Click()
    On Error Resume Next
ShowAttachments TxtNoteSerial1, "060520152"
 
End Sub

Private Sub Label10_Click()
    Frame3.Visible = False
End Sub
 
Private Sub Accredit_Click()
    Dim sql As String
    Dim BeginTrans As Boolean
    'sql = "update  Transactions  set Posted=" & user_id & "  where Transaction_ID=" & Val(XPTxtBillID.text)
    'Cn.Execute sql

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If
   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
    Retrive (val(XPTxtBillID.Text))

End Sub
Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(XPTxtBillID.Text)
                  RSApproval("NoteSerial").value = TxtNoteSerial1.Text
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function
Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

   
        StrSQL = "Select * from transactions where  Transaction_Type=43 and Order_no='" & order_no & "'"
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))

            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            If Transaction_Type = 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            End If
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Sub RetriveoOrderPO6(Optional TransID As Integer = 0, Optional Notserial As String = "")
Dim StrSQL As String
Dim RsDetails As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
Dim Num As Integer
If TransID <> 0 Then
StrSQL = "SELECT * FROM Transactions WHERE Transaction_ID=" & TransID
Else
StrSQL = "SELECT * FROM Transactions WHERE NoteSerial1='" & Notserial & " '"

End If
 StrSQL = StrSQL + " and  ( Transaction_Type =38  or Transaction_Type =46) "
    Set Rs1 = New ADODB.Recordset
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    TxtBillComment.Text = IIf(IsNull(Rs1("TransactionComment")), "", (Rs1("TransactionComment").value))
    DCboStoreName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
     DBCboClientName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
    DBCboClientName.BoundText = IIf(IsNull(Rs1("CusID")), 0, (Rs1("CusID").value))
'     DBCboClientName.BoundText = IIf(IsNull(Rs1("CusID")), 0, (Rs1("CusID").value))
     DcbDetpartment.BoundText = IIf(IsNull(Rs1("DepartementID")), 0, (Rs1("DepartementID").value))
    
   FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = " SELECT   dbo.Transaction_Details.unitid as itmemunitid,   dbo.TblItems.HaveSerial AS Expr1, *"
    StrSQL = StrSQL + "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + "                   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL + "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblProcessDEF ON dbo.Transaction_Details.Oper_ID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id"

    StrSQL = StrSQL + " where Transaction_ID=" & Rs1("Transaction_ID").value
    StrSQL = StrSQL + " order by Transaction_Details.id "
    
'order by id
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        ''//
       '  Fg.TextMatrix(Num, Fg.ColIndex("projectid")) = IIf(IsNull(RsDetails("project_ID1")), "", (RsDetails("project_ID1").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("pandid")) = IIf(IsNull(RsDetails("Pand_ID")), "", (RsDetails("Pand_ID").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("operaid")) = IIf(IsNull(RsDetails("Oper_ID")), "", (RsDetails("Oper_ID").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("pand")) = IIf(IsNull(RsDetails("des")), "", (RsDetails("des").value))
       '  If SystemOptions.UserInterface = ArabicInterface Then
       '  Fg.TextMatrix(Num, Fg.ColIndex("project")) = IIf(IsNull(RsDetails("Project_name")), "", (RsDetails("Project_name").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessName")), "", (RsDetails("ProcessName").value))
       '  Else
       '  Fg.TextMatrix(Num, Fg.ColIndex("project")) = IIf(IsNull(RsDetails("Project_nameE")), "", (RsDetails("Project_nameE").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessNameE")), "", (RsDetails("ProcessNameE").value))
       '  End If
        ''//
 '   Dim Item_ID As Double
  Dim movingqty As Double
   Dim actulaqty As Double
 'Item_ID = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
 ' actulaqty
  
 '       CheckintenalRequstQty Item_ID, TxtPO6
        
        
 
    Dim RsTest As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngColorID As Long
    Dim StrItemSize As String
    Dim LngClassId As Long
    
         LngItemID = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
         
              LngColorID = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            StrItemSize = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
          LngClassId = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
        
              Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.DCboStoreName.BoundText)), XPDtbBill.value, , , , , True, LngColorID, StrItemSize, LngClassId)
              If RsTest.EOF Or RsTest.BOF Then
                   actulaqty = 0
              Else
           
                  actulaqty = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
              
              End If
              
              If SystemOptions.poWithatotalQty = False Then ' Õ«·Â «·ﬂ„Ì… ﬂ«„·…
                         movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - actulaqty
              
              Else
                         movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
              End If
              
             If movingqty = 0 Then GoTo skiploop

              
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
           ' Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            
  '       If SystemOptions.poWithatotalQty = False Then
  '           FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - IIf(IsNull(RsDetails("ItemBalance")), 0, (RsDetails("ItemBalance").value))
  '        Else
  '        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
  '        End If
  '
                FG.TextMatrix(Num, FG.ColIndex("Count")) = movingqty
            
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), 0, (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
             FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("itmemunitid")), "", (RsDetails("itmemunitid").value))
             FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        

 ' FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID"))))
        



        FG.TextMatrix(Num, FG.ColIndex("RequestLimit")) = IIf(IsNull(RsDetails("RequestLimit")), 0, (RsDetails("RequestLimit").value))
       FG.TextMatrix(Num, FG.ColIndex("LastPurchaseDate")) = IIf(IsNull(RsDetails("LastPurchaseDate")), "", (RsDetails("LastPurchaseDate").value))
        FG.TextMatrix(Num, FG.ColIndex("LastPurchasePrice")) = IIf(IsNull(RsDetails("LastPurchasePrice")), 0, (RsDetails("LastPurchasePrice").value))
        FG.TextMatrix(Num, FG.ColIndex("LastPurchaseqty")) = IIf(IsNull(RsDetails("LastPurchaseqty")), 0, (RsDetails("LastPurchaseqty").value))
       FG.TextMatrix(Num, FG.ColIndex("AverageIssue")) = IIf(IsNull(RsDetails("AverageIssue")), 0, (RsDetails("AverageIssue").value))
skiploop:
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num
End If
    End If
End Sub

Private Sub TxtPO6_Change()
If Me.TxtModFlg.Text <> "R" And Me.TxtModFlg.Text <> "" Then
RetriveoOrderPO6 , Me.TxtPO6.Text
End If
End Sub

Private Sub TxtPO6_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
If KeyCode = vbKeyF3 Then
FrmBuySearch.DealingForm = GridTransType.internalorder
  FrmBuySearch.Index = 6
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…/ ⁄—÷ ”⁄— „‘ —Ì« "
            FrmBuySearch.show vbModal
       End If
       End If
End Sub

Private Sub TxtPONo_Change()
  If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TxtPONo
    End If
End Sub

Private Sub TxtPONo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Order_no_search4.show
        Order_no_search4.RetrunType = 43

        If val(Me.DBCboClientName.BoundText) <> 2 Then
        
            Order_no_search4.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
        End If
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtLcNo_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search3.show
        Order_no_search3.RetrunType = 1
         
    End If
        
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub

Private Sub VSFlexGrid1_Click()
    With FG
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
       
    End With
 
    fillOrders (1)
    
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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
        '    If Cmd(3).Enabled = False Then Exit Sub
        '    Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
       
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_Change()
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset
    Dim ShowTax As Boolean
    Dim Dcombos As ClsDataCombos
   If SystemOptions.POMustentryAndBillMustEntry = True Then
   TxtPO6.locked = True
End If


    On Error GoTo ErrTrap
'  If SystemOptions.UserInterface = ArabicInterface Then
'                FG.ColComboList(FG.ColIndex("Shipping")) = "#1;  «·„‘ —Ì|#2; «·»«∆⁄"
'            ElseIf SystemOptions.UserInterface = EnglishInterface Then
'               FG.ColComboList(FG.ColIndex("Shipping")) = "#1;Buyer |#2;Seller "
'            End If
            
   ' If GeneralPriceType = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   With Dcbpayment2
   .Clear
   .AddItem "‰ﬁœ«"
   .AddItem "«Ã·"
   End With
  Else
     With Dcbpayment2
   .Clear
   .AddItem "Cash"
   .AddItem "Debt"
   End With
  End If
        ScreenNameArabic = "  ÿ·»  «·‘—«¡ "
        ScreenNameEnglish = "Purchase   Request "
        CurrentTransactionType = 47
  
   ' End If

    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

'    Me.Caption = ScreenNameArabic
'   Ele(6).Caption = ScreenNameArabic

If SystemOptions.UserInterface = ArabicInterface Then
DcbReqStatus.AddItem "⁄«œÌ"
DcbReqStatus.AddItem "ÿ«—Ï"
    Me.Caption = ScreenNameArabic
 Else
 DcbReqStatus.AddItem "Normal"
DcbReqStatus.AddItem "Pressing"
 Me.Caption = ScreenNameEnglish
 End If
   Ele(6).Caption = Me.Caption
   
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang

    End If

    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set CmdConvert.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    Set CmdTemplate.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
     Set Accredit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Required").Picture
    Set NewGrid.Grid = FG
  NewGrid.GridTrans = GridTransType.purchaserequest
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.TxtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    'Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
Set NewGrid.Customer = Me.DBCboClientName
Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    
    
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DcboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.StoreName = DCboStoreName
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
        Set NewGrid.DtpBillDate = Me.XPDtbBill
        
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
      Resize_Form Me, TransactionSize
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FG.WallPaper = BGround.Picture
    AddTip
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos

   
        Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True '  2 supplier  1 customer
 
     Dcombos.GetPaymentMathods Me.DcbPayment
     Dcombos.GetShipingMathods Me.DcbShiping
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetEmpDepartments Me.DcbDetpartment
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName

    Dcombos.GetSalesRepDatapurchase Me.DcboEmp
 
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID

    NewGrid.FillGrid

    With CboPaymentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With
        
        
           With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        End With
        
    With Me.CboPriceType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
             
            .AddItem " ÿ·»«  «Ê«„— «·»Ì⁄ "
       
        Else
             
            .AddItem " Sales Order "
 
        End If

        .ListIndex = 0
    End With

    With Me.CboType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "   ÌœÊÌ "
            .AddItem "«·Ì ÿ»ﬁ« ·Õœ «·ÿ·» "
     
        Else
            .AddItem "Manual"
            .AddItem "Auto "
     
        End If

        .ListIndex = 0
    End With


        
        
   With XPCboDiscountType
           If SystemOptions.UserInterface = ArabicInterface Then
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
            Else
            .Clear
           .AddItem ("With out")
           .AddItem ("No Discount")
           .AddItem ("Discount With Value")
           .AddItem ("Discount With Percentage")
             End If
  End With
    'StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=6 or Transaction_Type=29  or Transaction_Type=17)" 'OR Transaction_Type=17
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=" & CurrentTransactionType

    StrSQL = StrSQL + " Order By Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim My_SQL As String
    My_SQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, My_SQL
    fill_combo Me.DataCombo11, My_SQL

    My_SQL = " select code,account_name from markaas_taklefa"
 
    fill_combo Me.DataCombo1, My_SQL

    My_SQL = " select id,Project_name from projects"
 
    fill_combo Me.DataCombo2, My_SQL

    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL

    My_SQL = " select id,name from Shipment_mode"
 
    fill_combo Me.DataCombo5, My_SQL
    CboPriceType.ListIndex = GeneralPriceType

    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ   " & Txt_order_no.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & "«‰Ê⁄ «·”‰œ  " & CboPriceType.Text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.Text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & CHR(13) & " —ﬁ„ «·«⁄ „«œ    " & TxtLcNo
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr . No   " & Txt_order_no.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Type  " & CboPriceType.Text & CHR(13) & " Store  " & DCboStoreName.Text & CHR(13) & " Customer/ Supplier " & DBCboClientName.Text & CHR(13) & " Lc NO    " & TxtLcNo
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , , Me.Txt_order_no
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , , Me.Txt_order_no
    End If
    
End Function

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            ' Me.Caption = "⁄—÷ √”⁄«—"
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
            XPBtnNewClients.Enabled = False
        
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            FG.Editable = flexEDNone
      '      Accredit.Enabled = True
            CmdConvert.Enabled = True
            '   CmdConvert.Visible = True
            CmdTemplate.Visible = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                CmdConvert.Enabled = False
                Accredit.Enabled = False
            End If

            Ele(2).Enabled = False

        Case "N"
            ' Me.Caption = "⁄—÷ √”⁄«—( ÃœÌœ )"
            Accredit.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Accredit.Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.Rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
        
            CmdConvert.Visible = False
            CmdTemplate.Enabled = True
            '  CmdTemplate.Visible = True
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"
            ' Me.Caption = "⁄—÷ √”⁄«—(  ⁄œÌ· )"
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
        
            FG.Enabled = True
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
            XPBtnNewClients.Enabled = True
        
            Accredit.Enabled = False
            CmdConvert.Visible = False
            CmdTemplate.Visible = False
            Ele(2).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim Num As Long
    Dim Dusername As String
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    
   If SystemOptions.PoCreateVoucher = True Then
  '  Me.TXTNoteID.text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
  'Me.TxtNoteSerial.text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
  End If


    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    Txt_order_no.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    TxtPONo.Text = IIf(IsNull(rs("PONo").value), "", rs("PONo").value)
    TxtBillComment.Text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
      CboPaymentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
Me.TxtPayment.Text = IIf(IsNull(rs("PaymentT").value), "", (rs("PaymentT").value))
XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
  XPTxtDiscountVal.Text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
  
    If rs("shipped").value = True Then
        chkshipped.value = vbChecked
    Else
        chkshipped.value = Unchecked
    End If

    Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").value), "", rs("countryid").value)

  If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If
    
    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    Me.DcbReqStatus.ListIndex = IIf(IsNull(rs("ReqStatus").value), -1, rs("ReqStatus").value)
    'If rs("Transaction_Type").value = 6 Then
    '    Me.CboPriceType.ListIndex = 1
    'ElseIf rs("Transaction_Type").value = 17 Then '17
    '    Me.CboPriceType.ListIndex = 0
    'ElseIf rs("Transaction_Type").value = 29 Then
    'Me.CboPriceType.ListIndex = 2
    'End If

 
   Me.CboPriceType.ListIndex = 0
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
''//
 Me.DcbDetpartment.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
 Me.TxtModeRecept.Text = IIf(IsNull(rs("ModeReceptEq").value), "", (rs("ModeReceptEq").value))
 Me.TxtModeSupply.Text = IIf(IsNull(rs("ModeSupply").value), "", (rs("ModeSupply").value))
''//

    XPTxtTaxValue.Text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    TxtLcNo.Text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
''/// 11 05 2015
Me.DcbPayment.BoundText = IIf(IsNull(rs("PaymentID").value), "", rs("PaymentID").value)
Me.DcbShiping.BoundText = IIf(IsNull(rs("ShipingID").value), "", rs("ShipingID").value)
''//
    If Txt_order_no <> "" Then
        Me.TxtNoteSerial1.Text = Txt_order_no
    End If
''// 25 05 2015
SippingDate.value = IIf(IsNull(rs("SippingDate").value), Date, (rs("SippingDate").value))
DeliverDate.value = IIf(IsNull(rs("DeliverDate").value), Date, (rs("DeliverDate").value))
TxtPO6.Text = IIf(IsNull(rs("NotSeialPO6").value), "", (rs("NotSeialPO6").value))

Me.Dcbpayment2.ListIndex = IIf(IsNull(rs("PaymentID2").value), -1, (rs("PaymentID2").value))
TxtPeriodReq.Text = IIf(IsNull(rs("PeriodReq").value), "", (rs("PeriodReq").value))
TxtDuring.Text = IIf(IsNull(rs("During").value), "", (rs("During").value))



    'Txt_order_no
If IsNull(rs("requestOrOrder").value) Then

Opt(0).value = True
Else
        If rs("requestOrOrder").value = 0 Then
            Opt(0).value = True
        Else
            Opt(1).value = True
        End If

End If

 
        
    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

'    DBPix202.ImageClear

'    If Dir(App.path & "\images\sign\sign" & rs("posted").value & ".JPG") <> "" Then
'
'        DBPix202.ImageLoadFile (App.path & "\images\sign\sign" & user_id & ".JPG")
'    End If

   If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   
    'If Not IsNull(rs("posted").value) Then
    '    Frame4.Visible = True
    '    GetUserData val(rs("posted").value), , , , , , , Dusername
    '    LblPostedPerson = Dusername

    '                If user_id = rs("posted").value Then
    '                                If CheckOrderNotInTransaction(21, TxtNoteSerial1) = False Then
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                    Accredit.Caption = "«·€«¡ «·«⁄ „«œ "
    '                                                Else
    '                                                    Accredit.Caption = "Cancel Accredit   "
    '                                                End If
    '
    '                                Else
    '
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                    Accredit.Caption = "  «—”«· ··«⁄ „«œ "
    '                                                Else
    '                                                    Accredit.Caption = " send to accredit   "
     '                                               End If
    '
    '                                End If
    '
    '                End If

    'Else
    '    Frame4.Visible = False
    '    Accredit.Caption = "     «—”«· ··«⁄ „«œ "
    'End If
  
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
   ' StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
   ' StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
  StrSQL = " SELECT     dbo.TblItems.HaveSerial AS Expr1, *"
  StrSQL = StrSQL & " FROM         dbo.TblItems INNER JOIN"
  StrSQL = StrSQL & "                    dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 
  StrSQL = StrSQL & "                   dbo.TblCountriesData ON dbo.Transaction_Details.countrisid = dbo.TblCountriesData.CountryID"
  StrSQL = StrSQL & " Where (dbo.Transaction_Details.Transaction_ID =" & val(rs("Transaction_ID").value) & ")"
   

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        ''//
          FG.TextMatrix(Num, FG.ColIndex("LastPurchasePrice")) = IIf(IsNull(RsDetails("LastPurchasePrice").value), "", (RsDetails("LastPurchasePrice").value))
          FG.TextMatrix(Num, FG.ColIndex("New_Rolling")) = IIf(IsNull(RsDetails("New_Rolling").value), "", (RsDetails("New_Rolling").value))
          FG.TextMatrix(Num, FG.ColIndex("ItemBalance")) = IIf(IsNull(RsDetails("ItemBalance").value), "", (RsDetails("ItemBalance").value))
          FG.TextMatrix(Num, FG.ColIndex("ItemSalesYear")) = IIf(IsNull(RsDetails("ItemSalesYear").value), "", (RsDetails("ItemSalesYear").value))
          FG.TextMatrix(Num, FG.ColIndex("catalog")) = IIf(IsNull(RsDetails("catalog").value), "", (RsDetails("catalog").value))
         '   FG.TextMatrix(Num, FG.ColIndex("countris")) = IIf(IsNull(RsDetails("CountryName")), "", Trim(RsDetails("CountryName").value))
         '   FG.TextMatrix(Num, FG.ColIndex("Shipping")) = IIf(IsNull(RsDetails("Shipping")), "", (RsDetails("Shipping").value))
        
           FG.TextMatrix(Num, FG.ColIndex("countris")) = IIf(IsNull(RsDetails("countrisid")), "", Trim(RsDetails("countrisid").value))
            FG.TextMatrix(Num, FG.ColIndex("Shipping")) = IIf(IsNull(RsDetails("Shipping")), "", (RsDetails("Shipping").value))
 
 
        ''//
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
           
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        If SystemOptions.UserInterface = ArabicInterface Then
           FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
           Else
           FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemNamee")), "", Trim(RsDetails("ItemNamee").value))
           End If
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
           
           
               FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
 FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
 
 
 
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))

            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
fillapprovData
FillOrderGrid
FillOrderGrid2
ReLineGrid

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Sub GRID3_Click()

    With FG
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
       
    End With
 
    fillOrders (0)

End Sub


Function fillOrders(Optional gridno As Integer = 0)
    Dim i As Integer
If gridno = 0 Then
    With Grid3

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            DCboStoreName.BoundText = val(.TextMatrix(i, .ColIndex("StoreID")))
            End If

        Next i

    End With

Else


    With VSFlexGrid1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            DCboStoreName.BoundText = val(.TextMatrix(i, .ColIndex("StoreID")))
            End If

        Next i

    End With
    
End If
End Function



Function Retrive_orders_data(Transaction_ID As Integer)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.Rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.Rows - 1 'RsDetails.RecordCount
    
'            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
'            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
'            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        
        Dim movingqty As Double
        Dim actulaqty As Double
        Dim RsTest As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngColorID As Long
    Dim StrItemSize As String
    Dim LngClassId As Long
    
         LngItemID = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
         
              LngColorID = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            StrItemSize = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
          LngClassId = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
        
              Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.DCboStoreName.BoundText)), XPDtbBill.value, , , , , True, LngColorID, StrItemSize, LngClassId)
              If RsTest.EOF Or RsTest.BOF Then
                   actulaqty = 0
              Else
           
                  actulaqty = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
              
              End If
              
              If SystemOptions.poWithatotalQty = False Then ' Õ«·Â «·ﬂ„Ì… ﬂ«„·…
                         movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - actulaqty
              
              Else
                         movingqty = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
              End If
              
             If movingqty = 0 Then GoTo skiploop
         
         
         
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = movingqty
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showprice")), "", (RsDetails("showprice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
skiploop:
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function

Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
'    StrSQL = "SELECT     TOP 100 PERCENT  dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, "
'StrSQL = StrSQL + "  dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TblEmployee.Emp_Code,"
'StrSQL = StrSQL + "   dbo.TblEmployee.emp_name , dbo.TblEmployee.Emp_Namee, dbo.TbLLevels.name, dbo.TbLLevels.namee"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + "   dbo.TblEmployee ON dbo.ApprovalData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
'StrSQL = StrSQL + "   dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + "  ORDER BY dbo.ApprovalData.levelorder"
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

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
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function



Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap
 
    Me.LblTotal.Caption = XPTxtSum.Text
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–« «·”‰œ   .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.TITLE) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ·    Â–« «·”‰œ .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.TITLE) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If XPTxtBillID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                          
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
     

    With Grid3

        For i = 1 To .Rows - 1
     
 
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
        
       
            Cn.Execute sql
 
        Next
       
    End With
    
                
      With VSFlexGrid1

        For i = 1 To .Rows - 1
     
 
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
        
       
            Cn.Execute sql
 
        Next
       
    End With
    
                rs.delete
                rs.MoveFirst

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.TITLE
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄—÷ ”⁄— ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷  ﬁ—Ì— »«·»Ì«‰«  «·Õ«·Ì… " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄—÷ «·”⁄— «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄—÷ ”⁄—" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub


Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub


Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.Text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
      '  lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
      '  lbl(8).Visible = False
    Else
     '   lbl(8).Visible = True
        XPTxtDiscountVal.Visible = True
      '  lbl(8).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub


Private Function CheckCashCustomer() As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Trim$(Me.TxtCashCustomerName.Text) = "" Then
        CheckCashCustomer = True
    Else
        StrSQL = "Select * From Transactions Where CashCustomerName='" & Trim$(Me.TxtCashCustomerName.Text) & "'"
    
    End If

End Function


Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    'Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim BeginTrans As Boolean
  ' On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.Text <> "R" Then
        If DBCboClientName.Text = "" And Opt(1).value = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
            Else
                Msg = "Please Select Vendor"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If DCboStoreName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
                Msg = "Select Inventory"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Dccurrency.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ «·⁄„·…"
            Else
                Msg = "Select Currency"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
            Dccurrency.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    
    If CboPaymentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
        CboPaymentType.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
        If Me.CboPriceType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄    «·«„—  ( )...!!!"
            Else
                Msg = "Specify Order Type"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
            CboPriceType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If



 If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·«„— " & CHR(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
                Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            Else
                Msg = Msg + " Must Enter Discount Value " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
            XPCboDiscountType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If




        If XPChkTAX.value = Checked Then
            If XPTxtTaxValue.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
                Else
                    Msg = "Insert Sales Tax"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
                XPTxtTaxValue.SetFocus
                FG.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
 
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If
Dim sql As String
sql = "Select * from Transaction_Details where 1=-1"
        Set RSTransDetails = New ADODB.Recordset
        RSTransDetails.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CboPriceType.ListIndex = 0 Then
            Transaction_Type = CurrentTransactionType
            Sanad_No = CurrentTransactionType
  
 
         
        End If

        my_branch = val(dcBranch.BoundText)

        If TxtNoteSerial1.Text = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ    " & CHR(13) & " Enter Vchr No": Exit Sub
                Else
                    TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 170, , Transaction_Type, , val(DCboStoreName.BoundText))
                End If
            End If
        End If
 
        Txt_order_no = Me.TxtNoteSerial1.Text
 
        Cn.BeginTrans
        BeginTrans = True
    
        If Me.TxtModFlg.Text = "N" Then
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=6"))
            
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
            rs.AddNew
         Else
           If SystemOptions.PoCreateVoucher = True Then
             '   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
             '     Cn.Execute StrSQL, , adExecuteNoRecords

         End If
        End If

        Screen.MousePointer = vbArrowHourglass
       rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rs("branchID").value = val(Me.dcBranch.BoundText)
     rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))
       rs("Transaction_ID").value = val(XPTxtBillID.Text)
        rs("order_no").value = Txt_order_no.Text
    
        If chkshipped.value = vbChecked Then
            rs("shipped").value = 1
        Else
            rs("shipped").value = 0
        End If
    
    
       If Opt(0).value = True Then
            rs("requestOrOrder").value = 0
        Else
            rs("requestOrOrder").value = 1
        End If
        
       rs("PaymentID2").value = val(Me.Dcbpayment2.ListIndex)
       rs("PeriodReq").value = TxtPeriodReq.Text
       rs("During").value = val(TxtDuring.Text)
       
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Serial").value = TxtTransSerial.Text

        rs("PONO").value = IIf(TxtPONo.Text = "", Null, (TxtPONo.Text))
        rs("Transaction_Type").value = CurrentTransactionType

        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("countryid").value = IIf(DataCombo4.BoundText = "", Null, val(DataCombo4.BoundText))
        rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
        rs("ReqStatus").value = IIf(Me.DcbReqStatus.ListIndex = -1, Null, val(DcbReqStatus.ListIndex))
    
       If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs("CashCustomerName").value = Null
    End If
    
       If CboPaymentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPaymentType.ListIndex)
    End If
    rs("PaymentT").value = Trim$(Me.TxtPayment.Text)
    
       If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
  rs("Trans_Discount").value = IIf(XPTxtDiscountVal.Text = "", Null, val(XPTxtDiscountVal.Text))
 ''//
   rs("DeptID").value = IIf(DcbDetpartment.BoundText = "", Null, val(Me.DcbDetpartment.BoundText))
   rs("ModeReceptEq").value = IIf(Me.TxtModeRecept.Text = "", Null, TxtModeRecept.Text)
   rs("ModeSupply").value = IIf(Me.TxtModeSupply.Text = "", Null, TxtModeSupply.Text)

 ''//
 
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
        rs("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
        rs("total").value = IIf(XPTxtSum.Text = "", Null, val(XPTxtSum.Text))
        rs("LcNo").value = IIf(TxtLcNo.Text = "", Null, (TxtLcNo.Text))
    ''//11 05 2015
    rs("ShipingID").value = val(Me.DcbShiping.BoundText)
    rs("PaymentID").value = val(Me.DcbPayment.BoundText)
    ''//
   ''//25 05 2015
   rs("SippingDate").value = SippingDate.value
   rs("DeliverDate").value = DeliverDate.value
   rs("NotSeialPO6").value = Me.TxtPO6.Text
   
        rs.update
    
       CuurentLogdata
  
        If Me.TxtModFlg.Text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
                RSTransDetails("order_id").value = val(XPTxtBillID.Text)
             
                RSTransDetails("order_no").value = Txt_order_no.Text
                ''//
               ' RSTransDetails("countrisid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("countrisid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("countrisid"))))
               ' RSTransDetails("Shipping").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Shipping")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Shipping"))))
                RSTransDetails("countrisid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("countris")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("countris"))))
                RSTransDetails("Shipping").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Shipping")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Shipping"))))
                RSTransDetails("New_Rolling").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("New_Rolling")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("New_Rolling"))))
                RSTransDetails("LastPurchasePrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LastPurchasePrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("LastPurchasePrice"))))
                RSTransDetails("ItemSalesYear").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSalesYear")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSalesYear"))))
                RSTransDetails("ItemBalance").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemBalance")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemBalance"))))
  
                ''//
                RSTransDetails("catalog").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("catalog")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("catalog"))))
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
 
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    'RSTransDetails("Price").value = Val(IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))) / RSTransDetails("Quantity").value
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                End If
RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))

                RSTransDetails.update
            End If

        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    Closeorders (0)
   Closeorders (1)
   If SystemOptions.PoCreateVoucher = True Then
   '    createVoucher
   '   updateNotesValueAndNobytext (val(TXTNoteID.text))
   End If
   


        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & CHR(13)
                    Msg = Msg + "do you new Operation?"
        
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.TITLE) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
                End If

        End Select

        TxtModFlg.Text = "R"
  FillOrderGrid
  FillOrderGrid2
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry... Error During Saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

  
  
            StrTempDes = " √„— «·‘—«¡ —ﬁ„ " & TxtNoteSerial1 & "  ··„Ê—œ   " & DBCboClientName.Text & " „·«ÕŸ«  " & TxtBillComment.Text
            LngDevNO = LngDevNO + 1
 
Notevalue = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
  
 Dim Account_Code_dynamic101 As String
  Dim Account_Code_dynamic102 As String
 
   Account_Code_dynamic101 = get_account_code_branch(101, my_branch)
            Account_Code_dynamic102 = get_account_code_branch(102, my_branch)
              
'll:
   LngDevNO = 0
  
  
 If Notevalue > 0 Then
       ' «·„œÌ‰
      
   LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic101
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = Account_Code_dynamic102
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

            
            
  End If
  
   
 
  
ErrTrap:
End Function

Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "«„— ‘—«— —ﬁ„  " & TxtNoteSerial1 & " „‰ «·„Ê—œ " & DBCboClientName.Text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "Transactions"
Filedname = "Transaction_ID"
ContNo = Me.XPTxtBillID.Text
Notevalue = LblTotalView.Caption


                     If Me.TxtModFlg = "N" Then
                                 CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 8064, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, ToHijriDate(XPDtbBill.value)
                                     TXTNoteID.Text = NoteID
                                    TxtNoteSerial.Text = NoteSerial
                    Else
                                      If TXTNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                    CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 8064, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, ToHijriDate(XPDtbBill.value)
                                                       TXTNoteID.Text = NoteID
                                                  TxtNoteSerial.Text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",remark='" & TxtNoteSerial1 & "'"
                                                    sql = sql & " where NoteID=" & val(TXTNoteID.Text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
ReLineGrid
CREATE_VOUCHER_GE val(TXTNoteID.Text), val(dcBranch.BoundText), user_id, XPDtbBill.value
rs.Resync adAffectCurrent


End Function



Private Sub XPBtnNewClients_Click()

    'With FrmAddNewCustemer
    '    .DealingForm = ShowPrice
    '    .show vbModal
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    'End With

End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        XPTxtTaxValue.locked = False
        lbl(4).Enabled = True
    Else
        XPTxtTaxValue.Text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
    End If

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

MySQL = " SELECT  dbo.Transactions.SippingDate , dbo.Transactions.DeliverDate ,  dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID, dbo.Transaction_Details.ItemDiscountType, "
MySQL = MySQL & "                      dbo.Transaction_Details.ItemDiscount, dbo.Transactions.order_no, dbo.Transactions.Currency_id, dbo.Transaction_Details.Item_ID,"
MySQL = MySQL & "                      dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ColorID,"
MySQL = MySQL & "                      dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
MySQL = MySQL & "                      dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblItemsclasses.SizeName AS ClassName,"
MySQL = MySQL & "                      dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Type, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "                      dbo.TblCustemers.E_mail, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.FaxNumber, dbo.Transaction_Details.ParrtNoCode,"
MySQL = MySQL & "                      dbo.TblUnites.UnitNamee, dbo.Transactions.ModeReceptEq, dbo.Transactions.ModeSupply, dbo.Transactions.DeptID, dbo.TblEmpDepartments.DepartmentName,"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.Transactions.PaymentType, dbo.ApprovalData.levelo, dbo.TbLLevels.Name, dbo.TbLLevels.Namee,"
MySQL = MySQL & "                      dbo.ApprovalData.EmpID, dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, dbo.ApprovalData.NoteID,"
MySQL = MySQL & "                      dbo.ApprovalData.Currcursor, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TbLLevels.LevelID, dbo.ApprovalData.id,"
MySQL = MySQL & "                      dbo.Transaction_Details.ID AS IDTr, dbo.Transactions.PaymentT, dbo.Transactions.ShipingID, dbo.TblShipingData.Name AS ShipName,"
MySQL = MySQL & "                      dbo.TblShipingData.NameE AS ShipNameE, dbo.Transactions.PaymentID, dbo.TblPaymetData.Name AS PaymName,"
MySQL = MySQL & "                      dbo.TblPaymetData.NameE AS PaymNameE"
MySQL = MySQL & " FROM         dbo.TblItemsColors RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItemsSizes RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblPaymetData RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.Transactions INNER JOIN"
MySQL = MySQL & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
MySQL = MySQL & "                     dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
MySQL = MySQL & "                      dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId ON dbo.TblPaymetData.ID = dbo.Transactions.PaymentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblShipingData ON dbo.Transactions.ShipingID = dbo.TblShipingData.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.ApprovalData ON dbo.Transactions.Transaction_ID = dbo.ApprovalData.Transaction_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.Transactions.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID ON dbo.TblItemsSizes.SizeId = dbo.Transaction_Details.ItemSize ON"
MySQL = MySQL & "                      dbo.TblItemsColors.ColorID = dbo.Transaction_Details.ColorID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
MySQL = MySQL & " WHERE      (dbo.Transactions.Transaction_ID =" & val(XPTxtBillID.Text) & ")"



 If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Inventory\PerformaInvoices7Sh.rpt"
              
     Else
       StrFileName = App.path & "\Reports\Inventory\PerformaInvoices7Sh.rpt"
      ' StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices7Sh.rpt"
       
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
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
        xReport.ParameterFields(12).AddCurrentValue TxtBillComment.Text  'RPTCompany_Name_Arabic
        'TxtBillComment
        
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(9).AddCurrentValue LblTotalAll.Caption
      xReport.ParameterFields(10).AddCurrentValue LblDiscountsTotalView.Caption
      xReport.ParameterFields(11).AddCurrentValue LblTotalView.Caption
   '    xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.TITLE
    xReport.ReportAuthor = App.TITLE
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Function print_report2(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
 MySQL = " SELECT     dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, dbo.Transactions.Trans_Discount, "
 MySQL = MySQL & "                     dbo.Transactions.Trans_DiscountType, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,"
 MySQL = MySQL & "                     dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.order_no, dbo.Transactions.Currency_id,"
 MySQL = MySQL & "                     dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemSize,"
 MySQL = MySQL & "                     dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
 MySQL = MySQL & "                     dbo.TblItems.ItemNamee, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblItemsclasses.SizeName AS ClassName,"
 MySQL = MySQL & "                     dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Type, dbo.TblCustemers.Fullcode,"
 MySQL = MySQL & "                     dbo.TblCustemers.E_mail, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.FaxNumber, dbo.Transaction_Details.ParrtNoCode,"
 MySQL = MySQL & "                     dbo.TblUnites.UnitNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.DepartementID,"
 MySQL = MySQL & "                     TblEmpDepartments_2.DepartmentName, TblEmpDepartments_2.DepartmentNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price,"
 MySQL = MySQL & "                     dbo.Transaction_Details.RequestLimit, dbo.Transaction_Details.LastPurchasePrice, dbo.Transaction_Details.LastPurchaseqty,"
 MySQL = MySQL & "                     dbo.Transaction_Details.AverageIssue, dbo.Transaction_Details.AverageIssueyraly, dbo.Transaction_Details.LastPurchaseDate, dbo.Transactions.InternalFlag,"
 MySQL = MySQL & "                     dbo.Transactions.FixesAssetsID, dbo.FixedAssets.code AS Fixedcode, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.Transactions.Emp_ID,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
 MySQL = MySQL & "                     dbo.TblItems.Fullcode AS ItemFullcode, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.ItemCase, dbo.markaas_taklefa.account_name,"
 MySQL = MySQL & "                     dbo.markaas_taklefa.Code AS TaklfaCode, dbo.Transactions.CBoBasedON, dbo.Transactions.OrderType, dbo.Transactions.TransactionComment,"
 MySQL = MySQL & "                     dbo.Transactions.DeptID, TblEmpDepartments_1.DepartmentName AS DepartmentName2, TblEmpDepartments_1.DepartmentNamee AS DepartmentName2E,"
 MySQL = MySQL & "                     dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Shipment_no, dbo.Transactions.shipped,"
 MySQL = MySQL & "                     dbo.Transactions.shipmentMethod, dbo.Transactions.ShipmentPrice, dbo.Transactions.ShipmentArae, dbo.Transactions.ShipmentStatus,"
 MySQL = MySQL & "                     dbo.Transactions.SippingDate, dbo.Transactions.shipmentType, dbo.Transactions.ShippingTypeID, dbo.Transactions.ShippingStatus, dbo.Transactions.Shipping_Pos,"
 MySQL = MySQL & "                     dbo.Transactions.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.ShipingID,"
 MySQL = MySQL & "                     dbo.TblShipingData.Name AS ShipName, dbo.TblShipingData.NameE AS ShipNameE, dbo.Transactions.DeliverDate, dbo.TblCustemers.CusID,"
 MySQL = MySQL & "                     dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress,"
 MySQL = MySQL & "                     dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, dbo.Transaction_Details.Height,"
 MySQL = MySQL & "                     dbo.Transaction_Details.Width, dbo.Transaction_Details.ItemsDetailsNewidea, dbo.Transactions.GrantyPeriod, dbo.Transactions.NoDelivery,"
 MySQL = MySQL & "                     dbo.Transactions.Granty, dbo.Transactions.Cityid, dbo.Transactions.Days1, dbo.TblEmployee.Fullcode AS EmpFullcode,"
 MySQL = MySQL & "                     dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCountriesGovernmentsCities.NameE AS CityNameE, dbo.Transactions.GovernmentID,"
 MySQL = MySQL & "                     dbo.TblCountriesGovernments.GovernmentName, dbo.Transactions.PaymentType, dbo.Transactions.PaymentID2, dbo.Transactions.PeriodReq,"
 MySQL = MySQL & "                     dbo.Transactions.During, dbo.Transaction_Details.Head_Details, dbo.TblCustemers.Address, dbo.TblCustemers.CustomerTypeID,"
 MySQL = MySQL & "                     dbo.TblCustomerType.name AS TypeCustomername, dbo.TblCustomerType.namee AS TypeCustomernameE, dbo.Transactions.ReqStatus, dbo.TblItems.GroupID,"
 MySQL = MySQL & "                     dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.Transaction_Details.New_Rolling, dbo.Transaction_Details.ItemBalance,"
 MySQL = MySQL & "                     dbo.Transaction_Details.ItemSalesYear , dbo.Transaction_Details.Remarks"
 MySQL = MySQL & " FROM         dbo.TblStore RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.Transactions INNER JOIN"
 MySQL = MySQL & "                     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
 MySQL = MySQL & "                     dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
 MySQL = MySQL & "                     dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId INNER JOIN"
 MySQL = MySQL & "                     dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCountriesGovernments ON dbo.Transactions.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Cityid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblShipingData ON dbo.Transactions.ShipingID = dbo.TblShipingData.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblPaymentType ON dbo.Transactions.PaymentID = dbo.TblPaymentType.PaymentID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartments TblEmpDepartments_1 ON dbo.Transactions.DeptID = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.markaas_taklefa ON dbo.Transactions.general_cost_center = dbo.markaas_taklefa.Code LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartments TblEmpDepartments_2 ON dbo.Transactions.DepartementID = TblEmpDepartments_2.DeparmentID ON"
 MySQL = MySQL & "                     dbo.TblStore.StoreID = dbo.Transactions.StoreID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustomerType RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustemers ON dbo.TblCustomerType.id = dbo.TblCustemers.CustomerTypeID ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
 MySQL = MySQL & " WHERE      (dbo.Transactions.Transaction_ID =" & val(XPTxtBillID.Text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PerformaInvoices701.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PerformaInvoices701.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.TITLE
    xReport.ReportAuthor = App.TITLE
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 7, DcboEmp.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.TITLE)

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

Public Sub Cala()
    NewGrid.Calculate 1
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial1.Text = ""
 TxtNoteSerial.Text = ""
 
End Sub

Private Sub XPTxtTaxValue_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    chkshipped.Caption = "shipped"
    Me.Caption = "Purchase Request"
    lbl(46).Caption = "Order Status"
    Me.Label1(2).Caption = "Purchase Request"
   'Me.Caption = "Order Request/Proforma   Invoice"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Approved Status"
    Me.XPTab301.TabCaption(2) = "Quotations"
    Me.XPTab301.TabCaption(3) = "Internal Orders"
    '''''''''''''''''''''''
    ISButton2.Caption = "Attachments"
    lbl(35).Caption = "Offer End"
    lbl(42).Caption = "Shipment Date"
    lbl(43).Caption = "Delivery Date"
    lbl(44).Caption = "internal Request"
    lbl(37).Caption = "Management Offer"
    lbl(36).Caption = "Supplier Name"
    lbl(35).Caption = "Value"
    lbl(41).Caption = "Shipping Method "
    lbl(34).Caption = "Discount Type"
    ''''''''''''''''''''''''''''''
    lbl(18).Caption = "Type"
    Label4.Caption = "ACC. BY"
    Label10.Caption = "Signature"
    lbl(32).Caption = "Sales Person"
    Accredit.Caption = "Accredit"
    Cmd(8).Caption = "Print Pur. Order"
    'Ele(6).Caption = Me.Caption
    lbl(50).Caption = "Discounts"
    lbl(49).Caption = "Net"
     With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
        End With
    lbl(5).Caption = "Ord/P INV. No"
    Frame3.Caption = "LC Data"
    ISButton1.Caption = "View"
    lbl(25).Caption = "Total"
    lbl(63).Caption = "Qty"
    Label2.Caption = "Branch"
    lbl(6).Caption = "Date"
    lbl(7).Caption = "Client"
    lbl(8).Caption = "Store"
    lbl(9).Caption = "Type"
    lbl(10).Caption = "Cost Center"
    lbl(11).Caption = "Project"
    lbl(16).Caption = "Article Section"
    lbl(12).Caption = "Currency"
    lbl(13).Caption = "Country"
    lbl(14).Caption = "Shipment Mode"
    lbl(17).Caption = "Value"
    lbl(15).Caption = "Payment M"
    lbl(28).Caption = "  Remarks"
    lbl(19).Caption = "Kind Of Order"
    lbl(24).Caption = "Expiry Date"
    lbl(20).Caption = "Credit Bank"
    lbl(21).Caption = "Credit Curr."
    lbl(22).Caption = "Credit No."
    lbl(23).Caption = "Value"
    'ISButton1.Caption = "Show Port Data"
  '  Label1.Caption = "LC NO:"
    'Label2.Caption = "Supp info No."
    Label3.Caption = "Supp info Date"
    Label5.Caption = "Exp Del Date"
    Label6.Caption = "Act Del Date"
    Label7.Caption = "Late Date"
    Label8.Caption = "Exp Arrival Date"
    Label9.Caption = "Comments"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "item name"

    lbl(29).Caption = "Status"
    lbl(27).Caption = "Qty"
    lbl(26).Caption = "Price"

    lbl(3).Caption = "Total"
    lbl(1).Caption = "By"
    lbl(0).Caption = "Currenr rec."
    lbl(2).Caption = "Total rec."

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    CmdConvert.Caption = "Convert To Bill"
    CmdTemplate.Caption = "Insert template"
    
     With Me.FG
        .TextMatrix(0, .ColIndex("FoxyNo")) = "Program NO."
        .TextMatrix(0, .ColIndex("Shipping")) = "Shipping"
    End With
    
    With Me.Grid3
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order_no"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date"
        .TextMatrix(0, .ColIndex("BranchName")) = "BranchNo"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"
    End With
    
    With Me.GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level Name"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approv Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
    
     With Me.VSFlexGrid1
         .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order_no"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date"
        .TextMatrix(0, .ColIndex("BranchName")) = "BranchNo"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"

    End With
           With CboPaymentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With
 End Sub

Function Closeorders(Optional gridno As Integer = 0)
   ' On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
     
If gridno = 0 Then
    With Grid3

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.Text) & ",nots2=" & Me.TxtNoteSerial1.Text & " where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
  Else
  
  With VSFlexGrid1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.Text) & ",nots2='" & Me.TxtNoteSerial1.Text & "'  where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
   
  End If
End Function
Function FillOrderGrid2()
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.VSFlexGrid1
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
  '  My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.Transaction_ID,dbo.Transactions.NoteSerial1 , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=38  AND  dbo.Transactions.Approved = 1   AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

If Me.TxtModFlg = "N" Then
My_SQL = "SELECT  dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved,dbo.Transactions.StoreID"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0" '  and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38 and OrderType=3 )   and CLOSED= 0  " '  and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
Else
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved,dbo.Transactions.StoreID "
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
 'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 "
My_SQL = My_SQL & "  WHERE    nots ='" & val(Me.XPTxtBillID.Text) & "'"

End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.VSFlexGrid1
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
         
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           .TextMatrix(i, .ColIndex("StoreId")) = IIf(IsNull(RsExp.Fields("StoreId").value), "", RsExp.Fields("StoreId").value)
           
 
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreName").value), "", RsExp.Fields("StoreName").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_name").value), "", RsExp.Fields("branch_name").value)
                                
Else
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreNamee").value), "", RsExp.Fields("StoreNamee").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_namee").value), "", RsExp.Fields("branch_namee").value)
                                
End If

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    'GRID2.Visible = True

End Function


Function FillOrderGrid()
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.Grid3
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
  '  My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.Transaction_ID,dbo.Transactions.NoteSerial1 , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=38  AND  dbo.Transactions.Approved = 1   AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

If Me.TxtModFlg = "N" Then
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 46) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 46)   and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)

Else
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 "
My_SQL = My_SQL & "  WHERE    nots ='" & val(Me.XPTxtBillID.Text) & "'"

End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid3
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
         
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           

                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreName").value), "", RsExp.Fields("StoreName").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_name").value), "", RsExp.Fields("branch_name").value)
                                
Else
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreNamee").value), "", RsExp.Fields("StoreNamee").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_namee").value), "", RsExp.Fields("branch_namee").value)
                                
End If

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID2.Visible = True

End Function


