VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmPO10 
   Caption         =   " √Ê«„— «·‘—«¡  "
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15855
   HelpContextID   =   340
   Icon            =   "FrmPO10.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   15855
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9825
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   15855
      _cx             =   27966
      _cy             =   17330
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
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   9255
         Width           =   18255
         _cx             =   32200
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   0
            Left            =   13410
            TabIndex        =   11
            Top             =   90
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
            Left            =   11970
            TabIndex        =   12
            Top             =   90
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
            Left            =   10095
            TabIndex        =   13
            Top             =   90
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
            Left            =   8730
            TabIndex        =   14
            Top             =   90
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
            Left            =   7800
            TabIndex        =   15
            Top             =   90
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
            Left            =   6885
            TabIndex        =   16
            Top             =   90
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
            Left            =   2910
            TabIndex        =   17
            Top             =   90
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   5880
            TabIndex        =   18
            Top             =   90
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   3705
            TabIndex        =   19
            Top             =   90
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   1920
            TabIndex        =   98
            Top             =   90
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   4680
            TabIndex        =   167
            Top             =   90
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   465
         Index           =   3
         Left            =   15
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   8775
         Width           =   15645
         _cx             =   27596
         _cy             =   820
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
            Left            =   15540
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   -210
            Visible         =   0   'False
            Width           =   1125
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3930
            TabIndex        =   22
            Top             =   45
            Width           =   2865
            _ExtentX        =   5054
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
            Height          =   405
            Left            =   7830
            TabIndex        =   146
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   315
            Index           =   49
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   75
            Width           =   645
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
            Height          =   405
            Left            =   7830
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   30
            Visible         =   0   'False
            Width           =   1425
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
            Height          =   405
            Left            =   10170
            TabIndex        =   143
            Top             =   0
            Width           =   1245
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
            Height          =   405
            Left            =   10275
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   0
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   315
            Index           =   50
            Left            =   11595
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   75
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   330
            Index           =   63
            Left            =   15525
            TabIndex        =   82
            Top             =   135
            Visible         =   0   'False
            Width           =   930
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
            Height          =   390
            Left            =   14730
            TabIndex        =   81
            Top             =   0
            Visible         =   0   'False
            Width           =   2445
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
            Height          =   435
            Left            =   12300
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   315
            Index           =   25
            Left            =   14175
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   60
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   285
            Index           =   3
            Left            =   16830
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   75
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   270
            Index           =   0
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   270
            Index           =   2
            Left            =   900
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   120
            Width           =   945
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   300
            Left            =   2025
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   105
            Width           =   615
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
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
            Height          =   345
            Index           =   1
            Left            =   6660
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   75
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3105
         Index           =   0
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   15840
         _cx             =   27940
         _cy             =   5477
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
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   840
            Width           =   1905
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√„—"
            Height          =   255
            Index           =   1
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ·»"
            Height          =   255
            Index           =   0
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtPO6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   165
            Width           =   1905
         End
         Begin VB.TextBox TxtPayment 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   960
            TabIndex        =   178
            Top             =   165
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox TxtModeSupply 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   8925
            TabIndex        =   171
            Top             =   4800
            Visible         =   0   'False
            Width           =   5445
         End
         Begin VB.TextBox TxtModeRecept 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   360
            TabIndex        =   168
            Top             =   3240
            Visible         =   0   'False
            Width           =   6450
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   3480
            TabIndex        =   162
            Top             =   1035
            Width           =   3330
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10110
            TabIndex        =   158
            Top             =   2220
            Width           =   1305
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   12015
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   2220
            Width           =   2355
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   255
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   -540
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.TextBox TxtPONo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8925
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   -315
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5850
            TabIndex        =   141
            Top             =   660
            Width           =   960
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   270
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   3360
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   4605
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12825
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   165
            Width           =   1545
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   810
            Left            =   855
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   1965
            Width           =   5955
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13335
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1365
            Width           =   1035
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13335
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   915
            Width           =   1035
         End
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   660
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3315
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  «·«⁄ „«œ"
            Height          =   840
            Left            =   -1875
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   -3540
            Visible         =   0   'False
            Width           =   4665
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
               Format          =   104660993
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
               Format          =   104660993
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
               Format          =   104660993
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
               Format          =   104660993
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
               Format          =   104660993
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
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2700
            Left            =   2475
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3660
            Visible         =   0   'False
            Width           =   6945
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
               Format          =   104660993
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
            Height          =   2520
            Left            =   -2400
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   -5685
            Visible         =   0   'False
            Width           =   8025
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
            Left            =   4650
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   -330
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   12825
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -330
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   -525
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2415
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   -1650
            Visible         =   0   'False
            Width           =   1020
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
            Width           =   2370
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   8925
            TabIndex        =   2
            Top             =   915
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   8925
            TabIndex        =   3
            Top             =   1365
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   12825
            TabIndex        =   1
            Top             =   540
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   615
            Left            =   8670
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   1740
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   1085
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
            ButtonImage     =   "FrmPO10.frx":038A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   855
            Left            =   1860
            TabIndex        =   33
            Top             =   3600
            Visible         =   0   'False
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1508
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
            Left            =   17880
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2835
            Width           =   4575
            _cx             =   8070
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
            Left            =   4665
            TabIndex        =   84
            Top             =   165
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   8925
            TabIndex        =   94
            Top             =   540
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   3495
            TabIndex        =   142
            Top             =   660
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   615
            Left            =   0
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   1085
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
            ButtonImage     =   "FrmPO10.frx":0724
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbDetpartment 
            Height          =   315
            Left            =   8925
            TabIndex        =   170
            Top             =   1800
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPayment 
            Height          =   315
            Left            =   240
            TabIndex        =   179
            Top             =   480
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbShiping 
            Height          =   315
            Left            =   840
            TabIndex        =   180
            Top             =   1560
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker SippingDate 
            Height          =   315
            Left            =   12480
            TabIndex        =   182
            Top             =   2640
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DeliverDate 
            Height          =   315
            Left            =   8925
            TabIndex        =   184
            Top             =   2640
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   11
            Left            =   0
            TabIndex        =   194
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   285
            Index           =   45
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ÿ·» œ«Œ·Ì"
            Height          =   285
            Index           =   44
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· Ê’Ì·"
            Height          =   270
            Index           =   43
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   2640
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·‘Õ‰"
            Height          =   270
            Index           =   42
            Left            =   14280
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   2640
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·‘Õ‰ Ê«· Ê—Ìœ"
            Height          =   375
            Index           =   41
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   1680
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ÿ—Ìﬁ… ««” ·«„ «·„Ê«œ"
            Height          =   375
            Index           =   38
            Left            =   7065
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   1440
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   375
            Index           =   28
            Left            =   7290
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   2325
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   390
            Index           =   32
            Left            =   7335
            TabIndex        =   174
            Top             =   585
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ «·‰ﬁœÌ"
            Height          =   600
            Index           =   36
            Left            =   6855
            TabIndex        =   173
            Top             =   1035
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «· Ê—Ìœ"
            Height          =   375
            Index           =   39
            Left            =   14265
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   2280
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «·ÿ«·»…"
            Height          =   375
            Index           =   37
            Left            =   14265
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   1800
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   450
            Index           =   35
            Left            =   11130
            TabIndex        =   161
            Top             =   2220
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   330
            Index           =   34
            Left            =   14655
            TabIndex        =   160
            Top             =   2220
            Width           =   855
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
            Height          =   360
            Index           =   55
            Left            =   9600
            TabIndex        =   159
            Top             =   2265
            Width           =   345
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   330
            Index           =   33
            Left            =   11550
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   -315
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·ÿ·»Ì…"
            Height          =   330
            Index           =   18
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   3360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   510
            Left            =   11595
            TabIndex        =   95
            Top             =   540
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·Â"
            Height          =   390
            Index           =   12
            Left            =   7155
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   165
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   330
            Index           =   9
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   480
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«„—"
            Height          =   375
            Index           =   5
            Left            =   14265
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   165
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«„—"
            Height          =   270
            Index           =   6
            Left            =   14265
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   405
            Index           =   7
            Left            =   14745
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   915
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   375
            Index           =   8
            Left            =   14265
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1365
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5895
         Left            =   0
         TabIndex        =   101
         Top             =   2805
         Width           =   15855
         _cx             =   27966
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
         Picture(0)      =   "FrmPO10.frx":0ABE
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   5430
            Left            =   16500
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   45
            Width           =   15765
            _cx             =   27808
            _cy             =   9578
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
               FormatString    =   $"FrmPO10.frx":0E58
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
            Height          =   5430
            Index           =   15
            Left            =   45
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   45
            Width           =   15765
            _cx             =   27808
            _cy             =   9578
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
            _GridInfo       =   $"FrmPO10.frx":0F9B
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5400
               Index           =   16
               Left            =   15
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   15
               Width           =   15735
               _cx             =   27755
               _cy             =   9525
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
                  Height          =   8205
                  Index           =   5
                  Left            =   0
                  TabIndex        =   112
                  TabStop         =   0   'False
                  Top             =   -465
                  Width           =   15795
                  _cx             =   27861
                  _cy             =   14473
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
                     Height          =   915
                     Left            =   600
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   5100
                     Visible         =   0   'False
                     Width           =   1935
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
                        _Image          =   "FrmPO10.frx":0FD1
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
                     Height          =   780
                     Index           =   2
                     Left            =   0
                     TabIndex        =   118
                     TabStop         =   0   'False
                     Top             =   1350
                     Width           =   15435
                     _cx             =   27226
                     _cy             =   1376
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
                        Left            =   -6060
                        MaxLength       =   20
                        RightToLeft     =   -1  'True
                        TabIndex        =   188
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   3585
                     End
                     Begin VB.ComboBox CboItemCase 
                        Height          =   315
                        Left            =   4725
                        RightToLeft     =   -1  'True
                        Style           =   2  'Dropdown List
                        TabIndex        =   121
                        Top             =   300
                        Width           =   1875
                     End
                     Begin VB.TextBox TxtQuantity 
                        Alignment       =   1  'Right Justify
                        Height          =   315
                        Left            =   2670
                        MaxLength       =   10
                        RightToLeft     =   -1  'True
                        TabIndex        =   120
                        Top             =   300
                        Width           =   1995
                     End
                     Begin VB.TextBox TxtPrice 
                        Alignment       =   1  'Right Justify
                        Height          =   315
                        Left            =   720
                        MaxLength       =   10
                        RightToLeft     =   -1  'True
                        TabIndex        =   119
                        Top             =   300
                        Width           =   1950
                     End
                     Begin MSDataListLib.DataCombo DCboItemsName 
                        Height          =   315
                        Left            =   6600
                        TabIndex        =   122
                        Top             =   300
                        Width           =   5565
                        _ExtentX        =   9816
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCboItemsCode 
                        Height          =   315
                        Left            =   12225
                        TabIndex        =   123
                        Top             =   300
                        Width           =   3030
                        _ExtentX        =   5345
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin ImpulseButton.ISButton CmdAdd 
                        Height          =   480
                        Left            =   60
                        TabIndex        =   124
                        Top             =   255
                        Width           =   600
                        _ExtentX        =   1058
                        _ExtentY        =   847
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
                        ButtonImage     =   "FrmPO10.frx":0FE9
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
                        Left            =   13140
                        RightToLeft     =   -1  'True
                        TabIndex        =   129
                        Top             =   0
                        Width           =   2955
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈”„ «·’‰›"
                        Height          =   255
                        Index           =   30
                        Left            =   9510
                        RightToLeft     =   -1  'True
                        TabIndex        =   128
                        Top             =   0
                        Width           =   2835
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Õ«·… «·’‰›"
                        Height          =   255
                        Index           =   29
                        Left            =   4965
                        RightToLeft     =   -1  'True
                        TabIndex        =   127
                        Top             =   0
                        Width           =   1635
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·ﬂ„Ì…"
                        Height          =   255
                        Index           =   27
                        Left            =   2910
                        RightToLeft     =   -1  'True
                        TabIndex        =   126
                        Top             =   0
                        Width           =   1755
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”⁄—"
                        Height          =   255
                        Index           =   26
                        Left            =   780
                        RightToLeft     =   -1  'True
                        TabIndex        =   125
                        Top             =   0
                        Width           =   1890
                     End
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   3135
                     Left            =   120
                     TabIndex        =   130
                     Top             =   2130
                     Width           =   15495
                     _cx             =   27331
                     _cy             =   5530
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
                     Cols            =   15
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmPO10.frx":1383
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
                     Top             =   5460
                     Width           =   3570
                     _ExtentX        =   6297
                     _ExtentY        =   1111
                     ButtonWidth     =   609
                     ButtonHeight    =   1005
                     Appearance      =   1
                     _Version        =   393216
                  End
                  Begin ImpulseButton.ISButton Accredit 
                     Height          =   510
                     Left            =   4230
                     TabIndex        =   151
                     Top             =   5400
                     Width           =   2190
                     _ExtentX        =   3863
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
                     Left            =   7080
                     RightToLeft     =   -1  'True
                     TabIndex        =   177
                     Top             =   6000
                     Width           =   1155
                  End
                  Begin VB.Label LblItemsCount 
                     Alignment       =   2  'Center
                     BackColor       =   &H00404040&
                     ForeColor       =   &H0000FFFF&
                     Height          =   285
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   132
                     Top             =   5340
                     Width           =   600
                  End
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Label12"
                  Height          =   885
                  Left            =   3390
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   240
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3285
                  Index           =   62
                  Left            =   3270
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   1395
                  Width           =   540
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5400
               Index           =   9
               Left            =   15
               TabIndex        =   105
               TabStop         =   0   'False
               Top             =   15
               Width           =   15735
               _cx             =   27755
               _cy             =   9525
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
                  Height          =   2880
                  Left            =   5325
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   1395
                  Width           =   1095
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   4320
                  Left            =   4050
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   900
                  Width           =   735
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
                  Height          =   3345
                  Index           =   69
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   1395
                  Width           =   300
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   2820
                  Index           =   68
                  Left            =   4785
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1680
                  Width           =   300
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   2880
                  Index           =   67
                  Left            =   3210
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   1395
                  Width           =   540
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5430
            Left            =   16800
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   45
            Width           =   15765
            _cx             =   27808
            _cy             =   9578
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
               FormatString    =   $"FrmPO10.frx":15B9
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
            Height          =   5430
            Left            =   17100
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   45
            Width           =   15765
            _cx             =   27808
            _cy             =   9578
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO10.frx":16D6
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
         Width           =   15795
         _cx             =   27861
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
         Caption         =   " √Ê«„— «·‘—«¡  "
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
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1755
            TabIndex        =   136
            Top             =   105
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":17F3
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
            Left            =   1005
            TabIndex        =   137
            Top             =   105
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":1B8D
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
            Left            =   2535
            TabIndex        =   138
            Top             =   105
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":1F27
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
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":22C1
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
            Left            =   10800
            Picture         =   "FrmPO10.frx":265B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
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
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   360
            Width           =   7140
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
Attribute VB_Name = "FrmPO10"
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

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            Dccurrency.BoundText = 1
        '    FG.SetFocus
            Fg.Col = Fg.ColIndex("Code")
            Fg.Row = Fg.Rows - 1
            Me.CboPriceType.ListIndex = 0
                   CboPayMentType.ListIndex = 0
                   
Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = False
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
                                   Me.dcBranch.Enabled = False
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = False
                                    
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

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id

        Case 2
            Dim Msg  As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
If SystemOptions.PoCreateVoucher = True Then
  If CheckAcconts = False Then Exit Sub
End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = GridTransType.purchaseOrderApproved
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ «Ê«„— «·‘—«¡ "
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            PrintReport
         Case 10
         print_report
         
         Case 11
           ShowGL_cc TxtNoteSerial.text, , 200
           
        
        Case 8
            On Error GoTo ErrTrap

            If XPTxtBillID.text <> "" Then
                Set SaleReport = New ClsSaleReport
                SaleReport.ShowPrice XPTxtBillID.text, 6, DcboEmp.text, val(DBCboClientName.BoundText)
            End If

            '        PrintReport1 (Txt_order_no.text)
        Case 6
            Unload Me
            
            
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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
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

        For RowNum = 1 To Fg.Rows - 1

            If .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Code")) <> "" Then
                .Fg.Rows = .Fg.Rows + 1
            End If

            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Name")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Name")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Name")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Code")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("ItemCase")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("HaveSerial")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("HaveSerial")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("HaveSerial")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Count")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Price")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Price")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("DiscountType")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
        
            StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .Fg.Cell(flexcpData, .Fg.Rows - 1, Fg.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .Fg.TextMatrix(.Fg.Rows - 1, Fg.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))
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
    TxtSearchCode.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

    TxtSearchCode.text = fullcode

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
 
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
                    
        FrmSearchSerial.XPTxtCode.text = DCboItemsCode.text
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
                    
        FrmSearchSerial.XPTxtCode.text = DCboItemsCode.text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 2
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
   If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

                 If CheckStoreCoding(val(dcBranch.BoundText), 29) = True Then
                ' TxtNoteSerial.text = ""
                TxtNoteSerial1.text = ""
            
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

Private Sub dcBranch_Change()
  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtNoteSerial1.text = ""
TxtNoteSerial.text = ""
End If



End Sub

Private Sub DcBranch_Click(Area As Integer)
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


Private Sub FG_AfterEdit(ByVal Row As Long, _
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
    If Col = Fg.ColIndex("Code") Or Col = Fg.ColIndex("Name") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , , , , , , , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("UnitID") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("UnitID")), , , , , , , , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("Count") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , (Fg.TextMatrix(Row, Fg.ColIndex("Count"))), , , , , , , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("Price") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , (Fg.TextMatrix(Row, Fg.ColIndex("Price"))), , , , , , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("ColorID") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , , , Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("ColorID")), , , , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("ItemSize") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , , , , Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("ItemSize")), , , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("ClassId") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , , , , , Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("ClassId")), , , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("DiscountType") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , , , , , , Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("DiscountType")), , , Me.Txt_order_no
    ElseIf Col = Fg.ColIndex("DiscountVal") Then
        RegisterItemData Me.name, Me.TxtModFlg, Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Code")), Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")), , , , , , , , , Fg.TextMatrix(Row, Fg.ColIndex("DiscountVal")), , Me.Txt_order_no

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
ReLineGrid
End Sub

Private Sub Fg_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
        FrmAddNewItem.DealingForm = ShowPrice
        FrmAddNewItem.show vbModal
    End If

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
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
   Dim sum As Double
   IntCounter = 0
   sum = 0
     With Fg

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
AddTofaforites Me.name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.text)
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
    Retrive (val(XPTxtBillID.text))

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
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(XPTxtBillID.text)
                  RSApproval("NoteSerial").value = TxtNoteSerial1.text
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
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
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Refresh

   
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
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))

            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            If Transaction_Type = 0 Then
                Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            End If
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                Fg.TextMatrix(Num, Fg.ColIndex("HaveSerial")) = True
            End If
        
            Fg.Cell(flexcpData, Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If Fg.Rows > 10 Then
                If Num = 8 Then Fg.Refresh
            End If

        Next Num

    End If

    TxtFillData.text = "F"
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
 StrSQL = StrSQL + " and Transaction_Type =38  "
    Set Rs1 = New ADODB.Recordset
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    TxtBillComment.text = IIf(IsNull(Rs1("TransactionComment")), "", (Rs1("TransactionComment").value))
    DCboStoreName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
     DBCboClientName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
'     DBCboClientName.BoundText = IIf(IsNull(Rs1("CusID")), 0, (Rs1("CusID").value))
'     DBCboClientName.BoundText = IIf(IsNull(Rs1("CusID")), 0, (Rs1("CusID").value))
     DcbDetpartment.BoundText = IIf(IsNull(Rs1("DepartementID")), 0, (Rs1("DepartementID").value))
    
   Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Refresh
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
        Fg.Rows = RsDetails.RecordCount + 1

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
    
        
            Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
           ' Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            
         If SystemOptions.poWithatotalQty = False Then
             Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - IIf(IsNull(RsDetails("ItemBalance")), 0, (RsDetails("ItemBalance").value))
          Else
          Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
          End If
            
            
            
            Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), 0, (RsDetails("showPrice").value))
        
            Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                Fg.TextMatrix(Num, Fg.ColIndex("HaveSerial")) = True
            End If
        
          '  Fg.Cell(flexcpData, Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("itmemunitid")), "", (RsDetails("itmemunitid").value))
          '  Fg.TextMatrix(Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        

 ' FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID"))))
        



      'sa  FG.TextMatrix(Num, FG.ColIndex("RequestLimit")) = IIf(IsNull(RsDetails("RequestLimit")), 0, (RsDetails("RequestLimit").value))
      'sa  FG.TextMatrix(Num, FG.ColIndex("LastPurchaseDate")) = IIf(IsNull(RsDetails("LastPurchaseDate")), "", (RsDetails("LastPurchaseDate").value))
      'sa  FG.TextMatrix(Num, FG.ColIndex("LastPurchasePrice")) = IIf(IsNull(RsDetails("LastPurchasePrice")), 0, (RsDetails("LastPurchasePrice").value))
      'sa  FG.TextMatrix(Num, FG.ColIndex("LastPurchaseqty")) = IIf(IsNull(RsDetails("LastPurchaseqty")), 0, (RsDetails("LastPurchaseqty").value))
      'sa  FG.TextMatrix(Num, FG.ColIndex("AverageIssue")) = IIf(IsNull(RsDetails("AverageIssue")), 0, (RsDetails("AverageIssue").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If Fg.Rows > 10 Then
                If Num = 8 Then Fg.Refresh
            End If

        Next Num
End If
    End If
End Sub

Private Sub TxtPO6_Change()
If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then
RetriveoOrderPO6 , Me.TxtPO6.text
End If
End Sub

Private Sub TxtPO6_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.text <> "R" Then
If KeyCode = vbKeyF3 Then
FrmBuySearch.DealingForm = GridTransType.internalorder
  FrmBuySearch.Index = 6
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
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
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
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
    With Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
       
    End With
 
    fillOrders (1)
    
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
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
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
       
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            
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
   
    On Error GoTo ErrTrap
'  If SystemOptions.UserInterface = ArabicInterface Then
'                FG.ColComboList(FG.ColIndex("Shipping")) = "#1;  «·„‘ —Ì|#2; «·»«∆⁄"
'            ElseIf SystemOptions.UserInterface = EnglishInterface Then
'               FG.ColComboList(FG.ColIndex("Shipping")) = "#1;Buyer |#2;Seller "
'            End If
            
   ' If GeneralPriceType = 0 Then
        ScreenNameArabic = "  √Ê«„— «·‘—«¡ "
        ScreenNameEnglish = "Purchase  Quotations Requests "
        CurrentTransactionType = 29
  
   ' End If

    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

    Me.Caption = ScreenNameArabic
    Ele(6).Caption = ScreenNameArabic

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
    Set NewGrid.Grid = Fg
  NewGrid.GridTrans = GridTransType.purchaseOrderApproved
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
Set NewGrid.customer = Me.DBCboClientName
Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    
    
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DcboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.STORENAME = DCboStoreName
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
        Set NewGrid.DtpBillDate = Me.XPDtbBill
        
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
      Resize_Form Me, TransactionSize
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Fg.WallPaper = BGround.Picture
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

    NewGrid.fillgrid

    With CboPayMentType
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
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish

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
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·”‰œ   " & Txt_order_no.text & Chr(13) & " «· «—ÌŒ " & XPDtbBill.value & Chr(13) & "«‰Ê⁄ «·”‰œ  " & CboPriceType.text & Chr(13) & " «·„Œ“‰  " & DCboStoreName.text & Chr(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & Chr(13) & " —ﬁ„ «·«⁄ „«œ    " & TxtLcNo
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & "Vchr . No   " & Txt_order_no.text & Chr(13) & " Date " & XPDtbBill.value & Chr(13) & " Type  " & CboPriceType.text & Chr(13) & " Store  " & DCboStoreName.text & Chr(13) & " Customer/ Supplier " & DBCboClientName.text & Chr(13) & " Lc NO    " & TxtLcNo
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", , , Me.Txt_order_no
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D", "", , , Me.Txt_order_no
    End If
    
End Function

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

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
            Fg.Editable = flexEDNone
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
            Fg.Enabled = True
            Fg.Rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Fg.Editable = flexEDKbdMouse
        
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
        
            Fg.Enabled = True
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Fg.Editable = flexEDKbdMouse
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

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    
   If SystemOptions.PoCreateVoucher = True Then
    Me.TXTNoteID.text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
  Me.TxtNoteSerial.text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
  End If


    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    Txt_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    TxtPONo.text = IIf(IsNull(rs("PONo").value), "", rs("PONo").value)
    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
      CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
Me.TxtPayment.text = IIf(IsNull(rs("PaymentT").value), "", (rs("PaymentT").value))
XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
  XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
  
    If rs("shipped").value = True Then
        chkshipped.value = vbChecked
    Else
        chkshipped.value = Unchecked
    End If

    Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").value), "", rs("countryid").value)

  If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If
    
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
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
 Me.TxtModeRecept.text = IIf(IsNull(rs("ModeReceptEq").value), "", (rs("ModeReceptEq").value))
 Me.TxtModeSupply.text = IIf(IsNull(rs("ModeSupply").value), "", (rs("ModeSupply").value))
''//

    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    TxtLcNo.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
''/// 11 05 2015
Me.DcbPayment.BoundText = IIf(IsNull(rs("PaymentID").value), "", rs("PaymentID").value)
Me.DcbShiping.BoundText = IIf(IsNull(rs("ShipingID").value), "", rs("ShipingID").value)
''//
    If Txt_order_no <> "" Then
        Me.TxtNoteSerial1.text = Txt_order_no
    End If
''// 25 05 2015
SippingDate.value = IIf(IsNull(rs("SippingDate").value), Date, (rs("SippingDate").value))
DeliverDate.value = IIf(IsNull(rs("DeliverDate").value), Date, (rs("DeliverDate").value))
TxtPO6.text = IIf(IsNull(rs("NotSeialPO6").value), "", (rs("NotSeialPO6").value))

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

 
        
    Lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

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
  
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Refresh
   ' StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
   ' StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
  StrSQL = " SELECT     dbo.TblItems.HaveSerial AS Expr1, *"
  StrSQL = StrSQL & " FROM         dbo.TblItems INNER JOIN"
  StrSQL = StrSQL & "                    dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 
  StrSQL = StrSQL & "                   dbo.TblCountriesData ON dbo.Transaction_Details.countrisid = dbo.TblCountriesData.CountryID"
  StrSQL = StrSQL & " Where (dbo.Transaction_Details.Transaction_ID =" & val(rs("Transaction_ID").value) & ")"
   

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        ''//
        
          
          Fg.TextMatrix(Num, Fg.ColIndex("catalog")) = IIf(IsNull(RsDetails("catalog").value), "", (RsDetails("catalog").value))
         '   FG.TextMatrix(Num, FG.ColIndex("countris")) = IIf(IsNull(RsDetails("CountryName")), "", Trim(RsDetails("CountryName").value))
         '   FG.TextMatrix(Num, FG.ColIndex("Shipping")) = IIf(IsNull(RsDetails("Shipping")), "", (RsDetails("Shipping").value))
        
           Fg.TextMatrix(Num, Fg.ColIndex("countris")) = IIf(IsNull(RsDetails("countrisid")), "", Trim(RsDetails("countrisid").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Shipping")) = IIf(IsNull(RsDetails("Shipping")), "", (RsDetails("Shipping").value))
 
 
        ''//
            Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
           
            Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                Fg.TextMatrix(Num, Fg.ColIndex("HaveSerial")) = True
            End If
            Fg.TextMatrix(Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        If SystemOptions.UserInterface = ArabicInterface Then
           Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
           Else
           Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemNamee")), "", Trim(RsDetails("ItemNamee").value))
           End If
            Fg.Cell(flexcpData, Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
           
           
               Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
 Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
 
 
 
            Fg.TextMatrix(Num, Fg.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If Fg.Rows > 10 Then
                If Num = 8 Then Fg.Refresh
            End If

        Next Num

    End If
fillapprovData
FillOrderGrid
FillOrderGrid2
ReLineGrid

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Sub GRID3_Click()

    With Fg
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
            
            End If

        Next i

    End With

Else


    With VSFlexGrid1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
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
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = Fg.Rows
    
        If Fg.TextMatrix(row_count - 1, Fg.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        Fg.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To Fg.Rows - 1 'RsDetails.RecordCount
    
'            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
'            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
'            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        
            Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
        
            Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("showprice")), "", (RsDetails("showprice").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            If RsDetails("HaveSerial") = True Then
                Fg.TextMatrix(Num, Fg.ColIndex("HaveSerial")) = True
            End If
        
            Fg.Cell(flexcpData, Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
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

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap
 
    Me.LblTotal.Caption = XPTxtSum.text
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–« «·”‰œ   .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ·    Â–« «·”‰œ .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
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

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtBillID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
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
        XPTxtDiscountVal.text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Fg.TextMatrix(1, Fg.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
    End If

    Me.Lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

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

    If Trim$(Me.TxtCashCustomerName.text) = "" Then
        CheckCashCustomer = True
    Else
        StrSQL = "Select * From Transactions Where CashCustomerName='" & Trim$(Me.TxtCashCustomerName.text) & "'"
    
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
'    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.text <> "R" Then
        If DBCboClientName.text = "" And Opt(1).value = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
            Else
                Msg = "Please Select Vendor"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If DCboStoreName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
                Msg = "Select Inventory"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Dccurrency.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ «·⁄„·…"
            Else
                Msg = "Select Currency"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Dccurrency.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    
    If CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPayMentType.SetFocus
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

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPriceType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If



 If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·«„— " & Chr(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & Chr(13)
                Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            Else
                Msg = Msg + " Must Enter Discount Value " & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPCboDiscountType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If




        If XPChkTAX.value = Checked Then
            If XPTxtTaxValue.text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
                Else
                    Msg = "Insert Sales Tax"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                XPTxtTaxValue.SetFocus
                Fg.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
 
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        Set RSTransDetails = New ADODB.Recordset
        RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CboPriceType.ListIndex = 0 Then
            Transaction_Type = CurrentTransactionType
            Sanad_No = CurrentTransactionType
  
 
         
        End If

        my_branch = val(dcBranch.BoundText)

        If TxtNoteSerial1.text = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ    " & Chr(13) & " Enter Vchr No": Exit Sub
                Else
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 170, , Transaction_Type, , val(DCboStoreName.BoundText))
                End If
            End If
        End If
 
        Txt_order_no = Me.TxtNoteSerial1.text
 
        Cn.BeginTrans
        BeginTrans = True
    
        If Me.TxtModFlg.text = "N" Then
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=6"))
            
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
            rs.AddNew
         Else
           If SystemOptions.PoCreateVoucher = True Then
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
                  Cn.Execute StrSQL, , adExecuteNoRecords

         End If
        End If

        Screen.MousePointer = vbArrowHourglass
       rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        rs("branchID").value = val(Me.dcBranch.BoundText)
     rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
       rs("Transaction_ID").value = val(XPTxtBillID.text)
        rs("order_no").value = Txt_order_no.text
    
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
        
        
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Serial").value = TxtTransSerial.text

      rs("PONO").value = IIf(TxtPONo.text = "", Null, (TxtPONo.text))
rs("Transaction_Type").value = CurrentTransactionType

        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("countryid").value = IIf(DataCombo4.BoundText = "", Null, val(DataCombo4.BoundText))
    
        rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    
       If Trim$(Me.TxtCashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
    Else
        rs("CashCustomerName").value = Null
    End If
    
       If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If
    rs("PaymentT").value = Trim$(Me.TxtPayment.text)
    
       If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
  rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
 ''//
   rs("DeptID").value = IIf(DcbDetpartment.BoundText = "", Null, val(Me.DcbDetpartment.BoundText))
   rs("ModeReceptEq").value = IIf(Me.TxtModeRecept.text = "", Null, TxtModeRecept.text)
   rs("ModeSupply").value = IIf(Me.TxtModeSupply.text = "", Null, TxtModeSupply.text)

 ''//
 
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
        rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
        rs("total").value = IIf(XPTxtSum.text = "", Null, val(XPTxtSum.text))
        rs("LcNo").value = IIf(TxtLcNo.text = "", Null, (TxtLcNo.text))
    ''//11 05 2015
    rs("ShipingID").value = val(Me.DcbShiping.BoundText)
    rs("PaymentID").value = val(Me.DcbPayment.BoundText)
    ''//
   ''//25 05 2015
   rs("SippingDate").value = SippingDate.value
   rs("DeliverDate").value = DeliverDate.value
   rs("NotSeialPO6").value = Me.TxtPO6.text
   
        rs.update
    
       CuurentLogdata
  
        If Me.TxtModFlg.text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        For RowNum = 1 To Fg.Rows - 1

            If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                RSTransDetails("order_id").value = val(XPTxtBillID.text)
             
                RSTransDetails("order_no").value = Txt_order_no.text
                ''//
               ' RSTransDetails("countrisid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("countrisid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("countrisid"))))
               ' RSTransDetails("Shipping").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Shipping")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Shipping"))))
                   RSTransDetails("countrisid").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("countris")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("countris"))))
                RSTransDetails("Shipping").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Shipping")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Shipping"))))
     
  
                ''//
                RSTransDetails("catalog").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("catalog")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("catalog"))))
                RSTransDetails("Item_ID").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))))
                RSTransDetails("ShowPrice").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal"))))
            
                RSTransDetails("ColorID").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ColorID")) = ""), 1, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ItemSize")) = ""), "", Trim$(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ClassId")) = ""), 1, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ClassId"))))
            
                RSTransDetails("UnitID").value = IIf(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")) = "", Null, (Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))))
 
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
                LngUnitID = val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")))
                DblQty = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    'RSTransDetails("Price").value = Val(IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))) / RSTransDetails("Quantity").value
                    RSTransDetails("Price").value = val(IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), 0, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                End If

                RSTransDetails.update
            End If

        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        Lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    Closeorders (0)
   Closeorders (1)
   If SystemOptions.PoCreateVoucher = True Then
       createVoucher
      updateNotesValueAndNobytext (val(TXTNoteID.text))
   End If
   


        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & Chr(13)
                    Msg = Msg + "do you new Operation?"
        
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.text = "R"
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
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry... Error During Saving " & Chr(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchId As Integer, UserID As Long _
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
     
    my_branch = BranchId

  
  
            StrTempDes = " √„— «·‘—«¡ —ﬁ„ " & TxtNoteSerial1 & "  ··„Ê—œ   " & DBCboClientName.text & " „·«ÕŸ«  " & TxtBillComment.text
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
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchId) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = Account_Code_dynamic102
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchId) = False Then
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
des = "«„— ‘—«— —ﬁ„  " & TxtNoteSerial1 & " „‰ «·„Ê—œ " & DBCboClientName.text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "Transactions"
Filedname = "Transaction_ID"
ContNo = Me.XPTxtBillID.text
Notevalue = LblTotalView.Caption


                     If Me.TxtModFlg = "N" Then
                                 CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 8064, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, ToHijriDate(XPDtbBill.value)
                                     TXTNoteID.text = NoteID
                                    TxtNoteSerial.text = NoteSerial
                    Else
                                      If TXTNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                    CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 8064, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, ToHijriDate(XPDtbBill.value)
                                                       TXTNoteID.text = NoteID
                                                  TxtNoteSerial.text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",remark='" & TxtNoteSerial1 & "'"
                                                    sql = sql & " where NoteID=" & val(TXTNoteID.text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
ReLineGrid
CREATE_VOUCHER_GE val(TXTNoteID.text), val(dcBranch.BoundText), user_id, XPDtbBill.value
rs.Resync adAffectCurrent


End Function



Private Sub XPBtnNewClients_Click()

    With FrmAddNewCustemer
        .DealingForm = ShowPrice
        .show vbModal
        .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
        .Lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
        .Lbl(0).Caption = "«”„ «·⁄„Ì·"
    End With

End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        XPTxtTaxValue.locked = False
        Lbl(4).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        Lbl(4).Enabled = False
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
MySQL = MySQL & " WHERE      (dbo.Transactions.Transaction_ID =" & val(XPTxtBillID.text) & ")"



 If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Reports\Inventory\PerformaInvoices7Sh.rpt"
     Else
       StrFileName = App.path & "\Reports\Inventory\PerformaInvoices7Sh.rpt"
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
        xReport.ParameterFields(12).AddCurrentValue TxtBillComment.text  'RPTCompany_Name_Arabic
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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.text, 7, DcboEmp.text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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

Public Sub Cala()
    NewGrid.Calculate 1
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial1.text = ""
 TxtNoteSerial.text = ""
 
End Sub

Private Sub XPTxtTaxValue_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
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
    Me.Caption = ScreenNameEnglish
    Ele(6).Caption = ScreenNameEnglish
    'Me.Caption = "Order Request/Proforma   Invoice"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Approved Status"
   Me.XPTab301.TabCaption(2) = "Quotations"
    Me.XPTab301.TabCaption(3) = "Internal Orders"
    
    Lbl(18).Caption = "Type"
    Label4.Caption = "ACC. BY"
    Label10.Caption = "Signature"
    Lbl(32).Caption = "Sales Person"
    Accredit.Caption = "Accredit"
    Cmd(8).Caption = "Print Pur. Order"
    'Ele(6).Caption = Me.Caption
    Lbl(50).Caption = "Discounts"
    Lbl(49).Caption = "Net"

   With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
        End With


    Lbl(5).Caption = "Ord/P INV. No"
    Frame3.Caption = "LC Data"
    ISButton1.Caption = "View"
    Lbl(25).Caption = "Total"
    Lbl(63).Caption = "Qty"
    Label2.Caption = "Branch"
    Lbl(6).Caption = "Date"
    Lbl(7).Caption = "Client"
    Lbl(8).Caption = "Store"
    Lbl(9).Caption = "Type"
    Lbl(10).Caption = "Cost Center"
    Lbl(11).Caption = "Project"
    Lbl(16).Caption = "Article Section"
    Lbl(12).Caption = "Currency"
    Lbl(13).Caption = "Country"
    Lbl(14).Caption = "Shipment Mode"
    Lbl(17).Caption = "Value"
    Lbl(15).Caption = "Payment M"
    Lbl(28).Caption = "  Remarks"
    Lbl(19).Caption = "Kind Of Order"
    Lbl(24).Caption = "Expiry Date"
    Lbl(20).Caption = "Credit Bank"
    Lbl(21).Caption = "Credit Curr."
    Lbl(22).Caption = "Credit No."
    Lbl(23).Caption = "Value"
    'ISButton1.Caption = "Show Port Data"
    Label1.Caption = "LC NO:"
    'Label2.Caption = "Supp info No."
    Label3.Caption = "Supp info Date"
    Label5.Caption = "Exp Del Date"
    Label6.Caption = "Act Del Date"
    Label7.Caption = "Late Date"
    Label8.Caption = "Exp Arrival Date"
    Label9.Caption = "Comments"

    Lbl(31).Caption = "Item Code"
    Lbl(30).Caption = "item name"

    Lbl(29).Caption = "Status"
    Lbl(27).Caption = "Qty"
    Lbl(26).Caption = "Price"

    Lbl(3).Caption = "Total"
    Lbl(1).Caption = "By"
    Lbl(0).Caption = "Currenr rec."
    Lbl(2).Caption = "Total rec."

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

    With Me.Grid3
 
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order_no"

        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date"
        .TextMatrix(0, .ColIndex("BranchName")) = "BranchNo"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"

    End With
 
 
     With Me.VSFlexGrid1
 
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order_no"

        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date"
        .TextMatrix(0, .ColIndex("BranchName")) = "BranchNo"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"

    End With
    
       With CboPayMentType
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
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2=" & Me.TxtNoteSerial1.text & " where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
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
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2='" & Me.TxtNoteSerial1.text & "'  where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
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
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0" '  and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38 and OrderType=3 )   and CLOSED= 0  " '  and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
Else
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
 'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 "
My_SQL = My_SQL & "  WHERE    nots = " & val(Me.XPTxtBillID.text)

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
My_SQL = My_SQL & "  WHERE    nots = " & val(Me.XPTxtBillID.text)

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


