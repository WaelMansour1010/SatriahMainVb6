VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«Ê«„— «·»Ì⁄ Ê«·‘—«¡ Ê«·›Ê« Ì— «·„«·Ì… Ê «·ÿ·»«  «·œ«Œ·Ì…"
   ClientHeight    =   9150
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14025
   HelpContextID   =   340
   Icon            =   "frmPO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14025
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9150
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   14025
      _cx             =   24739
      _cy             =   16140
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   1
         Left            =   15
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8580
         Width           =   16080
         _cx             =   28363
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
            Left            =   12090
            TabIndex        =   18
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
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
            Left            =   10650
            TabIndex        =   19
            Top             =   90
            Width           =   945
            _ExtentX        =   1667
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
            Left            =   8775
            TabIndex        =   20
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
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
            Left            =   7410
            TabIndex        =   21
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   6240
            TabIndex        =   22
            Top             =   90
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
            Left            =   5085
            TabIndex        =   23
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   30
            TabIndex        =   24
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
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
            Left            =   3960
            TabIndex        =   25
            Top             =   90
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   390
            Left            =   2145
            TabIndex        =   26
            Top             =   -270
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   1200
            TabIndex        =   104
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
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
            TabIndex        =   155
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
            Index           =   9
            Left            =   3000
            TabIndex        =   170
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â 2"
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
            Left            =   2040
            TabIndex        =   181
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â 2"
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
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8130
         Width           =   13560
         _cx             =   23918
         _cy             =   767
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
            Left            =   13710
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   30
            Width           =   1230
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4050
            TabIndex        =   29
            Top             =   45
            Width           =   1455
            _ExtentX        =   2566
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
            Height          =   375
            Left            =   6480
            TabIndex        =   152
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   285
            Index           =   49
            Left            =   8145
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   75
            Width           =   705
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
            Height          =   375
            Left            =   6450
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   30
            Width           =   1575
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
            Height          =   375
            Left            =   8880
            TabIndex        =   149
            Top             =   0
            Width           =   1065
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
            Height          =   375
            Left            =   8895
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   285
            Index           =   50
            Left            =   10140
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   75
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   300
            Index           =   63
            Left            =   4920
            TabIndex        =   90
            Top             =   135
            Visible         =   0   'False
            Width           =   960
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
            Height          =   360
            Left            =   3360
            TabIndex        =   89
            Top             =   0
            Visible         =   0   'False
            Width           =   2580
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
            Height          =   405
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   285
            Index           =   25
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   60
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   255
            Index           =   3
            Left            =   13950
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   75
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   0
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   930
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   105
            Width           =   690
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   135
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   1
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   75
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1875
         Index           =   0
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   13905
         _cx             =   24527
         _cy             =   3307
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
         Begin VB.TextBox txtDays2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1920
            TabIndex        =   10
            Top             =   1440
            Width           =   780
         End
         Begin VB.TextBox txtDays1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4800
            TabIndex        =   9
            Top             =   1440
            Width           =   780
         End
         Begin VB.ComboBox CBoStatus 
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4800
            TabIndex        =   147
            Top             =   495
            Width           =   780
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   2040
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   570
            Left            =   720
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   840
            Width           =   4905
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1215
            Width           =   825
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   840
            Width           =   825
         End
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   840
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  «·«⁄ „«œ"
            Height          =   615
            Left            =   -1560
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   360
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox TxtLcNo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4080
               TabIndex        =   73
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100270081
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4560
               TabIndex        =   74
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100270081
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   75
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100270081
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4560
               TabIndex        =   76
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100270081
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   77
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100270081
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   120
               TabIndex        =   91
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
               TabIndex        =   84
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   255
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·«⁄ „«œ"
               Height          =   255
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1815
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   2520
            Visible         =   0   'False
            Width           =   5700
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   62
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100270081
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   63
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
               TabIndex        =   64
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1695
            Left            =   -5520
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   840
            Visible         =   0   'False
            Width           =   6615
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   600
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   3120
               TabIndex        =   49
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
               TabIndex        =   50
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
               TabIndex        =   51
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
               TabIndex        =   52
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
               TabIndex        =   93
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
               TabIndex        =   95
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
               TabIndex        =   96
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
               TabIndex        =   94
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
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
               TabIndex        =   54
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
               TabIndex        =   53
               Top             =   960
               Width           =   1215
            End
         End
         Begin VB.ComboBox CboPriceType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7290
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2250
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   -240
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   -210
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1965
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   -150
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   30
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   -150
            Visible         =   0   'False
            Width           =   1920
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   7305
            TabIndex        =   6
            Top             =   840
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   7305
            TabIndex        =   7
            Top             =   1230
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   10560
            TabIndex        =   3
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100270081
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   450
            Left            =   6975
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   750
            Visible         =   0   'False
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   794
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
            ButtonImage     =   "frmPO.frx":038A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   480
            Left            =   1545
            TabIndex        =   40
            Top             =   915
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   847
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
            Height          =   630
            Index           =   4
            Left            =   14760
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3795
            _cx             =   6694
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
               TabIndex        =   13
               Top             =   210
               Width           =   1815
            End
            Begin VB.TextBox XPTxtTaxValue 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   14
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
               TabIndex        =   45
               Top             =   285
               Width           =   720
            End
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   525
            Left            =   1440
            TabIndex        =   85
            Top             =   1920
            Visible         =   0   'False
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   926
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕÊÌ· ≈·Ì ›« Ê—…"
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
         Begin MSDataListLib.DataCombo Dccurrency 
            Height          =   315
            Left            =   3840
            TabIndex        =   1
            Top             =   120
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   7320
            TabIndex        =   4
            Top             =   480
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   720
            TabIndex        =   5
            Top             =   480
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   390
            Index           =   37
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   390
            Index           =   36
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «· —ﬂÌ»"
            Height          =   390
            Index           =   35
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «· Ê—Ìœ"
            Height          =   390
            Index           =   34
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   1440
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Õ«·…"
            Height          =   285
            Index           =   33
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   120
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   285
            Index           =   32
            Left            =   5910
            TabIndex        =   148
            Top             =   510
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·ÿ·»Ì…"
            Height          =   240
            Index           =   18
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   9585
            TabIndex        =   101
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   270
            Index           =   28
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   840
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·Â"
            Height          =   285
            Index           =   12
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ÿ·»"
            Height          =   240
            Index           =   9
            Left            =   9540
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   270
            Index           =   5
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ÿ·»"
            Height          =   195
            Index           =   6
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì· / «·„Ê—œ"
            Height          =   240
            Index           =   7
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   270
            Index           =   8
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1200
            Width           =   1065
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5415
         Left            =   0
         TabIndex        =   107
         Top             =   2640
         Width           =   13920
         _cx             =   24553
         _cy             =   9551
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
         Caption         =   "«·√’‰«›|Õ«·Â «·«⁄ „«œ|„Õœœ«  «·‘«‘…|‘—Êÿ «·œ›⁄  Ê «·÷„«‰"
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
         Picture(0)      =   "frmPO.frx":0724
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4950
            Left            =   14565
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   45
            Width           =   13830
            _cx             =   24395
            _cy             =   8731
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
               Left            =   120
               TabIndex        =   140
               Tag             =   "1"
               Top             =   240
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
               FormatString    =   $"frmPO.frx":0ABE
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
               TabIndex        =   156
               Top             =   4560
               Width           =   3375
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4950
            Index           =   15
            Left            =   45
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   45
            Width           =   13830
            _cx             =   24395
            _cy             =   8731
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
            _GridInfo       =   $"frmPO.frx":0C01
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4920
               Index           =   16
               Left            =   15
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   15
               Width           =   13800
               _cx             =   24342
               _cy             =   8678
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
                  Height          =   7635
                  Index           =   5
                  Left            =   0
                  TabIndex        =   118
                  TabStop         =   0   'False
                  Top             =   -705
                  Width           =   13845
                  _cx             =   24421
                  _cy             =   13467
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
                     Height          =   855
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   119
                     Top             =   4680
                     Visible         =   0   'False
                     Width           =   1695
                     Begin DBPIXLib.DBPix20 DBPix202 
                        Height          =   855
                        Left            =   240
                        TabIndex        =   120
                        Top             =   -120
                        Width           =   2415
                        _Version        =   131072
                        _ExtentX        =   4260
                        _ExtentY        =   1508
                        _StockProps     =   1
                        _Image          =   "frmPO.frx":0C35
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
                        TabIndex        =   123
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
                        TabIndex        =   122
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
                        TabIndex        =   121
                        Top             =   240
                        Width           =   735
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   690
                     Index           =   2
                     Left            =   30
                     TabIndex        =   124
                     TabStop         =   0   'False
                     Top             =   750
                     Width           =   13500
                     _cx             =   23813
                     _cy             =   1217
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
                     Begin VB.ComboBox CboItemCase 
                        Height          =   315
                        Left            =   5040
                        RightToLeft     =   -1  'True
                        Style           =   2  'Dropdown List
                        TabIndex        =   127
                        Top             =   300
                        Width           =   1890
                     End
                     Begin VB.TextBox TxtQuantity 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Left            =   2820
                        MaxLength       =   10
                        RightToLeft     =   -1  'True
                        TabIndex        =   126
                        Top             =   300
                        Width           =   2160
                     End
                     Begin VB.TextBox TxtPrice 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Left            =   780
                        MaxLength       =   10
                        RightToLeft     =   -1  'True
                        TabIndex        =   125
                        Top             =   300
                        Width           =   2025
                     End
                     Begin MSDataListLib.DataCombo DCboItemsName 
                        Height          =   315
                        Left            =   6945
                        TabIndex        =   128
                        Top             =   300
                        Width           =   3255
                        _ExtentX        =   5741
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCboItemsCode 
                        Height          =   315
                        Left            =   10260
                        TabIndex        =   129
                        Top             =   300
                        Width           =   3195
                        _ExtentX        =   5636
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin ImpulseButton.ISButton CmdAdd 
                        Height          =   375
                        Left            =   75
                        TabIndex        =   130
                        Top             =   270
                        Width           =   630
                        _ExtentX        =   1111
                        _ExtentY        =   661
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
                        ButtonImage     =   "frmPO.frx":0C4D
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
                        Left            =   10440
                        RightToLeft     =   -1  'True
                        TabIndex        =   135
                        Top             =   0
                        Width           =   3015
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈”„ «·’‰›"
                        Height          =   255
                        Index           =   30
                        Left            =   7260
                        RightToLeft     =   -1  'True
                        TabIndex        =   134
                        Top             =   0
                        Width           =   3000
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Õ«·… «·’‰›"
                        Height          =   255
                        Index           =   29
                        Left            =   5280
                        RightToLeft     =   -1  'True
                        TabIndex        =   133
                        Top             =   0
                        Width           =   1680
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·ﬂ„Ì…"
                        Height          =   255
                        Index           =   27
                        Left            =   3060
                        RightToLeft     =   -1  'True
                        TabIndex        =   132
                        Top             =   0
                        Width           =   1890
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”⁄—"
                        Height          =   255
                        Index           =   26
                        Left            =   855
                        RightToLeft     =   -1  'True
                        TabIndex        =   131
                        Top             =   0
                        Width           =   1950
                     End
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   3150
                     Left            =   270
                     TabIndex        =   136
                     Top             =   1575
                     Width           =   13260
                     _cx             =   23389
                     _cy             =   5556
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
                     FormatString    =   $"frmPO.frx":0FE7
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
                     Left            =   495
                     TabIndex        =   137
                     Top             =   4980
                     Width           =   3195
                     _ExtentX        =   5636
                     _ExtentY        =   1111
                     ButtonWidth     =   609
                     ButtonHeight    =   1005
                     Appearance      =   1
                     _Version        =   393216
                  End
                  Begin ImpulseButton.ISButton Accredit 
                     Height          =   510
                     Left            =   3720
                     TabIndex        =   157
                     Top             =   4920
                     Width           =   1845
                     _ExtentX        =   3254
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
                  Begin VB.Label LblItemsCount 
                     Alignment       =   2  'Center
                     BackColor       =   &H00404040&
                     ForeColor       =   &H0000FFFF&
                     Height          =   285
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   138
                     Top             =   4860
                     Width           =   450
                  End
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Label12"
                  Height          =   825
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   240
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   2895
                  Index           =   62
                  Left            =   2805
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   1305
                  Width           =   495
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4920
               Index           =   9
               Left            =   15
               TabIndex        =   111
               TabStop         =   0   'False
               Top             =   15
               Width           =   13800
               _cx             =   24342
               _cy             =   8678
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
                  Height          =   2550
                  Left            =   4665
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   1305
                  Width           =   945
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   3900
                  Left            =   3570
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   840
                  Width           =   600
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
                  Height          =   2955
                  Index           =   69
                  Left            =   3300
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   1305
                  Width           =   270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   2490
                  Index           =   68
                  Left            =   4170
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   1590
                  Width           =   315
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   2550
                  Index           =   67
                  Left            =   2805
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   1305
                  Width           =   495
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   4950
            Left            =   14865
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   45
            Width           =   13830
            _cx             =   24395
            _cy             =   8731
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
            Begin VB.Frame Frame5 
               Caption         =   "œ·«·«  «·«·Ê«‰"
               Height          =   735
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   4200
               Width           =   5175
               Begin VB.Label LblAccCreitria 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF0000&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ã„«·Ì «·„Õœœ« "
                  Height          =   255
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label LbunlDone 
                  Alignment       =   2  'Center
                  BackColor       =   &H000000FF&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label LblDone 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FF00&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„Õﬁﬁ"
                  Height          =   255
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€Ì— „Õﬁﬁ"
                  Height          =   255
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   3990
               Left            =   240
               TabIndex        =   163
               Top             =   120
               Width           =   13260
               _cx             =   23389
               _cy             =   7038
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
               FormatString    =   $"frmPO.frx":1334
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   4950
            Left            =   15165
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   45
            Width           =   13830
            _cx             =   24395
            _cy             =   8731
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
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Text            =   "0"
               Top             =   3720
               Width           =   1095
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   3510
               Left            =   240
               TabIndex        =   169
               Top             =   120
               Width           =   13260
               _cx             =   23389
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmPO.frx":1573
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   21
               Left            =   12600
               TabIndex        =   175
               Top             =   3600
               Width           =   675
               _ExtentX        =   1191
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
               ButtonImage     =   "frmPO.frx":178B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   11520
               TabIndex        =   179
               Top             =   4200
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â  «·÷„«‰"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
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
               Index           =   11
               Left            =   9720
               TabIndex        =   180
               Top             =   4200
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â  ”‰œ «· ”·Ì„"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
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
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   270
               Index           =   39
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   3720
               Width           =   585
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·œ›⁄« "
               Height          =   270
               Index           =   38
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   3720
               Width           =   1065
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   0
         Width           =   13980
         _cx             =   24659
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
         Caption         =   "«Ê«„— «·»Ì⁄ Ê«·‘—«¡ Ê«·›Ê« Ì— «·„«·Ì… Ê «·ÿ·»«  «·œ«Œ·Ì… "
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1860
            TabIndex        =   142
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmPO.frx":1D25
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
            TabIndex        =   143
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
            ButtonImage     =   "frmPO.frx":20BF
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
            Left            =   2670
            TabIndex        =   144
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
            ButtonImage     =   "frmPO.frx":2459
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
            Left            =   165
            TabIndex        =   145
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmPO.frx":27F3
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
            Left            =   4560
            Picture         =   "frmPO.frx":2B8D
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
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   360
            Width           =   7755
         End
      End
   End
End
Attribute VB_Name = "FrmPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim NewGrid1 As New ClsGrid2
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)   As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    Dim intDef As Integer
 On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
                XPTab301.CurrTab = 0
       
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            DcCurrency.BoundText = 1
            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Me.CboPriceType.ListIndex = GeneralPriceType
            
            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                Dcbranch.Enabled = False
 
                DCboStoreName.Enabled = True
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
                DcboEmp.BoundText = EmpID
            Else
                Dcbranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.Dcbranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.Dcbranch.Enabled = True
                                   Else
                                    Me.Dcbranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If



            Me.Dcbranch.BoundText = Current_branch
            DBPix202.ImageClear

                                               
    
    
   'If SystemOptions.usertype <> UserAdminAll Then
  
  'GetUserData user_id, , , dstore, , , EmpID
'DcboEmp.BoundText = EmpID
'Me.DCboStoreName.BoundText = dstore
'  End If


Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
       
       NewGrid1.formname = Me.Name
    If NewGrid1.ShowCreteria() = True Then
  '  FrmShowScreenCreteria.show
  '  FG.Enabled = True
        CboStatus.ListIndex = 0
   Else
         CboStatus.ListIndex = 1
  ' FG.Enabled = False
   
   End If
    
        VSFlexGrid2.Rows = 2
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id
VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1

        Case 2
            Dim Msg  As String
  XPTab301.CurrTab = 0
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

         If val(LblDone.Caption) < val(LblAccCreitria.Caption) Then
                 If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Can't Save"
                Else
                    Msg = "·« Ì„ﬂ‰ Õ›Ÿ  Â–« «·”‰œ  ·⁄œ„  Õﬁﬁ „Õœœ«  «·‘«‘… "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         Exit Sub
         End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = GridTransType.salespricelistRequest
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄‰ ÿ·»«  ⁄—Ê÷ «·«”⁄«—"
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport
        
        Case 8
            On Error GoTo ErrTrap

            If XPTxtBillID.Text <> "" Then
                Set SaleReport = New ClsSaleReport
                SaleReport.ShowPrice XPTxtBillID.Text, 6, DcboEmp.Text, val(DBCboClientName.BoundText)
            End If

            '        PrintReport1 (Txt_order_no.text)
        Case 6
            Unload Me
     Case 9
     
            PrintReport (100)
            
     Case 10
            PrintReport (10)
            
       Case 11
            PrintReport (11)
            
               Case 12
            PrintReport (12)
            
            
            Case 21
                With Me.VSFlexGrid2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    
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
        .DcCurrency.BoundText = Me.DcCurrency.BoundText

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

Private Sub Command1_Click()

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
    End If
 
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 7
        FrmCustemerSearch.show vbModal
    End If
          
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos

        If GeneralPriceType = 0 Then
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
        ElseIf GeneralPriceType = 1 Then
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
        Else
            Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
        End If
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
        FrmItemSearch.RetrunType = 23
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
        FrmItemSearch.RetrunType = 2
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
     If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(Dcbranch.BoundText), 41) = True Then
  '   TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""

     End If
     
    End If

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
 
End If

End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches Dcbranch
    End If

End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = " select id,code from currency"
 
        fill_combo Me.DcCurrency, My_SQL

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

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), , Me.TXT_order_no

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
   '     FrmAddNewItem.DealingForm = ShowPrice
   '     FrmAddNewItem.show vbModal
    End If

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

    If TxtNoteSerial1.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
         
            MsgBox "·«»œ „‰ «Õ Ì«—  ”‰œ  «Ê·«": Exit Sub
        Else
            MsgBox "Select Voucher Firstly": Exit Sub
        End If
 
    End If

    Unload imaged
    imaged.show

    If SystemOptions.UserInterface = EnglishInterface Then

        imaged.Label9.Caption = "Po Request  #"
        imaged.Caption = "Po Requestt"
        imaged.txtopeation_type = "1010"
        imaged.SUBJECT_NO = TxtNoteSerial1.Text
        imaged.Label6.Caption = "Po Request  #"
    Else

        imaged.Label9.Caption = "  ÿ·»«  ⁄—Ê÷ «·»Ì⁄ —ﬁ„"
        imaged.Caption = "  ÿ·»«  ⁄—Ê÷  «·»Ì⁄ —ﬁ„    "
        imaged.txtopeation_type = "1010"
        imaged.SUBJECT_NO = TxtNoteSerial1.Text
        imaged.Label6.Caption = "„—›ﬁ«      ÿ·»«  ⁄—Ê÷ «·»Ì⁄ —ﬁ„"

    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type ='1010'  and subject_no='" & TxtNoteSerial1.Text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If


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

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
calPercentages
With VSFlexGrid2
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
  End With
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "value"
              
                Cancel = True
   
        End Select
End With
End Sub

 Function calPercentages()
 Dim i As Integer
 Dim Percentage As Double
  With Me.VSFlexGrid2
            For i = .FixedRows To .Rows - 1
                 Percentage = val(.TextMatrix(i, .ColIndex("percentage")))
                 If Percentage > 0 Then
                 .TextMatrix(i, .ColIndex("value")) = val(LblTotal.Caption) * Percentage / 100
                 
                 End If
                 
              Next i
              If .Rows > 1 Then
txtTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("percentage"), .Rows - 1, .ColIndex("percentage"))
End If
   End With
 End Function

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
    Dim CurrentTransactionType As Integer
'    On Error GoTo ErrTrap

    If GeneralPriceType = 0 Then
        ScreenNameArabic = "ÿ·»«  ⁄—Ê÷ «·«”⁄«— "
        ScreenNameEnglish = "Sales Quotations Reqest"
        CurrentTransactionType = 41
    ElseIf GeneralPriceType = 1 Then
        ScreenNameArabic = "«⁄ „«œ ⁄—÷ ”⁄—  "
        ScreenNameEnglish = "Sales QuotationsApproval"
        CurrentTransactionType = 42
    ElseIf GeneralPriceType = 2 Then
        ScreenNameArabic = "⁄—÷ ”⁄— ‰Â«∆Ì  "
        ScreenNameEnglish = "Final Sales Quotations "
        CurrentTransactionType = 43


    End If

    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Me.Caption = ScreenNameArabic
    Ele(6).Caption = ScreenNameArabic

With CboStatus
        .Clear
        .AddItem "„—›Ê÷"
        .AddItem "„ﬁ»Ê· "
End With

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
    NewGrid.GridTrans = salespricelistRequest
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
        Set NewGrid.DtpBillDate = Me.XPDtbBill
    'Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
      Set NewGrid1.LblDone = Me.LblDone
        Set NewGrid1.LblunDone = Me.LbunlDone
      Set NewGrid1.LblAccCreitria = Me.LblAccCreitria
     
    

    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.StoreName = DCboStoreName
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    ' Resize_Form Me, TransactionSize
    
    

    ' Set NewGrid1.Formname = Me.name
    
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FG.WallPaper = BGround.Picture
    AddTip
      
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos

    If GeneralPriceType = 0 Then
        Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    ElseIf GeneralPriceType = 1 Then
        Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
    Else
        Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    End If

    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.Dcbranch

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName

    Dcombos.GetSalesRepData Me.DcboEmp
 
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID

    NewGrid.FillGrid

    With Me.CboPriceType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "  ÿ·»«  ⁄—÷ ”⁄— "
            .AddItem "«⁄ „«œ ⁄— ÷ ”⁄—  "
            .AddItem " ⁄—÷ ”⁄— ‰Â«∆Ì"
       
        Else
            .AddItem "Sales Quotations Request"
            .AddItem "Sales Quotations  Approval"
            .AddItem "Final Sales Quotations "
 
        End If

        .ListIndex = 0
    End With

    With Me.CboType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "   ÌœÊÌ "
            .AddItem "«·Ì ÿ»ﬁ« ·Õœ «·ÿ·»«  "
     
        Else
            .AddItem "Manual"
            .AddItem "Auto "
     
        End If

        .ListIndex = 0
    End With

    Set NewGrid1.Grid = VSFlexGrid1
   
    'StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=6 or Transaction_Type=29  or Transaction_Type=17)" 'OR Transaction_Type=17
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=" & CurrentTransactionType
    StrSQL = StrSQL & "  AND      BranchId  in(" & Current_branchSql & ")"
    
  If SystemOptions.usertype <> UserAdminAll Then
 StrSQL = StrSQL & " and  Emp_id = " & user_id
End If

    StrSQL = StrSQL + " Order By Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim My_SQL As String
    My_SQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, My_SQL
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

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.Dcbranch.Enabled = True
        DcboEmp.Enabled = False
        TxtEmployeeID.Enabled = False
        DCboStoreName.Enabled = True
         TxtStoreID.Enabled = False
    End If
    
    
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
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ   " & TXT_order_no.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & "«‰Ê⁄ «·”‰œ  " & CboPriceType.Text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.Text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & CHR(13) & " —ﬁ„ «·«⁄ „«œ    " & TxtLcNo
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr . No   " & TXT_order_no.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Type  " & CboPriceType.Text & CHR(13) & " Store  " & DCboStoreName.Text & CHR(13) & " Customer/ Supplier " & DBCboClientName.Text & CHR(13) & " Lc NO    " & TxtLcNo
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , , Me.TXT_order_no
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , , Me.TXT_order_no
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
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TXT_order_no.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)

    If rs("shipped").value = True Then
        chkshipped.value = vbChecked
    Else
        chkshipped.value = Unchecked
    End If

    Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").value), "", rs("countryid").value)

txtDays1.Text = IIf(IsNull(rs("Days1").value), 0, (rs("Days1").value))
txtDays2.Text = IIf(IsNull(rs("Days2").value), 0, (rs("Days2").value))

    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
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

    XPTxtTaxValue.Text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    TxtLcNo.Text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
 TxtBillComment.Text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
 
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    If TXT_order_no <> "" Then
        Me.TxtNoteSerial1.Text = TXT_order_no
    End If

    'Txt_order_no

    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

   ' DBPix202.ImageClear

   ' If Dir(App.path & "\images\sign\sign" & rs("posted").value & ".JPG") <> "" Then
   '
   '     DBPix202.ImageLoadFile (App.path & "\images\sign\sign" & user_id & ".JPG")
   ' End If

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
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
        
       'GURANTEEE
              FG.TextMatrix(Num, FG.ColIndex("GranteeType")) = IIf(IsNull(RsDetails("GranteeType")), "", (RsDetails("GranteeType").value))
            FG.TextMatrix(Num, FG.ColIndex("GranteeStartDate")) = IIf(IsNull(RsDetails("GranteeStartDate")), "", (RsDetails("GranteeStartDate").value))
            FG.TextMatrix(Num, FG.ColIndex("GranteeEndDate")) = IIf(IsNull(RsDetails("GranteeEndDate")), "", (RsDetails("GranteeEndDate").value))
            FG.TextMatrix(Num, FG.ColIndex("RegularMaintenancedates")) = IIf(IsNull(RsDetails("RegularMaintenancedates")), "", (RsDetails("RegularMaintenancedates").value))
       FG.TextMatrix(Num, FG.ColIndex("RegularMaintenanceIDS")) = IIf(IsNull(RsDetails("RegularMaintenanceIDS")), "", (RsDetails("RegularMaintenanceIDS").value))
       
       
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

With VSFlexGrid1
     .Clear flexClearScrollable, flexClearEverything
     .Rows = 1
     .Clear flexClearScrollable, flexClearEverything
     .Refresh
 
Dim RsTransactionsCreteria As New ADODB.Recordset
 

    StrSQL = " SELECT     dbo.TransactionsCreteria.CreteriaID, dbo.tblScreenCriteria.Name, dbo.tblScreenCriteria.Namee, dbo.tblScreenCriteria.typeid, dbo.tblScreenCriteria.[value], "
    StrSQL = StrSQL + " dbo.tblScreenCriteria.Remarks, dbo.TransactionsCreteria.Enteredvalue, dbo.TransactionsCreteria.doneid, dbo.TransactionsCreteria.remarks AS userRemarks"
    StrSQL = StrSQL + ", dbo.TransactionsCreteria.remarksE  FROM         dbo.tblScreenCriteria RIGHT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TransactionsCreteria ON dbo.tblScreenCriteria.CriteriaID = dbo.TransactionsCreteria.CreteriaID"
    StrSQL = StrSQL + " WHERE     (dbo.TransactionsCreteria.Transaction_ID = " & val(rs("Transaction_ID").value) & ") "
     
  RsTransactionsCreteria.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   

    If Not (RsTransactionsCreteria.EOF Or RsTransactionsCreteria.BOF) Then
         .Rows = RsTransactionsCreteria.RecordCount + 1

        For Num = 1 To RsTransactionsCreteria.RecordCount
             .TextMatrix(Num, .ColIndex("PlainMessageID")) = IIf(IsNull(RsTransactionsCreteria("CreteriaID")), "", (RsTransactionsCreteria("CreteriaID").value))
          If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(RsTransactionsCreteria("Name")), "", (RsTransactionsCreteria("Name").value))
           Else
           .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(RsTransactionsCreteria("Namee")), "", (RsTransactionsCreteria("Namee").value))
           End If
           
           .TextMatrix(Num, .ColIndex("typeid")) = IIf(IsNull(RsTransactionsCreteria("typeid")), "", (RsTransactionsCreteria("typeid").value))
           .TextMatrix(Num, .ColIndex("typeName")) = ShowOperator(val(.TextMatrix(Num, .ColIndex("typeid"))))
            
           
           .TextMatrix(Num, .ColIndex("value")) = IIf(IsNull(RsTransactionsCreteria("value")), "", (RsTransactionsCreteria("value").value))
           .TextMatrix(Num, .ColIndex("Enteredvalue")) = IIf(IsNull(RsTransactionsCreteria("Enteredvalue")), "", (RsTransactionsCreteria("Enteredvalue").value))
           .TextMatrix(Num, .ColIndex("doneid")) = IIf(IsNull(RsTransactionsCreteria("doneid")), "", (RsTransactionsCreteria("doneid").value))
          If SystemOptions.UserInterface = ArabicInterface Then
                                  If val(.TextMatrix(Num, .ColIndex("doneid"))) = 0 Then
                                  .TextMatrix(Num, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                            Else
                            .TextMatrix(Num, .ColIndex("Done")) = "„Õﬁﬁ"
                            End If
           Else
                            If val(.TextMatrix(Num, .ColIndex("doneid"))) = 0 Then
                                  .TextMatrix(Num, .ColIndex("Done")) = " Denied"
                            Else
                            .TextMatrix(Num, .ColIndex("Done")) = "Dene"
                            End If
          
           End If
           
          .TextMatrix(Num, .ColIndex("remarks")) = IIf(IsNull(RsTransactionsCreteria("userRemarks")), "", (RsTransactionsCreteria("userRemarks").value))
.TextMatrix(Num, .ColIndex("remarksE")) = IIf(IsNull(RsTransactionsCreteria("remarksE")), "", (RsTransactionsCreteria("remarksE").value))

           

          RsTransactionsCreteria.MoveNext
         

        Next Num

    End If



End With

'«·œ›⁄« 
With VSFlexGrid2
     .Clear flexClearScrollable, flexClearEverything
     .Rows = 1
     .Clear flexClearScrollable, flexClearEverything
     .Refresh
 
Dim RsTransactionsvalues As New ADODB.Recordset

    StrSQL = " SELECT    *  from TransactionsValues "
    StrSQL = StrSQL + " WHERE     (dbo.TransactionsValues.Transaction_ID = " & val(rs("Transaction_ID").value) & ") "
     
  RsTransactionsvalues.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   

    If Not (RsTransactionsvalues.EOF Or RsTransactionsvalues.BOF) Then
         .Rows = RsTransactionsvalues.RecordCount + 1

        For Num = 1 To RsTransactionsvalues.RecordCount
             .TextMatrix(Num, .ColIndex("percentage")) = IIf(IsNull(RsTransactionsvalues("percentage")), "", (RsTransactionsvalues("percentage").value))
 
           
           .TextMatrix(Num, .ColIndex("value")) = IIf(IsNull(RsTransactionsvalues("value")), "", (RsTransactionsvalues("value").value))
           .TextMatrix(Num, .ColIndex("Des")) = IIf(IsNull(RsTransactionsvalues("value")), "", (RsTransactionsvalues("Des").value))
 
          RsTransactionsvalues.MoveNext
         

        Next Num

    End If



End With
calPercentages

NewGrid1.checkDataCreteria

fillapprovData
    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Function SaveGranteeData()
    Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDate As String
    Dim Alliids As String
    Dim RowNum As Integer
    strFilterText = ","
    Dim astrSplitItems1() As String
    
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete TBLRegularMaint   where Transaction_ID= " & val(Me.XPTxtBillID.Text)
    RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")) <> "" Then
                AllDate = FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates"))
                Alliids = FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenanceIDS"))
                
                astrSplitItems = Split(AllDate, strFilterText)
         astrSplitItems1 = Split(Alliids, strFilterText)
         
                For intX = 0 To UBound(astrSplitItems)
                        
                    If IsDate(astrSplitItems(intX)) Then
                        RsgGrantee.AddNew
                        RsgGrantee("DateOfRegularMaint").value = Format$(astrSplitItems(intX), "dd/mm/yyyy")
                        RsgGrantee("MaintenanceIDS").value = astrSplitItems1(intX)
                        
                        RsgGrantee("itemid").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                        RsgGrantee("GranteeType").value = val(FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")))
                        RsgGrantee("GranteeStartDate").value = FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate"))
                        RsgGrantee("GranteeEndDate").value = FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate"))
                        RsgGrantee("ItemSerial").value = FG.TextMatrix(RowNum, FG.ColIndex("Serial"))
                       RsgGrantee("Count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
       
                        RsgGrantee("Done").value = 0
                        RsgGrantee.update
                    End If
                       
                Next intX
                    
            End If

        End If

    Next RowNum

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
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
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

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

Cn.Execute " delete TBLRegularMaint where Transaction_ID=" & val(XPTxtBillID.Text)
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
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
     On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.Text <> "R" Then
        If DBCboClientName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
            Else
                Msg = "Please Select Vendor"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If DcCurrency.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ «·⁄„·…"
            Else
                Msg = "Select Currency"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcCurrency.SetFocus
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

        If XPChkTAX.value = Checked Then
            If XPTxtTaxValue.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
                Else
                    Msg = "Insert Sales Tax"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                XPTxtTaxValue.SetFocus
                FG.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
calPercentages

 If txtTotal.Text <> 0 And txtTotal.Text < 100 Then
     If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« Ì„ﬂ‰ «·Õ›Ÿ ·«‰ ‰”»… «·œ›⁄« "
                Else
                    Msg = "can't save  percentage not done"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 
Exit Sub
 End If
 
 
 
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        Set RSTransDetails = New ADODB.Recordset
        'RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
  
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CboPriceType.ListIndex = 0 Then
            Transaction_Type = 41
            Sanad_No = 41
        ElseIf Me.CboPriceType.ListIndex = 1 Then
            Transaction_Type = 42
            Sanad_No = 42
        ElseIf Me.CboPriceType.ListIndex = 2 Then
            Transaction_Type = 43
            Sanad_No = 43
 
         
        End If

        my_branch = val(Dcbranch.BoundText)

Dim TxtNoteSerial1str As String

        If TxtNoteSerial1.Text = "" Then
        TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText))
        
            If TxtNoteSerial1str = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If TxtNoteSerial1str = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ    " & CHR(13) & " Enter Vchr No": Exit Sub
                Else
                    TxtNoteSerial1.Text = TxtNoteSerial1str
                End If
            End If
        End If
 
        TXT_order_no = Me.TxtNoteSerial1.Text
 
        Cn.BeginTrans
        BeginTrans = True
    
        If Me.TxtModFlg.Text = "N" Then
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
                 XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=6"))
            
            rs.AddNew
        End If

        Screen.MousePointer = vbArrowHourglass
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rs("branchID").value = val(Me.Dcbranch.BoundText)
   
        rs("Transaction_ID").value = val(XPTxtBillID.Text)
        rs("order_no").value = TXT_order_no.Text
    
        If chkshipped.value = vbChecked Then
            rs("shipped").value = 1
        Else
            rs("shipped").value = 0
        End If
    
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Serial").value = TxtTransSerial.Text

        If Me.CboPriceType.ListIndex = 0 Then
            rs("Transaction_Type").value = 41
        ElseIf Me.CboPriceType.ListIndex = 1 Then
            rs("Transaction_Type").value = 42
        ElseIf Me.CboPriceType.ListIndex = 2 Then
            rs("Transaction_Type").value = 43
  
        End If
 rs("Days1").value = val(txtDays1.Text)
  rs("Days2").value = val(txtDays2.Text)
        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("countryid").value = IIf(DataCombo4.BoundText = "", Null, val(DataCombo4.BoundText))
    
        rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
        rs("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
        rs("total").value = IIf(XPTxtSum.Text = "", Null, val(XPTxtSum.Text))
        rs("LcNo").value = IIf(TxtLcNo.Text = "", Null, (TxtLcNo.Text))
        rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))
        rs.update
    
        CuurentLogdata
  
        If Me.TxtModFlg.Text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
                  StrSqlDel = "delete From TransactionsCreteria where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
                        StrSqlDel = "delete From TransactionsValues where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
        End If

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
                RSTransDetails("order_id").value = val(XPTxtBillID.Text)
             
                RSTransDetails("order_no").value = TXT_order_no.Text
             
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
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                End If
                
                           RSTransDetails("GranteeType").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")))
            RSTransDetails("GranteeStartDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate"))), "DD/mm/YYYY"))
            RSTransDetails("GranteeEndDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate"))), "DD/mm/YYYY"))
            RSTransDetails("RegularMaintenancedates").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")))
            
        RSTransDetails("RegularMaintenanceIDS").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenanceIDS")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenanceIDS")))



                RSTransDetails.update
            End If

        Next RowNum



   Dim RsTransactionsCreteria As ADODB.Recordset
     Set RsTransactionsCreteria = New ADODB.Recordset
        RsTransactionsCreteria.Open "[TransactionsCreteria]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
With VSFlexGrid1
    For RowNum = 1 To .Rows - 1

            If .TextMatrix(RowNum, .ColIndex("PlainMessageID")) <> "" Then
                RsTransactionsCreteria.AddNew
                RsTransactionsCreteria("Transaction_ID").value = val(XPTxtBillID.Text)
                 RsTransactionsCreteria("CreteriaID").value = val(.TextMatrix(RowNum, .ColIndex("PlainMessageID")))
                 RsTransactionsCreteria("Enteredvalue").value = val(.TextMatrix(RowNum, .ColIndex("Enteredvalue")))
                  RsTransactionsCreteria("doneid").value = val(.TextMatrix(RowNum, .ColIndex("doneid")))
                 RsTransactionsCreteria("remarks").value = (.TextMatrix(RowNum, .ColIndex("remarks")))
                 RsTransactionsCreteria("remarksE").value = (.TextMatrix(RowNum, .ColIndex("remarkse")))
                 
                RsTransactionsCreteria.update
             End If
    Next RowNum
  End With
  
  RsTransactionsCreteria.Close
  Set RsTransactionsCreteria = Nothing
  
  
     Dim RsTransactionsvalues As ADODB.Recordset
     Set RsTransactionsvalues = New ADODB.Recordset
        RsTransactionsvalues.Open "[TransactionsValues]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
With VSFlexGrid2
    For RowNum = 1 To .Rows - 1

            If val(.TextMatrix(RowNum, .ColIndex("percentage"))) > 0 Then
                RsTransactionsvalues.AddNew
                RsTransactionsvalues("Transaction_ID").value = val(XPTxtBillID.Text)
                 RsTransactionsvalues("percentage").value = val(.TextMatrix(RowNum, .ColIndex("percentage")))
                 RsTransactionsvalues("value").value = val(.TextMatrix(RowNum, .ColIndex("value")))
                   
                 RsTransactionsvalues("Des").value = (.TextMatrix(RowNum, .ColIndex("Des")))
                RsTransactionsvalues.update
             End If
    Next RowNum
  End With
  
  RsTransactionsvalues.Close
  Set RsTransactionsvalues = Nothing
  
  
  

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        
        lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
    SaveGranteeData
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & CHR(13)
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

        TxtModFlg.Text = "R"
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

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry... Error During Saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

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

Private Sub PrintReport(Optional reportid As Integer = 0)
    On Error GoTo ErrTrap
Dim RowNum As Integer
Dim stringItems As String
Dim stringcashing As String
Dim ItemName As String
    If XPTxtBillID.Text <> "" Then
        
If reportid = 0 Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 93, DcboEmp.Text
        
 ElseIf reportid = 10 Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 10, DcboEmp.Text

 ElseIf reportid = 11 Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 11, DcboEmp.Text
        
    
 ElseIf reportid = 12 Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 12, DcboEmp.Text
         
ElseIf reportid = 100 Then
Set SaleReport = New ClsSaleReport

'        SaleReport.ShowPrice XPTxtBillID.text, 100, DcboEmp.text


With FG
  For RowNum = 1 To .Rows - 1

    
        If .TextMatrix(RowNum, .ColIndex("Code")) <> "" Then
        
      GetItemData val(.TextMatrix(RowNum, .ColIndex("Code"))), , ItemName
        stringItems = stringItems & _
        ")" & val(.TextMatrix(RowNum, .ColIndex("Count"))) & "(" & "    " & _
          ItemName & CHR(13)
        
End If

Next RowNum

End With


With VSFlexGrid2
              For RowNum = 1 To .Rows - 1
            
                
                    If .TextMatrix(RowNum, .ColIndex("percentage")) <> "" Then
                    
                   
                    stringcashing = stringcashing & _
                    "" & val(.TextMatrix(RowNum, .ColIndex("percentage"))) & "%" & "            " & _
                        (.TextMatrix(RowNum, .ColIndex("des"))) & CHR(13)
                    
                    End If
            
            Next RowNum

End With



        SaleReport.ShowPrice XPTxtBillID.Text, 100, DcboEmp.Text, , stringItems, stringcashing, LblTotal.Caption
    
End If
    
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

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial1.Text = ""
 
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
    Me.Caption = ScreenNameEnglish
    Ele(6).Caption = ScreenNameEnglish
    'Me.Caption = "Order Request/Proforma   Invoice"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Internal Orders"
    lbl(18).Caption = "Type"
    Label4.Caption = "ACC. BY"
    Label10.Caption = "Signature"
    lbl(32).Caption = "Sales Person"
    Accredit.Caption = "Accredit"
    Cmd(8).Caption = "Print Pur. Order"
    'Ele(6).Caption = Me.Caption
    lbl(50).Caption = "Discounts"
    lbl(49).Caption = "Net"

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
    Label1.Caption = "LC NO:"
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

With CboStatus
.Clear
.AddItem "Denied"
.AddItem "Accept"
End With
Cmd(9).Caption = "Print"
lbl(34).Caption = "Deliveray "
lbl(36).Caption = "Day"

lbl(35).Caption = "Installation "
 lbl(37).Caption = "Day"
 

    With Me.VSFlexGrid2
 
        .TextMatrix(0, .ColIndex("percentage")) = "percentage"
         .TextMatrix(0, .ColIndex("value")) = "value"
'
         .TextMatrix(0, .ColIndex("Des")) = "Des"
 
    End With
 
 
    With Me.VSFlexGrid1
 
         .TextMatrix(0, .ColIndex("PlainMessageID")) = "Code"
         .TextMatrix(0, .ColIndex("TypeName")) = "Name"
         .TextMatrix(0, .ColIndex("value")) = "Default value"
 .TextMatrix(0, .ColIndex("Enteredvalue")) = "Entered value"
.TextMatrix(0, .ColIndex("done")) = "done"
 
         .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
 
    End With
 
 
End Sub
