VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCompany 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   HelpContextID   =   60
   Icon            =   "FrmCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   10800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Height          =   732
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   8760
      Width           =   10572
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9165
         TabIndex        =   83
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   8445
         TabIndex        =   84
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   7725
         TabIndex        =   85
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   86
         Top             =   240
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   6105
         TabIndex        =   87
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   5190
         TabIndex        =   88
         Top             =   240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   89
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   4380
         TabIndex        =   90
         Top             =   240
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   3510
         TabIndex        =   91
         Top             =   240
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   11
         Left            =   1080
         TabIndex        =   92
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   12
         Left            =   2160
         TabIndex        =   155
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "«·›Ê« Ì— «·”«»ﬁ…"
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin VB.TextBox txtcode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3405
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox XPTxtComID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3420
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1545
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   588
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10752
      _cx             =   18971
      _cy             =   1032
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
      Caption         =   " »Ì«‰«  «·„Ê—œÌ‰ "
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
         Left            =   1185
         TabIndex        =   3
         Top             =   120
         Width           =   495
         _ExtentX        =   873
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
         ButtonImage     =   "FrmCompany.frx":038A
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   873
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
         ButtonImage     =   "FrmCompany.frx":0724
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
         Left            =   1710
         TabIndex        =   2
         Top             =   120
         Width           =   495
         _ExtentX        =   873
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
         ButtonImage     =   "FrmCompany.frx":0ABE
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
         Left            =   645
         TabIndex        =   4
         Top             =   120
         Width           =   495
         _ExtentX        =   873
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
         ButtonImage     =   "FrmCompany.frx":0E58
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
         Left            =   6120
         Picture         =   "FrmCompany.frx":11F2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Image Img 
         Height          =   480
         Left            =   2280
         Picture         =   "FrmCompany.frx":4E5A
         Top             =   0
         Width           =   480
      End
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7812
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   10788
      _cx             =   19029
      _cy             =   13779
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
      Caption         =   "»Ì«‰«  «”«”Ì…|”Ì«—«  «·„ ⁄«ÂœÌ‰|»Ì«‰«  „ Œ’’…"
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
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬁ— «·„Ê—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   7395
         Index           =   5
         Left            =   11730
         RightToLeft     =   -1  'True
         TabIndex        =   187
         Top             =   45
         Width           =   10695
         Begin VB.TextBox txtPostalCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   315
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   2400
            Width           =   5445
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   5340
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   3180
            Width           =   1005
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   5340
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   3570
            Width           =   1005
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   8055
            MaxLength       =   4
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   3570
            Width           =   1005
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   315
            Index           =   10
            Left            =   8055
            MaxLength       =   2
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   3210
            Width           =   1005
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   315
            Index           =   2
            Left            =   8055
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   2760
            Width           =   1005
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   315
            Index           =   4
            Left            =   5340
            MaxLength       =   4
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Tag             =   "4 digit at least"
            Top             =   2790
            Width           =   1005
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   3600
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   2010
            Width           =   5445
         End
         Begin VB.TextBox TxtAddress 
            Alignment       =   1  'Right Justify
            Height          =   525
            Left            =   3600
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   188
            Top             =   1440
            Width           =   5445
         End
         Begin MSDataListLib.DataCombo DcboCountryID 
            Height          =   315
            Left            =   3600
            TabIndex        =   189
            Top             =   450
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboGovernmentID 
            Height          =   315
            Left            =   3600
            TabIndex        =   190
            Top             =   780
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboCityID 
            Height          =   315
            Left            =   3600
            TabIndex        =   191
            Top             =   1110
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—„“ «·»—ÌœÏ*"
            Height          =   285
            Index           =   42
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   2400
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œÌ‰… «·›—⁄Ì…"
            Height          =   375
            Index           =   89
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   209
            Top             =   3210
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘«—⁄2"
            Height          =   375
            Index           =   88
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   3660
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„Œÿÿ"
            Height          =   375
            Index           =   87
            Left            =   9390
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   3570
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·œÊ·…*"
            Height          =   255
            Index           =   86
            Left            =   9645
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   3210
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘«—⁄*"
            Height          =   375
            Index           =   90
            Left            =   9405
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   2820
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„»‰Ï*"
            Height          =   255
            Index           =   91
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   2790
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ 700"
            Height          =   375
            Index           =   85
            Left            =   9405
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   2040
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·œÊ·…"
            Height          =   225
            Index           =   22
            Left            =   9630
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   510
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Õ«›Ÿ…"
            Height          =   225
            Index           =   24
            Left            =   9630
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   810
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œÌ‰…"
            Height          =   225
            Index           =   25
            Left            =   9630
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   1140
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄‰Ê«‰ »«· ›’Ì·"
            Height          =   225
            Index           =   26
            Left            =   9270
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   1680
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7395
         Index           =   2
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   45
         Width           =   10695
         _cx             =   18865
         _cy             =   13044
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
         Begin VB.CheckBox chkTaxExempt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄›Ï „‰ «·÷—Ì»…"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2190
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   1170
            Width           =   1695
         End
         Begin VB.TextBox TXTTOPerson 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3360
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   6960
            Width           =   1785
         End
         Begin VB.CheckBox CHkMot3ahed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ ⁄ÂœÌ‰"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   360
            Width           =   1695
         End
         Begin VB.Frame Frame2 
            Height          =   1455
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   2940
            Visible         =   0   'False
            Width           =   5415
            Begin VB.TextBox XPMTxtRemarks2 
               Alignment       =   1  'Right Justify
               Height          =   795
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   480
               Width           =   5145
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "”»» «·«Ìﬁ«›"
               Height          =   285
               Index           =   32
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   240
               Width           =   825
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.TextBox TxtVATNO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   5880
            Width           =   2985
         End
         Begin VB.TextBox TxtBankAddress 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   50
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   153
            Top             =   6960
            Width           =   1905
         End
         Begin VB.TextBox TxtBankIBAN 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   6600
            Width           =   1905
         End
         Begin VB.ComboBox cbCategory 
            Height          =   315
            Left            =   2430
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   900
            Width           =   1455
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·‘Œ’ «·„”ƒÊ·"
            Height          =   975
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   1320
            Width           =   6012
            Begin VB.TextBox txtDegree 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   600
               Width           =   4785
            End
            Begin VB.TextBox txtRSID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   240
               Width           =   2025
            End
            Begin MSComCtl2.DTPicker dtpRsIDDate 
               Height          =   330
               Left            =   120
               TabIndex        =   137
               Top             =   -240
               Visible         =   0   'False
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   582
               _Version        =   393216
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   182452227
               CurrentDate     =   38718
            End
            Begin Dynamic_Byte.NourHijriCal dtpRsIDDate1 
               Height          =   255
               Left            =   120
               TabIndex        =   140
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ ’œÊ— «·ÂÊÌ…"
               Height          =   315
               Index           =   35
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   240
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ—Ã… «·ÊŸÌ›Ì…"
               Height          =   285
               Index           =   34
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   630
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÂÊÌ…"
               Height          =   285
               Index           =   18
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   270
               Width           =   915
            End
         End
         Begin VB.TextBox txtRecordNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8160
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   960
            Width           =   1125
         End
         Begin VB.TextBox txtBankAccount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   6240
            Width           =   1905
         End
         Begin VB.TextBox XPTxtComName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5376
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   576
            Width           =   3915
         End
         Begin VB.TextBox XPMTxtRemark 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   6240
            Width           =   2985
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ«·Õ«·Ï"
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   2
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   6840
            Width           =   4335
            Begin ImpulseButton.ISButton Cmd 
               Height          =   435
               Index           =   8
               Left            =   120
               TabIndex        =   66
               Top             =   150
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   767
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷  ﬁ—Ì— ﬂ‘› Õ”«»"
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
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   315
               Index           =   9
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   255
               Index           =   8
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   240
               Width           =   2145
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·√ ’«·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   2685
            Index           =   3
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1230
            Width           =   4245
            Begin VB.TextBox txtBoxNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   1920
               Width           =   2925
            End
            Begin VB.TextBox XPTxtPhone 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   990
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   570
               Width           =   2025
            End
            Begin VB.TextBox XPTxtMobile 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   990
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   915
               Width           =   2025
            End
            Begin VB.TextBox TxtE_mail 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1590
               Width           =   2925
            End
            Begin VB.TextBox TxtFaxNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   990
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   1260
               Width           =   2025
            End
            Begin VB.TextBox TxtResponsibleContact 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   96
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   210
               Width           =   2928
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’ .»"
               Height          =   288
               Index           =   13
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   1920
               Width           =   912
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÃÊ«·"
               Height          =   285
               Index           =   2
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   945
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Â« ›"
               Height          =   285
               Index           =   3
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   600
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ï"
               Height          =   285
               Index           =   12
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   1590
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·›«ﬂ”"
               Height          =   285
               Index           =   7
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1290
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”∆Ê· «·≈ ’«·"
               Height          =   315
               Index           =   23
               Left            =   2940
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
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
            Height          =   1965
            Index           =   1
            Left            =   96
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   3960
            Width           =   6012
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ ··„Ê—œ"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   975
               Index           =   0
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   840
               Width           =   5835
               Begin VB.CheckBox chkIsBranch 
                  Caption         =   "H"
                  Height          =   225
                  Index           =   4
                  Left            =   120
                  TabIndex        =   212
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   2910
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   4710
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   3510
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   1890
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   210
                  Width           =   915
               End
               Begin MSComCtl2.DTPicker Dtp 
                  Height          =   330
                  Left            =   150
                  TabIndex        =   45
                  Top             =   540
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   182452227
                  CurrentDate     =   38718
               End
               Begin VB.Label c 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   345
                  Index           =   5
                  Left            =   4530
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   510
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   315
                  Index           =   6
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   540
                  Width           =   1125
               End
            End
            Begin VB.TextBox TxtCreditLimit 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2910
               MaxLength       =   8
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   180
               Width           =   1185
            End
            Begin VB.TextBox TxtCreditlimitCredit 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2910
               MaxLength       =   8
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   540
               Width           =   1185
            End
            Begin VB.ComboBox dcCreditIntervalID 
               Height          =   288
               ItemData        =   "FrmCompany.frx":5B24
               Left            =   120
               List            =   "FrmCompany.frx":5B26
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   600
               Width           =   975
            End
            Begin VB.ComboBox dcDepitIntervalID 
               Height          =   288
               ItemData        =   "FrmCompany.frx":5B28
               Left            =   120
               List            =   "FrmCompany.frx":5B2A
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox TxtCreditInterval 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   600
               Width           =   495
            End
            Begin VB.TextBox TxtDepitInterval 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õœ «·√∆ „«‰(„œÌ‰)"
               Height          =   285
               Index           =   10
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   210
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õœ «·√∆ „«‰(œ«∆‰)"
               Height          =   285
               Index           =   11
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   570
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÂ «·«∆ „«‰"
               Height          =   285
               Index           =   31
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   600
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÂ «·«∆ „«‰"
               Height          =   285
               Index           =   30
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label c 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   345
               Index           =   0
               Left            =   -4680
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1080
               Width           =   885
            End
            Begin VB.Label c 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   345
               Index           =   1
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1080
               Width           =   885
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„«  Œ«’… ··⁄„Ì· ›Ï ›Ê« Ì— «·»Ì⁄"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1035
            Index           =   4
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   2310
            Width           =   5925
            Begin VB.TextBox TxtDiscountValue 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   3420
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   630
               Width           =   1425
            End
            Begin VB.ComboBox CboDiscountType 
               Height          =   288
               Left            =   3390
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   270
               Width           =   1455
            End
            Begin VB.CheckBox locked 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «· ⁄«„·"
               Height          =   255
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   120
               Width           =   1455
            End
            Begin ALLButtonS.ALLButton ALLButton1 
               Height          =   375
               Left            =   240
               TabIndex        =   27
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«·”»»"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmCompany.frx":5B2C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSDataListLib.DataCombo DcCustomerType 
               Height          =   315
               Left            =   0
               TabIndex        =   28
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   600
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
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
               Height          =   255
               Index           =   21
               Left            =   3180
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   720
               Width           =   195
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·Œ’„"
               Height          =   285
               Index           =   20
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   690
               Width           =   825
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Œ’„"
               Height          =   285
               Index           =   19
               Left            =   4770
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„Ê—œ"
               Height          =   285
               Index           =   2
               Left            =   1050
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   360
               Width           =   1890
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„«  Œ«’… ··„Ê—œ ›Ï ›Ê« Ì— «·‘—«¡"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1005
            Index           =   6
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   3360
            Width           =   5895
            Begin VB.ComboBox CboDiscountTypePur 
               Height          =   288
               Left            =   3390
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TxtDiscountValuePur 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   900
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Œ’„"
               Height          =   285
               Index           =   29
               Left            =   4770
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   270
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·Œ’„"
               Height          =   285
               Index           =   28
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   300
               Width           =   825
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
               Height          =   255
               Index           =   27
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   330
               Width           =   195
            End
         End
         Begin VB.TextBox XPTxtCusNamee 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   240
            MaxLength       =   50
            TabIndex        =   12
            Top             =   540
            Width           =   3645
         End
         Begin VB.TextBox txtopening_balance_voucher_id 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   780
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtid 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8160
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   180
            Width           =   1125
         End
         Begin VB.CheckBox chkCustomerandVendor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ì· Ê„Ê—œ"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   60
            Width           =   1695
         End
         Begin ImpulseButton.ISButton CmdPriceList 
            Height          =   255
            Left            =   930
            TabIndex        =   72
            Top             =   7050
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   450
            ButtonPositionImage=   1
            Caption         =   "ﬁ«∆„… √”⁄«— «·„Ê—œ"
            BackColor       =   14737632
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmCompany.frx":5B48
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   255
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DboParentAccount 
            Height          =   315
            Left            =   90
            TabIndex        =   73
            Top             =   5940
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCPreFix 
            Height          =   312
            Left            =   6840
            TabIndex        =   74
            Top             =   180
            Width           =   1152
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   312
            Left            =   240
            TabIndex        =   75
            Top             =   180
            Width           =   3648
            _ExtentX        =   6429
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal dtpRsIDDateH 
            Height          =   255
            Left            =   5400
            TabIndex        =   139
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo DcbCurrency 
            Height          =   315
            Left            =   240
            TabIndex        =   160
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCEmP 
            Height          =   315
            Left            =   6360
            TabIndex        =   179
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   6600
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo txtBankName 
            Height          =   315
            Left            =   3360
            TabIndex        =   184
            Top             =   6240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo TxtBankCode 
            Height          =   315
            Left            =   3360
            TabIndex        =   185
            Top             =   6600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„” ›Ìœ"
            Height          =   315
            Index           =   41
            Left            =   5310
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   6990
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰œÊ»"
            Height          =   285
            Index           =   1
            Left            =   9255
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   6600
            Width           =   1170
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„·…"
            Height          =   255
            Index           =   14
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
            Height          =   345
            Index           =   40
            Left            =   9330
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   5880
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄‰Ê«‰ «·»‰ﬂ"
            Height          =   285
            Index           =   38
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   6990
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—„“ «·»‰ﬂ"
            Height          =   315
            Index           =   37
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   6630
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«Ì»«‰"
            Height          =   285
            Index           =   36
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   6630
            Width           =   1275
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›∆…"
            Height          =   252
            Index           =   13
            Left            =   3852
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   960
            Width           =   1152
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”Ã·"
            Height          =   315
            Index           =   17
            Left            =   6780
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   312
            Index           =   1
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   960
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·»‰ﬂ"
            Height          =   285
            Index           =   16
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   6270
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»‰ﬂ"
            Height          =   315
            Index           =   15
            Left            =   4950
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   6240
            Width           =   1365
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„"
            Height          =   312
            Index           =   0
            Left            =   9336
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   588
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   165
            Index           =   1
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   6240
            Width           =   1335
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   312
            Index           =   2
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   300
            Width           =   1188
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «·«‰Ã·Ì“Ì"
            Height          =   252
            Index           =   4
            Left            =   3852
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   660
            Width           =   1152
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«» «·—∆Ì”Ì"
            Height          =   315
            Index           =   33
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   5940
            Width           =   1365
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4305
            TabIndex        =   76
            Top             =   180
            Width           =   690
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frame1 
         Height          =   7395
         Left            =   11430
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   45
         Width           =   10695
         _cx             =   18865
         _cy             =   13044
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
         Caption         =   "”Ì«—«  «·„ ⁄«ÂœÌ‰"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   1935
            Left            =   0
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   240
            Width           =   10695
            _cx             =   18865
            _cy             =   3413
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
            Begin VB.ComboBox DcbTypTrans1 
               Height          =   315
               ItemData        =   "FrmCompany.frx":5EE2
               Left            =   0
               List            =   "FrmCompany.frx":5EE4
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   0
               Visible         =   0   'False
               Width           =   3252
            End
            Begin VB.ComboBox DcbTypTrans 
               Height          =   315
               ItemData        =   "FrmCompany.frx":5EE6
               Left            =   6360
               List            =   "FrmCompany.frx":5EE8
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   1200
               Width           =   3252
            End
            Begin VB.TextBox TxtPartPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   1560
               Width           =   1128
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   1200
               Width           =   1128
            End
            Begin VB.TextBox accessoryTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   1560
               Width           =   3252
            End
            Begin VB.TextBox txtDriverTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   480
               Width           =   1128
            End
            Begin VB.TextBox txtDriverName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   480
               Width           =   3252
            End
            Begin VB.ComboBox DcCity 
               Height          =   315
               ItemData        =   "FrmCompany.frx":5EEA
               Left            =   6360
               List            =   "FrmCompany.frx":5EEC
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   960
               Visible         =   0   'False
               Width           =   3252
            End
            Begin VB.TextBox txtChassis 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   120
               Width           =   1092
            End
            Begin VB.ComboBox cbModel 
               Height          =   315
               ItemData        =   "FrmCompany.frx":5EEE
               Left            =   6360
               List            =   "FrmCompany.frx":5EF0
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   112
               Top             =   840
               Width           =   3252
            End
            Begin VB.TextBox txtCount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   120
               Width           =   1128
            End
            Begin C1SizerLibCtl.C1Elastic Frame6 
               Height          =   1455
               Left            =   0
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   0
               Width           =   2655
               _cx             =   4683
               _cy             =   2566
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
               Caption         =   "—ﬁ„ «··ÊÕ…"
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Height          =   492
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   840
                  Width           =   2172
                  Begin VB.TextBox ntxtNum4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   0
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtLetter4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1080
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   174
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtNum3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   240
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   173
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtNum2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   480
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   172
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtNum1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   720
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   171
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtLetter3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1320
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   170
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtLetter2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1560
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   169
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox ntxtLetter1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1800
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   168
                     Top             =   120
                     Width           =   288
                  End
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Height          =   492
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   360
                  Width           =   2172
                  Begin VB.TextBox txtLetter1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1800
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtLetter2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1560
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtLetter3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1320
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtNum1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   720
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   105
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtNum2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   480
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   106
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtNum3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   240
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtLetter4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1080
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   120
                     Width           =   288
                  End
                  Begin VB.TextBox txtNum4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   0
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   120
                     Width           =   288
                  End
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ » Ã  1 2 3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   192
                  Index           =   10
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   120
                  Width           =   1188
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„À«· "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   312
                  Index           =   3
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   120
                  Width           =   468
               End
            End
            Begin MSDataListLib.DataCombo dcBrand 
               Height          =   288
               Left            =   6360
               TabIndex        =   115
               Top             =   120
               Width           =   1212
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   14
               Left            =   2880
               TabIndex        =   116
               Top             =   1560
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   476
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
               ButtonImage     =   "FrmCompany.frx":5EF2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   1560
               TabIndex        =   148
               Top             =   1440
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› "
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
               ButtonImage     =   "FrmCompany.frx":628C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   120
               TabIndex        =   149
               Top             =   1440
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   688
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
               ButtonImage     =   "FrmCompany.frx":25476
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
               Height          =   315
               Index           =   18
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   1200
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… ··„·Õﬁ"
               Height          =   315
               Index           =   17
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   1560
               Width           =   1185
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Height          =   315
               Index           =   16
               Left            =   5340
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   1200
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„·Õﬁ"
               Height          =   315
               Index           =   15
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   1560
               Width           =   945
            End
            Begin VB.Label txtRate 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               Caption         =   "1.3"
               Height          =   315
               Left            =   3963
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· »⁄Ì…"
               Height          =   315
               Index           =   12
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   1200
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·”«∆ﬁ"
               Height          =   312
               Index           =   11
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   480
               Width           =   1188
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„«—ﬂ…"
               Height          =   228
               Index           =   5
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   120
               Width           =   768
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘«”Ì…"
               Height          =   312
               Index           =   5
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   120
               Width           =   1188
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ÊœÌ·"
               Height          =   312
               Index           =   6
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   840
               Width           =   1188
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·„ﬁ«⁄œ"
               Height          =   312
               Index           =   7
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   120
               Width           =   1068
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„⁄œ· «·«—ﬂ«»"
               Height          =   315
               Index           =   8
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ·Ì›Ê‰ «·”«∆ﬁ"
               Height          =   315
               Index           =   9
               Left            =   5340
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   480
               Width           =   945
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   4995
            Left            =   0
            TabIndex        =   123
            Top             =   2280
            Width           =   10665
            _cx             =   18812
            _cy             =   8811
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16776960
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
            Rows            =   50
            Cols            =   21
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCompany.frx":44660
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„œÌ‰…"
      Height          =   225
      Index           =   39
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   156
      Top             =   0
      Width           =   765
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   312
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   8520
      Width           =   612
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   312
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   8520
      Width           =   672
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   312
      Index           =   0
      Left            =   2076
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   8520
      Width           =   1632
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   312
      Index           =   4
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   8520
      Width           =   252
   End
End
Attribute VB_Name = "FrmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim TTP As clstooltip
Dim ComReport As ClsCompanyReport
Dim Dcombo As ClsDataCombos
Dim cSearch(2) As clsDCboSearch
Dim FirstPeriodDateInthisYear  As Date
 Dim rsVendor As ADODB.Recordset
Public mIndex As Integer
Dim s As String
Dim dummy As ADODB.Recordset
Sub ReloadCompo()
Dim sql As String

sql = "SELECT DISTINCT BankName,BankName"
sql = sql & " From dbo.TblCustemers"
sql = sql & " WHERE     (NOT (BankName IS NULL)) "
fill_combo txtBankName, sql
 

sql = "SELECT DISTINCT BankCode,BankCode"
sql = sql & " From dbo.TblCustemers"
sql = sql & " WHERE     (NOT (BankCode IS NULL)) "
fill_combo TxtBankCode, sql


End Sub

Private Sub ALLButton1_Click()

    Frame2.Visible = True
End Sub

Private Sub CboDiscountType_Change()
    Me.lbl(21).Visible = (Me.CboDiscountType.ListIndex = 2)

    If CboDiscountType.ListIndex = 0 Then
        lbl(20).Visible = False
        TxtDiscountValue.Visible = False
        lbl(21).Visible = False
    Else
        lbl(20).Visible = True
        TxtDiscountValue.Visible = True
        lbl(21).Visible = True
    End If

End Sub

Private Sub CboDiscountType_Click()
    CboDiscountType_Change
End Sub

Private Sub CboDiscountTypePur_Change()
    Me.lbl(27).Visible = (Me.CboDiscountTypePur.ListIndex = 2)

    If CboDiscountTypePur.ListIndex = 0 Then
        lbl(28).Visible = False
        TxtDiscountValuePur.Visible = False
        lbl(27).Visible = False
    Else
        lbl(28).Visible = True
        TxtDiscountValuePur.Visible = True
        lbl(27).Visible = True
    End If

End Sub

Private Sub CboDiscountTypePur_Click()
    CboDiscountTypePur_Change
End Sub

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.text = 0
    Cmd_Click (2)

End Function

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
 '   On Error GoTo ErrTrap

    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear
 
    Select Case Index

        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
 If mIndex = 1 Then
 CHkMot3ahed.value = vbChecked
 End If
       txtNoOFDigitUser(10) = "SA"
            DcboCountryID.BoundText = 1
            DcboGovernmentID.BoundText = 1
            DcboCityID.BoundText = 1
            
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.dcBranch.BoundText = Current_branch
   
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(9, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else
                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··„Ê—œÌ‰   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "A parent account has not been supplied for suppliers in this branch for this proccess", vbCritical
                    End If
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            OptType(2).value = True
            Grid.rows = Grid.FixedRows
            DcbCurrency.BoundText = MainCurrency()
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            '        If XPTxtComID.text = 1 Then
            '            Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·”Ã·"
            '            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '            Exit Sub
            '        End If
            TxtModFlg.text = "E"

        Case 2
            CREATEADDRESS
            Dim currentcode As String

            If txtid.text = "" Then
                currentcode = get_coding(branch_id, "TblCustemers", 5, Me.DCPreFix.text, True)

                If currentcode = "miniError" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Else
                        MsgBox "The number of digits chosen for this code is too small please change the coding policy in coding window or contact your administrator"
                    End If
                    Exit Sub
                ElseIf currentcode = "Manual" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Else
                        MsgBox "Please enter the code manually "
                  End If
                    Exit Sub
                    
                Else
                    txtid = currentcode
                End If
            End If

            SaveData

        Case 3
            Call Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If XPTxtComID.text = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã·"
                Else
                Msg = "This recored date can't be deleted"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
                Exit Sub
            End If

            Del_Company
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

         FrmCompanySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
print_report2
           ' printingReport

        Case 6
            Unload Me

        Case 8

            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
            ShowReport IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), XPTxtComName.text, FirstPeriod, Date
            Case 9
            
              If Grid.Row < Grid.FixedRows Then Exit Sub
                Dim StrSQL  As String
                
                If ISCarAllowDelete(val(Grid.TextMatrix(Grid.Row, Grid.ColIndex("ID")))) Then
                        '  strSQL = " delete from TblVendorCars where ID =    " & val(Grid.TextMatrix(Grid.Row, Grid.ColIndex("ID")))
                        '  Cn.Execute strSQL, , adExecuteNoRecords
                          Grid.RemoveItem (Grid.Row)
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox ("€Ì— „”„ÊÕ Õ–› «·„⁄œÂ/«·”Ì«—… " & Grid.TextMatrix(Grid.Row, Grid.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                    Else
                        MsgBox ("vehicle can't be deleted  " & Grid.TextMatrix(Grid.Row, Grid.ColIndex("BoardNo")) & "for data integration ")
                    End If
                End If
                
            Case 10
                   DelAll
            Case 11
            On Error Resume Next
ShowAttachments DCPreFix.text & txtid.text, "0701201402"
 Case 12
    Unload FrmOldContract
 
  FrmOldContract.ScrenFlg = 1
  FrmOldContract.show
Case 14
 addrow
 Grid.rows = Grid.rows + 1

    End Select

    Exit Sub
ErrTrap:
End Sub

 Public Function ISCarAllowDelete(CarID As Integer) As Boolean
Dim str As String, allow As Boolean

allow = True

'Check Attribution Contract
str = " Select * from TblVehicleAllocation_Details where type = 3 and CarID =   " & CarID
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        ISCarAllowDelete = False
        Exit Function
End If



'Check Confirm Violation
str = " Select * from TblConfirmViolation  where  CarID  =   " & CarID
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        ISCarAllowDelete = False
        Exit Function
End If

'Check Vendor Request
str = " Select * from TblExchangeReques_Detailst where carid   =   " & CarID
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        ISCarAllowDelete = False
        Exit Function
End If
 
'Check Stop Dealing
str = " select * from TblStopDealing where carid = " & CarID
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        ISCarAllowDelete = False
        Exit Function
End If
 
 ISCarAllowDelete = True
 
 End Function


Private Sub DelAll()
 If Grid.rows <= Grid.FixedRows Then Exit Sub
 Dim i  As Integer, m As Integer, StrSQL As String
 m = Grid.rows - Grid.FixedRows
 i = Grid.rows - 1
 Do While Grid.rows > Grid.FixedRows
      If ISCarAllowDelete(val(Grid.TextMatrix(i, Grid.ColIndex("ID")))) Then
                  StrSQL = " delete from TblVendorCars where ID =    " & val(Grid.TextMatrix(i, Grid.ColIndex("ID")))
                  Cn.Execute StrSQL, , adExecuteNoRecords
                  Grid.RemoveItem (i)
                  i = i - 1
        Else
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox ("€Ì— „”„ÊÕ Õ–› «·„⁄œÂ/«·”Ì«—… " & Grid.TextMatrix(i, Grid.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                Else
                    MsgBox ("vehicle can't be deleted  " & Grid.TextMatrix(Grid.Row, Grid.ColIndex("BoardNo")) & "for data integration ")
                End If
                
                  Exit Sub
        End If
 Loop
End Sub

Private Sub addrow()
If dcBrand.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("«Œ — «·„«—ﬂ… «Ê·«")
            Else
            MsgBox ("Select Brand")
            End If
dcBrand.SetFocus
  Sendkeys ("{F4}")
Exit Sub
End If

Dim board As String
Dim lettter As String
Dim Num As String
Dim nboard As String
Dim nlettter As String
Dim nNum As String


lettter = txtLetter1.text & " " & txtLetter2.text & " " & txtLetter3.text & " " & txtLetter4.text
Num = txtNum1.text & " " & txtNum2.text & " " & txtNum3.text & " " & txtNum4.text

nlettter = ntxtLetter1.text & " " & ntxtLetter2.text & " " & ntxtLetter3.text & " " & ntxtLetter4.text
nNum = ntxtNum1.text & " " & ntxtNum2.text & " " & ntxtNum3.text & " " & ntxtNum4.text

board = lettter & " " & Num

nboard = nlettter & " " & nNum

If Len(lettter) + Len(Num) < 2 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("·«»œ „‰ « »«⁄  ›Ê—„«  «·«œŒ«· ")
    Else
        MsgBox ("Should follow the input formate ")
    End If
    Exit Sub
End If
    
    
    Dim ss As String
   Set RsTemp = New ADODB.Recordset
  ss = "  select * from TblVendorCars where replace(BoardNo,' ','') = '" & Replace(board, " ", "") & "'"
    RsTemp.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If RsTemp.RecordCount > 0 Then
            frmCarDetails.ven = IIf(IsNull(RsTemp("CustomerID").value), 0, RsTemp("CustomerID").value)
              frmCarDetails.board = board
              frmCarDetails.show
            Exit Sub
    End If

    Dim s As Integer
    
    For s = 1 To Grid.rows - 1
            If Replace(Grid.TextMatrix(s, Grid.ColIndex("BoardNo")), " ", "") = Replace(board, " ", "") Then
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox (" „ «÷«›… «·„⁄œÂ/«·”Ì«—… Â–Â „‰ ﬁ»·  ")
                Else
                    MsgBox ("This vehicle has been add befor")
                End If
                Exit Sub
            End If
    Next

Dim i As Integer


Dim j As Integer, ad As Boolean

ad = False

For j = 1 To Grid.rows - 1
        If Grid.TextMatrix(j, Grid.ColIndex("BrandID")) = "" Then
                i = j
                ad = True
                Exit For
        End If
Next

If ad = False Then
Grid.rows = Grid.rows + 1
i = Grid.rows
i = i - 1
End If


With Grid
    .TextMatrix(i, .ColIndex("Serial")) = i - 1
    .TextMatrix(i, .ColIndex("BoardNo")) = board
    .TextMatrix(i, .ColIndex("nBoardNo")) = nboard
    .TextMatrix(i, .ColIndex("ChasisNo")) = txtChassis.text
    .TextMatrix(i, .ColIndex("BrandID")) = dcBrand.BoundText
    .TextMatrix(i, .ColIndex("Brand")) = dcBrand.text
    .TextMatrix(i, .ColIndex("ModelID")) = cbModel.ListIndex
    .TextMatrix(i, .ColIndex("Model")) = cbModel.text
    .TextMatrix(i, .ColIndex("Count")) = txtCount.text
    .TextMatrix(i, .ColIndex("Rate")) = txtRate.Caption
    .TextMatrix(i, .ColIndex("CityID")) = DcCity.ListIndex
    .TextMatrix(i, .ColIndex("City")) = DcCity.text
    .TextMatrix(i, .ColIndex("DriverName")) = txtDriverName.text
    .TextMatrix(i, .ColIndex("DriverTel")) = txtDriverTel.text
    .TextMatrix(i, .ColIndex("accessory")) = accessoryTxt.text
    .TextMatrix(i, .ColIndex("Price")) = val(TxtPrice.text)
    .TextMatrix(i, .ColIndex("PartPrice")) = val(TxtPartPrice.text)
    .TextMatrix(i, .ColIndex("TypeTransID")) = val(DcbTypTrans.ListIndex)
    .TextMatrix(i, .ColIndex("TypeTrans")) = DcbTypTrans.text
End With
'

txtChassis.text = ""
dcBrand.BoundText = ""
cbModel.ListIndex = -1
txtCount.text = ""
txtDriverName.text = ""
DcCity.ListIndex = -1
 
txtLetter1.text = ""
txtLetter2.text = ""
txtLetter3.text = ""
txtLetter4.text = ""

txtNum1.text = ""
txtNum2.text = ""
txtNum3.text = ""
txtNum4.text = ""

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdPriceList_Click()
    On Error GoTo ErrTrap
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    '«· √ﬂœ √‰ «·„Ê—œ ·Â ﬁ«∆„… √”⁄«—
    StrSQL = "SELECT CusJuncItem.ID,CusJuncItem.LastUpdate, CusJuncItem.CusID, CusJuncItem.ItemID, " & "CusJuncItem.ItemPrice, TblItems.ItemCode, TblItems.ItemName FROM TblItems " & " INNER JOIN CusJuncItem ON TblItems.ItemID = CusJuncItem.ItemID where CusID = " & XPTxtComID.text
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.EOF Or rs.BOF) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–« «·„Ê—œ ·Ì” ·Â ﬁ«∆„… √”⁄«—"
        Else
            Msg = "This Supplier doesn't have a price list"
            End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CmdPriceList.Enabled = False
        Exit Sub
    End If

    '⁄—÷ ﬁ«∆„… «·√”⁄«— «·Œ«’… »«·„Ê—œ
    If XPTxtComID.text <> "" Then

         

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 88
    End If

End Sub

Private Sub DcboCityID_Change()
    LoadDataCombos False, False, True
End Sub

Private Sub DcboCityID_Click(Area As Integer)
    DcboCityID_Change
End Sub

Private Sub DcboCityID_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

End Sub

Private Sub DcboCountryID_Change()
    LoadDataCombos True, False, False
End Sub

Private Sub DcboCountryID_Click(Area As Integer)

    If val(Me.DcboCountryID.BoundText) <> 0 Then
        DcboCountryID_Change
    End If

End Sub

Private Sub DcboCountryID_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

End Sub

Private Sub DcboGovernmentID_Change()
    LoadDataCombos False, True, False
End Sub

Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub
 
Private Sub DcboGovernmentID_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

End Sub

Private Sub Form_Activate()
    'XPTxtComID.SetFocus
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
MySQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_Phone, "
MySQL = MySQL & "                      dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Type, dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType,"
MySQL = MySQL & "                      dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.CreditLimit, dbo.TblCustemers.Account_Code_As_Client, dbo.TblCustemers.Account_Code_As_Supplier,"
MySQL = MySQL & "                      dbo.TblCustemers.CreditlimitCredit, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.SaleType, dbo.TblCustemers.Account_Code,"
MySQL = MySQL & "                      dbo.TblCustemers.Trans_Discount, dbo.TblCustemers.Trans_DiscountType, dbo.TblCustemers.CountryID, dbo.TblCountriesData.CountryName,"
MySQL = MySQL & "                      dbo.TblCustemers.CityID, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCustemers.GovernmentID, dbo.TblCountriesGovernments.GovernmentName,"
MySQL = MySQL & "                      dbo.TblCustemers.Address, dbo.TblCustemers.Trans_DiscountPur, dbo.TblCustemers.Trans_DiscountTypePur, dbo.TblCustemers.CountEmp, dbo.TblCustemers.ToTal,"
MySQL = MySQL & "                       dbo.TblCustemers.c1, dbo.TblCustemers.c2, dbo.TblCustemers.Remark2, dbo.TblCustemers.locked, dbo.TblCustemers.parent_account,"
MySQL = MySQL & "                      dbo.TblCustemers.opening_balance_voucher_id, dbo.TblCustemers.DepitInterval, dbo.TblCustemers.CreditInterval, dbo.TblCustemers.DepitIntervalID,"
MySQL = MySQL & "                      dbo.TblCustemers.CreditIntervalID, dbo.TblCustemers.EmpId, dbo.TblCustemers.prifix, dbo.TblCustemers.code, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "                      dbo.TblCustemers.CustomerandVendor, dbo.TblCustemers.CustomerTypeID, dbo.TblCustemers.BranchId, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.Company, dbo.TblCustemers.JobTitle,"
MySQL = MySQL & "                      dbo.TblCustemers.Salary, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel,"
MySQL = MySQL & "                      dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.CountryID2, dbo.TblCustemers.Sex, dbo.TblCustemers.Account_Code1,"
MySQL = MySQL & "                      dbo.TblCustemers.Account_Code2, dbo.TblCustemers.ParentAccount, dbo.TblCustemers.OpenBalanceType1, dbo.TblCustemers.OpenBalance1,"
MySQL = MySQL & "                      dbo.TblCustemers.OpenBalanceType2, dbo.TblCustemers.OpenBalance2, dbo.TblCustemers.ShowQty1, dbo.TblCustemers.showPrice1,"
MySQL = MySQL & "                      dbo.TblCustemers.showPrice2, dbo.TblCustemers.Salaries1, dbo.TblCustemers.Salaries2, dbo.TblCustemers.ShowQty1c, dbo.TblCustemers.showPrice1c,"
MySQL = MySQL & "                      dbo.TblCustemers.showPrice2c, dbo.TblCustemers.Salaries1c, dbo.TblCustemers.Salaries2c, dbo.TblCustemers.Totald, dbo.TblCustemers.Totalc,"
MySQL = MySQL & "                      dbo.TblCustemers.RecordDate, dbo.TblCustemers.balanced, dbo.TblCustemers.balancec, dbo.TblCustemers.TypeCustomer, dbo.TblCustemers.BoxMil,"
MySQL = MySQL & "                      dbo.TblCustemers.ZipCode , dbo.ACCOUNTS.account_serial  , dbo.TblCustemers.BankIBAN , dbo.TblCustemers.BankAccount , dbo.TblCustemers.VATNO"
MySQL = MySQL & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesData ON dbo.TblCustemers.CountryID = dbo.TblCountriesData.CountryID"
MySQL = MySQL & " Where (dbo.TblCustemers.CusID =" & val(XPTxtComID.text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repAents.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repAents.rpt"
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
            Msg = "There's no data to show "
            End If
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
End Function
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
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
    Dim StrSQL As String

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    On Error GoTo ErrTrap
    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "»Ì«‰«  «·„Ê—œÌ‰ "
    LogTexte = " Open Window " & " Suppliers Data"
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""
    StrSQL = " select id,code from currency"
    fill_combo Me.DcbCurrency, StrSQL
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    LoadDataCombos
    
    
     If SystemOptions.IsHiddenUser Then
        chkIsBranch(4).Visible = True
    Else
        chkIsBranch(4).Visible = False
    End If

With DcbTypTrans
.Clear
If SystemOptions.UserInterface = ArabicInterface Then
.AddItem "ﬂ«„·"
.AddItem "»œÊ‰ „·Õﬁ"
Else
.AddItem "Complete"
.AddItem "Without Part "
End If
End With
With DcbTypTrans1
.Clear
If SystemOptions.UserInterface = ArabicInterface Then
.AddItem "ﬂ«„·"
.AddItem "»œÊ‰ „·Õﬁ"
Else
.AddItem "Complete"
.AddItem "Without Part "
End If
End With
    With CboDiscountType
        .Clear
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With

    With CboDiscountTypePur
        .Clear
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With

    With Me.dcDepitIntervalID
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÌÊ„"
            .AddItem "‘Â—"
            .AddItem "”‰…"
        Else
            .AddItem "day"
            .AddItem "month"
            .AddItem "year"
        End If

    End With

    With Me.cbCategory
    .Clear
    
        If SystemOptions.UserInterface = ArabicInterface Then
                .AddItem "„Ê—œ"
                .AddItem "„ ⁄Âœ"
        Else
                .AddItem "Vendor"
                .AddItem "Contractor"
        End If
    End With


Dim k As Integer

With Me.cbModel
.Clear
For k = 1900 To 2050
        .AddItem k
Next
End With



    With Me.dcCreditIntervalID
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÌÊ„"
            .AddItem "‘Â—"
            .AddItem "”‰…"
        Else
            .AddItem "day"
            .AddItem "month"
            .AddItem "year"
        End If

    End With


Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
str = " Select id , name  from TBLCarTypes "
Else
str = " Select id , namee  from TBLCarTypes "
End If

fill_combo dcBrand, str


  '  Dcombo.getCountriesGovernments DcCity
   If SystemOptions.UserInterface = ArabicInterface Then
    With DcCity
    .Clear
    .AddItem ("—∆Ì”Ï")
    .AddItem ("»«ÿ‰")
    End With
    Else
    With DcCity
    .Clear
    .AddItem ("Main")
    .AddItem ("Inside")
    End With
  
    End If
  
  

    Me.Dtp.value = Date
    'Resize_Form Me
    AddTip
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True

    Dcombos.GetCodeing Me.DCPreFix, 5, "TblCustemers", "Type =2"
Dcombos.GetSalesRepDatapurchase Me.DCEmP


    StrSQL = "select * From TblCustemers where Type=2"
    
     If mIndex = 1 Then
     StrSQL = StrSQL & " and CHkMot3ahed =1"
     Else
     StrSQL = StrSQL & " and isnull(CHkMot3ahed,0) =0"
   End If
 
 
            If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & "  AND     (BranchId=0 or      BranchId in(" & Current_branchSql & "))"
            
     '       StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
        
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

'Grid.Rows = 1
txtLetter1.MaxLength = 1
txtLetter2.MaxLength = 1
txtLetter3.MaxLength = 1
txtLetter4.MaxLength = 1
txtNum1.MaxLength = 1
txtNum2.MaxLength = 1
txtNum3.MaxLength = 1
txtNum4.MaxLength = 1

C1Tab1.CurrTab = 0

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "   «·Œ—ÊÃ „‰ " & "»Ì«‰«  «·„Ê—œÌ‰ "
    LogTexte = " Exit Window " & " Suppliers Data"
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
    Set ComReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·„Ê—œÌ‰ " _
       & CHR(13) & " ﬂÊœ «·„Ê—œ  " & DCPreFix & txtid.text _
       & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtComName _
       & CHR(13) & "   „”∆Ê· «·« ’«·   " & TxtResponsibleContact _
       & CHR(13) & " —ﬁ„ «·Â« ›     " & XPTxtPhone _
       & CHR(13) & " —ﬁ„ «·ÃÊ«·     " & XPTxtmobile _
       & CHR(13) & " —ﬁ„ «·›«ﬂ”     " & TxtFaxNumber _
       & CHR(13) & "  «·»—Ìœ «·«·ﬂ —Ê‰Ì       " & TxtE_mail _
       & CHR(13) & " «·œÊ·Â   " & DcboCountryID.text _
       & CHR(13) & " «·„Õ«›Ÿ…   " & DcboGovernmentID.text _
       & CHR(13) & "  «·„œÌ‰…  " & DcboCityID.text _
       & CHR(13) & "  «·⁄‰Ê«‰ »«· ›’Ì· " & TxtAddress _
       & CHR(13) & " „·«ÕŸ«   " & XPMTxtRemark.text _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„»Ì⁄«    " & CboDiscountType.text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValue _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„‘ —Ì«    " & CboDiscountTypePur.text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValuePur _
       & CHR(13) & "  ‰Ê⁄ «·„Ê—œ  " & DcCustomerType.text _
       & CHR(13) & " Õœ «·«∆ „«‰ „œÌ‰  " & TxtCreditLimit _
       & CHR(13) & " „œ… «·«∆ „«‰     " & TxtDepitInterval.text & "   " & dcDepitIntervalID.text _
       & CHR(13) & " Õœ «·«∆ „«‰ œ«∆‰   " & TxtCreditlimitCredit _
       & CHR(13) & " „œ… «·«∆ „«‰      " & TxtCreditInterval.text & "   " & dcCreditIntervalID.text _

       LogTextA = LogTextA & CHR(13) & "⁄„Ì· „Ê—œ ø       "

    If chkCustomerandVendor.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
    Else
        LogTextA = LogTextA & "·«"
    End If

    LogTextA = LogTextA & CHR(13) & "«Ìﬁ«› «· ⁄«„·   ø     "

    If locked.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
        LogTextA = LogTextA & CHR(13) & "  ”»» «·«Ìﬁ«›   "
        LogTextA = LogTextA & CHR(13) & XPMTxtRemarks2
    Else
        LogTextA = LogTextA & "·«"
    End If

    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "   œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ     " & TxtOpenBalance
    LogTextA = LogTextA & CHR(13) & "«·Õ”«» «·—∆Ì”Ì    " & DboParentAccount

    LogTexte = "  Õ›Ÿ ‘«‘… " & " Customers Data  " _
       & CHR(13) & "  Code  " & DCPreFix & txtid.text _
       & CHR(13) & "Name " & XPTxtCusNamee _
       & CHR(13) & " Contact Person" & TxtResponsibleContact _
       & CHR(13) & " Tel " & XPTxtPhone _
       & CHR(13) & "Mob " & XPTxtmobile _
       & CHR(13) & " Fax  " & TxtFaxNumber _
       & CHR(13) & "  Email   " & TxtE_mail _
       & CHR(13) & " Contry   " & DcboCountryID.text _
       & CHR(13) & " City   " & DcboGovernmentID.text _
       & CHR(13) & "  Town  " & DcboCityID.text _
       & CHR(13) & " Address " & TxtAddress _
       & CHR(13) & " Remarks  " & XPMTxtRemark _
       & CHR(13) & " Sales Discount  type  " & CboDiscountType.text _
       & CHR(13) & " Discount Value " & TxtDiscountValue _
       & CHR(13) & " Purchase Discount type " & CboDiscountTypePur.text _
       & CHR(13) & "  Discount Value" & TxtDiscountValuePur _
       & CHR(13) & "  Supplier . Type " & DcCustomerType.text _
       & CHR(13) & "The limit for debit  " & TxtCreditLimit _
       & CHR(13) & " Period     " & TxtDepitInterval.text & "   " & dcDepitIntervalID.text _
       & CHR(13) & "The limit for Credit   " & TxtCreditlimitCredit _
       & CHR(13) & " Period " & TxtCreditInterval.text & "   " & dcCreditIntervalID.text _

       LogTexte = LogTexte & CHR(13) & "Customer & Supplier ?  "

    If chkCustomerandVendor.value = vbChecked Then
        LogTexte = LogTexte & " Yes "
    Else
        LogTexte = LogTexte & " No "
    End If

    LogTexte = LogTexte & CHR(13) & "Locked"

    If locked.value = vbChecked Then
        LogTexte = LogTexte & "Yes "
        LogTexte = LogTexte & CHR(13) & "  Reasons  "
        LogTexte = LogTexte & CHR(13) & XPMTxtRemarks2
    Else
        LogTexte = LogTexte & "No "
    End If

    LogTexte = LogTexte & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTexte = LogTexte & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTexte = LogTexte & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTexte = LogTexte & "€Ì— „Õœœ"
    End If

    LogTexte = LogTexte & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTexte = LogTexte & CHR(13) & "  Parent Acc. " & DboParentAccount
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    With Grid
     
         Select Case .ColKey(Col)
         
                    Case "Brand"
                        
                        .TextMatrix(.Row, .ColIndex("BrandID")) = .ComboData
                        Grid.rows = Grid.rows + 1
                        
                     Case "Model"
                     Dim k As Integer
                     k = val(.TextMatrix(.Row, .ColIndex("Model")))
                     k = k - 1900
                    .TextMatrix(.Row, .ColIndex("ModelID")) = k
                    
          End Select

    End With




   

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
     Dim i As Integer
     i = IIf(.TextMatrix(Row, .ColIndex("id")) = "", 0, val(.TextMatrix(Row, .ColIndex("id"))))
     If i > 0 Then
            If ISCarAllowDelete(i) = False Then
                    Cancel = True
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox ("€Ì— „”„ÊÕ «· ⁄œÌ· ⁄·Ï  «·„⁄œÂ/«·”Ì«—… " & Grid.TextMatrix(Grid.Row, Grid.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                    Else
                        MsgBox ("editing isn't allowed for this vehicle" & Grid.TextMatrix(Grid.Row, Grid.ColIndex("BoardNo")) & "for data integration")
                    End If
                    Exit Sub
            End If
     End If
     Select Case .ColKey(Col)

    Case "BoardNo"
            Cancel = True

End Select

End With


End Sub

Private Sub Grid_DblClick()

With Grid
        If .TextMatrix(.Row, .ColIndex("BoardNo")) <> "" Then
                
        End If
End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With Grid
     
     Select Case .ColKey(Col)
    
    Case "BoardNo"
         .ComboList = ""
   Case "ChasisNo"
            .ComboList = ""
     Case "Brand"
                    
          StrSQL = "  Select id , name  from TBLCarTypes ORDER BY ID "
          RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = Grid.BuildComboList(RsTemp, "Name", "ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
          
          
      Case "Model"
        '    StrSQL = "  Select id , name  from TBLCarTypes ORDER BY ID "
        '    rstemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
           ' StrComboList = Grid.BuildComboList(rstemp, "Name", "ID")
           ' If StrComboList <> "" Then
           '        StrComboList = "|" & StrComboList
           ' End If
            Dim str As String, k As Integer
           
             
             For k = 1900 To 2050
                    str = str & "|" & k
             Next
            .ComboList = str
        Case "Count"
                .ComboList = ""
                
        Case "Rate"
        .ComboList = ""
        
        Case "DriverName"
        .ComboList = ""
        
        Case "DriverTel"
        .ComboList = ""
        
        Case Is = "EndDate"
        .ComboList = ""
        
        
        
        
        
      
   End Select
End With


End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Label2_Click()
    Frame2.Visible = False
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub





Private Sub TxtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditLimit.text, 1)
End Sub

Private Sub TxtCreditlimitCredit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditlimitCredit.text, 1)
End Sub


Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.text = ""
If Len(txtLetter1.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select
End Sub


Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.text = ""
If Len(txtLetter2.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.text = ""
If Len(txtLetter3.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.text = ""
If Len(txtLetter4.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
End Sub


Private Sub ntxtLetter1_KeyPress(KeyAscii As Integer)
ntxtLetter1.text = ""
If Len(ntxtLetter1.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        ntxtLetter2.SetFocus
End Select
End Sub


Private Sub ntxtLetter2_KeyPress(KeyAscii As Integer)
ntxtLetter2.text = ""
If Len(ntxtLetter2.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtLetter3.SetFocus
End Select
End Sub

Private Sub ntxtLetter3_KeyPress(KeyAscii As Integer)
ntxtLetter3.text = ""
If Len(ntxtLetter3.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtLetter4.SetFocus
End Select
End Sub

Private Sub ntxtLetter4_KeyPress(KeyAscii As Integer)
ntxtLetter4.text = ""
If Len(ntxtLetter4.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtNum1.SetFocus
End Select
End Sub

Private Sub TxtModFlg_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„Ê—œÌ‰"
            Else
                Me.Caption = "Suppliers Data"
            End If

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
        
            Me.XPTxtComID.locked = True
            Me.XPTxtComName.locked = True
            Me.XPTxtPhone.locked = True
            Me.XPTxtmobile.locked = True
            Me.XPMTxtRemark.locked = True

            If XPTxtComID.text <> "" Then
                If XPTxtComID.text = 1 Then
                    CmdPriceList.Enabled = False
                Else
                    CmdPriceList.Enabled = True
                End If
            End If

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                CmdPriceList.Enabled = False
            End If

            Fra(0).Enabled = False

            '        Me.Dtp.Enabled = True
        Case "N"
            DboParentAccount.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„Ê—œÌ‰( ÃœÌœ )"
            Else
                Me.Caption = "Suppliers Data(Enter New Record)."
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtComID.locked = True
            Me.XPTxtComName.locked = False
            Me.XPTxtPhone.locked = False
            Me.XPMTxtRemark.locked = False
            Me.XPTxtmobile.locked = False
            CmdPriceList.Enabled = False
            Fra(0).Enabled = True

            '        Me.Dtp.Enabled = True
        Case "E"
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„Ê—œÌ‰(  ⁄œÌ· )"
            Else
                Me.Caption = "Suppliers Data(Edit Record)."
            End If

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
        
            Me.XPTxtComID.locked = True
            Me.XPTxtComName.locked = False
            Me.XPTxtmobile.locked = False
            Me.XPTxtPhone.locked = False
            Me.XPMTxtRemark.locked = False
            CmdPriceList.Enabled = False
            Fra(0).Enabled = True
            '        Me.Dtp.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim SngCusBegainAccount As Single

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
        rs.Find "CusID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    DcbCurrency.BoundText = IIf(IsNull(rs("CurrncyID").value), "", rs("CurrncyID").value)
    Me.DCEmP.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))

    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.txtid.text = IIf(IsNull(rs("code").value), "", rs("code").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     TxtBankCode.text = IIf(IsNull(rs("BankCode").value), "", rs("BankCode").value)
     
  If IsNull(rs("IsHiddenInv").value) Then
        Me.chkIsBranch(4).value = vbUnchecked
    Else
        Me.chkIsBranch(4).value = IIf(rs("IsHiddenInv").value = 0, vbUnchecked, vbChecked)
    End If
     TxtBankIBAN.text = IIf(IsNull(rs("BankIBAN").value), "", rs("BankIBAN").value)
     TxtBankAddress.text = IIf(IsNull(rs("BankAddress").value), "", rs("BankAddress").value)
     TxtVATNO.text = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), -100, rs("opening_balance_voucher_id").value)
    XPTxtComID.text = IIf(IsNull(rs("CusID").value), "", val(rs("CusID").value))
    Me.TxtCode = IIf(IsNull(rs("c1").value), "", rs("c1").value)
    XPTxtComName.text = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
    Me.TxtResponsibleContact.text = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
    XPTxtPhone.text = IIf(IsNull(rs("Cus_Phone").value), "", Trim(rs("Cus_Phone").value))
    XPTxtmobile.text = IIf(IsNull(rs("Cus_mobile").value), "", Trim(rs("Cus_mobile").value))
    XPMTxtRemark.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    
        txtNoOFDigitUser(2).text = IIf(IsNull(rs("StreetName").value), "", rs("StreetName").value)
txtNoOFDigitUser(4).text = IIf(IsNull(rs("BuildingNumber").value), "", rs("BuildingNumber").value)
'txtNoOFDigitUser(9).Text = IIf(IsNull(rs("CitySubdivisionName").value), "", rs("CitySubdivisionName").value)
'txtNoOFDigitUser(6).Text = IIf(IsNull(rs("CityName").value), "", rs("CityName").value)
'txtNoOFDigitUser(7).Text = IIf(IsNull(rs("PostalZone").value), "", rs("PostalZone").value)
txtNoOFDigitUser(10).text = IIf(IsNull(rs("IdentificationCode").value), "", rs("IdentificationCode").value)
txtNoOFDigitUser(5).text = IIf(IsNull(rs("PlotIdentification").value), "", rs("PlotIdentification").value)
txtNoOFDigitUser(3).text = IIf(IsNull(rs("AdditionalStreetName").value), "", rs("AdditionalStreetName").value)
txtNoOFDigitUser(8).text = IIf(IsNull(rs("CountrySubentity").value), "", rs("CountrySubentity").value)

txtNoOFDigitUser(0).text = IIf(IsNull(rs("Id700").value), "", rs("Id700").value)
 


    XPTxtCusNamee.text = IIf(IsNull(rs("CusNamee")), "", Trim(rs("CusNamee")))
    XPMTxtRemarks2.text = IIf(IsNull(rs("Remark2")), "", Trim(rs("Remark2")))
    locked.value = IIf(rs("locked") = True, 1, 0)
    Me.DboParentAccount.BoundText = IIf(IsNull(rs("parent_account")), "", rs("parent_account"))
    Me.DcCustomerType.BoundText = IIf(IsNull(rs("CustomerTypeID")), "", rs("CustomerTypeID"))
Dim Account_code As String
Account_code = IIf(IsNull(rs("Account_Code")), "", rs("Account_Code"))

    c(1).Caption = GetACCOUNTSCode(Account_code, 1)
    cbCategory.ListIndex = IIf(IsNull(rs("Category").value), -1, rs("Category").value)
       
      
    If rs("CustomerandVendor").value = True Then
        chkCustomerandVendor.value = vbChecked

    Else
        chkCustomerandVendor.value = vbUnchecked
    End If

      
      
   If rs("CustomerandVendor").value = True Then
        chkCustomerandVendor.value = vbChecked

    Else
        chkCustomerandVendor.value = vbUnchecked
    End If
   
   
    If IsNull(rs("chkTaxExempt").value) Then
        Me.chkTaxExempt.value = vbUnchecked
    Else
        Me.chkTaxExempt.value = IIf(rs("chkTaxExempt").value = 0, vbUnchecked, vbChecked)
    End If
    

   
    If rs("CHkMot3ahed").value = True Then
        CHkMot3ahed.value = vbChecked

    Else
        CHkMot3ahed.value = vbUnchecked
    End If
    

    If XPTxtComID.text = 1 Then
        CmdPriceList.Enabled = False
    Else
        CmdPriceList.Enabled = True
    End If

    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
         
    Else
    
        Me.Dtp.value = Date
        Me.Dtp.Enabled = False
    End If

    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
    
    Else
        Me.TxtOpenBalance.text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If

    TxtCreditLimit.text = IIf(IsNull(rs("CreditLimit").value), "0", rs("CreditLimit").value)
    Me.TxtCreditlimitCredit.text = IIf(IsNull(rs("CreditlimitCredit").value), "0", rs("CreditlimitCredit").value)
    Me.TxtFaxNumber.text = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
    Me.TxtE_mail.text = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
    'SngCusBegainAccount = GetCustomerAccount(val(XPTxtComID.text), True)

    Dim balanceString As String
    WriteCustomerBalPublic IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), , balanceString
    lbl(8).Caption = balanceString

    'If SngCusBegainAccount < 0 Then
    '    Me.lbl(8).Caption = Abs(SngCusBegainAccount)
    '    Me.lbl(9).Caption = "„œÌ‰"
    'ElseIf SngCusBegainAccount > 0 Then
    '    Me.lbl(8).Caption = Abs(SngCusBegainAccount)
    '    Me.lbl(9).Caption = "œ«∆‰"
    'Else
    '    Me.lbl(8).Caption = 0
    '    Me.lbl(9).Caption = ""
    'End If

    If IsNull(rs("Trans_DiscountType").value) Then
        Me.CboDiscountType.ListIndex = 0
        Me.TxtDiscountValue.text = 0
    ElseIf rs("Trans_DiscountType").value = 0 Then
        Me.CboDiscountType.ListIndex = 0
        Me.TxtDiscountValue.text = 0
    ElseIf rs("Trans_DiscountType").value = 1 Then
        Me.CboDiscountType.ListIndex = 1
        Me.TxtDiscountValue.text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
    ElseIf rs("Trans_DiscountType").value = 2 Then
        Me.CboDiscountType.ListIndex = 2
        Me.TxtDiscountValue.text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
    End If

    If IsNull(rs("Trans_DiscountTypePur").value) Then
        Me.CboDiscountTypePur.ListIndex = 0
        Me.TxtDiscountValuePur.text = 0
    ElseIf rs("Trans_DiscountTypePur").value = 0 Then
        Me.CboDiscountTypePur.ListIndex = 0
        Me.TxtDiscountValuePur.text = 0
    ElseIf rs("Trans_DiscountTypePur").value = 1 Then
        Me.CboDiscountTypePur.ListIndex = 1
        Me.TxtDiscountValuePur.text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
    ElseIf rs("Trans_DiscountTypePur").value = 2 Then
        Me.CboDiscountTypePur.ListIndex = 2
        Me.TxtDiscountValuePur.text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
    End If

    Me.DcboCountryID.BoundText = IIf(IsNull(rs("CountryID")), "", rs("CountryID"))
    Me.DcboGovernmentID.BoundText = IIf(IsNull(rs("GovernmentID")), "", rs("GovernmentID"))
    Me.DcboCityID.BoundText = IIf(IsNull(rs("CityID")), "", rs("CityID"))
    Me.TxtAddress.text = IIf(IsNull(rs("Address")), "", Trim(rs("Address")))
    TxtDepitInterval.text = IIf(IsNull(rs("DepitInterval")), 0, rs("DepitInterval"))
    TxtCreditInterval.text = IIf(IsNull(rs("CreditInterval")), 0, rs("CreditInterval"))
    
    dcDepitIntervalID.ListIndex = IIf(IsNull(rs("DepitIntervalID")), -1, rs("DepitIntervalID"))
    dcCreditIntervalID.ListIndex = IIf(IsNull(rs("CreditIntervalID")), -1, rs("CreditIntervalID"))

txtBoxNo.text = IIf(IsNull(rs("BoxNo").value), "", rs("BoxNo").value)
txtPostalCode.text = IIf(IsNull(rs("PostalCode").value), "", rs("PostalCode").value)
txtRSID.text = IIf(IsNull(rs("RSID").value), "", rs("RSID").value)
txtDegree.text = IIf(IsNull(rs("RSDegree").value), "", rs("RSDegree").value)
txtBankAccount.text = IIf(IsNull(rs("BankAccount").value), "", rs("BankAccount").value)
txtBankName.text = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
TXTTOPerson.text = IIf(IsNull(rs("TOPerson").value), "", rs("TOPerson").value)

TxtRecordNo.text = IIf(IsNull(rs("RecordNo").value), "", rs("RecordNo").value)
dtpRsIDDate1.value = IIf(IsNull(rs("RSIDDateH").value), ToHijriDate(Date), rs("RSIDDateH").value)
dtpRsIDDateH.value = IIf(IsNull(rs("RecordDateH").value), ToHijriDate(Date), rs("RecordDateH").value)

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    
       Dim j As Integer
       Dim VendorSQL As String
       
      VendorSQL = VendorSQL & "  SELECT  dbo.TblVendorCars.drivername , dbo.TblVendorCars.DriverTel , dbo.TblVendorCars.EndAllocationDate , dbo.TblVendorCars.ID, dbo.TblVendorCars.Serial, dbo.TblVendorCars.BoardNo, dbo.TblVendorCars.NBoardNo ,  dbo.TblVendorCars.ChasisNo, dbo.TblVendorCars.BrandID, dbo.TblVendorCars.ModelID,"
      VendorSQL = VendorSQL & "  dbo.TblVendorCars.count , dbo.TblVendorCars.CityID, dbo.TblVendorCars.rate, dbo.TblVendorCars.customerid, dbo.TBLCarTypes.name ,dbo.TblVendorCars.accessory ,dbo.TblVendorCars.Price ,dbo.TblVendorCars.PartPrice ,dbo.TblVendorCars.TypeTransID"
      VendorSQL = VendorSQL & "  FROM   dbo.TblVendorCars LEFT OUTER JOIN"
      VendorSQL = VendorSQL & "  dbo.TBLCarTypes ON dbo.TblVendorCars.BrandID = dbo.TBLCarTypes.id"
      VendorSQL = VendorSQL & "  where customerID = " & val(rs("cusID").value)
      
       Set rsVendor = New ADODB.Recordset
       rsVendor.Open VendorSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  '     rsVendor.MoveFirst
       
       
       Grid.rows = 1
       With Grid
       Grid.rows = rsVendor.RecordCount + 1
       For j = 1 To rsVendor.RecordCount
              .TextMatrix(j, .ColIndex("PartPrice")) = IIf(IsNull(rsVendor("PartPrice").value), "", rsVendor("PartPrice").value)
              .TextMatrix(j, .ColIndex("Price")) = IIf(IsNull(rsVendor("Price").value), "", rsVendor("Price").value)
              .TextMatrix(j, .ColIndex("serial")) = IIf(IsNull(rsVendor("serial").value), "", rsVendor("serial").value)
              .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(rsVendor("ID").value), "", rsVendor("ID").value)
              .TextMatrix(j, .ColIndex("BoardNo")) = IIf(IsNull(rsVendor("BoardNo").value), "", rsVendor("BoardNo").value)
              .TextMatrix(j, .ColIndex("nBoardNo")) = IIf(IsNull(rsVendor("nBoardNo").value), "", rsVendor("nBoardNo").value)
              .TextMatrix(j, .ColIndex("ChasisNo")) = IIf(IsNull(rsVendor("ChasisNo").value), "", rsVendor("ChasisNo").value)
              .TextMatrix(j, .ColIndex("BrandID")) = IIf(IsNull(rsVendor("BrandID").value), "", rsVendor("BrandID").value)
              .TextMatrix(j, .ColIndex("ModelID")) = IIf(IsNull(rsVendor("ModelID").value), "", rsVendor("ModelID").value)
              .TextMatrix(j, .ColIndex("Count")) = IIf(IsNull(rsVendor("Count").value), "", rsVendor("Count").value)
              .TextMatrix(j, .ColIndex("CityID")) = IIf(IsNull(rsVendor("CityID").value), "", rsVendor("CityID").value)
              .TextMatrix(j, .ColIndex("Rate")) = IIf(IsNull(rsVendor("Rate").value), "", rsVendor("Rate").value)
           .TextMatrix(j, .ColIndex("Model")) = IIf(IsNull(rsVendor("ModelID").value), "", val(rsVendor("ModelID").value) + 1900)
              .TextMatrix(j, .ColIndex("Brand")) = IIf(IsNull(rsVendor("name").value), "", rsVendor("name").value)
              .TextMatrix(j, .ColIndex("DriverName")) = IIf(IsNull(rsVendor("drivername").value), "", rsVendor("drivername").value)
              .TextMatrix(j, .ColIndex("DriverTel")) = IIf(IsNull(rsVendor("DriverTel").value), "", rsVendor("DriverTel").value)
              .TextMatrix(j, .ColIndex("EndDate")) = IIf(IsNull(rsVendor("EndAllocationDate").value), "", rsVendor("EndAllocationDate").value)
              .TextMatrix(j, .ColIndex("accessory")) = IIf(IsNull(rsVendor("accessory").value), "", rsVendor("accessory").value)
              DcbTypTrans1.ListIndex = IIf(IsNull(rsVendor("TypeTransID").value), -1, rsVendor("TypeTransID").value)
              .TextMatrix(j, .ColIndex("TypeTransID")) = val(DcbTypTrans1.ListIndex)
              .TextMatrix(j, .ColIndex("TypeTrans")) = DcbTypTrans1.text
              
          Dim d As Integer
          Dim str As String
          
        '  d = IIf(IsNull(rsVendor("GovernmentName").value), -1, rsVendor("GovernmentName").value)
        '' If SystemOptions.UserInterface = ArabicInterface Then
        '    If d = 0 Then Str = "—∆Ì”Ï"
        '    If d = 1 Then Str = "»«ÿ‰"
        ' Else
        '    If d = 0 Then Str = "Main"
        '    If d = 1 Then Str = "Inside"
        ' End If
          .TextMatrix(j, .ColIndex("City")) = str
             
                rsVendor.MoveNext
       Next
       End With
    
    
    
    Exit Sub
ErrTrap:
End Sub

Private Sub ShowCusBalance()
    Dim cReport As ClsCustemerReport
    Dim LngCusID As Long
    LngCusID = val(XPTxtComID.text)
    ShowCusBalDailog LngCusID, 0
End Sub


Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.text = ""
If Len(txtNum1.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.text = ""
If Len(txtNum2.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.text = ""
If Len(txtNum3.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus

End If
End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.text = ""
If Len(txtNum4.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
End Sub

Private Sub ntxtNum1_KeyPress(KeyAscii As Integer)
ntxtNum1.text = ""
If Len(ntxtNum1.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum2.SetFocus
End If
End Sub

Private Sub ntxtNum2_KeyPress(KeyAscii As Integer)
ntxtNum2.text = ""
If Len(ntxtNum2.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum3.SetFocus
End If
End Sub

Private Sub ntxtNum3_KeyPress(KeyAscii As Integer)
ntxtNum3.text = ""
If Len(ntxtNum3.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum4.SetFocus

End If
End Sub

Private Sub ntxtNum4_KeyPress(KeyAscii As Integer)
ntxtNum4.text = ""
If Len(ntxtNum4.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
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

Private Sub SaveData()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
Dim mSerial As String
 Dim mTxt As String
    If Me.TxtModFlg.text <> "R" Then
    
    
        If XPTxtComName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Please Enter the supplier name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
            XPTxtComName.SetFocus
            Exit Sub
        End If

        If Me.OptType(2).value = False Then
            If val(Me.TxtOpenBalance.text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ...!!!"
                Else
                    Msg = "balance value must be entered"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If TxtOpenBalance.Enabled = True Then
                    TxtOpenBalance.SetFocus
                End If

                Exit Sub
            End If
        End If

        If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
            Me.TxtDiscountValue.text = 0
        ElseIf Me.CboDiscountType.ListIndex = 1 Then

            If val(Me.TxtDiscountValue.text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„Ê—œ...!!!"
                Else
                    Msg = "Please enter the supplier discount"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValue.SetFocus
                Exit Sub
            End If

        ElseIf Me.CboDiscountType.ListIndex = 2 Then

            If val(Me.TxtDiscountValue.text) = 0 Then
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„Ê—œ...!!!"
                Else
                    Msg = "Please enter the supplier discount"
                End If
                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValue.SetFocus
                Exit Sub
            ElseIf val(Me.TxtDiscountValue.text) > 100 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
                Else
                    Msg = "Discount value must be bigger than 100"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValue.SetFocus
                Exit Sub
            End If
        End If

        If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
            Me.TxtDiscountValuePur.text = 0
        ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then

            If val(Me.TxtDiscountValuePur.text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„Ê—œ ›Ï ›Ê« Ì— «·‘—«¡...!!!"
                Else
                    Msg = "Discount value for supplier must be written in purchase invoices"
                    End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            End If

        ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then

            If val(Me.TxtDiscountValuePur.text) = 0 Then
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… »«·„Ê—œ ›Ï ›Ê« Ì— «·‘—«¡..!!!"
                Else
                    Msg = "Discount value for supplier must be written in purchase invoices"
                End If
                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            ElseIf val(Me.TxtDiscountValuePur.text) > 100 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
                Else
                    Msg = "Discount value must be bigger than 100"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            End If
        End If

        Select Case Me.TxtModFlg.text

            Case "N"
                XPTxtComID.text = CStr(new_id("TblCustemers", "CusID", "", True))
                StrSQL = "select * From  TblCustemers where Type=2 AND CusName='" & Trim(XPTxtComName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ „Ê—œ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ê—œ"
                    Else
                        Msg = "The a supplier with same name ... please make sure the name is correct or enter a different name"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtComName.SetFocus
                    Exit Sub
                End If

            
                RsTemp.Close
                StrSQL = "select * From  TblCustemers where   Type=2 AND fullcode='" & DCPreFix.text & txtid.text & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ „Ê—œ „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ—ﬂÊœ «·„Ê—œ"
                    Else
                        Msg = "The a supplier with same code ... please make sure the code is correct or enter a different code"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtComName.SetFocus
                    Exit Sub
                End If
                
                RsTemp.Close
                StrSQL = "select * From  TblCustemers where   Type=2 AND recordno ='" & TxtRecordNo.text & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 And TxtRecordNo.text <> "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ „Ê—œ „”Ã· „”»ﬁ« »—ﬁ„ «·”Ã· «·„œŒ·" & CHR(13)
                        Msg = Msg + "Â·  —Ìœ «·«” „—«— ø " & CHR(13)
                    Else
                        Msg = "There's a supplier with the same record number ... Do you want to continue"
                    End If
                   If MsgBox(Msg, vbOKCancel) = vbOK Then
                    
                   Else
                   Exit Sub
                   End If
                 
                End If
            
xx:
            
            Case "E"
                StrSQL = "select * From  TblCustemers where Type=2 AND CusName='" & Trim(XPTxtComName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("CusID").value <> val(XPTxtComID.text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ „Ê—œ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ê—œ"
                    Else
                        Msg = "The a supplier with same name ... please make sure the name is correct or enter a different name"
                    End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtComName.SetFocus
                        Exit Sub
                    End If
                End If
            
                RsTemp.Close

           
             
                StrSQL = "select * From  TblCustemers where Type=2 AND fullcode='" & DCPreFix.text & txtid.text & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("CusID").value <> val(XPTxtComID.text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÊÃœ „Ê—œ „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ—ﬂÊœ «·„Ê—œ"
                    Else
                        Msg = "The a supplier with same code ... please make sure the code is correct or enter a different code"
                    End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtComName.SetFocus
                        Exit Sub
                    End If
                End If
                
                
                RsTemp.Close
                StrSQL = "select * From  TblCustemers where   Type=2 AND recordno ='" & TxtRecordNo.text & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If RsTemp("CusID").value <> val(XPTxtComID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÊÃœ „Ê—œ „”Ã· „”»ﬁ« »—ﬁ„ «·”Ã· «·„œŒ·" & CHR(13)
                            Msg = Msg + "Â·  —Ìœ «·«” „—«— ø " & CHR(13)
                        Else
                            Msg = "There's a supplier with the same record number ... Do you want to continue"
                        End If
                                If MsgBox(Msg, vbOKCancel) = vbOK Then
                                 
                                Else
                                Exit Sub
                                End If
                    End If
                End If
            
ll:
        End Select

        Cn.BeginTrans
        BeginTrans = True

        If Me.TxtModFlg.text = "N" Then
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = Me.DboParentAccount.BoundText
            rs.AddNew
            rs("CusID").value = val(XPTxtComID.text)
       
        
        ElseIf Me.TxtModFlg.text = "E" Then
            '  StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtComID.text)
            '  Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            
        If SystemOptions.IsCreateOpenBalnceMan = True Then
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            Dim Account_code As String
            Account_code = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                If Account_code <> "" Then
                    StrSQL = " delete DOUBLE_ENTREY_VOUCHERS1"
                    StrSQL = StrSQL & " where  opening_balance_voucher_id in"
                    StrSQL = StrSQL & " ("
                    StrSQL = StrSQL & " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id"
                    StrSQL = StrSQL & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 INNER JOIN"
                    StrSQL = StrSQL & "                       dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code = dbo.ACCOUNTS.Account_Code"
                    StrSQL = StrSQL & " WHERE     ( Notes_ID=1 and dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code =  '" & Account_code & "')"
                    StrSQL = StrSQL & " )"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            End If
       '     strSQL = "delete From tblvendorcars where customerID=" & val(XPTxtComID.text)
       '     Cn.Execute strSQL, , adExecuteNoRecords
        
        End If
         rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
         rs("CurrncyID").value = IIf(Me.DcbCurrency.BoundText = "", 0, val(DcbCurrency.BoundText))
        rs("code").value = txtid.text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)

     rs("EmpId").value = IIf(Me.DCEmP.BoundText = "", Null, (Me.DCEmP.BoundText))

        rs("c1").value = Me.TxtCode.text
        rs("BankCode").value = Trim(TxtBankCode.text)
        rs("BankIBAN").value = Trim(TxtBankIBAN.text)
        rs("BankAddress").value = Trim(TxtBankAddress.text)
        rs("CusName").value = Trim(XPTxtComName.text)
        rs("VATNO").value = TxtVATNO.text
        rs("Cus_Phone").value = IIf(XPTxtPhone.text = "", "", Trim(XPTxtPhone.text))
        rs("Cus_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
        rs("Remark").value = IIf(XPMTxtRemark.text = "", "", Trim(XPMTxtRemark.text))
        rs("Remark2").value = IIf(XPMTxtRemarks2.text = "", "", Trim(XPMTxtRemarks2.text))
        rs("parent_account").value = IIf(Me.DboParentAccount.BoundText = "", Null, Me.DboParentAccount.BoundText)
        rs("Category").value = IIf(cbCategory.ListIndex = -1, Null, cbCategory.ListIndex)
           If chkTaxExempt.value = vbChecked Then
        rs("chkTaxExempt").value = 1
    Else
        rs("chkTaxExempt").value = 0
    End If
    
        If locked.value = vbChecked Then
            rs("locked").value = 1
        Else
            rs("locked").value = 0
        End If
    If Trim(XPTxtCusNamee.text) = "" Then XPTxtCusNamee.text = Trim(XPTxtComName)
    
        rs("CusNamee").value = Trim(XPTxtCusNamee.text)

        If chkCustomerandVendor.value = vbChecked Then
            rs("CustomerandVendor").value = 1

        Else
            rs("CustomerandVendor").value = 0
        End If


 
    If chkIsBranch(4).value = vbChecked Then
        rs("IsHiddenInv").value = 1
    Else
        rs("IsHiddenInv").value = 0
    End If
    

        If CHkMot3ahed.value = vbChecked Then
            rs("CHkMot3ahed").value = 1

        Else
            rs("CHkMot3ahed").value = 0
        End If
        
        
        rs("Type").value = 2

        If Me.OptType(2).value = True Then
            rs("OpenBalance").value = 0
            rs("OpenBalanceType").value = Null
        ElseIf Me.OptType(0).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
            rs("OpenBalanceType").value = 0
        ElseIf Me.OptType(1).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
            rs("OpenBalanceType").value = 1
        End If

        rs("OpenBalanceDate").value = Me.Dtp.value
        rs("CreditLimit").value = val(Me.TxtCreditLimit.text)
        rs("CreditlimitCredit").value = val(Me.TxtCreditlimitCredit.text)
        rs("FaxNumber").value = IIf(Trim$(Me.TxtFaxNumber.text) = "", Null, Trim$(Me.TxtFaxNumber.text))
        rs("E_mail").value = IIf(Trim$(Me.TxtE_mail.text) = "", Null, Trim$(Me.TxtE_mail.text))

        If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
            rs("Trans_DiscountType").value = 0
            rs("Trans_Discount").value = 0
        ElseIf Me.CboDiscountType.ListIndex = 1 Then
            rs("Trans_DiscountType").value = 1
            rs("Trans_Discount").value = val(Me.TxtDiscountValue.text)
        ElseIf Me.CboDiscountType.ListIndex = 2 Then
            rs("Trans_DiscountType").value = 2
            rs("Trans_Discount").value = val(Me.TxtDiscountValue.text)
        End If

        If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
            rs("Trans_DiscountTypePur").value = 0
            rs("Trans_DiscountPur").value = 0
        ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then
            rs("Trans_DiscountTypePur").value = 1
            rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.text)
        ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then
            rs("Trans_DiscountTypePur").value = 2
            rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.text)
        End If



                                             
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            If Me.TxtModFlg.text = "N" Then
                If SystemOptions.CustCreat4Acc = True Then
                      
                      
                      mTxt = ExtractCharacter(DCPreFix.text)
                      If mTxt = "" Then
                        mTxt = get_account_code_branch(218, my_branch, "T")  ' Account_Code_dynamic
                      End If
                      mSerial = GET_ACCOUNT_name_by_Code(DboParentAccount.BoundText, "T") & mTxt & txtid
                      
                      '    rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccountCurrentAss, Trim$(Me.XPTxtCusName.text), True, False, Trim$(Me.XPTxtCusNamee.text))
                      
                     ' rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text) & " Ã«—Ì «·⁄„· ", True, False, XPTxtCusNamee.text & " payable ", , , , , , mSerial)
        
                  
                  
                      rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , mSerial, , , , 1, 1, 1, 0, 0)
                Else
                    
                    rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , mTxt, , , , 1, 1, 1, 0, 0)
                End If
                'Rs("Account_Code").value = ModAccounts.AddNewAccount("a2a3a1", Trim$(Me.XPTxtComName.text), True, False)
            Else

                        If Not IsNull(rs("Account_Code").value) Then
                            If SystemOptions.CustCreat4Acc = True Then
                                s = "Select Account_Serial from accounts where account_Code = N'" & Trim(rs("Account_Code") & "") & "'"
                                Set dummy = New ADODB.Recordset
                                dummy.Open s, Cn, adOpenKeyset, adLockOptimistic
                                If Not dummy.EOF Then
                                    If ExtractCharacter(DCPreFix.text) <> ExtractCharacter(dummy!account_serial & "") Then
                                         mTxt = ExtractCharacter(DCPreFix.text)
                                         mSerial = GET_ACCOUNT_name_by_Code(DboParentAccount.BoundText, "T") & mTxt & txtid
                                         dummy!account_serial = mSerial
                                         dummy.update
                                         
                                    End If
                                End If
                                
                                
                            End If
                            ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtComName.text, Me.XPTxtCusNamee.text, , , , , , , , , 1, 1, 1, 0, 0, , , , True
                        Else
                            If SystemOptions.CustCreat4Acc = True Then
                                If mTxt = "" Then
                                    mTxt = get_account_code_branch(218, my_branch, "T")  ' Account_Code_dynamic
                                End If
                                mSerial = GET_ACCOUNT_name_by_Code(DboParentAccount.BoundText, "T") & mTxt & txtid
                                rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , mSerial, , , , 1, 1, 1, 0, 0)
                            Else
                                rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , mTxt, , , , 1, 1, 1, 0, 0)
                            End If
                        End If
            End If
        End If

        rs("CountryID").value = IIf(val(Me.DcboCountryID.BoundText) = 0, Null, val(Me.DcboCountryID.BoundText))
        rs("GovernmentID").value = IIf(val(Me.DcboGovernmentID.BoundText) = 0, Null, val(Me.DcboGovernmentID.BoundText))
        rs("CityID").value = IIf(val(Me.DcboCityID.BoundText) = 0, Null, val(Me.DcboCityID.BoundText))
        rs("ResponsibleContact").value = Trim$(Me.TxtResponsibleContact.text)
         rs("Address").value = Trim$(Me.TxtAddress.text)
        rs("CustomerTypeID").value = IIf(val(Me.DcCustomerType.BoundText) = 0, Null, val(Me.DcCustomerType.BoundText))
        rs("DepitInterval").value = val(TxtDepitInterval.text)
        rs("CreditInterval").value = val(TxtCreditInterval.text)
        rs("DepitIntervalID").value = val(dcDepitIntervalID.ListIndex)
        rs("CreditIntervalID").value = val(dcCreditIntervalID.ListIndex)
                
    rs("BoxNo").value = IIf(txtBoxNo.text = "", "", Trim(txtBoxNo.text))
    rs("PostalCode").value = IIf(txtPostalCode.text = "", "", Trim(txtPostalCode.text))
    rs("RSID").value = IIf(txtRSID.text = "", "", Trim(txtRSID.text))
    rs("RSDegree").value = IIf(txtDegree.text = "", "", Trim(txtDegree.text))
    rs("BankAccount").value = IIf(txtBankAccount.text = "", "", Trim(txtBankAccount.text))
    rs("BankName").value = IIf(txtBankName.text = "", "", Trim(txtBankName.text))
     rs("TOPerson").value = IIf(TXTTOPerson.text = "", "", Trim(TXTTOPerson.text))
     
    rs("RecordNo").value = IIf(TxtRecordNo.text = "", "", Trim(TxtRecordNo.text))
    rs("RSIDDateH").value = dtpRsIDDateH.value
    rs("RecordDateH").value = dtpRsIDDateH.value
    
    
         rs("StreetName").value = txtNoOFDigitUser(2).text
        rs("BuildingNumber").value = txtNoOFDigitUser(4).text
         rs("CitySubdivisionName").value = DcboCityID.text
          rs("CityName").value = DcboGovernmentID.text
           rs("PostalZone").value = txtPostalCode.text
            rs("IdentificationCode").value = txtNoOFDigitUser(10).text
             rs("PlotIdentification").value = txtNoOFDigitUser(5).text
              rs("AdditionalStreetName").value = txtNoOFDigitUser(3).text
              rs("CountrySubentity").value = txtNoOFDigitUser(8).text
              rs("Id700").value = txtNoOFDigitUser(0).text
       If val(TxtOpenBalance.text) = 0 Then
            txtopening_balance_voucher_id = 0
        End If
       
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
       
       '     If val(Me.txtopening_balance_voucher_id.text) = 0 Then
                txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
               rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
       '     End If '
        End If '
        
        
        rs.update
    
        Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtComName.text) & " "
        Else
            StrDes = " Opening Balance For: " & Trim(Me.XPTxtCusNamee.text) & " "
        End If
        If SystemOptions.IsCreateOpenBalnceMan = True Then
                If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
                    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                
                        Dim LngDevID As Long
                        Dim LngOpenID As Long
                        Dim Account_Code_dynamic1 As String
                
                        ' LngOpenID = ModAccounts.AddNewOpenBalance(Val(Me.XPTxtComID.text), Me.Dtp.value)
                        ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        LngOpenID = 1
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
                   
                        If Me.OptType(0).value = True Then
                            Account_Code_dynamic1 = get_account_code_branch(60, my_branch)
                
                            If Account_Code_dynamic1 = "NO branch" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                                Else
                                    MsgBox "Branch has not been created ", vbCritical
                                End If
                                GoTo ErrTrap
                            Else
                    
                                If Account_Code_dynamic1 = "NO account" Then
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                    Else
                                        MsgBox "An opening balance account was not selected for this branch for this process", vbCritical
                                    End If
                                    GoTo ErrTrap
                             
                                End If
                            End If
                
                            If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , IIf(chkIsBranch(4).value = vbChecked, 6, 0)) = False Then
                                GoTo ErrTrap
                            End If
                        
                            If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , IIf(chkIsBranch(4).value = vbChecked, 6, 0)) = False Then
                                GoTo ErrTrap
                            End If
                        
                            ' If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                            '     Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtComName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                            '        GoTo ErrTrap
                            ' End If
                        
                        ElseIf Me.OptType(1).value = True Then
                    
                            Account_Code_dynamic1 = get_account_code_branch(60, my_branch)
                         
                            If Account_Code_dynamic1 = "NO branch" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                                Else
                                    MsgBox "Branch has not been created ", vbCritical
                                End If
                                GoTo ErrTrap
                            Else
        
                                If Account_Code_dynamic1 = "NO account" Then
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                    Else
                                        MsgBox "An opening balance account was not selected for this branch for this process", vbCritical
                                    End If
                                    GoTo ErrTrap
                          
                                End If
                            End If
                         
                            If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , IIf(chkIsBranch(4).value = vbChecked, 6, 0)) = False Then
                                GoTo ErrTrap
                            End If
                        
                            'If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                            '    Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtComName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                            '        GoTo ErrTrap
                            'End If
                        
                            If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , IIf(chkIsBranch(4).value = vbChecked, 6, 0)) = False Then
                                GoTo ErrTrap
                            End If
                        End If
        
                        '         update_account_opening_balance rs("Account_Code").value
                        '     update_account_opening_balance Account_Code_dynamic1
                         
                    End If
                End If
            End If
        
       Dim j As Integer
       
  '     Dim rsVendor As New ADODB.Recordset
  '     rsVendor.Open "tblvendorcars", Cn, adOpenStatic, adLockPessimistic, adCmdTable
 
'      Dim rsVendor As ADODB.Recordset
'      Set rsVendor = New ADODB.Recordset
'      rsVendor.Open "tblvendorcars", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'       With Grid
'       For j = 1 To Grid.Rows - 1
'            If .TextMatrix(j, .ColIndex("BrandID")) <> "" Then
'                rsVendor.AddNew
'               rsVendor("ID").value = CStr(new_id("tblvendorcars", "ID", "", True))
'
'            '  Dim str as
'             ' str = IIf(.TextMatrix(j, .ColIndex("Serial")), "", .TextMatrix(j, .ColIndex("Serial")))
'
'                rsVendor("Serial").value = IIf(.TextMatrix(j, .ColIndex("serial")) = "", Null, .TextMatrix(j, .ColIndex("serial")))
'                rsVendor("BoardNo").value = IIf(.TextMatrix(j, .ColIndex("BoardNo")) = "", "", .TextMatrix(j, .ColIndex("BoardNo")))
'                rsVendor("ChasisNo").value = IIf(.TextMatrix(j, .ColIndex("ChasisNo")) = "", Null, .TextMatrix(j, .ColIndex("ChasisNo")))
'                rsVendor("BrandID").value = IIf(.TextMatrix(j, .ColIndex("BrandID")) = "", Null, .TextMatrix(j, .ColIndex("BrandID")))
'                rsVendor("ModelID").value = IIf(.TextMatrix(j, .ColIndex("ModelID")) = "", Null, .TextMatrix(j, .ColIndex("ModelID")))
'                rsVendor("Count").value = IIf(.TextMatrix(j, .ColIndex("Count")) = "", 0, .TextMatrix(j, .ColIndex("Count")))
'                rsVendor("CityID").value = IIf(.TextMatrix(j, .ColIndex("CityID")) = "", Null, .TextMatrix(j, .ColIndex("CityID")))
'                rsVendor("Rate").value = IIf(.TextMatrix(j, .ColIndex("Rate")) = "", 0, .TextMatrix(j, .ColIndex("Rate")))
'                rsVendor("CustomerID").value = val(XPTxtComID.text)
'
'               rsVendor("DriverName").value = IIf(.TextMatrix(j, .ColIndex("DriverName")) = "", Null, .TextMatrix(j, .ColIndex("DriverName")))
'               rsVendor("DriverTel").value = IIf(.TextMatrix(j, .ColIndex("DriverTel")) = "", Null, .TextMatrix(j, .ColIndex("DriverTel")))
'               'rsVendor("EndAllocationDate").value = IIf(.TextMatrix(j, .ColIndex("EndDate")) = "", Date, .TextMatrix(j, .ColIndex("EndDate")))
''
'                rsVendor.update
'            End If
'       Next
'       End With
       
       
        
        '////////////////////////////////////////////////////////////
        Dim rsVendor As ADODB.Recordset
        Set rsVendor = New ADODB.Recordset
        StrSQL = "select * from tblvendorcars  order by id "
        rsVendor.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        With Grid
     
       ' FgInstallments.Rows = val(TxtPaymentCount.text) + 1
        Dim AllID As String
    '    rsVendor.MoveFirst
       
        For j = Grid.FixedRows To Grid.rows - 1
           If .TextMatrix(j, .ColIndex("BoardNo")) <> "" Then
           
                    If .TextMatrix(j, .ColIndex("ID")) = "" Then
                            rsVendor.AddNew
                            rsVendor("ID") = CStr(new_id("tblvendorcars", "ID", "", True))
                    Else
                            rsVendor.Find " ID ='" & val(.TextMatrix(j, .ColIndex("ID"))) & "'", , adSearchForward, adBookmarkFirst
                            
                            If rsVendor.EOF Or rsVendor.BOF Then
                                    Exit Sub
                            End If
                    End If
                     rsVendor("PartPrice").value = IIf(.TextMatrix(j, .ColIndex("PartPrice")) = "", Null, val(.TextMatrix(j, .ColIndex("PartPrice"))))
                     rsVendor("Price").value = IIf(.TextMatrix(j, .ColIndex("Price")) = "", Null, val(.TextMatrix(j, .ColIndex("Price"))))
                     rsVendor("Serial").value = IIf(.TextMatrix(j, .ColIndex("serial")) = "", Null, .TextMatrix(j, .ColIndex("serial")))
                     rsVendor("BoardNo").value = IIf(.TextMatrix(j, .ColIndex("BoardNo")) = "", "", .TextMatrix(j, .ColIndex("BoardNo")))
                     rsVendor("nBoardNo").value = IIf(.TextMatrix(j, .ColIndex("nBoardNo")) = "", "", .TextMatrix(j, .ColIndex("nBoardNo")))
                     rsVendor("ChasisNo").value = IIf(.TextMatrix(j, .ColIndex("ChasisNo")) = "", Null, .TextMatrix(j, .ColIndex("ChasisNo")))
                     rsVendor("BrandID").value = IIf(.TextMatrix(j, .ColIndex("BrandID")) = "", Null, .TextMatrix(j, .ColIndex("BrandID")))
                      rsVendor("ModelID").value = IIf(.TextMatrix(j, .ColIndex("ModelID")) = "", Null, .TextMatrix(j, .ColIndex("ModelID")))
                     rsVendor("Count").value = IIf(.TextMatrix(j, .ColIndex("Count")) = "", 0, .TextMatrix(j, .ColIndex("Count")))
                     rsVendor("CityID").value = IIf(.TextMatrix(j, .ColIndex("CityID")) = "", Null, .TextMatrix(j, .ColIndex("CityID")))
                     rsVendor("Rate").value = IIf(.TextMatrix(j, .ColIndex("Rate")) = "", 0, .TextMatrix(j, .ColIndex("Rate")))
                     rsVendor("CustomerID").value = val(XPTxtComID.text)
                     rsVendor("DriverName").value = IIf(.TextMatrix(j, .ColIndex("DriverName")) = "", Null, .TextMatrix(j, .ColIndex("DriverName")))
                     rsVendor("DriverTel").value = IIf(.TextMatrix(j, .ColIndex("DriverTel")) = "", Null, .TextMatrix(j, .ColIndex("DriverTel")))
                     rsVendor("accessory").value = IIf(.TextMatrix(j, .ColIndex("accessory")) = "", Null, .TextMatrix(j, .ColIndex("accessory")))
                     rsVendor("TypeTransID").value = IIf(.TextMatrix(j, .ColIndex("TypeTransID")) = "", -1, val(.TextMatrix(j, .ColIndex("TypeTransID"))))
                     
                     rsVendor.update
                    
                If j = Grid.FixedRows Then
                    AllID = rsVendor("ID").value
                Else
                    AllID = AllID & "  ,  " & CStr(rsVendor("ID").value)
                End If
                    
            End If
           Next
        End With
        
        
         'Dim strSQL As String
         If AllID <> "" Then
                StrSQL = "delete from tblvendorcars  where customerid = " & val(XPTxtComID.text) & " and  id not in  ( " & AllID & "  ) "
                 Cn.Execute StrSQL, , adExecuteNoRecords
         End If
                     '//////////////////////////////////////////////////////////////////
        
        
        
        
        
        
        
        

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'update_account_opening_balance Me.DcboDebitSide.BoundText
        'update_account_opening_balance Me.DcboCreditSide.BoundText
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
        
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„Ê—œ" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Done, do you want new supplier"
                End If
            
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox " Update Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If

        End Select
ReloadCompo
        TxtModFlg.text = "R"
        Retrive (val(XPTxtComID.text))
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "CusID='" & val(XPTxtComID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Company()
    Dim Msg As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtComID.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·„Ê—œ   " & CHR(13)
            Msg = Msg + (XPTxtComName.text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Supplier data will be deleted" & CHR(13)
            Msg = Msg + (XPTxtComName.text) & CHR(13)
            Msg = Msg + "do you want to delete data ?"
        End If
        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
    
        If IntRes = vbYes Then
            If Not rs.RecordCount < 1 Then
                DeleteOpeningBalance
                Cn.BeginTrans
                BegainTrans = True
                'StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtComID.text)
                'Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                Dim Account_code As String
    Account_code = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
    If Account_code <> "" Then
StrSQL = " delete DOUBLE_ENTREY_VOUCHERS1"
StrSQL = StrSQL & " where  opening_balance_voucher_id in"
StrSQL = StrSQL & " ("
StrSQL = StrSQL & " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id"
StrSQL = StrSQL & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 INNER JOIN"
StrSQL = StrSQL & "                       dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL & " WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code =  '" & Account_code & "')"
StrSQL = StrSQL & " )"
    Cn.Execute StrSQL, , adExecuteNoRecords
 End If
 
                '   update_account_opening_balance get_account_code_branch(19, my_branch)
             
                Dim StrAccountCode As String
                StrAccountCode = rs("Account_Code").value
                '     If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                '         rs.delete
                '     Else
                '         Exit Sub
                '     End If
            
                StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords

                CuurentLogdata ("D")
                rs.delete
                
                StrSQL = "delete From tblvendorcars where customerID=" & val(XPTxtComID.text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                
                Cn.CommitTrans
                BegainTrans = False
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " „  ⁄„·Ì… «·Õ–›."
                Else
                    Msg = "Recored deleted successfully"
                End If
                
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPBtnMove_Click 2

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
            Msg = "sorry, this record cannot be deleted due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
    Else
        Msg = "sorry, this record cannot be deleted due to data integration"
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate

    If BegainTrans = True Then
        Cn.RollbackTrans
        BegainTrans = False
    End If

    'End If
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „Ê—œ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„Ê—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„Ê—œ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  „Ê—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „Ê—œ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ê—œÌ‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New Supplier Data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current Supplier data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit the current Supplier data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the current editing or Save the new Supplier data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the adding new record" & Wrap & "OR undo editing current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current Supplier data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search" & Wrap & "Search for a Supplier..." & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Suppliers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Show Help File", BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

'Private Sub printingReport()
'    Dim CusReport As ClsCustemerReport
'    On Error GoTo ErrTrap
'
'    If XPTxtComID.text <> "" Then
'        Set CusReport = New ClsCustemerReport
'        CusReport.CustemerData XPTxtComID.text, 2
'    End If
'
'    Exit Sub
'ErrTrap:

    'On Error GoTo ErrTrap
    'If XPTxtComID.text <> "" Then
    '    Set ComReport = New ClsCompanyReport
    '    ComReport.CompanyData XPTxtComID.text, 2
    'End If
    'Exit Sub
    'ErrTrap:
'End Sub

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

Private Sub ChangeLang()

    Dim XPic As IPictureDisp
    
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
     XPLbl(18).Caption = "Type"
    lbl(40).Caption = "VAT No."
    Label3.Caption = "Branch"

    With CboDiscountType
        .Clear
        .AddItem "No"
        .AddItem "Value"
        .AddItem "percentage"
    End With
    
    XPLbl(14).Caption = "Currency"
    
    With CboDiscountTypePur
        .Clear
        .AddItem "no"
        .AddItem "Value"
        .AddItem "percentage"
    End With

    lbl(23).Caption = "Contact person"
    lbl(19).Caption = " type"
    lbl(20).Caption = "Value"
    Cmd(12).Caption = "Old Bill"
    lbl(29).Caption = " type"
    lbl(28).Caption = "Value"
    lbl(22).Caption = "State"
    lbl(24).Caption = "Province"
    lbl(25).Caption = " City "
    lbl(26).Caption = "Address"
    Fra(5).Caption = "Work Address"
    Fra(4).Caption = "Discounts sales invoices"
    Fra(6).Caption = "Discounts purchase invoices"
    lbl(36).Caption = "IBAN"
    lbl(37).Caption = "Bank Code"
    lbl(38).Caption = "Bank Address"
    lbl(16).Caption = "Banck Account"
    lbl(17).Caption = "Record Date"
    XPLbl(1).Caption = "Record No"
    XPLbl(13).Caption = "Category"
    lbl(18).Caption = "ID No"
    lbl(34).Caption = "Degree"
    Frame3.Caption = "Data of Person Responsible "
    lbl(14).Caption = "Zip Code"
    lbl(35).Caption = "Issue Date"
    Me.Caption = "Suppliers Data"
    EleHeader.Caption = Me.Caption
    lbl(13).Caption = "P.O.b."
    XPLbl(2).Caption = "Code"
    lbl(15).Caption = "Banck Name"
    XPLbl(0).Caption = "Supplier Name"
    XPLbl(4).Caption = "English Name"
    lbl(3).Caption = "Phone"
    lbl(2).Caption = "Mobile"
    lbl(1).Caption = "Remarks"
    lbl(0).Caption = "Current Record"
    lbl(7).Caption = "Fax NO."
    lbl(10).Caption = "Credit Limit(Debit)"
    lbl(11).Caption = "Credit Limit(Credit)"
    lbl(12).Caption = "E-Mail."
    lbl(33).Caption = "Parent Acc"
    Me.Fra(1).Caption = "Open Balance"
    Me.Fra(0).Caption = "Open Balance State"
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    lbl(5).Caption = "Balance Value"
    lbl(6).Caption = "Record Date"
    Fra(3).Caption = "Contact Info."
    chkCustomerandVendor.Caption = "Customer & Supplier"
    CHkMot3ahed.Caption = "CHkMot3ahed"
    Label1(2).Caption = "Type"
    Me.Fra(2).Caption = "Current Balance State"
    Me.Cmd(8).Caption = "Customer Balance Report"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    Me.CmdPriceList.Caption = "Supplier Price List"
    locked.Caption = "locked"
    ALLButton1.Caption = "Reason"
    lbl(32).Caption = "reason"
    lbl(30).Caption = "period"
    lbl(31).Caption = "period"
    Frame1.Caption = "Cars Contractors"
    XPLbl(5).Caption = "Chasis No"
    XPLbl(6).Caption = "Model"
    lbl(5).Caption = "Brand"
    XPLbl(9).Caption = "For"
    XPLbl(7).Caption = "Set Count"
    XPLbl(8).Caption = "Pass. Rate"
    Frame6.Caption = "Board No."
    XPLbl(3).Caption = "Exp."
    XPLbl(10).Caption = "A B C 1 2 3"
    Cmd(14).Caption = "Add"

    With Grid
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("BoardNo")) = "Board No"
        .TextMatrix(0, .ColIndex("nBoardNo")) = "Board No E"
        .TextMatrix(0, .ColIndex("ChasisNo")) = "Chasis No"
        .TextMatrix(0, .ColIndex("Brand")) = "Brand"
        .TextMatrix(0, .ColIndex("Model")) = "Model"
        .TextMatrix(0, .ColIndex("count")) = "Set Count"
        .TextMatrix(0, .ColIndex("Rate")) = "Pass. Rate"
        .TextMatrix(0, .ColIndex("City")) = "For"
        
        .TextMatrix(0, .ColIndex("DriverName")) = "Driver Name"
        .TextMatrix(0, .ColIndex("DriverTel")) = "Driver Phone"
        .TextMatrix(0, .ColIndex("EndDate")) = "License End Date"
        .TextMatrix(0, .ColIndex("accessory")) = "Accessory"
        .TextMatrix(0, .ColIndex("Price")) = "Value"
        .TextMatrix(0, .ColIndex("TypeTrans")) = "Type"
        
    End With

    XPLbl(15).Caption = "Accessory"
    XPLbl(17).Caption = "Accessory Value"
    XPLbl(16).Caption = "Value"
    XPLbl(11).Caption = "Driver Name"
    Cmd(9).Caption = "Delete"
    Cmd(10).Caption = "Delete All"
    C1Tab1.Caption = "Cars Contractors |  Main Data "
    c(0).Caption = "account No."
    c(5).Caption = "Value"
    Cmd(11).Caption = "attachment"

End Sub

Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
ReloadCompo

    Set Dcombo = New ClsDataCombos

    If BolExceptCountries = False Then
        Dcombo.GetCountriesNames Me.DcboCountryID
        Set cSearch(0) = New clsDCboSearch
        Set cSearch(0).Client = Me.DcboCountryID
    End If

    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID.BoundText)
        Set cSearch(1) = New clsDCboSearch
        Set cSearch(1).Client = Me.DcboGovernmentID
    End If

    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID.BoundText), val(Me.DcboGovernmentID.BoundText)
        Set cSearch(2) = New clsDCboSearch
        Set cSearch(2).Client = Me.DcboCityID
    End If

    Dcombo.GetCustomerType Me.DcCustomerType
    Dcombo.GetBranches dcBranch
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

End Sub

Private Sub XPTxtComName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtCusNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub


 Public Function GetACCOUNTSCode(LngItemID As String, Optional ID As Integer = 0) As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
 
    If LngItemID <> "" Then
   If ID = 1 Then
       StrSQL = "Select Account_Serial  From ACCOUNTS Where Account_Code='" & LngItemID & "'"
     Else
     StrSQL = "Select Account_Code  From ACCOUNTS Where Account_Serial='" & LngItemID & "'"
     End If
        Set rs = New ADODB.Recordset
 
        If Cn.State = adStateClosed Then
            open_my_connection
        End If
 
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
        If Not (rs.BOF Or rs.EOF) Then
        If ID = 1 Then
            GetACCOUNTSCode = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            Else
            GetACCOUNTSCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
         End If
        Else
        End If
 
        rs.Close
        Set rs = Nothing
    End If
 
End Function


Function CREATEADDRESS()
If SystemOptions.IsBluee = True Then
TxtAddress = txtNoOFDigitUser(4) & " " & txtNoOFDigitUser(2) & " " & DcboCityID.text & " " & DcboGovernmentID.text & " " & DcboCountryID.text & " " & "«·—„“ «·»—ÌœÌ" & txtPostalCode
End If
End Function

Private Sub txtNoOFDigitUser_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Or Index = 10 Then Exit Sub
   KeyAscii = KeyAscii_Num(KeyAscii, txtNoOFDigitUser(4).text, 0)
End Sub




Function checkEeinvoice() As Boolean
If Trim(TxtRecordNo.text) = "" Then checkEeinvoice = True: Exit Function

  If Not SystemOptions.ApplyEinvoice Then checkEeinvoice = True: Exit Function
  If chkTaxExempt.value = Checked Then checkEeinvoice = True: Exit Function
'  If creditlocked.value = Checked Then checkEeinvoice = True: Exit Function
checkEeinvoice = False

If TxtRecordNo.text = "" And Trim(txtNoOFDigitUser(0).text) = "" Then

    
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "—ﬁ„ «·”Ã· «·“«„Ì", vbCritical
                Else
                MsgBox "enter CRN ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If Not SystemOptions.CustVatNoMandatory Then
    If (TxtVATNO.text = "" Or Len(TxtVATNO) < 15 Or mId(TxtVATNO, 15, 1) <> 3) And Trim(txtNoOFDigitUser(0)) = "" Then
          If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "  «·—ﬁ„ «·÷—ÌÌ 15 Œ«‰…  «·“«„Ì ÊÌ‰ ÂÌ »«·—ﬁ„ 3", vbCritical
                    Else
                    MsgBox "Vat No 15 digit ", vbCritical
           End If
            checkEeinvoice = False
            Exit Function
    End If
    
    
'     If (TxtVATNO.text = "" Or Len(TxtVATNO) < 15 Or mId(TxtVATNO, 15, 1) <> 3) And Trim(txtNoOFDigitUser(0)) = "" Then
'          If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "  «·—ﬁ„ «·÷—ÌÌ 15 Œ«‰…  «·“«„Ì ÊÌ‰ ÂÌ »«·—ﬁ„ 3", vbCritical
'                    Else
'                    MsgBox "Vat No 15 digit ", vbCritical
'           End If
'            checkEeinvoice = False
'            Exit Function
'    End If
End If


If txtNoOFDigitUser(4).text = "" Or Len(txtNoOFDigitUser(4)) < 4 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     —ﬁ„ «·„»‰Ì 4 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "bulding no 4 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If txtPostalCode.text = "" Or Len(txtPostalCode) < 5 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·—„“ «·»—ÌœÌ   5 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "Zib no 5 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If txtNoOFDigitUser(2).text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "    «”„ «·‘«—⁄  «·“«„Ì", vbCritical
                Else
                MsgBox "enter street name ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If txtNoOFDigitUser(10).text = "" Or Len(txtNoOFDigitUser(10)) < 2 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "   ﬂÊœ «·œÊ·…  «·“«„Ì 2 Œ«‰…", vbCritical
                Else
                MsgBox "must enter country code Code ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If val(DcboCountryID.BoundText) = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·œÊ·…  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter country  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If val(DcboGovernmentID.BoundText) = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·„œÌ‰…  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter city  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If val(DcboCityID.BoundText) = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·ÕÌ  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter distict  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If
checkEeinvoice = True


End Function

