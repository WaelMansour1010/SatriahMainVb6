VERSION 5.00
Object = "{88F7F54F-F24B-4B64-B0E0-2454E1E6DA40}#1.0#0"; "ciaXPButton30.ocx"
Object = "{81FEA250-2DA5-40F7-A3F1-6F8532B748DB}#1.0#0"; "ciaXPPanel30.ocx"
Object = "{46DBBAE5-ED3E-4D0A-BC4E-8031490B83C7}#1.0#0"; "ciaXPProgress30.ocx"
Object = "{D07199B2-D9E4-4704-B657-AACE5257AEFE}#1.0#0"; "ciaXPStatusBar30.ocx"
Object = "{798A85D3-625A-4512-A9E4-BA96E09CA6A6}#1.0#0"; "ciaXPIML30.ocx"
Object = "{E1BFA30F-D929-4F80-AEDD-76FC2BDF5E23}#1.0#0"; "ciaXPPopUp30.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„” ‘«— ‰Ê«›–"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "FrmMain.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9285
   Begin ciaXPProgress30.ProgressBar30 PrgBarLoad 
      Height          =   255
      Left            =   30
      TabIndex        =   11
      Top             =   6330
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      ApplyCandyEffects=   -1  'True
      LicValid        =   -1  'True
   End
   Begin ciaXPStatusBar30.XPStatusBar30 XPStusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   6270
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   688
      Panels          =   6
      PictureShadowEffect1=   -1  'True
      Enabled1        =   -1  'True
      Key1            =   ""
      Alignment1      =   1
      Caption1        =   "„Êﬁ⁄ «·‘—ﬂ…"
      CaptionForeColor1=   0
      CaptionBold1    =   0   'False
      CaptionItalic1  =   0   'False
      CaptionUnderline1=   0   'False
      CaptionPosition1=   0
      Picture1        =   "FrmMain.frx":0714
      PictureWidth1   =   32
      PictureHeight1  =   32
      PictureSize1    =   2
      PanelStyle1     =   0
      PanelWidth1     =   214
      PanelMinWidth1  =   100
      PicturePosition1=   1
      MaskColor1      =   16777215
      AutoSize1       =   0   'False
      ToolTipText1    =   "«·–Â«» ≈·Ï „Êﬁ⁄ «·‘—ﬂ… ⁄·Ï ‘»ﬂ… «·√‰ —‰ "
      PictureShadowEffect2=   0   'False
      Enabled2        =   -1  'True
      Key2            =   ""
      Alignment2      =   1
      Caption2        =   ""
      CaptionForeColor2=   0
      CaptionBold2    =   0   'False
      CaptionItalic2  =   0   'False
      CaptionUnderline2=   0   'False
      CaptionPosition2=   0
      Picture2        =   "FrmMain.frx":0A66
      PictureWidth2   =   32
      PictureHeight2  =   32
      PictureSize2    =   2
      PanelStyle2     =   2
      PanelWidth2     =   87
      PanelMinWidth2  =   100
      PicturePosition2=   1
      MaskColor2      =   16777215
      ToolTipText2    =   "«· «—ÌŒ «·Õ«·Ï ›Ï «·ÃÂ«“"
      PictureShadowEffect3=   0   'False
      Enabled3        =   -1  'True
      Key3            =   ""
      Alignment3      =   1
      Caption3        =   ""
      CaptionForeColor3=   0
      CaptionBold3    =   0   'False
      CaptionItalic3  =   0   'False
      CaptionUnderline3=   0   'False
      CaptionPosition3=   0
      Picture3        =   "FrmMain.frx":0DB8
      PictureWidth3   =   32
      PictureHeight3  =   32
      PictureSize3    =   2
      PanelStyle3     =   1
      PanelWidth3     =   79
      PanelMinWidth3  =   100
      PicturePosition3=   1
      MaskColor3      =   16777215
      ToolTipText3    =   "«·Êﬁ  «·Õ«·Ï ›Ï «·ÃÂ«“"
      PictureShadowEffect4=   0   'False
      Enabled4        =   -1  'True
      Key4            =   ""
      Alignment4      =   0
      Caption4        =   ""
      CaptionForeColor4=   0
      CaptionBold4    =   0   'False
      CaptionItalic4  =   0   'False
      CaptionUnderline4=   0   'False
      CaptionPosition4=   0
      PictureWidth4   =   32
      PictureHeight4  =   32
      PictureSize4    =   2
      PanelStyle4     =   3
      PanelWidth4     =   50
      PanelMinWidth4  =   100
      PicturePosition4=   1
      MaskColor4      =   16777215
      AutoSize4       =   0   'False
      PictureShadowEffect5=   0   'False
      Enabled5        =   -1  'True
      Key5            =   ""
      Alignment5      =   2
      Caption5        =   ""
      CaptionForeColor5=   0
      CaptionBold5    =   0   'False
      CaptionItalic5  =   0   'False
      CaptionUnderline5=   0   'False
      CaptionPosition5=   0
      PictureWidth5   =   32
      PictureHeight5  =   32
      PictureSize5    =   2
      PanelStyle5     =   4
      PanelWidth5     =   51
      PanelMinWidth5  =   100
      PicturePosition5=   1
      MaskColor5      =   16777215
      AutoSize5       =   0   'False
      ToolTipText5    =   "⁄—÷ Õ«·… «·„› «Õ Caps Lock ›Ï ·ÊÕ… «·„›« ÌÕ"
      PictureShadowEffect6=   0   'False
      Enabled6        =   -1  'True
      Key6            =   ""
      Alignment6      =   1
      Caption6        =   "»—‰«„Ã «·√”Â„ «·„«·Ì…"
      CaptionForeColor6=   0
      CaptionBold6    =   0   'False
      CaptionItalic6  =   0   'False
      CaptionUnderline6=   0   'False
      CaptionPosition6=   0
      Picture6        =   "FrmMain.frx":110A
      PictureWidth6   =   32
      PictureHeight6  =   32
      PictureSize6    =   2
      PanelStyle6     =   0
      PanelWidth6     =   118
      PanelMinWidth6  =   100
      PicturePosition6=   1
      MaskColor6      =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      BorderStyle     =   1
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LicValid        =   -1  'True
   End
   Begin ciaXPPanel30.XPPanel30 XPPanel201 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   1720
      LicValid        =   -1  'True
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   1
         Left            =   8310
         TabIndex        =   0
         Top             =   450
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "062E064A062706310627062A"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":145C
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   3
         Left            =   7302
         TabIndex        =   1
         Top             =   450
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "0645063306270639062F0629"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":17F6
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   4
         Left            =   6300
         TabIndex        =   2
         Top             =   450
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "002E002E002E06390646"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":1B90
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   9
         Left            =   5195
         TabIndex        =   16
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "062706440645062C0645064806390627062A"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":1F2A
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   7
         Left            =   6312
         TabIndex        =   17
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "0627064406390645064406270621"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":22C4
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   8
         Left            =   4198
         TabIndex        =   18
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "0627064406230635064606270641"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":265E
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   2
         Left            =   3201
         TabIndex        =   19
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "062706440645062E062706320646"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":29F8
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   10
         Left            =   2204
         TabIndex        =   20
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "062706440628064606480643"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":2D92
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   16
         Left            =   1207
         TabIndex        =   21
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "06270644062A064206270631064A0631"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":312C
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   6
         Left            =   7309
         TabIndex        =   22
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "062706440645064806380641064A0646"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":34C6
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   5
         Left            =   8310
         TabIndex        =   23
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "06270644064506480631062F064A0646"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":3860
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 XPBtn 
         Height          =   375
         Index           =   20
         Left            =   60
         TabIndex        =   25
         Top             =   60
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "0627064406450633062A062E062F0645064A0646"
         PicturePosition =   2
         ButtonStyle     =   2
         Picture         =   "FrmMain.frx":3BFA
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
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
   End
   Begin ciaXPImageList30.XPImageList30 img16 
      Left            =   2790
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      Size            =   22560
      Images          =   "FrmMain.frx":3F94
      KeyCount        =   24
      Keys            =   "NumbersˇˇˇˇˇˇˇˇˇSerialˇˇˇˇˇChartsˇNote1ˇˇˇCommentˇCompanyˇPersonˇNoteˇˇNews"
   End
   Begin VB.Timer TimAlert 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2490
      Top             =   840
   End
   Begin ciaXPPanel30.XPPanel30 XPPanel301 
      Height          =   5265
      Left            =   -60
      TabIndex        =   10
      Top             =   960
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9287
      BackStyle       =   1
      LicValid        =   -1  'True
      Begin ciaXPPanel30.XPPanel30 XPPanel302 
         Height          =   5265
         Left            =   7620
         TabIndex        =   12
         Top             =   30
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   9287
         BorderLeftStyle =   1
         BorderRightStyle=   1
         BorderBottomStyle=   2
         LicValid        =   -1  'True
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   1299
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "002000200020002000200020002006410627062A06480631062900200634063106270621"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":97D4
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   6
            Top             =   1712
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "00200020002000200020002000200020002006410627062A06480631062900200628064A0639"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":9B6E
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   7
            Top             =   2125
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "002000200020002000200020002000200020002000200020002000200020002000200635064A062706460629"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":9F08
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   3
            Top             =   60
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "0627064406310635064A062F00200627064406270641062A062A0627062D064A"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":A2A2
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   4
            Top             =   473
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "0020002000200020062706440623062C0647063206290020062706440645062A0627062D0629"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":A63C
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   13
            Top             =   886
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "00200020002000200020002000200020062C0631062F0020062706440645062E062706320646"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":A9D6
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   14
            Top             =   2951
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "002000200020062A06350645064A064500200627064406280627063106430648062F"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":AD70
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   15
            Top             =   3780
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "0020062706440627062A06350627064400200628062706440634063106430629"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":B10A
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   24
            Top             =   3364
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "0628062D062B00200639064600200633064A0631064A06270644"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":B4A4
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
         Begin ciaXPButton30.XPButton30 XPBtn 
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   26
            Top             =   2538
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            AutoSelectTheme =   -1  'True
            Caption         =   "06450631062A062C063900200627064406450634062A0631064A0627062A"
            PicturePosition =   2
            ButtonStyle     =   2
            Picture         =   "FrmMain.frx":B83E
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
            UseImageShadow  =   0   'False
            LicValid        =   -1  'True
         End
      End
      Begin ciaXPPopMenu30.XPPopUp30 XPPopUp 
         Left            =   0
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   873
         VisualStyle     =   0
         BeginProperty DefaultMenuItemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuItemSpacing =   0
         LicValid        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   795
         Left            =   390
         Top             =   2370
         Width           =   885
      End
      Begin VB.Image ImgBackGround 
         Height          =   5220
         Left            =   60
         Top             =   30
         Width           =   7560
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo ErrTrap
App.Title = "«·„—«ﬁ» ··Õ”«»«  «·»”Ìÿ…"
Me.Caption = App.Title
Me.XPStusBar.Caption(6) = App.Title
If Dir(App.Path & "\Garphics\wallpaper.bmp") <> "" Then
    ImgBackGround.Picture = LoadPicture(App.Path & "\Garphics\wallpaper.bmp")
End If
XPPanel301.BackColor = vbWhite
Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'CloseApplication
End Sub

Private Sub TimAlert_Timer()
'Dim Msg As String
'Static I As Single
'I = I + 10
'If I = 50 Then
'
'    Msg = "«Â·« »ﬂ„ ›Ï »—‰«„Ã „” ‘«— ‰Ê«›–"
'    Msg = Msg & " œ·Ì·ﬂ «·œ«∆„ ›Ï ⁄«·„ «·»Ê—’…"
'    Msg = Msg & " Ê«·√„Ê«· "
'    DisplayAlert Msg, 5000, True, 0
'    'Me.SetFocus
'ElseIf I = 130 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "„‰ Œ·«· «·»—‰«„Ã Ì„ﬂ‰ﬂ «‰  ﬁÊ„ "
'    Msg = Msg & "»„‘«Âœ… «—«¡ Ê ⁄·Ìﬁ«  Œ»—«¡ «·»Ê—’…"
'    Msg = Msg & " ⁄·Ï «·√”Â„ «·„Œ ·›… ··‘—ﬂ« "
'    DisplayAlert Msg, 5000, True, 3
'    'Me.SetFocus
'ElseIf I = 70 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "ﬁ«„ «·Œ»Ì— ›Ì’· «·Õ„«Ì‰Ï "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "» ”ÃÌ· «·œŒÊ· «·√‰ "
'    Msg = Msg & " "
'    DisplayAlert Msg, 5000, True, 1
'    'Me.SetFocus
'ElseIf I = 100 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & " ‰»ÌÂ "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "Â‰«ﬂ √Õœ «·√”Â„ ﬁœ «— ﬁ⁄ "
'    Msg = Msg & " »’Ê—… ﬂ»Ì—… "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'    DisplayAlert Msg, 5000, True, 2
'    'Me.SetFocus
'
'ElseIf I = 160 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & " ‰»ÌÂ "
'    Msg = Msg & "‘—ﬂ… «·„ﬁ«Ê·«  «·⁄«„… ﬁœ ≈‰Œ›÷ "
'    Msg = Msg & " ”⁄— ”Â„Â« ≈·Ï 100 —Ì«· "
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'    Msg = Msg & Chr(13)
'    DisplayAlert Msg, 5000, True, 4
'    'Me.SetFocus
'ElseIf I = 200 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "Â‰«ﬂ —”«·… ÃœÌœ…"
'    Msg = Msg & "·ﬂ „‰ «·Œ»Ì—  "
'    Msg = Msg & Chr(13)
'    Msg = Msg & " „Õ„œ «·”Ìœ „Õ„œ "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'    Msg = Msg & Chr(13)
'    DisplayAlert Msg, 5000, True, 11
'    'Me.SetFocus
'ElseIf I = 220 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "‘«Âœ „Ì“«‰ «·„œ›Ê⁄« "
'    Msg = Msg & "«·ÃœÌœ "
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'    Msg = Msg & Chr(13)
'    DisplayAlert Msg, 5000, True, 8
'    'Me.SetFocus
'ElseIf I = 250 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "Â‰«ﬂ ≈— ›«⁄ ›Ï ”⁄— «·œÂ»"
'    Msg = Msg & " "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'
'    DisplayAlert Msg, 5000, True, 9
'    'Me.SetFocus
'ElseIf I = 290 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "”Ê› ”Õ» «·⁄„·… «·„⁄œÌ‰… "
'    Msg = Msg & " „‰ «·”Êﬁ "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'
'    DisplayAlert Msg, 5000, True, 5
'    'Me.SetFocus
'ElseIf I = 320 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "‘«Âœ «·—”„ «·»Ì«‰Ï «·Œ«’ "
'    Msg = Msg & " »√”Â„ ‘—ﬂ… «·„ Õœ… «·⁄—»Ì… "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'    AlertBox.DisplayAlert Msg, 5000, True, 6
'    'Me.SetFocus
'ElseIf I = 400 Then
'    Set AlertBox = New frmAlert
'    Msg = Msg & "‘«Âœ «· €Ì—«  ›Ï „Ì“«‰ «”Â„ﬂ "
'    Msg = Msg & " «·„‘«—ﬂ… ›Ï «·»Ê—’… «·ÌÊ„ "
'    Msg = Msg & Chr(13)
'    Msg = Msg & "≈÷€ÿ Â‰« Õ Ï  ‘«Âœ «· ›«’Ì·"
'    DisplayAlert Msg, 5000, True, 10
'    'Me.SetFocus
'    I = 0
'End If
ErrTrap:
End Sub
Private Sub XPBtn_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        FrmBillBuy.Show
    Case 1
        FrmOptions.Show
    Case 2
        FrmStoreData.Show
    Case 3
       SendKeys "{f1}"
    Case 4
        TimAlert.Enabled = False
        frmAbout.Show
        TimAlert.Enabled = True
    Case 5
        FrmCompany.Show
    Case 6
        FrmEmployee.Show
    Case 7
        FrmCustemers.Show
    Case 8
        FrmItems.Show
    Case 9
        FrmGroups.Show
    Case 10
        FrmBanksData.Show
    Case 11
        FrmSaleBill.Show
    Case 12
        FrmMaintenence.Show
    Case 13
        FrmOpeningBalance.Show
    Case 14
        FrmSearchSerial.Show
    Case 15
        FrmGard.Show
    Case 16
        FrmReports.Show
    Case 17
        FrmBarcode.Show
    Case 18
        FrmConect_US.Show
    Case 19
        FrmSerialData.Show vbModal
    Case 20
        FrmUsers.Show vbModal
    Case 21
        FrmReturnpurchases.Show
End Select
Exit Sub
ErrTrap:
End Sub
Private Sub XPListView301_ItemMouseUp(Button As Integer, _
Shift As Integer, X As Single, Y As Single, _
Item As ciaXPListView30.tListItemInfo, ItemIndex As Long)
On Error GoTo ErrTrap
Dim tp            As PointAPI
Dim lX            As Single
Dim lY            As Single
Dim tr            As RECT
If Button = vbRightButton Then
    GetCursorPos tp
    lX = (tp.X) * Screen.TwipsPerPixelX
    lY = tp.Y * Screen.TwipsPerPixelY
    Me.XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
End If
Exit Sub
ErrTrap:
End Sub
Private Sub XPListView301_SubItemClick(Item As ciaXPListView30.tListItemInfo, ItemIndex As Long, SubItemIndex As Integer)
On Error GoTo ErrTrap
Dim tp            As PointAPI
Dim lX            As Single
Dim lY            As Single
Dim tr            As RECT
'If Button = vbRightButton Then
    GetCursorPos tp
    lX = (tp.X) * Screen.TwipsPerPixelX
    lY = tp.Y * Screen.TwipsPerPixelY
    Me.XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
'End If
Exit Sub
ErrTrap:
End Sub

Private Sub XPStusBar_MouseDown(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, _
    ByVal Y As Single, ByVal PanelIndex As Integer, _
    Panel As ciaXPStatusBar30.cPanel)
If PanelIndex = 1 Then
    OpenWebSite
End If
End Sub
