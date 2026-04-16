VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCashing0 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„ﬁ»Ê÷« "
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   DrawWidth       =   10
   HelpContextID   =   290
   Icon            =   "FrmCashing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   8085
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   4200
      Top             =   840
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.TextBox txtperson 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   5040
      Width           =   2685
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì Õ«·… «·„‘«—Ì⁄"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   2040
      Width           =   4215
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬁ«Ê· »«ÿ‰"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄„Ì· ‰Â«∆Ì"
         Height          =   195
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   600
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAdv_payment_value 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   91
      Top             =   2760
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ŒÌ«—« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   600
      Width           =   3735
      Begin VB.OptionButton Option6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌœ „” Œ·’« "
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
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌœ ›Ê« Ì—"
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
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "FIFO"
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
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "œ›⁄Â „ﬁœ„Â"
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
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ÕœÌœ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCashing.frx":038A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton4 
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ÕœÌœ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCashing.frx":03A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.TextBox XPTxtBillID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   2040
      TabIndex        =   80
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«ŸÂ«— «·«ﬁ”«ÿ"
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
      MICON           =   "FrmCashing.frx":03C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
      Height          =   885
      Index           =   1
      Left            =   1500
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   6150
      Width           =   6495
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   200
         Width           =   1875
      End
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   72
         Top             =   180
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   73
         Top             =   510
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   33
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   510
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   29
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   30
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   31
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   32
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame FraInfo 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«   Â„ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2235
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   3510
      Width           =   3705
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   0
         Left            =   1830
         TabIndex        =   62
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":03DE
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0540
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   2
         Left            =   1830
         TabIndex        =   64
         Top             =   1350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":06A2
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   65
         Top             =   1350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0804
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   4
         Left            =   1830
         TabIndex        =   66
         Top             =   1920
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0966
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   67
         Top             =   1920
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0AC8
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   68
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0C2A
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   69
         Top             =   1110
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0D8C
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmCashing.frx":0EEE
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
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
         Height          =   225
         Index           =   28
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
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
         Height          =   225
         Index           =   27
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
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
         Height          =   225
         Index           =   26
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
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
         Height          =   225
         Index           =   25
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
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
         Height          =   225
         Index           =   24
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Ã„«·Ï „ﬁ»Ê÷«  «·ÌÊ„:"
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
         Height          =   225
         Index           =   23
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   540
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈”»Ê⁄ «·Õ«·Ï"
         Height          =   255
         Index           =   22
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
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
         Height          =   225
         Index           =   21
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ﬁ»Ê÷«  ›Ï «·‘Â— «·Õ«·Ï :"
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
         Height          =   225
         Index           =   20
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1680
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ﬁ»Ê÷«  ›Ï «·≈”»Ê⁄ «·Õ«·Ï:"
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
         Height          =   225
         Index           =   19
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1110
         Width           =   2235
      End
   End
   Begin VB.Frame FraNote 
      BackColor       =   &H00E2E9E9&
      Height          =   1485
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   3510
      Width           =   4155
      Begin MSComCtl2.DTPicker DtpChequeDueDate 
         Height          =   315
         Left            =   30
         TabIndex        =   14
         Top             =   1140
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Format          =   91619329
         CurrentDate     =   39614
      End
      Begin VB.TextBox TxtChequeNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   810
         Width           =   2685
      End
      Begin MSDataListLib.DataCombo DcboBankName 
         Height          =   315
         Left            =   30
         TabIndex        =   12
         Top             =   480
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   30
         TabIndex        =   11
         Top             =   150
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
         Height          =   285
         Index           =   17
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   285
         Index           =   16
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   285
         Index           =   15
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Œ“‰…"
         Height          =   285
         Index           =   9
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   3840
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3150
      Width           =   2685
   End
   Begin MSDataListLib.DataCombo DcboRevenuesTypes 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.CheckBox ChkTrans 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ Õ”«» ›« Ê—…"
      Height          =   195
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   510
      Width           =   1575
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3840
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2445
      Width           =   2685
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   585
      Left            =   3810
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   5520
      Width           =   2715
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   4260
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1365
      Width           =   2265
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   5100
      TabIndex        =   0
      Top             =   1020
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   91619329
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   1680
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   4680
      TabIndex        =   20
      Top             =   7080
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   540
      Index           =   0
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7440
      Width           =   7995
      _cx             =   14102
      _cy             =   953
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
         Height          =   375
         Index           =   0
         Left            =   7140
         TabIndex        =   22
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6244
         TabIndex        =   23
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   5355
         TabIndex        =   24
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   4460
         TabIndex        =   25
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   3568
         TabIndex        =   26
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   6
         Left            =   0
         TabIndex        =   27
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   892
         TabIndex        =   28
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   5
         Left            =   2676
         TabIndex        =   29
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   7
         Left            =   1784
         TabIndex        =   41
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1005
      Index           =   0
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   660
      Width           =   3735
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox TxtTransID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin ImpulseButton.ISButton CmdSearchTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   8
         Top             =   570
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCashing.frx":1050
      End
      Begin ImpulseButton.ISButton CmdOpenTrans 
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   570
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCashing.frx":13EA
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   12
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   10
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Width           =   1305
      End
   End
   Begin ImpulseAniLabel.ISAniLabel LblLink 
      Height          =   315
      Left            =   90
      TabIndex        =   42
      Top             =   2070
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   556
      ActiveUnderline =   -1  'True
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   4210688
      MousePointer    =   99
      MouseIcon       =   "FrmCashing.frx":1784
      BackColor       =   14871017
      Alignment       =   1
      Caption         =   ""
      ColorHover      =   16711680
      RightToLeft     =   -1  'True
      ImageCount      =   0
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   2040
      TabIndex        =   81
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«ŸÂ«— ”‰œ «·„œÌÊ‰Ì…"
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
      MICON           =   "FrmCashing.frx":18E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DCPROJECT 
      Height          =   315
      Left            =   8640
      TabIndex        =   82
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcCostCenter 
      Bindings        =   "FrmCashing.frx":1902
      Height          =   315
      Left            =   0
      TabIndex        =   93
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   9
      Left            =   3360
      TabIndex        =   101
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·ﬁÌœ"
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   585
      Index           =   1
      Left            =   0
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   0
      Width           =   8025
      _cx             =   14155
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "«·„ﬁ»Ê÷« "
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
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   5460
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1125
         TabIndex        =   109
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
         ButtonImage     =   "FrmCashing.frx":1917
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
         Left            =   60
         TabIndex        =   110
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
         ButtonImage     =   "FrmCashing.frx":1CB1
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
         Left            =   1650
         TabIndex        =   111
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
         ButtonImage     =   "FrmCashing.frx":204B
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
         Left            =   585
         TabIndex        =   112
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
         ButtonImage     =   "FrmCashing.frx":23E5
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   2400
         TabIndex        =   113
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   1680
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   -360
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Index           =   11
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   114
         Top             =   60
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ «·„ﬂ—„"
      Height          =   285
      Index           =   36
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "œ›⁄Â „ﬁœ„Â"
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
      Height          =   285
      Index           =   35
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   92
      Top             =   2775
      Width           =   1395
   End
   Begin VB.Label lblsqlstring 
      Alignment       =   1  'Right Justify
      Height          =   855
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   85
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‘—Ê⁄"
      Height          =   285
      Index           =   34
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Index           =   18
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2430
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      Height          =   315
      Index           =   14
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   3150
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—’Ìœ «·Õ«·Ï:"
      Height          =   315
      Index           =   13
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   2190
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "–·ﬂ „ﬁ«»·"
      Height          =   285
      Index           =   5
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   5520
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   285
      Index           =   4
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   690
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· √Ê «·„Ê—œ"
      Height          =   315
      Index           =   3
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
      Height          =   285
      Index           =   2
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   2460
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1035
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„ﬁ»Ê÷« "
      Height          =   285
      Index           =   0
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1380
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   8
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   7080
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1770
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7080
      Width           =   825
   End
End
Attribute VB_Name = "FrmCashing0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim Line1 As Double
Dim Line2 As Double
Dim Line3 As Double
Dim Line4 As Double

Dim departement_name As Integer
Dim numbering_type As Integer

Private Sub ALLButton1_Click()
If IsNumeric(Me.DBCboClientName.BoundText) Then
INSTALLMENT_DATA1.Show
INSTALLMENT_DATA1.Adodc1.CommandType = adCmdText
INSTALLMENT_DATA1.Adodc1.RecordSource = "select *  FROM INSTALLMENT_DETAILS where payed=0 and cust_id =" & Me.DBCboClientName.BoundText
INSTALLMENT_DATA1.Adodc1.Refresh
 
INSTALLMENT_DATA1.id.text = Me.DBCboClientName.BoundText
INSTALLMENT_DATA1.lblcustid = Me.DBCboClientName.BoundText
INSTALLMENT_DATA1.txtname.text = Me.DBCboClientName.text
End If
End Sub

Private Sub ALLButton2_Click()
If IsNumeric(Me.DBCboClientName.BoundText) Then
sanad_dean.Show
sanad_dean.lblid = DBCboClientName.BoundText
sanad_dean.lblname = DBCboClientName.text
'sanad_dean.lblaccountcode.Caption = txtaccount.text
sanad_dean.Adodc1.CommandType = adCmdText
sanad_dean.Adodc1.RecordSource = "select*  FROM sanad_dean where cust_id=" & DBCboClientName.BoundText
sanad_dean.Adodc1.Refresh
sanad_dean.ALLButton1.Visible = False
sanad_dean.ALLButton1.Visible = False


sanad_dean.Adodc2.CommandType = adCmdText
sanad_dean.Adodc2.RecordSource = "select *  FROM member_child where cust_id=" & DBCboClientName.BoundText
sanad_dean.Adodc2.Refresh
End If
End Sub

Private Sub ALLButton3_Click()
lblsqlstring.Caption = ""
FrmPaymentTime1.Show
FrmPaymentTime1.lblcusid = DBCboClientName.BoundText
FrmPaymentTime1.LblValue = Val(XPTxtVal.text)
End Sub

Private Sub ALLButton4_Click()
If DCboCashType.ListIndex <> 5 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â–… «·⁄„·Ì… „ «Õ… „⁄ ›Ê« Ì— «·„‘«—Ì⁄ ›ﬁÿ", vbInformation
    Else
    MsgBox "This Process For Project Bill Only", vbInformation
    
    End If
            DCboCashType.SetFocus
        SendKeys "{F4}"
    Exit Sub
End If

If Val(DBCboClientName.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "«Œ — „‘—Ê⁄ «Ê·«", vbInformation
    Else
    MsgBox "select Project Firstly, vbInformation"
    
    End If
         DBCboClientName.SetFocus
            SendKeys "{F4}"
    Exit Sub


End If
 
End Sub

Private Sub CboPayMentType_Change()
If Me.CboPaymentType.ListIndex = 0 Then
    Me.lbl(9).Enabled = True
    Me.DcboBox.Enabled = True
    Me.lbl(15).Enabled = False
    Me.lbl(16).Enabled = False
    Me.lbl(17).Enabled = False
    Me.DcboBankName.Enabled = False
    Me.TxtChequeNumber.Enabled = False
    Me.DtpChequeDueDate.Enabled = False
ElseIf Me.CboPaymentType.ListIndex = 1 Then
    Me.lbl(9).Enabled = False
    Me.DcboBox.Enabled = False
    Me.lbl(15).Enabled = True
    Me.lbl(16).Enabled = True
    Me.lbl(17).Enabled = True
    Me.DcboBankName.Enabled = True
    Me.TxtChequeNumber.Enabled = True
    Me.DtpChequeDueDate.Enabled = True
Else
    Me.lbl(9).Enabled = False
    Me.DcboBox.Enabled = False
    Me.lbl(15).Enabled = False
    Me.lbl(16).Enabled = False
    Me.lbl(17).Enabled = False
    Me.DcboBankName.Enabled = False
    Me.TxtChequeNumber.Enabled = False
    Me.DtpChequeDueDate.Enabled = False
End If
End Sub

Private Sub CboPayMentType_Click()
CboPayMentType_Change
End Sub

Private Sub ChkTrans_Click()
Me.lbl(10).Enabled = ChkTrans.value
Me.lbl(12).Enabled = ChkTrans.value
Me.CboTrans.Enabled = ChkTrans.value
Me.TxtTransID.Enabled = ChkTrans.value
Me.TxtTransSerial.Enabled = ChkTrans.value
Me.CmdSearchTrans.Enabled = ChkTrans.value
Me.CmdOpenTrans.Enabled = ChkTrans.value
End Sub
Function sand_numbering() As String
On Error Resume Next
Dim start_at As Integer
Dim end_at As Integer
Dim auto_sanad_no As String
Dim no As Integer
auto_sanad_no = ""
departement_name = 1
 
connection_string = Cn.ConnectionString
numbering.ConnectionString = connection_string
numbering.CommandType = adCmdText
numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=2"
numbering.Refresh
If numbering.Recordset.RecordCount = 0 Then
numbering_type = 0
Else
numbering_type = numbering.Recordset.Fields!numbering_id
start_at = numbering.Recordset.Fields!start_at
end_at = numbering.Recordset.Fields!end_at

End If

If numbering_type = 1 Then
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
detect_no.Refresh

 If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
 If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
 If detect_no.Recordset.Fields!last_sand_no >= end_at Then
 sand_numbering = "error"
 Exit Function
 End If
 End If
Else
If numbering_type = 2 Then
 
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2)
'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
detect_no.Refresh

If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
   no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
   If end_at = 0 Then end_at = no + 1
 If no >= end_at Then
 sand_numbering = "error"
 Exit Function
 End If
 End If


Else
If numbering_type = 3 Then
 
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4)
'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
detect_no.Refresh
If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
If end_at = 0 Then end_at = no + 1
 If no >= end_at Then
 sand_numbering = "error"
 Exit Function
 End If
 End If

 
End If
 
End If
End If

If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

                If numbering_type = 0 Then
                 ' auto_sanad_no = 1
                Else
                    If numbering_type = 1 Then
                    auto_sanad_no = start_at
                Else
                
                    If numbering_type = 2 Then
                    auto_sanad_no = Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2) & start_at

                Else
                     If numbering_type = 3 Then
                        auto_sanad_no = Mid(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & start_at

                  End If
                  End If
                  End If
                  End If

Else
                If numbering_type = 0 Then
                'auto_sanad_no = x + 1
                Else
                    If numbering_type = 1 Then
                  auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
                Else
                
                    If numbering_type = 2 Then
                  '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                   ' no = 1
                  '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                  '  Else
                    no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (no + 1)
                  '  End If
                    
                    
                      
                Else
                     If numbering_type = 3 Then
                  '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                      'no = 1
                  '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                  '    Else
                           no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                      auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (no + 1)

                  '    End If

                  End If
                  End If
                  End If
                  End If

End If
sand_numbering = auto_sanad_no

'MsgBox auto_sanad_no

End Function

Private Sub Cmd_Click(Index As Integer)
Dim cNoteReport As ClsNotesReports
Dim Msg As String
'On Error GoTo ErrTrap
Select Case Index
    Case 0
        If SystemOptions.SysRegisterState = DemoRun Then
            If Not rs Is Nothing Then
                If Not (rs.BOF Or rs.EOF) Then
                    If rs.RecordCount >= 25 Then
                        Msg = "›Ï «·‰”Œ… «· Ã—Ì»Ì… ·«Ì„ﬂ‰  ”ÃÌ· «ﬂÀ— „‰ 25 ⁄„·Ì… ﬁ»÷ «Ê œ›⁄"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If
            End If
        End If
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
        clear_all Me
        TxtModFlg.text = "N"
 '       XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
       ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
        Me.DCboUserName.BoundText = user_id
        XPDtbTrans.SetFocus
        Text1.text = setfoxy
        Option4.value = True
    Case 1
        If DoPremis(Do_Edit, Me.name, True) = False Then
            Exit Sub
        End If
        TxtModFlg.text = "E"
        Me.DCboUserName.BoundText = user_id
    Case 2
        
        If Option2.value = True And lblsqlstring.Caption = "" Then MsgBox "·«»œ „‰  ÕœÌœ ›Ê« Ì—": Exit Sub
         If Me.TxtModFlg.text = "N" Then
         
                If TxtNoteSerial.text = "" Then
                       If Notes_coding(Val(my_branch), XPDtbTrans.value) = "error" Then
                       MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                       Else
                       
                       If Notes_coding(Val(my_branch), XPDtbTrans.value) = "" Then
                       MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                       Else
                       TxtNoteSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)
                       End If
                       End If
                End If
        
             If TxtNoteSerial1.text = "" Then
                If Voucher_coding(Val(my_branch), XPDtbTrans.value, 2, 4) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁ»÷ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                
                If Voucher_coding(Val(my_branch), XPDtbTrans.value, 2, 4) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                TxtNoteSerial1.text = Voucher_coding(Val(my_branch), XPDtbTrans.value, 2, 4)
                End If
                End If
             End If
         End If
    
    
 
'TxtNoteSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)
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
        Load FrmNotesSearch
        FrmNotesSearch.SearchType = 4
        FrmNotesSearch.Show vbModal
    Case 6
        Unload Me
    Case 7
        If DoPremis(Do_Print, Me.name, True) = False Then
            Exit Sub
        End If
       ' If Val(Me.XPTxtID.text) <> 0 Then
       '     Set cNoteReport = New ClsNotesReports
       '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
       '     Set cNoteReport = Nothing
       ' End If
       If TxtNoteSerial <> "" Then
       print_report TxtNoteSerial
       End If
    Case 8
        'ViewDataList
    Case 9
    ShowGL_cc Me.TxtNoteSerial.text, , 200
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

MySQL = "Select * From payment_voucher  where noteserial='" & NoteSerial & "'"

 

If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.Path & "\Reports\" & "Payment_voucher.rpt"
Else
    StrFileName = App.Path & "\Reports\" & "Payment_voucher.rpt"
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
   xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
   
    StrReportTitle = "" '& StrAccountName
 
Else
 
    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
     xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
    StrReportTitle = ""
 
End If
xReport.ParameterFields(3).AddCurrentValue user_name
xReport.ReportTitle = StrReportTitle
xReport.EnableParameterPrompting = False
xReport.ApplicationName = App.Title
xReport.ReportAuthor = App.Title
Set CViewer = New ClsReportViewer
CViewer.FireReport xReport, WindowTarget, ""

RsData.Close
Set RsData = Nothing
Screen.MousePointer = vbDefault

End Function


Private Sub ViewDataList()
Dim FrmView As FrmViewList
Dim Fg As VSFlex8UCtl.vsFlexGrid
Dim StrSQL As String
Dim rs As ADODB.Recordset
Dim StrComboList As String
Dim GrdBack As ClsBackGroundPic
'Dim cProgress As ClsProgress
Dim BolFrmLoaded As Boolean
Set FrmView = New FrmViewList
Set Fg = FrmView.vsfGroup1.vsFlexGrid

With Fg
    .Cols = 18
    .RowHeightMin = 320
    .ExplorerBar = flexExSortShowAndMove
    .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
    .ColKey(0) = "NoteID"
    .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
    .ColKey(1) = "NoteSerial"
    .TextMatrix(0, 2) = "«· «—ÌŒ"
    .ColKey(2) = "NoteDate"
    .TextMatrix(0, 3) = " ‰Ê⁄ «·„ﬁ»Ê÷« "
    .ColKey(3) = "Name"
    .TextMatrix(0, 4) = "ﬁÌ„… «·„ﬁ»Ê÷« "
    .ColKey(4) = "Note_Value"
    .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
    .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
    .ColKey(5) = "BoxName"
    .TextMatrix(0, 6) = "„·«ÕŸ« "
    .ColKey(6) = "Remark"
    .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
    .ColKey(7) = "UserName"
    
    StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & _
    "Remark, UserName From ExpensesReport"
    StrSQL = StrSQL + " Order By NoteID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
    '------------------------------------
    '
    '
    '
    '
    
    '------------------------------------
    Set .DataSource = rs
    .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
    .ColKey(0) = "NoteID"
    .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
    .ColKey(1) = "NoteSerial"
    .TextMatrix(0, 2) = "«· «—ÌŒ"
    .ColKey(2) = "NoteDate"
    .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
    .ColKey(3) = "Name"
    .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
    .ColKey(4) = "Note_Value"
    .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
    .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
    .ColKey(5) = "BoxName"
    .TextMatrix(0, 6) = "„·«ÕŸ« "
    .ColKey(6) = "Remark"
    .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
    .ColKey(7) = "UserName"
    
    'Rs.Close
    'Set Rs = Nothing
    .AutoSize 0, .Cols - 1, False
End With
Set GrdBack = New ClsBackGroundPic
FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
FrmView.vsfGroup1.SetRTL = True
FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
FrmView.vsfGroup1.sql = StrSQL
FrmView.vsfGroup1.ShowTreeGroups = True
FrmView.vsfGroup1.update
FrmView.SetDblClickRetrun Me, "NoteID"
FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
FrmView.Show
End Sub
Private Sub CmdHelp_Click()
SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdSearchTrans_Click()
Dim Msg As String
If Me.CboTrans.ListIndex = -1 Then
    Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·Õ—ﬂ… «·„—«œ «·»ÕÀ ⁄‰Â«..."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    CboTrans.SetFocus
    SendKeys "{F4}"
    Exit Sub
End If
If Me.CboTrans.ListIndex = 0 Then
   ' ›« Ê—… „»Ì⁄« 
    Load FrmBuySearch
    FrmBuySearch.DealingForm = InvoiceTransaction
    Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
    FrmBuySearch.CboPaymentType.ListIndex = 1
    FrmBuySearch.CboPaymentType.Enabled = False
    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
    FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
    FrmBuySearch.Show
ElseIf Me.CboTrans.ListIndex = 1 Then
    '›« Ê—… „— Ã⁄ „‘ —Ì« 
    Load FrmBuySearch
    FrmBuySearch.DealingForm = ReturnTransaction
    Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
    FrmBuySearch.CboPaymentType.ListIndex = 1
    FrmBuySearch.CboPaymentType.Enabled = False
    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
    FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
    FrmBuySearch.Show vbModal
ElseIf Me.CboTrans.ListIndex = 2 Then
    '›« Ê—… ’Ì«‰…
    Load FrmMaintanenceSearch
    Set FrmMaintanenceSearch.ExtraRetrunObject = Me.TxtTransID
    FrmMaintanenceSearch.CboPaymentType.ListIndex = 1
    FrmMaintanenceSearch.SearchType = 4
    FrmMaintanenceSearch.CboPaymentType.Enabled = False
    FrmMaintanenceSearch.Show vbModal
End If
End Sub

Private Sub Command1_Click()
 
End Sub

Private Sub DBCboClientName_Change()
If DBCboClientName.BoundText = "" Then Exit Sub
WriteCustomerBal
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", Val(Me.DBCboClientName.BoundText))
    If DCboCashType.ListIndex = 5 Then
    If Option4.value = True Then
    Me.DcboCreditSide.BoundText = get_project_customer_account(DBCboClientName.BoundText, "End_user_Account")
     
    Else
    
    Me.DcboCreditSide.BoundText = get_project_customer_account(DBCboClientName.BoundText, "sub_contractor_Account")
    End If
    
    
    End If
End If
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
'WriteCustomerBal
End Sub

Private Sub DcboBankName_Click(Area As Integer)
If DcboBankName.BoundText = "" Then Exit Sub
Dim RsSavRec As ADODB.Recordset
Dim My_SQL As String

If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    'Me.DcboDebitSide.BoundText =   "a1a2a4"
    My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

Set RsSavRec = New ADODB.Recordset
RsSavRec.CursorLocation = adUseClient
RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 

 Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value

End If
End Sub

Private Sub DcboBox_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
End If
End Sub

Private Sub DCboCashType_Change()
On Error GoTo ErrTrap
Frame2.Enabled = False
Dim StrSQL As String
Dim intDef As Integer
Select Case DCboCashType.ListIndex
    Case 0
        Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
        Me.DBCboClientName.Visible = True
        Me.DcboRevenuesTypes.Visible = False
        ChkTrans.Visible = True
        Fra(0).Visible = True
                If SystemOptions.UserInterface <> EnglishInterface Then
                    Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
                Else
                   Me.lbl(3).Caption = "Customer Name"
                End If
        
        Me.lbl(13).Visible = True
        Me.LblLink.Visible = True
    Case 1
        Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
        Me.DBCboClientName.Visible = True
        Me.DcboRevenuesTypes.Visible = False
        ChkTrans.Visible = True
        Fra(0).Visible = True
               If SystemOptions.UserInterface <> EnglishInterface Then
                    Me.lbl(3).Caption = "«”„ «·„Ê—œ"
                Else
                   Me.lbl(3).Caption = "Vendor Name"
                End If
                
        
        Me.lbl(13).Visible = True
        Me.LblLink.Visible = True
    Case 2
        Dcombos.GetPersons Me.DBCboClientName
        Me.DBCboClientName.Visible = True
        Me.DcboRevenuesTypes.Visible = False
        ChkTrans.Visible = False
        Fra(0).Visible = False

                If SystemOptions.UserInterface = EnglishInterface Then
                    Me.lbl(3).Caption = "name"
                Else
                   Me.lbl(3).Caption = "„ﬁ«Ê· «·»«ÿ‰"
                End If
                
        Me.lbl(13).Visible = True
        Me.LblLink.Visible = True
    Case 3
        '≈Ì—«œ«  ≈Œ—Ï
        Me.DBCboClientName.Visible = False
        Me.DcboRevenuesTypes.Visible = True
        Me.ChkTrans.Visible = False
        Fra(0).Visible = False
        
               If SystemOptions.UserInterface <> EnglishInterface Then
                    Me.lbl(3).Caption = "‰Ê⁄ «·«Ì—«œ"
                Else
                   Me.lbl(3).Caption = "RVN Type"
                End If
                
        Me.lbl(13).Visible = False
        Me.LblLink.Visible = False
        
 Case 4
         Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
        Me.DBCboClientName.Visible = True
        Me.DcboRevenuesTypes.Visible = False
        ChkTrans.Visible = True
        Fra(0).Visible = True
                If SystemOptions.UserInterface <> EnglishInterface Then
                    Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
                Else
                   Me.lbl(3).Caption = "Customer Name"
                End If
        
        Me.lbl(13).Visible = True
        Me.LblLink.Visible = True
        
 Case 5
 Dim My_SQL As String
     My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
      fill_combo Me.DBCboClientName, My_SQL
      
         
        Me.DBCboClientName.Visible = True
        Me.DcboRevenuesTypes.Visible = False
 
                If SystemOptions.UserInterface <> EnglishInterface Then
                    Me.lbl(3).Caption = "«”„ «·„‘—Ê⁄"
                Else
                   Me.lbl(3).Caption = "project Name"
                End If
        
 Frame2.Enabled = True
        
        
End Select
cSearchDcbo.Refresh
Exit Sub
ErrTrap:
End Sub
Private Sub DCboCashType_Click()
DCboCashType_Change
End Sub

Private Sub DcboRevenuesTypes_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
End If
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
CostCenterSearch.Show
CostCenterSearch.RetrunType = 6
End If
End Sub

Private Sub Form_Load()
 Me.left = (MDIFrmMain.Width - Me.Width) / 2
    Me.top = (MDIFrmMain.Height - Me.Height) / 2 - 500

On Error GoTo ErrTrap
Dim StrSQL As String
Dim Msg As String
Set Dcombos = New ClsDataCombos
StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
fill_combo Me.DcCostCenter, StrSQL

Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
Set Cmd(8).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
'Resize_Form Me
AddTip
DCboCashType.AddItem "„‰ ⁄„Ì·"
DCboCashType.AddItem "„‰ „Ê—œ"
DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
DCboCashType.AddItem "≈Ì—«œ«  ≈Œ—Ï"
DCboCashType.AddItem "„œ›Ê⁄«  „ﬁœ„Â"
DCboCashType.AddItem "„‘—Ê⁄"

With Me.CboPaymentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "‘Ìﬂ"
End With
Dcombos.GetUsers Me.DCboUserName
Dcombos.GetBoxes Me.DcboBox
Dcombos.GetBanks Me.DcboBankName
Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
Dcombos.GetRevenuesTypes Me.DcboRevenuesTypes
Set cSearchDcbo = New clsDCboSearch
Set cSearchDcbo.Client = Me.DBCboClientName
Dcombos.GetAccountingCodes Me.DcboDebitSide
Dcombos.GetAccountingCodes Me.DcboCreditSide

Set rs = New ADODB.Recordset
StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Not (rs.BOF Or rs.EOF) Then
    rs.MoveLast
End If
SetDtpickerDate Me.XPDtbTrans
SetDtpickerDate Me.DtpChequeDueDate
With Me.CboTrans
    .Clear
    .AddItem "›« Ê—… „»Ì⁄« "
    .AddItem "„— Ã⁄ „‘ —Ì« "
    .AddItem " ”·Ì„ ’Ì«‰… ·⁄„Ì·"
    .AddItem "Œœ„« "
End With
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If


Msg = "„·ÕÊŸ…:-"
Msg = Msg & Chr(13) & "≈–« ﬂ«‰  Â–Â «·„ﬁ»Ê÷«   Õ’Ì· ·›« Ê—… „⁄Ì‰…"
Msg = Msg & "›ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ » ÕœÌœ Â–Â «·›« Ê—… "
Msg = Msg & "Õ Ï Ì „ —»ÿ ⁄„·Ì… «· Õ’Ì· Â–Â „⁄ «·›« Ê—…"
Me.lbl(11).Caption = Msg
SetDtpickerDate Me.XPDtbTrans
ChkTrans.value = Unchecked
ChkTrans_Click
XPBtnMove_Click 2
Me.TxtModFlg.text = "R"
WriteInfo
  
Dim My_SQL As String

'My_SQL = "  select expanses_account,account_name from projects  where not (account_no is null)"
     My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
fill_combo DCPROJECT, My_SQL

If OPEN_NEW_SCREEN = True Then
Cmd_Click (0)
End If
Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
If rs.state = adStateOpen Then
    If Not (rs.EOF Or rs.BOF) Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If
    End If
    rs.Close
    Set rs = Nothing
End If
Set TTP = Nothing
'Set EmpReport = Nothing
Exit Sub
ErrTrap:
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Index = 18 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).ToolTipText = "ﬁÌ„… „»·€ «·„ﬁ»Ê÷« :" & lbl(18).Caption
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        Me.lbl(18).ToolTipText = "Notes Recivable Value:" & lbl(18).Caption
    End If
End If
End Sub


Private Sub LblLink_Click()
Dim LngCusID As Long
If DoPremis(Do_Print, "ReportCustomers", True) = False Then
    Exit Sub
End If
LngCusID = Val(Me.DBCboClientName.BoundText)
OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0
End Sub

Private Sub Option1_Click()
If Option2.value = True Then
ALLButton3.Enabled = True
Else
ALLButton3.Enabled = False
End If
If Option6.value = True Then
ALLButton4.Enabled = True
Else
ALLButton4.Enabled = False
End If

End Sub

Private Sub Option2_Click()
If Option2.value = True Then
ALLButton3.Enabled = True
Else
ALLButton3.Enabled = False
End If
If Option6.value = True Then
ALLButton4.Enabled = True
Else
ALLButton4.Enabled = False
End If

End Sub

Private Sub Option3_Click()
If Option2.value = True Then
ALLButton3.Enabled = True
Else
ALLButton3.Enabled = False
End If

If Option6.value = True Then
ALLButton4.Enabled = True
Else
ALLButton4.Enabled = False
End If

End Sub

Private Sub Option4_Click()
If DCboCashType.ListIndex <> 5 Then Exit Sub
    If Option4.value = True Then
    Me.DcboCreditSide.BoundText = get_project_customer_account(Val(DBCboClientName.BoundText), "End_user_Account")
     
    Else
    
    Me.DcboCreditSide.BoundText = get_project_customer_account(Val(DBCboClientName.BoundText), "sub_contractor_Account")
    End If
End Sub

Private Sub Option5_Click()
If DCboCashType.ListIndex <> 5 Then Exit Sub
    If Option4.value = True Then
    Me.DcboCreditSide.BoundText = get_project_customer_account(DBCboClientName.BoundText, "End_user_Account")
     
    Else
    
    Me.DcboCreditSide.BoundText = get_project_customer_account(DBCboClientName.BoundText, "sub_contractor_Account")
    End If
End Sub

Private Sub Option6_Click()
If Option6.value = True Then
ALLButton4.Enabled = True
Else
ALLButton4.Enabled = False
End If
If Option6.value = True Then
ALLButton4.Enabled = True
Else
ALLButton4.Enabled = False
End If

End Sub

Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Select Case Me.TxtModFlg.text
    Case "R"
    
        If SystemOptions.UserInterface = EnglishInterface Then
         Me.Caption = "Receipts"
        Else
'        Me.Caption = "«·„ﬁ»Ê÷« "
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
        XPTxtVal.locked = True
        XPDtbTrans.Enabled = False
        XPMTxtRemarks.locked = True
        DBCboClientName.locked = True
        DCboCashType.locked = True
        Me.CboPaymentType.locked = True
        Me.DcboBox.locked = True
        Me.DcboBankName.locked = True
        Me.TxtChequeNumber.locked = True
        Me.DtpChequeDueDate.Enabled = False
        If rs.RecordCount < 1 Then
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        End If
        Fra(0).Enabled = False
        ChkTrans.Enabled = False
    Case "N"
'        Me.Caption = "«·„ﬁ»Ê÷« ( ÃœÌœ )"
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        Me.Cmd(5).Enabled = False
        Me.Cmd(7).Enabled = False
    '    Me.XPBtnMove(0).Enabled = False
    '    Me.XPBtnMove(1).Enabled = False
    '    Me.XPBtnMove(2).Enabled = False
    '    Me.XPBtnMove(3).Enabled = False
        XPDtbTrans.Enabled = True
        XPTxtVal.locked = False
        XPMTxtRemarks.locked = False
        DBCboClientName.locked = False
        XPDtbTrans.value = Date
        DCboCashType.locked = False
        DCboCashType.ListIndex = 0
        
        Me.CboPaymentType.locked = False
        Me.DcboBox.locked = False
        Me.DcboBankName.locked = False
        Me.TxtChequeNumber.locked = False
        Me.DtpChequeDueDate.Enabled = True
        
        Fra(0).Enabled = True
        ChkTrans.Enabled = True
    Case "E"
'        Me.Caption = "«·„ﬁ»Ê÷« (  ⁄œÌ· )"
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
        
        XPTxtVal.locked = False
        XPDtbTrans.Enabled = True
'        XPCboProfLevel.Locked = False
'        XPTxtProfMail.Locked = False
'        XPTxtPhone.Locked = False
'        XPTxtMobile.Locked = False
        XPMTxtRemarks.locked = False
        DBCboClientName.locked = False
        DCboCashType.locked = False
        Fra(0).Enabled = True
        ChkTrans.Enabled = True
        Me.CboPaymentType.locked = False
        Me.DcboBox.locked = False
        Me.DcboBankName.locked = False
        Me.TxtChequeNumber.locked = False
        Me.DtpChequeDueDate.Enabled = True
End Select
Exit Sub
ErrTrap:
End Sub

Private Sub TxtTransID_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If Me.TxtTransID.text <> "" Then
        If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
            Me.TxtTransSerial.text = GetTransIDSerial(1, Val(Me.TxtTransID.text))
        Else
            Me.TxtTransSerial.text = Me.TxtTransID.text
        End If
    End If
End If
End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
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
Public Sub Retrive(Optional Lngid As Long = 0)
Dim RsTemp As ADODB.Recordset
Dim StrSQL As String
Dim RsDev As ADODB.Recordset
Dim I As Integer
On Error GoTo ErrTrap
If rs.RecordCount < 1 Then
    XPTxtCurrent.Caption = 0
    XPTxtCount.Caption = 0
    Exit Sub
End If
If rs.EOF Or rs.BOF Then
    Exit Sub
Else
    If Lngid <> 0 Then
        rs.find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst
        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If
End If
 If Not IsNull(rs("general_cost_center").value) Then
    Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
 End If
Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", Val(rs("NoteID").value))
TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
txtperson.text = IIf(IsNull(rs("person").value), "", rs("person").value)

XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
txtAdv_payment_value.text = IIf(IsNull(rs("Adv_payment_value").value), "", Trim(rs("Adv_payment_value").value))

XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
'dcproject.BoundText = IIf(IsNull(Rs("Remark").value), "", Trim(Rs("Remark").value))

XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)

DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)


Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)

'-----------------------------------------------------------------------------
If IsNull(rs("NoteCashingType").value) Then
    Me.CboPaymentType.ListIndex = 0
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    
    'project_Expensen_account
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
ElseIf rs("NoteCashingType").value = 0 Then
    Me.CboPaymentType.ListIndex = 0
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
ElseIf rs("NoteCashingType").value = 1 Then
    Me.CboPaymentType.ListIndex = 1
    Me.DcboBox.BoundText = ""
    Me.DcboBankName.BoundText = rs("BankID").value
    Me.TxtChequeNumber.text = rs("ChqueNum").value
    Me.DtpChequeDueDate.value = rs("DueDate").value
End If

CboPayMentType_Change
'-----------------------------------------------------------------------------
If Not IsNull(rs("Transaction_ID").value) Then
    Me.ChkTrans.value = vbChecked
    'Me.ChkTrans.Enabled = True
    Set RsTemp = New ADODB.Recordset
    StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsTemp.BOF Or RsTemp.EOF) Then
        Me.TxtTransID.text = RsTemp("Transaction_ID").value
        Me.TxtTransSerial.text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)
        If Not (IsNull(RsTemp("Transaction_Type").value)) Then
            If RsTemp("Transaction_Type").value = 5 Then
                Me.CboTrans.ListIndex = 1
            ElseIf RsTemp("Transaction_Type").value = 2 Then
                Me.CboTrans.ListIndex = 0
            End If
        End If
    End If
ElseIf Not IsNull(rs("MaintananceID").value) Then
    Me.ChkTrans.value = vbChecked
    Me.CboTrans.ListIndex = 2
    Me.TxtTransID.text = rs("MaintananceID").value
    Me.TxtTransSerial.text = rs("MaintananceID").value
ElseIf Not IsNull(rs("RevenuesID").value) Then
    Me.DcboRevenuesTypes.BoundText = rs("RevenuesID").value
    Me.ChkTrans.value = vbUnchecked
    Me.CboTrans.ListIndex = -1
    Me.TxtTransID.text = ""
    Me.TxtTransSerial.text = ""
Else
    Me.ChkTrans.value = vbUnchecked
    Me.CboTrans.ListIndex = -1
    Me.TxtTransID.text = ""
    Me.TxtTransSerial.text = ""
End If

If DCboCashType.ListIndex = 5 Then
Dim My_SQL As String
  My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
      fill_combo Me.DBCboClientName, My_SQL
      
DBCboClientName.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
Dim cus_or_sub As Integer
cus_or_sub = IIf(IsNull(rs("cus_or_sub").value), 0, rs("cus_or_sub").value)
If cus_or_sub = 0 Then
Option4.value = True
Else
Option5.value = True
End If


End If

'-----------------------------------------------------------------------------
If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
    StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
    StrSQL = StrSQL + " Order By DEV_ID_Line_No "
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or rs.EOF) Then
        Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
        Me.lbl(33).Caption = RsDev("Account_Interval_ID").value
        RsDev.MoveFirst
        For I = 1 To 2 ' RsDev.RecordCount
            If RsDev("Credit_Or_Debit").value = 0 Then
                Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
            ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
            End If
            RsDev.MoveNext
        Next I
    End If
End If
'-----------------------------------------------------------------------------
ChkTrans_Click
XPTxtCurrent.Caption = rs.AbsolutePosition
XPTxtCount.Caption = rs.RecordCount
Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
Dim Msg As String
Dim RsTemp As New ADODB.Recordset
Dim StrSQL As String
Dim StrTemp As String
Dim LngDevID As Long
Dim RsDev As ADODB.Recordset

Dim BeginTrans As Boolean
'On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
    If DCboCashType.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„ﬁ»Ê÷«  "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboCashType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If
    If Me.DCboCashType.ListIndex = 3 Then
        If Val(Me.DcboRevenuesTypes.BoundText) = 0 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·≈Ì—«œ«  «·√Œ—Ï...!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            If Me.DcboRevenuesTypes.Visible = True Then
                DcboRevenuesTypes.SetFocus
                SendKeys "{F4}"
            End If
            Exit Sub
        End If
    End If
    If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Then
        If DBCboClientName.text = "" Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    End If
    
    
        If Me.DCboCashType.ListIndex = 5 Then
        If DBCboClientName.text = "" Then
            Msg = "ÌÃ» «Œ Ì«— «”„ ««·„‘—Ê⁄"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    End If
    
    
    If XPTxtVal.text = "" Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„ﬁ»Ê÷«  "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtVal.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(XPTxtVal.text) Then
        Msg = "ﬁÌ„… «·„ﬁ»Ê÷«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtVal.SetFocus
        SelectText XPTxtVal
        Exit Sub
    End If
    If Me.ChkTrans.value = vbChecked Then
        If Me.CboTrans.ListIndex = -1 Then
            Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboTrans.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        If Trim(Me.TxtTransSerial.text) = "" Then
            Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Sub
        Else
            If Me.CboTrans.ListIndex = 0 Then
                StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)
                If CheckDebitTrans(Val(StrTemp)) = False Then
                    Exit Sub
                End If
            ElseIf Me.CboTrans.ListIndex = 1 Then
                StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 5)
                If CheckDebitTrans(Val(StrTemp)) = False Then
                    Exit Sub
                End If
            ElseIf Me.CboTrans.ListIndex = 2 Then
                If CheckDebitMaintaince(Val(Me.TxtTransSerial.text)) = False Then
                    Exit Sub
                End If
            ElseIf Me.CboTrans.ListIndex = 3 Then
                Msg = "⁄›Ê« .. Ã«—Ï  ÿÊÌ— «·»—‰«„Ã .. ·⁄„· «·„ﬁ»Ê÷«  „‰ «·Œœ„« "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If
    End If
    If Me.CboPaymentType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄...!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPaymentType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If
    If Me.CboPaymentType.ListIndex = 0 Then
        If Me.DcboBox.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        If Me.DcboBankName.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBankName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        If Trim$(Me.TxtChequeNumber.text) = "" Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtChequeNumber.SetFocus
            Exit Sub
        End If
        If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DtpChequeDueDate.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    End If
    Cn.BeginTrans
    BeginTrans = True
    If TxtModFlg.text = "N" Then
        XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
        'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
        rs.AddNew
        rs("NoteID").value = Val(XPTxtID.text)
    ElseIf TxtModFlg.text = "E" Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
        Cn.Execute StrSQL, , adExecuteNoRecords



    End If
    rs("foxy_no").value = Val(Text1.text)
    rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("person").value = IIf(txtperson.text = "", "", Trim(txtperson.text))
    rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, Val(XPTxtVal.text))
    rs("Adv_payment_value").value = IIf(txtAdv_payment_value.text = "", Null, Val(txtAdv_payment_value.text))
    
'    Rs("Remark").value = IIf(dcproject.BoundText = "", "", Trim(dcproject.BoundText))
     rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))

    rs("NoteType").value = 4
    rs("NoteDate").value = XPDtbTrans.value
    Select Case DCboCashType.ListIndex
        Case 0, 1
            If Me.ChkTrans.value = vbChecked Then
                If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                    rs("Transaction_ID").value = Val(Me.TxtTransID.text)
                    rs("MaintananceID").value = Null
                ElseIf Me.CboTrans.ListIndex = 2 Then
                    rs("Transaction_ID").value = Null
                    rs("MaintananceID").value = Val(Me.TxtTransID.text)
                End If
            Else
                rs("Transaction_ID").value = Null
                rs("MaintananceID").value = Null
            End If
            rs("RevenuesID").value = Null
        Case 2
            rs("Transaction_ID").value = Null
            rs("MaintananceID").value = Null
            rs("RevenuesID").value = Null
        Case 3
            rs("RevenuesID").value = Val(Me.DcboRevenuesTypes.BoundText)
            rs("Transaction_ID").value = Null
            rs("MaintananceID").value = Null
        Case 4
 '       Set rs1 = New ADODB.Recordset
 '       StrSQL = "select * From Transactions"
 '       rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
 '       rs1.AddNew
 '       rs1("Transaction_ID").value = Val(XPTxtBillID.text)
 '       rs1("Transaction_Date").value = XPDtbTrans.value
 '       rs1("Transaction_Type").value = 23
 '       rs1.update
 '
 '        Rs("Transaction_ID").value = Val(XPTxtBillID.text)
 '
    End Select
    rs("CashingType").value = DCboCashType.ListIndex
    
    If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Then
        rs("CusID").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
     
    ElseIf Me.DCboCashType.ListIndex = 5 Then
           Dim x As Double
            
            If Option4.value = True Then
                x = get_project_customer_id(DBCboClientName.BoundText, "End_user_Account")
            Else
                x = get_project_customer_id(DBCboClientName.BoundText, "sub_contractor_Account")
            End If
    rs("CusID").value = x
     
    Else
        rs("CusID").value = Null
    End If
    '--------------------------------------------------------------------------
    'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
    If Me.CboPaymentType.ListIndex = 0 Then
        rs("NoteCashingType").value = 0
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        rs("NoteCashingType").value = 1
        rs("BoxID").value = Null
        rs("BankID").value = Val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
    End If
    '--------------------------------------------------------------------------
    rs("UserID").value = user_id
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    rs("numbering_type1").value = sand_numbering_type(2) '”‰œ «·ﬁ»÷
    
    If DCboCashType.ListIndex = 5 Then
        rs("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
    End If
    rs("sanad_year").value = year(XPDtbTrans.value)
    rs("sanad_month").value = Month(XPDtbTrans.value)
    If DCboCashType.ListIndex = 5 Then
    rs("note_value_by_characters").value = WriteNo(Val(Me.XPTxtVal.text) * 2, 0, True)
    Else
    rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    End If
    If Option4.value = True Then
      rs("cus_or_sub").value = 0 '⁄„Ì· ‰Â«∆Ì
    Else
     rs("cus_or_sub").value = 1 '⁄„Ì· »«ÿ‰
    End If
    
    
    rs.update
    



    '==========================================================================
  
    
     Line1 = setfoxy_Line
    Line2 = setfoxy_Line
         Line3 = setfoxy_Line
    Line4 = setfoxy_Line
    ' ”ÃÌ· ﬁÌÊœ
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
        Set RsDev = New ADODB.Recordset
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '«·ÿ—› «·„œÌ‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("DEV_ID_Line_No1").value = Line1
            
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = Val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = Val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                   If DCboCashType.ListIndex = 5 Then
               RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
           End If
        RsDev.update
        '«·ÿ—› «·œ«∆‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 2
            RsDev("DEV_ID_Line_No1").value = Line2
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = Val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
           ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            RsDev("Notes_ID").value = Val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
           If DCboCashType.ListIndex = 5 Then
               RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
           End If
    
        RsDev.update
        If DCboCashType.ListIndex = 5 Then
       '«·„‘«—Ì⁄
       Dim account_codeLegal As String
       Dim account_codeREVENUE_account As String
       
       account_codeLegal = get_project_Account(Val(DBCboClientName.BoundText), "legal")
       account_codeREVENUE_account = get_project_Account(Val(DBCboClientName.BoundText), "REVENUE_account")
       If account_codeLegal = "" Or account_codeREVENUE_account = "" Then GoTo LL
       
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 3
            RsDev("DEV_ID_Line_No1").value = Line3
            
            RsDev("Account_Code").value = account_codeLegal
            RsDev("Value").value = Val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = Val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                   If DCboCashType.ListIndex = 5 Then
               RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
           End If
        RsDev.update
        '«·ÿ—› «·œ«∆‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 4
            RsDev("DEV_ID_Line_No1").value = Line4
            RsDev("Account_Code").value = account_codeREVENUE_account
            RsDev("Value").value = Val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
           ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            RsDev("Notes_ID").value = Val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
           If DCboCashType.ListIndex = 5 Then
               RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
           End If
    
        RsDev.update
LL:
  End If
        LblDevID.Caption = LngDevID
        lbl(33).Caption = SystemOptions.SysCurrentAccountIntervalID
    End If
    '==========================================================================
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
   If DCboCashType.ListIndex = 5 Then
    saveprojectBillPayment TxtNoteSerial.text, Val(XPTxtVal.text)
    End If
    
    If Me.ChkTrans.value = vbUnchecked Then
        Me.CboTrans.ListIndex = -1
        Me.TxtTransSerial.text = ""
        Me.TxtTransID.text = ""
    End If
    Select Case Me.TxtModFlg.text
    Case "N"
        Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
        Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
        Cmd_Click (0)
        Exit Sub
        End If
        
    Case "E"
        MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End Select
    
    If Me.DcCostCenter.BoundText <> "" Then
     save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "„ﬁ»Ê÷« ", Me.XPDtbTrans.value
    End If
        
    TxtModFlg.text = "R"
End If

WriteCustomerBal
WriteInfo
   If Option1.value = True Then
     FIFO_FUNCTION Val(DBCboClientName.BoundText)
   End If
   
   If Option2.value Then
    Distribute_to_bills Me.lblsqlstring, Val(DBCboClientName.BoundText)
   End If
   
  TxtModFlg.text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Public Function save_General_cost_center(cost_center_id As String, cost_center, opr_type As String, record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
 Dim I As Integer
 Dim rs As New ADODB.Recordset
 
 Dim StrSQL As String

StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    
 
rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

 'ÿ—› „œÌ‰
        rs.AddNew
        rs("cost_center_id").value = cost_center_id
        rs("cost_center").value = cost_center
        rs("value").value = XPTxtVal.text
        rs("depit_or_credit").value = "„œÌ‰"
        rs("opr_id").value = Me.Text1.text
        rs("kedno").value = Me.Text1.text
        
        rs("opr_type").value = opr_type
        rs("account_name").value = DcboDebitSide.text
        rs("account_no").value = DcboDebitSide.BoundText
        rs("line_no").value = Line1
        rs("record_date").value = record_date
        rs.update
 'ÿ—› œ«∆‰
        rs.AddNew
        rs("cost_center_id").value = cost_center_id
        rs("cost_center").value = cost_center
        rs("value").value = XPTxtVal.text
        rs("depit_or_credit").value = "œ«∆‰"
        rs("opr_id").value = Me.Text1.text
        rs("kedno").value = Me.Text1.text
        
        rs("opr_type").value = opr_type
        rs("account_name").value = DcboCreditSide.text
        rs("account_no").value = DcboCreditSide.BoundText
        rs("line_no").value = Line2
        rs("record_date").value = record_date
        rs.update
 
rs.Close
End Function

Function change_adv_payment_value(note_id As Double, value As Double)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
Dim I As Integer


sql = "SELECT * from notes   where  NoteID=" & note_id
 
  Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
  If Rs3.RecordCount = 0 Then Exit Function
  Rs3("Adv_payment_value").value = value
  Rs3.update
  
  
End Function
Function Distribute_to_bills(Sql1 As String, CusID As Double)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
Dim I As Integer

sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where  requiredvalue>0 and " & Sql1
 
  Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
  If Rs3.RecordCount = 0 Then Exit Function
 Dim total_value As Double
 Dim current_value As Double
  total_value = Val(txtAdv_payment_value.text)
  
  For I = 1 To Rs3.RecordCount
        If total_value > Rs3("requiredvalue") Then
        current_value = Rs3("requiredvalue")
        total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
         current_value = total_value
         total_value = 0
        ElseIf total_value = 0 Then
         Exit Function
        End If
 
  
  Add_new_notes Me.XPDtbTrans, 2000, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, Val(DCboUserName.BoundText)
  Rs3.MoveNext
  Next I
  txtAdv_payment_value.text = total_value
  change_adv_payment_value XPTxtID.text, total_value


 ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  
 ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
Rs3.Close


 
End Function
Function FIFO_FUNCTION(CusID As Double)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
Dim I As Integer

sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where   (cusid=" & CusID & " and requiredvalue>0)"
 
  Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
  If Rs3.RecordCount = 0 Then Exit Function
 Dim total_value As Double
 Dim current_value As Double
  total_value = Val(txtAdv_payment_value.text)
  
  For I = 1 To Rs3.RecordCount
        If total_value > Rs3("requiredvalue") Then
        current_value = Rs3("requiredvalue")
        total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
         current_value = total_value
         total_value = 0
        ElseIf total_value = 0 Then
         Exit Function
        End If
 
  
  Add_new_notes Me.XPDtbTrans, 2000, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, Val(DCboUserName.BoundText)
  Rs3.MoveNext
  Next I
 ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  txtAdv_payment_value.text = total_value
  change_adv_payment_value XPTxtID.text, total_value
 ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
Rs3.Close


 

End Function
Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Integer, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
            Dim RsDev As New ADODB.Recordset
        RsDev.Open "notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '
        
         RsDev.AddNew
      
            RsDev("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
            RsDev("NoteSerial").value = TxtNoteSerial.text ' CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=2000"))
              
            RsDev("NoteDate").value = NoteDate
            RsDev("NoteType").value = NoteType
           
            RsDev("Note_Value").value = Note_Value
            RsDev("Transaction_ID").value = Transaction_ID
            RsDev("CusID").value = CusID
            RsDev("BoxID").value = BoxID
            RsDev("UserID").value = UserID
            RsDev("displayed").value = 0
            
            
            
           
        RsDev.update
 

End Function
Private Sub Undo()
On Error GoTo ErrTrap
Select Case TxtModFlg.text
    Case "N"
         clear_all Me
         Me.TxtModFlg.text = "R"
         XPBtnMove_Click (1)
    Case "E"
         rs.find "NoteID='" & Val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst
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
Private Sub Del_Trans()
Dim Msg As String
On Error GoTo ErrTrap
If XPTxtID.text <> "" Then
    If Me.CboPaymentType.ListIndex = 0 Then
        If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), Date, False) = False Then
            Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
            Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If
    Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
    Msg = Msg + (TxtNoteSerial.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
       If Not rs.RecordCount < 1 Then
          '  Rs.Delete
       Dim StrSQL As String
       StrSQL = "Delete From notes  Where  (NoteType=2000 OR NoteType=4 ) AND  NoteSerial=" & Val(TxtNoteSerial.text)
       Cn.Execute StrSQL, , adExecuteNoRecords
        
       StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
       Cn.Execute StrSQL, , adExecuteNoRecords
            rs.MoveFirst
            If rs.RecordCount < 1 Then
                clear_all Me
                
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
            Else
            clear_all Me
                Retrive
            End If
            '--------
            WriteInfo
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
            vbExclamation, App.Title
    rs.CancelUpdate
End Sub
Private Sub ChangeLang()

Dim XPic As IPictureDisp
Set XPic = Me.XPBtnMove(1).ButtonImage
Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
Set Me.XPBtnMove(2).ButtonImage = XPic
Set XPic = Me.XPBtnMove(0).ButtonImage
Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
Set Me.XPBtnMove(3).ButtonImage = XPic
lbl(35).Caption = "Adv. Payment"
Frame1.Caption = "Options"
Option3.Caption = "Adv. Payment"
Option2.Caption = "Select Invoice"
 ALLButton3.Caption = "Select"
 lbl(22).Caption = "Current Week"
  Label8.Caption = "General C.C."
 lbl(36).Caption = "From"
 
 Cmd(9).Caption = "GL Print"
 
 
Frame2.Caption = "Project"
Option4.Caption = "End User"
Option5.Caption = "Sub-contractor"


LblLink.Visible = False
lbl(18).Visible = False
ALLButton1.Caption = "Installment view"
ALLButton2.Caption = "debt Voucher"
Me.Caption = "Receipts"
Ele(1).Caption = Me.Caption
lbl(4).Caption = "Opr Code"
lbl(1).Caption = "Date"
lbl(0).Caption = "Type"
lbl(3).Caption = "Name"
lbl(2).Caption = "Value"
lbl(14).Caption = "Payemnt Method"
lbl(9).Caption = "Box Name"
lbl(15).Caption = "Bank Name"
lbl(16).Caption = "Cheque #"
lbl(17).Caption = "Cheque Name"
lbl(5).Caption = "Note"
ChkTrans.Caption = "From bill"
lbl(12).Caption = "Bill type"
lbl(10).Caption = "Bill #"
lbl(13).Caption = "Current Balance"
FraInfo.Caption = "Information"
lbl(22).Caption = "Current Week"

lbl(23).Caption = "Today Receipts "
lbl(27).Caption = "Cash"
lbl(28).Caption = "Cheque"

lbl(19).Caption = "Week Receipts "

lbl(21).Caption = "Cash"
lbl(24).Caption = "Cheque"

lbl(20).Caption = "Month Receipts "

lbl(25).Caption = "Cash"
lbl(26).Caption = "Cheque"
Fra(1).Caption = "GL"

lbl(30).Caption = "GL#"
lbl(29).Caption = "Interval"

lbl(31).Caption = "Depit"
lbl(32).Caption = "Credit"
Cmd(8).Caption = "Table view"
lbl(8).Caption = "By"
lbl(7).Caption = "Current "
lbl(6).Caption = "Records Count "

Cmd(0).Caption = "New"
Cmd(1).Caption = "Edit"
Cmd(2).Caption = "Save"
Cmd(3).Caption = "Undo"
Cmd(4).Caption = "Delete"
Cmd(5).Caption = "Search"
Cmd(7).Caption = "Print"
Cmd(6).Caption = "Exit"
CmdHelp.Caption = "Help"
DCboCashType.Clear
DCboCashType.AddItem "To Customer"
DCboCashType.AddItem "To Vendor"
DCboCashType.AddItem "Sub-contractor"
DCboCashType.AddItem "Another Revenues"
DCboCashType.AddItem "Advanced Payment"
DCboCashType.AddItem "Projects"



With Me.CboPaymentType
    .Clear
    .AddItem "Cash"
    .AddItem "Cheque"
End With

With Me.CboTrans
    .Clear
    .AddItem "Sales invoice"
    .AddItem "Returned purchases"
    .AddItem "Delivery of maintenance for a client"
    .AddItem "Services"
End With

 
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTrap
If KeyCode = vbKeyReturn Then
    If Me.TxtModFlg.text = "R" Then
        Cmd_Click (0)
    Else
        SendKeys "{TAB}"
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
If Shift = 2 Then
    If KeyCode = vbKeyX Then
        If Cmd(6).Enabled = False Then Exit Sub
        'Cmd_Click (6)
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub AddTip()
Dim Wrap As String
On Error GoTo ErrTrap
Wrap = Chr(13) + Chr(10)
Set TTP = New clstooltip
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(0), _
    "ÃœÌœ ..." & Wrap & _
    "· ”ÃÌ· »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(1), _
    " ⁄œÌ· ..." & Wrap & _
    "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(2), _
    "Õ›Ÿ ..." & Wrap & _
    "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & _
     "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(3), _
    " —«Ã⁄ ..." & Wrap & _
    "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
     "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
 With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(4), _
    "Õ–› ..." & Wrap & _
    "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(6), _
    "Œ—ÊÃ ..." & Wrap & _
    "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(1), _
    "«·√Ê· ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(0), _
    "«·”«»ﬁ ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(3), _
    "«· «·Ì ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(2), _
    "«·√ŒÌ— ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl CmdHelp, _
    "„”«⁄œ… ..." & Wrap & _
    "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
    "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
    "≈÷€ÿ Â‰«" & Wrap, True
End With
Exit Sub
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


Dim IntResult As String
Dim StrMSG As String
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

Private Sub XPTxtVal_Change()
Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)

If TxtModFlg.text = "N" Then
txtAdv_payment_value.text = XPTxtVal.text
End If

End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Function CheckDebitTrans(LngTransID As Long) As Boolean
Dim Msg As String
Dim RsTemp As ADODB.Recordset
Dim DblCreditNoteValue As Double
Dim LngDebitNoteID As Long
Dim StrSQL As String

CheckDebitTrans = False
If LngTransID = 0 Then
    Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
    Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtTransSerial.SetFocus
    Exit Function
ElseIf LngTransID <> 0 Then
    Set RsTemp = New ADODB.Recordset
    StrSQL = "Select CusID,PaymentType From Transactions where Transaction_ID=" & LngTransID & ""
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not (RsTemp.BOF Or RsTemp.EOF) Then
        If RsTemp("PaymentType").value = 0 Then
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & Chr(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
        If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
        If LngTransID <> Val(Me.TxtTransID.text) Then
            Me.TxtTransID.text = LngTransID
        End If
        
        DblCreditNoteValue = 0
        StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
        "Transactions.Transaction_Type, Transactions.PaymentType, " & _
        "Notes.Note_Value, Notes.NoteID "
        StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & _
        "Notes.Transaction_ID WHERE (Notes.NoteType=1) AND Transactions.Transaction_ID= " & LngTransID & ""
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            LngDebitNoteID = RsTemp("NoteID").value
            DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
            '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
            'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
            StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If RsTemp.RecordCount > 0 Then
                    Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                    Msg = Msg & Chr(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                End If
            End If
        Else
        'LngDebitNoteID
            Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Function
        End If
        If DblCreditNoteValue < Val(Me.XPTxtVal.text) Then
            Msg = "⁄›Ê« ..."
            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
            Msg = Msg & Chr(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
            Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.XPTxtVal.SetFocus
            Exit Function
        End If
        Set RsTemp = New ADODB.Recordset
        StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
        "Transactions.Transaction_Type, Transactions.PaymentType," & _
        "Sum(Notes.Note_Value) AS SumNote_Value "
        StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & _
        "Notes.Transaction_ID " & _
        " Where ((Notes.NoteType = 4 OR Notes.NoteType = 9) And Transactions.Transaction_ID = " & LngTransID & ")"
        If Me.TxtModFlg.text = "E" Then
            StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
        End If
        StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
        "Transactions.Transaction_Type, Transactions.PaymentType "
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                Msg = "⁄›Ê« ...!!!!!" & Chr(13)
                Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            ElseIf RsTemp("SumNote_Value").value + Val(Me.XPTxtVal.text) > _
                DblCreditNoteValue Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                Msg = Msg & Chr(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                Msg = Msg & Chr(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            End If
        End If
    Else
        Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
        Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Exit Function
    End If
End If
CheckDebitTrans = True
Exit Function
ErrTrap:
End Function

Private Function CheckDebitMaintaince(LngTransID As Long) As Boolean
Dim Msg As String
Dim RsTemp As ADODB.Recordset
Dim DblCreditNoteValue As Double
Dim LngDebitNoteID As Long
Dim StrSQL As String

CheckDebitMaintaince = False
If LngTransID = 0 Then
    Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
    Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtTransSerial.SetFocus
    Exit Function
ElseIf LngTransID <> 0 Then
    Set RsTemp = New ADODB.Recordset
    StrSQL = "Select CusID,PaymentType From TblMaintenece where MaintananceID=" & LngTransID & ""
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not (RsTemp.BOF Or RsTemp.EOF) Then
        If RsTemp("PaymentType").value = 0 Then
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & Chr(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
        If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
        If LngTransID <> Val(Me.TxtTransID.text) Then
            Me.TxtTransID.text = LngTransID
        End If
        
        DblCreditNoteValue = 0
        StrSQL = "SELECT Notes.Note_Value, Notes.NoteID, TblMaintenece.MaintananceID," & _
        "TblMaintenece.PaymentType, TblMaintenece.MType "
        StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON " & _
        "TblMaintenece.MaintananceID = Notes.MaintananceID " & _
        " WHERE (((Notes.NoteType)=1)) AND TblMaintenece.MaintananceID=" & LngTransID & ""
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            LngDebitNoteID = RsTemp("NoteID").value
            DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
            '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
            'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
            StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If RsTemp.RecordCount > 0 Then
                    Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                    Msg = Msg & Chr(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                End If
            End If
        Else
        'LngDebitNoteID
            Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Function
        End If
        If DblCreditNoteValue < Val(Me.XPTxtVal.text) Then
            Msg = "⁄›Ê« ..."
            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
            Msg = Msg & Chr(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
            Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.XPTxtVal.SetFocus
            Exit Function
        End If
        Set RsTemp = New ADODB.Recordset
        
        StrSQL = "SELECT  TblMaintenece.MaintananceID," & _
        "TblMaintenece.MType, TblMaintenece.PaymentType," & _
        "Sum(Notes.Note_Value) AS SumNote_Value "
        StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON TblMaintenece.MaintananceID =" & _
        "Notes.MaintananceID " & _
        " Where ((Notes.NoteType = 4) And TblMaintenece.MaintananceID = " & LngTransID & ")"
        If Me.TxtModFlg.text = "E" Then
            StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
        End If
        StrSQL = StrSQL + " GROUP BY TblMaintenece.MaintananceID," & _
        "TblMaintenece.MType, TblMaintenece.PaymentType"
        
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                Msg = "⁄›Ê« ...!!!!!"
                Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            ElseIf RsTemp("SumNote_Value").value + Val(Me.XPTxtVal.text) > _
                DblCreditNoteValue Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  „”»ﬁ« ·Â–Â «·›« Ê—…"
                Msg = Msg & Chr(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                Msg = Msg & Chr(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            End If
        End If
    Else
        Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
        Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Exit Function
    End If
End If
CheckDebitMaintaince = True
Exit Function
ErrTrap:
End Function

Public Function CheckDebitService()

End Function

Private Sub WriteCustomerBal()
Dim StrTemp As String
Dim SngCusBegainAccount As Single
Me.MousePointer = vbArrowHourglass
If Val(Me.DBCboClientName.BoundText) <> 0 Then
    SngCusBegainAccount = GetCustomerAccount(Val(Me.DBCboClientName.BoundText), True)
    If SngCusBegainAccount < 0 Then
        StrTemp = Abs(SngCusBegainAccount) & " „œÌ‰ "
    ElseIf SngCusBegainAccount > 0 Then
        StrTemp = Abs(SngCusBegainAccount) & " œ«∆‰ "
    Else
        StrTemp = Abs(SngCusBegainAccount) & " Œ«·’ "
    End If
Else
    StrTemp = "0" & " Œ«·’ "
End If
Me.MousePointer = vbDefault
Me.LblLink.Caption = StrTemp
End Sub

Private Sub WriteInfo()
Dim rs As ADODB.Recordset
Dim StrSQL As String
Dim StartWeekDate As Date
Dim EndWeekDate As Date
Dim StrTemp As String
Dim I As Integer

StartWeekDate = GetWeekStartEND(Date, 0)
EndWeekDate = DateAdd("d", 7, StartWeekDate)
StrTemp = "«·≈”»Ê⁄ «·Õ«·Ï „‰ " & DisplayDate(StartWeekDate)
StrTemp = StrTemp & " ≈·Ï " & DisplayDate(EndWeekDate)
Me.lbl(22).Caption = StrTemp
For I = LblLinkInfo.LBound To LblLinkInfo.UBound
    LblLinkInfo(I).Caption = "0"
Next I
'------------------------------------------------------------------------------
'„ﬁ»Ê÷«  «·ÌÊ„
StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
StrSQL = StrSQL + " From Notes "
StrSQL = StrSQL + " Where (NoteType = 4) "
StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
StrSQL = StrSQL + " GROUP BY NoteCashingType"
StrSQL = StrSQL + " Order BY NoteCashingType"
Set rs = New ADODB.Recordset
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (rs.BOF Or rs.EOF) Then
    rs.MoveFirst
    For I = 0 To rs.RecordCount - 1
        If rs("NoteCashingType").value = 0 Then
            Me.LblLinkInfo(0).Caption = rs("SumX").value
        ElseIf rs("NoteCashingType").value = 1 Then
            Me.LblLinkInfo(1).Caption = rs("SumX").value
        End If
        rs.MoveNext
    Next
    Me.LblLinkInfo(6).Caption = Val(Me.LblLinkInfo(0).Caption) + _
    Val(Me.LblLinkInfo(1).Caption)
Else
    Me.LblLinkInfo(0).Caption = 0
    Me.LblLinkInfo(1).Caption = 0
    Me.LblLinkInfo(6).Caption = 0
End If
'------------------------------------------------------------------------------
'„ﬁ»Ê÷«  «·√”»Ê⁄ «·Õ«·Ï
StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
StrSQL = StrSQL + " From Notes "
StrSQL = StrSQL + " Where (NoteType = 4) "
StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(StartWeekDate, True)
StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(EndWeekDate, True)
StrSQL = StrSQL + " GROUP BY NoteCashingType"
StrSQL = StrSQL + " Order BY NoteCashingType"
Set rs = New ADODB.Recordset
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (rs.BOF Or rs.EOF) Then
    rs.MoveFirst
    For I = 0 To rs.RecordCount - 1
        If rs("NoteCashingType").value = 0 Then
            Me.LblLinkInfo(2).Caption = rs("SumX").value
        ElseIf rs("NoteCashingType").value = 1 Then
            Me.LblLinkInfo(3).Caption = rs("SumX").value
        End If
        rs.MoveNext
    Next
    Me.LblLinkInfo(7).Caption = Val(Me.LblLinkInfo(2).Caption) + _
    Val(Me.LblLinkInfo(3).Caption)
Else
    Me.LblLinkInfo(0).Caption = 0
    Me.LblLinkInfo(1).Caption = 0
    Me.LblLinkInfo(7).Caption = 0
End If
'------------------------------------------------------------------------------
'„ﬁ»Ê÷«  «·‘Â— «·Õ«·Ï
StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
StrSQL = StrSQL + " From Notes "
StrSQL = StrSQL + " Where (NoteType = 4) "
StrSQL = StrSQL + " AND Month(NoteDate)=" & Month(Date) & ""
StrSQL = StrSQL + " AND Year(NoteDate)=" & year(Date) & ""
StrSQL = StrSQL + " GROUP BY NoteCashingType"
StrSQL = StrSQL + " Order BY NoteCashingType"
Set rs = New ADODB.Recordset
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (rs.BOF Or rs.EOF) Then
    rs.MoveFirst
    For I = 0 To rs.RecordCount - 1
        If rs("NoteCashingType").value = 0 Then
            Me.LblLinkInfo(4).Caption = rs("SumX").value
        ElseIf rs("NoteCashingType").value = 1 Then
            Me.LblLinkInfo(5).Caption = rs("SumX").value
        End If
        rs.MoveNext
    Next
    Me.LblLinkInfo(8).Caption = Val(Me.LblLinkInfo(4).Caption) + _
    Val(Me.LblLinkInfo(5).Caption)
Else
    Me.LblLinkInfo(4).Caption = 0
    Me.LblLinkInfo(5).Caption = 0
    Me.LblLinkInfo(8).Caption = 0
End If
End Sub
