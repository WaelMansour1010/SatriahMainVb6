VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmExpenses1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„’—Ê›« "
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   HelpContextID   =   280
   Icon            =   "FrmExpenses1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10650
   Begin VB.TextBox txtto 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   68
      Top             =   2760
      Width           =   4755
   End
   Begin VB.Frame FraNote 
      BackColor       =   &H00E2E9E9&
      Height          =   1725
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   1680
      Width           =   4155
      Begin VB.TextBox TxtChequeNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   840
         Width           =   2685
      End
      Begin MSComCtl2.DTPicker DtpChequeDueDate 
         Height          =   315
         Left            =   30
         TabIndex        =   60
         Top             =   1140
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104071169
         CurrentDate     =   39614
      End
      Begin MSDataListLib.DataCombo DcboBankName 
         Height          =   315
         Left            =   30
         TabIndex        =   61
         Top             =   480
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   0
         TabIndex        =   66
         Top             =   120
         Width           =   2835
         _ExtentX        =   5001
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
         Index           =   19
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   285
         Index           =   18
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   285
         Index           =   17
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Œ“‰…"
         Height          =   285
         Index           =   16
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   6720
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   840
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   1335
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
      Height          =   1035
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   8700
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   37
         Top             =   270
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
         TabIndex        =   39
         Top             =   600
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
         Index           =   12
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   13
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   11
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   9
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2040
      Width           =   4755
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   6360
      Width           =   1905
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _cx             =   18653
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
      Picture         =   "FrmExpenses1.frx":038A
      Caption         =   "«·„’—Ê›«  "
      Align           =   0
      AutoSizeChildren=   0
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   7
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
         ButtonImage     =   "FrmExpenses1.frx":1064
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
         TabIndex        =   8
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
         ButtonImage     =   "FrmExpenses1.frx":13FE
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
         Left            =   2220
         TabIndex        =   9
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
         ButtonImage     =   "FrmExpenses1.frx":1798
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
         TabIndex        =   10
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
         ButtonImage     =   "FrmExpenses1.frx":1B32
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   1200
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
         Left            =   0
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
      Begin VB.Label LblShortcutKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
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
         Height          =   195
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Top             =   810
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   104071169
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   11280
      TabIndex        =   2
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   7680
      TabIndex        =   17
      Top             =   7290
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   30
      TabIndex        =   24
      Top             =   840
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   7980
      TabIndex        =   26
      Top             =   6750
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   27
      Top             =   6750
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      Height          =   495
      Index           =   2
      Left            =   6270
      TabIndex        =   28
      Top             =   6720
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      Height          =   495
      Index           =   3
      Left            =   5115
      TabIndex        =   29
      Top             =   6750
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      Height          =   495
      Index           =   4
      Left            =   4200
      TabIndex        =   30
      Top             =   6750
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   31
      Top             =   6750
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
   Begin ImpulseButton.ISButton CmdHelp 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1080
      TabIndex        =   32
      Top             =   6750
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   5
      Left            =   3150
      TabIndex        =   33
      Top             =   6750
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
   Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
      Height          =   2340
      Left            =   240
      TabIndex        =   45
      Top             =   3840
      Width           =   10320
      _cx             =   18203
      _cy             =   4128
      Appearance      =   1
      BorderStyle     =   1
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
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
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
      Rows            =   10
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses1.frx":1ECC
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      Begin VB.PictureBox PicDes 
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   240
         RightToLeft     =   -1  'True
         ScaleHeight     =   1635
         ScaleWidth      =   2925
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox TxtDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1125
            Left            =   30
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label LblDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
            ForeColor       =   &H0000C8FF&
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   0
            Width           =   2445
         End
      End
      Begin VDSCOMBOLibCtl.SmartCombo CboDes 
         Height          =   315
         Left            =   240
         TabIndex        =   52
         ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
         Top             =   600
         Visible         =   0   'False
         Width           =   2955
         _cx             =   1973752924
         _cy             =   1973748268
         Alignment       =   0
         Appearance      =   3
         AutoSearch      =   0   'False
         BackColor       =   -2147483624
         BackgroundColor =   -2147483633
         BorderColor     =   0
         BorderVisible   =   -1  'True
         Caption         =   "SmartCombo1"
         CaptionAlignment=   4
         CaptionBackColor=   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   -2147483630
         CaptionHeight   =   15
         CaptionOnTop    =   0   'False
         CaptionMultiLine=   0
         Checkbox3D      =   0   'False
         CheckboxAlignment=   5
         CheckboxBackColor=   16777215
         CheckboxSize    =   13
         CheckboxValue   =   0
         BrowsePictureAlignment=   5
         BrowsePictureStretchH=   0
         BrowsePictureStretchV=   0
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
         ForeColor       =   0
         Gap             =   0
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   0
         MultiLine       =   0
         OnFocus         =   3
         PasswordChar    =   ""
         Picture         =   "FrmExpenses1.frx":1FA3
         PictureAlignment=   5
         PictureBackColor=   -2147483624
         PictureStretchH =   0
         PictureStretchV =   0
         Redraw          =   -1  'True
         ScrollBar       =   0
         Style           =   0
         Text            =   ""
         UnderLine       =   0   'False
         Enabled0        =   -1  'True
         Position0       =   0
         Tip0            =   "Caption"
         Visible0        =   0   'False
         Width0          =   90
         Enabled1        =   -1  'True
         Position1       =   1
         Tip1            =   ""
         Visible1        =   -1  'True
         Width1          =   32
         Enabled2        =   -1  'True
         Position2       =   2
         Tip2            =   "Check Box (Space, Ctrl + Space)"
         Visible2        =   0   'False
         Width2          =   16
         Enabled3        =   -1  'True
         Position3       =   3
         Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
         Visible3        =   -1  'True
         Width3          =   145
         Enabled4        =   -1  'True
         Position4       =   4
         Tip4            =   "Left Spinner (Alt + Left)"
         Visible4        =   0   'False
         Width4          =   16
         Enabled5        =   -1  'True
         Position5       =   5
         Tip5            =   "Right Spinner (Alt + Right)"
         Visible5        =   0   'False
         Width5          =   16
         Enabled6        =   -1  'True
         Position6       =   6
         Tip6            =   "Up Spinner (Ctrl + Up)"
         Visible6        =   0   'False
         Width6          =   16
         Enabled7        =   -1  'True
         Position7       =   7
         Tip7            =   "Down Spinner (Ctrl + Down)"
         Visible7        =   0   'False
         Width7          =   16
         Enabled8        =   -1  'True
         Position8       =   8
         Tip8            =   "Browse (Alt + Enter)"
         Visible8        =   0   'False
         Width8          =   16
         Enabled9        =   -1  'True
         Position9       =   9
         Tip9            =   " (Alt + Down, F4)"
         Visible9        =   -1  'True
         Width9          =   16
         Enabled10       =   -1  'True
         Position10      =   10
         Tip10           =   "Right Arrow (Alt + >)"
         Visible10       =   0   'False
         Width10         =   16
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   8760
      TabIndex        =   46
      Top             =   6840
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "„—«ﬂ“ «· ﬂ·›…"
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
      BCOL            =   255
      BCOLO           =   192
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "FrmExpenses1.frx":253D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo dcproject 
      Height          =   315
      Left            =   1680
      TabIndex        =   54
      Top             =   1680
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   55
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â"
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
      Height          =   375
      Index           =   9
      Left            =   3120
      TabIndex        =   67
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      Caption         =   "»‰«¡ ⁄·Ï"
      Height          =   285
      Index           =   0
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      Height          =   195
      Index           =   15
      Left            =   9060
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   1320
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‘—Ê⁄"
      Height          =   255
      Index           =   14
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   1680
      Width           =   915
   End
   Begin VB.Image ImgNote 
      Height          =   240
      Left            =   0
      Picture         =   "FrmExpenses1.frx":2559
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
      Height          =   255
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LblValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6300
      Width           =   5175
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   7290
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   7290
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   435
      Index           =   6
      Left            =   690
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   7290
      Width           =   165
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
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
      Height          =   435
      Index           =   7
      Left            =   1500
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   7290
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   390
      Index           =   8
      Left            =   9345
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   7305
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   825
      Width           =   555
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6480
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„’—Ê›« "
      Height          =   285
      Index           =   3
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   285
      Index           =   4
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2160
      Width           =   1515
   End
End
Attribute VB_Name = "FrmExpenses1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String

Private Sub ALLButton1_Click()
'On Error GoTo ErrTrap
Dim opr_id As Double
'If Me.TxtModFlg.text = "N" Then
opr_id = Val(Me.TxtNoteSerial.text)
'Else
'opr_id = TxtDEV_NO.text
'End If


       If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
            If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) = "0" Then
            marakes_taklefa_tawze3.Show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.LineNo = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))
        
            End If

            marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
            marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
            marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
            marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
            marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))
            marakes_taklefa_tawze3.Adodc3.Refresh
        '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

End If
Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_AfterAutoCloseUp()
PutData
CboDes.Visible = False
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

Private Sub Cmd_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
        TxtModFlg.text = "N"
        clear_all Me
        XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
        Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=3"))
        Me.DCboUserName.BoundText = user_id
        XPDtbTrans.SetFocus
        Fg_Journal.Clear flexClearScrollable, flexClearEverything
          Fg_Journal.Rows = 3
          Fg_Journal.Enabled = True
    Case 1
        If DoPremis(Do_Edit, Me.name, True) = False Then
            Exit Sub
        End If
        TxtModFlg.text = "E"
        Me.DCboUserName.BoundText = user_id
        Fg_Journal.Rows = Fg_Journal.Rows + 1
        Fg_Journal.Enabled = True
         

    Case 2
    If TxtSerial.text = "" Then
         If sand_numbering = "error" Then
         MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
         Else
         
         If sand_numbering = "" Then
         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
         Else
         TxtSerial.text = sand_numbering
         End If
         End If
    End If
    
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
        FrmNotesSearch.SearchType = 3
        FrmNotesSearch.Show vbModal
    Case 6
        Unload Me
    Case 7
        ViewDataList
    Case 8
         print_report (TxtSerial.text)
         Case 9
         print_Cheque TxtChequeNumber.text, get_Cheque_report_no(Val(DcboBankName.BoundText))
    
End Select
Exit Sub
ErrTrap:
End Sub
Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "")
    hide_logo = True
Dim MySQL As String
Dim RsData As New ADODB.Recordset
Dim xApp As New CRAXDRT.Application
Dim xReport As CRAXDRT.Report
Dim CViewer As ClsReportViewer
Dim StrReportTitle As String
Dim StrFileName As String
Dim Msg As String

MySQL = "Select * From Expanses_Order  where ChqueNum='" & ChqueNum & "'"

 

If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.Path & "\Reports\Chque\" & report_no & ".rpt"
Else
    StrFileName = App.Path & "\Reports\Chque\" & report_no & ".rpt"
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
 'MsgBox ToHijriDate(Date)

 xReport.ParameterFields(5).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 1, 2)
 xReport.ParameterFields(6).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 4, 2)
 xReport.ParameterFields(7).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 9, 2)
 

  xReport.ParameterFields(8).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
 xReport.ParameterFields(9).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
 xReport.ParameterFields(10).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
 xReport.ParameterFields(11).AddCurrentValue CStr(txtto.text)
 
 
 
xReport.EnableParameterPrompting = False
xReport.ApplicationName = App.Title
xReport.ReportAuthor = App.Title
Set CViewer = New ClsReportViewer
CViewer.FireReport xReport, WindowTarget, ""

RsData.Close
Set RsData = Nothing
Screen.MousePointer = vbDefault

End Function




Function print_report(Optional NoteSerial As String)
    
Dim MySQL As String
Dim RsData As New ADODB.Recordset
Dim xApp As New CRAXDRT.Application
Dim xReport As CRAXDRT.Report
Dim CViewer As ClsReportViewer
Dim StrReportTitle As String
Dim StrFileName As String
Dim Msg As String

MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"

 
'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
'End If
'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
'End If

If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.Path & "\Reports\" & "Expenses_order2.rpt"
Else
    StrFileName = App.Path & "\Reports\" & "Expenses_order_Eng.rpt"
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
   ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    StrReportTitle = "" '& StrAccountName
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
    'End If
Else
 
    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    StrReportTitle = ""
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
    'End If
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

Private Sub CmdHelp_Click()
SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DcboBankName_Change()
On Error Resume Next
If DcboBankName.BoundText = "" Then Exit Sub
Dim RsSavRec As ADODB.Recordset
Dim My_SQL As String
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

Set RsSavRec = New ADODB.Recordset
RsSavRec.CursorLocation = adUseClient
RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 

 Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value



End If
End Sub

Private Sub DcboBox_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
End If
End Sub



Private Sub DcboBox_Click(Area As Integer)
DcboBox_Change
End Sub

Private Sub Fg_Journal_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 
Dim StrAccountCode As String
Dim Msg As String
Dim Rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
With Fg_Journal
    Select Case .ColKey(Col)
 
        Case "AccountName"
          '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
            StrAccountCode = .ComboData
            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
            If LngRow <> -1 Then
                'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '.TextMatrix(Row, Col) = ""
                '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                'Exit Sub
            End If
           ' Set ClsAcc = New ClsAccounts
            'If BolEditOnMainAccounts = False Then
               ' If LastAccount(StrAccountCode) = False Then
                '    .TextMatrix(Row, Col) = ""
                '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
               ' Else
'
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
            '      '  .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
              '  End If
           ' Else
                 .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                 '.TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
          '  End If
            'Set ClsAcc = Nothing
              Case "Value"
              
                
               
      Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
      Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    End Select
     Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
     Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    'to Add new row if needed
    If Row = .Rows - 1 Then
        .Rows = .Rows + 1
    End If
   ' ReLineGrid
End With
ReLineGrid
End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg_Journal
    If Row > .FixedRows Then
      '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
      '      Cancel = True
      '  End If
    End If
    Select Case .ColKey(Col)
        Case "value"
            .ComboList = ""
      Case "des"
        .ComboList = ""
          '  Cancel = True
            
    End Select
End With
End Sub

Private Sub Fg_Journal_DblClick()

Static lNoteRow&, lNoteCol&, r&, C&
With Fg_Journal
    ' clicking? no work
    'If Button <> 0 Then Exit Sub
    ' get mouse coordinates
    r = Fg_Journal.Row
    C = Fg_Journal.Col
    If Fg_Journal.ColKey(C) <> "Des" Then
        CboDes.Visible = False
        Exit Sub
    End If
    If Fg_Journal.TextMatrix(r, C) = "" Then
        'Exit Sub
    End If
    If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
        Exit Sub
    End If
    ' same cell or neighbour? no work
'    If r = lNoteRow And C = lNoteCol Then Exit Sub
'    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

    ' other cell, hide current note, if any
    If lNoteRow >= 0 And lNoteCol >= 0 Then
        Fg_Journal.SetFocus
        lNoteRow = -1
        lNoteCol = -1
    End If

    ' no note to show? then bail out
    If r <= 0 Or C <= 0 Then Exit Sub
    If TypeName(Fg_Journal.Cell(flexcpData, r, C)) <> "String" Then
        TxtDes.text = ""
    Else
        '
        TxtDes.text = Fg_Journal.Cell(flexcpData, r, C)
    End If
    ' show new note
    CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
    CboDes.Visible = True
    CboDes.ZOrder 0
    CboDes.SetFocus
    'save coordinates for next time
    lNoteRow = r
    lNoteCol = C
End With
End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Rs As New ADODB.Recordset
Dim StrSQL  As String
Dim StrAccountType As String
Dim StrComboList As String
Dim Msg As String
'Case "DebitName"
'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
'Case "CreditName"
'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
With Fg_Journal
    Select Case .ColKey(Col)
         Case "AccountName"
       Dim Account_Code_dynamic As String
        Account_Code_dynamic = get_account_code_branch(33, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
          Exit Sub
        Else
        If Account_Code_dynamic = "NO account" Then
           MsgBox "·„ Ì „ «‰‘«¡ Õ”«» «·„’—Ê›«  ·Â–« «·›—⁄", vbCritical
        Exit Sub
         
        End If
        End If
        
                     'Full Path Display
                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName " & _
                " FROM ACCOUNTS  Where ACCOUNTS.Account_Code" & _
                " LIKE '" & Account_Code_dynamic & "%' "
              '  If ChkLastAccount.value = vbChecked Then
              '      If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= 1) "
              '      Else
              '          StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
              '      End If
               ' End If
               ' If OptSort(1).value = True Then
               '     StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
               ' Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
               ' End If
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(Rs, "RootName,ParentName,*FirstName", "Account_Code")
               ' Debug.Print StrSQL
                
                
           ' ElseIf Opt(1).value = True Then
           '     'Full Path Display
           '     StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & _
           '     "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & _
           '     " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & _
           ''     "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & _
            '    "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & _
            '    "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
            '    If ChkLastAccount.value = vbChecked Then
            '        If SystemOptions.SysDataBaseType = AccessDataBase Then
            '            StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
            '        Else
            '            StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
            '        End If
            '    End If
               ' If OptSort(1).value = True Then
               '     StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
               ' Else
             '       StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                'End If
                
             '   Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             '   StrComboList = Fg_Journal.BuildComboList(Rs, "RootName,ParentName,*FirstName", "Account_Code")
             '   Debug.Print StrSQL
            'ElseIf Opt(2).value = True Then 'the normal Display
            '    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & _
            '    "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del " & _
            '    "From ACCOUNTS Where  ACCOUNTS.Account_Code <>'r' "
            '    If ChkLastAccount.value = vbChecked Then
            '        If SystemOptions.SysDataBaseType = AccessDataBase Then
            '            StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
            '        Else
            '            StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
            '        End If
            '    End If
            '    If OptSort(1).value = True Then
            '        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
            '    Else
            '        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
            '    End If
            '    Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            '    StrComboList = Fg_Journal.BuildComboList(Rs, "Account_Name", "Account_Code")
            'End If
            If StrComboList <> "" Then
                StrComboList = "|" & StrComboList
            End If
            .ComboList = StrComboList
    End Select
End With

End Sub

Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
Dim StrSQL As String

On Error GoTo ErrTrap

If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If
Set TTD = New clstooltipdemand
Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
Resize_Form Me
AddTip
SetDtpickerDate XPDtbTrans
Set Dcombos = New ClsDataCombos
Dcombos.GetBoxes Me.DcboBox
Dcombos.GetBanks Me.DcboBankName
Dcombos.GetUsers Me.DCboUserName
Dcombos.GetExpensesType XPCboExpensesType
Set cSearchDcbo = New clsDCboSearch
Set cSearchDcbo.Client = Me.XPCboExpensesType

Dcombos.GetAccountingCodes Me.DcboDebitSide
Dcombos.GetAccountingCodes Me.DcboCreditSide


With Me.CboPaymentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "‘Ìﬂ"
End With

StrSQL = " select expanses_account,Project_name from projects"
fill_combo dcproject, StrSQL

Set Rs = New ADODB.Recordset
StrSQL = "select * From Notes where NoteType=3 Order By NoteID"
Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
XPBtnMove_Click 2
Me.TxtModFlg.text = "R"
Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
hide_logo = False
If Rs.state = adStateOpen Then
    If Not (Rs.EOF Or Rs.BOF) Then
        If Rs.EditMode <> adEditNone Then
            Rs.CancelUpdate
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End If
Set TTP = Nothing
'Set EmpReport = Nothing
TTD.Destroy
Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, ByVal SpinningEnded As Boolean)
If ButtonID = vdsDownArrow Then
    If CboDes.IsDropped = False Then
        If PicHeight > 0 Then
            PicDes.Height = PicHeight
            PicDes.Width = PicWidth
        Else
            PicDes.Width = CboDes.Width - 10
            PicDes.Height = CboDes.Height * 8
        End If
        Debug.Print PicHeight
        Debug.Print PicWidth
        TxtDes.Visible = True
        TxtDes.text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
        CboDes.DropDown PicDes.hWnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
        Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
    Else
        CboDes.CloseUp
    End If
End If
End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{F4}"
End If
End Sub

Private Sub PicDes_Resize()
With PicDes
    LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
    TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
'    PicHeight = PicDes.Height
'    PicWidth = PicDes.Width
End With

End Sub

 

Private Sub TxtDes_LostFocus()
PicHeight = PicDes.Height
PicWidth = PicDes.Width
CboDes.CloseUp
CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    PutData
    CboDes.CloseUp
End If
End Sub

Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Select Case Me.TxtModFlg.text
    Case "R"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "«·„’—Ê›« "
        Else
            Me.Caption = "Expenses"
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
        
        XPTxtVal.locked = True
'        XPCboProfLevel.Locked = True
'        XPTxtProfMail.Locked = True
'        XPTxtPhone.Locked = True
'        XPTxtMobile.Locked = True
        XPMTxtRemarks.locked = True
        XPCboExpensesType.locked = True
        Me.DcboBox.locked = True
        XPDtbTrans.Enabled = False
        If Rs.RecordCount < 1 Then
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            
        End If
    Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "«·„’—Ê›« (ÃœÌœ)"
        Else
            Me.Caption = "Expenses(New Record)"
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
        
        XPTxtVal.locked = False
'        XPCboProfLevel.Locked = False
'        XPTxtProfMail.Locked = False
'        XPTxtPhone.Locked = False
'        XPTxtMobile.Locked = False
        XPMTxtRemarks.locked = False
        XPCboExpensesType.locked = False
        Me.DcboBox.locked = False
        XPDtbTrans.Enabled = True
        XPDtbTrans.value = Date
    Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "«·„’—Ê›« (  ⁄œÌ· )"
        Else
            Me.Caption = "Expenses(Edit Current Record)"
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
        
        XPTxtVal.locked = False
'        XPCboProfLevel.Locked = False
'        XPTxtProfMail.Locked = False
'        XPTxtPhone.Locked = False
'        XPTxtMobile.Locked = False
        XPMTxtRemarks.locked = False
        XPCboExpensesType.locked = False
        Me.DcboBox.locked = False
        XPDtbTrans.Enabled = True
End Select
Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MovePrevious
            If Rs.BOF Then Rs.MoveFirst
        End If
    Case 1
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveFirst
        End If
    Case 2
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveLast
        End If
    Case 3
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveNext
            If Rs.EOF Then Rs.MoveLast
        End If
End Select
Retrive
Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
Dim RsDev As ADODB.Recordset
Dim StrSQL As String
Dim I As Integer

On Error GoTo ErrTrap

If Rs.RecordCount < 1 Then
    XPTxtCurrent.Caption = 0
    XPTxtCount.Caption = 0
    Exit Sub
End If
If Rs.EOF Or Rs.BOF Then
    Exit Sub
Else
'Lngid
  '  If XPTxtID.text <> 0 Then
  '      Rs.find "NoteID=" & XPTxtID.text, , adSearchForward, adBookmarkFirst
  '      If Rs.EOF Or Rs.BOF Then
  '          Exit Sub
  '      End If
  '  End If
End If
XPTxtID.text = IIf(IsNull(Rs("NoteID").value), "", Val(Rs("NoteID").value))
Me.TxtNoteSerial.text = IIf(IsNull(Rs("NoteSerial").value), "", Rs("NoteSerial").value)
XPTxtVal.text = IIf(IsNull(Rs("Note_Value").value), "", Rs("Note_Value").value)
XPMTxtRemarks.text = IIf(IsNull(Rs("Remark").value), "", Rs("Remark").value)
txtto.text = IIf(IsNull(Rs("too").value), "", Rs("too").value)

XPDtbTrans.value = IIf(IsNull(Rs("NoteDate").value), Date, Rs("NoteDate").value)
XPCboExpensesType.BoundText = IIf(IsNull(Rs("ExpensesID").value), "", Rs("ExpensesID").value)


If IsNull(Rs("NoteCashingType").value) Then
    Me.CboPaymentType.ListIndex = 0
    Me.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), 0, Rs("BoxID").value)
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
ElseIf Rs("NoteCashingType").value = 0 Then
    Me.CboPaymentType.ListIndex = 0
    Me.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), 0, Rs("BoxID").value)
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
ElseIf Rs("NoteCashingType").value = 1 Then
    Me.CboPaymentType.ListIndex = 1
    Me.DcboBox.BoundText = ""
    Me.DcboBankName.BoundText = Rs("BankID").value
    Me.TxtChequeNumber.text = Rs("ChqueNum").value
    Me.DtpChequeDueDate.value = Rs("DueDate").value
End If
CboPayMentType_Change

'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

      If Rs("NoteCashingType").value = 0 Then
                DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
            ElseIf Rs("NoteCashingType").value = 1 Then
               DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", Val(Me.DcboBankName.BoundText))
       End If
            

Me.DCboUserName.BoundText = IIf(IsNull(Rs("UserID").value), "", Rs("UserID").value)
Me.Txt_Numorder.text = IIf(IsNull(Rs("NumOrderInpot").value), "", Rs("NumOrderInpot").value)
Me.TxtSerial.text = IIf(IsNull(Rs("NoteSerial").value), "", Rs("NoteSerial").value)
Me.dcproject.BoundText = IIf(IsNull(Rs("project_Expensen_account").value), "", Rs("project_Expensen_account").value)

'-----------------------------------------------------------------------------
If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
 '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
 '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"
StrSQL = "SELECT      dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & ") "
StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or Rs.EOF) Then
        Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
        Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
        RsDev.MoveFirst
        For I = 1 To RsDev.RecordCount
            If RsDev("Credit_Or_Debit").value = 0 Then
                Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
            ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
            End If
            RsDev.MoveNext
        Next I
    
    
      RsDev.MoveFirst
    
    With Me.Fg_Journal
    If Me.dcproject.BoundText = "" Then
   .Rows = .FixedRows + RsDev.RecordCount
    Else
    .Rows = .FixedRows + RsDev.RecordCount - 1
    End If
    For I = .FixedRows To .Rows - 1
        .TextMatrix(I, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), _
            "", RsDev("DEV_ID_Line_No").value)
            
        .TextMatrix(I, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), _
            "", RsDev("Account_Code").value)
            
                .TextMatrix(I, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), _
            "", RsDev("Account_Name").value)
            'Double_Entry_Vouchers_Description
        .TextMatrix(I, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), _
            "", RsDev("Double_Entry_Vouchers_Description").value)
  

            
    '    .TextMatrix(I, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), _
    '        "", RsDev("Account_Name").value)
        
                .TextMatrix(I, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), _
            "", RsDev("Value").value)
            
       
 
        RsDev.MoveNext
    Next I
        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
          Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
  '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
  '  .Rows - 1, .ColIndex("CreditValue"))
  '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
  '  .Rows - 1, .ColIndex("DebitValue"))
End With
End If

End If
'-----------------------------------------------------------------------------
XPTxtCurrent.Caption = Rs.AbsolutePosition
XPTxtCount.Caption = Rs.RecordCount
Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
Dim Msg As String
Dim RsTemp As New ADODB.Recordset
Dim StrSQL As String
Dim BeginTrans As Boolean
Dim RsDev As ADODB.Recordset
Dim LngDevID As Long

On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
'    If XPCboExpensesType.text = "" Then
'        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„’—Ê›«  "
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        XPCboExpensesType.SetFocus
'        SendKeys "{F4}"
'        Exit Sub
'    End If
'    If XPTxtVal.text = "" Then
'        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„’—Ê›«  "
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        XPTxtVal.SetFocus
'        Exit Sub
'    End If
'    If Not IsNumeric(XPTxtVal.text) Then
'        Msg = "ﬁÌ„… «·„’—Ê›«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        XPTxtVal.SetFocus
'        Exit Sub
'    End If
   ' If Trim(Me.DcboBox.BoundText) = "" Then
   '     Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
   '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
   '     DcboBox.SetFocus
   '     SendKeys "{F4}"
   '     Exit Sub
   ' End If
    
    
   If Me.CboPaymentType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPaymentType.SetFocus
        Exit Sub
    End If
    If Me.CboPaymentType.ListIndex = 0 Then
        If Trim(Me.DcboBox.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        If DateDiff("d", Me.DtpChequeDueDate.value, Date) >= 0 Then
            Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DtpChequeDueDate.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    End If
    
    If Me.TxtModFlg.text = "N" Then
        If Me.CboPaymentType.ListIndex = 0 Then
            If Val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), _
                    XPDtbTrans.value) = False Then
                    Exit Sub
                End If
            End If
        End If
    ElseIf Me.TxtModFlg.text = "E" Then
        If Me.CboPaymentType.ListIndex = 0 Then
            If Val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), _
                    XPDtbTrans.value, , , Val(Me.XPTxtID.text)) = False Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '-------------------------------------------------------------------------------------------
 
    '-------------------------------------------------------------------------------------------
    Cn.BeginTrans
    BeginTrans = True
    If TxtModFlg.text = "N" Then
        XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
        Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=3"))
        Rs.AddNew
        Rs("NoteID").value = Val(XPTxtID.text)
    ElseIf Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    
    Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    Rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, XPTxtVal.text)
    Rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
    Rs("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
    
  '  Rs("BankID").value = Null
    Rs("CusID").value = Null
    Rs("NoteType").value = 3
    Rs("NoteDate").value = XPDtbTrans.value
    Rs("UserID").value = user_id
    Rs("ExpensesID").value = IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
  
  
  If Me.CboPaymentType.ListIndex = 0 Then
        Rs("BoxID").value = Val(DcboBox.BoundText)
        Rs("BankID").value = Null
        Rs("ChqueNum").value = Null
        Rs("DueDate").value = Null
        Rs("NoteCashingType").value = 0
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        Rs("BoxID").value = Null
        Rs("BankID").value = Val(Me.DcboBankName.BoundText)
        Rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        Rs("DueDate").value = Me.DtpChequeDueDate.value
        Rs("NoteCashingType").value = 1
    End If
    
    
    
  '  Rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
    Rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
    Rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    Rs("Buy").value = "0"
    Rs("NoteSerial").value = Trim$(Me.TxtSerial.text)
    Rs("Remark").value = XPMTxtRemarks.text
    Rs("numbering_type").value = sand_numbering_type(1) 'numbering_type
    Rs("sanad_year").value = year(Date)
    Rs("sanad_month").value = Month(Date)
    Rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
    Rs.update
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
        Set RsDev = New ADODB.Recordset
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '«·ÿ—› «·„œÌ‰
 Dim I As Integer
Dim ExpensesID As Double

Dim line_no As Integer
line_no = 1
  With Fg_Journal
    For I = .FixedRows To .Rows - 2
        Dim IntDEV_Type As Integer
        Dim SngDEV_Value As Single
        If .TextMatrix(I, .ColIndex("AccountCode")) <> "" Then
  '.TextMatrix(i, .ColIndex("LineNo"))
'          ExpensesID = get_expanses_id(.TextMatrix(I, .ColIndex("AccountCode")))
          
            If ModAccounts.AddNewDev(LngDevID, line_no, _
                .TextMatrix(I, .ColIndex("AccountCode")), .TextMatrix(I, .ColIndex("value")), 0, _
                 .TextMatrix(I, .ColIndex("des")), Val(XPTxtID.text), , , _
                SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(I, .ColIndex("value")), , , , , setfoxy_Line) = False Then
                    GoTo ErrTrap
                    
            End If
            line_no = line_no + 1
        End If
    Next I
End With

     '   RsDev.AddNew
     '       RsDev("Double_Entry_Vouchers_ID").value = LngDevID
     '       RsDev("DEV_ID_Line_No").value = 1
     '       RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
     '       RsDev("Value").value = Val(Me.XPTxtVal.text)
     '       RsDev("Credit_Or_Debit").value = 0
     '       RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
     '       RsDev("Notes_ID").value = Val(XPTxtID.text)
     '       RsDev("RecordDate").value = Me.XPDtbTrans.value
     '       RsDev("UserID").value = Me.DCboUserName.BoundText
     '       RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
     '   RsDev.update
     
        '«·ÿ—› «·œ«∆‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
              RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = Me.XPTxtVal.text
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = Val(XPTxtID.text)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev.update
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 If Me.dcproject.BoundText <> "" Then
 line_no = line_no + 1
                ' «·ÿ—› «·„œÌ‰ 2
                RsDev.AddNew
                    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                    RsDev("DEV_ID_Line_No").value = line_no
                    RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                    RsDev("Account_Code").value = Me.dcproject.BoundText
                    RsDev("Value").value = Val(Me.XPTxtVal.text)
                    RsDev("Credit_Or_Debit").value = 0
                    RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                    RsDev("RecordDate").value = Me.XPDtbTrans.value
                    RsDev("Notes_ID").value = Val(XPTxtID.text)
                    RsDev("UserID").value = Me.DCboUserName.BoundText
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    
                RsDev.update
                '«ÿ—› «·œ«∆‰ 2'
                line_no = line_no + 1
                  With Fg_Journal
            For I = .FixedRows To .Rows - 2
          
                If .TextMatrix(I, .ColIndex("AccountCode")) <> "" Then
          '.TextMatrix(i, .ColIndex("LineNo"))
        '          ExpensesID = get_expanses_id(.TextMatrix(I, .ColIndex("AccountCode")))
                  
                    If ModAccounts.AddNewDev(LngDevID, line_no, _
                        .TextMatrix(I, .ColIndex("AccountCode")), .TextMatrix(I, .ColIndex("value")), 1, _
                         .TextMatrix(I, .ColIndex("des")), Val(XPTxtID.text), , , _
                        SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(I, .ColIndex("value"))) = False Then
                            GoTo ErrTrap
                          
                    End If
                      line_no = line_no + 1
                End If
            Next I
        End With
End If
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        LblDevID.Caption = LngDevID
        lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
    End If
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = Rs.AbsolutePosition
    XPTxtCount.Caption = Rs.RecordCount
    Select Case Me.TxtModFlg.text
        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
              Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            Cmd_Click (0)
            Exit Sub
            End If
        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Fg_Journal.Enabled = False
    End Select
    TxtModFlg.text = "R"
End If

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
Private Sub Undo()
On Error GoTo ErrTrap
Select Case TxtModFlg.text
    Case "N"
         clear_all Me
         Me.TxtModFlg.text = "R"
         XPBtnMove_Click (1)
    Case "E"
         Rs.find "NoteID='" & Val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst
         If Rs.EOF Or Rs.BOF Then
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
    Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
    Msg = Msg + (TxtNoteSerial.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
       If Not Rs.RecordCount < 1 Then
            Rs.Delete
            Rs.MoveFirst
            If Rs.RecordCount < 1 Then
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
    Rs.CancelUpdate
End Sub
Function FillGridWithData()


End Function
Private Sub ReLineGrid()
Dim I As Integer
Dim IntCounter As Integer
With Fg_Journal
    For I = .FixedRows To .Rows - 1
        If .TextMatrix(I, .ColIndex("AccountCode")) <> "" Then
            IntCounter = IntCounter + 1
            .TextMatrix(I, .ColIndex("LineNo")) = IntCounter
        End If
    Next I
End With
End Sub
Private Sub PutData()
'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
With Fg_Journal
    If Len(TxtDes.text) > 0 Then
        .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.text
        .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
        .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
    Else
        .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
        .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
        .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
    End If
End With
End Sub

Function sand_numbering() As String
On Error Resume Next
Dim start_at As Integer
Dim end_at As Integer
Dim auto_sanad_no As String
Dim no As String
auto_sanad_no = ""
departement_name = 1
branch_no = 1
connection_string = Cn.ConnectionString
numbering.ConnectionString = connection_string
numbering.CommandType = adCmdText
numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
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
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=3 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
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
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
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
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
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
                    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else
                     If numbering_type = 3 Then
                        auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

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
 Function setfoxy_Line() As Double
    
Dim X As Double
X = CStr(new_id("foxy", "id1", "", True))
setfoxy_Line = X
   Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable

 
 Rs("id1").value = X ' last_line_id
 
 Rs.update
 
    
End Function

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
        Cmd_Click (6)
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub AddTip()
Dim Wrap As String
Dim BolRtl As Boolean
If SystemOptions.UserInterface = ArabicInterface Then
    BolRtl = True
Else
    BolRtl = False
End If
On Error GoTo ErrTrap
Wrap = Chr(13) + Chr(10)
Set TTP = New clstooltip
If BolRtl = True Then
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(0), _
        "ÃœÌœ ..." & Wrap & _
        "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(1), _
        " ⁄œÌ· ..." & Wrap & _
        "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
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
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
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
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(4), _
        "Õ–› ..." & Wrap & _
        "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(6), _
        "Œ—ÊÃ ..." & Wrap & _
        "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(1), _
        "«·√Ê· ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(0), _
        "«·”«»ﬁ ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(3), _
        "«· «·Ì ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(2), _
        "«·√ŒÌ— ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl CmdHelp, _
        "„”«⁄œ… ..." & Wrap & _
        "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
        "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
        "≈÷€ÿ Â‰«" & Wrap, True
    End With
Else
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(0), _
        "Add New Record..." & Wrap & _
        "Shortcut Key F12 OR Enter" & Wrap & _
        "OR Alt+N", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(1), _
        "Edit the Current Record..." & Wrap & _
        "Shortcut Key F11 " & Wrap & _
        "OR Alt+E", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(2), _
        "Save the New Record OR Save the Editing in the Current Record..." & Wrap & _
        "Shortcut Key F10 " & Wrap & _
        "OR Alt+S", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(3), _
        "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & _
        "Shortcut Key F9 " & Wrap & _
        "OR Alt+U", BolRtl
    End With
     With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(4), _
       "Delete the Current Record..." & Wrap & _
        "Shortcut Key F8 " & Wrap & _
        "OR Alt+D", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(6), _
       "Close this Screen" & Wrap & _
        "OR Alt+X", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(1), _
        "«·√Ê· ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(0), _
        "«·”«»ﬁ ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(3), _
        "«· «·Ì ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(2), _
        "«·√ŒÌ— ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl CmdHelp, _
        "Help..." & Wrap & _
        "Display Help for this Screen" & Wrap & _
        "Shortcut Key F1" & Wrap, BolRtl
    End With
End If
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
        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
        StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    Case "E"
        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
        StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
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

Private Sub XPCboExpensesType_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", Val(Me.XPCboExpensesType.BoundText))
End If
End Sub

Private Sub XPTxtVal_Change()
'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0)
Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")
    

    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
'If Val(XPTxtVal.Text) = 0 Then
'    Set TTD = New clstooltipdemand
'    TTD.Style = TTBalloon
'    TTD.Icon = TTIconWarning
'    TTD.Centered = True
'    TTD.RightToLeft = True
'    TTD.VisibleTime = 600
'    TTD.BackColor = 0
'    TTD.Title = "ﬁÌ„… «·„’—Ê›« "
'    TTD.TipText = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·„’—Ê›« "
'    TTD.PopupOnDemand = True
'    TTD.CreateToolTip XPTxtVal.hwnd
'    TTD.Show 0, XPTxtVal.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
'    Cancel = True
'Else
'    TTD.Destroy
'End If
End Sub
Private Sub ViewDataList()
Dim FrmView As FrmViewList
Dim FG As VSFlex8UCtl.vsFlexGrid
Dim StrSQL As String
Dim Rs As ADODB.Recordset
Dim StrComboList As String
Dim GrdBack As ClsBackGroundPic
'Dim cProgress As ClsProgress
Dim BolFrmLoaded As Boolean
Set FrmView = New FrmViewList
Set FG = FrmView.vsfGroup1.vsFlexGrid

With FG
    .Cols = 18
    .RowHeightMin = 320
    .ExplorerBar = flexExSortShowAndMove
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
    
    StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & _
    "Remark, UserName From ExpensesReport"
    StrSQL = StrSQL + " Order By NoteID"
    Set Rs = New ADODB.Recordset
    Rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
    '------------------------------------
    '
    '
    '
    '
    
    '------------------------------------
    Set .DataSource = Rs
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
FrmView.vsfGroup1.Sql = StrSQL
FrmView.vsfGroup1.ShowTreeGroups = True
FrmView.vsfGroup1.update
FrmView.SetDblClickRetrun Me, "NoteID"
FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
FrmView.Show
End Sub
Private Sub ChangeLang()
Dim XPic As IPictureDisp
Set XPic = Me.XPBtnMove(1).ButtonImage
Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
Set Me.XPBtnMove(2).ButtonImage = XPic
Set XPic = Me.XPBtnMove(0).ButtonImage
Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
Set Me.XPBtnMove(3).ButtonImage = XPic
LblValue.Visible = False
lbl(14).Caption = "Project#"
Label1.Caption = "Manual #"
Me.ALLButton1.Caption = "Cost Center"
With Me.CboPaymentType
    .Clear
    .AddItem "Cash"
    .AddItem "Cheque"
End With

Me.Caption = "Expenses"
Me.Ele.Caption = "Expenses"
Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
Me.lbl(4).Caption = "Operation ID"
Me.lbl(1).Caption = "Operation Date"
Me.lbl(3).Caption = "Expenses Type"
Me.lbl(2).Caption = "Expenses Value"
Me.lbl(0).Caption = "Box Name"
Me.lbl(5).Caption = "Remarks"
Me.lbl(8).Caption = "Issued By."
Me.lbl(7).Caption = "Current Record."
Fra.Caption = "GL"
lbl(11).Caption = "GL#"
lbl(13).Caption = "interval"
lbl(9).Caption = "Depit"
lbl(10).Caption = "Credit"



Me.Cmd(0).Caption = "&New"
Me.Cmd(1).Caption = "&Edit"
Me.Cmd(2).Caption = "&Save"
Me.Cmd(3).Caption = "&Undo"
Me.Cmd(4).Caption = "&Delete"
Me.Cmd(5).Caption = "Sear&ch"
Me.Cmd(6).Caption = "E&xit"
Me.Cmd(7).Caption = "&Table View"
Me.CmdHelp.Caption = "&Help"

With Me.Fg_Journal
.TextMatrix(0, .ColIndex("LineNo")) = "Index"
.TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
.TextMatrix(0, .ColIndex("value")) = "value"
.TextMatrix(0, .ColIndex("des")) = "description"
End With



End Sub
