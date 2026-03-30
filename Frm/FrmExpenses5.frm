VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmExpenses5 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ ’—› -  Õ·Ì·Ì „’—Ê›« "
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16125
   HelpContextID   =   280
   Icon            =   "FrmExpenses5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   16125
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1ElastiMain 
      Height          =   9885
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16125
      _cx             =   28443
      _cy             =   17436
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   3735
         Left            =   0
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   720
         Width           =   16215
         _cx             =   28601
         _cy             =   6588
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
         Begin VB.TextBox TxtVATCustoms 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9420
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   136
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox XPTxtID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   1050
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtOrderID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3090
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   600
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox TxtNoteserial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   510
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Text            =   "Text1"
            Top             =   990
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   150
            Width           =   1785
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            Left            =   12120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   2880
            Width           =   2775
         End
         Begin VB.TextBox TxtManulaNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11220
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   120
            Width           =   915
         End
         Begin VB.CheckBox ChkCCDES 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«œ—«Ã «·‘—Õ  «·⁄«„ ›Ì ‘—Õ  „—ﬂ“ «· ﬂ·›…"
            Height          =   255
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   3360
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.TextBox TxtNoteserial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   600
            Width           =   2505
         End
         Begin VB.TextBox txt_general_des 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   99
            Top             =   2640
            Width           =   7275
         End
         Begin VB.TextBox txtto 
            Alignment       =   1  'Right Justify
            Height          =   525
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            Top             =   2070
            Width           =   7275
         End
         Begin VB.CheckBox chkDestribute 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ê“⁄"
            Enabled         =   0   'False
            Height          =   195
            Left            =   13740
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox XPMTxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   525
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   96
            Top             =   1440
            Width           =   7275
         End
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   11700
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13740
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txt_ORDER_NO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9420
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox TXT_A_NoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7860
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Text            =   "Text2"
            Top             =   3390
            Visible         =   0   'False
            Width           =   1095
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   1845
            Left            =   9180
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   960
            Width           =   6795
            _cx             =   11986
            _cy             =   3254
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
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   840
               Width           =   5475
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   120
               TabIndex        =   83
               Top             =   1140
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   556
               _Version        =   393216
               Format          =   110624769
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   120
               TabIndex        =   84
               Top             =   480
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   120
               TabIndex        =   85
               Top             =   120
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCAccounts 
               Height          =   315
               Left            =   120
               TabIndex        =   86
               Top             =   1440
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«»"
               Height          =   285
               Index           =   26
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1440
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   16
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   17
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   510
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   285
               Index           =   18
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
               Height          =   285
               Index           =   19
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   1140
               Width           =   1215
            End
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmExpenses5.frx":038A
            Height          =   315
            Left            =   3060
            TabIndex        =   103
            Top             =   120
            Width           =   4335
            _ExtentX        =   7646
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
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmExpenses5.frx":039F
            Height          =   315
            Left            =   4260
            TabIndex        =   104
            Top             =   1110
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo DCPreFix 
            Height          =   315
            Left            =   12900
            TabIndex        =   105
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   123
            Top             =   1110
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
         Begin MSDataListLib.DataCombo dcproject 
            Height          =   315
            Left            =   90
            TabIndex        =   126
            Top             =   1080
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9120
            TabIndex        =   129
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   110624769
            CurrentDate     =   38784
            MinDate         =   -292192
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   315
            Left            =   9120
            TabIndex        =   135
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «· VAT ··Ã„«—ﬂ"
            Height          =   255
            Index           =   28
            Left            =   12240
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Image ImgNote 
            Height          =   240
            Left            =   120
            Picture         =   "FrmExpenses5.frx":03B4
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   255
            Index           =   14
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   285
            Index           =   53
            Left            =   12060
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ÿ·» —ﬁ„"
            Height          =   255
            Left            =   7260
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   8100
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   195
            Index           =   22
            Left            =   14640
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2910
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«„—/«·„” ›Ìœ"
            Height          =   285
            Index           =   5
            Left            =   7500
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   285
            Index           =   4
            Left            =   14580
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   10620
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   135
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   195
            Index           =   15
            Left            =   14760
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   630
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   285
            Index           =   0
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   2310
            Width           =   1395
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
            Height          =   255
            Left            =   7500
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   1140
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—Õ «·⁄«„"
            Height          =   285
            Index           =   20
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   3150
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   405
            Index           =   21
            Left            =   9420
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   3000
            Width           =   1275
         End
      End
      Begin VB.TextBox XPTxtVal 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   8520
         Width           =   2265
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   11220
         Width           =   6465
         Begin MSDataListLib.DataCombo DcboDebitSide 
            Height          =   315
            Left            =   90
            TabIndex        =   7
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
            TabIndex        =   8
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
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—› „œÌ‰"
            Height          =   285
            Index           =   9
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   270
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—› œ«∆‰"
            Height          =   285
            Index           =   10
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ:"
            Height          =   315
            Index           =   11
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   270
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·› —… :"
            Height          =   315
            Index           =   13
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.Label LblDevID 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   3870
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   12
            Left            =   3870
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   570
            Width           =   1485
         End
      End
      Begin VB.TextBox Txt_Numorder 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   3735
         Left            =   -17760
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   16215
         Begin VB.Frame FraNote 
            BackColor       =   &H00E2E9E9&
            Height          =   1845
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   4470
            Visible         =   0   'False
            Width           =   6795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„’—Ê›« "
            Height          =   285
            Index           =   3
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   1080
            Width           =   1515
         End
      End
      Begin VB.TextBox XPTxtValView 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   8520
         Width           =   2265
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Index           =   0
         Left            =   0
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   16095
         _cx             =   28390
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
         BackColor       =   12648447
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Picture         =   "FrmExpenses5.frx":093E
         Caption         =   "”‰œ ’—› -  Õ·Ì·Ì „’—Ê›«  "
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
         Begin VB.TextBox oldTxtSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   18
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
            ButtonImage     =   "FrmExpenses5.frx":1618
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
            TabIndex        =   19
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
            ButtonImage     =   "FrmExpenses5.frx":19B2
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
            TabIndex        =   20
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
            ButtonImage     =   "FrmExpenses5.frx":1D4C
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
            TabIndex        =   21
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
            ButtonImage     =   "FrmExpenses5.frx":20E6
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
            Left            =   4680
            Top             =   600
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
            Left            =   2640
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
            TabIndex        =   22
            Top             =   510
            Visible         =   0   'False
            Width           =   5445
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   5040
            Picture         =   "FrmExpenses5.frx":2480
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
      End
      Begin MSDataListLib.DataCombo XPCboExpensesType 
         Height          =   315
         Left            =   16080
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10560
         TabIndex        =   24
         Top             =   9450
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   0
         Left            =   10860
         TabIndex        =   25
         Top             =   8880
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
         Left            =   9960
         TabIndex        =   26
         Top             =   8880
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
         Left            =   9150
         TabIndex        =   27
         Top             =   8880
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
         Left            =   7995
         TabIndex        =   28
         Top             =   8910
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
         Left            =   7080
         TabIndex        =   29
         Top             =   8910
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
         Left            =   3120
         TabIndex        =   30
         Top             =   8880
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
         Left            =   3960
         TabIndex        =   31
         Top             =   8910
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
         Left            =   6030
         TabIndex        =   32
         Top             =   8910
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
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   255
         Left            =   12000
         TabIndex        =   33
         Top             =   9000
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
         MICON           =   "FrmExpenses5.frx":60E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   8
         Left            =   5040
         TabIndex        =   34
         Top             =   8880
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   873
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
         Left            =   9120
         TabIndex        =   35
         Top             =   9360
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
      Begin ALLButtonS.ALLButton CmdRemove 
         Height          =   375
         Left            =   15120
         TabIndex        =   36
         Tag             =   "Delete Row"
         Top             =   8520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ–› ”ÿ—"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   0
         BCOLO           =   0
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmExpenses5.frx":6104
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         CausesValidation=   0   'False
         Height          =   375
         Index           =   10
         Left            =   7920
         TabIndex        =   37
         Top             =   9360
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   3855
         Left            =   0
         TabIndex        =   38
         Top             =   4440
         Width           =   16065
         _cx             =   28337
         _cy             =   6800
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
         Caption         =   "«·„’—Ê›« |‰”» «· Ê“Ì⁄|«··«∆ÕÂ «·œ«Œ·Ì…|Õ«·… «·«⁄ „«œ"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3435
            Left            =   16710
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   45
            Width           =   15975
            _cx             =   28178
            _cy             =   6059
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
            GridRows        =   10
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
            Begin VSFlex8Ctl.VSFlexGrid GridEstimatedCost 
               Height          =   3195
               Left            =   120
               TabIndex        =   40
               Top             =   120
               Width           =   15705
               _cx             =   27702
               _cy             =   5636
               Appearance      =   2
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   20
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmExpenses5.frx":6120
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3435
            Index           =   2
            Left            =   45
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   45
            Width           =   15975
            _cx             =   28178
            _cy             =   6059
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5955
               Index           =   1
               Left            =   0
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   -120
               Width           =   15975
               _cx             =   28178
               _cy             =   10504
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
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   -60
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   9585
                  Width           =   30
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
                  Height          =   3390
                  Left            =   60
                  TabIndex        =   43
                  Top             =   180
                  Width           =   15990
                  _cx             =   28205
                  _cy             =   5980
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
                  GridLines       =   3
                  GridLinesFixed  =   2
                  GridLineWidth   =   5
                  Rows            =   1
                  Cols            =   41
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmExpenses5.frx":6412
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
                  Begin VB.PictureBox PicDes 
                     BorderStyle     =   0  'None
                     Height          =   1635
                     Left            =   240
                     RightToLeft     =   -1  'True
                     ScaleHeight     =   1635
                     ScaleWidth      =   2925
                     TabIndex        =   44
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
                        TabIndex        =   45
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
                        TabIndex        =   46
                        Top             =   0
                        Width           =   2445
                     End
                  End
                  Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   47
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
                     Picture         =   "FrmExpenses5.frx":6A37
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
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                  Height          =   3630
                  Left            =   0
                  TabIndex        =   49
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   15990
                  _cx             =   28205
                  _cy             =   6403
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
                  GridLines       =   3
                  GridLinesFixed  =   2
                  GridLineWidth   =   5
                  Rows            =   2
                  Cols            =   34
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmExpenses5.frx":6FD1
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
                  Begin VB.Frame Frame3 
                     Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
                     Height          =   1215
                     Left            =   -120
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   3720
                     Visible         =   0   'False
                     Width           =   4215
                     Begin VB.CommandButton Command5 
                        Caption         =   "‰”Œ"
                        Height          =   255
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   63
                        Top             =   720
                        Width           =   1215
                     End
                     Begin VB.TextBox Text4 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   62
                        Top             =   240
                        Width           =   2175
                     End
                     Begin VB.Label Label7 
                        Alignment       =   1  'Right Justify
                        Caption         =   "—ﬁ„ «·ﬁÌœ"
                        Height          =   255
                        Left            =   2640
                        RightToLeft     =   -1  'True
                        TabIndex        =   64
                        Top             =   240
                        Width           =   1335
                     End
                  End
                  Begin VB.PictureBox Picture1 
                     BorderStyle     =   0  'None
                     Height          =   3915
                     Left            =   2550
                     RightToLeft     =   -1  'True
                     ScaleHeight     =   3915
                     ScaleWidth      =   9405
                     TabIndex        =   50
                     Top             =   810
                     Visible         =   0   'False
                     Width           =   9405
                     Begin VB.TextBox TxtDese 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000018&
                        BorderStyle     =   0  'None
                        Height          =   1485
                        Left            =   120
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   54
                        Top             =   2040
                        Width           =   8955
                     End
                     Begin VB.TextBox txtcodesub 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   5400
                        RightToLeft     =   -1  'True
                        TabIndex        =   53
                        Top             =   3600
                        Width           =   855
                     End
                     Begin VB.CommandButton Command4 
                        Caption         =   "Add des"
                        Height          =   255
                        Left            =   7440
                        RightToLeft     =   -1  'True
                        TabIndex        =   52
                        Top             =   3600
                        Width           =   1350
                     End
                     Begin VB.CommandButton Command3 
                        Caption         =   "Call des"
                        Height          =   255
                        Left            =   6240
                        RightToLeft     =   -1  'True
                        TabIndex        =   51
                        Top             =   3600
                        Width           =   1095
                     End
                     Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                        Height          =   3900
                        Left            =   120
                        TabIndex        =   55
                        TabStop         =   0   'False
                        Top             =   0
                        Width           =   10905
                        _cx             =   19235
                        _cy             =   6879
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial (Arabic)"
                           Size            =   20.25
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
                        BackColor       =   16777215
                        ForeColor       =   4210688
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
                        CaptionPos      =   7
                        WordWrap        =   -1  'True
                        MaxChildSize    =   0
                        MinChildSize    =   0
                        TagWidth        =   0
                        TagPosition     =   0
                        Style           =   0
                        TagSplit        =   2
                        PicturePos      =   7
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
                        Begin VB.TextBox Text3 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H80000018&
                           BorderStyle     =   0  'None
                           Height          =   1605
                           Left            =   0
                           MultiLine       =   -1  'True
                           RightToLeft     =   -1  'True
                           ScrollBars      =   3  'Both
                           TabIndex        =   56
                           Top             =   480
                           Visible         =   0   'False
                           Width           =   8955
                        End
                        Begin VB.Label Label2 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H8000000C&
                           Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                           ForeColor       =   &H0000C8FF&
                           Height          =   315
                           Left            =   6840
                           RightToLeft     =   -1  'True
                           TabIndex        =   57
                           Top             =   0
                           Width           =   2445
                        End
                     End
                     Begin VB.Label Label6 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Code"
                        Height          =   495
                        Left            =   1920
                        RightToLeft     =   -1  'True
                        TabIndex        =   60
                        Top             =   3480
                        Width           =   735
                     End
                     Begin VB.Label Label5 
                        Alignment       =   1  'Right Justify
                        Height          =   495
                        Left            =   1560
                        RightToLeft     =   -1  'True
                        TabIndex        =   59
                        Top             =   1200
                        Width           =   975
                     End
                     Begin VB.Label Label4 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Code"
                        Height          =   255
                        Left            =   1680
                        RightToLeft     =   -1  'True
                        TabIndex        =   58
                        Top             =   1320
                        Width           =   735
                     End
                  End
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   465
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   960
                  Width           =   15
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   23
               Left            =   8385
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   90
               Width           =   1140
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3435
            Left            =   17010
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   45
            Width           =   15975
            _cx             =   28178
            _cy             =   6059
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
            GridRows        =   10
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
            Begin VB.Frame Frame2 
               Height          =   3330
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   0
               Width           =   15615
               Begin VB.TextBox TxtScreenDesc 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2565
                  Left            =   120
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   69
                  Top             =   480
                  Width           =   15195
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   3435
            Left            =   17310
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   45
            Width           =   15975
            _cx             =   28178
            _cy             =   6059
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
               Height          =   3135
               Left            =   120
               TabIndex        =   131
               Tag             =   "1"
               Top             =   120
               Width           =   15855
               _cx             =   27966
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
               FormatString    =   $"FrmExpenses5.frx":74FA
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
               Left            =   6450
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   3240
               Width           =   3375
            End
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   11025
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   3480
               Width           =   3390
            End
         End
      End
      Begin ImpulseButton.ISButton CmdAttach 
         Height          =   375
         Left            =   6720
         TabIndex        =   70
         Top             =   9360
         Width           =   915
         _ExtentX        =   1614
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
      Begin ALLButtonS.ALLButton CmdRemoveAll 
         Height          =   375
         Left            =   14040
         TabIndex        =   71
         Tag             =   "Delete Row"
         Top             =   8520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ–› «·ﬂ·"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   0
         BCOLO           =   0
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmExpenses5.frx":763D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ImpulseButton.ISButton Accredit 
         Height          =   345
         Left            =   120
         TabIndex        =   134
         Top             =   9480
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   609
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«Ã„«·Ì"
         Height          =   285
         Index           =   2
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   8520
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   390
         Index           =   8
         Left            =   12225
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   9465
         Width           =   900
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
         Left            =   4380
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   9450
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "/"
         Height          =   435
         Index           =   6
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   9450
         Width           =   165
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   435
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   9450
         Width           =   525
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   435
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   9450
         Width           =   555
      End
      Begin VB.Label LblValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   405
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   8460
         Width           =   5895
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
         ForeColor       =   &H000000C0&
         Height          =   435
         Index           =   27
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   8400
         Width           =   5835
      End
   End
End
Attribute VB_Name = "FrmExpenses5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim Destribute As Boolean

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub TxtVATCustoms_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtVATCustoms.Text, 0)
End Sub
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·”‰œ " & TxtSerial1.Text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   «·›—⁄ " & Dcbranch & CHR(13) & "   „—ﬂ“ «· ﬂ·›… «·⁄«„  " & DcCostCenter & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄  " & CboPaymentType & CHR(13) & "   «·„‘—Ê⁄  " & dcproject & CHR(13) & "   «·Œ“Ì‰… " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "  »‰«¡ ⁄·Ï " & txtto & CHR(13) & "   »‰«¡ ⁄·Ï  " & CBoBasedON & "  »—ﬁ„  " & txt_ORDER_NO & CHR(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & CHR(13) & "   «Ã„«·Ì «·”‰œ    " & XPTxtValView
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. No " & TxtSerial1.Text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "   Branch " & Dcbranch & CHR(13) & "   CC  " & DcCostCenter & CHR(13) & "  Payment Type  " & CboPaymentType & CHR(13) & "   Project  " & dcproject & CHR(13) & "   Box " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No:   " & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Based On " & txtto & CHR(13) & "   Based On  " & CBoBasedON & "  No:  " & txt_ORDER_NO & CHR(13) & "  Remarks  " & txt_general_des & CHR(13) & "   Vchr Total   " & XPTxtValView
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 3, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtSerial, TxtSerial1
    Else
        AddToLogFile CInt(user_id), 3, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial, TxtSerial1
    End If
    
End Function

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
 
    SendTopost Me.Name, "notes_all", "NoteID", 0, val(Dcbranch.BoundText), val(XPTxtID.Text), TxtSerial1.Text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
    fillapprovData
End Sub

Private Sub ALLButton1_Click()
    On Error GoTo ErrTrap

    If DcCostCenter.BoundText <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Else
        MsgBox "It can not be the cost of distribution centers because you chose in distribution", vbCritical
        End If
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.Text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.Text)
    'Else
    'opr_id = TxtDEV_NO.text
    'End If

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 Then
            marakes_taklefa_tawze3.Show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… «Ê·« ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If

        marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
        marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
        marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
        marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
        marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        marakes_taklefa_tawze3.Adodc3.Refresh
        '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ALLButton2_Click()

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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
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
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.BackColor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.BackColor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function

Private Sub CBoBasedON_Change()
TxtVATCustoms.Visible = False
lbl(28).Visible = False
    With Me.Fg_Journal
        .ColHidden(.ColIndex("order_no")) = False

        If Me.CBoBasedON.ListIndex = 0 Then
 
            .ColHidden(.ColIndex("order_no")) = True

        ElseIf Me.CBoBasedON.ListIndex = 1 Then
TxtVATCustoms.Visible = True
lbl(28).Visible = True
            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·«„—"
            Else
                lbl(21).Caption = "  Order No"
            End If

        ElseIf Me.CBoBasedON.ListIndex = 2 Then
TxtVATCustoms.Visible = True
lbl(28).Visible = True
            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·› Ê—… «·„»œ∆ÌÂ"
            Else
                lbl(21).Caption = "Performa Invoice NO"
            End If

        ElseIf Me.CBoBasedON.ListIndex = 3 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·«„—"
            Else
                lbl(21).Caption = "  Order No"
            End If
        
        End If

        .TextMatrix(0, .ColIndex("order_no")) = lbl(21).Caption

    End With

End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboPayMentType_Change()
DcboBox.Enabled = False
DcboBankName.Enabled = False
    If Me.TxtModFlg.Text = "E" Then
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        Me.DcboBox.Text = ""

    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
        lbl(19).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(18).Caption = "Cheque No"
        lbl(19).Caption = "Due Date"
    End If

    If Me.CboPaymentType.ListIndex = 0 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        DCAccounts.Enabled = False
        DCAccounts.Text = ""
    ElseIf Me.CboPaymentType.ListIndex = 1 Or Me.CboPaymentType.ListIndex = 3 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        DCAccounts.Enabled = False
        DCAccounts.Text = ""
    
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame3.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·ÕÊ«·… "
            lbl(19).Caption = " «—ÌŒÂ«"
    
        Else
            lbl(18).Caption = "Transfer No"
            lbl(19).Caption = "Date"
        End If

        DCAccounts.Enabled = False
        DCAccounts.Text = ""
     
    ElseIf Me.CboPaymentType.ListIndex = 5 Then
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame3.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·«„— "
            lbl(19).Caption = " «—ÌŒÂ"
    
        Else
            lbl(18).Caption = "Bank O No"
            lbl(19).Caption = "Date"
        End If
    
        DCAccounts.Enabled = False
        DCAccounts.Text = ""
    ElseIf Me.CboPaymentType.ListIndex = 4 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        '     Me.DCVendor.Enabled = False
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        '    DCVendor.BoundText = ""
        DcboBox.BoundText = ""
        DcboBankName.BoundText = ""
        DCAccounts.Enabled = True
    
    Else
     '   Me.lbl(16).Enabled = True
     '   Me.DcboBox.Enabled = True
     '   Me.lbl(19).Enabled = False
     '   Me.lbl(18).Enabled = False
     '   Me.lbl(17).Enabled = False
     '   Me.DcboBankName.Enabled = False
     '   Me.TxtChequeNumber.Enabled = False
     '   Me.DtpChequeDueDate.Enabled = False
     '   DCAccounts.Enabled = False
     '   DCAccounts.Text = ""
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Function setfoxy()
    Text1.Text = CStr(new_id("foxy", "id", "", True))

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.Text
 
    rs.update
    
End Function
'Sub DeleteGridAccontVAT()
'Dim i As Integer
'With Fg_Journal
'i = .Rows
'Do
'i = i - 1
'If val(.TextMatrix(i, .ColIndex("FlgVat"))) = 1 Then
'.RemoveItem i
'End If
'Loop While i > 1
'End With
'End Sub

Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            DcCostCenter.Text = ""
            dcproject.Text = ""
Accredit.Caption = ""
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            GRID2.Clear flexClearScrollable, flexClearEverything
              GRID2.Rows = 1
           ChkCCDES.value = vbChecked
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Me.VSFlexGrid1.Rows = 2
          
            Fg_Journal.Enabled = True
            DtpChequeDueDate.value = Date
            setfoxy
            CBoBasedON.ListIndex = 0
            Me.Dcbranch.BoundText = Current_branch
            Txt_DateHigri.value = ToHijriDate(Date)


C1Tab1.CurrTab = 0
      XPDtbTrans.SetFocus
TxtChequeNumber.Text = 0
CboPayMentType_Change
        Case 1
                             If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            Dim Msg As String

            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    Else
                      Msg = " Can Not Edit this Process"
                      Msg = Msg & CHR(13) & " There is the Process of Payment checks "
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If
             If CheAssetPayd(val(Me.XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ≈÷«›… ··«’Ê·   "
                    Else
                    Msg = " Can Not Edit this Process"
                    Msg = Msg & CHR(13) & " There is the Process of adding Assest "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
                
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
        '    Me.DCboUserName.BoundText = user_id
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
         
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            VSFlexGrid1.Enabled = True
            CuurentLogdata
           ' DeleteGridAccontVAT
        Case 2
            If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            C1Tab1.CurrTab = 0
  
            If CBoBasedON.ListIndex > 0 And Trim(txt_ORDER_NO.Text) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify NO For"
                Else
                    Msg = "Õœœ —ﬁ„ "
                End If

                Msg = Msg & "  " & CBoBasedON.Text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                txt_ORDER_NO.SetFocus
                SendKeys "{F4}"
                 Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = val(Me.Dcbranch.BoundText)

            DcboBox_Change
            DcboBankName_Change
            DCAccounts_Change
               Dim Account_Code_dynamic82 As String
         If val(TxtVATCustoms.Text) > 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(148, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·Ã„«—ﬂ", vbCritical
                                                            Else
                                                                MsgBox "Please Select Customs Account", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
          End If
                         
            SaveData
           
        Case 3
            Undo

        Case 4
                     If ChekClodePeriod(XPDtbTrans.value) = True Then
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

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
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
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report (TxtSerial.Text)

        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber.Text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.Text

        Case 10
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc TxtSerial.Text, , 3, , , TxtSerial1.Text
    
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From notes  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
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
        Msg = "No Data"
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
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(txtto.Text)
  If right(XPTxtValView, 2) = "00" Then
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.Text)
    Else
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtValView.Text)
    End If
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.Text)
    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
    '  xReport.ParameterFields(15).AddCurrentValue Format$(DtpChequeDueDate.value, "dd/mm/yyyy")
 
    If SystemOptions.DateOpt = 0 Then
        xReport.ParameterFields(15).AddCurrentValue Format$(DtpChequeDueDate.value, "dd/mm/yyyy")
    Else
        xReport.ParameterFields(15).AddCurrentValue Format$(ToHijriDate(DtpChequeDueDate.value), "yyyy/mm/dd")
    End If
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Public Function print_report(Optional NoteSerial As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

  '  MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
 If NoteSerial = "" Then Exit Function
 MySQL = "SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
 MySQL = MySQL + "                     dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID,"
 MySQL = MySQL + "                     dbo.TblUsers.UserName, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.BanksData.BankNamee Namee, dbo.notes_all.too, dbo.Notes.Note_Value AS [Sub-value],"
 MySQL = MySQL + "                     dbo.Notes.note_value_by_characters AS sub_note_value_by_char, dbo.Notes.Remark AS sub_remark, dbo.ExpensesType.Name AS Sub_expenses_name,"
 MySQL = MySQL + "                     dbo.Notes.NoteType AS sub_notetype, dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.Notes.ExpensesRemark,"
 MySQL = MySQL + "                     dbo.ExpensesType.NameE As ExpensesNameE"
MySQL = MySQL + " ,Account_Serial="
MySQL = MySQL + "  ("
MySQL = MySQL + "  SELECT     dbo.ACCOUNTS.Account_Serial"
MySQL = MySQL + "  FROM         dbo.ExpensesType E INNER JOIN"
MySQL = MySQL + "                       dbo.ACCOUNTS ON dbo.ExpensesType.Account_Code = dbo.ACCOUNTS.Account_Code"
MySQL = MySQL + " Where (E.ID = dbo.ExpensesType.ID)"
MySQL = MySQL + " )"

 MySQL = MySQL + " FROM         dbo.ExpensesType RIGHT OUTER JOIN"
 MySQL = MySQL + "                     dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID LEFT OUTER JOIN"
 MySQL = MySQL + "                     dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
 MySQL = MySQL + "                     dbo.TblUsers ON dbo.Notes.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
 MySQL = MySQL + "                     dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID RIGHT OUTER JOIN"
 MySQL = MySQL + "                     dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
 
'MySQL = MySQL + "                     LEFT OUTER JOIN "
'MySQL = MySQL + "                     dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.Bankid "
 
MySQL = MySQL + "  WHERE     (dbo.Notes.NoteType = 3) "
MySQL = MySQL + "  and notes_all.noteserial='" & NoteSerial & "'"
MySQL = MySQL & "    AND (dbo.Notes.NoteSerial1 = " & TxtSerial1 & ")"

    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If
    If SystemOptions.DateOpt = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Expenses_orderH.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Expenses_orderH.rpt"
        End If

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
        Msg = "No Data"
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
    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1, "0612201402"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdRemove_Click()
        If Fg_Journal.Rows > 1 Then
If val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("FlgVat"))) = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ Õ–› ”ÿ— «·›«  .Ì—ÃÏ  ’›Ì— ‰”»… «·›« "
Else
MsgBox "Can not delete VAT  "
End If
Exit Sub
End If
End If

  '      If VSFlexGrid1.Rows > 1 Then
  '      If val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("FlgVat"))) = 1 Then
  '                      If SystemOptions.UserInterface = ArabicInterface Then
  '                      MsgBox "·«Ì„ﬂ‰ Õ–› ”ÿ— «·›«  .Ì—ÃÏ  ’›Ì— ‰”»… «·›« "
  '                      Else
  '                      MsgBox "Can not delete VAT  "
  '                      End If
  '          Exit Sub
  '          End If
'End If

Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub

  '  If dcproject.Text = "" Then
        Dim sql As String

        sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
        Cn.Execute sql, , adExecuteNoRecords
    
        If Fg_Journal.Rows > 1 Then
            If Fg_Journal.Rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.Rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                         
                        With Me.Fg_Journal

                         '   If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·„’—Ê›   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                            LogTexte = "  Delete  Expensen   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), val(TxtSerial1)
                        End With
 
                        Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                    End If
                End If
            End If
  '      End If
            
        With Fg_Journal
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With
            
    Else

        If VSFlexGrid1.Rows > 1 Then
            If VSFlexGrid1.Rows = 2 Then
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid1.Rows > 1 Then
                    If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                         
                        With Me.VSFlexGrid1

                            If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·Õ”«»   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                            LogTexte = "  Delete  Account   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
                        End With
 
                        Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid1
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    End If

End Sub

Private Sub CmdRemoveAll_Click()
   Dim X As Integer
Dim i As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub

  '  If dcproject.Text = "" Then
        Dim sql As String
For i = 1 To Fg_Journal.Rows - 1

        sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("LineNo1")))
        Cn.Execute sql, , adExecuteNoRecords
    
        If Fg_Journal.Rows > 1 Then
            If Fg_Journal.Rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.Rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                         
                        With Me.Fg_Journal

                         '   If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·„’—Ê›   " & .Cell(flexcpTextDisplay, i, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, i, .ColIndex("Value"))
                            LogTexte = "  Delete  Expensen   " & .Cell(flexcpTextDisplay, i, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, i, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), val(TxtSerial1)
                        End With
 
                    End If
                End If
           End If
           End If
            Next i
             Fg_Journal.Rows = 2
             Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything

    
End Sub

Private Sub DCAccounts_Change()
If val(CboPaymentType.ListIndex) <> 4 Then Exit Sub
    If DCAccounts.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = DCAccounts.BoundText
    End If

End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.Show
        Account_search.case_id = 201302
    End If

End Sub

Private Sub DCAccounts_Click(Area As Integer)
    DCAccounts_Change
End Sub

Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
        
        If CboPaymentType.ListIndex = 2 Or CboPaymentType.ListIndex = 3 Or CboPaymentType.ListIndex = 5 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If

End Sub

Private Sub DcboBox_Change()

    If DcboBox.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.Text = ""
    TxtSerial1.Text = ""
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
    
        CostCenterSearch.Show
        CostCenterSearch.RetrunType = 4
    End If

End Sub

Private Sub DCPreFix_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    TxtSerial.Text = ""
    TxtSerial1.Text = ""
    
End If
End Sub

Private Sub DCPreFix_Click(Area As Integer)
    TxtSerial.Text = ""
    TxtSerial1.Text = ""
End Sub

Private Sub dcproject_Change()
'If Me.TxtModFlg.Text <> "R" Then
'    Fg_Journal.Clear flexClearScrollable, flexClearEverything
'    Fg_Journal.Rows = 2
'
'    Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
'    Me.VSFlexGrid1.Rows = 2
'    End If
'    If dcproject.Text = "" Then
'        VSFlexGrid1.Visible = False
'        Me.Fg_Journal.Visible = True
'    Else
'
'        VSFlexGrid1.Visible = True
'        Me.Fg_Journal.Visible = False
'    End If
 
End Sub

Private Sub dcproject_Click(Area As Integer)
 Exit Sub
   ' If dcproject.Text = "" Then Exit Sub

   ' If SystemOptions.gldetails_or_gl_general = 0 Then 'Õ”«»«  «·„‘—Ê⁄
   '     VSFlexGrid1.Visible = True
   '     Me.Fg_Journal.Visible = False
   '     Fg_Journal.Clear flexClearScrollable, flexClearEverything
   '     Fg_Journal.Rows = 2
   '
   '     Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
   '     Me.VSFlexGrid1.Rows = 2
   ' Else
   '     VSFlexGrid1.Visible = False
   '     Me.Fg_Journal.Visible = True
   '     Fg_Journal.Clear flexClearScrollable, flexClearEverything
   '     Fg_Journal.Rows = 2
   '
   '     Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
   '     Me.VSFlexGrid1.Rows = 2
'
'    End If

End Sub

Function CheckAllExpensesDistributed() As Boolean
    CheckAllExpensesDistributed = False
    Dim i As Integer
    Dim zeroExist As Boolean
    Dim oneexist As Boolean

    With Fg_Journal

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Destribute")) = "0" Or .TextMatrix(i, .ColIndex("Destribute")) = "" Then
                zeroExist = True
            End If
        
            If .TextMatrix(i, .ColIndex("Destribute")) = "1" Then
                oneexist = True
            End If
        
            If zeroExist = True And oneexist = True Then
                CheckAllExpensesDistributed = False
                Exit Function
            End If
        
        Next i

    End With

    CheckAllExpensesDistributed = True
End Function

Function FillDestributionsToAll() As Boolean
    GridEstimatedCost.Clear flexClearScrollable, flexClearEverything
    GridEstimatedCost.Rows = 1
    Dim Msg As String

    If CheckAllExpensesDistributed = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " Â–« «·”‰œ ÌÕ ÊÏ ⁄·Ï „’«—Ì› „Ê“⁄Â Ê«Œ—Ï €Ì— „Ê“⁄Â Ê·« Ì„ﬂ‰ «·Õ›Ÿ  " & CHR(13)
                          
        Else
            Msg = " This Expenses Voucher  Have  Destribute and not  Destribute Expenses " & CHR(13)
            Msg = Msg + "can't Save"
                    
        End If
                                 
        FillDestributionsToAll = False
        Exit Function
            
    End If
 
    Dim i As Integer
    GridEstimatedCost.Clear flexClearScrollable, flexClearEverything
    GridEstimatedCost.Rows = 1
          
    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                FillDestributions .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("AccountName")), val(.TextMatrix(i, .ColIndex("value")))
        
            End If
        
        Next i

    End With
 
End Function
 
Public Function FillDestributions(AcountCode As String, _
                                  AcountName As String, _
                                  value As Double)
 
    Dim StrSQL  As String
    StrSQL = "SELECT     dbo.TblAccountsDestributions.AccountMaster, dbo.TblAccountsDestributionsDetails.ACode, dbo.TblAccountsDestributionsDetails.Percentage, "
    StrSQL = StrSQL + "  dbo.TblAccountsDestributions.DistType , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL + " FROM         dbo.TblAccountsDestributions INNER JOIN"
    StrSQL = StrSQL + " dbo.TblAccountsDestributionsDetails ON"
    StrSQL = StrSQL + " dbo.TblAccountsDestributions.TblAccountsDestributionsid = dbo.TblAccountsDestributionsDetails.TblAccountsDestributionsid INNER JOIN"
    StrSQL = StrSQL + "  dbo.TblBranchesData ON dbo.TblAccountsDestributionsDetails.ACode = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL + " WHERE     (dbo.TblAccountsDestributions.DistType IS NULL) AND (dbo.TblAccountsDestributions.AccountMaster = N'" & AcountCode & "')"
     
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
 
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
 
        row_count = GridEstimatedCost.Rows
    
        If GridEstimatedCost.TextMatrix(row_count - 1, GridEstimatedCost.ColIndex("AcountCode")) = "" Then
            row_count = row_count - 1
        End If
     
        GridEstimatedCost.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To GridEstimatedCost.Rows - 1 'RsDetails.RecordCount
    
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Ser")) = Num
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("AcountCode")) = AcountCode
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("AcountName")) = AcountName
           
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("BranchId")) = IIf(IsNull(RsDetails("Acode")), "", (RsDetails("Acode").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_name")), "", (RsDetails("branch_name").value))
            Else
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_namee")), "", (RsDetails("branch_namee").value))
            End If
         
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Percentage")) = IIf(IsNull(RsDetails("Percentage")), 0, (RsDetails("Percentage").value))
         
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("value")) = value
            
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Netvalue")) = Round(value * GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Percentage")) / 100, 2)
         
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If GridEstimatedCost.Rows > 10 Then
            '     If Num = 8 Then GridEstimatedCost.Refresh
            ' End If
        Next Num
 
    End If
            
End Function

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 4
             FrmProjectSearch.Show vbModal
           
        End If
End Sub
Sub DeleteGridCurrRow(Optional CurrRow As Long)
Dim i As Integer
With Fg_Journal
i = .Rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub AddVAT(Optional Row As Long)
If mdifrmmain.taxes = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts
With Fg_Journal
.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(.TextMatrix(Row, .ColIndex("branch_id"))), ForcedFlg)
.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
.TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)
End If
valuee = val(.TextMatrix(Row, .ColIndex("value")))
.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg

.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
GetValueAddedAccount XPDtbTrans.value, AccountVATDept
If AccountVATDept = "" And val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«» «·„œÌ‰ ›Ì ‘«‘… «⁄œ«œ  «·›« "
Else
MsgBox "Please Enter Account In VAT Settings"
End If
.TextMatrix(Row, .ColIndex("Vat")) = 0
.TextMatrix(Row, .ColIndex("Vatyo")) = 0
Exit Sub
End If
''/////////////

If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(Fg_Journal.Row, .ColIndex("AccountCode")) = "" Then
    DeleteGridCurrRow Row
   For i = 1 To 1
         .AddItem " ", Fg_Journal.Row + i
  k = .Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
.TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
.TextMatrix(k, .ColIndex("ExpensesID")) = 0
.TextMatrix(k, .ColIndex("branch_id")) = .TextMatrix(Row, .ColIndex("branch_id"))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " ﬁÌ„… „÷«›…"
Else
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " VAT  "
End If
.TextMatrix(k, .ColIndex("Destribute")) = .TextMatrix(Row, .ColIndex("Destribute"))
.TextMatrix(k, .ColIndex("branch_name")) = .TextMatrix(Row, .ColIndex("branch_name"))
.TextMatrix(k, .ColIndex("opr_fullcode")) = .TextMatrix(Row, .ColIndex("opr_fullcode"))
.TextMatrix(k, .ColIndex("CarName")) = .TextMatrix(Row, .ColIndex("CarName"))
.TextMatrix(k, .ColIndex("CarId")) = .TextMatrix(Row, .ColIndex("CarId"))
'.TextMatrix(k, .ColIndex("projectid2")) = .TextMatrix(Row, .ColIndex("projectid2"))
.TextMatrix(k, .ColIndex("Fixes")) = .TextMatrix(Row, .ColIndex("Fixes"))
'.TextMatrix(k, .ColIndex("PrjectCode")) = .TextMatrix(Row, .ColIndex("PrjectCode"))
'.TextMatrix(k, .ColIndex("project")) = .TextMatrix(Row, .ColIndex("project"))
'.TextMatrix(k, .ColIndex("pand")) = .TextMatrix(Row, .ColIndex("pand"))
'.TextMatrix(k, .ColIndex("oper")) = .TextMatrix(Row, .ColIndex("oper"))
'.TextMatrix(k, .ColIndex("operid")) = .TextMatrix(Row, .ColIndex("operid"))
'.TextMatrix(k, .ColIndex("pandid")) = .TextMatrix(Row, .ColIndex("pandid"))
'.TextMatrix(k, .ColIndex("projectid")) = .TextMatrix(Row, .ColIndex("projectid"))
.TextMatrix(k, .ColIndex("pandid2")) = .TextMatrix(Row, .ColIndex("pandid2"))
.TextMatrix(k, .ColIndex("pand")) = .TextMatrix(Row, .ColIndex("pand"))
.TextMatrix(k, .ColIndex("oper")) = .TextMatrix(Row, .ColIndex("oper"))
.TextMatrix(k, .ColIndex("operid2")) = .TextMatrix(Row, .ColIndex("operid2"))
.TextMatrix(k, .ColIndex("fixedid")) = .TextMatrix(Row, .ColIndex("fixedid"))
.TextMatrix(k, .ColIndex("fixCode")) = .TextMatrix(Row, .ColIndex("FixCode"))
.TextMatrix(k, .ColIndex("deptid")) = .TextMatrix(Row, .ColIndex("deptid"))
.TextMatrix(k, .ColIndex("dept")) = .TextMatrix(Row, .ColIndex("dept"))
.TextMatrix(k, .ColIndex("FlgVat")) = 1
    If SystemOptions.IsMergeVat Then
        .RowHidden(k) = True
    End If
Next i
End If
End If
End With
End If
End Sub
Sub HidFat()
With Fg_Journal
If mdifrmmain.taxes = True Then
.ColHidden(.ColIndex("Vat")) = False
.ColHidden(.ColIndex("Vatyo")) = False
Else
.ColHidden(.ColIndex("Vat")) = True
.ColHidden(.ColIndex("Vatyo")) = True
End If
End With
End Sub
Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sgl As String
    Dim str  As String
    Dim rsDummy As New ADODB.Recordset
    With Fg_Journal
        sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
        Cn.Execute sgl, , adExecuteNoRecords
        If .TextMatrix(Row, .ColIndex("project")) = "" Then
        .TextMatrix(Row, .ColIndex("project")) = (dcproject.Text)
        .TextMatrix(Row, .ColIndex("projectid2")) = val(dcproject.BoundText)
        End If



        Select Case .ColKey(Col)
        Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .Rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
        
        Case "PriceTotal"
              If val(.TextMatrix(Row, .ColIndex("branch_id"))) = 0 Then
                .TextMatrix(Row, .ColIndex("branch_id")) = val(Dcbranch.BoundText)
                .TextMatrix(Row, .ColIndex("branch_name")) = Dcbranch.Text
                
        End If
        AddVAT Row
        Case "branch_name"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("branch_id")) = StrAccountCode
                AddVAT Row
       Case "Supplier"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SupplierID"), False, True)
                .TextMatrix(Row, .ColIndex("SupplierID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("SupplierName")) = .TextMatrix(Row, .ColIndex("Supplier"))
                StrSQL = "Select * From TblCustemers Where CusID=" & val(.TextMatrix(Row, .ColIndex("SupplierID")))
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CusVATNO")) = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
                Else
                .TextMatrix(Row, .ColIndex("CusVATNO")) = ""
                End If
                
             AddVAT Row
        Case "dept"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("deptid")) = StrAccountCode
             AddVAT Row
        Case "FixCode"
               
                str = " SELECT   TblCarsData.ID,   TblCarsData.Fullcode, fixedassetid ,TblCarsData.EqupName,TblCarsData.BoardNO                FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  dbo.TblCarsData.Fullcode like '%" & Trim(.TextMatrix(Row, .ColIndex("FixCode"))) & "%'  "
                rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("fixedid")) = val(rsDummy!FixedassetId & "")
                    .TextMatrix(Row, .ColIndex("Fixes")) = Trim(rsDummy!EqupName & "")
                    .TextMatrix(Row, .ColIndex("CarId")) = val(rsDummy!ID & "")
                    .TextMatrix(Row, .ColIndex("CarName")) = Trim(rsDummy!BoardNO & "")
  If CheckEqp(val(.TextMatrix(Row, .ColIndex("fixedid")))) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
               Msg = " ÊÃœ „‘«—Ì⁄ ·Â–Â «·„⁄œ… Â·  —Ìœ  Õ„Ì·Â«  ·ﬁ«∆Ì«"
               Else
               Msg = "There are projects. Do you want downloaded automatically"
               End If
               If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
               FillGridEqup val(.TextMatrix(Row, .ColIndex("fixedid"))), Row
               End If
               End If
                 AddVAT Row
                   ' Fg_Journal_AfterEdit Row, Fg_Journal.ColIndex("Fixes")
                Else
                  .TextMatrix(Row, .ColIndex("fixedid")) = ""
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("Fixes")) = ""
                    .TextMatrix(Row, .ColIndex("CarName")) = ""
                End If
        Case "Fixes"
                StrAccountCode = .ComboData
              .TextMatrix(Row, .ColIndex("fixedid")) = StrAccountCode
                               If CheckEqp(val(.TextMatrix(Row, .ColIndex("fixedid")))) = True Then
                               
                 
                               
               If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ÊÃœ „‘«—Ì⁄ ·Â–Â «·„⁄œ… Â·  —Ìœ  Õ„Ì·Â«  ·ﬁ«∆Ì«"
               Else
                    Msg = "There are projects. Do you want downloaded automatically"
               End If
               If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
               FillGridEqup val(.TextMatrix(Row, .ColIndex("fixedid"))), Row
               End If
               End If
                 
                    str = " SELECT       TblCarsData.Fullcode,fixedassetid ,TblCarsData.EqupName,TblCarsData.BoardNO                FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  dbo.TblCarsData.fixedassetid = " & val(StrAccountCode)
                    rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummy.EOF Then
                       ' .TextMatrix(Row, .ColIndex("fixedid")) = val(rsDummy!FixedassetId & "")
                        '.TextMatrix(Row, .ColIndex("Fixes")) = Trim(rsDummy!EqupName & "")
                        .TextMatrix(Row, .ColIndex("CarName")) = Trim(rsDummy!BoardNO & "")
                        .TextMatrix(Row, .ColIndex("FixCode")) = Trim(rsDummy!Fullcode & "")
                    End If
                 AddVAT Row
         Case "project"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("projectid2")) = StrAccountCode
                        If StrAccountCode <> "" Then
                StrSQL = " SELECT Fullcode  From dbo.Projects where id =" & val(StrAccountCode) & ""
                End If
                     Set rs = New ADODB.Recordset
                     If StrSQL = "" Then Exit Sub
      rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      .TextMatrix(Row, .ColIndex("PrjectCode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
       Case "PrjectCode"
       If .TextMatrix(Row, .ColIndex("PrjectCode")) <> "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name , id From dbo.Projects where not(Project_name is null) and Project_name <>N'""' "
           Else
               StrSQL = " SELECT  LTRIM(RTRIM( Project_nameE )) as Project_nameE , id From dbo.Projects where not(Project_nameE is null) and Project_nameE <>N'""' "
       End If
       StrSQL = StrSQL & " and Fullcode= N'" & .TextMatrix(Row, .ColIndex("PrjectCode")) & "'"
       Set rs = New ADODB.Recordset
      rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      If rs.RecordCount > 0 Then
       .TextMatrix(Row, .ColIndex("projectid2")) = IIf(IsNull(rs("id").value), 0, rs("id").value)
       If SystemOptions.UserInterface = ArabicInterface Then
       .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
       Else
       .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(rs("Project_nameE").value), "", rs("Project_nameE").value)
       End If
       End If
       End If
        AddVAT Row
                  Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid2")) = StrAccountCode
                  Case "oper"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid2")) = StrAccountCode
  AddVAT Row
            Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
    AddVAT Row
            Case "CarName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
     
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CarID"), False, True)
                .TextMatrix(Row, .ColIndex("CarID")) = StrAccountCode
            
                .TextMatrix(Row, .ColIndex("des")) = "’—›  ⁄·Ï «·„⁄œÂ/«·”Ì«—…  : " & .TextMatrix(Row, .ColIndex("CarName"))
                str = " SELECT      TblCarsData.Fullcode, fixedassetid ,TblCarsData.EqupName,TblCarsData.BoardNO                FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  dbo.TblCarsData.fixedassetid = " & val(StrAccountCode)
                rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("fixedid")) = val(rsDummy!FixedassetId & "")
                    .TextMatrix(Row, .ColIndex("Fixes")) = Trim(rsDummy!EqupName & "")
                    .TextMatrix(Row, .ColIndex("FixCode")) = Trim(rsDummy!Fullcode & "")
                   ' .TextMatrix(Row, .ColIndex("CarName")) = Trim(rsDummy!BoardNO & "")
                End If
             AddVAT Row
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
     
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                    
                .TextMatrix(Row, .ColIndex("Destribute")) = 0
                StrAccountCode = .TextMatrix(Row, .ColIndex("AccountCode"))

                If CheckAccountHaveDestributions(StrAccountCode) = True Then
             
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " Â–« «·„’—Ê› ·Â ŒÿÂ  Ê“Ì⁄  ⁄·Ï «·›—Ê⁄ Â·  —Ìœ «· Ê“Ì⁄  " & CHR(13)
                        Msg = Msg + "‰⁄„ «„ ·« "
                          
                    Else
                        Msg = " This Expenses Have Destribution Plan Do you want  Destribute  " & CHR(13)
                        Msg = Msg + "Yes Or No"
                    
                    End If
                                 
                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                        .TextMatrix(Row, .ColIndex("Destribute")) = 1
         
                    Else
                        .TextMatrix(Row, .ColIndex("Destribute")) = 0
                    End If
            
                End If
 
                FillDestributionsToAll
             
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = txt_ORDER_NO.Text
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                     .TextMatrix(Row, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If
                AddVAT Row
Case "Account_Serial"
If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Serial='" & .TextMatrix(Row, .ColIndex("Account_Serial")) & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Serial='" & .TextMatrix(Row, .ColIndex("Account_Serial")) & "'"
                End If
                  rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    End If
                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                  .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                    .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(.TextMatrix(Row, .ColIndex("AccountCode")))
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = txt_ORDER_NO.Text
                      If CheckAccountHaveDestributions(.TextMatrix(Row, .ColIndex("AccountCode"))) = True Then
             
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " Â–« «·„’—Ê› ·Â ŒÿÂ  Ê“Ì⁄  ⁄·Ï «·›—Ê⁄ Â·  —Ìœ «· Ê“Ì⁄  " & CHR(13)
                        Msg = Msg + "‰⁄„ «„ ·« "
                          
                    Else
                        Msg = " This Expenses Have Destribution Plan Do you want  Destribute  " & CHR(13)
                        Msg = Msg + "Yes Or No"
                    
                    End If
                                 
                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                        .TextMatrix(Row, .ColIndex("Destribute")) = 1
         
                    Else
                        .TextMatrix(Row, .ColIndex("Destribute")) = 0
                    End If
            
                End If
            
                End If
                AddVAT Row
            Case "value", "opr_fullcode"
        
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
    
                FillDestributionsToAll
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        
        If val(.TextMatrix(Row, .ColIndex("branch_id"))) = 0 Then
                .TextMatrix(Row, .ColIndex("branch_id")) = val(Dcbranch.BoundText)
                .TextMatrix(Row, .ColIndex("branch_name")) = Dcbranch.Text
                
        End If
        AddVAT Row
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

    With Me.Fg_Journal

        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·„’—Ê› «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Value") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··„’—Ê›   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change value" & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Expenses " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··„’—Ê›   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Expenses " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        End If

        AddToLogFile CInt(user_id), 3, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), TxtSerial1
    End With

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If
If val(.TextMatrix(Row, .ColIndex("FlgVat"))) <> 0 Then
Cancel = True
Else
 Select Case .ColKey(Col)
        Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
           Case "BillNo", "FixCode"
                .ComboList = ""
         Case "LineNo"
                .ComboList = ""
         Case "CusVATNO"
                .ComboList = ""
         Case "SupplierName"
                .ComboList = ""
         Case "PriceTotal"
                .ComboList = ""
                
          Case "PrjectCode"
                .ComboList = ""
     
     Case "value"
                .ComboList = ""
            Case "Account_Serial"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
            Case "TradingContractID"
                .ComboList = ""
                '  Cancel = True
            
        End Select
End If
    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
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
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then Exit Sub
   SendKeys "{F4}"
     SendKeys "{BACKSPACE}"
   SendKeys CHR(KeyAscii)
 
End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
             
                   Order_no_search.Show
                     Order_no_search.RetrunType = 4
                   
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    
                    FrmExpensesSearch.Show
                    FrmExpensesSearch.RetrunType = 1
                   
                End If
 
  Case "PrjectCode", "project"
    If KeyCode = vbKeyF3 Then
        Unload FrmProjectSearch
        FrmProjectSearch.Show
        FrmProjectSearch.lblSearchtype.Caption = 23
    End If
            Case "Account_Serial"

                If KeyCode = vbKeyF3 Then
                    
                    FrmExpensesSearch.Show
                    FrmExpensesSearch.RetrunType = 1
                    
                End If
 
        End Select

    End With

End Sub
Public Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                ByVal Col As Long, _
                                Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim StrComboList2 As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
           Case "Supplier"
            StrSQL = " SELECT     CusID, CusName, CusNamee"
            StrSQL = StrSQL & "            From dbo.TblCustemers"
            StrSQL = StrSQL & "       WHERE     (Type = 2)"
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "CusNamee", "CusID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
        Case "branch_name"
         If SystemOptions.UserInterface = ArabicInterface Then
       StrSQL = " SELECT     branch_id, branch_name From TblBranchesData"
      
       Else
        StrSQL = " SELECT     branch_id , branch_namee From TblBranchesData "
     
        End If
       StrSQL = StrSQL & " where  branch_id in(" & Current_branchSql & ")"
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_name", "branch_id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_namee", "branch_id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
        Case "dept"
         If SystemOptions.UserInterface = ArabicInterface Then
       StrSQL = " SELECT     DeparmentID, DepartmentName From dbo.TblEmpDepartments "
      
       Else
        StrSQL = " SELECT     DeparmentID , DepartmentNamee From dbo.TblEmpDepartments"
     
        End If
        
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
''///
        Case "Fixes"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE   ISEQUP=1 or id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE  ISEQUP=1 or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
                    Case "project"
If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name , id From dbo.Projects  "
                StrSQL = StrSQL & " where Project_name<>N'""' and not (Project_name is null)"
   Else
    StrSQL = " SELECT  LTRIM(RTRIM( Project_nameE )) as Project_nameE , id From dbo.Projects  "
                StrSQL = StrSQL & " where Project_nameE<>N'""' and not (Project_nameE is null)"
End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
           If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Project_name", "id")
              Else
                     StrComboList = Fg_Journal.BuildComboList(rs, "Project_nameE", "id")
              End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid2")) = "" Then
             If SystemOptions.UserInterface = ArabicInterface Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Else
             MsgBox "Please Select Project"
             End If
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ")"
           
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = Fg_Journal.BuildComboList(rs, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  Case "oper"
                   
If .TextMatrix(Row, .ColIndex("projectid2")) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
Else
MsgBox "Please Select Project"
End If
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("pandid2")) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
Else
MsgBox "Please Select Des"
End If
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("pandid2"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ")"
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
''//
            Case "AccountName"
                '      StrSQL = "select * from Expenses_accounts"
                             
                If SystemOptions.UserInterface = ArabicInterface Then

                    StrSQL = "select * from Expenses_accounts order by Account_Name"
                Else
                    StrSQL = "select * from Expenses_accounts_eng order by Account_Nameeng"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Nameeng", "Account_Code")
                End If
           
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
            
                '«ŸÂ«— «·„⁄œ« /«·”Ì«—« 
            Case "CarName"
        
                StrSQL = "  select id,BoardNO from TblCarsData"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList2 = Fg_Journal.BuildComboList(rs, "BoardNO", "id")
       
                If StrComboList2 <> "" Then
                    StrComboList2 = "|" & StrComboList2
                End If

                .ComboList = StrComboList2
         
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String

    On Error GoTo ErrTrap
    TxtScreenDesc.Text = GetScreenDescription(Me.Name)
   
'    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
'    fill_combo Me.DcCostCenter, StrSQL

    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If
    'mdifrmmain.taxes = False
    HidFat
    ScreenNameArabic = "”‰œ ’—› -  Õ·Ì·Ì „’—Ê›« "
    ScreenNameEnglish = "Payments Voucher "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 3

    If SystemOptions.DateOpt = 1 Then
        Label8.Visible = False
        DcCostCenter.Visible = False
        lbl(14).Visible = False
        dcproject.Visible = False

    End If

    With Fg_Journal

        If mdifrmmain.MnuProjects.Visible = False Then
            .ColHidden(.ColIndex("opr_fullcode")) = True
       '     .ColHidden(.ColIndex("project")) = True
            .ColHidden(.ColIndex("pand")) = True
            .ColHidden(.ColIndex("oper")) = True
        '    .ColHidden(.ColIndex("PrjectCode")) = True
            
       End If
 
        If mdifrmmain.TransporterMain.Visible = False Then
            .ColHidden(.ColIndex("CarName")) = True
        End If

    End With

    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    Resize_Form Me
    AddTip
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Dcombos.GetAccountingCodes Me.DCAccounts, True
'    Dim Dcombos As ClsDataCombos
'Set Dcombos = New ClsDataCombos
Dcombos.GetCostCenter DcCostCenter
   Dcombos.GetPrefix2 Me.DCPreFix, 1, Current_branch
   

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Dcbranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ/ ⁄ÂœÂ"
        .AddItem "‘Ìﬂ"
        .AddItem " ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "‘Ìﬂ „”œœ"
        .AddItem "Õ”«»"
        .AddItem "√„— »‰ﬂÌ"
    End With

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
        .AddItem " «„— «‰ «Ã  "
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = " select expanses_account,Project_name from projects  where not(expanses_account is null) and (Project_name<>N'""') and not (Project_name is null)"
    StrSQL = StrSQL & "  order by Project_name"
  Else
    StrSQL = " select expanses_account,Project_nameE from projects  where not(expanses_account is null) and Project_nameE<>N'""' and not (Project_nameE is null)"
    StrSQL = StrSQL & "  order by Project_nameE"
End If
    fill_combo dcproject, StrSQL

    Set rs = New ADODB.Recordset
    StrSQL = "select * From notes_all where notetype=3"
    
            If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
       
    StrSQL = "select * From notes_all where  (ToPriodDateH is null) and notetype=3 AND branch_no in(" & Current_branchSql & ")"
     If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    'MsgBox ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    hide_logo = False
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 3

    If rs.State = adStateOpen Then
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
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

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
            TxtDes.Text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hwnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{F4}"
    End If

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If CBoBasedON.ListIndex = 3 Then
        If KeyCode = vbKeyF3 Then
             Order_no_search2.Show
             Order_no_search2.RetrunType = 3
         
        End If

    Else

        If KeyCode = vbKeyF3 Then
             Order_no_search.Show
             Order_no_search.RetrunType = 0
        End If

    End If

End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„’—Ê›« "
            Else
                Me.Caption = "Expenses"
            End If
        
            Me.VSFlexGrid1.Enabled = False
            Me.Fg_Journal.Enabled = False
            Frame1.Enabled = False
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            CmdRemove.Enabled = False
            CmdRemoveAll.Enabled = False
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

            If rs.RecordCount < 1 Then
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
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            CmdRemove.Enabled = True
            CmdRemoveAll.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            'Me.XPBtnMove(0).Enabled = False
            'Me.XPBtnMove(1).Enabled = False
            'Me.XPBtnMove(2).Enabled = False
            'Me.XPBtnMove(3).Enabled = False
        
            ' XPTxtVal.locked = False
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
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
        
            CmdRemove.Enabled = True
            CmdRemoveAll.Enabled = True
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

Private Sub TxtNoteSerial1_Change()
If TxtNoteserial1.Text <> "" Then
Dim Type1 As Integer
Dim TXTperson As String
Dim des As String
Dim EmpID As Integer
Dim Price As Double
If Me.TxtModFlg.Text <> "R" Then
OrderExchange TxtNoteserial1.Text, Type1, TXTperson, des, Price, EmpID
CboPaymentType.ListIndex = Type1
txtto.Text = TXTperson
txt_general_des.Text = des
End If
End If
End Sub

Private Sub TxtNoteserial1_KeyUp(KeyCode As Integer, Shift As Integer)
    
        If KeyCode = vbKeyF3 Then
            FrmReqExchangeSearch.Show
            FrmReqExchangeSearch.lbltype.Caption = 2
        End If
End Sub
Function CheckEqp(Optional FixedID As Double) As Boolean
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     FixedID"
sql = sql & " From dbo.TblSpecificFixedDeti"
sql = sql & " WHERE     (FixedID = " & FixedID & ") AND (ToDate >=" & SQLDate(XPDtbTrans.value, True) & ") AND (FromDate <=" & SQLDate(XPDtbTrans.value, True) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckEqp = True
Else
CheckEqp = False
End If
End Function
Sub FillGridEqup(Optional FixedID As Double, Optional Row As Long)
Dim i As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "SELECT     dbo.TblSpecificFixedDeti.ID, dbo.TblSpecificFixedDeti.LngT, dbo.TblSpecificFixedDeti.Price, dbo.TblSpecificFixedDeti.total, dbo.FixedAssets.code, "
sql = sql & "                      dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.TblSpecificFixedDeti.ToDate, dbo.TblSpecificFixedDeti.FromDate, dbo.TblSpecificFixedDeti.ProjectID,"
sql = sql & "                       dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblSpecificFixedDeti.PandID, dbo.projects_des.des, dbo.TblSpecificFixedDeti.OperID,"
sql = sql & "                       dbo.TblProcessDEF.ProcessName , dbo.TblProcessDEF.ProcessNameE, dbo.TblSpecificFixedDeti.FixedID, dbo.TblSpecificFixedDeti.SPFixID"
sql = sql & "  FROM         dbo.TblSpecificFixedDeti LEFT OUTER JOIN"
sql = sql & "                       dbo.TblProcessDEF ON dbo.TblSpecificFixedDeti.OperID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
sql = sql & "                       dbo.projects_des ON dbo.TblSpecificFixedDeti.PandID = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
sql = sql & "                       dbo.projects ON dbo.TblSpecificFixedDeti.ProjectID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                       dbo.FixedAssets ON dbo.TblSpecificFixedDeti.FixedID = dbo.FixedAssets.id"
sql = sql & "  WHERE     (dbo.TblSpecificFixedDeti.FixedID = " & FixedID & ") AND (dbo.TblSpecificFixedDeti.ToDate >=" & SQLDate(XPDtbTrans.value, True) & ") AND"
sql = sql & "                        (dbo.TblSpecificFixedDeti.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
With Fg_Journal
If .Rows = 2 Then
.Rows = .Rows + Rs3.RecordCount - 1
Else
.Rows = .Rows + Rs3.RecordCount - 2
End If
For i = Row To .Rows - 1
.TextMatrix(i, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("value"))) / Rs3.RecordCount
.TextMatrix(i, .ColIndex("branch_id")) = .TextMatrix(Row, .ColIndex("branch_id"))
.TextMatrix(i, .ColIndex("AccountCode")) = .TextMatrix(Row, .ColIndex("AccountCode"))
.TextMatrix(i, .ColIndex("ExpensesID")) = .TextMatrix(Row, .ColIndex("ExpensesID"))
.TextMatrix(i, .ColIndex("Destribute")) = .TextMatrix(Row, .ColIndex("Destribute"))
.TextMatrix(i, .ColIndex("branch_name")) = .TextMatrix(Row, .ColIndex("branch_name"))
.TextMatrix(i, .ColIndex("AccountName")) = .TextMatrix(Row, .ColIndex("AccountName"))
.TextMatrix(i, .ColIndex("Account_Serial")) = .TextMatrix(Row, .ColIndex("Account_Serial"))
.TextMatrix(i, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("value"))
.TextMatrix(i, .ColIndex("opr_fullcode")) = .TextMatrix(Row, .ColIndex("opr_fullcode"))
.TextMatrix(i, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des"))
.TextMatrix(i, .ColIndex("Des")) = .TextMatrix(Row, .ColIndex("Des"))
.TextMatrix(i, .ColIndex("CarName")) = .TextMatrix(Row, .ColIndex("CarName"))
.TextMatrix(i, .ColIndex("CarId")) = .TextMatrix(Row, .ColIndex("CarId"))
.TextMatrix(i, .ColIndex("Fixes")) = .TextMatrix(Row, .ColIndex("Fixes"))
.TextMatrix(i, .ColIndex("fixedid")) = .TextMatrix(Row, .ColIndex("fixedid"))
.TextMatrix(i, .ColIndex("deptid")) = .TextMatrix(Row, .ColIndex("deptid"))
.TextMatrix(i, .ColIndex("dept")) = .TextMatrix(Row, .ColIndex("dept"))
.TextMatrix(i, .ColIndex("projectid2")) = IIf(IsNull(Rs3("ProjectID").value), "", Rs3("ProjectID").value)
.TextMatrix(i, .ColIndex("pandid2")) = IIf(IsNull(Rs3("PandID").value), "", Rs3("PandID").value)
.TextMatrix(i, .ColIndex("operid2")) = IIf(IsNull(Rs3("OperID").value), "", Rs3("OperID").value)
.TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(Rs3("des").value), "", Rs3("des").value)
.TextMatrix(i, .ColIndex("LineNo1")) = setfoxy_Line
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(Rs3("ProcessName").value), "", Rs3("ProcessName").value)
.TextMatrix(i, .ColIndex("project")) = IIf(IsNull(Rs3("Project_name").value), "", Rs3("Project_name").value)
Else
.TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(Rs3("ProcessNameE").value), "", Rs3("ProcessNameE").value)
.TextMatrix(i, .ColIndex("project")) = IIf(IsNull(Rs3("Project_nameE").value), "", Rs3("Project_nameE").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
    Dim project_id As Integer
Dim rsDummy As New ADODB.Recordset
    With VSFlexGrid1

        Select Case .ColKey(Col)
         Case "branch_name"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("branch_id")) = StrAccountCode
         Case "dept"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("deptid")) = StrAccountCode
                
    Case "FixCode"
               Dim str As String
                str = " SELECT      TblCarsData.Fullcode, fixedassetid ,TblCarsData.EqupName,TblCarsData.BoardNO                FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  dbo.TblCarsData.Fullcode like '%" & Trim(.TextMatrix(Row, .ColIndex("FixCode"))) & "%'  "
                rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("fixedid")) = val(rsDummy!FixedassetId & "")
                    .TextMatrix(Row, .ColIndex("Fixes")) = Trim(rsDummy!EqupName & "")
                    '.TextMatrix(Row, .ColIndex("CarName")) = Trim(rsDummy!BoardNO & "")
            
                   ' Fg_Journal_AfterEdit Row, Fg_Journal.ColIndex("Fixes")
                Else
                  .TextMatrix(Row, .ColIndex("fixedid")) = ""
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("Fixes")) = ""
                    '.TextMatrix(Row, .ColIndex("CarName")) = ""
                End If
        Case "Fixes"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("fixedid")) = StrAccountCode

         Case "project"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("projectid2")) = StrAccountCode
                  Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid2")) = StrAccountCode
                  Case "oper"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid2")) = StrAccountCode
                
    
            Case "Value", "opr_fullcode"
    
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
    
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If

                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
   
            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If

                    '
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
            
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
       
                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                    
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
        
                'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                'Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                'End If
           
                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ReLineGrid

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "Value"
                .ComboList = ""

            Case "Account_Serial"
                .ComboList = ""
        
            Case "Des"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
    SendKeys "{F4}"
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.Show
        Account_search.case_id = 80

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Msg As String
    Dim project_id As Integer
    Dim whrstring As String

    With VSFlexGrid1

        Select Case .ColKey(Col)
''///
     Case "branch_name"
         If SystemOptions.UserInterface = ArabicInterface Then
       StrSQL = " SELECT     branch_id, branch_name From TblBranchesData"
      
       Else
        StrSQL = " SELECT     branch_id , branch_namee From TblBranchesData "
     
        End If
        
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_name", "branch_id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_namee", "branch_id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
     Case "dept"
         If SystemOptions.UserInterface = ArabicInterface Then
       StrSQL = " SELECT     DeparmentID, DepartmentName From dbo.TblEmpDepartments "
      
       Else
        StrSQL = " SELECT     DeparmentID , DepartmentNamee From dbo.TblEmpDepartments"
     
        End If
        
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
        Case "Fixes"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
                    Case "project"

               
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name , id From dbo.Projects  "

         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(rs, "Project_name", "id")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid2")) = "" Then
             If SystemOptions.UserInterface = ArabicInterface Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Else
             MsgBox "Please Select Project"
             End If
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ")"
           
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(rs, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  Case "oper"
                   
If .TextMatrix(Row, .ColIndex("projectid2")) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
Else
MsgBox "Please Select Project"
End If
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("pandid2")) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
Else
MsgBox "Please Select Des"
End If
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("pandid2"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ")"
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
''//
            Case "opr_fullcode"
            
                    
                project_id = get_project_id(dcproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = .BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = .BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
            
            Case "AccountName"
         
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                whrstring = getProjectAccountwhereString(project_id)
                
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '   End If
                    StrSQL = StrSQL + "and (" + whrstring + ")"
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    StrSQL = StrSQL + "and (" + whrstring + ")"
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                
                End If
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim CarID As Integer
    Dim CarName As String
    Dim mCode As String
    On Error GoTo ErrTrap
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 3
                 
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
          
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        'Lngid
        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If

  DCPreFix.Text = IIf(IsNull(rs("Prefix").value), "", rs("Prefix").value)

    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.txt_ORDER_NO.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    TXT_A_NoteID.Text = IIf(IsNull(rs("A_NoteID").value), "", (rs("A_NoteID").value))
    XPTxtID.Text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteserial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    XPMTxtRemarks.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtto.Text = IIf(IsNull(rs("too").value), "", rs("too").value)
    TxtManulaNO.Text = IIf(IsNull(rs("ManualNo").value), "", rs("ManualNo").value)
    
    txt_general_des.Text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
''//
 Me.TxtOrderID.Text = IIf(IsNull(rs("OrderID").value), "", rs("OrderID").value)
  Me.TxtNoteserial1.Text = IIf(IsNull(rs("Noteseril2").value), "", rs("Noteseril2").value)
''//
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(XPDtbTrans.value), rs("NoteDateH").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)

    If IsNull(rs("Destribute").value) Then
        chkDestribute.value = vbUnchecked
    ElseIf (rs("Destribute").value) = False Then
        chkDestribute.value = vbUnchecked
    Else
        chkDestribute.value = vbChecked
    End If
TxtVATCustoms.Text = IIf(IsNull(rs("VATCustoms").value), 0, rs("VATCustoms").value)
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPaymentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPaymentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
    ElseIf rs("NoteCashingType").value = 5 Then
        Me.CboPaymentType.ListIndex = 5
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPaymentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
    ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPaymentType.ListIndex = 4
        Me.DCAccounts.BoundText = IIf(IsNull(rs("AccountCode").value), "", rs("AccountCode").value)
        DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        '    DCVendor.BoundText = ""
     
    End If

    CboPayMentType_Change

    If Not IsNull(rs("BasedONID").value) Then
        Me.CBoBasedON.ListIndex = rs("BasedONID").value
    Else
        Me.CBoBasedON.ListIndex = 0
 
    End If
 
    'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
    'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

    If rs("NoteCashingType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Or rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.Text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

    Me.oldTxtSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)

    Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then 'Õ”«Ì« 
   '     Me.VSFlexGrid1.Visible = True
        Me.Fg_Journal.Visible = True

      '  StrSQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
      '  StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
      '  StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.projects.Project_name, dbo.projects.Project_nameE,"
      '  StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.projects_des.des, dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName,"
      '  StrSQL = StrSQL + "              dbo.TblProcessDEF.ProcessNameE , dbo.DOUBLE_ENTREY_VOUCHERS.FixedassetId, dbo.FixedAssets.name, dbo.FixedAssets.NameE"
      '  StrSQL = StrSQL + "   FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
      '  StrSQL = StrSQL + "             dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
      '  StrSQL = StrSQL + "              dbo.FixedAssets ON dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = dbo.FixedAssets.id LEFT OUTER JOIN"
      '  StrSQL = StrSQL + "              dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
      '  StrSQL = StrSQL + "              dbo.projects_des ON dbo.DOUBLE_ENTREY_VOUCHERS.pandid = dbo.projects_des.oprid LEFT OUTER JOIN"
      '  StrSQL = StrSQL + "              dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.projectid = dbo.projects.id"
      '  StrSQL = StrSQL + " Where (   dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
      '  StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
StrSQL = " SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID,"
StrSQL = StrSQL + "                        dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
StrSQL = StrSQL + "                        dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.projects.Project_name, dbo.projects.Project_nameE,"
StrSQL = StrSQL + "                        dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.projects_des.des, dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName,"
StrSQL = StrSQL + "                        dbo.TblProcessDEF.ProcessNameE, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, dbo.FixedAssets.Name, dbo.FixedAssets.namee,"
StrSQL = StrSQL + "                        dbo.DOUBLE_ENTREY_VOUCHERS.Departementid , dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,FixedAssets.Fullcode FixCode"
StrSQL = StrSQL + "   FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
StrSQL = StrSQL + "                        dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.FixedAssets ON dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.projects_des ON dbo.DOUBLE_ENTREY_VOUCHERS.pandid = dbo.projects_des.oprid and dbo.projects_des.oprid<>0 LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.projectid = dbo.projects.id"
StrSQL = StrSQL + "   Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
StrSQL = StrSQL + "   ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsDev.RecordCount > 0 Then
            RsDev.MoveFirst
        End If
    
        With Me.VSFlexGrid1
 
            .Rows = .FixedRows + RsDev.RecordCount
 
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            '    .TextMatrix(i, .ColIndex("AccountCodeVat")) = IIf(IsNull(RsDev("AccountCodeVat").value), "", RsDev("AccountCodeVat").value)
            
                .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
            
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                End If
        
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            .TextMatrix(i, .ColIndex("deptid")) = IIf(IsNull(RsDev("Departementid").value), "", RsDev("Departementid").value)
                .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
              ''//
                    .TextMatrix(i, .ColIndex("projectid2")) = IIf(IsNull(RsDev("projectid").value), "", RsDev("projectid").value)
                    .TextMatrix(i, .ColIndex("pandid2")) = IIf(IsNull(RsDev("pandid").value), "", RsDev("pandid").value)
                    .TextMatrix(i, .ColIndex("operid2")) = IIf(IsNull(RsDev("operid").value), "", RsDev("operid").value)
                    .TextMatrix(i, .ColIndex("fixedid")) = IIf(IsNull(RsDev("FixedAssetId").value), "", RsDev("FixedAssetId").value)
                    .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                    .TextMatrix(i, .ColIndex("FixCode")) = IIf(IsNull(RsDev("FixCode").value), "", RsDev("FixCode").value)
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("dept")) = IIf(IsNull(RsDev("DepartmentName").value), "", RsDev("DepartmentName").value)
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                    .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                    .TextMatrix(i, .ColIndex("Fixes")) = IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                    Else
                    .TextMatrix(i, .ColIndex("dept")) = IIf(IsNull(RsDev("DepartmentNamee").value), "", RsDev("DepartmentNamee").value)
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_nameE").value), "", RsDev("Project_nameE").value)
                    .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessNameE").value), "", RsDev("ProcessNameE").value)
                    .TextMatrix(i, .ColIndex("Fixes")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                    
                    End If
                RsDev.MoveNext
            Next i
    
        End With

        'Exit Sub
        GoTo EndMEe
    End If

    Me.VSFlexGrid1.Visible = False
    Me.Fg_Journal.Visible = True

    '«·„’—Ê›« 
    '-----------------------------------------------------------------------------
    If chkDestribute.value = vbUnchecked Then
        '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
        '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        ' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"

        'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
       ' StrSQL = "SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, "
       ' StrSQL = StrSQL + "               dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,"
       ' StrSQL = StrSQL + "               dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
       ' StrSQL = StrSQL + "               dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
       ' StrSQL = StrSQL + "               dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
       ' StrSQL = StrSQL + "                dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.ORDER_NO, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId,"
       ' StrSQL = StrSQL + "               dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.projects.Project_name, dbo.projects.Project_nameE,"
       ' StrSQL = StrSQL + "               dbo.DOUBLE_ENTREY_VOUCHERS.pandid , dbo.projects_des.des, dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName , dbo.TblProcessDEF.ProcessNameE "
       ' StrSQL = StrSQL + " FROM         dbo.ACCOUNTS INNER JOIN"
       ' StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
       ' StrSQL = StrSQL + "               dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
       ' StrSQL = StrSQL + "               dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
       ' StrSQL = StrSQL + "               dbo.projects_des ON dbo.DOUBLE_ENTREY_VOUCHERS.pandid = dbo.projects_des.oprid LEFT OUTER JOIN"
       ' StrSQL = StrSQL + "               dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.projectid = dbo.projects.id LEFT OUTER JOIN"
       ' StrSQL = StrSQL + "               dbo.FixedAssets ON dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = dbo.FixedAssets.id"
       ' StrSQL = StrSQL + " Where ( hideline is null and dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
       ' StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
     StrSQL = " SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, "
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
     StrSQL = StrSQL + "                  dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.ORDER_NO, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId,"
     StrSQL = StrSQL + "                 dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.projects.Project_name, dbo.projects.Project_nameE,"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.projects_des.des, dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName,"
     StrSQL = StrSQL + "                 dbo.TblProcessDEF.ProcessNameE, dbo.DOUBLE_ENTREY_VOUCHERS.Departementid, dbo.TblEmpDepartments.DepartmentName,"
     StrSQL = StrSQL + "                 dbo.TblEmpDepartments.DepartmentNamee, dbo.ACCOUNTS.Account_Serial, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.branch_name,"
     StrSQL = StrSQL + "                 dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Vat, dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo,"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat, dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierName,"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS.CusVATNO, dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID,"
     StrSQL = StrSQL + "                 dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS Expr2, dbo.DOUBLE_ENTREY_VOUCHERS.Rate2 , dbo.DOUBLE_ENTREY_VOUCHERS.BillNo,"
     StrSQL = StrSQL + "                 Notes.TradingContractID,TblCarsData.Fullcode FixCode"
     StrSQL = StrSQL + "      FROM         dbo.ACCOUNTS INNER JOIN"
     StrSQL = StrSQL + "                 dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
     StrSQL = StrSQL + "                 dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.TblCustemers ON dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.projects_des ON dbo.DOUBLE_ENTREY_VOUCHERS.pandid = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.projectid = dbo.projects.id LEFT OUTER JOIN"
     StrSQL = StrSQL + "                 dbo.FixedAssets ON dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = dbo.FixedAssets.id"
     StrSQL = StrSQL + "                 Left Outer Join TblCarsData On TblCarsData.FixedassetId = dbo.FixedAssets.id"
     
     StrSQL = StrSQL + "       WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.hideline IS NULL) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
     StrSQL = StrSQL + "                  (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.Text) & ")"
     StrSQL = StrSQL + "  ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal

                If Me.dcproject.BoundText = "" Then
                    .Rows = .FixedRows + RsDev.RecordCount
                Else
                    .Rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("SupplierID")) = IIf(IsNull(RsDev("SupplierID").value), 0, RsDev("SupplierID").value)
                    .TextMatrix(i, .ColIndex("CusVATNO")) = IIf(IsNull(RsDev("CusVATNO").value), "", RsDev("CusVATNO").value)
                    .TextMatrix(i, .ColIndex("SupplierName")) = IIf(IsNull(RsDev("SupplierName").value), "", RsDev("SupplierName").value)
                   ' .TextMatrix(i, .ColIndex("AccountCodeVat")) = IIf(IsNull(RsDev("AccountCodeVat").value), "", RsDev("AccountCodeVat").value)
                    .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsDev("PriceTotal").value), 0, RsDev("PriceTotal").value)
                    .TextMatrix(i, .ColIndex("Rate")) = IIf(IsNull(RsDev("Rate2").value), 0, RsDev("Rate2").value)
                    .TextMatrix(i, .ColIndex("BillNo")) = IIf(IsNull(RsDev("BillNo").value), "", RsDev("BillNo").value)
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
                    .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), 0, RsDev("FlgVat").value)
                    .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                    .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), 0, RsDev("Vat").value)
                    .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value))
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
                    .TextMatrix(i, .ColIndex("PrjectCode")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                    .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(RsDev("Account_Serial").value), "", RsDev("Account_Serial").value)
                    CarID = IIf(IsNull(RsDev("CarID").value), 0, RsDev("CarID").value)
                    .TextMatrix(i, .ColIndex("FixCode")) = IIf(IsNull(RsDev("FixCode").value), "", RsDev("FixCode").value)
                    If CarID <> 0 Then
                        
                        GetCarName CarID, CarName
                        .TextMatrix(i, .ColIndex("CarId")) = IIf(IsNull(RsDev("CarID").value), "", RsDev("CarID").value)
                        
                        .TextMatrix(i, .ColIndex("CarName")) = CarName
                       
                    End If
                    .TextMatrix(i, .ColIndex("branch_id")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
                     .TextMatrix(i, .ColIndex("deptid")) = IIf(IsNull(RsDev("Departementid").value), "", RsDev("Departementid").value)
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("Supplier")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                        .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                        .TextMatrix(i, .ColIndex("dept")) = IIf(IsNull(RsDev("DepartmentName").value), "", RsDev("DepartmentName").value)
                    Else
                        .TextMatrix(i, .ColIndex("Supplier")) = IIf(IsNull(RsDev("CusNamee").value), "", RsDev("CusNamee").value)
                        .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                        .TextMatrix(i, .ColIndex("dept")) = IIf(IsNull(RsDev("DepartmentNamee").value), "", RsDev("DepartmentNamee").value)
                    End If
        
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
        
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
                    ''//
                    .TextMatrix(i, .ColIndex("projectid2")) = IIf(IsNull(RsDev("projectid").value), "", RsDev("projectid").value)
                    .TextMatrix(i, .ColIndex("pandid2")) = IIf(IsNull(RsDev("pandid").value), "", RsDev("pandid").value)
                    .TextMatrix(i, .ColIndex("operid2")) = IIf(IsNull(RsDev("operid").value), "", RsDev("operid").value)
                    .TextMatrix(i, .ColIndex("fixedid")) = IIf(IsNull(RsDev("FixedAssetId").value), "", RsDev("FixedAssetId").value)
                    .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                    .TextMatrix(i, .ColIndex("TradingContractID")) = IIf(IsNull(RsDev("TradingContractID").value), 0, RsDev("TradingContractID").value)
                    
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                    .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                    .TextMatrix(i, .ColIndex("Fixes")) = IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                    Else
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_nameE").value), "", RsDev("Project_nameE").value)
                    .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessNameE").value), "", RsDev("ProcessNameE").value)
                    .TextMatrix(i, .ColIndex("Fixes")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)

                    
                    End If
                    If val(.TextMatrix(i, .ColIndex("PriceTotal"))) = 0 And SystemOptions.IsMergeVat Then
                        If i > 1 Then
                            If val(.TextMatrix(i - 1, .ColIndex("Vatyo"))) <> 0 Then
                                .RowHidden(i) = True
                            End If
                        End If
                        
                    End If
                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

            End With

        End If

    End If

    '-----------------------------------------------------------------------------«·„’—Ê›«  «·„Ê“⁄Â
    If chkDestribute.value = vbChecked Then
    
        'StrSQL = "SELECT dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
        'StrSQL = StrSQL + " FROM         dbo.ACCOUNTS INNER JOIN"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        'StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & Val(Me.XPTxtID.text) & ")"
        'StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        StrSQL = "Select * from ExpensesDetails where noteid=" & val(XPTxtID.Text)
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            '  Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            '  Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst
            '        For i = 1 To RsDev.RecordCount
            '            If RsDev("Credit_Or_Debit").value = 0 Then
            '                Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
            '            ElseIf RsDev("Credit_Or_Debit").value = 1 Then
            '                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
            '            End If
            '            RsDev.MoveNext
            '        Next i
    
            '      RsDev.MoveFirst
    
            With Me.Fg_Journal

                If Me.dcproject.BoundText = "" Then
                    .Rows = .FixedRows + RsDev.RecordCount
                Else
                    .Rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("Destribute")) = IIf(IsNull(RsDev("Destribute").value), 0, RsDev("Destribute").value)
            
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
            
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
            
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value))
            
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
            
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("ExpensesName").value), "", RsDev("ExpensesName").value)
                    Else
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("ExpensesName").value), "", RsDev("ExpensesName").value)
                    End If
        
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Des").value), "", RsDev("Des").value)
        
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
'                     If val(.TextMatrix(i, .ColIndex("PriceTotal"))) = 0 And SystemOptions.IsMergeVat Then
'                        .RowHidden = True
'                    End If

                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

            End With

        End If

    End If

EndMEe:

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ReLineGrid
    FillDestributionsToAll
fillapprovData
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
     Dim Msg As String
     Dim StrAccount As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim bankDes As String
    Dim OtherInformation As New ClsGLOther
  
    On Error GoTo ErrTrap
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    If Me.TxtModFlg.Text <> "R" Then

        If Me.CboPaymentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPaymentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPaymentType.ListIndex = 4 Then
            If Trim(Me.DCAccounts.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— ·Õ”«»..!!"
                Else
                    Msg = "Select Account..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DCAccounts.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        
        End If
   
        If Me.CboPaymentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPaymentType.ListIndex = 1 Or Me.CboPaymentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
                    
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        
            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
         
        ElseIf Me.CboPaymentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
                                        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·…...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            
            End If
       
        ElseIf Me.CboPaymentType.ListIndex = 5 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
                                        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·«„—...!!"
                Else
                    Msg = "Enter Bank Order No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            
            End If
       
        End If

        '      If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
        '      If SystemOptions.UserInterface = ArabicInterface Then
        '          Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
        '      Else
        '      Msg = "Cheque Due Date Not Valid...!!"
        '
        '      End If
        '          MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '          DtpChequeDueDate.SetFocus
        '          SendKeys "{F4}"
        '          Exit Sub
        '      End If
    
    End If
   
    If CheckAllExpensesDistributed = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·”‰œ ÌÕ ÊÏ ⁄·Ï „’«—Ì› „Ê“—⁄Â «Œ—Ï €Ì— „Ê“⁄Â Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
        Else
            MsgBox "This Voucher Have Distributed and not Distributed Expenses", vbCritical
        End If
               
        Exit Sub
    End If
 
    If Me.TxtModFlg.Text = "N" Then
        If Me.CboPaymentType.ListIndex = 0 Then
            If val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value) = False Then
                    Exit Sub
                End If
            End If
        End If

    ElseIf Me.TxtModFlg.Text = "E" Then

        If Me.CboPaymentType.ListIndex = 0 Then
            If val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value, , , val(Me.XPTxtID.Text)) = False Then
                    Exit Sub
                End If
            End If
        End If
    End If

    Dim xrow As Integer

    With Fg_Journal

        For xrow = .Rows - 1 To 2 Step -1

            If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                .Rows = .Rows - 1
            End If

        Next xrow

    End With
    
    With Me.VSFlexGrid1

        For xrow = .Rows - 1 To 2 Step -1

            If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                .Rows = .Rows - 1
            End If

        Next xrow

    End With

    Dim i As Integer

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then

        With VSFlexGrid1

            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                    '////////////////////////////////////////notes
               
                    If SystemOptions.UserInterface = ArabicInterface Then
              '          MsgBox "·«  ÌÊÃœ Õ”«»  ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                    Else
              '          MsgBox "Select Acc in line no" & i, vbCritical
                    End If

              '      Exit Sub
              
                End If
        
            Next i

        End With

        With VSFlexGrid1

            For i = .FixedRows To .Rows - 1

                If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
                    '////////////////////////////////////////notes
               
                    If SystemOptions.UserInterface = ArabicInterface Then
             '           MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                    Else
             '           MsgBox "Enter Value in line no" & i, vbCritical
                    End If
               
             '       Exit Sub
                End If
        
            Next i

        End With

        GoTo xx
    End If
 If dcproject.Text = "" Then
    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                '////////////////////////////////////////notes
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«  ÌÊÃœ „’—Ê› ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                Else
                    MsgBox "Select Expenses in line no" & i, vbCritical
                End If

                Exit Sub
              
            End If
        
        Next i

    End With

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
                '////////////////////////////////////////notes
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                Else
                    MsgBox "Enter Value in line no" & i, vbCritical
                End If
               
                Exit Sub
            End If
        
        Next i

    End With
    End If
    Dim ISVAT As Boolean
    ISVAT = False
With Fg_Journal
    For i = .FixedRows To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("Vat"))) > 0 Then
      ISVAT = True
      End If
     Next i
 End With
     
Dim AccountVATDept As String
If ISVAT = True Then
If GetValueAddedAccount(XPDtbTrans.value, AccountVATDept) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
xx:
    calcnets     '-------------------------------------------------------------------------------------------
 
    '-------------------------------------------------------------------------------------------
    
    If TxtSerial.Text = "" Then
        If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
            End If

        Else
         
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    MsgBox "You must Define JE Coding ": Exit Sub
                End If

            Else
                TxtSerial.Text = Notes_coding(val(my_branch), XPDtbTrans.value)
            End If
        End If
    End If

    ' TxtSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value) 'kk
    If TxtSerial1.Text = "" Then
        If Voucher_coding(val(my_branch), XPDtbTrans.value, 1, 3, , , DCPreFix.Text) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
            End If

        Else
         
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 1, 3, , , DCPreFix.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                End If

            Else
                TxtSerial1.Text = Voucher_coding(val(my_branch), XPDtbTrans.value, 1, 3, , , DCPreFix.Text)
            End If
        End If
    End If
    
    Cn.BeginTrans
    BeginTrans = True
    Dim A_NoteID As Long

    '///////////////NOTESALL
    If TxtModFlg.Text = "N" Then
        XPTxtID.Text = CStr(new_id("notes_all", "NoteID", "", True))
        Me.TxtNoteserial.Text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        rs.AddNew
        rs("NoteID").value = val(XPTxtID.Text)
        Me.oldTxtSerial1.Text = Trim$(Me.TxtSerial1.Text)
         
    ElseIf Me.TxtModFlg.Text = "E" Then
     '   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.Text)
     '   Cn.Execute StrSQL, , adExecuteNoRecords
        
        StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        If DcCostCenter.BoundText <> "" Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        
        StrSQL = "Delete From ExpensesDetails Where Noteid =" & val(XPTxtID.Text) & "  or NoteSerial1='" & Me.TxtSerial1.Text & "'"
        Cn.Execute StrSQL, , adExecuteNoRecords
        
    End If
    
    '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
   rs("Prefix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)
   
    rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    rs("foxy_no").value = val(Text1.Text)
    rs("order_no").value = txt_ORDER_NO.Text
    
    rs("Note_Value").value = IIf(XPTxtVal.Text = "", Null, XPTxtVal.Text)
    rs("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
    rs("too").value = IIf(txtto.Text = "", "", Trim(txtto.Text))
    rs("general_des").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text)) & bankDes
    rs("ManualNo").value = IIf(TxtManulaNO.Text = "", "", Trim(TxtManulaNO.Text))
    
    rs("branch_no").value = val(Me.Dcbranch.BoundText)
    ''/
     rs("OrderID").value = IIf(Me.TxtOrderID.Text = "", Null, Trim(TxtOrderID.Text))
     rs("Noteseril2").value = IIf(Me.TxtNoteserial1.Text = "", "", Trim(TxtNoteserial1.Text))
    ''/
    rs("CusID").value = Null
    rs("NoteType").value = 3
    rs("NoteDate").value = XPDtbTrans.value
   ' rs("NoteDate").value = Format$(Date, "dd-mm-yyyy")
    rs("NoteDateH").value = Me.Txt_DateHigri.value

    rs("UserID").value = user_id

    If chkDestribute.value = vbChecked Then
        Destribute = True
    Else
        Destribute = False
    End If

    rs("Destribute").value = Destribute
    rs("ExpensesID").value = IIf(XPCboExpensesType.Text = "", Null, XPCboExpensesType.BoundText)
    rs("VATCustoms").value = val(TxtVATCustoms.Text)
    If CBoBasedON.ListIndex > -1 Then
        rs("BasedONID").value = CBoBasedON.ListIndex
    Else
        rs("BasedONID").value = 0
    End If
  
    If Me.CboPaymentType.ListIndex = 0 Then
        rs("BoxID").value = val(DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 0

        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "   ’—› „‰  " & DcboBox.Text
        Else
            bankDes = "   Payed From  " & DcboBox.Text
        End If

    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 1

        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "  ’—› »‘Ìﬂ —ﬁ„  " & TxtChequeNumber.Text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.Text
        Else
            bankDes = "  Check No:  " & TxtChequeNumber.Text & "  Bank:  " & DcboBankName.Text
        
        End If
        
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 3

        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "  ’—› »‘Ìﬂ „”œœ —ﬁ„  " & TxtChequeNumber.Text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.Text
        Else
            bankDes = "  Check No:  " & TxtChequeNumber.Text & "  Bank:  " & DcboBankName.Text
        
        End If
  
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 2

        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "  ’—› »ÕÊ«·…  —ﬁ„  " & TxtChequeNumber.Text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.Text
        Else
            bankDes = "  Bank Transfere No:  " & TxtChequeNumber.Text & "  Bank:  " & DcboBankName.Text
        End If
        
    ElseIf Me.CboPaymentType.ListIndex = 5 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 5

        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "  ’—› »√„— »‰ﬂÌ   —ﬁ„  " & TxtChequeNumber.Text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.Text
        Else
            bankDes = "  Bank Transfere No:  " & TxtChequeNumber.Text & "  Bank:  " & DcboBankName.Text
        End If
        
    ElseIf Me.CboPaymentType.ListIndex = 4 Then
        rs("BoxID").value = Null
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 4
        rs("AccountCode").value = (Me.DCAccounts.BoundText)
        
        If SystemOptions.UserInterface = ArabicInterface Then
            '    bankDes = txt_general_des.text
        Else
            '    bankDes = txt_general_des.text
        
        End If
        
    End If
    
    rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
    rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.Text) = "", Null, Trim$(Me.Txt_Numorder.Text))
    rs("Buy").value = "0"
    rs("Remark").value = XPMTxtRemarks.Text
    rs("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
    rs("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
    rs("OldNoteSerial1").value = Trim$(Me.oldTxtSerial1.Text) '
'   rs("ManualNo").value = IIf(Trim(Me.TxtManulaNO.text) = "", Null, Trim(Me.TxtManulaNO.text))
     rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    rs("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
     
    rs("sanad_year").value = year(XPDtbTrans.value)
    rs("sanad_month").value = month(XPDtbTrans.value)
'    rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
    
    If Me.TxtModFlg.Text = "N" Then
        A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
        TXT_A_NoteID.Text = A_NoteID
    Else
        A_NoteID = val(TXT_A_NoteID.Text)
    End If
    
    rs("A_NoteID").value = val(A_NoteID)
     
    rs.update
    Dim project_id As Integer
    project_id = get_project_id(dcproject.BoundText, "expanses_account")
    '/////////////////////Accounts Õ”«Ì« 
    Dim line_no  As Integer

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then
        Set RsNotes = New ADODB.Recordset
        'RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
        If TxtModFlg.Text = "N" Then
           
        ElseIf Me.TxtModFlg.Text = "E" Then
     '       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
     '       Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If
    
        '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        ' rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        ' rs("foxy_no").value = Val(Text1.text)
        'œ«∆‰ Õ”«»«  «·„‘—Ê€
        RsNotes.AddNew
        RsNotes("NoteID").value = A_NoteID
         RsNotes.update
        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
        RsNotes("order_no").value = txt_ORDER_NO.Text
        RsNotes("notes_all").value = Me.XPTxtID.Text
        RsNotes("Note_Value").value = IIf(Not IsNumeric(XPTxtVal.Text), 0, val(XPTxtVal.Text)) + val(TxtVATCustoms.Text)
        RsNotes("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
     RsNotes("ManualNO").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))

        RsNotes("too").value = IIf(txtto.Text = "", "", Trim(txtto.Text))
        '    RsNotes("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
        If Me.CboPaymentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 1
        
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 3
            
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 2
        
        ElseIf Me.CboPaymentType.ListIndex = 5 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 5
        
            ' ElseIf Me.CboPaymentType.ListIndex = 2 Then
            ' RsNotes("CusID").value = DCVendor.BoundText
        End If
     
        '     RsNotes("BasedONID").value = Me.CBoBasedON.ListIndex
    
        RsNotes("NoteType").value = 3
         RsNotes("NoteDate").value = XPDtbTrans.value
       ' RsNotes("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        RsNotes("NoteDateH").value = Me.Txt_DateHigri.value
     
        RsNotes("UserID").value = user_id
    
        'rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
        'rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
        RsNotes("Buy").value = "0"
        RsNotes("Remark").value = txt_general_des.Text & bankDes
        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„   ”‰œ ’—›
     RsNotes("ManualNO").value = IIf(Trim(TxtManulaNO.Text) = "", Null, Trim(TxtManulaNO.Text))
        RsNotes("sanad_year").value = year(XPDtbTrans.value)
        RsNotes("sanad_month").value = month(XPDtbTrans.value)
        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
   RsNotes("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
        RsNotes.update
    
        Dim IntDEV_Type As Integer
        Dim SngDEV_Value As Single
        line_no = 1
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            
        If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, IIf(Not IsNumeric(XPTxtVal.Text), 0, val(XPTxtVal.Text)), 1, txt_general_des.Text & bankDes, A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
            GoTo ErrTrap
                    
        End If
            
        '„œÌ‰ Õ”«»«  «·„‘—Ê€
        With VSFlexGrid1
            line_no = 2
 
            For i = .FixedRows To .Rows - 1
                
                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
                    If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("Value")), 0, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.Text), project_id, .TextMatrix(i, .ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , val(Dcbranch.BoundText), , , , , , , val(.TextMatrix(i, .ColIndex("deptid"))), , , , , , , val(.TextMatrix(i, .ColIndex("projectid2"))), val(.TextMatrix(i, .ColIndex("pandid2"))), val(.TextMatrix(i, .ColIndex("operid2"))), , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                    
                    End If

                    line_no = line_no + 1
            
                End If

            Next i

        End With
            If val(TxtVATCustoms.Text) > 0 Then
    
    StrAccount = get_account_code_branch(148, my_branch)
    
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccount, val(TxtVATCustoms.Text), 0, txt_general_des.Text & "Õ”«» «·ﬁÌ„… «·„÷«›… ··Ã„«—ﬂ ›Ì ”‰œ ’—›  Õ·Ì· „’—Ê›« ", A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
            GoTo ErrTrap
        End If
        line_no = line_no + 1
    End If
        ' TxtModFlg.text = "R"
        GoTo ll
      
    End If

    '„’—Ê›« 
    
    '//////////////////////////////////////Notes////////////////////////////////////
    If Destribute = True Then
        If createDest = True Then
            GoTo ll
        Else
            Exit Sub
        End If
    End If

    Set RsNotes = New ADODB.Recordset
   ' RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
       
        Set RsDev = New ADODB.Recordset
       ' RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
        '«·ÿ—› «·„œÌ‰
  
        Dim ExpensesID As Double
 
        Dim NoteID As String

        With Fg_Journal
 
            line_no = 1

            For i = .FixedRows To .Rows - 1
 
                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                    '////////////////////////////////////////notes
                
                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                        Else
                            MsgBox "Cant save no value in line no:  " & i - 1, vbCritical: GoTo ErrTrap
                        End If
               
                    End If

                    RsNotes.AddNew
                    NoteID = CStr(new_id("Notes", "NoteID", "", True))
                    RsNotes("NoteID").value = CStr(NoteID)
                
                
                    RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value"))
                    RsNotes("Destribute").value = IIf(.TextMatrix(i, .ColIndex("Destribute")) = "", 0, Destribute)
                
                    RsNotes("Remark").value = txt_general_des.Text & bankDes
                    RsNotes("ExpensesRemark").value = .TextMatrix(i, .ColIndex("des"))
                 RsNotes("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
                    RsNotes("foxy_no").value = val(Text1.Text)
                    RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
RsNotes("ManualNO").value = IIf(Trim(TxtManulaNO.Text) = "", Null, Trim(TxtManulaNO.Text))
RsNotes("Prefix").value = IIf(Trim(DCPreFix.Text) = "", Null, Trim(DCPreFix.Text))

 
                    If Me.CboPaymentType.ListIndex = 0 Then
                        RsNotes("BoxID").value = val(DcboBox.BoundText)
                        RsNotes("BankID").value = Null
                        RsNotes("ChqueNum").value = Null
                        RsNotes("DueDate").value = Null
                        RsNotes("NoteCashingType").value = 0
                    ElseIf Me.CboPaymentType.ListIndex = 1 Then
                        RsNotes("BoxID").value = Null
                        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        RsNotes("NoteCashingType").value = 1
                            
                    ElseIf Me.CboPaymentType.ListIndex = 3 Then
                        RsNotes("BoxID").value = Null
                        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        RsNotes("NoteCashingType").value = 3
                            
                    ElseIf Me.CboPaymentType.ListIndex = 2 Then
                        RsNotes("BoxID").value = Null
                        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        RsNotes("NoteCashingType").value = 2
                        
                    ElseIf Me.CboPaymentType.ListIndex = 5 Then
                        RsNotes("BoxID").value = Null
                        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        RsNotes("NoteCashingType").value = 5
                            
                    End If

                    If txt_ORDER_NO.Text <> "" Then
                        RsNotes("order_no").value = txt_ORDER_NO.Text
                    Else
                        RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                    End If

                    RsNotes("CusID").value = Null
                    RsNotes("NoteType").value = 3
                    RsNotes("NoteDate").value = XPDtbTrans.value
                   ' RsNotes("NoteDate").value = Format$(Date, "dd-mm-yyyy")
                    RsNotes("NoteDateH").value = Me.Txt_DateHigri.value
                    RsNotes("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
   'RsNotes("ManualNO").value = IIf(Trim(TxtManulaNO.text) = "", Null, Trim(TxtManulaNO.text))
                    RsNotes("UserID").value = user_id
                    RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                    RsNotes("notes_all").value = Me.XPTxtID.Text
                    RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
                    RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
                    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                    RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
                    RsNotes("Prefix").value = IIf(Trim(DCPreFix.Text) = "", Null, Trim(DCPreFix.Text))
                    RsNotes("sanad_year").value = year(XPDtbTrans.value)
                    RsNotes("sanad_month").value = month(XPDtbTrans.value)
                    RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                    RsNotes("remark").value = txt_general_des.Text & bankDes
                      RsNotes("ProjectID").value = val(.TextMatrix(i, .ColIndex("projectid2")))
                      
                        RsNotes("Pand").value = val(.TextMatrix(i, .ColIndex("pandid2")))
                        RsNotes("Oper").value = val(.TextMatrix(i, .ColIndex("operid2")))
                        RsNotes("fixedid").value = val(.TextMatrix(i, .ColIndex("fixedid")))
                        RsNotes("TradingContractID").value = val(.TextMatrix(i, .ColIndex("TradingContractID")))
                        
                    RsNotes.update
              
                    '////////////////////////////////////////notes
   
                    project_id = get_project_id(dcproject.BoundText, "expanses_account")
                    
     If project_id = 0 Then
project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
     End If
project_id = val(.TextMatrix(i, .ColIndex("projectid2")))

Dim Material_account As String
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                     If Destribute = False Then
                     If 1 = 1 Then
                     project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
                     OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                     OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                     OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                     OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.SupplierID = val(.TextMatrix(i, .ColIndex("SupplierID")))
                    OtherInformation.CusVATNO = (.TextMatrix(i, .ColIndex("CusVATNO")))
                    OtherInformation.SupplierName = (.TextMatrix(i, .ColIndex("SupplierName")))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                    OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))
                    OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
                        
                             If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("value"))), 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, val(Me.DCboUserName.BoundText), , , , val(.TextMatrix(i, .ColIndex("value"))), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.Text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , val(.TextMatrix(i, .ColIndex("branch_id"))), val(.TextMatrix(i, .ColIndex("CarId"))), , , , , , val(.TextMatrix(i, .ColIndex("deptid"))), , , , , , .TextMatrix(i, .ColIndex("BillNo")), project_id, val(.TextMatrix(i, .ColIndex("pandid2"))), val(.TextMatrix(i, .ColIndex("operid2"))), , , , , , , , , Posted, , OtherInformation) = False Then
                                 GoTo ErrTrap
                            End If

        line_no = line_no + 1
                        End If
            
    '    MsgBox "xx"
        
            On Error GoTo ErrTrap
                               'ÃÊ«—Ì
                                  Dim BranchID As Integer
    Dim BranchID2 As Integer
    
                       Dim DeptSide As String
                        Dim credit_side As String
                       Dim total_value As Double
BranchID = val(Me.Dcbranch.BoundText)
BranchID2 = val(.TextMatrix(i, .ColIndex("branch_id")))
DeptSide = getBranchCurrentAccount(BranchID)
  credit_side = getBranchCurrentAccount(BranchID2)
   total_value = Round(.TextMatrix(i, .ColIndex("value")), 2)


   If BranchID <> BranchID2 Then
                                                              line_no = line_no + 1
                                                  '????
                                               If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, total_value, 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                                                   
                                                End If
                                                
                                                              
                                                              line_no = line_no + 1
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, line_no, DeptSide, total_value, 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                        line_no = line_no + 1
     End If
     
     
      
    If project_id <> 0 Then
      line_no = line_no + 1
      
      Material_account = get_project_Account(project_id, "expanses_account")
                If SystemOptions.gldetails_or_gl_general = 1 Then
                project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
                     OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                    OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                    OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                    OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.SupplierID = val(.TextMatrix(i, .ColIndex("SupplierID")))
                    OtherInformation.CusVATNO = (.TextMatrix(i, .ColIndex("CusVATNO")))
                    OtherInformation.SupplierName = (.TextMatrix(i, .ColIndex("SupplierName")))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                    OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))
                    OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
                    
                                               If ModAccounts.AddNewDev(LngDevID, line_no, Material_account, .TextMatrix(i, .ColIndex("value")), 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.Text), , .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , val(.TextMatrix(i, .ColIndex("branch_id"))), val(.TextMatrix(i, .ColIndex("CarId"))), , , , , , , , , , , , .TextMatrix(i, .ColIndex("BillNo")), , val(.TextMatrix(i, .ColIndex("pandid2"))), val(.TextMatrix(i, .ColIndex("operid2"))), , 1, , , , , , , Posted, , OtherInformation) = False Then
                                                       GoTo ErrTrap
                                               
                                              End If
                                            line_no = line_no + 1
                                            If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")), 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.Text), , .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , val(.TextMatrix(i, .ColIndex("branch_id"))), val(.TextMatrix(i, .ColIndex("CarId"))), , , , , , , , , , , , , , val(.TextMatrix(i, .ColIndex("pandid2"))), val(.TextMatrix(i, .ColIndex("operid2"))), , 1, , , , , , , Posted, , OtherInformation) = False Then
                                                  GoTo ErrTrap
                                          
                                              End If
                        
               End If
        End If
      
                    End If
                End If
            Next i
        End With
    ''/////////////
    
    If val(TxtVATCustoms.Text) > 0 Then
                    OtherInformation.FlgVat = 0
                    OtherInformation.Vat = 0
                    OtherInformation.Vatyo = 0
                    OtherInformation.CurrRow = 0
                    OtherInformation.SupplierID = 0
                    OtherInformation.CusVATNO = ""
                    OtherInformation.SupplierName = ""
                    OtherInformation.PriceTotal = 0
                    OtherInformation.Rate = 0
                    OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
    StrAccount = get_account_code_branch(148, my_branch)
    line_no = line_no + 1
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccount, val(TxtVATCustoms.Text), 0, txt_general_des.Text & "Õ”«» «·ﬁÌ„… «·„÷«›… ··Ã„«—ﬂ ›Ì ”‰œ ’—›  Õ·Ì· „’—Ê›« ", val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , 1, , , , , , , Posted, , OtherInformation) = False Then
            GoTo ErrTrap
        End If
    End If
    
        '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
        RsNotes.AddNew
        NoteID = CStr(new_id("Notes", "NoteID", "", True))
        RsNotes("NoteID").value = CStr(NoteID)
       RsNotes.update
        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
 
        RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
        RsNotes("Remark").value = txt_general_des.Text & bankDes 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsNotes("foxy_no").value = val(Text1.Text)
'RsNotes("ManualNO").value = IIf(Trim(TxtManulaNO.text) = "", Null, Trim(TxtManulaNO.text))
RsNotes("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))

        If Me.CboPaymentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPaymentType.ListIndex = 1 Or Me.CboPaymentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 1
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 3
                            
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 2
        ElseIf Me.CboPaymentType.ListIndex = 5 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 5
        End If
                        
        '                       If txt_ORDER_NO.text <> "" Then
        '           RsNotes("order_no").value = txt_ORDER_NO.text
        '       Else
        '        RsNotes("order_no").value = IIf(Me.Fg_Journal.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
        '       End If
            
        RsNotes("CusID").value = Null
        RsNotes("NoteType").value = 3
        RsNotes("NoteDate").value = XPDtbTrans.value
       ' RsNotes("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        RsNotes("NoteDateH").value = Me.Txt_DateHigri.value
   RsNotes("ManualNO").value = IIf(Trim(TxtManulaNO.Text) = "", Null, Trim(TxtManulaNO.Text))
        RsNotes("UserID").value = user_id
        ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
        RsNotes("notes_all").value = Me.XPTxtID.Text
        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
        RsNotes("sanad_year").value = year(XPDtbTrans.value)
        RsNotes("sanad_month").value = month(XPDtbTrans.value)
        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
        RsNotes("Remark").value = txt_general_des.Text & bankDes 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsNotes.update
     If SystemOptions.AnalyticPaymentVouchr = True Or SystemOptions.AllowAnalyticJL = True Then
      With Fg_Journal
       For i = .FixedRows To .Rows - 1
        If SystemOptions.IsMergeVat And .RowHidden(i) Then GoTo NextRow
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
        If Not SystemOptions.IsMergeVat Then
            total_value = Round(.TextMatrix(i, .ColIndex("value")), 2)
        Else
             total_value = Round(.TextMatrix(i, .ColIndex("value")), 2) + Round(.TextMatrix(i, .ColIndex("Vat")), 2)
        End If
    '  total_value = Round(.TextMatrix(i, .ColIndex("value")), 2)
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = DcboCreditSide.BoundText
        RsDev("NextAccount_Code").value = .TextMatrix(i, .ColIndex("AccountCode"))
        RsDev("Value").value = total_value
        'RsDev("Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
        RsDev("Credit_Or_Debit").value = 1
        '     rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
        RsDev("RecordDate").value = Me.XPDtbTrans.value
        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
        RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = .TextMatrix(i, .ColIndex("des")) 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
                        
        RsDev("UserID").value = Me.DCboUserName.BoundText
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("notes_all").value = Me.XPTxtID.Text
        RsDev.update
End If
NextRow:
Next i
   If val(TxtVATCustoms.Text) > 0 Then
    
    StrAccount = get_account_code_branch(148, my_branch)
    line_no = line_no + 1
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccount, val(TxtVATCustoms.Text), 1, txt_general_des.Text & "Õ”«» «·ﬁÌ„… «·„÷«›… ··Ã„«—ﬂ ›Ì ”‰œ ’—›  Õ·Ì· „’—Ê›« ", A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
            GoTo ErrTrap
        End If
    End If
End With
Else

        '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = DcboCreditSide.BoundText
        RsDev("Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) + val(TxtVATCustoms.Text) '.TextMatrix(I, .ColIndex("VALUE"))
        RsDev("Credit_Or_Debit").value = 1
        If Posted = 1 Then
        RsDev("Posted").value = 1
        Else
        RsDev("Posted").value = Null
        End If
        RsDev("FlgVat").value = 0
        '     rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
        RsDev("RecordDate").value = Me.XPDtbTrans.value
        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
        RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.Text & bankDes 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("UserID").value = Me.DCboUserName.BoundText
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("notes_all").value = Me.XPTxtID.Text
        RsDev.update

End If
            '       With Fg_Journal
            '    For i = .FixedRows To .Rows - 1
            '        ' line_no = 2
            '        If val(.TextMatrix(i, .ColIndex("Vat"))) <> 0 Then
            '            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            '            OtherInformation.FlgVat = 1
            '            OtherInformation.Vat = 0
              '          OtherInformation.Vatyo = 0
              '          project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
              '          If ModAccounts.AddNewDev(LngDevID, line_no, AccountVATDept, .TextMatrix(i, .ColIndex("Vat")), 0, txt_general_des.Text & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("Vat")), , , , , setfoxy_Line, val(Me.XPTxtID.Text), , , , , , , , val(.TextMatrix(i, .ColIndex("branch_id"))), , , , , , , , , , , , , , project_id, , , , , , , , , , , Posted, , OtherInformation) = False Then
              '              GoTo ErrTrap
              '          End If
              '          line_no = line_no + 1
              '          If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, .TextMatrix(i, .ColIndex("Vat")), 1, txt_general_des.Text & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("Vat")), , , , , setfoxy_Line, val(Me.XPTxtID.Text), , , , , , , , val(.TextMatrix(i, .ColIndex("branch_id"))), , , , , , , , , , , , , , project_id, , , , , , , , , , , Posted, , OtherInformation) = False Then
              '              GoTo ErrTrap
              '          End If
              '          line_no = line_no + 1
              '
              '      End If
              '  Next i
            'End With
            

        'GoTo ll
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        If Me.dcproject.BoundText <> "" Then
            '«·ÿ—› «·„œÌ‰   „’—Ê›«  «·„‘—Ê⁄
            RsNotes.AddNew
            NoteID = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteID").value = CStr(NoteID)
             RsNotes.update
            RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
          
            RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
            RsNotes("Remark").value = txt_general_des.Text & bankDes

            If Me.CboPaymentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPaymentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
            ElseIf Me.CboPaymentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
                            
            End If
               
            ' If txt_ORDER_NO.text <> "" Then
            '       RsNotes("order_no").value = txt_ORDER_NO.text
            '   Else
            '   RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
            '  End If
            
            RsNotes("CusID").value = Null
            RsNotes("NoteType").value = 3
            RsNotes("NoteDate").value = XPDtbTrans.value
           ' RsNotes("NoteDate").value = Format$(Date, "dd-mm-yyyy")
            RsNotes("NoteDateH").value = Me.Txt_DateHigri.value
   
            RsNotes("UserID").value = user_id
            ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
            RsNotes("notes_all").value = Me.XPTxtID.Text
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = month(XPDtbTrans.value)
 '                RsNotes("ManualNO").value = IIf(Trim(TxtManulaNO.text) = "", Null, Trim(TxtManulaNO.text))
RsNotes("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
            RsNotes.update
          
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = dcproject.BoundText
            RsDev("Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) + val(TxtVATCustoms.Text) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.Text & bankDes  ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("FlgVat").value = 0
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.Text
            '                      RsDev("project_id").value = project_id
                        
            RsDev.update
                    
            line_no = line_no + 1
            With Fg_Journal
                For i = .FixedRows To .Rows - 1
                    ' line_no = 2
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                        '////////////////////////////////////////notes
                        If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                            Else
                                MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
                            End If
                        End If
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
                        OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                        OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                        OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                        OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                        OtherInformation.SupplierID = val(.TextMatrix(i, .ColIndex("SupplierID")))
                        OtherInformation.CusVATNO = (.TextMatrix(i, .ColIndex("CusVATNO")))
                        OtherInformation.SupplierName = (.TextMatrix(i, .ColIndex("SupplierName")))
                        OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                        OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))
                       ' OtherInformation.BillNo = (.TextMatrix(i, .ColIndex("BillNo")))
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")), 1, txt_general_des.Text & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , setfoxy_Line, val(Me.XPTxtID.Text), , , , , , , , val(.TextMatrix(i, .ColIndex("branch_id"))), , , , , , , , , , , , , .TextMatrix(i, .ColIndex("BillNo")), project_id, , , , , , , , , , , Posted, , OtherInformation) = False Then
                            GoTo ErrTrap
                        End If
                        line_no = line_no + 1
                    End If
                Next i
            End With
            '''////////VAT
            
              '     With Fg_Journal
              '  For i = .FixedRows To .Rows - 1
              '      ' line_no = 2
              '      If val(.TextMatrix(i, .ColIndex("Vat"))) <> 0 Then
              '          LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
              '          project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
              '           OtherInformation.FlgVat = 1
              '          OtherInformation.Vat = 0
              '          OtherInformation.Vatyo = 0
              '          If ModAccounts.AddNewDev(LngDevID, line_no, AccountVATDept, .TextMatrix(i, .ColIndex("Vat")), 0, txt_general_des.Text & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("Vat")), , , , , setfoxy_Line, val(Me.XPTxtID.Text), , , , , , , , val(.TextMatrix(i, .ColIndex("branch_id"))), , , , , , , , , , , , , , project_id, , , , , , , , , , , Posted, , OtherInformation) = False Then
              '              GoTo ErrTrap
              '          End If
              '          line_no = line_no + 1
              '          If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("Vat")), 1, txt_general_des.Text & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("Vat")), , , , , setfoxy_Line, val(Me.XPTxtID.Text), , , , , , , , val(.TextMatrix(i, .ColIndex("branch_id"))), , , , , , , , , , , , , , project_id, , , , , , , , , , , Posted, , OtherInformation) = False Then
              '              GoTo ErrTrap
              '          End If
              '          line_no = line_no + 1
              '
              '      End If
              '  Next i
            'End With
            ''//////////

            Dim sql As String
            sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=3" & "and NoteSerial1=" & TxtSerial1
            Cn.Execute sql
            sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=3" & "and NoteSerial1=" & TxtSerial1
            Cn.Execute sql
 
        End If

        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        LblDevID.Caption = LngDevID
        lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
    End If

ll:

 '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
    '     If Me.DcCostCenter.BoundText <> "" Then
    save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "”‰œ ’—›", Me.XPDtbTrans.value
    '     End If
    save_cost_center
        
    'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„’«—Ì›
     
    If saveExpensesDetails(0, TxtSerial.Text, TxtSerial1.Text, txt_ORDER_NO.Text, XPDtbTrans.value, val(XPTxtID.Text)) = True Then
    End If
    
    'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
    saveChequeBoxContents1 (val(Me.XPTxtID.Text))
    
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
 sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) + val(TxtVATCustoms.Text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=3" & "and NoteSerial1=" & TxtSerial1
            Cn.Execute sql
    CuurentLogdata
    
    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
        
            End If

            Fg_Journal.Enabled = False

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
            End If
        
            lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)
        
            Fg_Journal.Enabled = False
    End Select
    
   
    TxtModFlg.Text = "R"
fillapprovData
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
            Msg = "cant save " & CHR(13)
            Msg = Msg + "Invalid entry value " & CHR(13)
            Msg = Msg + "Check data and try again"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorr.... Error during saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function saveChequeBoxContents1(NoteID As Double)
 
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    If SystemOptions.banks_Accounts3 = False Then Exit Function
    'rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
          StrSQL = "SELECT     * from dbo.TblChecqueBoxContent1 Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If CboPaymentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.Text
        rs("ChequeNo").value = TxtChequeNumber.Text
        rs("ChequeValue").value = val(XPTxtVal.Text)
        rs("Remarks").value = Me.DcboDebitSide.Text
        rs("Payed").value = 0
        rs("DepitAccount").value = (DcboDebitSide.BoundText)
        rs("notes_all").value = NoteID
        rs.update
    End If

    rs.Close
End Function

Function createDest() As Boolean
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    '„’—Ê›« 
    If CheckAllExpensesDistributed = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·”‰œ ÌÕ ÊÏ ⁄·Ï „’«—Ì› „Ê“—⁄Â «Œ—Ï €Ì— „Ê“⁄Â Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
        Else
            MsgBox "This Voucher Have Distributed and not Distributed Expenses", vbCritical
        End If

        createDest = False
        Exit Function
    End If

    '//////////////////////////////////////Notes////////////////////////////////////
    Dim RsNotes As ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
   ' RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  Dim StrSQL  As String
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
    Dim ExpensesID As Double
    Dim NoteID As String
 
    RsNotes.AddNew
    NoteID = CStr(new_id("Notes", "NoteID", "", True))
    RsNotes("NoteID").value = CStr(NoteID)
     RsNotes.update
     
    RsNotes("Note_Value").value = val(XPTxtVal.Text)
    RsNotes("Remark").value = txt_general_des.Text
    RsNotes("foxy_no").value = val(Text1.Text)
    RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)

    If Me.CboPaymentType.ListIndex = 0 Then
        RsNotes("BoxID").value = val(DcboBox.BoundText)
        RsNotes("BankID").value = Null
        RsNotes("ChqueNum").value = Null
        RsNotes("DueDate").value = Null
        RsNotes("NoteCashingType").value = 0
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 1
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 3
                            
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 2
                        
    ElseIf Me.CboPaymentType.ListIndex = 5 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 5
                        
    End If

    If txt_ORDER_NO.Text <> "" Then
        RsNotes("order_no").value = txt_ORDER_NO.Text
    Else
              
    End If

    RsNotes("CusID").value = Null
    RsNotes("NoteType").value = 3
    RsNotes("NoteDate").value = XPDtbTrans.value
   ' RsNotes("NoteDate").value = Format$(Date, "dd-mm-yyyy")
    RsNotes("NoteDateH").value = Me.Txt_DateHigri.value
   
    RsNotes("UserID").value = user_id
    'RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
    RsNotes("notes_all").value = Me.XPTxtID.Text
    RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
    RsNotes("sanad_year").value = year(XPDtbTrans.value)
    RsNotes("sanad_month").value = month(XPDtbTrans.value)
    RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
    RsNotes("remark").value = txt_general_des.Text
    RsNotes.update
              
    Dim line_no As Integer
    Dim i As Integer
    Dim project_id As Integer
    Dim LngDevID As Long

    With GridEstimatedCost
 
        line_no = 1

        For i = .FixedRows To .Rows - 1
   
            If .TextMatrix(i, .ColIndex("AcountCode")) <> "" Then
                '////////////////////////////////////////notes
   
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                If Destribute = True Then
                    If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AcountCode")), .TextMatrix(i, .ColIndex("Netvalue")), 0, .TextMatrix(i, .ColIndex("Remarks")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.Text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                              
                    End If
                     
                    line_no = line_no + 1

                    If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, .TextMatrix(i, .ColIndex("Netvalue")), 1, .TextMatrix(i, .ColIndex("Remarks")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.Text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                              
                    End If
     
                    line_no = line_no + 1
                End If
        
            End If

        Next i

    End With

    createDest = True
    '
ErrTrap:
End Function

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.Text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.Text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtSerial.Text
        rs("Remark").value = "”‰œ ’—› —ﬁ„ " & TxtSerial1 & "    " & Me.txt_general_des
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 and  kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
        
   ' rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("general_des").value = 1
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = val(.TextMatrix(i, .ColIndex("LineNo1")))
                rs("record_date").value = record_date
                If ChkCCDES.value = vbChecked Then
                rs("description").value = txt_general_des.Text
              Else
              rs("description").value = ""
              End If
                
                rs.update
        
            End If

        Next i

    End With

    rs.Close
End Function

Function calcnets()

    If GridEstimatedCost.Rows > 1 Then
        chkDestribute.value = vbChecked
    Else
        chkDestribute.value = vbUnchecked
    End If

    With Fg_Journal
        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then

        With Me.VSFlexGrid1
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    End If

End Function

Private Sub Undo()
    On Error GoTo ErrTrap
    Dim sql As String
    Dim sgl As String

    Select Case TxtModFlg.Text

        Case "N"
            sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords

            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            sgl = "delete  marakes_taklefa_temp  where ok is null and  kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            rs.Find "NoteID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
Function CheAssetPayd(Optional NoteID As Double = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
CheAssetPayd = False
sql = "select NoteID from notes_all where NoteID=" & NoteID & " and (AssestPayd =1) "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheAssetPayd = True
Else
CheAssetPayd = False
End If
End Function
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If SystemOptions.banks_Accounts3 = True Then
        If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
            Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If
      If CheAssetPayd(val(Me.XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ≈÷«›… ··«’Ê·   "
                    Else
                    Msg = " Can Not Delete this Process"
                    Msg = Msg & CHR(13) & " There is the Process of adding Assest "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
                
    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteserial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        MsgBox "Confirm Delete"
       End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.Text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            '        StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & Val(Me.TXT_A_NoteID)
            '   Cn.Execute StrSQL, , adExecuteNoRecords
    
            StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.Rows = 3
                    Fg_Journal.Enabled = False
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
        Msg = "This is Process UnAvailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error douring delete " & CHR(13)
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Function FillGridWithData()

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub PutData()

    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
    With Fg_Journal

        If Len(TxtDes.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
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
    Dim NO As String
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
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
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
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
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
                    auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & mId(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

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
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

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
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = X ' last_line_id
 
    rs.update
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    Exit Sub
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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
   '         Cmd_Click (6)
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
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If BolRtl = True Then

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Display Help for this Screen" & Wrap & "Shortcut Key F1" & Wrap, BolRtl
        End With

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

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtSerial1.Text) <> "" Then
        oldTxtSerial1.Text = TxtSerial1.Text
    End If

    TxtSerial.Text = ""
    TxtSerial1.Text = ""
    Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
End Sub

Private Sub Txt_DateHigri_LostFocus()
    XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
 
End Sub

Private Sub XPTxtVal_Change()
    XPTxtValView.Text = Format(val(XPTxtVal.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 1)

    End If
    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
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
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(28).Caption = "Customs Value"
 CmdAttach.Caption = "Attachments"
    lbl(14).Caption = "Project#"
    Label1.Caption = "Voucher #"
    Me.C1Tab1.TabCaption(0) = "Expenses"
    Me.C1Tab1.TabCaption(1) = "Distributions"
    Me.C1Tab1.TabCaption(2) = "Internal Rules"
lbl(22).Caption = "According to"
ChkCCDES.Caption = "Add Des To CC Des"
lbl(53).Caption = "Manual No"

    With Me.CBoBasedON
        .Clear
        .AddItem "Without"
        .AddItem "Purchase Invoices"
        .AddItem "Performa Invoices"
        .AddItem "Production Order"
    
    End With
Label10.Caption = "Based Request"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(15).Caption = "Payment Method"
    lbl(16).Caption = "Box Name"
    lbl(20).Caption = "General Des"
    lbl(21).Caption = "Order No:"

    lbl(26).Caption = "Account No:"

    Label8.Caption = "General C. C."

    With Me.CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Bank Transfer"
        .AddItem "P Cheque"
        .AddItem "Account"
        .AddItem "Bank Order"
    
    End With

    CmdRemove.Caption = "Delete Row"
    CmdRemoveAll.Caption = "Delete All"
    Me.Caption = "Payments Voucher"
    Me.Ele(0).Caption = "Payments Voucher"
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = "Operation ID"
    Me.lbl(1).Caption = "Date"
    Me.lbl(3).Caption = "Expenses Type"
    Me.lbl(2).Caption = "Total"
    Me.lbl(0).Caption = "Based On"
    Me.lbl(22).Caption = "Based On"
    Label3.Caption = "Branch"

    Me.lbl(5).Caption = "TO"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."
    Fra.Caption = "GL"
    lbl(11).Caption = "GL#"
    lbl(13).Caption = "interval"
    lbl(9).Caption = "Depit"
    lbl(10).Caption = "Credit"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque#"
    lbl(19).Caption = "Due Date"

    Me.Cmd(0).Caption = "&New"
    Me.Cmd(1).Caption = "&Edit"
    Me.Cmd(2).Caption = "&Save"
    Me.Cmd(3).Caption = "&Undo"
    Me.Cmd(4).Caption = "&Delete"
    Me.Cmd(5).Caption = "Sear&ch"
    Me.Cmd(6).Caption = "E&xit"
    Me.Cmd(7).Caption = "&Table View"
    Cmd(8).Caption = "Print"
    Cmd(9).Caption = "Cheque Print"
    Cmd(10).Caption = "GL Print "

    Me.CmdHelp.Caption = "&Help"

 With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Account Project"
        .TextMatrix(0, .ColIndex("value")) = "value"
        
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("pand")) = "Pand"
        .TextMatrix(0, .ColIndex("oper")) = "Process"
        .TextMatrix(0, .ColIndex("Fixes")) = "Machine"
        .TextMatrix(0, .ColIndex("dept")) = "Department"
        .TextMatrix(0, .ColIndex("Des")) = "Description"
       ' .TextMatrix(0, .ColIndex("order_no")) = "order no"
       ' .TextMatrix(0, .ColIndex("des")) = "description"
       ' .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
       ' .TextMatrix(0, .ColIndex("order_no")) = "order no"

    End With
    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("PrjectCode")) = "Prject Code"
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Expenses Code"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("CarName")) = "CarName "
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("pand")) = "Pand"
        .TextMatrix(0, .ColIndex("oper")) = "Process"
        .TextMatrix(0, .ColIndex("Fixes")) = "Machine"
        .TextMatrix(0, .ColIndex("dept")) = "Department"
        .TextMatrix(0, .ColIndex("des")) = "Description"
        .TextMatrix(0, .ColIndex("order_no")) = "Order no"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
        .TextMatrix(0, .ColIndex("order_no")) = "Order no"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
        .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        .TextMatrix(0, .ColIndex("PriceTotal")) = "Price Total"
        .TextMatrix(0, .ColIndex("SupplierName")) = "Cash Supplier"
        .TextMatrix(0, .ColIndex("CusVATNO")) = "VAT NO."
        .TextMatrix(0, .ColIndex("Supplier")) = "Supplier Name"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
    End With

    With Me.GridEstimatedCost
        .TextMatrix(0, .ColIndex("Ser")) = "Index"
        .TextMatrix(0, .ColIndex("AcountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("BranchName")) = " Branch Name "

        .TextMatrix(0, .ColIndex("value")) = "Total Value"
        .TextMatrix(0, .ColIndex("Percentage")) = "Percentage"
        .TextMatrix(0, .ColIndex("Netvalue")) = "Distr Value"
        .TextMatrix(0, .ColIndex("REMARKS")) = "REMARKS "

    End With
    
    Accredit.Caption = "Send For Approval"
    Me.C1Tab1.TabCaption(3) = "Approval Status"
    Label1100.Caption = "Approval Requested By"
    Label11.Caption = "Approval Requested By"
    
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With

End Sub
