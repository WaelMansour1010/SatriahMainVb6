VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBuyGoodsInst 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmBuyGoodsInst.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBuyGoodsInst.frx":6852
      Left            =   15480
      List            =   "FrmBuyGoodsInst.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   82
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   0
      Width           =   14505
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   77
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyGoodsInst.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   78
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyGoodsInst.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   79
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyGoodsInst.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   80
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuyGoodsInst.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ·» ‘—«¡ ”·⁄… »«· ﬁ”Ìÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmBuyGoodsInst.frx":76E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Caption         =   "«·„—›ﬁ« "
      Enabled         =   0   'False
      Height          =   7335
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   720
      Width           =   14235
      Begin VB.Frame d 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   73
         Top             =   120
         Width           =   14175
         Begin VB.TextBox TxtTypeRequest 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtNoteIDx 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8640
            TabIndex        =   1
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94830593
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmBuyGoodsInst.frx":8AE8
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
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
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   255
            Left            =   7200
            TabIndex        =   2
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   450
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ÿ·»"
            Height          =   285
            Index           =   1
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   0
            Left            =   3270
            TabIndex        =   104
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   285
            Index           =   4
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10410
            TabIndex        =   74
            Top             =   255
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2505
         Index           =   0
         Left            =   0
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   840
         Width           =   14175
         _cx             =   25003
         _cy             =   4419
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
         Caption         =   " "
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
         Begin VB.TextBox TxtCus_Name 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   240
            Width           =   4140
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   2040
            Width           =   9975
         End
         Begin VB.TextBox TxtTotal_liab 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11355
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox TxtStreet 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox TxtDistrict 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox TxtNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox TxtHome_Tel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11355
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox TxtIBN 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox TxtBankName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox TxtExt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox TxtTelephone 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11355
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox TxtPlaceID 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtSalary_Acc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TxtSalary 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11355
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox TxtCompany 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox TxtCity 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox TxtCust_Mobile 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox TxtItemRequest 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   11760
            TabIndex        =   5
            Top             =   -240
            Visible         =   0   'False
            Width           =   810
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   8400
            TabIndex        =   6
            Top             =   -240
            Visible         =   0   'False
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker BrithDate 
            Height          =   315
            Left            =   11355
            TabIndex        =   9
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94830593
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal BrithDateH 
            Height          =   255
            Left            =   8640
            TabIndex        =   10
            Top             =   600
            Width           =   1455
            _extentx        =   2566
            _extenty        =   450
         End
         Begin Dynamic_Byte.NourHijriCal ExpDate 
            Height          =   255
            Left            =   1800
            TabIndex        =   16
            Top             =   960
            Width           =   1335
            _extentx        =   2355
            _extenty        =   450
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   26
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«· “«„« "
            Height          =   285
            Index           =   25
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   2040
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘«—⁄"
            Height          =   285
            Index           =   24
            Left            =   3120
            TabIndex        =   124
            Top             =   1680
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÕÌ"
            Height          =   285
            Index           =   23
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   1680
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   285
            Index           =   22
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Â« › «·„‰“·"
            Height          =   285
            Index           =   21
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   1680
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«» «·»‰ﬂÌ"
            Height          =   285
            Index           =   20
            Left            =   3120
            TabIndex        =   120
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»‰ﬂ «·Õ«·Ì"
            Height          =   285
            Index           =   19
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕÊÌ·…"
            Height          =   285
            Index           =   18
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Â« › «·⁄„·"
            Height          =   285
            Index           =   17
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«‰ Â«¡/„ﬂ«‰ «·«’œ«—"
            Height          =   285
            Index           =   16
            Left            =   3120
            TabIndex        =   116
            Top             =   960
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„Ì·«œ"
            Height          =   285
            Index           =   15
            Left            =   10080
            TabIndex        =   115
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—« » ›Ì «·ﬂ‘›"
            Height          =   285
            Index           =   14
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—« »"
            Height          =   285
            Index           =   13
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   960
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÂÊÌ…"
            Height          =   285
            Index           =   11
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÂ… «·⁄„·"
            Height          =   285
            Index           =   10
            Left            =   3120
            TabIndex        =   111
            Top             =   600
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ…"
            Height          =   285
            Index           =   6
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÊ«· «·⁄„Ì·"
            Height          =   285
            Index           =   9
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”·⁄… «·„ÿ·Ê»Â"
            Height          =   285
            Index           =   5
            Left            =   3120
            TabIndex        =   108
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„Ì·«œ"
            Height          =   285
            Index           =   3
            Left            =   12570
            TabIndex        =   107
            Top             =   615
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   7
            Left            =   12570
            TabIndex        =   106
            Top             =   300
            Width           =   1965
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1065
         Index           =   1
         Left            =   0
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   4680
         Width           =   14175
         _cx             =   25003
         _cy             =   1879
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
         Caption         =   " «·ﬂ›Ì· «·€«—„"
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
         Begin VB.TextBox TxtGurSalary 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TxtGurCompany 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   9240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox TxtGurType_liab 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1080
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox TxtGurTotal_liab 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtGurName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   3735
         End
         Begin MSComCtl2.DTPicker GurBrithDate 
            Height          =   315
            Left            =   3240
            TabIndex        =   52
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94830593
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal GurBrithDateH 
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1455
            _extentx        =   2566
            _extenty        =   450
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   145
            Top             =   240
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«· «ﬂÌœ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«· “«„« "
            Height          =   285
            Index           =   27
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—« » «·«Ã„«·Ì"
            Height          =   285
            Index           =   32
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„Ì·«œ"
            Height          =   285
            Index           =   31
            Left            =   1320
            TabIndex        =   133
            Top             =   600
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„Ì·«œ"
            Height          =   285
            Index           =   30
            Left            =   4530
            TabIndex        =   132
            Top             =   615
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÂ… «·⁄„·"
            Height          =   285
            Index           =   29
            Left            =   13080
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«· “«„« "
            Height          =   285
            Index           =   28
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·ﬂ›Ì·"
            Height          =   285
            Index           =   12
            Left            =   13080
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   240
            Width           =   1035
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1305
         Index           =   2
         Left            =   0
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   3360
         Width           =   14175
         _cx             =   25003
         _cy             =   2302
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
         Caption         =   " «·„—›ﬁ« "
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
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   28
            Top             =   240
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "»ÿ«ﬁ… «·«ÕÊ«·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   1
            Left            =   12360
            TabIndex        =   34
            Top             =   600
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ—ÊﬂÌ «·⁄„·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   2
            Left            =   12000
            TabIndex        =   40
            Top             =   960
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " ”·Ì„ «·»÷«⁄…"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   3
            Left            =   10080
            TabIndex        =   29
            Top             =   240
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ÿ·» ‘—«¡"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   4
            Left            =   10320
            TabIndex        =   35
            Top             =   600
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«„— „” œÌ„"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   5
            Left            =   10320
            TabIndex        =   41
            Top             =   960
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "⁄ﬁœ «ÌÃ«—"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   6
            Left            =   7920
            TabIndex        =   30
            Top             =   240
            Width           =   2175
            _Version        =   786432
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Œÿ«»  ⁄—Ì› »«·—« »"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   7
            Left            =   7920
            TabIndex        =   36
            Top             =   600
            Width           =   2175
            _Version        =   786432
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "⁄ﬁœ «· ﬁ”Ìÿ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   8
            Left            =   8400
            TabIndex        =   42
            Top             =   960
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Õ”„ „»«‘—"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   9
            Left            =   5760
            TabIndex        =   31
            Top             =   240
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ‘› Õ”«» «·»‰ﬂ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   10
            Left            =   6000
            TabIndex        =   37
            Top             =   600
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "”‰œ«  ·√„—"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   11
            Left            =   6000
            TabIndex        =   43
            Top             =   960
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ›«·… €—«„ Ê«œ«¡"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   32
            Top             =   240
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "›« Ê—… ﬂÂ—»«¡"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   13
            Left            =   3840
            TabIndex        =   38
            Top             =   600
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«ﬁ—«— Ê ⁄Âœ «·„Õ«„«…"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   14
            Left            =   3840
            TabIndex        =   44
            Top             =   960
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " ÊﬂÌ· »Ì⁄"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   15
            Left            =   1800
            TabIndex        =   33
            Top             =   240
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ—ÊﬂÌ «·„‰“·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   16
            Left            =   2040
            TabIndex        =   39
            Top             =   600
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ÃœÊ· «·«ﬁ”«ÿ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   17
            Left            =   2280
            TabIndex        =   45
            Top             =   960
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "”‰œ ﬁ»÷"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "”‰œ ’—›"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CHATT 
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   146
            Top             =   240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "’Ê—… ’ﬂ «·„‰“·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1545
         Index           =   3
         Left            =   0
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   5640
         Width           =   14175
         _cx             =   25003
         _cy             =   2725
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
         Caption         =   " "
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
         Begin VB.TextBox TxtNameAdmin_Mobile 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   960
            Width           =   7575
         End
         Begin VB.TextBox TxtNameOffice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   9600
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox TxtGurExt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox TxtGurTele 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox TxtDirectManager 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3960
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox TxtAdress2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   9600
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox TxtRelation2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox TxtAdress1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3960
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox TxtRelation1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«—ﬁ«„ «·ÂÊ« › Ê«”„ «·„”ƒÊ·"
            Height          =   285
            Index           =   40
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„«¡ „ﬂ« » «· ﬁ”Ìÿ"
            Height          =   285
            Index           =   38
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Â« ›/ ÕÊÌ·…"
            Height          =   285
            Index           =   37
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œÌ— «·„»«‘—"
            Height          =   285
            Index           =   36
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÃÊ«· Ê«·⁄‰Ê«‰ ﬂ«„·"
            Height          =   285
            Index           =   35
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «Õœ «·«ﬁ«—»"
            Height          =   285
            Index           =   34
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÃÊ«· Ê«·⁄‰Ê«‰ ﬂ«„·"
            Height          =   285
            Index           =   33
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «Õœ «·«ﬁ«—»"
            Height          =   285
            Index           =   39
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   84
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   85
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1905
      Left            =   0
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   7320
      Width           =   14235
      _cx             =   25109
      _cy             =   3360
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   88
         Top             =   600
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   87
         Top             =   1200
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   64
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":8AFD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   66
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":F35F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   65
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":F6F9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   67
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":15F5B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   68
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":162F5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   69
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":1688F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   4320
            TabIndex        =   100
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":16C29
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   3000
            TabIndex        =   101
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":1D48B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   330
            Left            =   1440
            TabIndex        =   148
            Top             =   240
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBuyGoodsInst.frx":1D825
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
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10560
         TabIndex        =   93
         Top             =   840
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   -720
         Width           =   13965
         _cx             =   24633
         _cy             =   1005
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
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13320
         TabIndex        =   94
         Top             =   840
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":24087
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":24421
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":247BB
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":24B55
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":24EEF
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":25289
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":25623
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuyGoodsInst.frx":25BBD
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   95
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmBuyGoodsInst.frx":25F57
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmBuyGoodsInst.frx":2C7B9
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmBuyGoodsInst.frx":3301B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmBuyGoodsInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double

Private Sub BrithDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         BrithDateH.value = ToHijriDate(BrithDate.value)
 End If
End Sub

Private Sub BrithDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
BrithDate.value = ToGregorianDate(BrithDateH.value)
End If
End Sub


Private Sub CmdAttach_Click()
    On Error Resume Next
ShowAttachments TxtSerial1, "0703201701"

End Sub

Private Sub GurBrithDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         GurBrithDateH.value = ToHijriDate(GurBrithDate.value)
 End If
End Sub

Private Sub GurBrithDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
GurBrithDate.value = ToGregorianDate(GurBrithDateH.value)
End If
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub ISButton8_Click()
FrmClearanceCerificateSearch.Ind = 1
            Load FrmClearanceCerificateSearch
            FrmClearanceCerificateSearch.show vbModal
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
XPDtbTrans.value = ToGregorianDate(RecordDateH.value)
End If
End Sub

'Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
'  Dim CUSTID As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
'        DBCboClientName.BoundText = CUSTID
'    End If

'End Sub
'Private Sub DBCboClientName_Change()
'If val(DBCboClientName.BoundText) <> 0 Then
'    Dim FullCode As String
'    Dim ExpairDate As String
'    Dim company As String
'    Dim salary As Double
'    Dim JobTel As String
'    Dim mobile As String
'    Dim jobConvert As String
'    Dim hay As String
'    Dim City As String
'    Dim HomeTel As String
'    Dim cusid As Double
'    Dim Address As String
'    GetCustomersDetail val(DBCboClientName.BoundText), , FullCode
'    TxtSearchCode.text = FullCode
'    If Me.TxtModFlg.text <> "R" Then
'    GetCustomerAllData val(DBCboClientName.BoundText), , ExpairDate, company, , salary, , JobTel, jobConvert, HomeTel, mobile, , , City, hay, , , cusid, Address
'    ExpDate.value = ExpairDate
'    TxtCompany.text = company
'    TxtSalary.text = salary
'    TxtTelephone.text = JobTel
'    TxtExt.text = jobConvert
'    TxtHome_Tel.text = HomeTel
'    TxtCust_Mobile.text = mobile
'    TxtCity.text = City
'    TxtDistrict.text = hay
'    TxtCusID.text = cusid
'    TxtStreet.text = Address
'    End If
'    End If
'End Sub

'Private Sub DBCboClientName_Click(Area As Integer)
'    DBCboClientName_Change
'End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblInstallmentsReq order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
   
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    
    Dcombos.GetCustomersSuppliers 1, DBCboClientName
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
    FiLLTXT
   Me.Refresh
   
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
     'On Error GoTo ErrTrap
    Dim Sql As String
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecordDateH").value = RecordDateH.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("TypeRequest").value = TxtTypeRequest.Text
   ' RsSavRec.Fields("Cus_ID").value = val(Me.DBCboClientName.BoundText)
    RsSavRec.Fields("Cust_Mobile").value = TxtCust_Mobile.Text
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("ItemRequest").value = TxtItemRequest.Text
    RsSavRec.Fields("BrithDateH").value = BrithDateH.value
    RsSavRec.Fields("BrithDate").value = BrithDate.value
    RsSavRec.Fields("Company").value = TxtCompany.Text
    RsSavRec.Fields("City").value = TxtCity.Text
    RsSavRec.Fields("CusID").value = TxtCusID.Text
    RsSavRec.Fields("ExpDate").value = ExpDate.value
    RsSavRec.Fields("PlaceID").value = TxtPlaceID.Text
    RsSavRec.Fields("Salary").value = val(TxtSalary.Text)
    RsSavRec.Fields("Salary_Acc").value = val(TxtSalary_Acc.Text)
    RsSavRec.Fields("BankName").value = TxtBankName.Text
    RsSavRec.Fields("Telephone").value = TxtTelephone.Text
    RsSavRec.Fields("Ext").value = TxtExt.Text
    RsSavRec.Fields("Home_Tel").value = TxtHome_Tel.Text
    RsSavRec.Fields("Street").value = TxtStreet.Text
    RsSavRec.Fields("District").value = TxtDistrict.Text
    RsSavRec.Fields("No").value = TxtNo.Text
    RsSavRec.Fields("Cus_Name").value = TxtCus_Name.Text
    
    RsSavRec.Fields("IBN").value = TxtIBN.Text
    RsSavRec.Fields("Total_liab").value = val(TxtTotal_liab.Text)
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("GurName").value = TxtGurName.Text
    RsSavRec.Fields("GurBrithDate").value = GurBrithDate.value
    RsSavRec.Fields("GurBrithDateH").value = GurBrithDateH.value
    RsSavRec.Fields("GurCompany").value = TxtGurCompany.Text
    RsSavRec.Fields("GurSalary").value = val(TxtGurSalary.Text)
    RsSavRec.Fields("GurTotal_liab").value = val(TxtGurTotal_liab.Text)
    RsSavRec.Fields("Relation1").value = TxtRelation1.Text
    RsSavRec.Fields("Relation2").value = TxtRelation2.Text
    RsSavRec.Fields("GurType_liab").value = TxtGurType_liab.Text
    RsSavRec.Fields("Adress1").value = TxtAdress1.Text
    RsSavRec.Fields("Adress2").value = TxtAdress2.Text
    RsSavRec.Fields("DirectManager").value = TxtDirectManager.Text
    RsSavRec.Fields("GurTele").value = TxtGurTele.Text
    RsSavRec.Fields("GurExt").value = TxtGurExt.Text
    RsSavRec.Fields("NameOffice").value = TxtNameOffice.Text
    RsSavRec.Fields("NameAdmin_Mobile").value = TxtNameAdmin_Mobile.Text
    If Me.CHATT(0).value = vbChecked Then
    RsSavRec.Fields("Attch0").value = 1
    Else
    RsSavRec.Fields("Attch0").value = 0
    End If
    If Me.CHATT(1).value = vbChecked Then
    RsSavRec.Fields("Attch1").value = 1
    Else
    RsSavRec.Fields("Attch1").value = 0
    End If
     If Me.CHATT(2).value = vbChecked Then
    RsSavRec.Fields("Attch2").value = 1
    Else
    RsSavRec.Fields("Attch2").value = 0
    End If
    If Me.CHATT(3).value = vbChecked Then
    RsSavRec.Fields("Attch3").value = 1
    Else
    RsSavRec.Fields("Attch3").value = 0
    End If
    If Me.CHATT(4).value = vbChecked Then
    RsSavRec.Fields("Attch4").value = 1
    Else
    RsSavRec.Fields("Attch4").value = 0
    End If
     If Me.CHATT(5).value = vbChecked Then
    RsSavRec.Fields("Attch5").value = 1
    Else
    RsSavRec.Fields("Attch5").value = 0
    End If
    If Me.CHATT(6).value = vbChecked Then
    RsSavRec.Fields("Attch6").value = 1
    Else
    RsSavRec.Fields("Attch6").value = 0
    End If
    If Me.CHATT(7).value = vbChecked Then
    RsSavRec.Fields("Attch7").value = 1
    Else
    RsSavRec.Fields("Attch7").value = 0
    End If
    If Me.CHATT(8).value = vbChecked Then
    RsSavRec.Fields("Attch8").value = 1
    Else
    RsSavRec.Fields("Attch8").value = 0
    End If
    If Me.CHATT(9).value = vbChecked Then
    RsSavRec.Fields("Attch9").value = 1
    Else
    RsSavRec.Fields("Attch9").value = 0
    End If
    If Me.CHATT(10).value = vbChecked Then
    RsSavRec.Fields("Attch10").value = 1
    Else
    RsSavRec.Fields("Attch10").value = 0
    End If
    If Me.CHATT(11).value = vbChecked Then
    RsSavRec.Fields("Attch11").value = 1
    Else
    RsSavRec.Fields("Attch11").value = 0
    End If
    If Me.CHATT(12).value = vbChecked Then
    RsSavRec.Fields("Attch12").value = 1
    Else
    RsSavRec.Fields("Attch12").value = 0
    End If
        If Me.CHATT(13).value = vbChecked Then
    RsSavRec.Fields("Attch13").value = 1
    Else
    RsSavRec.Fields("Attch13").value = 0
    End If
        If Me.CHATT(14).value = vbChecked Then
    RsSavRec.Fields("Attch14").value = 1
    Else
    RsSavRec.Fields("Attch14").value = 0
    End If
        If Me.CHATT(15).value = vbChecked Then
    RsSavRec.Fields("Attch15").value = 1
    Else
    RsSavRec.Fields("Attch15").value = 0
    End If
    If Me.CHATT(16).value = vbChecked Then
    RsSavRec.Fields("Attch16").value = 1
    Else
    RsSavRec.Fields("Attch16").value = 0
    End If
    If Me.CHATT(17).value = vbChecked Then
    RsSavRec.Fields("Attch17").value = 1
    Else
    RsSavRec.Fields("Attch17").value = 0
    End If
    If Me.CHATT(18).value = vbChecked Then
    RsSavRec.Fields("Attch18").value = 1
    Else
    RsSavRec.Fields("Attch18").value = 0
    End If
    If Me.CHATT(19).value = vbChecked Then
    RsSavRec.Fields("Accept").value = 1
    Else
    RsSavRec.Fields("Accept").value = 0
    End If
     If Me.CHATT(20).value = vbChecked Then
    RsSavRec.Fields("Attch19").value = 1
    Else
    RsSavRec.Fields("Attch19").value = 0
    End If
    
    RsSavRec.update

      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This Recpord alredy saved... " & Chr(13)
                Msg = Msg + " you want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
             
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "  Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
            
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
       
       RsSavRec.Resync adAffectCurrent
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim I As Integer
    Dim ContactTime  As Date
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    
     Me.TxtCus_Name.Text = IIf(IsNull(RsSavRec.Fields("Cus_Name").value), "", RsSavRec.Fields("Cus_Name").value)
    Me.TxtTypeRequest.Text = IIf(IsNull(RsSavRec.Fields("TypeRequest").value), "", RsSavRec.Fields("TypeRequest").value)
    Me.TxtCust_Mobile.Text = IIf(IsNull(RsSavRec.Fields("Cust_Mobile").value), "", RsSavRec.Fields("Cust_Mobile").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.TxtItemRequest.Text = IIf(IsNull(RsSavRec.Fields("ItemRequest").value), "", RsSavRec.Fields("ItemRequest").value)
   ' Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("Cus_ID").value), "", RsSavRec.Fields("Cus_ID").value)
    BrithDateH.value = IIf(IsNull(RsSavRec.Fields("BrithDateH").value), ToHijriDate(Date), RsSavRec.Fields("BrithDateH").value)
    Me.TxtCompany.Text = IIf(IsNull(RsSavRec.Fields("Company").value), "", RsSavRec.Fields("Company").value)
    BrithDate.value = IIf(IsNull(RsSavRec.Fields("BrithDate").value), "", RsSavRec.Fields("BrithDate").value)
    Me.TxtCusID.Text = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    Me.TxtCity.Text = IIf(IsNull(RsSavRec.Fields("City").value), "", RsSavRec.Fields("City").value)
    Me.ExpDate.value = IIf(IsNull(RsSavRec.Fields("ExpDate").value), ToHijriDate(Date), RsSavRec.Fields("ExpDate").value)
    Me.TxtPlaceID.Text = IIf(IsNull(RsSavRec.Fields("PlaceID").value), "", RsSavRec.Fields("PlaceID").value)
    Me.TxtSalary.Text = IIf(IsNull(RsSavRec.Fields("Salary").value), 0, RsSavRec.Fields("Salary").value)
    Me.TxtBankName.Text = IIf(IsNull(RsSavRec.Fields("BankName").value), "", RsSavRec.Fields("BankName").value)
    Me.TxtSalary_Acc.Text = IIf(IsNull(RsSavRec.Fields("Salary_Acc").value), 0, RsSavRec.Fields("Salary_Acc").value)
    Me.TxtTelephone.Text = IIf(IsNull(RsSavRec.Fields("Telephone").value), "", RsSavRec.Fields("Telephone").value)
    Me.TxtExt.Text = IIf(IsNull(RsSavRec.Fields("Ext").value), "", RsSavRec.Fields("Ext").value)
    Me.TxtHome_Tel.Text = IIf(IsNull(RsSavRec.Fields("Home_Tel").value), "", RsSavRec.Fields("Home_Tel").value)
    Me.TxtStreet.Text = IIf(IsNull(RsSavRec.Fields("Street").value), "", RsSavRec.Fields("Street").value)
    Me.TxtDistrict.Text = IIf(IsNull(RsSavRec.Fields("District").value), "", RsSavRec.Fields("District").value)
    Me.TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    Me.TxtNo.Text = IIf(IsNull(RsSavRec.Fields("No").value), "", RsSavRec.Fields("No").value)
    Me.TxtIBN.Text = IIf(IsNull(RsSavRec.Fields("IBN").value), "", RsSavRec.Fields("IBN").value)
    Me.TxtTotal_liab.Text = IIf(IsNull(RsSavRec.Fields("Total_liab").value), 0, RsSavRec.Fields("Total_liab").value)
    Me.TxtGurName.Text = IIf(IsNull(RsSavRec.Fields("GurName").value), "", RsSavRec.Fields("GurName").value)
    Me.GurBrithDate.value = IIf(IsNull(RsSavRec.Fields("GurBrithDate").value), Date, RsSavRec.Fields("GurBrithDate").value)
    Me.GurBrithDateH.value = IIf(IsNull(RsSavRec.Fields("GurBrithDateH").value), ToHijriDate(Date), RsSavRec.Fields("GurBrithDateH").value)
    Me.TxtGurCompany.Text = IIf(IsNull(RsSavRec.Fields("GurCompany").value), "", RsSavRec.Fields("GurCompany").value)
    Me.TxtGurSalary.Text = IIf(IsNull(RsSavRec.Fields("GurSalary").value), 0, RsSavRec.Fields("GurSalary").value)
    Me.TxtGurTotal_liab.Text = IIf(IsNull(RsSavRec.Fields("GurTotal_liab").value), 0, RsSavRec.Fields("GurTotal_liab").value)
    Me.TxtRelation1.Text = IIf(IsNull(RsSavRec.Fields("Relation1").value), "", RsSavRec.Fields("Relation1").value)
    Me.TxtRelation2.Text = IIf(IsNull(RsSavRec.Fields("Relation2").value), "", RsSavRec.Fields("Relation2").value)
    Me.TxtAdress1.Text = IIf(IsNull(RsSavRec.Fields("Adress1").value), "", RsSavRec.Fields("Adress1").value)
    Me.TxtAdress2.Text = IIf(IsNull(RsSavRec.Fields("Adress2").value), "", RsSavRec.Fields("Adress2").value)
    Me.TxtGurType_liab.Text = IIf(IsNull(RsSavRec.Fields("GurType_liab").value), "", RsSavRec.Fields("GurType_liab").value)
    Me.TxtDirectManager.Text = IIf(IsNull(RsSavRec.Fields("DirectManager").value), "", RsSavRec.Fields("DirectManager").value)
    Me.TxtGurTele.Text = IIf(IsNull(RsSavRec.Fields("GurTele").value), "", RsSavRec.Fields("GurTele").value)
    Me.TxtGurExt.Text = IIf(IsNull(RsSavRec.Fields("GurExt").value), "", RsSavRec.Fields("GurExt").value)
    Me.TxtNameOffice.Text = IIf(IsNull(RsSavRec.Fields("NameOffice").value), "", RsSavRec.Fields("NameOffice").value)
    Me.TxtNameAdmin_Mobile.Text = IIf(IsNull(RsSavRec.Fields("NameAdmin_Mobile").value), "", RsSavRec.Fields("NameAdmin_Mobile").value)
    If RsSavRec.Fields("Attch0").value = True Then
    Me.CHATT(0).value = vbChecked
    Else
    Me.CHATT(0).value = vbUnchecked
    End If
    
    If RsSavRec.Fields("Attch1").value = True Then
    Me.CHATT(1).value = vbChecked
    Else
    Me.CHATT(1).value = vbUnchecked
    End If
        If RsSavRec.Fields("Attch2").value = True Then
    Me.CHATT(2).value = vbChecked
    Else
    Me.CHATT(2).value = vbUnchecked
    End If
        If RsSavRec.Fields("Attch3").value = True Then
    Me.CHATT(3).value = vbChecked
    Else
    Me.CHATT(3).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch4").value = True Then
    Me.CHATT(4).value = vbChecked
    Else
    Me.CHATT(4).value = vbUnchecked
    End If
        If RsSavRec.Fields("Attch5").value = True Then
    Me.CHATT(5).value = vbChecked
    Else
    Me.CHATT(5).value = vbUnchecked
    End If
        If RsSavRec.Fields("Attch6").value = True Then
    Me.CHATT(6).value = vbChecked
    Else
    Me.CHATT(6).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch7").value = True Then
    Me.CHATT(7).value = vbChecked
    Else
    Me.CHATT(7).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch8").value = True Then
    Me.CHATT(8).value = vbChecked
    Else
    Me.CHATT(8).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch9").value = True Then
    Me.CHATT(9).value = vbChecked
    Else
    Me.CHATT(9).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch10").value = True Then
    Me.CHATT(10).value = vbChecked
    Else
    Me.CHATT(10).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch11").value = True Then
    Me.CHATT(11).value = vbChecked
    Else
    Me.CHATT(11).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch12").value = True Then
    Me.CHATT(12).value = vbChecked
    Else
    Me.CHATT(12).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch13").value = True Then
    Me.CHATT(13).value = vbChecked
    Else
    Me.CHATT(13).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch14").value = True Then
    Me.CHATT(14).value = vbChecked
    Else
    Me.CHATT(14).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch15").value = True Then
    Me.CHATT(15).value = vbChecked
    Else
    Me.CHATT(15).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch16").value = True Then
    Me.CHATT(16).value = vbChecked
    Else
    Me.CHATT(16).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch17").value = True Then
    Me.CHATT(17).value = vbChecked
    Else
    Me.CHATT(17).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch18").value = True Then
    Me.CHATT(18).value = vbChecked
    Else
    Me.CHATT(18).value = vbUnchecked
    End If
    If RsSavRec.Fields("Accept").value = True Then
    Me.CHATT(19).value = vbChecked
    Else
    Me.CHATT(19).value = vbUnchecked
    End If
    If RsSavRec.Fields("Attch19").value = True Then
    Me.CHATT(20).value = vbChecked
    Else
    Me.CHATT(20).value = vbUnchecked
    End If
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount

ErrTrap:
 
End Sub



' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
 
      If dcBranch.Text = "" Or val(dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             End If
             dcBranch.SetFocus
            Exit Sub
     End If
     ' If Me.TXtMangerName.text = "" Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·„œÌ— «·⁄«„ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     '             Else
     '       MsgBox "Please Eneter Manager Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '        End If
     '       TXtMangerName.SetFocus
     '       Exit Sub
     'End If
     
    If (Me.TxtCus_Name.Text) = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...  Ì—ÃÏ ≈œŒ«·  «”„  «·⁄„Ì·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Enter Customer ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
        TxtCus_Name.SetFocus
         Exit Sub
     End If
       If TxtTypeRequest.Text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...  Ì—ÃÏ ≈œŒ«· ‰Ê⁄ «·ÿ·» ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Eneter Type Request ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
         TxtTypeRequest.SetFocus
         Exit Sub
     End If
  


            '+++++++++++++++++++++++++++++++++++++++++++++++
    ' For Each CtrlTxt In Me.Controls
    '    If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
    '        If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
    '            MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
    '            CtrlTxt.SetFocus
    '            Exit Sub
    '        End If
    '    End If
    'Next
    '------------------------------ check if Empcode exist ----------------------
'   StrVacName = IsRecExist("TblEmploymentModel", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtSerial1.text) & "'")
  ' If StrVacName <> "" Then
 '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
  '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
  '    TxtVacName.SetFocus
 '     Exit Sub
'   End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry...Error douring saving ", vbOKOnly + vbMsgBoxRight, App.title
  End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblInstallmentsReq", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub




Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecID, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub

 

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim I As Integer
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If
    Dim x As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If x = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Delete Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
          
              
     End If
                            '------------------------------ Move Next ---------------------------.
                            Unche
                            
           LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
                 
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Can Not Delete this record .it's related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecID As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        BtnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
       
        
        
        
    ElseIf TxtModFlg.Text = "R" Then
       
        
        
        btnModify.Enabled = False
        BtnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            BtnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
     
      
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        BtnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
             Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
        
           End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
        
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
             Msg = "Sorry you can not edit this record now" & Chr(13)
            Msg = Msg & "Where it was being edited by another user on the network " & Chr(13)
            
        
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
  Unche
    
 '  Rd(0).value = True
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    dcBranch.SetFocus
          XPDtbTrans.value = Date

ErrTrap:
End Sub
Sub Unche()
  CHATT(0).value = vbUnchecked
    CHATT(1).value = vbUnchecked
    CHATT(3).value = vbUnchecked
    CHATT(4).value = vbUnchecked
    CHATT(5).value = vbUnchecked
    CHATT(6).value = vbUnchecked
    CHATT(7).value = vbUnchecked
    CHATT(8).value = vbUnchecked
    CHATT(9).value = vbUnchecked
    CHATT(10).value = vbUnchecked
    CHATT(11).value = vbUnchecked
   CHATT(12).value = vbUnchecked
   CHATT(13).value = vbUnchecked
   CHATT(14).value = vbUnchecked
   CHATT(15).value = vbUnchecked
   CHATT(16).value = vbUnchecked
   CHATT(17).value = vbUnchecked
   CHATT(18).value = vbUnchecked
   CHATT(19).value = vbUnchecked
   CHATT(20).value = vbUnchecked
   CHATT(2).value = vbUnchecked
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
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
MySQL = "SELECT     dbo.TblInstallmentsReq.ID, dbo.TblInstallmentsReq.RecordDate, dbo.TblInstallmentsReq.RecordDateH, dbo.TblInstallmentsReq.BranchID, "
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblInstallmentsReq.TypeRequest, dbo.TblInstallmentsReq.Cust_Mobile,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.ItemRequest, dbo.TblInstallmentsReq.Cus_ID, dbo.TblInstallmentsReq.BrithDateH, dbo.TblInstallmentsReq.BrithDate,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Company, dbo.TblInstallmentsReq.City, dbo.TblInstallmentsReq.CusID, dbo.TblInstallmentsReq.ExpDate, dbo.TblInstallmentsReq.PlaceID,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Salary, dbo.TblInstallmentsReq.BankName, dbo.TblInstallmentsReq.Salary_Acc, dbo.TblInstallmentsReq.Telephone,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Ext, dbo.TblInstallmentsReq.Home_Tel, dbo.TblInstallmentsReq.Street, dbo.TblInstallmentsReq.District, dbo.TblInstallmentsReq.Remarks,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.[No], dbo.TblInstallmentsReq.IBN, dbo.TblInstallmentsReq.Total_liab, dbo.TblInstallmentsReq.TypeAccept, dbo.TblInstallmentsReq.GurName,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.GurBrithDate, dbo.TblInstallmentsReq.GurBrithDateH, dbo.TblInstallmentsReq.GurCompany, dbo.TblInstallmentsReq.GurSalary,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.GurTotal_liab, dbo.TblInstallmentsReq.Relation1, dbo.TblInstallmentsReq.Relation2, dbo.TblInstallmentsReq.Adress1,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Adress2, dbo.TblInstallmentsReq.GurType_liab, dbo.TblInstallmentsReq.DirectManager, dbo.TblInstallmentsReq.GurTele,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.GurExt, dbo.TblInstallmentsReq.NameOffice, dbo.TblInstallmentsReq.NameAdmin_Mobile, dbo.TblInstallmentsReq.Attch0,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Attch1, dbo.TblInstallmentsReq.Attch2, dbo.TblInstallmentsReq.Attch3, dbo.TblInstallmentsReq.Attch4, dbo.TblInstallmentsReq.Attch5,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Attch6, dbo.TblInstallmentsReq.Attch7, dbo.TblInstallmentsReq.Attch8, dbo.TblInstallmentsReq.Attch9, dbo.TblInstallmentsReq.Attch10,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Attch11, dbo.TblInstallmentsReq.Attch12, dbo.TblInstallmentsReq.Attch13, dbo.TblInstallmentsReq.Attch14, dbo.TblInstallmentsReq.Attch15,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Attch16, dbo.TblInstallmentsReq.Attch17, dbo.TblInstallmentsReq.Attch18, dbo.TblInstallmentsReq.Accept, dbo.TblInstallmentsReq.Attch19,"
MySQL = MySQL & "                      dbo.TblInstallmentsReq.Cus_Name"
MySQL = MySQL & " FROM         dbo.TblInstallmentsReq LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblInstallmentsReq.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblInstallmentsReq.id =" & val(TxtSerial1.Text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
             StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepInstalmentRequest.rpt"
        Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepInstalmentRequest.rpt"
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

    Else

    End If
 xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    xReport.ParameterFields(3).AddCurrentValue user_name

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If BtnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
 '  If val(Me.TxtSerial1.text) <> 0 Then
 '      print_report
 '  End If
ErrTrap:
End Sub

Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Buy Goods in Installments Request     "

    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Req. ID"
    Me.lbl(2).Caption = "Date"
    Me.lbl(0).Caption = "Branch"
    Me.lbl(1).Caption = "Type Request"
    lbl(7).Caption = "Customer"
    Me.lbl(9).Caption = "Mobile"
    Me.lbl(5).Caption = "Goods"
    Me.lbl(10).Caption = "Company"
    lbl(6).Caption = "City"
    lbl(15).Caption = "Birth Date"
    lbl(3).Caption = "Birth Date"
    lbl(13).Caption = "Salary"
    lbl(14).Caption = "Salary in Acc."
    lbl(11).Caption = "ID"
    lbl(16).Caption = "ExpDate/Place"
    lbl(20).Caption = "IBN"
    lbl(24).Caption = "Street"
    lbl(19).Caption = "Bank"
   lbl(23).Caption = "District"
    lbl(22).Caption = "No"
    lbl(18).Caption = "Ext"
    lbl(26).Caption = "Remarks"
    lbl(21).Caption = "Home Phone"
    lbl(17).Caption = "Work Phone"
    lbl(25).Caption = "Obligations"
    Ele(2).Caption = "Attachments"
    Ele(1).Caption = "Guarantor"
    lbl(12).Caption = "Name"
    lbl(29).Caption = "Company"
    lbl(27).Caption = "Total Obligations"
    lbl(28).Caption = "Type Obligations"
    lbl(30).Caption = "Birth Date"
    lbl(31).Caption = "Birth Date"
    lbl(32).Caption = "Total Salary"
    CHATT(19).RightToLeft = False
    CHATT(19).Caption = "Sure"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' ISButton2.Caption = "Add"
    lbl(39).Caption = "Relationship"
    lbl(34).Caption = "Relationship"
    lbl(33).Caption = "Mobile,Address"
    lbl(35).Caption = "Mobile, Address"
    lbl(36).Caption = "Direct manager"
    lbl(37).Caption = "Phone/Ext"
    lbl(38).Caption = "Installment Offices"
    lbl(40).Caption = "Phone and Admin"
    '''/////////////
    CHATT(0).RightToLeft = False
    CHATT(1).RightToLeft = False
    CHATT(2).RightToLeft = False
    CHATT(3).RightToLeft = False
    CHATT(4).RightToLeft = False
    CHATT(5).RightToLeft = False
    CHATT(6).RightToLeft = False
    CHATT(7).RightToLeft = False
    CHATT(8).RightToLeft = False
    CHATT(9).RightToLeft = False
    CHATT(10).RightToLeft = False
    CHATT(11).RightToLeft = False
    CHATT(12).RightToLeft = False
    CHATT(13).RightToLeft = False
    CHATT(14).RightToLeft = False
    CHATT(15).RightToLeft = False
    CHATT(16).RightToLeft = False
    CHATT(17).RightToLeft = False
    CHATT(18).RightToLeft = False
    CHATT(20).RightToLeft = False
    '''/////////////////////////////
    
    CHATT(0).Caption = "ID"
    CHATT(1).Caption = "Office Google Site"
    CHATT(2).Caption = "Delivery of the goods"
    CHATT(3).Caption = "Purchase Requisition"
    CHATT(4).Caption = "Standing Order"
    CHATT(5).Caption = "Lease"
    CHATT(6).Caption = "ID letter Salary"
    CHATT(7).Caption = "Installment Contract"
    CHATT(8).Caption = "Direct Debit"
    CHATT(9).Caption = "Bank Statement"
    CHATT(10).Caption = "Voucher Order"
    CHATT(11).Caption = "Guarantee"
    CHATT(12).Caption = "Electricity Bill"
    CHATT(13).Caption = "Law Firm Pledged"
    CHATT(14).Caption = "Franchise Sale"
    CHATT(15).Caption = "Home Google Site"
    CHATT(16).Caption = "Installments Table"
    CHATT(17).Caption = "Cashing Voucher"
    CHATT(18).Caption = "Payment Voucher"
    CHATT(20).Caption = "Instrument Home"
    
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
 
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    BtnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblInstallmentsReq"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end



Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(XPDtbTrans.value)
 End If
End Sub
