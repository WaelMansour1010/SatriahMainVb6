VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmWarrantyOffer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14190
   Icon            =   "FrmWarrantyOffer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   14190
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
      Height          =   735
      Left            =   0
      TabIndex        =   38
      Top             =   840
      Width           =   14295
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   8520
         TabIndex        =   40
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93192193
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmWarrantyOffer.frx":6852
         Height          =   315
         Left            =   2280
         TabIndex        =   41
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
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
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   2
         Left            =   10170
         TabIndex        =   44
         Top             =   255
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         Height          =   285
         Index           =   4
         Left            =   12960
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   7
         Left            =   7320
         TabIndex        =   42
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frmam 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   35
      Top             =   1560
      Width           =   14295
      Begin VB.TextBox TxtAllIDS 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Top             =   1680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox TxtAllDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   66
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   1815
         Left            =   10200
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   840
         Width           =   3975
         Begin VB.TextBox txtvlaue 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Text            =   "12"
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton GranteeTypeopt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»œÊ‰ «·ﬁÿ⁄"
            Height          =   195
            Index           =   0
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton GranteeTypeopt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄ «·ﬁÿ⁄"
            Height          =   195
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin MSComCtl2.DTPicker GranteeStartDate 
            Height          =   330
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            _Version        =   393216
            Format          =   93192193
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker GranteeEndDate 
            Height          =   330
            Left            =   120
            TabIndex        =   57
            Top             =   1320
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            _Version        =   393216
            Format          =   93192193
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘Â—"
            Height          =   255
            Index           =   23
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   960
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "› —… «·÷„«‰"
            Height          =   255
            Index           =   22
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   960
            Width           =   2235
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·÷„«‰"
            Height          =   255
            Index           =   21
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ‰Â«Ì… «·÷„«‰"
            Height          =   255
            Index           =   20
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1440
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   19
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   120
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ »œ«Ì… «·÷„«‰"
            Height          =   255
            Index           =   16
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   600
            Width           =   1995
         End
      End
      Begin VB.ComboBox DcbSandType 
         Height          =   315
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   3645
      End
      Begin VB.TextBox TxtOrderNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtRemark 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   960
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo DcbProject 
         Bindings        =   "FrmWarrantyOffer.frx":6867
         Height          =   315
         Left            =   2400
         TabIndex        =   64
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
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
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   315
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         Caption         =   " ›«’Ì· «·÷„«‰"
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
         ButtonImage     =   "FrmWarrantyOffer.frx":687C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         LowerToggledContent=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ "
         Height          =   285
         Index           =   0
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   18
         Left            =   1320
         TabIndex        =   47
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   46
         Top             =   960
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ« "
         Height          =   285
         Index           =   13
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ì"
         Height          =   285
         Index           =   5
         Left            =   12840
         TabIndex        =   36
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmWarrantyOffer.frx":D0DE
      Left            =   15480
      List            =   "FrmWarrantyOffer.frx":D0EE
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
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
      TabIndex        =   3
      Top             =   0
      Width           =   14505
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   4
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
         ButtonImage     =   "FrmWarrantyOffer.frx":D107
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   5
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
         ButtonImage     =   "FrmWarrantyOffer.frx":D4A1
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   6
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
         ButtonImage     =   "FrmWarrantyOffer.frx":D83B
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
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
         ButtonImage     =   "FrmWarrantyOffer.frx":DBD5
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ”ÃÌ· «·÷„«‰"
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
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmWarrantyOffer.frx":DF6F
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   11
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
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
            Picture         =   "FrmWarrantyOffer.frx":F374
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":F70E
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":FAA8
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":FE42
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":101DC
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":10576
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":10910
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWarrantyOffer.frx":10EAA
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   13
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
      ButtonImage     =   "FrmWarrantyOffer.frx":11244
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   15
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
      ButtonImage     =   "FrmWarrantyOffer.frx":17AA6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   16
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
      ButtonImage     =   "FrmWarrantyOffer.frx":1E308
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   6705
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2400
      Width           =   14235
      _cx             =   25109
      _cy             =   11827
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   23
         Top             =   6000
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   24
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
            ButtonImage     =   "FrmWarrantyOffer.frx":1E6A2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9480
            TabIndex        =   25
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
            ButtonImage     =   "FrmWarrantyOffer.frx":24F04
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   26
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
            ButtonImage     =   "FrmWarrantyOffer.frx":2529E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7680
            TabIndex        =   27
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
            ButtonImage     =   "FrmWarrantyOffer.frx":2BB00
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5760
            TabIndex        =   28
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
            ButtonImage     =   "FrmWarrantyOffer.frx":2BE9A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   29
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
            ButtonImage     =   "FrmWarrantyOffer.frx":2C434
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmWarrantyOffer.frx":2C7CE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmWarrantyOffer.frx":33030
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   5520
         Width           =   3855
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   255
            Width           =   675
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   540
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9600
         TabIndex        =   32
         Top             =   5640
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   -840
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
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   5010
         Left            =   0
         TabIndex        =   68
         Top             =   360
         Width           =   14175
         _cx             =   25003
         _cy             =   8837
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmWarrantyOffer.frx":333CA
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
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·«’‰«›"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   9
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   12840
         TabIndex        =   34
         Top             =   5640
         Width           =   900
      End
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·«”Â„ «·„ﬂ  »…"
      Height          =   285
      Index           =   15
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   1725
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
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmWarrantyOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double
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
    Cn.Execute "delete TBLRegularMaint   where WarntID= " & val(Me.TxtSerial1.Text)
    RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        If val(DcbProject.BoundText) <> 0 Then
            If (TxtAllDate.Text) <> "" Then
                AllDate = TxtAllDate.Text
                Alliids = TxtAllIDS.Text
                
                astrSplitItems = Split(AllDate, strFilterText)
         astrSplitItems1 = Split(Alliids, strFilterText)
         
                For intX = 0 To UBound(astrSplitItems)
                        
                    If IsDate(astrSplitItems(intX)) Then
                        RsgGrantee.AddNew
                        RsgGrantee("DateOfRegularMaint").value = Format$(astrSplitItems(intX), "dd/mm/yyyy")
                        RsgGrantee("MaintenanceIDS").value = astrSplitItems1(intX)
                        If val(Me.DcbSandType.ListIndex) = 1 Then
                        RsgGrantee("itemid").value = val(DcbProject.BoundText)
                        Else
                        RsgGrantee("itemid").value = val(XPTxtBillID.Text)
                        End If
                        RsgGrantee("WarntID").value = val(Me.TxtSerial1.Text)
                        If GranteeTypeopt(0).value = True Then
                        RsgGrantee("GranteeType").value = 0
                        Else
                        RsgGrantee("GranteeType").value = 1
                        End If
                        RsgGrantee("GranteeStartDate").value = GranteeStartDate.value
                        RsgGrantee("GranteeEndDate").value = GranteeEndDate.value
                        RsgGrantee("Count").value = 1
                        RsgGrantee("TypeTrnas").value = 1
                        RsgGrantee("Done").value = 0
                       

                        RsgGrantee.update
                    End If
                       
                Next intX
                    
            End If

        End If
End Function

Private Sub DcbProject_Change()
DcbProject_Click (0)
End Sub

Private Sub DcbProject_Click(Area As Integer)
Dim Fullcode As String
If val(DcbProject.BoundText) <> 0 Then
'SALAHX GetProjectsCode_ID val(DcbProject.BoundText), fullcode
TxtOrderNo.Text = Fullcode
End If
End Sub

Private Sub DcbProject_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Load FrmProjectSearch
FrmProjectSearch.lblSearchtype.Caption = 20
FrmProjectSearch.show
End If
End Sub

Private Sub DcbSandType_Change()
DcbProject.Visible = False
ISButton2.Visible = False
lbl(9).Visible = False
FG.Visible = False
If val(DcbSandType.ListIndex) = 1 Then
DcbProject.Visible = True
ISButton2.Visible = True
ElseIf val(DcbSandType.ListIndex) = 2 Then
ISButton2.Visible = True
ElseIf val(DcbSandType.ListIndex) = 0 Then
lbl(9).Visible = True
FG.Visible = True
End If
End Sub

Private Sub DcbSandType_Click()
DcbSandType_Change
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With FG
Select Case .ColKey(Col)
Case "Show"
Load FRMGranteeData
FRMGranteeData.inde = 2
FRMGranteeData.LngRow = Row

FRMGranteeData.lbl(4).Caption = .TextMatrix(Row, .ColIndex("Code"))
FRMGranteeData.lbl(5).Caption = .TextMatrix(Row, .ColIndex("Name"))
FRMGranteeData.AllDate = .TextMatrix(Row, .ColIndex("RegularMaintenancedates"))
FRMGranteeData.AllIDS = .TextMatrix(Row, .ColIndex("RegularMaintenanceIDS"))
FRMGranteeData.GranteeStartDate.value = IIf((Not IsDate(.TextMatrix(Row, .ColIndex("GranteeStartDate")))), Date, (.TextMatrix(Row, .ColIndex("GranteeStartDate"))))
FRMGranteeData.GranteeEndDate.value = IIf(Not IsDate(.TextMatrix(Row, .ColIndex("GranteeEndDate"))), Date, (.TextMatrix(Row, .ColIndex("GranteeEndDate"))))
FRMGranteeData.txtvlaue.Text = .TextMatrix(Row, .ColIndex("Period"))
FRMGranteeData.FillGridWithData
If val(.TextMatrix(Row, .ColIndex("GranteeType"))) = 0 Then
FRMGranteeData.GranteeTypeopt(0).value = True
Else
FRMGranteeData.GranteeTypeopt(1).value = True
End If
FRMGranteeData.show
End Select
End With

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "Show"
.ColComboList(.ColIndex("Show")) = "..."
End Select
End With
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblWarrantyOffer order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbSandType
    .Clear
    .AddItem "„»Ì⁄« "
    .AddItem "„‘«—Ì⁄"
    .AddItem "⁄ﬁÊœ ﬁœÌ„…"
    End With
   Else
 With DcbSandType
    .Clear
    .AddItem "Sales"
    .AddItem "Project"
    .AddItem "Old Contract"
    End With
    End If
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
      Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetProjects Me.DcbProject
    
    BtnLast_Click
    ShowTip
    
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
   
   Me.Refresh
ErrTrap:
End Sub
Function SaveGranteeDataSales()
  
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
    Cn.Execute "delete TBLRegularMaint   where WarntID= " & val(Me.TxtSerial1.Text) & "  Or Transaction_ID =" & val(Me.XPTxtBillID.Text) & ""
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
                        RsgGrantee("WarntID").value = val(Me.TxtSerial1.Text)
                        RsgGrantee("itemid").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                        RsgGrantee("GranteeType").value = val(FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")))
                        RsgGrantee("GranteeStartDate").value = FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate"))
                        RsgGrantee("GranteeEndDate").value = FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate"))
                        'RsgGrantee("ItemSerial").value = FG.TextMatrix(RowNum, FG.ColIndex("Serial"))
                        RsgGrantee("Count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
                        RsgGrantee("Done").value = 0
                        RsgGrantee("Count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
                        RsgGrantee("TypeTrnas").value = 0
                        RsgGrantee.update
                    End If
                       
                Next intX
                    
            End If

        End If

    Next RowNum
    
End Function
Function GetContractID(Optional ContractNo As String) As Double
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
sql = "SELECT ID FROm TblOLDContract WHERE  ContractNo='" & (ContractNo) & "'"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetContractID = IIf(IsNull(Rs6("ID").value), 0, Rs6("ID").value)
Else
GetContractID = 0
End If
End Function
' save new reco

Function GetTransID(Optional NoteSerial1 As String) As Double
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
sql = "SELECT Transaction_ID FROm Transactions WHERE Transaction_Type=21 and NoteSerial1='" & (NoteSerial1) & "'"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetTransID = IIf(IsNull(Rs6("Transaction_ID").value), 0, Rs6("Transaction_ID").value)
Else
GetTransID = 0
End If
End Function
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
  Dim StrSQL As String
  If Me.TxtModFlg.Text = "E" Then
  StrSQL = "Delete From TblWarrantyOfferDet Where WrantID='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
  
  End If
    RsSavRec.Fields("RecorDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("SandType").value = val(Me.DcbSandType.ListIndex)
    RsSavRec.Fields("Remarks").value = Me.txtremark.Text
    RsSavRec.Fields("OrderNo").value = Me.TxtOrderNo.Text
    RsSavRec.Fields("USerID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("AllDate").value = IIf(TxtAllDate.Text <> "", Trim(TxtAllDate.Text), Null)
    RsSavRec.Fields("AllIDS").value = IIf(TxtAllIDS.Text <> "", Trim(TxtAllIDS.Text), Null)
    RsSavRec.Fields("ProjectID").value = IIf(Me.DcbProject.BoundText <> "", val(DcbProject.BoundText), Null)
    RsSavRec.Fields("PriodG").value = IIf(Me.txtvlaue.Text <> "", val(txtvlaue.Text), Null)
    RsSavRec.Fields("GranteeStartDate").value = GranteeStartDate.value
    RsSavRec.Fields("GranteeEndDate").value = GranteeEndDate.value
    If GranteeTypeopt(0).value = True Then
    RsSavRec.Fields("GranteeTypeopt").value = 0
    Else
    RsSavRec.Fields("GranteeTypeopt").value = 1
    End If
    RsSavRec.Fields("TransectionID").value = val(XPTxtBillID.Text)
    
    RsSavRec.update
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblWarrantyOfferDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With FG
       For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("ItemID")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("WrantID").value = val(Me.TxtSerial1.Text)
                RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                RsDevsub("RegularMaintenancedates").value = IIf((.TextMatrix(i, .ColIndex("RegularMaintenancedates"))) = "", Null, .TextMatrix(i, .ColIndex("RegularMaintenancedates")))
                RsDevsub("RegularMaintenanceIDS").value = IIf((.TextMatrix(i, .ColIndex("RegularMaintenanceIDS"))) = "", Null, .TextMatrix(i, .ColIndex("RegularMaintenanceIDS")))
                RsDevsub("Price").value = IIf((.TextMatrix(i, .ColIndex("Price"))) = "", Null, val(.TextMatrix(i, .ColIndex("Price"))))
                RsDevsub("Count").value = IIf((.TextMatrix(i, .ColIndex("Count"))) = "", Null, val(.TextMatrix(i, .ColIndex("Count"))))
                RsDevsub("GranteeStartDate").value = IIf(Not IsDate(.TextMatrix(i, .ColIndex("GranteeStartDate"))), Null, (.TextMatrix(i, .ColIndex("GranteeStartDate"))))
                RsDevsub("GranteeEndDate").value = IIf(Not IsDate(.TextMatrix(i, .ColIndex("GranteeEndDate"))), Null, (.TextMatrix(i, .ColIndex("GranteeEndDate"))))
                RsDevsub("GranteeType").value = IIf((.TextMatrix(i, .ColIndex("GranteeType"))) = "", Null, val(.TextMatrix(i, .ColIndex("GranteeType"))))
                RsDevsub("Period").value = IIf((.TextMatrix(i, .ColIndex("Period"))) = "", Null, val(.TextMatrix(i, .ColIndex("Period"))))
                RsDevsub.update
        End If
      Next i
     End With
     If val(Me.DcbSandType.ListIndex) = 0 Then
     SaveGranteeDataSales
    Else
    SaveGranteeData
   End If
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
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
    Dim i As Integer
  
     TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
     XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecorDate").value), Date, RsSavRec.Fields("RecorDate").value)
     dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
     
     txtremark.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
     DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("USerID").value), "", RsSavRec.Fields("UserID").value)
     TxtOrderNo.Text = IIf(IsNull(RsSavRec.Fields("OrderNo").value), "", RsSavRec.Fields("OrderNo").value)
     TxtAllDate.Text = IIf(IsNull(RsSavRec.Fields("AllDate").value), "", RsSavRec.Fields("AllDate").value)
     TxtAllIDS.Text = IIf(IsNull(RsSavRec.Fields("AllIDS").value), "", RsSavRec.Fields("AllIDS").value)
     Me.DcbProject.BoundText = IIf(IsNull(RsSavRec.Fields("ProjectID").value), "", RsSavRec.Fields("ProjectID").value)
     Me.txtvlaue.Text = IIf(IsNull(RsSavRec.Fields("PriodG").value), 0, RsSavRec.Fields("PriodG").value)
     GranteeStartDate.value = IIf(IsNull(RsSavRec.Fields("GranteeStartDate").value), Date, RsSavRec.Fields("GranteeStartDate").value)
     GranteeEndDate.value = IIf(IsNull(RsSavRec.Fields("GranteeEndDate").value), Date, RsSavRec.Fields("GranteeEndDate").value)
     XPTxtBillID.Text = IIf(IsNull(RsSavRec.Fields("TransectionID").value), 0, RsSavRec.Fields("TransectionID").value)
     If Not IsNull(RsSavRec.Fields("GranteeTypeopt").value) Then
     If RsSavRec.Fields("GranteeTypeopt").value = 0 Then
     GranteeTypeopt(0).value = True
     Else
     GranteeTypeopt(1).value = True
     End If
     End If
     Me.DcbSandType.ListIndex = IIf(IsNull(RsSavRec.Fields("SandType").value), -1, RsSavRec.Fields("SandType").value)
     FullGridData
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 80
     LabCountRec.Caption = RsSavRec.RecordCount
ErrTrap:
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
  sql = " SELECT     dbo.TblWarrantyOfferDet.ID, dbo.TblWarrantyOfferDet.WrantID, dbo.TblWarrantyOfferDet.RegularMaintenancedates, "
  sql = sql + "                     dbo.TblWarrantyOfferDet.RegularMaintenanceIDS, dbo.TblWarrantyOfferDet.Price, dbo.TblWarrantyOfferDet.[Count], dbo.TblWarrantyOfferDet.ItemID,"
  sql = sql + "                     dbo.TblItems.itemname , dbo.TblItems.fullcode, dbo.TblItems.ItemNamee ,dbo.TblWarrantyOfferDet.GranteeEndDate ,dbo.TblWarrantyOfferDet.GranteeStartDate ,dbo.TblWarrantyOfferDet.GranteeType ,dbo.TblWarrantyOfferDet.Period"
  sql = sql + "  FROM         dbo.TblWarrantyOfferDet LEFT OUTER JOIN"
  sql = sql + "                      dbo.TblItems ON dbo.TblWarrantyOfferDet.ItemID = dbo.TblItems.ItemID"
  sql = sql + "  Where (dbo.TblWarrantyOfferDet.WrantID = " & val(TxtSerial1.Text) & ") "
  
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.FG
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RegularMaintenancedates")) = IIf(IsNull(Rs1("RegularMaintenancedates").value), "", Rs1("RegularMaintenancedates").value)
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
                   .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs1("Price").value), 0, Rs1("Price").value)
                   .TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(Rs1("Count").value), 0, Rs1("Count").value)
                   .TextMatrix(i, .ColIndex("RegularMaintenanceIDS")) = IIf(IsNull(Rs1("RegularMaintenanceIDS").value), "", Rs1("RegularMaintenanceIDS").value)
                    .TextMatrix(i, .ColIndex("GranteeEndDate")) = IIf(IsNull(Rs1("GranteeEndDate").value), Date, Rs1("GranteeEndDate").value)
                    .TextMatrix(i, .ColIndex("GranteeStartDate")) = IIf(IsNull(Rs1("GranteeStartDate").value), Date, Rs1("GranteeStartDate").value)
                    .TextMatrix(i, .ColIndex("GranteeType")) = IIf(IsNull(Rs1("GranteeType").value), 0, Rs1("GranteeType").value)
                    .TextMatrix(i, .ColIndex("Period")) = IIf(IsNull(Rs1("Period").value), 0, Rs1("Period").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                  
                    Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                  
                    End If
                   .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
            
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
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
      If dcBranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
           dcBranch.SetFocus
            Exit Sub
     End If

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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblWarrantyOffer", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Sub ReturnItems(Optional Transaction_ID As Double = 0)
Dim StrSQL As String
Dim RsDetails  As ADODB.Recordset
Dim i As Integer
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblItems.ItemName, dbo.TblItems.ItemID, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee, dbo.Transaction_Details.Transaction_ID , dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice"
    StrSQL = StrSQL & " ,RegularMaintenancedates, RegularMaintenanceIDS , dbo.Transaction_Details.guaranteeTime, dbo.Transaction_Details.GranteeStartDate, dbo.Transaction_Details.GranteeEndDate, dbo.Transaction_Details.GranteeType"
    StrSQL = StrSQL & " FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID"
    StrSQL = StrSQL & " Where (dbo.Transaction_Details.Transaction_ID = " & Transaction_ID & ")"
    StrSQL = StrSQL & " ORDER BY dbo.Transaction_Details.ID"
    

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' XPTxtSum.text = ""

    If RsDetails.RecordCount > 0 Then
        FG.Rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
              FG.Cell(flexcpData, i, FG.ColIndex("Ser")) = i
              FG.TextMatrix(i, FG.ColIndex("RegularMaintenanceIDS")) = IIf(IsNull(RsDetails("RegularMaintenanceIDS")), "", (RsDetails("RegularMaintenanceIDS").value))
              FG.TextMatrix(i, FG.ColIndex("RegularMaintenancedates")) = IIf(IsNull(RsDetails("RegularMaintenancedates")), "", (RsDetails("RegularMaintenancedates").value))
             ''////////////
              FG.TextMatrix(i, FG.ColIndex("Period")) = IIf(IsNull(RsDetails("guaranteeTime")), 0, (RsDetails("guaranteeTime").value))
              FG.TextMatrix(i, FG.ColIndex("GranteeStartDate")) = IIf(IsNull(RsDetails("GranteeStartDate")), Date, (RsDetails("GranteeStartDate").value))
              FG.TextMatrix(i, FG.ColIndex("GranteeEndDate")) = IIf(IsNull(RsDetails("GranteeEndDate")), Date, (RsDetails("GranteeEndDate").value))
              FG.TextMatrix(i, FG.ColIndex("GranteeType")) = IIf(IsNull(RsDetails("GranteeType")), 0, (RsDetails("GranteeType").value))
             ''////
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Fullcode")), "", (RsDetails("Fullcode").value))
            If SystemOptions.UserInterface = ArabicInterface Then
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim$(RsDetails("ItemName").value))
            Else
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemNamee")), "", Trim$(RsDetails("ItemNamee").value))
            End If
            FG.TextMatrix(i, FG.ColIndex("ItemID")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
           
             FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
         
            RsDetails.MoveNext
        
            If FG.Rows > 10 Then
                If i = 8 Then FG.Refresh
            End If

        Next i
End If
End Sub




Private Sub ISButton2_Click()
Load FRMGranteeData
FRMGranteeData.inde = 1

FRMGranteeData.lbl(4).Caption = TxtOrderNo.Text
FRMGranteeData.lbl(5).Caption = DcbProject.Text
FRMGranteeData.AllDate = Me.TxtAllDate.Text
FRMGranteeData.AllIDS = Me.TxtAllIDS.Text
FRMGranteeData.GranteeStartDate.value = GranteeStartDate.value
FRMGranteeData.GranteeEndDate.value = GranteeEndDate.value
FRMGranteeData.txtvlaue.Text = txtvlaue.Text
FRMGranteeData.FillGridWithData
If GranteeTypeopt(0).value = True Then
FRMGranteeData.GranteeTypeopt(0).value = True
Else
FRMGranteeData.GranteeTypeopt(1).value = True
End If
FRMGranteeData.show
End Sub



Private Sub TxtOrderNo_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbSandType.ListIndex) = 0 Then
XPTxtBillID.Text = GetTransID(Me.TxtOrderNo.Text)
ReturnItems val(XPTxtBillID.Text)
ElseIf val(DcbSandType.ListIndex) = 2 Then
XPTxtBillID.Text = GetContractID(Me.TxtOrderNo.Text)
End If
End If
End Sub

Private Sub TxtOrderNo_KeyPress(KeyAscii As Integer)
Dim ProjectID As Double
If val(DcbSandType.ListIndex) = 1 Then
If TxtOrderNo.Text <> "" Then
'SALAHXX GetProjectsCode_ID ProjectID, TxtOrderNo.text
DcbProject.BoundText = ProjectID
End If
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
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
    Dim Msg As String
    On Error GoTo ErrTrap
    
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
  Dim StrSQL As String
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
         StrSQL = "Delete From TblWarrantyOfferDet Where WrantID=" & val(TxtSerial1.Text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
             Cn.Execute "delete TBLRegularMaint   where WarntID= " & val(Me.TxtSerial1.Text)
            
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                     

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

'               'cleargriid
            LabCurrRec.Caption = 0
     LabCountRec.Caption = 0

   End If                         '------------------------------ Move Next ---------------------------.
        Me.Refresh
       
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
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
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    DcbProject.Enabled = True
    TxtOrderNo.Enabled = True
    DcbSandType.Enabled = True
    XPDtbTrans.Enabled = True
        
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
       XPDtbTrans.Enabled = False
       DcbProject.Enabled = False
       TxtOrderNo.Enabled = False
       DcbSandType.Enabled = False
    
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
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
       DcbProject.Enabled = True
    TxtOrderNo.Enabled = True
    DcbSandType.Enabled = True
 XPDtbTrans.Enabled = True
      
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
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
'        'cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    'cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
    
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    
    clear_all Me

    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = branch_id
    dcBranch.SetFocus

ErrTrap:
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
      'cleargriid
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
    'cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
        'cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     'cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
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
        If btnDelete.Enabled = False Then Exit Sub
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


Private Sub ChangeLang()
On Error GoTo ErrTrap
  
    Me.Caption = "Guarantee "
    
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
  
    Me.lbl(7).Caption = "Branch"
  lbl(0).Caption = "No"
    lbl(5).Caption = "Type "
 
    lbl(9).Caption = "Items"
ISButton2.Caption = "Show Guarantee "
   
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
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
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
   With Me.FG
'
      .TextMatrix(0, .ColIndex("Ser")) = "Ser"
      .TextMatrix(0, .ColIndex("Code")) = "Code No"
       .TextMatrix(0, .ColIndex("Name")) = "Item Name"
      .TextMatrix(0, .ColIndex("Count")) = "Qty"
       .TextMatrix(0, .ColIndex("Price")) = "Price "
     .TextMatrix(0, .ColIndex("Show")) = "Guarantee"
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblWarrantyOffer"
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
