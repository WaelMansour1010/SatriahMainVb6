VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCustCash2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… »Ì«‰«  «·⁄„·«¡ «·‰ﬁœÌ"
   ClientHeight    =   12165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12735
   ClipControls    =   0   'False
   Icon            =   "FrmCustCash2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmCustCash2.frx":058A
   RightToLeft     =   -1  'True
   ScaleHeight     =   12165
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.PictureBox Picture1 
      Height          =   10935
      Left            =   120
      Picture         =   "FrmCustCash2.frx":246A8C
      RightToLeft     =   -1  'True
      ScaleHeight     =   10875
      ScaleWidth      =   12555
      TabIndex        =   12
      Top             =   405
      Width           =   12615
      Begin VB.Frame Frame1 
         Height          =   540
         Left            =   5145
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   7095
         Width           =   3780
         Begin VB.OptionButton optIsPoliticianNo 
            Alignment       =   1  'Right Justify
            Caption         =   "·«"
            Height          =   195
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   225
            Width           =   840
         End
         Begin VB.OptionButton optIsPoliticianYes 
            Alignment       =   1  'Right Justify
            Caption         =   "‰⁄„"
            Height          =   195
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   225
            Width           =   945
         End
      End
      Begin VB.OptionButton optMailAdressNo 
         Alignment       =   1  'Right Justify
         Height          =   240
         Left            =   5475
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   2085
         Width           =   210
      End
      Begin VB.OptionButton optMailAdressYes 
         Alignment       =   1  'Right Justify
         Height          =   240
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   2115
         Width           =   210
      End
      Begin VB.TextBox txtCardSource 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4350
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   4470
         Width           =   1980
      End
      Begin VB.CheckBox chkIsNegativListNo 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   10290
         Width           =   495
      End
      Begin VB.CheckBox chkIsNegativList 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   10200
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   10290
         Width           =   495
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5130
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   1590
         Width           =   4560
      End
      Begin VB.TextBox txtMailAdress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   2550
         Width           =   4530
      End
      Begin VB.TextBox txtPoliticianJobName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5370
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   7740
         Width           =   4530
      End
      Begin VB.TextBox txtWorkAddress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5340
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   8280
         Width           =   4530
      End
      Begin VB.TextBox txtPhoneNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   5010
         Width           =   4530
      End
      Begin VB.TextBox txtCompanyName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   6660
         Width           =   4530
      End
      Begin VB.TextBox txtJobName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6030
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   6180
         Width           =   4530
      End
      Begin VB.TextBox txtPhone2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5970
         MaxLength       =   11
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   5520
         Width           =   4530
      End
      Begin VB.TextBox TxtCardNo 
         Alignment       =   1  'Right Justify
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2910
         MaxLength       =   18
         TabIndex        =   53
         Top             =   9420
         Width           =   6915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   4080
         MaxLength       =   18
         TabIndex        =   52
         Top             =   30
         Width           =   6915
      End
      Begin VB.TextBox txtId 
         Alignment       =   1  'Right Justify
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   5160
         MaxLength       =   14
         TabIndex        =   50
         Top             =   3960
         Width           =   5475
      End
      Begin VB.TextBox TXTEmail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   8910
         Width           =   3840
      End
      Begin VB.TextBox TxtCashCustomerName 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4980
         TabIndex        =   48
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox txtEnglishName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   9420
         MaxLength       =   50
         TabIndex        =   47
         Top             =   1155
         Width           =   1095
      End
      Begin VB.TextBox txtEnglishName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   8325
         MaxLength       =   50
         TabIndex        =   46
         Top             =   1155
         Width           =   1095
      End
      Begin VB.TextBox txtEnglishName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   7245
         MaxLength       =   50
         TabIndex        =   45
         Top             =   1155
         Width           =   1095
      End
      Begin VB.TextBox txtEnglishName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   6090
         MaxLength       =   50
         TabIndex        =   44
         Top             =   1155
         Width           =   1095
      End
      Begin VB.TextBox txtEnglishName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4980
         MaxLength       =   50
         TabIndex        =   43
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Frame Frm2 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2085
         Left            =   -6060
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   6450
         Width           =   10035
         Begin VB.ComboBox CmbType 
            BackColor       =   &H80000018&
            Height          =   315
            ItemData        =   "FrmCustCash2.frx":40F18E
            Left            =   2280
            List            =   "FrmCustCash2.frx":40F19E
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2310
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8010
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   45
            Width           =   1065
         End
         Begin VB.TextBox TxtVacName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5235
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   390
            Width           =   3840
         End
         Begin VB.TextBox TxtVacNamee 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   405
            Width           =   3840
         End
         Begin VB.TextBox Txtdiscount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   720
            Width           =   3840
         End
         Begin VB.ComboBox CBOCartTYpe 
            Height          =   315
            ItemData        =   "FrmCustCash2.frx":40F1B7
            Left            =   2895
            List            =   "FrmCustCash2.frx":40F1C1
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ  "
            Height          =   285
            Index           =   3
            Left            =   8460
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   285
            Index           =   0
            Left            =   8820
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ì"
            Height          =   285
            Index           =   1
            Left            =   3225
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   480
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·Œ’„"
            Height          =   285
            Index           =   4
            Left            =   4065
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   795
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·›Ê‰"
            Height          =   285
            Index           =   5
            Left            =   8850
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   780
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬂ— "
            Height          =   285
            Index           =   6
            Left            =   8820
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1155
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ﬂ«— "
            Height          =   285
            Index           =   7
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’·«ÕÌ…"
            Height          =   285
            Index           =   8
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ì„Ì·"
            Height          =   285
            Index           =   9
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1440
            Width           =   1155
         End
      End
      Begin MSComCtl2.DTPicker txtBirthDate 
         Height          =   330
         Left            =   9120
         TabIndex        =   51
         Top             =   3540
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   187236353
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker dtCardDate 
         Height          =   330
         Left            =   9390
         TabIndex        =   54
         Top             =   4410
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   187236355
         CurrentDate     =   38718
      End
      Begin MSComCtl2.DTPicker dtCardEndDate 
         Height          =   330
         Left            =   390
         TabIndex        =   55
         Top             =   4440
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   187236355
         CurrentDate     =   38718
      End
      Begin MSDataListLib.DataCombo DCNationality 
         Height          =   315
         Left            =   7830
         TabIndex        =   67
         Top             =   3060
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3945
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   2
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
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
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
         Top             =   0
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
               Picture         =   "FrmCustCash2.frx":40F1D0
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":40F56A
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":40F904
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":40FC9E
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":410038
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":4103D2
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":41076C
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustCash2.frx":410D06
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmCustCash2.frx":4110A0
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   7
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmCustCash2.frx":41143A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmCustCash2.frx":4117D4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmCustCash2.frx":411B6E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰«  «·⁄„·«¡ «·‰ﬁœÌ"
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
         Left            =   5295
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   90
         Width           =   4320
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   105
      Left            =   0
      TabIndex        =   11
      Top             =   570
      Width           =   10095
      _cx             =   17806
      _cy             =   185
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCustCash2.frx":411F08
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
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   960
      Left            =   150
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   11340
      Width           =   12540
      _cx             =   22119
      _cy             =   1693
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
      Begin ImpulseButton.ISButton btnNew 
         Height          =   330
         Left            =   8175
         TabIndex        =   30
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
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
         ButtonImage     =   "FrmCustCash2.frx":412011
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   6630
         TabIndex        =   31
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
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
         ButtonImage     =   "FrmCustCash2.frx":4123AB
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   7395
         TabIndex        =   32
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
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
         ButtonImage     =   "FrmCustCash2.frx":412745
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   5865
         TabIndex        =   33
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
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
         ButtonImage     =   "FrmCustCash2.frx":412ADF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   4860
         TabIndex        =   34
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
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
         ButtonImage     =   "FrmCustCash2.frx":412E79
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   3840
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   555
         Width           =   840
         _ExtentX        =   1482
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
         ButtonImage     =   "FrmCustCash2.frx":413413
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   -135
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
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
         ButtonImage     =   "FrmCustCash2.frx":4137AD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   2445
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   555
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â «·ﬂ«— "
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
         ButtonImage     =   "FrmCustCash2.frx":413B47
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   1425
         TabIndex        =   38
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
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
         ButtonImage     =   "FrmCustCash2.frx":413EE1
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnExport 
         Height          =   330
         Left            =   390
         TabIndex        =   73
         Top             =   510
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Export"
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
         ButtonImage     =   "FrmCustCash2.frx":41427B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   5025
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   225
         Width           =   975
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmCustCash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    'If TxtVac_ID.text <> "" Then
    '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
    '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
    Else
    MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
    End If

    If MSGType = vbYes Then
        RsSavRec.Find "id=" & val(TxtVac_ID.text), , adSearchForward, 1
        RsSavRec.delete
       If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
       Else
       MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
       End If
        '------------------------------ Move Next ---------------------------.
        FillGridWithData
        BtnNext_Click
    End If

    'End If
    Exit Sub
ErrTrap:
 
    Select Case err.Number

        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
         '   StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Public Sub btnExport_Click()
    On Error GoTo err
    Dim FileName As String
    '***********************
    CommonDialog1.filter = "Excel (*.xls)|*.xls"
    CommonDialog1.DefaultExt = "xls"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen
    FileName = CommonDialog1.FileName
    '***********************
    Dim rst As New ADODB.Recordset

    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWs As Object

    Dim recArray As Variant
    
    Dim fldCount As Integer
    Dim recCount As Long
    Dim iCol As Integer
    Dim iRow As Integer
    
    rst.Open "Select * From TblCusCsh", Cn
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    
    Set xlWB = xlApp.Workbooks.Open(FileName)
   
    Set xlWs = xlWB.Worksheets("Sheet1")
    xlApp.UserControl = True
    '1 Token
    '2 embossing Name(25)
    '3 extension Name(25)
    '4 magstripe Name(25)
    '5 national id(14)
    '6 address1 (35)
    '7 address2 (35)
    '8 address3 (35)
    '9 sms Flag(1)
    '10 Mobile Number(10)
    '11 birth date (10)
    
    Dim row As Integer
    row = 3
    Do While Not rst.EOF
    
        ' xlWs.Cells(row, 1).value = rst.Fields(iCol - 1).Name
        Dim arr
        arr = Split(rst!Name, " ")
        If UBound(arr) = 0 Then
            xlWs.Cells(row, 2).value = arr(0)
        ElseIf UBound(arr) = 1 Then
            xlWs.Cells(row, 2).value = arr(0)
            xlWs.Cells(row, 3).value = arr(1)
        ElseIf UBound(arr) = 2 Then
            xlWs.Cells(row, 2).value = arr(0)
            xlWs.Cells(row, 3).value = arr(1)
            xlWs.Cells(row, 4).value = arr(2)
        End If
        
        xlWs.Cells(row, 5).value = rst!CardID & ""
        xlWs.Cells(row, 6).value = rst!Address & ""
        xlWs.Cells(row, 10).value = rst!PhoneNo2 & ""
        xlWs.Cells(row, 11).value = rst!BirthDate & ""
        row = row + 1
        rst.MoveNext
    Loop

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.rows.AutoFit
    
    rst.Close
    
    Set rst = Nothing
    
    
    Set xlWs = Nothing
    Set xlWB = Nothing

    Set xlApp = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
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

    Select Case err.Number

        Case -2147217885
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
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

    Select Case err.Number

        Case -2147217885
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case err.Number

        Case -2147467259
            'Could not update; currently locked.
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
           Else
            Msg = "Sorry..." & CHR(13)
            Msg = Msg & " This record can not be edited at this time" & CHR(13)
            Msg = Msg & "Because it was modified by another user on the network"
         
           End If
              MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim Rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set Rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.text = "N"

    My_SQL = "TblCusCsh"
    Rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If Rs.RecordCount > 0 Then
        TxtSerial.text = Rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    Rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
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

    Select Case err.Number

        Case -2147217885
     If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
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

    Select Case err.Number

        Case -2147217885
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrint_Click()
print_report2
End Sub

Private Sub btnQuery_Click()
Unload frmCashCustomerSearch
frmCashCustomerSearch.RetrunType = 4
Load frmCashCustomerSearch
frmCashCustomerSearch.RetrunType = 4
frmCashCustomerSearch.show
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg        As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt    As Control
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblCusCsh", "nameˆ", Trim(TxtVacName.text), "name", "ID<>'" & Trim(TxtVac_ID.text) & "'")

    If StrVacName <> "" Then
        ' Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "I have already registered this type before"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    ' MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Error...douring entering data", vbOKOnly + vbMsgBoxRight, App.Title
    End If

End Sub
Sub LoadData(ID)
    Dim s As String
    s = ""
    s = s & "SELECT * "
    s = s & "FROM TblCusCsh "
    s = s & "WHERE Id = " & ID & ";"
    Dim Rs As New ADODB.Recordset
    Rs.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If Not Rs.EOF Then
        TxtCashCustomerName = Rs!Name & ""
        Dim arr
        arr = Split(Rs!NameE, " ")
        If UBound(arr) = 0 Then
            txtEnglishName(0) = arr(0)
        ElseIf UBound(arr) = 1 Then
            txtEnglishName(0) = arr(0)
            txtEnglishName(1) = arr(1)
        ElseIf UBound(arr) = 2 Then
            txtEnglishName(0) = arr(0)
            txtEnglishName(1) = arr(1)
            txtEnglishName(2) = arr(2)
        ElseIf UBound(arr) = 3 Then
            txtEnglishName(0) = arr(0)
            txtEnglishName(1) = arr(1)
            txtEnglishName(2) = arr(2)
            txtEnglishName(3) = arr(3)
        ElseIf UBound(arr) = 4 Then
            txtEnglishName(0) = arr(0)
            txtEnglishName(1) = arr(1)
            txtEnglishName(2) = arr(2)
            txtEnglishName(3) = arr(3)
            txtEnglishName(4) = arr(4)
        End If
       
        txtPhoneNo = Rs!tel & ""
        '  card,
        ' discount,
        ' CartTYpe,
        'RecordDate,
        TXTEmail = Rs!Email & ""
        txtAddress = Rs!Address & ""
        If IsNull(Rs!IsMailAdress) Then
            optMailAdressYes = False
            optMailAdressNo = True
        Else
            optMailAdressYes = Rs!IsMailAdress
            optMailAdressNo = Not Rs!IsMailAdress
        End If
    
        txtMailAdress = Rs!MailAdress & ""
        DCNationality.SelectedItem = val(Rs!Nationality & "")
        txtBirthDate = Rs!BirthDate & ""
        TxtCardNo = Rs!CardID & ""
        dtCardDate = Rs!CardDate & ""
        txtCardSource = Rs!CardSource & ""
        dtCardEndDate = Rs!CardEndDate & ""
        '  txtPhoneNo PhoneNo,
        txtPhone2 = Rs!PhoneNo2 & ""
        txtJobName = Rs!jobname & ""
        txtCompanyName = Rs!CompanyName & ""
        If IsNull(Rs!IsPolitician) Then
            optIsPoliticianYes = False
            optIsPoliticianNo = True
        Else
            optIsPoliticianYes = Rs!IsPolitician
            optIsPoliticianNo = Not Rs!IsPolitician
        End If
       
        txtPoliticianJobName = Rs!PoliticianJobName & ""
        txtWorkAddress = Rs!WorkAddress & ""
        TxtCardNo = Rs!CardNO & ""
        If IsNull(Rs!IsNegativList) Then
            chkIsNegativList = False
            chkIsNegativListNo = True
        Else
            chkIsNegativList = Rs!IsNegativList
            chkIsNegativListNo = Not Rs!IsNegativList
        End If
        
    End If

End Sub

Sub SaveData()
    Dim s As String
    s = ""
    s = s & "SELECT * "
    s = s & "FROM TblCusCsh "
    s = s & "WHERE  1 = 2 "
    Dim Rs As New ADODB.Recordset
    Rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!Name = TxtCashCustomerName
    
    Rs!NameE = txtEnglishName(0) & " " & _
       txtEnglishName(1) & " " & _
       txtEnglishName(2) & " " & _
       txtEnglishName(3) & " " & _
       txtEnglishName(4)
     
    Rs!tel = txtPhoneNo
    Rs!Email = TXTEmail
    Rs!Address = txtAddress
    Rs!IsMailAdress = optMailAdressYes
    Rs!MailAdress = txtMailAdress
    Rs!Nationality = val(DCNationality.SelectedItem)
    Rs!BirthDate = CDate(txtBirthDate)
    Rs!CardID = TxtCardNo
    Rs!CardDate = CDate(dtCardDate)
    Rs!CardSource = txtCardSource
    Rs!CardEndDate = CDate(dtCardEndDate)
    Rs!PhoneNo2 = txtPhone2
    Rs!jobname = txtJobName
    Rs!CompanyName = txtCompanyName
    Rs!IsPolitician = optIsPoliticianYes
    Rs!PoliticianJobName = txtPoliticianJobName
    Rs!WorkAddress = txtWorkAddress
    Rs!CardNO = TxtCardNo
    Rs!IsNegativList = chkIsNegativList = False
    Rs.update
End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.text)
    Me.TxtModFlg.text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
If SystemOptions.UserInterface = ArabicInterface Then
    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

Else
    If FristCount = LastCount Then
        Msg = "No new data"
    Else
        Msg = "Number of records before update" & vbCrLf & FristCount & vbCrLf & "Number of records after  update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "Number of new records" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "Number of records deleted" & vbCrLf & FristCount - LastCount
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
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
MySQL = "SELECT     dbo.TblCusCsh.*"
MySQL = MySQL & " From dbo.TblCusCsh"
MySQL = MySQL & " Where (id =" & val(TxtVac_ID.text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repCashCustomer.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repCashCustomer.rpt"
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
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String

    My_SQL = "TblCusCsh"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
   btnQuery.Caption = "Search"
   BtnPrint.Caption = "Print Card"
    Me.Caption = "Data Of Customers Cash"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name AR"
    Label1(1).Caption = "Name ENG"
      Label1(6).Caption = "Card No"
    Label1(5).Caption = "Telephone"
    Label1(4).Caption = "Discount"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
Label1(7).Caption = "Type Card "
Label1(8).Caption = "Validity"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Id"
        .TextMatrix(0, .ColIndex("name")) = "Name AR"
        .TextMatrix(0, .ColIndex("namee")) = "Name ENG"
         '.TextMatrix(0, .ColIndex("discount")) = "Discount"
        .TextMatrix(0, .ColIndex("card")) = "Card No"
        .TextMatrix(0, .ColIndex("tel")) = "Telephone "
        .TextMatrix(0, .ColIndex("CartTYpe")) = "Type Card "
   
    End With

End Sub

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
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
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

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblCusCsh", "id", "")
    RsSavRec.AddNew
    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
      RsSavRec.Fields("tel").value = IIf(txtTel.text <> "", Trim(txtTel.text), Null)
      RsSavRec.Fields("Email").value = IIf(TXTEmail.text <> "", Trim(TXTEmail.text), Null)
      
      
    RsSavRec.Fields("card").value = IIf(Txtcard.text <> "", Trim(Txtcard.text), Null)
RsSavRec.Fields("discount").value = IIf(Txtdiscount.text <> "", Trim(Txtdiscount.text), 0)
RsSavRec.Fields("CartTYpe").value = IIf(CBOCartTYpe.text <> "", (CBOCartTYpe.text), "")
          RsSavRec("RecordDate").value = XPDtbBill.value
          
    RsSavRec.update
 If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 End If

    FillGridWithData
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("namee").value), "", RsSavRec.Fields("namee").value)
    
   ' Txttel.text = IIf(IsNull(RsSavRec.Fields("tel").value), "", RsSavRec.Fields("tel").value)
   ' Txtcard.text = IIf(IsNull(RsSavRec.Fields("card").value), "", RsSavRec.Fields("card").value)
    Txtdiscount.text = IIf(IsNull(RsSavRec.Fields("discount").value), "", RsSavRec.Fields("discount").value)
    TXTEmail.text = IIf(IsNull(RsSavRec.Fields("Email").value), "", RsSavRec.Fields("Email").value)
    
    
    
        CBOCartTYpe.text = IIf(IsNull(RsSavRec.Fields("CartTYpe").value), "", RsSavRec.Fields("CartTYpe").value)
 'XPDtbBill.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
 
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub


Private Sub Tet_NumPoket_Change()
 Dim mInput As String
   
    Dim mTxt As String
    mInput = Tet_NumPoket.text
    
    If Len(mInput) > 13 Then
        Dim firstThree As String
        firstThree = mId(mInput, 1, 3)
        
        Dim lastThree As String
        lastThree = mId(mInput, Len(mInput) - 2, 3)
        
        Dim asterisks As String
        asterisks = String(Len(mInput) - 6, "*")
        
        mTxt = firstThree & asterisks & lastThree ' ????? ???? ?????? ??? ??? ???????
    End If
End Sub

Function ExtractInfoFromNationalID(nationalID As String) As Date
    Dim day As Integer
    Dim Month As Integer
    Dim year As Integer
    Dim centuryCode As Integer
    Dim governorateCode As Integer
    Dim governorate As String
    
    ' Extract information from national ID
    centuryCode = CInt(mId(nationalID, 1, 1))
    year = CInt(mId(nationalID, 2, 2))
    Month = CInt(mId(nationalID, 4, 2))
    day = CInt(mId(nationalID, 6, 2))
    governorateCode = CInt(mId(nationalID, 8, 2))
    
    ' Determine century
    If centuryCode = 2 Then
        year = year + 1900
    ElseIf centuryCode = 3 Then
        year = year + 2000
    End If
    
    ' Determine governorate
    Select Case governorateCode
        Case 1
            governorate = "«·ﬁ«Â—…"
        Case 2
            governorate = "«·≈”ﬂ‰œ—Ì…"
        ' Add more cases for other governorates...
        Case Else
            governorate = "€Ì— „⁄—Ê›"
    End Select
    
    ' Format the date
    Dim dateOfBirth As String
    dateOfBirth = Format(day, "00") & "/" & Format(Month, "00") & "/" & year
    
    ' Return the information
    'ExtractInfoFromNationalID = " «—ÌŒ «·„Ì·«œ: " & dateOfBirth & vbCrLf & "«·„Õ«›Ÿ…: " & governorate
    
    ExtractInfoFromNationalID = dateOfBirth
End Function



Private Sub Tet_NumPoket_Validate(Cancel As Boolean)
   Dim mInput As String
   
    Dim mTxt As String
    mInput = Tet_NumPoket.text
    
    If Len(mInput) > 13 Then
        Dim firstThree As String
        firstThree = mId(mInput, 1, 3)
        
        Dim lastThree As String
        lastThree = mId(mInput, Len(mInput) - 2, 3)
        
        Dim asterisks As String
        asterisks = String(Len(mInput) - 6, "*")
        
        mTxt = firstThree & asterisks & lastThree ' ????? ???? ?????? ??? ??? ???????
    End If
    txtBirthDate.value = ExtractInfoFromNationalID(Tet_NumPoket)
End Sub

Private Sub TxtCashCustomerName_Change()
If Me.TxtModFlg.text <> "R" Then

Dim sentence As String
sentence = TxtCashCustomerName.text

Dim words() As String
Dim word As Variant
words = ExtractWords(sentence)
Dim i As Long
i = 0
For Each word In words
    txtEnglishName(i).text = Translate(0, CStr(word))
    i = i + 1
    If i > 4 Then Exit Sub
    'MsgBox word
Next word


'txtEnglishName(0).Text = Translate(0, TxtCashCustomerName.Text)
End If

End Sub
Function ExtractWords(ByVal sentence As String) As String()
    Dim words() As String
    Dim delimiter As String
    delimiter = " " ' Ì„ﬂ‰ﬂ  €ÌÌ— «·„Õœœ Õ”» «·Õ«Ã…
    
    words = Split(sentence, delimiter)
    
    ExtractWords = words
End Function


Private Sub TxtCashCustomerName_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
Private Sub txtEnglishName_GotFocus(Index As Integer)
SwitchKeyboardLang LANG_ENGLISH
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '  btnNext.Enabled = False
        '  btnPrevious.Enabled = False
        '  btnFirst.Enabled = False
        '  btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim Rs As ADODB.Recordset
    Dim My_SQL As String

    Set Rs = New ADODB.Recordset
    My_SQL = "select * From TblCusCsh order by id"
    Rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If Rs.RecordCount > 0 Then
            .rows = Rs.RecordCount + 1
            Rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs.Fields("name").value), "", Rs.Fields("name").value)
               
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs.Fields("namee").value), "", Rs.Fields("namee").value)
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs.Fields("id").value), "", Rs.Fields("id").value)
                
                .TextMatrix(i, .ColIndex("tel")) = IIf(IsNull(Rs.Fields("tel").value), "", Rs.Fields("tel").value)
               
                .TextMatrix(i, .ColIndex("card")) = IIf(IsNull(Rs.Fields("card").value), "", Rs.Fields("card").value)
               
                .TextMatrix(i, .ColIndex("CartTYpe")) = IIf(IsNull(Rs.Fields("CartTYpe").value), "", Rs.Fields("CartTYpe").value)
                Rs.MoveNext
            Next

            Rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

'-------------------------------------------------------------
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
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

'Private Function CheckDelCountry(Lngid As Long) As Boolean
'    Dim rs As ADODB.Recordset
'    Dim StrSQL As String
'    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        CheckDelCountry = False
'    Else
'        CheckDelCountry = True
'    End If

'    rs.Close
'    Set rs = Nothing
'End Function

Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
