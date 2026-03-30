VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVizits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "“Ì«—«  «·⁄„·«¡"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13440
   Icon            =   "FrmVizits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   3870
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4035
      Width           =   13485
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·“Ì«—…"
         Height          =   3015
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   13455
         Begin VB.TextBox TxtEmpRemark 
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
            Height          =   1035
            Left            =   0
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   1680
            Width           =   12090
         End
         Begin VB.TextBox TxtCusRemark 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   15
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   600
            Width           =   12090
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
            Left            =   10680
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   8220
            TabIndex        =   40
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94765057
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker FromTime 
            Height          =   315
            Left            =   5760
            TabIndex        =   42
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94765058
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToTime 
            Height          =   315
            Left            =   3000
            TabIndex        =   43
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94765058
            CurrentDate     =   38784
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·„‰œÊ»"
            Height          =   285
            Index           =   6
            Left            =   12120
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·⁄„Ì·"
            Height          =   285
            Index           =   7
            Left            =   12120
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label RecDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï «·”«⁄…"
            Height          =   285
            Index           =   1
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label RecDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·”«⁄…"
            Height          =   285
            Index           =   0
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   690
         End
         Begin VB.Label RecDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ "
            Height          =   285
            Index           =   11
            Left            =   9900
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   195
            Index           =   3
            Left            =   12585
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   510
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·„‰œÊ»"
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   0
         Width           =   6615
         Begin VB.TextBox TxtEmpPass 
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
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   50
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1665
         End
         Begin MSDataListLib.DataCombo DcbEmpUsrID 
            Height          =   315
            Left            =   2760
            TabIndex        =   34
            Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
            Top             =   240
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„"
            Height          =   285
            Index           =   5
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»«”Ê—œ"
            Height          =   285
            Index           =   1
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·⁄„Ì·"
         Height          =   735
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   0
         Width           =   6615
         Begin VB.TextBox TxtUserPass 
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
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   50
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   1665
         End
         Begin MSDataListLib.DataCombo DcbUserID 
            Height          =   315
            Left            =   2760
            TabIndex        =   28
            Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
            Top             =   240
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»«”Ê—œ"
            Height          =   285
            Index           =   4
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„"
            Height          =   285
            Index           =   0
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "FrmVizits.frx":57E2
         Left            =   2280
         List            =   "FrmVizits.frx":57F2
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -90
      Width           =   13515
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   4
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Visible         =   0   'False
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
            TabIndex        =   5
            Top             =   45
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   510
         Visible         =   0   'False
         Width           =   945
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
               Picture         =   "FrmVizits.frx":580B
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":5BA5
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":5F3F
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":62D9
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":6673
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":6A0D
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":6DA7
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmVizits.frx":7341
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   150
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
         ButtonImage     =   "FrmVizits.frx":76DB
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
         Top             =   150
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
         ButtonImage     =   "FrmVizits.frx":7A75
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   150
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
         ButtonImage     =   "FrmVizits.frx":7E0F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   150
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
         ButtonImage     =   "FrmVizits.frx":81A9
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "“Ì«—«  «·⁄„·«¡"
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
         Left            =   9855
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   3390
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1260
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7815
      Width           =   13440
      _cx             =   23707
      _cy             =   2223
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
      Begin VB.CommandButton Command1 
         Caption         =   "Send SMS"
         Height          =   375
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "Send Email"
         Height          =   375
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin ImpulseButton.ISButton btnNew 
         Height          =   330
         Left            =   10800
         TabIndex        =   14
         Top             =   675
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
         ButtonImage     =   "FrmVizits.frx":8543
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   7710
         TabIndex        =   15
         Top             =   675
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
         ButtonImage     =   "FrmVizits.frx":88DD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   9195
         TabIndex        =   16
         Top             =   675
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
         ButtonImage     =   "FrmVizits.frx":8C77
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   6465
         TabIndex        =   17
         Top             =   675
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
         ButtonImage     =   "FrmVizits.frx":9011
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   5220
         TabIndex        =   18
         Top             =   675
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
         ButtonImage     =   "FrmVizits.frx":93AB
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   90
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14737632
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
         ButtonImage     =   "FrmVizits.frx":9945
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   105
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
         ButtonImage     =   "FrmVizits.frx":9CDF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   345
         TabIndex        =   21
         Top             =   675
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
         ButtonImage     =   "FrmVizits.frx":A079
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         ButtonImage     =   "FrmVizits.frx":A413
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… «·ﬂ·"
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
         ButtonImage     =   "FrmVizits.frx":10C75
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton CMDSendMAil 
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«—”«· „Ì·"
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
         ButtonImage     =   "FrmVizits.frx":174D7
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   2505
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   975
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   3405
      Left            =   0
      TabIndex        =   26
      Top             =   600
      Width           =   13485
      _cx             =   23786
      _cy             =   6006
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmVizits.frx":1DD39
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
End
Attribute VB_Name = "FrmVizits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim OurEmployee As String

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Old Contract Data"
    Label1(2).Caption = Me.Caption

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("CusID")) = "Customer Name"
        .TextMatrix(0, .ColIndex("ContractNo")) = "Contract No."
        .TextMatrix(0, .ColIndex("ContractDate")) = "Contract Date"
.TextMatrix(0, .ColIndex("ContractValue")) = "Contract Value"
.TextMatrix(0, .ColIndex("EndGuranteeDate")) = "End Gurantee Date"
.TextMatrix(0, .ColIndex("NetValue")) = "Maintenance Value"
.TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With
    Label1(0).Caption = "Cont.No"
    Label1(4).Caption = "Contr.Date"
    Label1(1).Caption = "Customer"
    Label1(6).Caption = "End Gurantee"
    Label1(9).Caption = "Type"
     Label1(5).Caption = "Maint. Value"
    Label1(8).Caption = "Contract Value"
    Label1(7).Caption = "Remarks"
'XPPnlTime.Caption = "Print End Contruct"
    Label1(3).Caption = "ID"
 
 
    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."
 
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

'    If DoPremis(Do_Delete, Me.name, True) = False Then
'        Exit Sub
'    End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
Else
MsgBox "Please Select Employee"
End If
Exit Sub
End If
    If TxtVac_ID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
        MSGType = MsgBox("ÂConfirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
     End If

        If MSGType = vbYes Then
            RsSavRec.find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry... Can not Delete.  is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
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
        FindRec val(TxtVac_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

'    If DoPremis(Do_Edit, Me.name, True) = False Then
'        Exit Sub
'    End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
Else
MsgBox "Please Select Employee"
End If
Exit Sub
End If
Frame3.Enabled = False
Frame2.Enabled = False
    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
       ' Frm2.Enabled = True
      
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry" & CHR(13)
            Msg = Msg & " ·Currently can not be edited" & CHR(13)
            Msg = Msg & "Where it was being edited by another user on the network"
           
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

'    If DoPremis(Do_New, Me.name, True) = False Then
'        Exit Sub
'    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   ' Frm2.Enabled = True
    Frame2.Enabled = False
    '-----------------------------------
    Me.TxtVac_ID.Text = ""
Label3.Caption = ""
    Frame3.Enabled = False
    '-----------------------------------
    TxtModFlg.Text = "N"
clear_all Me
FillGridWithData
    My_SQL = "TblVisit"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If
Me.DcbUserID.BoundText = user_id
    rs.Close
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
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
        FindRec val(TxtVac_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub



Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
OurEmployee = DcbEmpUsrID.Text
    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ”„ «·„‰œÊ»"
Else
MsgBox "Please Enter Employee Name"
End If
Exit Sub
End If
If TxtEmpRemark.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· „·«ÕŸ«  «·„‰œÊ»"
Else
MsgBox "Please Enter Remarks"
End If
Exit Sub
End If
    '------------------------------ check if Empcode exist ----------------------

    'StrVacName = IsRecExist("TblVisit", "ID", Trim(TxtContractNo.Text), "ID", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    'If StrVacName <> "" Then
    'If SystemOptions.UserInterface = ArabicInterface Then
    '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
    '    Else
    '    Msg = "This type already exists"
    '    End If
    '    MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
    '    TxtContractNo.SetFocus
    '
    '    Exit Sub

    'End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    CreateMessage
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
   Else
   MsgBox "Sorry...error in douring enter data", vbOKOnly + vbMsgBoxRight, App.title
   End If

End Sub
 Function CreateMessage()
Dim subject As String
    'Dim msg As String
    Dim CompanyName As String
        Dim cOptions As ClsCompanyInfo
        Dim Msg As String
    Set cOptions = New ClsCompanyInfo
 
    CompanyName = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName
 

  '    subject = "“Ì«—… «·⁄„Ì· " & CompanyName & " »Ê«”ÿÂ  " & OurEmployee & "  » «—ÌŒ    " & RecordDate.value
     subject = CompanyName
     
      Msg = " ›«’Ì· «·“Ì«—…" & CHR(13)
      Msg = Msg & "«”„  «·⁄„Ì·:" & CompanyName & CHR(13)
      Msg = Msg & "«·‘Œ’ «·„”∆Ê·:" & DcbUserID & CHR(13)
      Msg = Msg & " «—ÌŒ «·“Ì«—…:" & RecordDate.value & CHR(13)
      Msg = Msg & "Êﬁ  «·“Ì«—…  :" & CHR(13)
      Msg = Msg & "„‰:" & FromTime & CHR(13)
      Msg = Msg & "«·Ì:" & ToTime & CHR(13)
       Msg = Msg & "Êﬂ«‰  „·«ÕŸ«  «·⁄„Ì· ﬂ«· «·Ì: " & TxtCusRemark & CHR(13)
       Msg = Msg & "Êﬂ«‰  „·«ÕŸ«  «·„‰œÊ» ﬂ«· «·Ì :" & TxtEmpRemark & CHR(13)
      Dim RetVal As String
   RetVal = SendMail("", _
        Trim$(subject), _
        "»ﬁŒ…", _
        Msg, _
        "txtServer", _
        25, _
         "txtUsername", _
        "txtPassword", _
       "", _
         False, True)
            Label3.Caption = IIf(RetVal = "ok", " „ «—”«· «·“Ì«—…", RetVal)
 
 End Function
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
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
        Msg = "No new data found"
    Else
        Msg = "No Rec.Before Update " & vbCrLf & FristCount & vbCrLf & "No Rec.After Update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "No New Record" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "No Deleted Record"" & vbCrLf & FristCount - LastCount"
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub




Private Sub CMDSendMAil_Click(Index As Integer)
CreateMessage
End Sub

 

Private Sub cmdSendMessage_Click()
Dim subject As String
Dim Msg As String
Dim msgstatus As String
    Dim CompanyName As String
        Dim cOptions As ClsCompanyInfo
       
    Set cOptions = New ClsCompanyInfo
 
    CompanyName = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName
 
 
subject = " »Œ’Ê’  ⁄ÃÌ· «·œ›⁄ "
     Msg = " ›«’Ì·  «·—”«·Â" & CHR(13)
      Msg = Msg & " „ ⁄„Ì· Œ’„ „”„ÊÕ »Â · ⁄ÃÌ· «·œ›⁄ ›Ì ”‰œ «·ﬁ»÷ —ﬁ„" & "1111" & CHR(13)
      Msg = Msg & "ﬁÌ„… «·”‰œ:" & "5500" & "ﬁÌ„… «·Œ’„:" & "550" & CHR(13)
      Msg = Msg & " «—ÌŒ «·”‰œ:" & "01/01/2017" & CHR(13)
      Msg = Msg & "Êﬁ  «·”‰œ  :" & Time & CHR(13)
 
       Msg = Msg & "Â–… «·—”«·Â „—”·Â «·Ì«.." & CHR(13)
       Msg = Msg & "„⁄ «· ÕÌ… " & CHR(13)
       Msg = Msg & CompanyName & CHR(13)
       


SendEmailForCustomer 4, subject, Msg, msgstatus
If msgstatus = "ok" Then
MsgBox "Done"
End If
End Sub

Private Sub Command1_Click()
Dim subject As String
Dim Msg As String
Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
     CompanyName = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName
 subject = " »Œ’Ê’  ⁄ÃÌ· «·œ›⁄ "
    Msg = Msg & " „ ⁄„Ì· Œ’„ „”„ÊÕ »Â · ⁄ÃÌ· «·œ›⁄ ›Ì ”‰œ «·ﬁ»÷ —ﬁ„" & "1111" & CHR(13)
      Msg = Msg & "ﬁÌ„… «·”‰œ:" & "5500" & "ﬁÌ„… «·Œ’„:" & "550" & CHR(13)
      Msg = Msg & " «—ÌŒ «·”‰œ:" & "01/01/2017" & CHR(13)
   Msg = Msg & "„⁄ «· ÕÌ… " & CHR(13)
       Msg = Msg & CompanyName & CHR(13)
 Dim t As String
t = sendMessageM("user", "password", Msg, CompanyName, GetCustomerNumber(4))
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos


    RecordDate = Date
    ToTime.value = Time
    FromTime.value = Time

    My_SQL = "TblVisit"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------

    Set Dcombos = New ClsDataCombos
    Set cSearch = New clsDCboSearch
    Dcombos.GetUsers DcbUserID
    Dcombos.GetUserComp DcbEmpUsrID
    FillGridWithData

    With Me.Grid
'        .Cell(flexcpPicture, 0, .ColIndex("ContractNo")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
'        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

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
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
         Dim xApp As New CRAXDRT.Application

    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report

    Dim rs As ADODB.Recordset
    Dim cCompanyInfo As ClsCompanyInfo
    Set cCompanyInfo = New ClsCompanyInfo
 Dim sql As String
 
sql = " SELECT     TOP 100 PERCENT dbo.TblVisit.ID, dbo.TblUsers.UserName, dbo.TblVisit.UserID, dbo.TblVisit.EmpUsrID, dbo.TblUserComp.Name, dbo.TblUserComp.NameE, "
sql = sql & "                       dbo.TblVisit.UserPass, dbo.TblVisit.EmpPass, dbo.TblVisit.RecordDate, dbo.TblVisit.FromTime, dbo.TblVisit.ToTime, dbo.TblVisit.CusRemark,"
sql = sql & "                       dbo.TblVisit.EmpRemark"
sql = sql & "  FROM         dbo.TblVisit INNER JOIN"
sql = sql & "                       dbo.TblUsers ON dbo.TblVisit.UserID = dbo.TblUsers.UserID INNER JOIN"
sql = sql & "                       dbo.TblUserComp ON dbo.TblVisit.EmpUsrID = dbo.TblUserComp.ID"
If Ind = 0 Then
sql = sql & " where dbo.TblVisit.ID=" & val(TxtVac_ID.Text) & " "
End If
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\REPORTS NEW\RepVisits.rpt")
    Else
    
        Set xReport = xApp.OpenReport(App.path & "\reports\REPORTS NEW\RepVisits.rpt")
    End If

   
    xReport.Database.SetDataSource rs
     Dim cAccountReport As New ClsReportViewer
    Dim FrmReport As New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.TxtPath = (App.path & "\reports\REPORTS NEW\RepVisits.rpt")
  
    
       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
       xReport.ParameterFields(2).AddCurrentValue user_name
    FrmReport.CRViewer.viewReport
 cAccountReport.CreateLogo xReport
  
    FrmReport.show
    Screen.MousePointer = vbDefault


End Function

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

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

    Set cSearch = Nothing
ErrTrap:
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblVisit", "id", "")
    RsSavRec.AddNew
    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap
     RsSavRec.Fields("UserID").value = IIf(DcbUserID.BoundText <> 0, val(DcbUserID.BoundText), Null)
     RsSavRec.Fields("EmpUsrID").value = IIf(DcbEmpUsrID.BoundText <> 0, val(DcbEmpUsrID.BoundText), Null)
     RsSavRec.Fields("UserPass").value = TxtUserPass.Text
     RsSavRec.Fields("EmpPass").value = TxtEmpPass.Text
     RsSavRec.Fields("RecordDate").value = RecordDate.value
     RsSavRec.Fields("FromTime").value = FormatDateTime(Me.FromTime.value, vbShortTime)
     RsSavRec.Fields("ToTime").value = FormatDateTime(Me.ToTime.value, vbShortTime)
     RsSavRec.Fields("CusRemark").value = TxtCusRemark.Text
     RsSavRec.Fields("EmpRemark").value = TxtEmpRemark.Text
     RsSavRec.update
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
    MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FiLLTXT
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
    Dim Shifttime As Date
   ' Frm2.Enabled = False
    Frame3.Enabled = False
Frame2.Enabled = False
Me.DcbEmpUsrID.BoundText = 0
TxtEmpPass.Text = ""
TxtUserPass.Text = ""
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtSerial.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    Me.DcbUserID.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
   ' Me.DcbEmpUsrID.BoundText = IIf(IsNull(RsSavRec.Fields("EmpUsrID").value), "", RsSavRec.Fields("EmpUsrID").value)
  '  TxtEmpPass.Text = IIf(IsNull(RsSavRec.Fields("EmpPass").value), "", RsSavRec.Fields("EmpPass").value)
  '  TxtUserPass.Text = IIf(IsNull(RsSavRec.Fields("UserPass").value), "", RsSavRec.Fields("UserPass").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    TxtCusRemark.Text = IIf(IsNull(RsSavRec.Fields("CusRemark").value), "", RsSavRec.Fields("CusRemark").value)
    TxtEmpRemark.Text = IIf(IsNull(RsSavRec.Fields("EmpRemark").value), "", RsSavRec.Fields("EmpRemark").value)
        If Not IsNull(RsSavRec("FromTime").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromTime").value, vbShortTime)
        Me.FromTime.value = Shifttime
    End If
        If Not IsNull(RsSavRec("ToTime").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToTime").value, vbShortTime)
        Me.ToTime.value = Shifttime
    End If


    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    FiLLRec

End Sub

Private Sub Grid_Click()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
    FiLLTXT
ErrTrap:
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Function CheckPassworUserComp() As Double
Dim StrSQL As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
       StrSQL = "Select * From TblUserComp Where  Password='" & Trim(Me.TxtEmpPass.Text) & "'"
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPassworUserComp = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
CheckPassworUserComp = 0
End If
End Function


Function CheckPassworUser() As Boolean
Dim StrSQL As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
       StrSQL = "Select * From TblUsers Where UserID=" & Me.DcbUserID.BoundText & " AND PassWord='" & Trim(Me.TxtUserPass.Text) & "'"
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPassworUser = True
Else
CheckPassworUser = False
End If
End Function

Private Sub ISButton2_Click(Index As Integer)
print_report , Index
End Sub

Private Sub TxtEmpPass_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.DcbEmpUsrID.BoundText = CheckPassworUserComp()
End If
End Sub

Private Sub TxtUserPass_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If CheckPassworUser() = True Then
Frame3.Enabled = True
Frame2.Enabled = True
Else
Frame3.Enabled = False
Frame2.Enabled = False
End If
End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "id=" & RecId, , adSearchForward, 1

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

Private Sub TxtModFlg_Change()
btnDelete.Enabled = False
    If TxtModFlg.Text = "N" Then
       ' Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    CMDSendMAil(2).Enabled = False
    ISButton2(0).Enabled = False
    ISButton2(1).Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
    ISButton2(0).Enabled = True
    ISButton2(1).Enabled = True
    CMDSendMAil(2).Enabled = True
     '   Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.Text <> "" Then
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
    
    ElseIf TxtModFlg.Text = "E" Then
    CMDSendMAil(2).Enabled = False
        ISButton2(0).Enabled = False
        ISButton2(1).Enabled = False
       ' Frm2.Enabled = True
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
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT     dbo.TblVisit.ID, dbo.TblUsers.UserName, dbo.TblVisit.UserID, dbo.TblVisit.EmpUsrID, dbo.TblUserComp.Name, dbo.TblUserComp.NameE, dbo.TblVisit.UserPass, "
    My_SQL = My_SQL & "                   dbo.TblVisit.EmpPass , dbo.TblVisit.RecordDate, dbo.TblVisit.FromTime, dbo.TblVisit.ToTime, dbo.TblVisit.CusRemark, dbo.TblVisit.EmpRemark"
    My_SQL = My_SQL & "    FROM         dbo.TblVisit INNER JOIN"
    My_SQL = My_SQL & "                  dbo.TblUsers ON dbo.TblVisit.UserID = dbo.TblUsers.UserID INNER JOIN"
    My_SQL = My_SQL & "                  dbo.TblUserComp ON dbo.TblVisit.EmpUsrID = dbo.TblUserComp.ID"
    My_SQL = My_SQL & " order by  dbo.TblVisit.ID                  "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs.Fields("RecordDate").value), "", rs.Fields("RecordDate").value)
                .TextMatrix(i, .ColIndex("FromTime")) = IIf(IsNull(rs.Fields("FromTime").value), "", rs.Fields("FromTime").value)
                .TextMatrix(i, .ColIndex("ToTime")) = IIf(IsNull(rs.Fields("ToTime").value), "", rs.Fields("ToTime").value)
                .TextMatrix(i, .ColIndex("CusRemark")) = IIf(IsNull(rs.Fields("CusRemark").value), "", rs.Fields("CusRemark").value)
                .TextMatrix(i, .ColIndex("EmpRemark")) = IIf(IsNull(rs.Fields("EmpRemark").value), "", rs.Fields("EmpRemark").value)
               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
               Else
               .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("NameE").value), "", rs.Fields("NameE").value)
               End If
                rs.MoveNext
            Next

            rs.Close
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



