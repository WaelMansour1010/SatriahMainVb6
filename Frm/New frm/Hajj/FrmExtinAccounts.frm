VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmExtinAccounts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   Icon            =   "FrmExtinAccounts.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      TabIndex        =   7
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmExtinAccounts.frx":6852
      Left            =   15480
      List            =   "FrmExtinAccounts.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   8
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Visible         =   0   'False
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
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
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
            Picture         =   "FrmExtinAccounts.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExtinAccounts.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   10
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
      ButtonImage     =   "FrmExtinAccounts.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Visible         =   0   'False
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
      ButtonImage     =   "FrmExtinAccounts.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Visible         =   0   'False
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
      ButtonImage     =   "FrmExtinAccounts.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8880
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      _cx             =   23627
      _cy             =   15663
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   13440
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   540
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   13455
         _cx             =   23733
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
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5460
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   360
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox TxtOrderID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   1485
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10965
            TabIndex        =   0
            Top             =   120
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   8415
            TabIndex        =   1
            Top             =   120
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102957057
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   6960
            TabIndex        =   46
            Top             =   120
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   3600
            TabIndex        =   47
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   3480
            TabIndex        =   60
            Top             =   360
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   315
            Index           =   27
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï «„— ‘€· —ﬁ„"
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   49
            Top             =   120
            Width           =   1605
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   11
            Left            =   5910
            TabIndex        =   41
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   25
            Left            =   9795
            TabIndex        =   40
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   255
            Index           =   4
            Left            =   12510
            TabIndex        =   16
            Top             =   120
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   990
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   7905
         Width           =   13365
         _cx             =   23574
         _cy             =   1746
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12090
            TabIndex        =   18
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   1110
            _ExtentX        =   1958
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
            ButtonImage     =   "FrmExtinAccounts.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   10485
            TabIndex        =   19
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
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
            ButtonImage     =   "FrmExtinAccounts.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9420
            TabIndex        =   2
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   885
            _ExtentX        =   1561
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
            ButtonImage     =   "FrmExtinAccounts.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8220
            TabIndex        =   20
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   1035
            _ExtentX        =   1826
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
            ButtonImage     =   "FrmExtinAccounts.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7230
            TabIndex        =   21
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
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
            ButtonImage     =   "FrmExtinAccounts.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   330
            Left            =   3990
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… Œÿ«» «· —‘ÌÕ"
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExtinAccounts.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   5760
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmExtinAccounts.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   1800
            TabIndex        =   24
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   600
            Width           =   915
            _ExtentX        =   1614
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
            ButtonImage     =   "FrmExtinAccounts.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   25
            Top             =   90
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   315
            TabIndex        =   30
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   29
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   28
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   195
            Index           =   0
            Left            =   3255
            TabIndex        =   27
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   345
            Index           =   14
            Left            =   12270
            TabIndex        =   26
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Width           =   13440
         _cx             =   23707
         _cy             =   1376
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
         BackColor       =   16777215
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
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   135
            TabIndex        =   35
            Top             =   240
            Width           =   465
            _ExtentX        =   820
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
            ButtonImage     =   "FrmExtinAccounts.frx":2A8D1
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   675
            TabIndex        =   36
            Top             =   240
            Width           =   450
            _ExtentX        =   794
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
            ButtonImage     =   "FrmExtinAccounts.frx":2AC6B
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1350
            TabIndex        =   37
            Top             =   240
            Width           =   465
            _ExtentX        =   820
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
            ButtonImage     =   "FrmExtinAccounts.frx":2B005
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1950
            TabIndex        =   38
            Top             =   240
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "FrmExtinAccounts.frx":2B39F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12375
            Picture         =   "FrmExtinAccounts.frx":2B739
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„”«—«  «·⁄„—… «·„Õ–Ê›…"
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
            Left            =   6975
            TabIndex        =   39
            Top             =   240
            Width           =   4665
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   6015
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
         Width           =   13455
         _cx             =   23733
         _cy             =   10610
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
         ForeColor       =   8388608
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
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   5355
            Left            =   90
            TabIndex        =   43
            Top             =   180
            Width           =   13215
            _cx             =   23310
            _cy             =   9446
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmExtinAccounts.frx":2CB3E
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   0
            Left            =   12360
            TabIndex        =   44
            Top             =   5625
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmExtinAccounts.frx":2CCCE
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   1
            Left            =   10200
            TabIndex        =   45
            Top             =   5625
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ·"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmExtinAccounts.frx":2D268
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   300
            Index           =   2
            Left            =   6840
            TabIndex        =   51
            Top             =   5625
            Width           =   2085
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   300
            Index           =   1
            Left            =   8880
            TabIndex        =   50
            Top             =   5625
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   11
         Left            =   0
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   7320
         Width           =   13455
         _cx             =   23733
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
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   120
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   3210
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   120
            Width           =   2385
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   165
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   120
            Width           =   2340
         End
         Begin ImpulseButton.ISButton BtnCreateVou 
            Height          =   375
            Left            =   9840
            TabIndex        =   57
            Top             =   120
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "«‰‘«¡ «·ﬁÌœ"
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
            ButtonImage     =   "FrmExtinAccounts.frx":2D802
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton BtnDeleteVoun 
            Height          =   375
            Left            =   7200
            TabIndex        =   58
            Top             =   120
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬁÌœ"
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
            ButtonImage     =   "FrmExtinAccounts.frx":34064
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   35
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   765
         End
      End
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
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmExtinAccounts"
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
Dim StrDepID As String

Private Sub BtnCreateVou_Click()
        If ChekClodePeriod(RecordDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰   ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "This is Period is Closed"
           End If
              Exit Sub
        End If
If TxtNoteSerial.Text = "" Then
    Dim Account_Code_dynamic As String
   Account_Code_dynamic = get_account_code_branch(135, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
        MsgBox "Please Create Branch"
        End If
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «Ì—«œ«  «·⁄„—…", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
                
    
createVoucher
MsgBox " „ «‰‘«¡ «·ﬁÌœ"
Else
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
End If
End Sub


Private Sub BtnDeleteVoun_Click()
       If ChekClodePeriod(RecordDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰   ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "This is Period is Closed"
           End If
              Exit Sub
        End If
       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
       Cn.Execute "Update TblExtinAccounts set NoteSerial=null ,NoteID=null where id=" & val(TxtSerial1.Text) & ""
TxtNoteSerial.Text = ""
TxtNoteID.Text = 0
MsgBox " „ Õ–› «·ﬁÌœ"
RsSavRec.Resync adAffectCurrent
End Sub

Private Sub Cmd_Click(Index As Integer)
Dim i As Integer
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
ReLineGrid
Case 1
With Fg1
     .Clear flexClearScrollable, flexClearEverything
     .Rows = 1
   ReLineGrid
 End With
End Select
End If
End Sub
Private Sub RemoveGridRow()
    With Me.Fg1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Sub ReLineGrid()
Dim i As Integer
Dim Conter As Integer
Conter = 0
Dim SumValue As Double
SumValue = 0
With Fg1
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Name")) <> "" Then
Conter = Conter + 1
.TextMatrix(i, .ColIndex("Serial")) = Conter
 If val(.TextMatrix(i, .ColIndex("PathID"))) <> 0 Then
 SumValue = SumValue + val(.TextMatrix(i, .ColIndex("Valuee")))
 End If
End If
Next i
lbl(2).Caption = SumValue

End With

End Sub
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblExtinAccounts", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Private Sub Fill_Combos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetCompany OutClientID, 2, 0
End Sub


Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub Fg1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub
 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "  «ÿ›«¡ —ﬁ„ " & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim Sql As String
tablename = "TblExtinAccounts"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.Text)
NoteSerial = ""
Notevalue = 0
 notytype = 9058
Notevalue = val(lbl(2).Caption)
BranchID = val(Me.DcbBranch.BoundText)
NoteDate = (RecordDate.value)
 
If Notevalue > 0 Then
                              
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, RecordDateH.value         ',
                                              TxtNoteID.Text = NoteID
                                                    TxtNoteSerial.Text = NoteSerial
                                  '  Else
                                  '               If TxtNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                  '       CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                  '                            TxtNoteID.text = NoteID
                                  '                           TxtNoteSerial.text = NoteSerial
                                  '               Else
                                  '                            Sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                  '                            Sql = Sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                  '                               Sql = Sql & " where NoteID=" & val(TxtNoteID.text)
                                  '                               Cn.Execute Sql
                                        
                                  '              End If
                            

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim StrTempCustomerCodeInsuranceAccount  As String
    
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 Dim valuee As Double
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
 LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID
valuee = val(lbl(2).Caption)
 StrTempAccountCode = get_account_code_branch(136, my_branch)
      StrTempDes = "«ÿ›«¡ —ﬁ„   " & TxtSerial1.Text

             LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & "      Õ”«» Œ’„ „”«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
         StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.OutClientID.BoundText))
         
          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» «·⁄„Ì·  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
ErrTrap:
End Function
Function print_report1(Optional NoteSerial As String, Optional ID As Double)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.TblExtinAccounts.ID, dbo.TblExtinAccounts.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
 MySQL = MySQL & "                     dbo.TblExtinAccounts.RecordDate, dbo.TblExtinAccounts.RecordDateH, dbo.TblExtinAccounts.OrderID, dbo.TblExtinAccounts.Total, dbo.TblExtinAccountsDer.Reson,"
 MySQL = MySQL & "                      dbo.TblExtinAccountsDer.PathStus, dbo.TblExtinAccountsDer.Total AS TotalDet, dbo.TblExtinAccountsDer.TotaMorhl, dbo.TblExtinAccountsDer.Remain,"
 MySQL = MySQL & "                     dbo.TblExtinAccountsDer.Valuee, dbo.TblExtinAccountsDer.PathID, dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblExtinAccountsDer.ID AS IDDet"
 MySQL = MySQL & " FROM         dbo.TblShrines RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblExtinAccountsDer ON dbo.TblShrines.ID = dbo.TblExtinAccountsDer.PathID RIGHT OUTER JOIN"
 MySQL = MySQL & "                       dbo.TblExtinAccounts ON dbo.TblExtinAccountsDer.ExtAccID = dbo.TblExtinAccounts.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblExtinAccounts.BranchID = dbo.TblBranchesData.branch_id"
 MySQL = MySQL & " Where (dbo.TblExtinAccountsDer.PathStus = 3) And (dbo.TblExtinAccounts.ID = " & val(TxtSerial1.Text) & ") And (dbo.TblExtinAccountsDer.ID = " & ID & ")"
   
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExtinAccountsCancel.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExtinAccountsCancel.rpt"
        End If
    If Dir(StrFileName) = "" Then
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
    
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     
     
        StrReportTitle = ""
 
    End If
    Dim DisCont As Double
    xReport.ParameterFields(4).AddCurrentValue DisCont
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
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.TblExtinAccounts.ID, dbo.TblExtinAccounts.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
 MySQL = MySQL & "                     dbo.TblExtinAccounts.RecordDate, dbo.TblExtinAccounts.RecordDateH, dbo.TblExtinAccounts.OrderID, dbo.TblExtinAccounts.Total, dbo.TblExtinAccountsDer.Reson,"
 MySQL = MySQL & "                     dbo.TblExtinAccountsDer.PathStus, dbo.TblExtinAccountsDer.Total AS TotalDet, dbo.TblExtinAccountsDer.TotaMorhl, dbo.TblExtinAccountsDer.Remain,"
 MySQL = MySQL & "                     dbo.TblExtinAccountsDer.valuee , dbo.TblExtinAccountsDer.PathID, dbo.TblShrines.name, dbo.TblShrines.NameE"
 MySQL = MySQL & " FROM         dbo.TblShrines RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblExtinAccountsDer ON dbo.TblShrines.ID = dbo.TblExtinAccountsDer.PathID RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblExtinAccounts ON dbo.TblExtinAccountsDer.ExtAccID = dbo.TblExtinAccounts.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblExtinAccounts.BranchID = dbo.TblBranchesData.branch_id"
 MySQL = MySQL & " Where (dbo.TblExtinAccounts.ID = " & val(TxtSerial1.Text) & ")"
   
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExtinAccounts.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExtinAccounts.rpt"
        End If
    If Dir(StrFileName) = "" Then
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
    
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     
     
        StrReportTitle = ""
 
    End If
    Dim DisCont As Double
    xReport.ParameterFields(4).AddCurrentValue DisCont
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
Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "Name"
Cancel = True
Case "Valuee"
If val(.TextMatrix(Row, .ColIndex("PathStus"))) = 3 Then
.ComboList = ""
Else
Cancel = True
End If
Case "Reson"
If val(.TextMatrix(Row, .ColIndex("PathStus"))) = 3 Then
.ComboList = ""
Else
Cancel = True
End If
Case "Total"
Cancel = True
Case "TotaMorhl"
Cancel = True
Case "Remain"
Cancel = True
Case "Show1"
.ComboList = ""
End Select
End With
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Fg1
Select Case .ColKey(Col)
Case "Show1"
print_report1 , val(.TextMatrix(Row, .ColIndex("ID")))
End Select
End With
End Sub

Private Sub Fg1_Click()
ReLineGrid
End Sub


Private Sub Fg1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "Show1"
.ColComboList(.ColIndex("Show1")) = "..."
End Select
End With
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
  If SystemOptions.AllowCreateHajomraVoucher = True Then
   Ele(11).Enabled = True
 Else
    Ele(11).Enabled = False
 End If
        With Fg1
     If SystemOptions.UserInterface = ArabicInterface Then
        .ColComboList(.ColIndex("PathStus")) = "#1;„› ÊÕ |#2;  „‰›–|#3;„·€Ì "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("PathStus")) = "#1;Open  |#2;No Executed  |#3;Cancel"
            End If
    End With
   
    conection = "select * from  TblExtinAccounts  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
Fill_Combos
   
   
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
   FiLLTXT
ErrTrap:
End Sub


Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim Sql As String
    Dim ID As Double
   RsSavRec.Fields("RecordDateH").value = RecordDateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("OrderID").value = val(TxtOrderID.Text)
   RsSavRec.Fields("Total").value = val(lbl(2).Caption)
   RsSavRec.Fields("CusID").value = val(Me.OutClientID.BoundText)
   RsSavRec.Update
   Cn.Execute "Update tblbookingrequest2 set FlgExAcc=1 where ID=" & val(TxtOrderID.Text) & ""
  ''//////////////////////////
  If Me.TxtModFlg.Text = "E" Then
 Cn.Execute "delete from TblExtinAccountsDer where ExtAccID=" & val(TxtSerial1.Text) & " "
 End If
  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblExtinAccountsDer Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.Fg1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("PathID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ExtAccID").value = val(Me.TxtSerial1.Text)
                RsDevsub("Valuee").value = IIf((.TextMatrix(i, .ColIndex("Valuee"))) = "", Null, val(.TextMatrix(i, .ColIndex("Valuee"))))
                RsDevsub("PathID").value = IIf((.TextMatrix(i, .ColIndex("PathID"))) = "", Null, val(.TextMatrix(i, .ColIndex("PathID"))))
                RsDevsub("Reson").value = IIf((.TextMatrix(i, .ColIndex("Reson"))) = "", Null, (.TextMatrix(i, .ColIndex("Reson"))))
                RsDevsub("PathStus").value = IIf((.TextMatrix(i, .ColIndex("PathStus"))) = "", Null, val(.TextMatrix(i, .ColIndex("PathStus"))))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val(.TextMatrix(i, .ColIndex("Total"))))
                RsDevsub("TotaMorhl").value = IIf((.TextMatrix(i, .ColIndex("TotaMorhl"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotaMorhl"))))
                RsDevsub("Remain").value = IIf((.TextMatrix(i, .ColIndex("Remain"))) = "", Null, val(.TextMatrix(i, .ColIndex("Remain"))))
      RsDevsub.Update
      End If
     Next i
    End With

    FiLLTXT
'''///////////////
   
    Dim Msg As String
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & Chr(13)
                Msg = Msg + " You want to enter another record?"
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
                TxtModFlg = "R"
                FiLLTXT
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
  Sub FullGri()
Dim Rs3 As ADODB.Recordset
Dim i As Integer
Dim Sql As String
    Fg1.Clear flexClearScrollable, flexClearEverything
    Fg1.Rows = 1
 Set Rs3 = New ADODB.Recordset
Sql = " SELECT     dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblExtinAccountsDer.*"
Sql = Sql & " FROM         dbo.TblExtinAccountsDer LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblShrines ON dbo.TblExtinAccountsDer.PathID = dbo.TblShrines.ID"
Sql = Sql & " Where (dbo.TblExtinAccountsDer.ExtAccID = " & val(TxtSerial1.Text) & ")"
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Fg1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
.TextMatrix(i, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(i, .ColIndex("Reson")) = IIf(IsNull(Rs3("Reson").value), "", Rs3("Reson").value)
.TextMatrix(i, .ColIndex("PathStus")) = IIf(IsNull(Rs3("PathStus").value), "", Rs3("PathStus").value)
.TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs3("Total").value), 0, Rs3("Total").value)
.TextMatrix(i, .ColIndex("TotaMorhl")) = IIf(IsNull(Rs3("TotaMorhl").value), 0, Rs3("TotaMorhl").value)
.TextMatrix(i, .ColIndex("Remain")) = IIf(IsNull(Rs3("Remain").value), 0, Rs3("Remain").value)
.TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(Rs3("Valuee").value), 0, Rs3("Valuee").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
 
   On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
    Me.OutClientID.BoundText = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    lbl(2).Caption = IIf(IsNull(RsSavRec.Fields("Total").value), 0, RsSavRec.Fields("Total").value)
    TxtOrderID.Text = IIf(IsNull(RsSavRec.Fields("OrderID").value), 0, RsSavRec.Fields("OrderID").value)
     TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
    Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim total As Double
         Dim i As Integer
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
           If ChekClodePeriod(RecordDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰   ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "This is Period is Closed"
           End If
              Exit Sub
        End If
    If val(DcbBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
    Else
    MsgBox "Please Select Branch"
    End If
    DcbBranch.SetFocus
    Exit Sub
    End If
   If val(OutClientID.BoundText) = 0 Or OutClientID.Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ «„— «·‘€· ·«ÌÊÃœ »Â ⁄„Ì·"
    Else
    MsgBox "Can Not Save .No Found Customer"
    End If
    Exit Sub
    End If
    
With Fg1
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PathID"))) <> 0 Then
If val(.TextMatrix(i, .ColIndex("PathStus"))) <> 3 Then
If (val(.TextMatrix(i, .ColIndex("PathStus"))) = 1 Or val(.TextMatrix(i, .ColIndex("Remain")))) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ  «ﬂœ „‰ «‰Â ·«ÌÊÃœ „”«— „› ÊÕ Ê«ﬂ „«· «· —ÕÌ·"
Else
MsgBox "Save can not be sure that the path is not open"
End If
Exit Sub
End If
End If
End If
Next i
End With
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
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub





Private Sub ISButton5_Click()
print_report
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(RecordDate.value)
End If
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecordDateH.value)
End If
End Sub
Function TbleVehicle(Optional OrderID As Double, Optional PathID As Double, Optional ByRef CuntNo As Integer)
Dim Sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Sql = " SELECT     COUNT(PathID) AS CuntNo"
Sql = Sql & " From dbo.TblDeported"
Sql = Sql & " Where(OrderID = " & OrderID & ") And (PathID = " & PathID & ")"
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
'TbleVehicle = True
CuntNo = IIf(IsNull(Rs3("CuntNo").value), 0, Rs3("CuntNo").value)
Else
CuntNo = 0
'TbleVehicle = False
End If
End Function
Sub GetOperNo(Optional ID As Double)
Dim Sql As String
Dim i As Integer
Dim VehicleNo As Double
Dim CountNo As Integer
  Fg1.Clear flexClearScrollable, flexClearEverything
    Fg1.Rows = 1
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Sql = " SELECT     dbo.TblFlightDetails2.PathID, dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.tblbookingrequest2.OutClientID"
Sql = Sql & " FROM         dbo.tblbookingrequest2 LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblFlightDetails2 ON dbo.tblbookingrequest2.ID = dbo.TblFlightDetails2.HID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblShrines ON dbo.TblFlightDetails2.PathID = dbo.TblShrines.ID"
Sql = Sql & " WHERE     (dbo.TblFlightDetails2.HID = " & ID & ") AND (dbo.tblbookingrequest2.FlgExAcc IS NULL)"
rs1.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs1.RecordCount > 0 Then
With Fg1
rs1.MoveFirst
Fg1.Rows = rs1.RecordCount + 1
RetriveOrderInformation ID, , VehicleNo
'VehicleNo = VehicleNo * Rs1.RecordCount
For i = 1 To .Rows - 1
OutClientID.BoundText = IIf(IsNull(rs1("OutClientID").value), "", rs1("OutClientID").value)
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("Total")) = VehicleNo
.TextMatrix(i, .ColIndex("PathID")) = IIf(IsNull(rs1("PathID").value), 0, rs1("PathID").value)
 TbleVehicle ID, val(.TextMatrix(i, .ColIndex("PathID"))), CountNo
 If CountNo > 0 Then
.TextMatrix(i, .ColIndex("PathStus")) = 2
Else
.TextMatrix(i, .ColIndex("PathStus")) = 1
End If
.TextMatrix(i, .ColIndex("TotaMorhl")) = CountNo
.TextMatrix(i, .ColIndex("Remain")) = val(.TextMatrix(i, .ColIndex("Total"))) - val(.TextMatrix(i, .ColIndex("TotaMorhl")))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs1("Name").value), "", rs1("Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs1("NameE").value), "", rs1("NameE").value)
End If
rs1.MoveNext
Next i
End With
Else
End If
End Sub

Private Sub TxtOrderID_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg.Text = "E" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«Õ–› «·Õ—ﬂ… À„ «œŒ· „‰ ÃœÌœ"
Else
MsgBox "You can notv Edite .Delete Transection"
End If
BtnUndo_Click
Exit Sub
End If
If Me.TxtModFlg.Text = "N" Then

If val(TxtOrderID.Text) <> 0 Then
GetOperNo val(TxtOrderID.Text)
End If
End If
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
    Dim Msg As String
    Dim Sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
           If ChekClodePeriod(RecordDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «·Õ–›  ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "This is Period is Closed"
           End If
              Exit Sub
        End If
    If TxtNoteSerial.Text <> "" Then
       MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
       Exit Sub
   End If
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
          Cn.Execute "Update tblbookingrequest2 set FlgExAcc=null where ID=" & val(TxtOrderID.Text) & ""
         Cn.Execute "Delete from TblExtinAccountsDer where ExtAccID=" & val(TxtSerial1.Text) & " "

                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.Delete
                  Fg1.Clear flexClearScrollable, flexClearEverything
                  Fg1.Rows = 1
                  lbl(2).Caption = 0
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
     End If                       '------------------------------ Move Next ---------------------------.
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
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
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
    'XPDtbTrans.Enabled = True
      '  Frm2.Enabled = True
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
   ' XPDtbTrans.Enabled = False
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
   
  ' XPDtbTrans.Enabled = True
  '     Frm2.Enabled = True
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
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
           If ChekClodePeriod(RecordDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·  ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "This is Period is Closed"
           End If
              Exit Sub
        End If
        
    If TxtSerial1.Text <> "" Then
    If TxtNoteSerial.Text <> "" Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Exit Sub
End If
    Fg1.Rows = Fg1.Rows + 1

        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
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
            Msg = "Sorry.." & Chr(13)
            Msg = Msg & " You can not edit this the record now" & Chr(13)
            Msg = Msg & "It was being edited by another user on the network"
           
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
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    clear_all Me
    TxtModFlg.Text = "N"
    Me.DcbBranch.BoundText = Current_branch
      Fg1.Clear flexClearScrollable, flexClearEverything
     Fg1.Rows = 1
    Me.DCboUserName.BoundText = user_id
    lbl(2).Caption = 0
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

'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)
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
   ' form name
  Label1(2).Caption = "Approve the Nomination"
lbl(4).Caption = "No"
lbl(11).Caption = "Branch"
'ISButton3.Caption = "Attachments"
lbl(25).Caption = "Date"
Label1(0).Caption = "No.Nomination"
Label1(5).Caption = "Companies"
Label1(1).Caption = "No. Student "
Label1(3).Caption = "Remarks"
Label1(14).Caption = "No. Remaining"
Label1(10).Caption = "No. Accepted"
Label1(4).Caption = "Supervisor"
Label1(6).Caption = "No. Accepted"
lbl(0).Caption = "Date"
 
 
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

  
  With Fg1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("selected")) = "Approve"
  .TextMatrix(0, .ColIndex("Name")) = "Student Name"
  .TextMatrix(0, .ColIndex("Supervisor")) = "Supervisor"
  .TextMatrix(0, .ColIndex("QName")) = "Course"
  .TextMatrix(0, .ColIndex("CursNoHour")) = "Number of Hours"
  .TextMatrix(0, .ColIndex("CursValue")) = "Value"
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblExtinAccounts"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

