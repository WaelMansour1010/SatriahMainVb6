VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmStudCandidAccept 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   Icon            =   "FrmStudCandidAccept.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
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
      TabIndex        =   15
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmStudCandidAccept.frx":6852
      Left            =   15480
      List            =   "FrmStudCandidAccept.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   16
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
      TabIndex        =   17
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
            Picture         =   "FrmStudCandidAccept.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudCandidAccept.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   18
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
      ButtonImage     =   "FrmStudCandidAccept.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   20
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
      ButtonImage     =   "FrmStudCandidAccept.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   21
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
      ButtonImage     =   "FrmStudCandidAccept.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8190
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      _cx             =   23627
      _cy             =   14446
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
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   540
         Left            =   0
         TabIndex        =   23
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
         Begin VB.TextBox TxtContNoID 
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
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox TxtCandidacyID 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   120
            Width           =   1260
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10605
            TabIndex        =   0
            Top             =   120
            Width           =   1725
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   6600
            TabIndex        =   2
            Top             =   120
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   8055
            TabIndex        =   1
            Top             =   120
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   2520
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «· —‘ÌÕ"
            Height          =   195
            Index           =   0
            Left            =   1290
            TabIndex        =   55
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   11
            Left            =   5430
            TabIndex        =   51
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   25
            Left            =   9435
            TabIndex        =   50
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   255
            Index           =   4
            Left            =   12510
            TabIndex        =   24
            Top             =   120
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   7065
         Width           =   13365
         _cx             =   23574
         _cy             =   1958
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
            TabIndex        =   26
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
            ButtonImage     =   "FrmStudCandidAccept.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   10485
            TabIndex        =   27
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
            ButtonImage     =   "FrmStudCandidAccept.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9420
            TabIndex        =   10
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
            ButtonImage     =   "FrmStudCandidAccept.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8220
            TabIndex        =   28
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
            ButtonImage     =   "FrmStudCandidAccept.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7230
            TabIndex        =   29
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
            ButtonImage     =   "FrmStudCandidAccept.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   3510
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… Œÿ«» «· —‘ÌÕ"
            Top             =   600
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… Œÿ«» «· —‘ÌÕ"
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
            ButtonImage     =   "FrmStudCandidAccept.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   5760
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
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
            ButtonImage     =   "FrmStudCandidAccept.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   32
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
            ButtonImage     =   "FrmStudCandidAccept.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   33
            Top             =   90
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   330
            Left            =   5760
            TabIndex        =   48
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1665
            _ExtentX        =   2937
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
            ButtonImage     =   "FrmStudCandidAccept.frx":2A8D1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   420
            Left            =   1320
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   600
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··»—Ìœ «·«·ﬂ —Ê‰Ì"
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
            ButtonImage     =   "FrmStudCandidAccept.frx":31133
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   315
            TabIndex        =   38
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   345
            Index           =   14
            Left            =   12270
            TabIndex        =   34
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   42
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
            TabIndex        =   43
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
            ButtonImage     =   "FrmStudCandidAccept.frx":37995
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   675
            TabIndex        =   44
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
            ButtonImage     =   "FrmStudCandidAccept.frx":37D2F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1350
            TabIndex        =   45
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
            ButtonImage     =   "FrmStudCandidAccept.frx":380C9
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1950
            TabIndex        =   46
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
            ButtonImage     =   "FrmStudCandidAccept.frx":38463
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12375
            Picture         =   "FrmStudCandidAccept.frx":387FD
            Stretch         =   -1  'True
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ê«›ﬁ… ⁄·Ï «„—  —‘ÌÕ"
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
            Left            =   5655
            TabIndex        =   47
            Top             =   240
            Width           =   4665
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1170
         Left            =   0
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1320
         Width           =   13455
         _cx             =   23733
         _cy             =   2064
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
         Begin VB.TextBox TxtNoStudAccept 
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
            Height          =   330
            Left            =   120
            MaxLength       =   50
            TabIndex        =   68
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TxtCode 
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
            Height          =   345
            Left            =   10110
            MaxLength       =   50
            TabIndex        =   5
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox TxtRemarks 
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
            Height          =   375
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   630
            Width           =   4020
         End
         Begin VB.TextBox TxtNoStudRemain 
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
            Height          =   345
            Left            =   10110
            MaxLength       =   50
            TabIndex        =   8
            Top             =   630
            Width           =   1500
         End
         Begin VB.TextBox TxtNoStudCon 
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
            Height          =   345
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   7
            Top             =   180
            Width           =   1260
         End
         Begin MSDataListLib.DataCombo DcbCompany 
            Height          =   315
            Left            =   5760
            TabIndex        =   6
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   180
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal AcceptDateH 
            Height          =   315
            Left            =   5760
            TabIndex        =   70
            Top             =   600
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker AcceptDate 
            Height          =   315
            Left            =   7155
            TabIndex        =   71
            Top             =   600
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„Ê«›ﬁ…"
            Height          =   285
            Index           =   0
            Left            =   8835
            TabIndex        =   72
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» «·„ﬁ»Ê·Ì‰"
            Height          =   285
            Index           =   10
            Left            =   1170
            TabIndex        =   69
            Top             =   195
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   3
            Left            =   4080
            TabIndex        =   63
            Top             =   630
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ…"
            Height          =   285
            Index           =   5
            Left            =   11520
            TabIndex        =   57
            Top             =   195
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» «·„ »ﬁÌ‰"
            Height          =   285
            Index           =   14
            Left            =   11520
            TabIndex        =   56
            Top             =   630
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» ›Ì «·⁄ﬁœ"
            Height          =   285
            Index           =   1
            Left            =   4080
            TabIndex        =   52
            Top             =   195
            Width           =   1725
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   4455
         Left            =   0
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2550
         Width           =   13455
         _cx             =   23733
         _cy             =   7858
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
         Begin VB.TextBox TxtSuperVisor 
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
            Left            =   3345
            TabIndex        =   64
            Top             =   120
            Width           =   7305
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   3525
            Left            =   120
            TabIndex        =   54
            Top             =   510
            Width           =   13185
            _cx             =   23257
            _cy             =   6218
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmStudCandidAccept.frx":39C02
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
            Height          =   270
            Index           =   0
            Left            =   12360
            TabIndex        =   59
            Top             =   4140
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   476
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
            ButtonImage     =   "FrmStudCandidAccept.frx":39DC9
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   10200
            TabIndex        =   60
            Top             =   4140
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
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
            ButtonImage     =   "FrmStudCandidAccept.frx":3A363
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   360
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   120
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   635
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmStudCandidAccept.frx":3A8FD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   270
            Left            =   11880
            TabIndex        =   73
            Top             =   120
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   " ÕœÌœ «·ﬂ·"
            ForeColor       =   8388608
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‘—›"
            Height          =   285
            Index           =   4
            Left            =   10320
            TabIndex        =   65
            Top             =   120
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   7
            Left            =   4200
            TabIndex        =   62
            Top             =   4140
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» «·„ﬁ»Ê·Ì‰ Õ«·Ì«"
            Height          =   195
            Index           =   6
            Left            =   6000
            TabIndex        =   61
            Top             =   4140
            Width           =   2085
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
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmStudCandidAccept"
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
  
Sub SaveStudent(Optional AcceptCandID As Double, Optional Name As String, Optional SuperVisorName As String, Optional UQama As String, Optional DcbQualiID As Double _
, Optional SexID As Integer, Optional StudentEmail As String, Optional StudentPhone As String, Optional DateBrithH As String, Optional DateBrith As Date _
, Optional StudentAddres As String, Optional Mobile As String, Optional BranchID As Integer = 0)
Dim Rs9 As ADODB.Recordset
Dim Sql As String
Dim Stud As Double
Dim currentcode As String
Set Rs9 = New ADODB.Recordset
If Me.TxtModFlg.Text = "E" Then
Cn.Execute "Delete From TblStudent  where AcceptCandID=" & AcceptCandID & ""
End If
Sql = "Select * from TblStudent where 1=-1"
Rs9.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Stud = CStr(new_id("TblStudent", "id", "", True))
Rs9.AddNew
currentcode = get_coding(Current_branch, "TblStudent", 12, "")

Rs9("ID").value = Stud
Rs9("AcceptCandID").value = AcceptCandID
Rs9("Name").value = Name
Rs9("NameE").value = Name
Rs9("SuperVisorName").value = SuperVisorName
Rs9("TypeContract").value = 1
Rs9("UQama").value = UQama
Rs9("DcbQualiID").value = DcbQualiID
Rs9("code").value = currentcode
Rs9("Fullcode").value = currentcode
Rs9("SexID").value = SexID
Rs9("StudentEmail").value = StudentEmail
Rs9("StudentPhone").value = StudentPhone
Rs9("Mobile").value = Mobile
Rs9("DateBrithH").value = DateBrithH
Rs9("DateBrith").value = DateBrith
Rs9("CompID").value = val(DcbCompany.BoundText)
Rs9("StudentAddres").value = StudentAddres
Rs9("StutsID").value = 0
Rs9("BranchID").value = BranchID

Rs9.Update
Cn.Execute "update  TblStuCandidacyDet set AccptedID=1  where ID=" & AcceptCandID & ""
Cn.Execute "update TblStuCandidacyDet set StuID =" & Stud & " where  id=" & AcceptCandID & ""
RsSavRec.Resync
End Sub

Private Sub AcceptDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         AcceptDateH.value = ToHijriDate(AcceptDate.value)
End If
End Sub

Private Sub AcceptDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 AcceptDate.value = ToGregorianDate(AcceptDateH.value)
End If
End Sub

Private Sub CheckBox1_Click()
Dim i As Integer
If CheckBox1.value = vbChecked Then
With Fg1
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("selected")) = 1
Next i
End With
Else
With Fg1
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("selected")) = 0
Next i
End With
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
Dim i As Integer
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
Case 1
With Fg1
For i = 1 To .Rows - 1
Cn.Execute "update  TblStuCandidacyDet set AccptedID=null  where ID=" & val(.TextMatrix(i, .ColIndex("CandIDDet"))) & " "
 Cn.Execute "Delete From TblStudent  where AcceptCandID=" & val(.TextMatrix(i, .ColIndex("CandIDDet"))) & " "
  Cn.Execute " update  TblStuCandidacyDet set AccptedID=null,StuID=null  where ID=" & val(.TextMatrix(i, .ColIndex("CandIDDet"))) & " "
            Cn.Execute "update TblTrainingRequest set FlagAccept=null where id=" & val((.TextMatrix(i, .ColIndex("TrainingID")))) & ""
Next i
     .Clear flexClearScrollable, flexClearEverything
     .Rows = 1
 End With
End Select
End If
End Sub
Private Sub RemoveGridRow()
    With Me.Fg1

        If .Row <= 0 Then Exit Sub
        Cn.Execute "update  TblStuCandidacyDet set AccptedID=null ,StuID=null where ID=" & val(.TextMatrix(.Row, .ColIndex("CandIDDet"))) & " "
        Cn.Execute "Delete From TblStudent  where AcceptCandID=" & val(.TextMatrix(.Row, .ColIndex("CandIDDet"))) & " "
         Cn.Execute "update TblTrainingRequest set FlagAccept=null where id=" & val((.TextMatrix(.Row, .ColIndex("TrainingID")))) & ""
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Sub ReLineGrid()
Dim i As Integer
Dim Conter As Integer
Conter = 0
Dim cunt As Integer
cunt = 0
With Fg1
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Name")) <> "" Then
Conter = Conter + 1
.TextMatrix(i, .ColIndex("Serial")) = Conter
 If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
 cunt = cunt + 1
 End If
End If
Next i
Label1(7).Caption = cunt
End With
End Sub
Private Sub DcbCompany_Change()
DcbCompany_Click (0)
End Sub
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblStuCandidacyAccept", "ID", "")
       Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Private Sub DcbCompany_Click(Area As Integer)
  If val(DcbCompany.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany.BoundText, EmpCode
    Me.TxtCode.Text = EmpCode
End Sub



Private Sub Fg1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     Dim StrAccountCode As String
     Dim Rs2 As New ADODB.Recordset
    Dim Sql As String
    Dim LngRow As Long
    With Fg1

     Select Case .ColKey(Col)
              Case "Name"
                 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CandIDDet"), False, True)
                .TextMatrix(Row, .ColIndex("CandIDDet")) = StrAccountCode
             Case "QName"
                 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CursID"), False, True)
                .TextMatrix(Row, .ColIndex("CursID")) = StrAccountCode
               Sql = "select * from  TblStudentCurs where id =" & val(.TextMatrix(Row, .ColIndex("CursID"))) & ""
               Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs2.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("CursValue")) = IIf(IsNull(Rs2("Price").value), 0, Rs2("Price").value)
               .TextMatrix(Row, .ColIndex("CursNoHour")) = IIf(IsNull(Rs2("NoHour").value), 0, Rs2("NoHour").value)
               Else
               .TextMatrix(Row, .ColIndex("CursValue")) = 0
               .TextMatrix(Row, .ColIndex("CursNoHour")) = 0
               End If
    End Select
    End With
    ReLineGrid
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "CursNoHour"
Cancel = True
Case "CursValue"
Cancel = True
Case "Name"
Cancel = True

Case "SuperVisor"
.ComboList = ""
End Select
End With
End Sub

Private Sub Fg1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim Rs2 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    With Fg1

        Select Case .ColKey(Col)
 Case "Name"
     StrSQL = " SELECT     Name, ID"
     StrSQL = StrSQL & "    From dbo.TblStuCandidacyDet"
     StrSQL = StrSQL & "  Where (AccptedID Is Null) And (StudCandID = " & val(TxtCandidacyID.Text) & ")"
                Rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = .BuildComboList(Rs2, "Name", "ID")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
  Case "QName"
     StrSQL = "SELECT   *"
     StrSQL = StrSQL & "          From dbo.TblStudentCurs"
                Rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(Rs2, "Name", "ID")
                Else
                StrComboList = .BuildComboList(Rs2, "NameE", "ID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
  End Select
  
  End With
End Sub
Function GetNoAccept() As Double
Dim Sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Sql = " SELECT     ContNoID, SUM(NoStudACNow) AS SmNoStudACNow"
Sql = Sql & " From dbo.TblStuCandidacyAccept"
Sql = Sql & " Where   (ContNoID = " & val(TxtContNoID.Text) & ")"
If Me.TxtModFlg.Text = "E" Then
Sql = Sql & " and  (ID <> " & val(TxtSerial1.Text) & ")"
End If
Sql = Sql & " GROUP BY ContNoID"
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetNoAccept = IIf(IsNull(Rs4("SmNoStudACNow").value), 0, Rs4("SmNoStudACNow").value)
Else
GetNoAccept = 0
End If
End Function
 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblStuCandidacyAccept  "
    conection = conection & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    conection = conection & " Order By ID"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany
   
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

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
Dim BranchID As Integer
Dim Qualid As Double
Dim SexID As Integer
Dim UQama As String
Dim phone As String
Dim Email As String
Dim Address As String
Dim DateBrithH As String
Dim DateBrith As Date
  '  On Error GoTo ErrTrap
    Dim Sql As String
    Dim ID As Double
   RsSavRec.Fields("RecordDateH").value = RecordDateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("CompID").value = val(DcbCompany.BoundText)
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("Remarks").value = TxtRemarks.Text
   RsSavRec.Fields("CandidacyID").value = val(TxtCandidacyID.Text)
   RsSavRec.Fields("NoStudCon").value = val(TxtNoStudCon.Text)
   RsSavRec.Fields("NoStudRemain").value = val(TxtNoStudRemain.Text)
   RsSavRec.Fields("NoStudAccept").value = val(TxtNoStudAccept.Text)
   RsSavRec.Fields("ContNoID").value = val(TxtContNoID.Text)
   'RsSavRec.Fields("NoStusCandid").value = val(TxtNoStusCandid.text)
   RsSavRec.Fields("AcceptDateH").value = AcceptDateH.value
   RsSavRec.Fields("AcceptDate").value = AcceptDate.value
   RsSavRec.Fields("SuperVisor").value = TxtSuperVisor.Text
  RsSavRec.Fields("NoStudACNow").value = val(Label1(7).Caption)
   RsSavRec.Update
  ''//////////////////////////
  If Me.TxtModFlg.Text = "E" Then
 Cn.Execute "delete from TblStuCandidacyAcceptDet where AcceptCandID=" & val(TxtSerial1.Text) & " "
 End If
  Dim RsDevsub As ADODB.Recordset
  Dim Mobile As String
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblStuCandidacyAcceptDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.Fg1
       For i = .FixedRows To .Rows - 1
       If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
       RsDevsub.AddNew
                RsDevsub("AcceptCandID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TrainingID").value = IIf((.TextMatrix(i, .ColIndex("TrainingID"))) = "", Null, val(.TextMatrix(i, .ColIndex("TrainingID"))))
                RsDevsub("ContNoID").value = IIf((.TextMatrix(i, .ColIndex("ContNoID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ContNoID"))))
                RsDevsub("CandIDDet").value = IIf((.TextMatrix(i, .ColIndex("CandIDDet"))) = "", Null, val(.TextMatrix(i, .ColIndex("CandIDDet"))))
                RsDevsub("Name").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, (.TextMatrix(i, .ColIndex("Name"))))
                RsDevsub("CursID").value = IIf((.TextMatrix(i, .ColIndex("CursID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CursID"))))
                RsDevsub("SuperVisor").value = IIf((.TextMatrix(i, .ColIndex("SuperVisor"))) = "", Null, (.TextMatrix(i, .ColIndex("SuperVisor"))))
                RsDevsub("CursValue").value = IIf((.TextMatrix(i, .ColIndex("CursValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("CursValue"))))
                RsDevsub("CursNoHour").value = IIf((.TextMatrix(i, .ColIndex("CursNoHour"))) = "", Null, val(.TextMatrix(i, .ColIndex("CursNoHour"))))
                RsDevsub("UQama").value = IIf((.TextMatrix(i, .ColIndex("UQama"))) = "", Null, (.TextMatrix(i, .ColIndex("UQama"))))
          RsDevsub.Update
          GetTriningStudentInformation val((.TextMatrix(i, .ColIndex("TrainingID")))), Qualid, SexID, UQama, phone, Email, Address, DateBrithH, DateBrith, Mobile, BranchID
          SaveStudent val(.TextMatrix(i, .ColIndex("CandIDDet"))), (.TextMatrix(i, .ColIndex("Name"))), (.TextMatrix(i, .ColIndex("SuperVisor"))), UQama, Qualid, SexID, Email, phone, DateBrithH, DateBrith, Address, Mobile, BranchID
         Cn.Execute "update TblTrainingRequest set FlagAccept=1 where id=" & val((.TextMatrix(i, .ColIndex("TrainingID")))) & ""
          Else
          Cn.Execute "update TblTrainingRequest set FlagAccept=null where id=" & val((.TextMatrix(i, .ColIndex("TrainingID")))) & ""
          Cn.Execute "update  TblStuCandidacyDet set AccptedID=null where ID=" & val((.TextMatrix(i, .ColIndex("CandIDDet")))) & ""
         Cn.Execute "Delete From TblStudent  where AcceptCandID=" & val(.TextMatrix(i, .ColIndex("CandIDDet"))) & ""
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
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Dim Sql As String
    Fg1.Clear flexClearScrollable, flexClearEverything
    Fg1.Rows = 2
Sql = " SELECT     dbo.TblStuCandidacyAcceptDet.AcceptCandID, dbo.TblStuCandidacyAcceptDet.CandIDDet, dbo.TblStuCandidacyAcceptDet.SuperVisor, "
Sql = Sql & "                       dbo.TblStuCandidacyAcceptDet.CursValue, dbo.TblStuCandidacyAcceptDet.CursNoHour, dbo.TblStuCandidacyAcceptDet.CursID, dbo.TblStudentCurs.Name AS QName,"
Sql = Sql & "                       dbo.TblStudentCurs.NameE AS QNameE, dbo.TblStuCandidacyDet.Name ,dbo.TblStuCandidacyAcceptDet.ContNoID,dbo.TblStuCandidacyAcceptDet.TrainingID ,dbo.TblStuCandidacyAcceptDet.UQama"
Sql = Sql & "  FROM         dbo.TblStuCandidacyAcceptDet LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblStuCandidacyDet ON dbo.TblStuCandidacyAcceptDet.CandIDDet = dbo.TblStuCandidacyDet.ID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblStudentCurs ON dbo.TblStuCandidacyAcceptDet.CursID = dbo.TblStudentCurs.ID"
Sql = Sql & "  Where (dbo.TblStuCandidacyAcceptDet.AcceptCandID =" & val(TxtSerial1.Text) & ")"
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Fg1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("selected")) = 1
.TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(Rs3("UQama").value), "", Rs3("UQama").value)
.TextMatrix(i, .ColIndex("TrainingID")) = IIf(IsNull(Rs3("TrainingID").value), 0, Rs3("TrainingID").value)
.TextMatrix(i, .ColIndex("ContNoID")) = IIf(IsNull(Rs3("ContNoID").value), 0, Rs3("ContNoID").value)
.TextMatrix(i, .ColIndex("CandIDDet")) = IIf(IsNull(Rs3("CandIDDet").value), 0, Rs3("CandIDDet").value)
.TextMatrix(i, .ColIndex("SuperVisor")) = IIf(IsNull(Rs3("SuperVisor").value), "", Rs3("SuperVisor").value)
.TextMatrix(i, .ColIndex("CursValue")) = IIf(IsNull(Rs3("CursValue").value), 0, Rs3("CursValue").value)
.TextMatrix(i, .ColIndex("CursNoHour")) = IIf(IsNull(Rs3("CursNoHour").value), 0, Rs3("CursNoHour").value)
.TextMatrix(i, .ColIndex("CursID")) = IIf(IsNull(Rs3("CursID").value), 0, Rs3("CursID").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("QName")) = IIf(IsNull(Rs3("QName").value), "", Rs3("QName").value)
Else
.TextMatrix(i, .ColIndex("QName")) = IIf(IsNull(Rs3("QNameE").value), "", Rs3("QNameE").value)
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
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtContNoID.Text = IIf(IsNull(RsSavRec.Fields("ContNoID").value), 0, RsSavRec.Fields("ContNoID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbCompany.BoundText = IIf(IsNull(RsSavRec.Fields("CompID").value), "", RsSavRec.Fields("CompID").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
   Me.TxtCandidacyID.Text = IIf(IsNull(RsSavRec.Fields("CandidacyID").value), "", RsSavRec.Fields("CandidacyID").value)
    TxtNoStudCon.Text = IIf(IsNull(RsSavRec.Fields("NoStudCon").value), 0, RsSavRec.Fields("NoStudCon").value)
     TxtNoStudAccept.Text = IIf(IsNull(RsSavRec.Fields("NoStudAccept").value), 0, RsSavRec.Fields("NoStudAccept").value)
     Me.TxtNoStudRemain.Text = IIf(IsNull(RsSavRec.Fields("NoStudRemain").value), 0, RsSavRec.Fields("NoStudRemain").value)
    'TxtNoStusCandid = IIf(IsNull(RsSavRec.Fields("NoStusCandid").value), 0, RsSavRec.Fields("NoStusCandid").value)
    Label1(7).Caption = IIf(IsNull(RsSavRec.Fields("NoStudACNow").value), 0, RsSavRec.Fields("NoStudACNow").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    If val(DcbBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
    Else
    MsgBox "Please Select Branch"
    End If
    DcbBranch.SetFocus
    Exit Sub
    End If
    If Round(val(TxtNoStudRemain.Text), 2) < Round(val(Label1(7).Caption), 2) Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ⁄œœ «·ÿ·«» «·„Ê«›ﬁ ⁄·ÌÂ„ «ﬂ»— „‰ «·„ÿ·Ê»Ì‰ "
    Else
    MsgBox "Can not Save .Number of students  large"
    End If
    Exit Sub
    End If
    If val(TxtCandidacyID.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ ÿ·» «· —‘ÌÕ "
    Else
    MsgBox "Please Eneter No. of Nomination"
    End If
    TxtCandidacyID.SetFocus
    Exit Sub
    End If
        Dim currentcode As String

            If TxtSerial1.Text = "" Then
                currentcode = get_coding(Current_branch, "TblStudent", 12, " ")

                   If currentcode = "miniError" Then
                    MsgBox "⁄œœ Œ«‰«   ﬂÊÌœ ﬂÊœ «·ÿ·«» «· Ì ﬁ„  » ÕœÌœ…  ·Â–« «·ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„ "
                    Exit Sub
                End If
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
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub

Private Sub ISButton2_Click()
    If val(TxtCandidacyID.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ ÿ·» «· —‘ÌÕ "
    Else
    MsgBox "Please Eneter No. of Nomination"
    End If
    TxtCandidacyID.SetFocus
    Exit Sub
    End If
  filgrid1
  ReLineGrid
End Sub
Sub filgrid1()
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim i As Integer
Dim k As Integer
Sql = "SELECT     dbo.TblStuCandidacyDet.*"
Sql = Sql & "  FROM         dbo.TblStuCandidacyDet LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblTrainingRequest ON dbo.TblStuCandidacyDet.TrainingID = dbo.TblTrainingRequest.ID"
Sql = Sql & "  Where (dbo.TblStuCandidacyDet.AccptedID Is Null) And (dbo.TblStuCandidacyDet.StudCandID = " & val(TxtCandidacyID.Text) & ") And (dbo.TblTrainingRequest.FlagAccept Is Null)"
If Me.TxtModFlg.Text <> "R" Then
Sql = Sql & " or dbo.TblStuCandidacyDet.ID IN (select CandIDDet from TblStuCandidacyAcceptDet where AcceptCandID=" & val(TxtSerial1.Text) & " )"
End If
Sql = Sql & " order by dbo.TblStuCandidacyDet.name"
Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With Fg1
.Rows = 1
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For i = 1 To .Rows - 1

.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
.TextMatrix(i, .ColIndex("CandIDDet")) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
.TextMatrix(i, .ColIndex("TrainingID")) = IIf(IsNull(Rs2("TrainingID").value), 0, Rs2("TrainingID").value)
.TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(Rs2("UQama").value), "", Rs2("UQama").value)
.TextMatrix(i, .ColIndex("SuperVisor")) = TxtSuperVisor.Text
Rs2.MoveNext
Next i
End With
End If
ReLineGrid
End Sub
Private Sub ISButton3_Click()
            On Error Resume Next
ShowAttachments TxtSerial1.Text, "04092016002"
ErrTrap:
End Sub

Private Sub ISButton8_Click()
FrmSearStudent.inde = 6
Load FrmSearStudent
FrmSearStudent.show vbModal
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

Private Sub TxtCandidacyID_Change()
Dim CompID As Double
Dim NoStud As Double
Dim ContNoID As Double
If Me.TxtModFlg.Text <> "R" Then
If val(Me.TxtCandidacyID.Text) <> 0 Then
GetNominStudentInformation val(Me.TxtCandidacyID.Text), CompID, NoStud, ContNoID
DcbCompany.BoundText = CompID
TxtNoStudCon.Text = NoStud
TxtContNoID.Text = ContNoID
TxtNoStudRemain.Text = val(TxtNoStudCon.Text) - val(TxtNoStudAccept.Text)
TxtNoStudAccept.Text = GetNoAccept()
End If
End If
End Sub

Private Sub TxtCandidacyID_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCandidacyID.Text, 0)
End Sub

Private Sub TxtCandidacyID_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 501
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtCode.Text, EmpID
        DcbCompany.BoundText = EmpID
    End If
End Sub


Private Sub TxtNoStudAccept_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoStudRemain.Text = val(TxtNoStudCon.Text) - val(TxtNoStudAccept.Text)
End If
End Sub

Private Sub TxtNoStudCon_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoStudRemain.Text = val(TxtNoStudCon.Text) - val(TxtNoStudAccept.Text)
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
         Cn.Execute "Delete from TblStuCandidacyAcceptDet where AcceptCandID=" & val(TxtSerial1.Text) & " "
         With Fg1
         For i = 1 To .Rows - 1
         Cn.Execute "update TblTrainingRequest set FlagAccept=null where id=" & val((.TextMatrix(i, .ColIndex("TrainingID")))) & ""
           Cn.Execute "Delete From TblStudent  where AcceptCandID= " & val(.TextMatrix(i, .ColIndex("CandIDDet"))) & " "
           Cn.Execute " update  TblStuCandidacyDet set AccptedID=null,StuID=null  where ID=" & val(.TextMatrix(i, .ColIndex("CandIDDet"))) & " "
          Next i
          End With
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.Delete
                  Fg1.Clear flexClearScrollable, flexClearEverything
                  Fg1.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            Label1(7).Caption = 0
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
    If TxtSerial1.Text <> "" Then
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
 Label1(7).Caption = 0
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
ISButton3.Caption = "Attachments"
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
ISButton2.Caption = "Add"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete All"
lbl(14).Caption = "By"
ISButton4.Caption = "Send an Email"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"

    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    
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
   My_SQL = "TblStuCandidacyAccept"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
