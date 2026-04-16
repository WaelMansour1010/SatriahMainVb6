VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmApproveRequset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " √ﬂÌœ ÿ·»«  «·ÕÃ“ ··ÕÃ Ê«·⁄„—…"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   Icon            =   "FrmApproveRequset.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
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
      TabIndex        =   5
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmApproveRequset.frx":6852
      Left            =   15480
      List            =   "FrmApproveRequset.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   6
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
      TabIndex        =   7
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
            Picture         =   "FrmApproveRequset.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveRequset.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   8
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
      ButtonImage     =   "FrmApproveRequset.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   10
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
      ButtonImage     =   "FrmApproveRequset.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   11
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
      ButtonImage     =   "FrmApproveRequset.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8100
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      _cx             =   23627
      _cy             =   14288
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
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   28
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   8295
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
            TabIndex        =   14
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
            ButtonImage     =   "FrmApproveRequset.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   10365
            TabIndex        =   15
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   1365
            _ExtentX        =   2408
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
            ButtonImage     =   "FrmApproveRequset.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8940
            TabIndex        =   0
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   1125
            _ExtentX        =   1984
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
            ButtonImage     =   "FrmApproveRequset.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7140
            TabIndex        =   16
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   1515
            _ExtentX        =   2672
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
            ButtonImage     =   "FrmApproveRequset.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5790
            TabIndex        =   17
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "FrmApproveRequset.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   3870
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   741
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
            ButtonImage     =   "FrmApproveRequset.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
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
            ButtonImage     =   "FrmApproveRequset.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   1200
            TabIndex        =   20
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
            ButtonImage     =   "FrmApproveRequset.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   21
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
            TabIndex        =   26
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   30
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
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmApproveRequset.frx":2A8D1
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   675
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmApproveRequset.frx":2AC6B
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1350
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmApproveRequset.frx":2B005
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1950
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmApproveRequset.frx":2B39F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   3480
            TabIndex        =   39
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   65273857
            CurrentDate     =   38784
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12375
            Picture         =   "FrmApproveRequset.frx":2B739
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " √ﬂÌœ ÿ·»«  «·ÕÃ“ ··ÕÃ Ê«·⁄„—…"
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
            TabIndex        =   35
            Top             =   240
            Width           =   4665
         End
      End
      Begin C1SizerLibCtl.C1Elastic CompnyOut 
         Height          =   7395
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   13455
         _cx             =   23733
         _cy             =   13044
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
         Begin VB.TextBox TxtCompnyIn 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6675
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   4605
         End
         Begin VB.TextBox TxtCompnyOut 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   4365
         End
         Begin VB.ComboBox DcbStus 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmApproveRequset.frx":2CB3E
            Left            =   2280
            List            =   "FrmApproveRequset.frx":2CB48
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   600
            Width           =   2235
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   375
            Left            =   6360
            TabIndex        =   49
            Top             =   960
            Visible         =   0   'False
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«ŸÂ«— €Ì— «·„ƒﬂœ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   315
            Left            =   9840
            TabIndex        =   45
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   65273857
            CurrentDate     =   38784
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3510
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   720
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10500
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   720
            Visible         =   0   'False
            Width           =   780
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   5505
            Left            =   120
            TabIndex        =   37
            Top             =   1320
            Width           =   13425
            _cx             =   23680
            _cy             =   9710
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmApproveRequset.frx":2CB58
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
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   315
            Left            =   240
            TabIndex        =   38
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   6960
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   556
            Caption         =   " ‰›Ì–"
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
            ButtonImage     =   "FrmApproveRequset.frx":2CDCC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   2160
            TabIndex        =   42
            Top             =   960
            Visible         =   0   'False
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo InClientID 
            Height          =   315
            Left            =   6675
            TabIndex        =   48
            Top             =   720
            Visible         =   0   'False
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   555
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   600
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   979
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
            ButtonImage     =   "FrmApproveRequset.frx":3362E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   375
            Left            =   3720
            TabIndex        =   51
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   6960
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   661
            Caption         =   "„”Õ"
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
            ButtonImage     =   "FrmApproveRequset.frx":39E90
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   315
            Left            =   6675
            TabIndex        =   55
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   65273857
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·ÕÃ“"
            Height          =   210
            Index           =   3
            Left            =   4860
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   345
            Index           =   2
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —… „‰"
            Height          =   345
            Index           =   1
            Left            =   11760
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… «·”⁄ÊœÌ…"
            Height          =   345
            Index           =   6
            Left            =   11700
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… „‰ «·Œ«—Ã"
            Height          =   210
            Index           =   0
            Left            =   4860
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   1320
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
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmApproveRequset"
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
 Dim II As Long
 Public LngRow As Long
  Public LngCol As Long
Function print_report2(Optional ID As Double)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.TblBookingRequest.ApproveFlag, dbo.TblBookingRequest.ApproveDate, dbo.TblBookingRequest.ApproveTime, dbo.TblBookingRequest.GroupName, "
 MySQL = MySQL & "                      dbo.TblBookingRequest.ModelID, dbo.TblBookingRequest.CreationDate, dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBookingRequest.InClientID,"
 MySQL = MySQL & "                      TblCustemers_2.CusName, TblCustemers_2.CusNamee, TblCustemers_2.Fullcode, dbo.TblBookingRequest.OutClientID, TblCustemers_1.CusName AS OutCusName,"
 MySQL = MySQL & "                      TblCustemers_1.CusNamee AS OutCusNameE, TblCustemers_1.Fullcode AS OutFullcode, dbo.TblBookingRequest.AirPortID, dbo.TblAirport.Name,"
 MySQL = MySQL & "                      dbo.TblAirport.NameE, dbo.TblBookingRequest.AirLineID, dbo.TblAirlines.Name AS AirLineName, dbo.TblAirlines.NameE AS AirLineNameE,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.EmpName, dbo.TblBookingRequest.EmpCode, dbo.TblBookingRequest.EmpMbile, dbo.TblBookingRequest.ArriveDate,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.ArriveTime, dbo.TblBookingRequest.emp, dbo.TblBookingRequest.other, dbo.TblBookingRequest.FlightNo, dbo.TblBookingRequest.UserID,"
 MySQL = MySQL & "                      TblUsers_2.UserName, dbo.TblBookingRequest.UserID2, TblUsers_1.UserName AS UserName2, dbo.TblBookingRequest.ProgrammID,"
 MySQL = MySQL & "                      dbo.TblProgrammTypes.Name AS ProgName, dbo.TblProgrammTypes.NameE AS ProgNameE, dbo.TblBookingRequest.VehicleNo,"
 MySQL = MySQL & "                      dbo.TBLCarTypes.name AS VehicTyname, dbo.TBLCarTypes.namee AS VehicTynameE, dbo.TblFlightDetails.[Date], dbo.TblFlightDetails.[Time],"
 MySQL = MySQL & "                      dbo.TblFlightDetails.Remarks, dbo.TblFlightDetails.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.RemarkApprove, dbo.TblBookingRequest.StusID, dbo.TblBookingRequest.UseFlag, dbo.TblBookingRequest.ReservNo,"
 MySQL = MySQL & "                      dbo.TblCompaniesGroup.Name AS SeasoName, dbo.TblCompaniesGroup.NameE AS SeasoNameE, dbo.TblBookingRequest.SeasonsID,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.HotelMakh, dbo.TblBookingRequest.HotelMadinh, dbo.TblBookingRequest.HotelJaddah, dbo.TblBookingRequest.CusNo,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.VehicleType , dbo.TblBookingRequest.CompnyIn, dbo.TblBookingRequest.CompnyOut"
 MySQL = MySQL & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBookingRequest LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCompaniesGroup ON dbo.TblBookingRequest.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblBookingRequest.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblUsers TblUsers_1 ON dbo.TblBookingRequest.UserID2 = TblUsers_1.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblUsers TblUsers_2 ON dbo.TblBookingRequest.UserID = TblUsers_2.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblAirlines ON dbo.TblBookingRequest.AirLineID = dbo.TblAirlines.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblAirport ON dbo.TblBookingRequest.AirPortID = dbo.TblAirport.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblBookingRequest.OutClientID = TblCustemers_1.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_2 ON dbo.TblBookingRequest.InClientID = TblCustemers_2.CusID ON"
 MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.TblBookingRequest.BranchID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblShrines RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblFlightDetails ON dbo.TblShrines.ID = dbo.TblFlightDetails.PathID ON dbo.TblBookingRequest.ID = dbo.TblFlightDetails.HID"
 MySQL = MySQL & "  Where (dbo.TblBookingRequest.ID = " & ID & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequestOrde.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequestOrde.rpt"
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

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
Sub filgrid1()
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim Period As Integer
Dim StrPeriod As String
Dim I As Integer
Dim k As Integer
Sql = " SELECT dbo.TblBookingRequest.CusNo,     dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.BranchID, dbo.TblBookingRequest.InClientID, TblCustemers_2.CusName, "
Sql = Sql & "                       TblCustemers_2.CusNamee, TblCustemers_2.Fullcode, dbo.TblBookingRequest.OutClientID, TblCustemers_1.CusName AS OutCusName,"
Sql = Sql & "                       TblCustemers_1.CusNamee AS OutCusNameE, TblCustemers_1.Fullcode AS OutFullcode, dbo.TblBookingRequest.RemarkApprove,"
Sql = Sql & "                       dbo.TblBookingRequest.UserID2, dbo.TblBookingRequest.StusID , dbo.TblBookingRequest.ApproveTime, dbo.TblBookingRequest.ApproveDate ,"
Sql = Sql & "   dbo.TblBookingRequest.CompnyOut ,dbo.TblBookingRequest.CompnyIn,dbo.TblBookingRequest.NoteSerial1"
Sql = Sql & "  FROM         dbo.TblBookingRequest LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblCustemers TblCustemers_1 ON dbo.TblBookingRequest.OutClientID = TblCustemers_1.CusID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblCustemers TblCustemers_2 ON dbo.TblBookingRequest.InClientID = TblCustemers_2.CusID"
Sql = Sql & "     where dbo.TblBookingRequest.UseFlag is null"
If val(DcbStus.ListIndex) = 0 Or val(DcbStus.ListIndex) = -1 Then
Sql = Sql & "  and (dbo.TblBookingRequest.StusID is null or   dbo.TblBookingRequest.StusID=3)"
ElseIf val(DcbStus.ListIndex) = 1 Then
Sql = Sql & "  and dbo.TblBookingRequest.StusID =1"
ElseIf val(DcbStus.ListIndex) = 2 Then
Sql = Sql & "  and dbo.TblBookingRequest.StusID =2"
End If
If Not IsNull(FromDate.value) Then
Sql = Sql & "  and dbo.TblBookingRequest.SDate<= " & SQLDate(FromDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
Sql = Sql & "  and dbo.TblBookingRequest.SDate>= " & SQLDate(ToDate.value, True) & ""
End If
If TxtCompnyIn.Text <> "" Then
Sql = Sql & " AND dbo.TblBookingRequest.CompnyIn like '%" & Me.TxtCompnyIn.Text & "%'"
End If
If TxtCompnyOut.Text <> "" Then
Sql = Sql & " AND dbo.TblBookingRequest.CompnyOut like '%" & Me.TxtCompnyOut.Text & "%'"
End If


'If val(OutClientID.BoundText) <> 0 And OutClientID.text <> "" Then
'Sql = Sql & "  and dbo.TblBookingRequest.OutClientID=" & val(OutClientID.BoundText) & ""
'End If
'If val(InClientID.BoundText) <> 0 And InClientID.text <> "" Then
'Sql = Sql & "  and dbo.TblBookingRequest.InClientID=" & val(InClientID.BoundText) & ""
'End If
   Fg1.Clear flexClearScrollable, flexClearEverything
   Fg1.Rows = 1
Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With Fg1
k = .Rows
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For I = k To .Rows - 1
'CusNo
.TextMatrix(I, .ColIndex("CusNo")) = IIf(IsNull(Rs2("CusNo").value), "", Rs2("CusNo").value)
.TextMatrix(I, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs2("NoteSerial1").value), "", Rs2("NoteSerial1").value)
.TextMatrix(I, .ColIndex("ApproveDate")) = IIf(IsNull(Rs2("ApproveDate").value), Date, Rs2("ApproveDate").value)
.TextMatrix(I, .ColIndex("ApproveTime")) = IIf(IsNull(Rs2("ApproveTime").value), Time, Rs2("ApproveTime").value)
.TextMatrix(I, .ColIndex("StatusID")) = IIf(IsNull(Rs2("StusID").value), 3, Rs2("StusID").value)
.TextMatrix(I, .ColIndex("UserID2")) = IIf(IsNull(Rs2("UserID2").value), "", Rs2("UserID2").value)
.TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(Rs2("ID").value), "", Rs2("ID").value)
.TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(Rs2("SDate").value), "", Rs2("SDate").value)
.TextMatrix(I, .ColIndex("RemarkApprove")) = IIf(IsNull(Rs2("RemarkApprove").value), "", Rs2("RemarkApprove").value)
.TextMatrix(I, .ColIndex("Serial")) = I
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("OutCusName")) = IIf(IsNull(Rs2("CompnyOut").value), "", Rs2("CompnyOut").value)
.TextMatrix(I, .ColIndex("CusName")) = IIf(IsNull(Rs2("CompnyIn").value), "", Rs2("CompnyIn").value)
Else
.TextMatrix(I, .ColIndex("CusName")) = IIf(IsNull(Rs2("CompnyIn").value), "", Rs2("CompnyIn").value)
.TextMatrix(I, .ColIndex("OutCusName")) = IIf(IsNull(Rs2("CompnyOut").value), "", Rs2("CompnyOut").value)
End If
Rs2.MoveNext
Next I
End With
End If
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "StatusID"
If .Cell(flexcpChecked, Row, .ColIndex("Slect")) = flexChecked Then
.ComboList = ""
Else
Cancel = True
End If
Case "ApproveTime"
If .Cell(flexcpChecked, Row, .ColIndex("Slect")) = flexChecked Then
.ComboList = ""
Else
Cancel = True
End If
Case "ApproveDate"
If .Cell(flexcpChecked, Row, .ColIndex("Slect")) = flexChecked Then
.ComboList = ""
Else
Cancel = True
End If
End Select
End With
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Fg1
Select Case .ColKey(Col)
Case "Printer1"
If val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0 Then
print_report2 val(.TextMatrix(.Row, .ColIndex("ID")))
End If

Case "Show"
If val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0 Then
FrmBookingRequest.show
       FrmBookingRequest.Retrive (val(.TextMatrix(.Row, .ColIndex("ID"))))
End If


       
 Case "ApproveTime"
        LngRow = Row
        LngCol = Col
        FrmDateOpProject.Index = 29
        Load FrmDateOpProject
        FrmDateOpProject.Index = 29
        FrmDateOpProject.show vbModal
  Case "ApproveDate"
        LngRow = Row
        LngCol = Col
        FrmDateOpProject.Index = 29
        Load FrmDateOpProject
        FrmDateOpProject.Index = 29
        FrmDateOpProject.show vbModal
End Select
End With
End Sub

Private Sub Fg1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "Printer1"
.ColComboList(.ColIndex("Printer1")) = "..."
Case "Show"
.ColComboList(.ColIndex("Show")) = "..."

Case "ApproveTime"
.ColComboList(.ColIndex("ApproveTime")) = "..."
Case "ApproveDate"
.ColComboList(.ColIndex("ApproveDate")) = "..."
End Select
End With
End Sub

Private Sub ISButton4_Click()
clear_all Me
ToDate.value = ""
FromDate.value = ""
CheckBox1.value = vbUnchecked
   Fg1.Clear flexClearScrollable, flexClearEverything
   Fg1.Rows = 1
End Sub

Private Sub OutClientID_Change()
OutClientID_Click (0)
End Sub

Private Sub OutClientID_Click(Area As Integer)
   Dim fullcode As String
    GetCustomersDetail val(OutClientID.BoundText), , fullcode, 1
    Text2.Text = fullcode
End Sub
Private Sub InClientID_Change()
InClientID_Click (0)
End Sub

Private Sub InClientID_Click(Area As Integer)
   Dim fullcode As String
    GetCustomersDetail val(InClientID.BoundText), , fullcode, 1
    Text1.Text = fullcode
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
  
    With Fg1
     If SystemOptions.UserInterface = ArabicInterface Then
        .ColComboList(.ColIndex("StatusID")) = "#1;ÕÃ“ „ƒﬂœ |#2;ÕÃ“ €Ì— „ƒﬂœ|#3;ÃœÌœ "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("StatusID")) = "#1;Confirmed Reservation |#2;No Confirmed Reservation |#3;New"
            End If
    End With
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbStus
    .Clear
    .AddItem "«·ÃœÌœ ›ﬁÿ"
    .AddItem "«·„ƒﬂœ ›ﬁÿ"
    .AddItem "«·€Ì— „ƒﬂœ ›ﬁÿ"
    .AddItem "«·ﬂ·"
    End With
    Else
   With DcbStus
    .Clear
    .AddItem "New "
    .AddItem "Confirmed Reservation"
    .AddItem "No Confirmed Reservation"
    .AddItem "ALL "
    End With
    End If
    CheckBox1.value = vbUnchecked
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
      Dcombos.GetCompany InClientID, 0, 0
   Dcombos.GetCompany OutClientID, 1, 0
   '  Dcombos.GetCustomersSuppliers 2, InClientID
   'Dcombos.GetCustomersSuppliers 1, OutClientID
ToDate.value = ""
FromDate.value = ""
   filgrid1
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
   Me.Refresh
ErrTrap:
End Sub

Private Sub ISButton2_Click()
filgrid1
End Sub


  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
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

Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
  Label1(2).Caption = "Alarms Of Duration Of The Test  "
 ISButton2.Caption = "Update"
 ISButton3.Caption = "Excute"

  With Fg1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("Slect")) = "Select"
  .TextMatrix(0, .ColIndex("Fullcode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("Emp_Name")) = " Employee Name"
  .TextMatrix(0, .ColIndex("Contract_date")) = "Contract Date"
  .TextMatrix(0, .ColIndex("DateH")) = "Contract Date"
  .TextMatrix(0, .ColIndex("test_period_no")) = "Period"
  .TextMatrix(0, .ColIndex("Contract_Enddate")) = "End Period Date"
  .TextMatrix(0, .ColIndex("DateH1")) = "End Period Date"
  .TextMatrix(0, .ColIndex("StatusID")) = "Procedure"

  End With
ErrTrap:
End Sub

Private Sub ISButton3_Click()
Dim I As Integer
Dim Sql As String
With Fg1
For I = 1 To .Rows - 1
If .Cell(flexcpChecked, I, .ColIndex("Slect")) = flexChecked Then
Sql = " Update tblbookingrequest set StusID=" & val(.TextMatrix(I, .ColIndex("StatusID"))) & ""
Sql = Sql & ",   RemarkApprove='" & (.TextMatrix(I, .ColIndex("RemarkApprove"))) & "'"
Sql = Sql & ",   UserID2=" & val(.TextMatrix(I, .ColIndex("UserID2"))) & ""
Sql = Sql & ",   ApproveDate=" & SQLDate((.TextMatrix(I, .ColIndex("ApproveDate"))), True) & ""
Sql = Sql & ",   ApproveTime='" & (.TextMatrix(I, .ColIndex("ApproveTime"))) & "'"
Sql = Sql & " where ID=" & val(.TextMatrix(I, .ColIndex("ID"))) & " "
Cn.Execute Sql
End If
Next I
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «· ÕœÌÀ »‰Ã«Õ"
Else
MsgBox "Update Successfully"
End If
End With

filgrid1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
   If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text1.Text, 2
        InClientID.BoundText = CUSTID
    End If
    InClientID.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.Text, 1
        OutClientID.BoundText = CUSTID
        OutClientID.SetFocus
    End If
End Sub
