VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmEndExtenGroups 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   Icon            =   "FrmEndExtenGroups.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
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
      ItemData        =   "FrmEndExtenGroups.frx":6852
      Left            =   15480
      List            =   "FrmEndExtenGroups.frx":6862
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
            Picture         =   "FrmEndExtenGroups.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEndExtenGroups.frx":83B1
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
      ButtonImage     =   "FrmEndExtenGroups.frx":874B
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
      ButtonImage     =   "FrmEndExtenGroups.frx":EFAD
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
      ButtonImage     =   "FrmEndExtenGroups.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8100
      Left            =   0
      TabIndex        =   14
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
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   15
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
            TabIndex        =   16
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
            ButtonImage     =   "FrmEndExtenGroups.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   10365
            TabIndex        =   17
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
            ButtonImage     =   "FrmEndExtenGroups.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8940
            TabIndex        =   2
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
            ButtonImage     =   "FrmEndExtenGroups.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7140
            TabIndex        =   18
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
            ButtonImage     =   "FrmEndExtenGroups.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5790
            TabIndex        =   19
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
            ButtonImage     =   "FrmEndExtenGroups.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   3870
            TabIndex        =   20
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   120
            TabIndex        =   21
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   1200
            TabIndex        =   22
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   23
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
            TabIndex        =   28
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   32
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2A8D1
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   675
            TabIndex        =   34
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2AC6B
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1350
            TabIndex        =   35
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2B005
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1950
            TabIndex        =   36
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2B39F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12375
            Picture         =   "FrmEndExtenGroups.frx":2B739
            Stretch         =   -1  'True
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«‰Â«¡/  „œÌœ „Ã„Ê⁄« "
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
            TabIndex        =   37
            Top             =   240
            Width           =   4665
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1140
         Left            =   0
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   720
         Width           =   13455
         _cx             =   23733
         _cy             =   2011
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
         Begin VB.TextBox Text1 
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
            Left            =   4080
            TabIndex        =   51
            Top             =   180
            Width           =   1050
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   435
            Left            =   240
            TabIndex        =   1
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   600
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   767
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
            ButtonImage     =   "FrmEndExtenGroups.frx":2CB3E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   7080
            TabIndex        =   0
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   180
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal DataFomH 
            Height          =   315
            Left            =   8760
            TabIndex        =   43
            Top             =   600
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker DataFom 
            Height          =   330
            Left            =   10185
            TabIndex        =   44
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal DateToH 
            Height          =   315
            Left            =   4440
            TabIndex        =   46
            Top             =   600
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker DateTo 
            Height          =   330
            Left            =   5910
            TabIndex        =   47
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   104660993
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbInstrucor 
            Height          =   315
            Left            =   240
            TabIndex        =   52
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   180
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "·„œ—» „⁄Ì‰"
            Height          =   195
            Index           =   0
            Left            =   5040
            TabIndex        =   49
            Top             =   225
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   195
            Index           =   1
            Left            =   7515
            TabIndex        =   48
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   195
            Index           =   0
            Left            =   11595
            TabIndex        =   45
            Top             =   600
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ ÿ·» «· œ—Ì»"
            Height          =   195
            Index           =   3
            Left            =   8400
            TabIndex        =   42
            Top             =   1530
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "·›—⁄ „⁄Ì‰"
            Height          =   195
            Index           =   10
            Left            =   11520
            TabIndex        =   41
            Top             =   225
            Width           =   1725
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   6315
         Left            =   0
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1800
         Width           =   13455
         _cx             =   23733
         _cy             =   11139
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
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   255
            Left            =   12120
            TabIndex        =   53
            Top             =   120
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " ÕœÌœ «·ﬂ·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   5265
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   13425
            _cx             =   23680
            _cy             =   9287
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
            FormatString    =   $"FrmEndExtenGroups.frx":333A0
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
            TabIndex        =   50
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   5880
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
            ButtonImage     =   "FrmEndExtenGroups.frx":33569
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
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
Attribute VB_Name = "FrmEndExtenGroups"
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


Private Sub CheckBox1_Click()
Dim i As Integer
With Fg1
If CheckBox1.value = xtpChecked Then
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Name")) <> "" Then
.Cell(flexcpChecked, i, .ColIndex("Slect")) = flexChecked
End If
Next i
Else
For i = 1 To .Rows - 1
.Cell(flexcpChecked, i, .ColIndex("Slect")) = flexUnchecked
Next i
End If
End With
End Sub

Private Sub DataFom_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(DataFom.value) Then
         DataFomH.value = ToHijriDate(DataFom.value)
         filgrid1
End If
End If
End Sub

Private Sub DataFomH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 DataFom.value = ToGregorianDate(DataFomH.value)
 filgrid1
End If
End Sub

Private Sub DateTo_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(DateTo.value) Then
         DateToH.value = ToHijriDate(DateTo.value)
         filgrid1
End If
End If
End Sub

Private Sub DateToH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 DateTo.value = ToGregorianDate(DateToH.value)
 filgrid1
End If
End Sub
Sub filgrid1()
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim i As Integer
Dim k As Integer
Sql = "SELECT * from dbo.TblStuGroup "
Sql = Sql & "where (FlgEnd IS NULL)"
If Not IsNull(DataFom.value) Then
Sql = Sql & " and EndDate>=" & SQLDate(DataFom.value, True) & ""
End If
If Not IsNull(DateTo.value) Then
Sql = Sql & " and EndDate<=" & SQLDate(DateTo.value, True) & ""
End If
If val(DcbInstrucor.BoundText) <> 0 And DcbInstrucor.Text <> "" Then
Sql = Sql & " and InstrcID =" & val(DcbInstrucor.BoundText) & ""
End If
If val(Me.DcbBranch.BoundText) <> 0 And DcbBranch.Text <> "" Then
Sql = Sql & " and BranchID =" & val(DcbBranch.BoundText) & ""
End If
   Fg1.Clear flexClearScrollable, flexClearEverything
      Fg1.Rows = 1
Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With Fg1
k = .Rows
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs2("EndDateH").value), "", Rs2("EndDateH").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs2("EndDate").value), "", Rs2("EndDate").value)
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs2("Code").value), "", Rs2("Code").value)
.TextMatrix(i, .ColIndex("Proceder")) = 1
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("NameE").value), "", Rs2("NameE").value)
End If
Rs2.MoveNext
Next i
End With
End If
End Sub

Private Sub DcbBranch_Change()
DcbBranch_Click (0)
End Sub

Private Sub DcbBranch_Click(Area As Integer)
filgrid1
End Sub

Private Sub Fg1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Fg1
Select Case .ColKey(Col)
Case "NoDay"
If val(.TextMatrix(Row, .ColIndex("Proceder"))) = 2 Then
.TextMatrix(Row, .ColIndex("NewEndDate")) = DateAdd("d", val(.TextMatrix(Row, .ColIndex("NoDay"))), (.TextMatrix(Row, .ColIndex("RecordDate"))))
.TextMatrix(Row, .ColIndex("NewEndDateH")) = ToHijriDate(.TextMatrix(Row, .ColIndex("NewEndDate")))
End If
Case "Proceder"
If val(.TextMatrix(Row, .ColIndex("Proceder"))) = 1 Then
.TextMatrix(Row, .ColIndex("NoDay")) = ""
.TextMatrix(Row, .ColIndex("NewEndDate")) = ""
.TextMatrix(Row, .ColIndex("NewEndDateH")) = ""
Else
.TextMatrix(Row, .ColIndex("NewEndDate")) = DateAdd("d", val(.TextMatrix(Row, .ColIndex("NoDay"))), (.TextMatrix(Row, .ColIndex("RecordDate"))))
.TextMatrix(Row, .ColIndex("NewEndDateH")) = ToHijriDate(.TextMatrix(Row, .ColIndex("NewEndDate")))
End If
End Select
End With
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "Slect"
.ComboList = ""
Case "NoDay"
If .Cell(flexcpChecked, Row, .ColIndex("Slect")) = flexChecked Then
.ComboList = ""
Else
Cancel = True
End If
Case "NewEndDate"
Cancel = True
Case "FullCode"
Cancel = True
Case "Name"
Cancel = True
Case "RecordDate"
Cancel = True
Case "RecordDateH"
Cancel = True
Case "Proceder"
If .Cell(flexcpChecked, Row, .ColIndex("Slect")) = flexChecked Then
Cancel = False
Else
Cancel = True
End If
End Select

End With
End Sub

Private Sub ISButton3_Click()
Dim Sql As String
Dim i As Integer
With Fg1
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("Slect")) = flexChecked Then
If val(.TextMatrix(i, .ColIndex("Proceder"))) = 2 Then
Sql = "Update TblStuGroup Set EndDate=" & SQLDate(.TextMatrix(i, .ColIndex("NewEndDate")), True) & ""
Sql = Sql & ", EndDateH=" & .TextMatrix(i, .ColIndex("NewEndDateH")) & ""
Sql = Sql & " where id=" & val(.TextMatrix(i, .ColIndex("ID"))) & ""
Cn.Execute Sql
ElseIf val(.TextMatrix(i, .ColIndex("Proceder"))) = 1 Then
Sql = "Update TblStuGroup set FlgEnd =1,UsrtID2=" & user_id & "  ,NEndDate=" & SQLDate(Date, True)
Sql = Sql & ",NEndDateH='" & ToHijriDate(Date) & "'"
Sql = Sql & " where id=" & val(.TextMatrix(i, .ColIndex("ID"))) & ""
Cn.Execute Sql
End If
End If
Next i
End With
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «· ‰›Ì– »‰Ã«Õ"
Else
MsgBox "Excute Saccessfully"
End If
filgrid1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetInstructorCode EmpID, Text1.Text, 1
        DcbInstrucor.BoundText = EmpID
        filgrid1
    End If
End Sub
Private Sub DcbInstrucor_Change()
DcbInstrucor_Click (0)
End Sub

Private Sub DcbInstrucor_Click(Area As Integer)
Dim UQama As String
  If val(DcbInstrucor.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetInstructorCode val(DcbInstrucor.BoundText), EmpCode, 0
    Me.Text1.Text = EmpCode
    filgrid1
End Sub
 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    DateTo.value = Date
    DateTo.value = ""
    DataFom.value = Date
    DataFom.value = ""
    With Fg1
     If SystemOptions.UserInterface = ArabicInterface Then
           .ColComboList(.ColIndex("Proceder")) = "#1;«‰Â«¡ |#2; „œÌœ "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("Proceder")) = "#1;Terminate |#2;Extend "
            End If
    End With
    
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GeInstructor Me.DcbInstrucor
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

' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
 
End Sub


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
  Label1(2).Caption = "Terminate / Extend Groups "
 ISButton2.Caption = "Add"
 ISButton3.Caption = "Excute"
  Label1(10).Caption = "Branch"
  Label1(0).Caption = "Instructor"
  lbl(0).Caption = "From Date"
  lbl(1).Caption = "To Date"
  CheckBox1.RightToLeft = False
  CheckBox1.Caption = "Select All"
  With Fg1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("FullCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Group Name"
  .TextMatrix(0, .ColIndex("RecordDate")) = " Date"
  .TextMatrix(0, .ColIndex("RecordDateH")) = " Date"
  .TextMatrix(0, .ColIndex("Slect")) = "Select"
  .TextMatrix(0, .ColIndex("NoDay")) = "No. of Days"
  .TextMatrix(0, .ColIndex("NewEndDate")) = "New Date"
  .TextMatrix(0, .ColIndex("NewEndDateH")) = "New Date"
  .TextMatrix(0, .ColIndex("Proceder")) = "Procedure"

  End With
ErrTrap:
End Sub
