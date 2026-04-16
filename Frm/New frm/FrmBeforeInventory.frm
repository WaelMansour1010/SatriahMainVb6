VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBeforeInventory 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16005
   Icon            =   "FrmBeforeInventory.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10665
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   16005
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
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBeforeInventory.frx":6852
      Left            =   16560
      List            =   "FrmBeforeInventory.frx":6862
      RightToLeft     =   -1  'True
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
      Left            =   16680
      RightToLeft     =   -1  'True
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
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16320
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   16920
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
      Left            =   16560
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
      Left            =   16680
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
            Picture         =   "FrmBeforeInventory.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventory.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   16680
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
      ButtonImage     =   "FrmBeforeInventory.frx":874B
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
      ButtonImage     =   "FrmBeforeInventory.frx":EFAD
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
      ButtonImage     =   "FrmBeforeInventory.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   10665
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   16005
      _cx             =   28231
      _cy             =   18812
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
         Height          =   705
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   0
         Width           =   15975
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   16
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
            ButtonImage     =   "FrmBeforeInventory.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   17
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
            ButtonImage     =   "FrmBeforeInventory.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   18
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
            ButtonImage     =   "FrmBeforeInventory.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   19
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
            ButtonImage     =   "FrmBeforeInventory.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈⁄œ«œ«   Õœ «·ÿ·»"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   5160
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   15120
            Picture         =   "FrmBeforeInventory.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   9555
         Width           =   16005
         _cx             =   28231
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
         Align           =   2
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
            Left            =   14385
            TabIndex        =   22
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   1365
            _ExtentX        =   2408
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
            ButtonImage     =   "FrmBeforeInventory.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   12405
            TabIndex        =   23
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   1380
            _ExtentX        =   2434
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
            ButtonImage     =   "FrmBeforeInventory.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   10605
            TabIndex        =   24
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
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
            ButtonImage     =   "FrmBeforeInventory.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8730
            TabIndex        =   25
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   1620
            _ExtentX        =   2858
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
            ButtonImage     =   "FrmBeforeInventory.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   6840
            TabIndex        =   26
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   1515
            _ExtentX        =   2672
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
            ButtonImage     =   "FrmBeforeInventory.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   5745
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   600
            Width           =   1110
            _ExtentX        =   1958
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
            ButtonImage     =   "FrmBeforeInventory.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1830
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
            Visible         =   0   'False
            Width           =   1050
            _ExtentX        =   1852
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
            ButtonImage     =   "FrmBeforeInventory.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   3780
            TabIndex        =   29
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            ButtonImage     =   "FrmBeforeInventory.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   11655
            TabIndex        =   36
            Top             =   120
            Width           =   2940
            _ExtentX        =   5186
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
            Height          =   210
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   600
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1965
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   255
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   885
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2745
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   8
            Left            =   14925
            TabIndex        =   37
            Top             =   120
            Width           =   1005
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8160
         Left            =   0
         TabIndex        =   30
         Top             =   1410
         Width           =   15975
         _cx             =   28178
         _cy             =   14393
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
         Caption         =   "»Ì«‰«  «”«”Ì…|«·„—›ﬁ« "
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
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   7740
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   15885
            _cx             =   28019
            _cy             =   13653
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   2415
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1200
               Width           =   7245
               _cx             =   12779
               _cy             =   4260
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
               Begin XtremeSuiteControls.CheckBox ChAllStore 
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   48
                  Top             =   120
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ﬂ· «·„Œ«“‰"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbStore 
                  Bindings        =   "FrmBeforeInventory.frx":2CB3E
                  Height          =   315
                  Left            =   960
                  TabIndex        =   49
                  Top             =   120
                  Width           =   4095
                  _ExtentX        =   7223
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
               Begin ImpulseButton.ISButton BtonAdd 
                  Height          =   390
                  Left            =   120
                  TabIndex        =   50
                  Top             =   0
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2CB53
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid FgStore 
                  Height          =   1395
                  Left            =   120
                  TabIndex        =   52
                  Top             =   600
                  Width           =   7065
                  _cx             =   12462
                  _cy             =   2461
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
                  Rows            =   1
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBeforeInventory.frx":2CEED
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   13
                  Left            =   6000
                  TabIndex        =   53
                  Top             =   2040
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2CF81
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   15
                  Left            =   3720
                  TabIndex        =   54
                  Top             =   2040
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2D51B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Œ“‰ „Õœœ"
                  Height          =   285
                  Index           =   17
                  Left            =   4560
                  TabIndex        =   51
                  Top             =   120
                  Width           =   1365
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3375
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   4320
               Width           =   15885
               _cx             =   28019
               _cy             =   5953
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   405
                  Index           =   3
                  Left            =   14400
                  TabIndex        =   42
                  Top             =   2940
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ— "
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2DAB5
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   405
                  Index           =   4
                  Left            =   12945
                  TabIndex        =   43
                  Top             =   2940
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2E04F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid FgItem 
                  Height          =   2775
                  Left            =   120
                  TabIndex        =   75
                  Top             =   120
                  Width           =   15705
                  _cx             =   27702
                  _cy             =   4895
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
                  Rows            =   1
                  Cols            =   23
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBeforeInventory.frx":2E5E9
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   2415
               Left            =   7320
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1200
               Width           =   8565
               _cx             =   15108
               _cy             =   4260
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
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ã„Ê⁄«  „Õœœ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   210
                  Index           =   1
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   120
                  Width           =   3225
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰› „Õœœ  ≈Œ «— «·’‰›"
                  ForeColor       =   &H000000FF&
                  Height          =   210
                  Index           =   2
                  Left            =   6000
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   2040
                  Width           =   2265
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·«’‰«›"
                  ForeColor       =   &H00FF0000&
                  Height          =   210
                  Index           =   0
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1185
               End
               Begin VB.ListBox ListGroupAll 
                  Height          =   1620
                  ItemData        =   "FrmBeforeInventory.frx":2E966
                  Left            =   4440
                  List            =   "FrmBeforeInventory.frx":2E96D
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   360
                  Width           =   3975
               End
               Begin VB.ListBox ListGroupSelected 
                  Height          =   1620
                  ItemData        =   "FrmBeforeInventory.frx":2E97F
                  Left            =   120
                  List            =   "FrmBeforeInventory.frx":2E986
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   360
                  Width           =   3855
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   390
                  Left            =   240
                  TabIndex        =   73
                  Top             =   1920
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2E99D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   375
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   375
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   1440
                  Width           =   495
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   720
                  Width           =   495
               End
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   360
                  Width           =   495
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   690
               Left            =   0
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   3600
               Width           =   15885
               _cx             =   28019
               _cy             =   1217
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
               Begin VB.TextBox TxtQty 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1200
                  TabIndex        =   67
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox TxtCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   66
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1575
               End
               Begin MSDataListLib.DataCombo DcbItem 
                  Bindings        =   "FrmBeforeInventory.frx":2ED37
                  Height          =   315
                  Left            =   6960
                  TabIndex        =   68
                  Top             =   240
                  Width           =   6015
                  _ExtentX        =   10610
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
               Begin MSDataListLib.DataCombo DcbUnitDit 
                  Bindings        =   "FrmBeforeInventory.frx":2ED4C
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1815
                  _ExtentX        =   3201
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
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   390
                  Left            =   240
                  TabIndex        =   74
                  Top             =   120
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
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
                  ButtonImage     =   "FrmBeforeInventory.frx":2ED61
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊÕœ…"
                  Height          =   285
                  Index           =   0
                  Left            =   5160
                  TabIndex        =   72
                  Top             =   240
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·’‰›"
                  Height          =   285
                  Index           =   51
                  Left            =   14160
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ„Ì… Õœ «·ÿ·»"
                  Height          =   285
                  Index           =   49
                  Left            =   2280
                  TabIndex        =   70
                  Top             =   240
                  Width           =   1485
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   1170
               Left            =   0
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   0
               Width           =   15885
               _cx             =   28019
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
               Begin VB.TextBox TxtOperReQst 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   840
                  Width           =   2175
               End
               Begin VB.ComboBox DcbPriodType 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.TextBox TxtSafetyRate 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  TabIndex        =   93
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   2175
               End
               Begin VB.TextBox TxtPriod 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   975
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                  Height          =   1170
                  Left            =   12600
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   3285
                  _cx             =   5794
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
                  Begin VB.OptionButton Auto_Manula 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÌœÊÌ"
                     ForeColor       =   &H00000000&
                     Height          =   210
                     Index           =   0
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   480
                     Value           =   -1  'True
                     Width           =   1185
                  End
                  Begin VB.OptionButton Auto_Manula 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ì"
                     ForeColor       =   &H00000000&
                     Height          =   210
                     Index           =   1
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   480
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
                     ForeColor       =   &H00C00000&
                     Height          =   285
                     Index           =   3
                     Left            =   360
                     TabIndex        =   80
                     Top             =   120
                     Width           =   2565
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic8 
                  Height          =   1170
                  Left            =   8280
                  TabIndex        =   81
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   4125
                  _cx             =   7276
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
                  Begin VB.OptionButton Mont_Day 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÌÊ„Ì"
                     ForeColor       =   &H00000000&
                     Height          =   210
                     Index           =   1
                     Left            =   720
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   600
                     Width           =   825
                  End
                  Begin VB.OptionButton Mont_Day 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—Ì"
                     ForeColor       =   &H00000000&
                     Height          =   210
                     Index           =   0
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   600
                     Value           =   -1  'True
                     Width           =   1185
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "›Ì Õ«·… «·Ì"
                     ForeColor       =   &H00C00000&
                     Height          =   285
                     Index           =   5
                     Left            =   1200
                     TabIndex        =   85
                     Top             =   0
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» „⁄œ· «·«” Â·«ﬂ"
                     ForeColor       =   &H00404040&
                     Height          =   285
                     Index           =   6
                     Left            =   600
                     TabIndex        =   84
                     Top             =   240
                     Width           =   3405
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic10 
                  Height          =   1170
                  Left            =   3960
                  TabIndex        =   86
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   4245
                  _cx             =   7488
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
                  Begin VB.OptionButton TypExpenses 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„’—Ê› ›ﬁÿ"
                     ForeColor       =   &H00000000&
                     Height          =   210
                     Index           =   0
                     Left            =   2160
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   600
                     Value           =   -1  'True
                     Width           =   1785
                  End
                  Begin VB.OptionButton TypExpenses 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„’—Ê›+ «·ÕÃÊ“"
                     ForeColor       =   &H00000000&
                     Height          =   210
                     Index           =   1
                     Left            =   -480
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   600
                     Width           =   2385
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„ Ê”ÿ ··«” Â·«ﬂ ÂÊ"
                     ForeColor       =   &H00C00000&
                     Height          =   285
                     Index           =   7
                     Left            =   240
                     TabIndex        =   89
                     Top             =   120
                     Width           =   3645
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„⁄«„· Õœ «·ÿ·»"
                  ForeColor       =   &H00C00000&
                  Height          =   285
                  Index           =   11
                  Left            =   2400
                  TabIndex        =   96
                  Top             =   840
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„⁄œ· «·«„«‰"
                  ForeColor       =   &H00C00000&
                  Height          =   285
                  Index           =   10
                  Left            =   2400
                  TabIndex        =   92
                  Top             =   480
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·ÿ·»"
                  ForeColor       =   &H00C00000&
                  Height          =   285
                  Index           =   9
                  Left            =   2400
                  TabIndex        =   90
                  Top             =   120
                  Width           =   1365
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   690
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   735
         Width           =   16005
         _cx             =   28231
         _cy             =   1217
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
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   255
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   240
            Width           =   6705
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1965
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9105
            TabIndex        =   44
            Top             =   240
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   94306305
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   2
            Left            =   7065
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   1665
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   11085
            TabIndex        =   45
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   14760
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1080
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
      Left            =   16560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmBeforeInventory"
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

Private Sub Auto_Manula_Click(Index As Integer)
If Me.Auto_Manula(0).value Then
C1Elastic2.Enabled = False
'C1Elastic8.Enabled = False
'C1Elastic10.Enabled = False
Else
C1Elastic10.Enabled = True
C1Elastic8.Enabled = True
C1Elastic2.Enabled = True
End If
End Sub

Private Sub BtonAdd_Click()
RetriveStore
End Sub
Function GetUnitFuctor(Optional ItemID As Double, Optional UnitID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     *"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (unitid = " & UnitID & ") And (ItemID = " & ItemID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs3.RecordCount > 0 Then
GetUnitFuctor = IIf(IsNull(Rs3("UnitFactor").value), 0, Rs3("UnitFactor").value)
 Else
 GetUnitFuctor = 0
 End If
End Function
Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
Case 13
RemoveGridStoreRow
Case 15
RemoveGridAllRowStore
End Select
End Sub
Sub RetriveStore()
Dim ID As Integer
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
Dim i As Integer
If Me.ChAllStore.value = xtpChecked Then
   Set RsDetails = New ADODB.Recordset
StrSQL = " select * from TblStore where 1=1"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Else

     Set RsDetails = New ADODB.Recordset
    StrSQL = " select * from TblStore where StoreID = " & val(Me.DcbStore.BoundText) & ""
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
End If
  If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        With Me.FgStore
        ID = .Rows
       .Rows = .Rows + RsDetails.RecordCount
        For i = ID To .Rows - 1
           .TextMatrix(i, .ColIndex("Ser")) = i
           .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(RsDetails("Code").value), "", RsDetails("Code").value)
           .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(RsDetails("StoreID").value), 0, RsDetails("StoreID").value)
           If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDetails("StoreName").value), "", RsDetails("StoreName").value)
           Else
           .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDetails("StoreNamee").value), "", RsDetails("StoreNamee").value)
           End If
            RsDetails.MoveNext
        Next i
End With
    End If
End Sub

Sub Retrivetitems()
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
  bool = True
  Dim IK As Integer
  For IK = FgStore.FixedRows To FgStore.Rows - 1
       If val(FgStore.TextMatrix(IK, FgStore.ColIndex("StoreID"))) <> 0 Then
  With FgItem
   If XPOptShowType(2).value = True Or Auto_Manula(0).value = True Then
  If val(DcbItem.BoundText) = 0 Then
  MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
  DcbItem.SetFocus
  Exit Sub
  End If
  End If
   
   j = .Rows
 If XPOptShowType(2).value = True Or XPOptShowType(0).value = True Or Auto_Manula(0).value = True Then
Set Rs1 = New ADODB.Recordset
     sql = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, "
     sql = sql & "                  dbo.TblItems.fullcode , dbo.TblItemsUnits.unitid, dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee , dbo.TblItemsUnits.UnitFactor"
     sql = sql & "      FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
     sql = sql & "                  dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
     sql = sql & "                  dbo.TblItems ON dbo.TblItemsUnits.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
     sql = sql & "                  dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"
     
If XPOptShowType(2).value = True Or Auto_Manula(0).value = True Then
     sql = sql & "  Where (dbo.TblItems.ItemID =" & val(Me.DcbItem.BoundText) & ")"
     sql = sql & " and dbo.TblItemsUnits.unitid=" & val(DcbUnitDit.BoundText) & ""
     Else
    sql = sql & " where (dbo.TblItemsUnits.UnitFactor= dbo.GetItemMaxUnitFactor(dbo.TblItems.ItemID)) "
End If
     Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
   If Rs1.RecordCount > 0 Then
 .Rows = .Rows + Rs1.RecordCount
 Rs1.MoveFirst
        For i = j To .Rows - 1
        .TextMatrix(i, .ColIndex("Ser")) = i
        .TextMatrix(i, .ColIndex("StoreID")) = val(FgStore.TextMatrix(IK, FgStore.ColIndex("StoreID")))
         .TextMatrix(i, .ColIndex("StoreName")) = FgStore.TextMatrix(IK, FgStore.ColIndex("StoreName"))
               .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
               .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
               .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
         If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
         Else
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
         End If
         .TextMatrix(i, .ColIndex("OperReQst")) = val(Me.TxtOperReQst.Text)
         If XPOptShowType(0).value = True Then
         .TextMatrix(i, .ColIndex("UnitFactor")) = IIf(IsNull(Rs1("UnitFactor").value), 0, Rs1("UnitFactor").value)
         If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
         Else
          .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
         End If
           .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
           Else
            .TextMatrix(i, .ColIndex("UnitName")) = Me.DcbUnitDit.Text
           .TextMatrix(i, .ColIndex("UnitID")) = val(Me.DcbUnitDit.BoundText)
           .TextMatrix(i, .ColIndex("UnitFactor")) = GetUnitFuctor(val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("UnitID"))))
           End If
           
           .TextMatrix(i, .ColIndex("SafetyRate")) = val(Me.TxtSafetyRate.Text)
           .TextMatrix(i, .ColIndex("Qty")) = val(Me.txtQty.Text)
           .TextMatrix(i, .ColIndex("Priod")) = val(Me.TxtPriod.Text)
           If Mont_Day(1).value = True Then
           .TextMatrix(i, .ColIndex("Mont_Day")) = 2
           Else
           .TextMatrix(i, .ColIndex("Mont_Day")) = 1
           End If
             If val(DcbPriodType.ListIndex) = 1 Then
           .TextMatrix(i, .ColIndex("PriodType")) = 2
           Else
           .TextMatrix(i, .ColIndex("PriodType")) = 1
           End If
           Rs1.MoveNext
        Next i
   
  End If
  End If
      If XPOptShowType(1).value = True Then
          Dim GROUPIDS As String
          For k = 1 To ListGroupSelected.ListCount
          Set Rs1 = New ADODB.Recordset
        '   sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
        GROUPIDS = GetallChilddata(ListGroupSelected.ItemData(k - 1))
        If Len(GROUPIDS) > 2 Then GROUPIDS = mId(GROUPIDS, 2, Len(GROUPIDS))
        Debug.Print GROUPIDS
        If GROUPIDS = "" Then GROUPIDS = ListGroupSelected.ItemData(k - 1)
          sql = " SELECT dbo.TblItems.Fullcode,     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
           sql = sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee ,dbo.TblItemsUnits.UnitFactor"
           sql = sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
           sql = sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
           sql = sql & "            dbo.TblUnites RIGHT OUTER JOIN"
           sql = sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
           
       sql = sql & "  where dbo.TblItems.GroupID IN ( " & GROUPIDS & ") and(dbo.TblItemsUnits.UnitFactor= dbo.GetItemMaxUnitFactor(dbo.TblItems.ItemID)) "
        
       '(GetallChilddata
           Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
 j = .Rows
.Rows = .Rows + Rs1.RecordCount

        For i = j To .Rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
             .TextMatrix(i, .ColIndex("UnitFactor")) = IIf(IsNull(Rs1("UnitFactor").value), 0, Rs1("UnitFactor").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
            Else
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
            End If
            .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
            .TextMatrix(i, .ColIndex("StoreID")) = val(FgStore.TextMatrix(IK, FgStore.ColIndex("StoreID")))
            .TextMatrix(i, .ColIndex("StoreName")) = FgStore.TextMatrix(IK, FgStore.ColIndex("StoreName"))
         
            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
            .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
            .TextMatrix(i, .ColIndex("SafetyRate")) = val(Me.TxtSafetyRate.Text)
            .TextMatrix(i, .ColIndex("Qty")) = val(Me.txtQty.Text)
            .TextMatrix(i, .ColIndex("Priod")) = val(Me.TxtPriod.Text)
            .TextMatrix(i, .ColIndex("OperReQst")) = val(Me.TxtOperReQst.Text)
           If Mont_Day(1).value = True Then
           .TextMatrix(i, .ColIndex("Mont_Day")) = 2
           Else
           .TextMatrix(i, .ColIndex("Mont_Day")) = 1
           End If
             If val(DcbPriodType.ListIndex) = 1 Then
           .TextMatrix(i, .ColIndex("PriodType")) = 2
           Else
           .TextMatrix(i, .ColIndex("PriodType")) = 1
           End If
          Rs1.MoveNext
        Next i

    End If
       
       
         Next k
  End If

   End With
   End If
   Next IK
         DcbItem.Text = ""
        txtCode.Text = ""
        txtQty.Text = ""
   ' ReLineGrid
End Sub

Private Sub DcbItem_Change()
DcbItem_Click (0)
End Sub

Private Sub DcbItem_Click(Area As Integer)
Dim UnitName As String
Dim UnitID As Long
Me.txtCode.Text = GetItemCode(val(Me.DcbItem.BoundText))
Dim Dcombos As New ClsDataCombos
Dcombos.GetItemsUnitsDetai DcbUnitDit, val(DcbItem.BoundText)
   GetDefaultItemUnit val(Me.DcbItem.BoundText), UnitID, UnitName
    DcbUnitDit.Text = UnitName
    DcbUnitDit.BoundText = UnitID
End Sub

Private Sub FgItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Auto_Manula(1).value = True Then
If FgItem.ColKey(Col) = "OperReQst" Then
FgItem.ComboList = ""
Else
Cancel = True
End If
Else
If FgItem.ColKey(Col) = "OperReQst" Then
FgItem.ComboList = ""
End If
Cancel = False
End If
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    With FgItem
    If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("PriodType")) = "#1; ÌÊ„|#2; ‘Â—"
                .ColComboList(.ColIndex("Mont_Day")) = "#1; ‘Â—Ì|#2; ÌÊ„Ì"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               .ColComboList(.ColIndex("Mont_Day")) = "#1;Monthly   |#2;Daily "
               .ColComboList(.ColIndex("PriodType")) = "#1;Day  |#2;Month "
            End If
    End With
    conection = "select * from  TblSettsRequestLimit  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    FillMylist
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetStores Me.DcbStore
    Dcombos.GetUsers Me.DCboUserName
   ' Dcombos.GetItemsUnits Me.DcbUnitDit
    Dcombos.GetItemsNamesupdate Me.DcbItem
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbPriodType
    .Clear
    .AddItem "ÌÊ„"
    .AddItem "‘Â—"
    End With
    Else
   With DcbPriodType
    .Clear
    .AddItem "Day"
    .AddItem "Month"
    End With
    End If
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
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                  StrSQL = "Delete From TblSettsRequestLimitDet Where SetReqLID=" & val(Me.TxtSerial1.Text)
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
   RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("Remarks").value = TxtRemarks.Text
   RsSavRec.Fields("StoreID").value = val(Me.DcbStore.BoundText)
   RsSavRec.Fields("ItemID").value = val(Me.DcbItem.BoundText)
   RsSavRec.Fields("UnitID").value = val(Me.DcbUnitDit.BoundText)
   RsSavRec.Fields("Qty").value = val(Me.txtQty.Text)
   RsSavRec.Fields("SafetyRate").value = val(Me.TxtSafetyRate.Text)
   RsSavRec.Fields("Priod").value = val(Me.TxtPriod.Text)
   RsSavRec.Fields("PriodType").value = val(Me.DcbPriodType.ListIndex)
   If TypExpenses(1).value = True Then
   RsSavRec.Fields("TypExpenses").value = 1
   Else
   RsSavRec.Fields("TypExpenses").value = 0
   End If
   If Mont_Day(1).value = True Then
   RsSavRec.Fields("Mont_Day").value = 1
   Else
   RsSavRec.Fields("Mont_Day").value = 0
   End If
   If Auto_Manula(1).value = True Then
   RsSavRec.Fields("Auto_Manula").value = 1
   Else
   RsSavRec.Fields("Auto_Manula").value = 0
   End If
   If ChAllStore.value = vbChecked Then
   RsSavRec.Fields("AllStore").value = 1
   Else
   RsSavRec.Fields("AllStore").value = 0
   End If
   If XPOptShowType(0).value = True Then
   RsSavRec.Fields("SelectType").value = 0
   ElseIf XPOptShowType(1).value = True Then
   RsSavRec.Fields("SelectType").value = 1
   ElseIf XPOptShowType(1).value = True Then
   RsSavRec.Fields("SelectType").value = 2
   End If
   RsSavRec.Fields("OperReQst").value = val(Me.TxtOperReQst.Text)
   
    RsSavRec.update
    Dim TransType As String
    If TypExpenses(1).value = True Then
    TransType = "27,61,19"
    Else
    TransType = "27,19"
    End If
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSettsRequestLimitDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.FgItem
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("SetReqLID").value = val(Me.TxtSerial1.Text)
                RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val((.TextMatrix(i, .ColIndex("StoreID")))))
                RsDevsub("UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
                RsDevsub("UnitFactor").value = IIf((.TextMatrix(i, .ColIndex("UnitFactor"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitFactor"))))
                RsDevsub("GroupID").value = IIf((.TextMatrix(i, .ColIndex("GroupID"))) = "", Null, val(.TextMatrix(i, .ColIndex("GroupID"))))
                RsDevsub("Mont_Day").value = IIf((.TextMatrix(i, .ColIndex("Mont_Day"))) = "", Null, val(.TextMatrix(i, .ColIndex("Mont_Day"))))
                RsDevsub("Priod").value = IIf((.TextMatrix(i, .ColIndex("Priod"))) = "", Null, val(.TextMatrix(i, .ColIndex("Priod"))))
                RsDevsub("PriodType").value = IIf((.TextMatrix(i, .ColIndex("PriodType"))) = "", Null, val(.TextMatrix(i, .ColIndex("PriodType"))))
                RsDevsub("SafetyRate").value = IIf((.TextMatrix(i, .ColIndex("SafetyRate"))) = "", Null, val(.TextMatrix(i, .ColIndex("SafetyRate"))))
                If Auto_Manula(1).value = True Then
                 If Mont_Day(0).value = True Then
                .TextMatrix(i, .ColIndex("ConsuRateLowQty")) = GetQtyMonthaly(IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", 0, val(.TextMatrix(i, .ColIndex("ItemID")))), TransType)
                 Else
                .TextMatrix(i, .ColIndex("ConsuRateLowQty")) = GetQtyDailay(IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", 0, val(.TextMatrix(i, .ColIndex("ItemID")))), TransType)
                 End If
                .TextMatrix(i, .ColIndex("ConsuRate")) = Round(val(.TextMatrix(i, .ColIndex("ConsuRateLowQty"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
                .TextMatrix(i, .ColIndex("Minimum")) = val(.TextMatrix(i, .ColIndex("ConsuRate"))) * val(.TextMatrix(i, .ColIndex("Priod"))) * val(TxtSafetyRate.Text)
                .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("Minimum"))) + (val(.TextMatrix(i, .ColIndex("ConsuRate"))) * val(.TextMatrix(i, .ColIndex("Priod"))))
                .TextMatrix(i, .ColIndex("Maximum")) = val(.TextMatrix(i, .ColIndex("Qty"))) + val(.TextMatrix(i, .ColIndex("Minimum")))
                 End If
                 If val(.TextMatrix(i, .ColIndex("UnitFactor"))) <> 0 Then
                 RsDevsub("MaxLowQty").value = Round(val(.TextMatrix(i, .ColIndex("Maximum"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
                 RsDevsub("MinLowQty").value = Round(val(.TextMatrix(i, .ColIndex("Minimum"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
                 RsDevsub("ConsuRate").value = Round(val(.TextMatrix(i, .ColIndex("ConsuRateLowQty"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
                 End If
                RsDevsub("ConsuRateLowQty").value = val(.TextMatrix(i, .ColIndex("ConsuRateLowQty")))
                RsDevsub("Minimum").value = val(.TextMatrix(i, .ColIndex("Minimum")))
                RsDevsub("Qty").value = val(.TextMatrix(i, .ColIndex("Qty")))
                RsDevsub("Maximum").value = val(.TextMatrix(i, .ColIndex("Maximum")))
                RsDevsub("OperReQst").value = val(.TextMatrix(i, .ColIndex("OperReQst")))
                
                RsDevsub("Typ").value = 0
       RsDevsub.update
      End If
     Next i
    End With
    
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSettsRequestLimitDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.FgStore
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("StoreID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("SetReqLID").value = val(Me.TxtSerial1.Text)
                RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val((.TextMatrix(i, .ColIndex("StoreID")))))
                RsDevsub("Typ").value = 1
       RsDevsub.update
      End If
     Next i
    End With
      Set RsDevsub = New ADODB.Recordset
   StrSQL = "SELECT     *  from dbo.TblSettsRequestLimitDet Where (1 = -1)"
   RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 0 To ListGroupSelected.ListCount - 1
             RsDevsub.AddNew
             RsDevsub("SetReqLID").value = val(Me.TxtSerial1.Text)
             RsDevsub("GroupID").value = val(ListGroupSelected.ItemData(i))
             RsDevsub("Typ").value = 2
             RsDevsub.update
       Next i
    
        'Cn.CommitTrans
      '  BeginTrans = False
       RsDevsub.Close
    Dim Msg As String
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
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

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbStore.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), "", RsSavRec.Fields("StoreID").value)
    Me.DcbItem.BoundText = IIf(IsNull(RsSavRec.Fields("ItemID").value), "", RsSavRec.Fields("ItemID").value)
    Me.DcbUnitDit.BoundText = IIf(IsNull(RsSavRec.Fields("UnitID").value), "", RsSavRec.Fields("UnitID").value)
    txtQty.Text = IIf(IsNull(RsSavRec.Fields("Qty").value), 0, RsSavRec.Fields("Qty").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtPriod.Text = IIf(IsNull(RsSavRec.Fields("Priod").value), 0, RsSavRec.Fields("Priod").value)
    TxtSafetyRate.Text = IIf(IsNull(RsSavRec.Fields("SafetyRate").value), 0, RsSavRec.Fields("SafetyRate").value)
    DcbPriodType.ListIndex = IIf(IsNull(RsSavRec.Fields("PriodType").value), -1, RsSavRec.Fields("PriodType").value)
    TxtOperReQst.Text = IIf(IsNull(RsSavRec.Fields("OperReQst").value), 0, RsSavRec.Fields("OperReQst").value)
    
     If Not IsNull(RsSavRec.Fields("Auto_Manula").value) Then
    If RsSavRec.Fields("Auto_Manula").value = 1 Then
    Auto_Manula(1).value = True
    Else
    Auto_Manula(0).value = True
    End If
    Else
    Auto_Manula(0).value = True
    End If
    If Not IsNull(RsSavRec.Fields("Mont_Day").value) Then
    If RsSavRec.Fields("Mont_Day").value = 1 Then
    Mont_Day(1).value = True
    Else
    Mont_Day(0).value = True
    End If
    Else
    Mont_Day(0).value = True
    End If
    If Not IsNull(RsSavRec.Fields("TypExpenses").value) Then
    If RsSavRec.Fields("TypExpenses").value = 1 Then
    TypExpenses(1).value = True
    Else
    TypExpenses(0).value = True
    End If
    Else
    TypExpenses(0).value = True
    End If
    
    If Not IsNull(RsSavRec.Fields("AllStore").value) Then
    If RsSavRec.Fields("AllStore").value = True Then
    ChAllStore.value = vbChecked
    Else
    ChAllStore.value = vbUnchecked
    End If
    Else
    ChAllStore.value = vbUnchecked
    End If
    
   If Not IsNull(RsSavRec.Fields("SelectType").value) Then
    If RsSavRec.Fields("SelectType").value = 2 Then
    XPOptShowType(2).value = True
    ElseIf RsSavRec.Fields("SelectType").value = 1 Then
    XPOptShowType(1).value = True
     ElseIf RsSavRec.Fields("SelectType").value = 0 Then
    XPOptShowType(0).value = True
    End If
    End If
    
    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData

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
    Dim Sm As Double

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
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblSettsRequestLimit", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Function GetQtyMonthaly(Optional ItemID As Double, Optional TransType As String) As Double
Dim sql As String
Dim CounQty As Double
Dim i As Integer
Dim SumQty As Double
CounQty = 0
SumQty = 0
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     MONTH(dbo.Transactions.Transaction_Date) as Monthaly , dbo.Transaction_Details.Item_ID, SUM(dbo.Transaction_Details.Quantity) AS SumQty"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & " WHERE     (dbo.Transaction_Details.Item_ID = " & ItemID & ") AND (dbo.Transactions.Transaction_Type in (" & TransType & " ))"
sql = sql & "   GROUP BY MONTH(dbo.Transactions.Transaction_Date), dbo.Transaction_Details.Item_ID"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Rs7.MoveFirst
For i = 1 To Rs7.RecordCount
SumQty = SumQty + IIf(IsNull(Rs7("SumQty").value), 0, Rs7("SumQty").value)
CounQty = CounQty + 1
Rs7.MoveNext
Next i
Else
GetQtyMonthaly = 0
End If
If CounQty > 0 Then
GetQtyMonthaly = Round(SumQty / CounQty, 2)
Else
GetQtyMonthaly = 0
End If
End Function

Function GetQtyDailay(Optional ItemID As Double, Optional TransType As String) As Double
Dim sql As String
Dim CounQty As Double
Dim i As Integer
Dim SumQty As Double
CounQty = 0
SumQty = 0
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID, SUM(dbo.Transaction_Details.Quantity) AS SumQty"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & " WHERE     (dbo.Transaction_Details.Item_ID = " & ItemID & ") AND (dbo.Transactions.Transaction_Type in (" & TransType & " ))"
sql = sql & "   GROUP BY dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Rs7.MoveFirst
For i = 1 To Rs7.RecordCount
SumQty = SumQty + IIf(IsNull(Rs7("SumQty").value), 0, Rs7("SumQty").value)
CounQty = CounQty + 1
Rs7.MoveNext
Next i
Else
GetQtyDailay = 0
End If
If CounQty > 0 Then
GetQtyDailay = Round(SumQty / CounQty, 2)
Else
GetQtyDailay = 0
End If
End Function
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    FgItem.Clear flexClearScrollable, flexClearEverything
    ListGroupSelected.Clear
            FgItem.Rows = 1
sql = " SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.Typ, dbo.TblSettsRequestLimitDet.SetReqLID, dbo.TblSettsRequestLimitDet.ID, "
sql = sql & "                      dbo.TblSettsRequestLimitDet.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblItems.ItemName,"
sql = sql & "                      dbo.TblItems.ItemNamee, dbo.TblSettsRequestLimitDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblSettsRequestLimitDet.StoreID,"
sql = sql & "                      dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.TblSettsRequestLimitDet.Minimum, dbo.TblSettsRequestLimitDet.Maximum,"
sql = sql & "                      dbo.TblSettsRequestLimitDet.SafetyRate, dbo.TblSettsRequestLimitDet.Priod, dbo.TblSettsRequestLimitDet.PriodType, dbo.TblSettsRequestLimitDet.Mont_Day,"
sql = sql & "                      dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblItems.Fullcode, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblSettsRequestLimitDet.ConsuRateLowQty,"
sql = sql & "                      dbo.TblSettsRequestLimitDet.MinLowQty , dbo.TblSettsRequestLimitDet.MaxLowQty,dbo.TblSettsRequestLimitDet.OperReQst"
sql = sql & " FROM         dbo.TblSettsRequestLimitDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStore ON dbo.TblSettsRequestLimitDet.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.TblSettsRequestLimitDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TblSettsRequestLimitDet.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
sql = sql & "                      dbo.Groups ON dbo.TblSettsRequestLimitDet.GroupID = dbo.Groups.GroupID"
sql = sql & " Where (dbo.TblSettsRequestLimitDet.SetReqLID = " & val(TxtSerial1.Text) & ") And (dbo.TblSettsRequestLimitDet.typ = 0)"
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.FgItem
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   
                   .TextMatrix(i, .ColIndex("OperReQst")) = IIf(IsNull(Rs1("OperReQst").value), 0, Rs1("OperReQst").value)
                   .TextMatrix(i, .ColIndex("ConsuRate")) = IIf(IsNull(Rs1("ConsuRate").value), 0, Rs1("ConsuRate").value)
                   .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(Rs1("Qty").value), 0, Rs1("Qty").value)
                   .TextMatrix(i, .ColIndex("UnitFactor")) = IIf(IsNull(Rs1("UnitFactor").value), 0, Rs1("UnitFactor").value)
                   .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), 0, Rs1("GroupID").value)
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
                   .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("UnitID").value), 0, Rs1("UnitID").value)
                   .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs1("StoreID").value), 0, Rs1("StoreID").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreName").value), "", Rs1("StoreName").value)
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                   Else
                   .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreNamee").value), "", Rs1("StoreNamee").value)
                   End If
                   .TextMatrix(i, .ColIndex("Mont_Day")) = IIf(IsNull(Rs1("Mont_Day").value), "", Rs1("Mont_Day").value)
                   .TextMatrix(i, .ColIndex("PriodType")) = IIf(IsNull(Rs1("PriodType").value), "", Rs1("PriodType").value)
                   .TextMatrix(i, .ColIndex("Priod")) = IIf(IsNull(Rs1("Priod").value), 0, Rs1("Priod").value)
                   .TextMatrix(i, .ColIndex("SafetyRate")) = IIf(IsNull(Rs1("SafetyRate").value), 0, Rs1("SafetyRate").value)
                   .TextMatrix(i, .ColIndex("MaxLowQty")) = IIf(IsNull(Rs1("MaxLowQty").value), 0, Rs1("MaxLowQty").value)
                   .TextMatrix(i, .ColIndex("MinLowQty")) = IIf(IsNull(Rs1("MinLowQty").value), 0, Rs1("MinLowQty").value)
                   .TextMatrix(i, .ColIndex("ConsuRateLowQty")) = IIf(IsNull(Rs1("ConsuRateLowQty").value), 0, Rs1("ConsuRateLowQty").value)
                   .TextMatrix(i, .ColIndex("Maximum")) = IIf(IsNull(Rs1("Maximum").value), 0, Rs1("Maximum").value)
                   .TextMatrix(i, .ColIndex("Minimum")) = IIf(IsNull(Rs1("Minimum").value), 0, Rs1("Minimum").value)
                    
                   Rs1.MoveNext
             Next i
End With
   ''//////////////////
       FgStore.Clear flexClearScrollable, flexClearEverything
            FgStore.Rows = 1
sql = "SELECT     dbo.TblSettsRequestLimitDet.Typ, dbo.TblSettsRequestLimitDet.SetReqLID, dbo.TblSettsRequestLimitDet.ID, dbo.TblSettsRequestLimitDet.StoreID, "
sql = sql & "                      dbo.TblStore.STORENAME , dbo.TblStore.StoreNamee, dbo.TblStore.code"
sql = sql & "  FROM         dbo.TblSettsRequestLimitDet LEFT OUTER JOIN"
sql = sql & "                       dbo.TblStore ON dbo.TblSettsRequestLimitDet.StoreID = dbo.TblStore.StoreID"
sql = sql & " Where (dbo.TblSettsRequestLimitDet.SetReqLID = " & val(TxtSerial1.Text) & ") And (dbo.TblSettsRequestLimitDet.typ = 1)"
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     
     With Me.FgStore
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                
                   .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs1("StoreID").value), 0, Rs1("StoreID").value)
                   
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreName").value), "", Rs1("StoreName").value)
                   Else
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreNamee").value), "", Rs1("StoreNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
End With
Dim RsDetails As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblSettsRequestLimitDet.Typ, dbo.TblSettsRequestLimitDet.SetReqLID, dbo.TblSettsRequestLimitDet.ID, dbo.TblSettsRequestLimitDet.GroupID,"
StrSQL = StrSQL & "                       dbo.Groups.GroupName , dbo.Groups.fullcode, dbo.Groups.GroupNamee"
StrSQL = StrSQL & " FROM         dbo.TblSettsRequestLimitDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblSettsRequestLimitDet.GroupID = dbo.Groups.GroupID"
StrSQL = StrSQL & " Where (dbo.TblSettsRequestLimitDet.SetReqLID = " & val(TxtSerial1.Text) & ") And (dbo.TblSettsRequestLimitDet.typ = 2)"
RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 0 To RsDetails.RecordCount - 1
  If SystemOptions.UserInterface = ArabicInterface Then
  ListGroupSelected.AddItem IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value)
  Else
  ListGroupSelected.AddItem IIf(IsNull(RsDetails("GroupNamee").value), "", RsDetails("GroupNamee").value)
  End If
  ListGroupSelected.ItemData(i) = IIf(IsNull(RsDetails("GroupID").value), "", RsDetails("GroupID").value)
   RsDetails.MoveNext
  Next i
  
        Exit Sub
ErrTrap:
    End Sub



Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
If FgStore.Rows = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„Œ“‰ «Ê·«"
Else
MsgBox "Please Select Store"
End If
FgStore.SetFocus
Exit Sub
End If
Retrivetitems
End If
End Sub

Private Sub ISButton3_Click()
If Me.TxtModFlg.Text <> "R" Then
If FgStore.Rows = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„Œ“‰ «Ê·«"
Else
MsgBox "Please Select Store"
End If
FgStore.SetFocus
Exit Sub
End If
If val(DcbUnitDit.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "—ÃÏ «Œ Ì«— «·ÊÕœ…"
Else
MsgBox ""
End If
Exit Sub
End If
Retrivetitems
End If
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub Mont_Day_Click(Index As Integer)
If Me.Mont_Day(0).value = True Then
DcbPriodType.ListIndex = 1
Else
DcbPriodType.ListIndex = 0
End If
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtCode.Text = "" Then
            Me.DcbItem.BoundText = ""
        Else
            Me.DcbItem.BoundText = GetItemID(Trim$(Me.txtCode.Text))
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
    Dim sql As String
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

          StrSQL = "Delete From TblSettsRequestLimitDet Where   SetReqLID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            FgItem.Clear flexClearScrollable, flexClearEverything
            FgItem.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
             
               '''''''''''''''''''''''''''''''

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
    XPDtbTrans.Enabled = True
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
    XPDtbTrans.Enabled = False
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
        TxtModFlg = "E"
        FgItem.Rows = FgItem.Rows + 1
        FgStore.Rows = FgStore.Rows + 1
        Me.DCboUserName.BoundText = user_id
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
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
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
    XPOptShowType_Click (0)
    Auto_Manula(1).value = True
    Auto_Manula_Click (1)
    TypExpenses(0).value = True
    ChAllStore.value = vbUnchecked
    Mont_Day(1).value = True
    Label6_Click
    FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.Rows = 1
     FgStore.Clear flexClearScrollable, flexClearEverything
    FgStore.Rows = 1
    ListGroupSelected.Clear
    Me.DCboUserName.BoundText = user_id
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
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
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
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
  Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    
   Label1(2).Caption = "Settings Request limit"
lbl(4).Caption = "No"
lbl(1).Caption = "Date"
lbl(2).Caption = "Remarks"
lbl(5).Caption = "Auto"
lbl(49).Caption = "Order Qty limit"
ChAllStore.Caption = "All Store"
lbl(17).Caption = "Store"
lbl(51).Caption = "Item"
lbl(0).Caption = "Unit"
C1Tab1.Caption = "Data"
lbl(7).Caption = "Average Consumption"
XPOptShowType(0).RightToLeft = False
XPOptShowType(1).RightToLeft = False
XPOptShowType(2).RightToLeft = False
XPOptShowType(0).Caption = "Select All"
XPOptShowType(1).Caption = "Select Group"
XPOptShowType(2).Caption = "Select Item"
Auto_Manula(0).RightToLeft = False
Auto_Manula(1).RightToLeft = False
Mont_Day(0).RightToLeft = False
Mont_Day(1).RightToLeft = False
TypExpenses(0).RightToLeft = False
TypExpenses(1).RightToLeft = False
TypExpenses(0).Caption = "Only Expenses"
lbl(3).Caption = "Calculation Method"
TypExpenses(1).Caption = "Expense and Reserved"
Mont_Day(0).Caption = "Monthly"
Mont_Day(1).Caption = "Daily"
Auto_Manula(0).Caption = "Manual"
Auto_Manula(1).Caption = "Auto"
lbl(9).Caption = "Period"
lbl(10).Caption = "Safety Rate"
ISButton2.Caption = "Add"
BtonAdd.Caption = "Add"
ISButton3.Caption = "Add"
   Cmd(13).Caption = "Delete"
    Cmd(15).Caption = "Delete All"
    Cmd(3).Caption = "Delete"
    Cmd(4).Caption = "Delete All"
    
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"

    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
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
    lbl(6).Caption = "Average Calculation"
  With Me.FgItem
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
  .TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
  .TextMatrix(0, .ColIndex("ItemName")) = "Item Name "
  .TextMatrix(0, .ColIndex("UnitName")) = "Unit"
  .TextMatrix(0, .ColIndex("ConsuRate")) = "Consumption Rate"
  .TextMatrix(0, .ColIndex("Mont_Day")) = "Average Calculation"
  .TextMatrix(0, .ColIndex("Priod")) = "Priod "
  .TextMatrix(0, .ColIndex("PriodType")) = "Priod Type"
  .TextMatrix(0, .ColIndex("SafetyRate")) = "Safety Rate"
  .TextMatrix(0, .ColIndex("Qty")) = "Order Qty limit"
  .TextMatrix(0, .ColIndex("Minimum")) = "Minimum"
  .TextMatrix(0, .ColIndex("Maximum")) = "Maximum"
  .TextMatrix(0, .ColIndex("OperReQst")) = "Factor"
  End With
  lbl(11).Caption = "Factor"
    With Me.FgStore
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("Code")) = "Code "
  .TextMatrix(0, .ColIndex("StoreName")) = "Store Name "
  End With
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String, Optional Index As Integer = 0)
        
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.TblSettsRequestLimit.ID, dbo.TblSettsRequestLimit.RecordDate, dbo.TblSettsRequestLimit.AllStore, dbo.TblSettsRequestLimit.Remarks, "
MySQL = MySQL & "                      dbo.TblSettsRequestLimit.SelectType, dbo.TblSettsRequestLimit.Qty, dbo.TblSettsRequestLimit.Mont_Day, dbo.TblSettsRequestLimit.Auto_Manula,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimit.PriodType, dbo.TblSettsRequestLimit.TypExpenses, dbo.TblSettsRequestLimit.SafetyRate, dbo.TblSettsRequestLimit.Priod,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet.Typ, dbo.TblSettsRequestLimitDet.Qty AS DetQty, dbo.TblSettsRequestLimitDet.ConsuRate,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet.Mont_Day AS DetMont_Day, dbo.TblSettsRequestLimitDet.PriodType AS DetPriodType, dbo.TblSettsRequestLimitDet.Priod AS DetPriod,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet.SafetyRate AS DetSafetyRate, dbo.TblSettsRequestLimitDet.Maximum, dbo.TblSettsRequestLimitDet.Minimum,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblSettsRequestLimitDet.MaxLowQty, dbo.TblSettsRequestLimitDet.MinLowQty,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.TblSettsRequestLimitDet.StoreID AS DetStoreID, dbo.TblSettsRequestLimitDet.ItemID AS DetItemID,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet.UnitID AS DetUnitID, dbo.TblSettsRequestLimitDet.GroupID, dbo.TblSettsRequestLimit.StoreID, dbo.TblStore.StoreName,"
MySQL = MySQL & "                      dbo.TblStore.StoreNamee, dbo.TblSettsRequestLimit.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
MySQL = MySQL & "                      dbo.TblSettsRequestLimit.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, TblStore_1.StoreName AS DetStoreName,"
MySQL = MySQL & "                      TblStore_1.StoreNamee AS DetStoreNameE, TblStore_1.Code, TblItems_1.ItemName AS DetItemName, TblItems_1.ItemNamee AS DetItemNameE,"
MySQL = MySQL & "                      TblItems_1.Fullcode AS DetFullcode, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.Groups.Fullcode AS GropFullcode,"
MySQL = MySQL & "                      TblUnites_1.UnitName AS DetUnitName, TblUnites_1.UnitNamee AS DetUnitNameE"
MySQL = MySQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUnites RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblSettsRequestLimit ON dbo.TblUnites.UnitID = dbo.TblSettsRequestLimit.UnitID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimit.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblStore TblStore_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblSettsRequestLimitDet LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUnites TblUnites_1 ON dbo.TblSettsRequestLimitDet.UnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Groups ON dbo.TblSettsRequestLimitDet.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems TblItems_1 ON dbo.TblSettsRequestLimitDet.ItemID = TblItems_1.ItemID ON TblStore_1.StoreID = dbo.TblSettsRequestLimitDet.StoreID ON"
MySQL = MySQL & "                      dbo.TblSettsRequestLimit.ID = dbo.TblSettsRequestLimitDet.SetReqLID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblStore ON dbo.TblSettsRequestLimit.StoreID = dbo.TblStore.StoreID"
MySQL = MySQL & " Where (dbo.TblSettsRequestLimit.ID = " & val(TxtSerial1.Text) & ")"

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSettingsRequestlimi.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSettingsRequestlimi.rpt"
        End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
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
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblSettsRequestLimit"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
Private Sub Label8_Click()
Dim GROUPIDS, sql As String
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.XPOptShowType(1).value = True Then
 If ListGroupAll.ListIndex > -1 Then
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)

            End If
            End If
End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
  

  sql = " SELECT * from  Groups where GroupID>1"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label5_Click()

    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If

End Sub
Private Sub Label7_Click()
    Dim i As Integer
    If Me.XPOptShowType(1).value = True Then
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End If
End Sub
Private Sub RemoveGridAllRowStore()
 FgStore.Clear flexClearScrollable, flexClearEverything
            FgStore.Rows = 1
End Sub

Private Sub RemoveGridAllRow()
 FgItem.Clear flexClearScrollable, flexClearEverything
            FgItem.Rows = 1
'    ReLineGrid
End Sub
Private Sub RemoveGridRow()
    With Me.FgItem
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub
Private Sub RemoveGridStoreRow()
    With Me.FgStore
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub


Private Sub XPOptShowType_Click(Index As Integer)
DcbUnitDit.BoundText = 0
DcbItem.BoundText = 0
txtQty.Text = ""
txtCode.Text = ""

If XPOptShowType(1).value = True Or XPOptShowType(0).value = True Then
C1Elastic3.Enabled = False
Else
C1Elastic3.Enabled = True
End If
End Sub
