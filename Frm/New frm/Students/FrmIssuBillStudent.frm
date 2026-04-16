VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmIssuBillStudent 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13290
   Icon            =   "FrmIssuBillStudent.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8850
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13290
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
      TabIndex        =   12
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmIssuBillStudent.frx":6852
      Left            =   16560
      List            =   "FrmIssuBillStudent.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16320
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   16920
      TabIndex        =   13
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
      TabIndex        =   14
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
            Picture         =   "FrmIssuBillStudent.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIssuBillStudent.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   16680
      TabIndex        =   15
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
      ButtonImage     =   "FrmIssuBillStudent.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   17
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
      ButtonImage     =   "FrmIssuBillStudent.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   18
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
      ButtonImage     =   "FrmIssuBillStudent.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8850
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   13290
      _cx             =   23442
      _cy             =   15610
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
         Left            =   -480
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   0
         Width           =   14040
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   23
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
            ButtonImage     =   "FrmIssuBillStudent.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   24
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
            ButtonImage     =   "FrmIssuBillStudent.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   25
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
            ButtonImage     =   "FrmIssuBillStudent.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   26
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
            ButtonImage     =   "FrmIssuBillStudent.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«’œ«— ›Ê« Ì— «·‘—ﬂ« "
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
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   5160
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12840
            Picture         =   "FrmIssuBillStudent.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   7740
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
            Left            =   12135
            TabIndex        =   29
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11100
            TabIndex        =   30
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9780
            TabIndex        =   31
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8625
            TabIndex        =   32
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7560
            TabIndex        =   33
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   330
            Left            =   6465
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   4350
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   5340
            TabIndex        =   36
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmIssuBillStudent.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   9990
            TabIndex        =   42
            Top             =   120
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   555
            Index           =   11
            Left            =   0
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   480
            Width           =   4335
            _cx             =   7646
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
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   120
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   1530
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   120
               Width           =   1665
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   165
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   120
               Width           =   1185
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   240
               Width           =   765
            End
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   8
            Left            =   12300
            TabIndex        =   43
            Top             =   120
            Width           =   1005
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6360
         Left            =   0
         TabIndex        =   37
         Top             =   1410
         Width           =   13320
         _cx             =   23495
         _cy             =   11218
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
            Height          =   5940
            Left            =   45
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   45
            Width           =   13230
            _cx             =   23336
            _cy             =   10478
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
               BackColor       =   &H00FFFFFF&
               Height          =   1365
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   1035
               Width           =   4515
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3375
               Left            =   0
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   2520
               Width           =   13245
               _cx             =   23363
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
                  Left            =   11880
                  TabIndex        =   48
                  Top             =   2940
                  Visible         =   0   'False
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
                  ButtonImage     =   "FrmIssuBillStudent.frx":2CB3E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   405
                  Index           =   4
                  Left            =   10425
                  TabIndex        =   49
                  Top             =   2940
                  Visible         =   0   'False
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
                  ButtonImage     =   "FrmIssuBillStudent.frx":2D0D8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid FgItem 
                  Height          =   3135
                  Left            =   120
                  TabIndex        =   56
                  Top             =   120
                  Width           =   13065
                  _cx             =   23045
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
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   17
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmIssuBillStudent.frx":2D672
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   6
                  Left            =   1680
                  TabIndex        =   66
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   3165
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   5040
                  TabIndex        =   65
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   885
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   2415
               Left            =   4680
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   0
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
               Begin VB.ListBox ListGroupAll 
                  Height          =   1425
                  ItemData        =   "FrmIssuBillStudent.frx":2D8F6
                  Left            =   4440
                  List            =   "FrmIssuBillStudent.frx":2D8FD
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   360
                  Width           =   3975
               End
               Begin VB.ListBox ListGroupSelected 
                  Height          =   1425
                  ItemData        =   "FrmIssuBillStudent.frx":2D90F
                  Left            =   120
                  List            =   "FrmIssuBillStudent.frx":2D916
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   360
                  Width           =   3855
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   390
                  Left            =   240
                  TabIndex        =   7
                  Top             =   1920
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "«’œ«—"
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
                  ButtonImage     =   "FrmIssuBillStudent.frx":2D92D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‘—ﬂ« "
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   59
                  Top             =   120
                  Width           =   870
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
                  Height          =   255
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
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
                  Height          =   255
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   1320
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
                  Height          =   255
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   840
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
                  Height          =   255
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   600
                  Width           =   495
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   585
               Index           =   3
               Left            =   0
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   120
               Width           =   4695
               _cx             =   8281
               _cy             =   1032
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
               Caption         =   " Õœœ «·› —…"
               Align           =   0
               AutoSizeChildren=   7
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
               Begin VB.ComboBox CmbMonth 
                  Height          =   315
                  Left            =   75
                  Style           =   2  'Dropdown List
                  TabIndex        =   69
                  Top             =   180
                  Width           =   1485
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   2355
                  Style           =   2  'Dropdown List
                  TabIndex        =   68
                  Top             =   165
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   195
                  Index           =   9
                  Left            =   1425
                  TabIndex        =   71
                  Top             =   150
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   240
                  Index           =   7
                  Left            =   2955
                  TabIndex        =   70
                  Top             =   150
                  Width           =   1020
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   2
               Left            =   1800
               TabIndex        =   58
               Top             =   720
               Width           =   885
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   690
         Left            =   0
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   735
         Width           =   13500
         _cx             =   23813
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
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10620
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   1635
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   6480
            TabIndex        =   2
            Top             =   240
            Width           =   1350
            _extentx        =   2355
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   7935
            TabIndex        =   1
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   103284737
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   240
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   0
            Left            =   5430
            TabIndex        =   57
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   9405
            TabIndex        =   50
            Top             =   255
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   12345
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   915
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
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmIssuBillStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim strSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim II As Long

Private Sub CboYear_Change()
ConveDate
End Sub
Sub ConveDate()
On Error GoTo errortrap
'Exit Sub
Dim Day1 As Integer
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
If val(CmbMonth.ListIndex) <> -1 And val(CboYear.Text) <> 0 Then
    Dim str As String
 Day1 = day(RecordDate.value)

 If val(CmbMonth.ListIndex + 1) <= 9 Then
    str = Day1 & "/0" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text
    Else
    str = Day1 & "/" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text
    End If
    
RecordDate.value = CDate(str)
RecordDateH.value = ToHijriDate(RecordDate.value)
Exit Sub
errortrap:
     RecordDate.value = MonthLastDay(CDate("01" & Mid(str, 3, 8)))
   RecordDateH.value = ToHijriDate(RecordDate.value)
    End If
    End If
End Sub

Private Sub CboYear_Click()
CboYear_Change
End Sub

Private Sub CmbMonth_Change()
ConveDate
End Sub

Private Sub CmbMonth_Click()
CmbMonth_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
End Select
End Sub


 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "  «’œ«— ›Ê« Ì— «·‘—ﬂ«  —ﬁ„" & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblIssuBillStudent"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.Text)
Notevalue = 0
 notytype = 9056
Notevalue = val(Lbl(6).Caption)
 BranchID = val(DcbBranch.BoundText)
NoteDate = (RecordDate.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              txtnoteid.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                    Else
                                                 If txtnoteid.Text = "" Or TxtNoteSerial.Text = "" Then
                                         CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                              txtnoteid.Text = NoteID
                                                             TxtNoteSerial.Text = NoteSerial
                                                 Else
                                                              sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                              sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                 sql = sql & " where NoteID=" & val(txtnoteid.Text)
                                                                 Cn.Execute sql
                                        
                                                End If
                                       
                                End If

CREATE_VOUCHER_GE val(txtnoteid.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function
Function GetAccountSuperVisor(Optional ByRef ID As Double) As String
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "Select AccountCode from TblContrStudent where id=" & ID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetAccountSuperVisor = IIf(IsNull(Rs4("AccountCode").value), "", Rs4("AccountCode").value)
Else
GetAccountSuperVisor = ""
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
 
 Dim strSQL As String
 
         strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute strSQL, , adExecuteNoRecords
 LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID

  ' StrTempAccountCode = get_account_code_branch(64, my_branch)
  Dim Comm As Double
  Dim EmpID As String
   Dim TypeSuper As Integer
   Dim sql As String
   Dim ContID As Double
   Dim cusid As Long
   Dim CountStud As Double
   Dim valuee As Double
   Dim EmpAccount As String
   Dim AmolatAccountCode As String
   Dim Rs1 As ADODB.Recordset
'   Set rs1 = New ADODB.Recordset
'       Sql = " SELECT     ContNo"
''Sql = Sql & " From dbo.TblIssuBillStudentDet"
'Sql = Sql & " GROUP BY IsuBillID, TypeTrns, ContNo"
'Sql = Sql & " Having (IsuBillID = " & val(TxtSerial1.text) & ") And (TypeTrns = 0)"
'rs1.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If rs1.RecordCount > 0 Then
'      For i = 1 To rs1.RecordCount
'      ContID = IIf(IsNull(rs1("ContNo").value), 0, rs1("ContNo").value)
'      If ContID <> 0 Then
'      CountStud = GetCountStudent(ContID)
'      If CountStud <> 0 Then
'      GetConInformation ContID, CusID, valuee, TypeSuper, Comm, EmpID
'      valuee = valuee * CountStud
With FgItem

For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("CompID"))) <> 0 Then
cusid = val(.TextMatrix(i, .ColIndex("CompID")))
      StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", cusid)
      StrTempDes = "«À»«  «·»œ·«  «·„” Õﬁ… »—ﬁ —ﬁ„   " & TxtSerial1.Text
      StrTempDes = "·⁄ﬁœ —ﬁ„ " & .TextMatrix(i, .ColIndex("Fullcode"))
      valuee = val(.TextMatrix(i, .ColIndex("Price"))) - val(.TextMatrix(i, .ColIndex("MarkTotal")))
      If valuee > 0 Then
             LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & "      Õ”«» «·‘—ﬂ… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
       End If
          StrTempAccountCode = get_account_code_branch(134, my_branch)
          valuee = val(.TextMatrix(i, .ColIndex("MarkTotal")))
      If valuee > 0 Then
             LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & "      Õ”«» „’«—Ì› «· ”ÊÌﬁ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
       End If
       
       valuee = val(.TextMatrix(i, .ColIndex("Price")))
          StrTempAccountCode = get_account_code_branch(132, my_branch)
         
          LngDevNO = LngDevNO + 1
          If valuee > 0 Then
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» „»Ì⁄«  «·„⁄«Âœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
           End If
          ''''////////////////
         valuee = val(.TextMatrix(i, .ColIndex("TotalSup")))
        If valuee <> 0 Then
          AmolatAccountCode = get_account_code_branch(133, my_branch) ''33
                  LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, AmolatAccountCode, valuee, 0, StrTempDes & "      Õ”«» «·⁄„Ê·«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
         If val(.TextMatrix(i, .ColIndex("TypeSuper"))) = 1 Then
         EmpAccount = GetAccountSuperVisor(val(.TextMatrix(i, .ColIndex("ContNo"))))
          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, EmpAccount, valuee, 1, StrTempDes & "    Õ”«» „»Ì⁄«  «·„⁄«Âœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
           Else
         EmpAccount = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("SuperID"))), "Account_code1")
          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, EmpAccount, valuee, 1, StrTempDes & "    Õ”«» „»Ì⁄«  «·„⁄«Âœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
         End If
        End If
       ''/////////
          End If
           

      Next i

  End With
      updateNotesValueAndNobytext (general_noteid)

ErrTrap:
End Function

Sub Relin()
Dim i As Integer
Dim SumVal As Double
SumVal = 0
With FgItem
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
SumVal = SumVal + val(.TextMatrix(i, .ColIndex("Price")))
End If
Next i
End With
Lbl(6).Caption = SumVal
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub FgItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    With FgItem
    If SystemOptions.UserInterface = ArabicInterface Then
              .ColComboList(.ColIndex("MrkExp")) = "#1;  ﬁÌ„…|#2; ‰”»…"
            Else
            .ColComboList(.ColIndex("MrkExp")) = "#1;Value  |#2;Percentage "
     End If
    End With
    YearMonth
    conection = "select * from  TblIssuBillStudent "
     conection = conection & "  where  (BrnchID=0 or BrnchID is null or         BrnchID in(" & Current_branchSql & "))"
    conection = conection & " Order By ID"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    FillMylist
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.DcbBranch
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
                  strSQL = "Delete From TblIssuBillStudentDet Where IsuBillID=" & val(Me.TxtSerial1.Text)
                  Cn.Execute strSQL, , adExecuteNoRecords
                   strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.txtnoteid.Text)
          Cn.Execute strSQL, , adExecuteNoRecords
              End If
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("RecordDateH").value = RecordDateH.value
   RsSavRec.Fields("BrnchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("Remarks").value = TxtRemarks.Text
   RsSavRec.Fields("TotalValue").value = val(Lbl(6).Caption)
   RsSavRec.Fields("YearID").value = val(CboYear.ListIndex)
   RsSavRec.Fields("MonthID").value = val(CmbMonth.ListIndex)
    RsSavRec.Update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    strSQL = "SELECT  *  from TblIssuBillStudentDet Where (1 = -1)"
    RsDevsub.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.FgItem
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("IsuBillID").value = val(Me.TxtSerial1.Text)
                RsDevsub("YearID").value = val(Me.CboYear.ListIndex)
                RsDevsub("MothID").value = val(Me.CmbMonth.ListIndex)
                RsDevsub("TypeSuper").value = IIf((.TextMatrix(i, .ColIndex("TypeSuper"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeSuper"))))
                RsDevsub("CompID").value = IIf((.TextMatrix(i, .ColIndex("CompID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CompID"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val((.TextMatrix(i, .ColIndex("EmpID")))))
                RsDevsub("ContNo").value = IIf((.TextMatrix(i, .ColIndex("ContNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("ContNo"))))
                RsDevsub("Price").value = IIf((.TextMatrix(i, .ColIndex("Price"))) = "", Null, val(.TextMatrix(i, .ColIndex("Price"))))
                RsDevsub("SuperName").value = IIf((.TextMatrix(i, .ColIndex("SuperName"))) = "", Null, (.TextMatrix(i, .ColIndex("SuperName"))))
                RsDevsub("MrkPers").value = IIf((.TextMatrix(i, .ColIndex("MrkPers"))) = "", Null, val(.TextMatrix(i, .ColIndex("MrkPers"))))
                RsDevsub("MrkExp").value = IIf((.TextMatrix(i, .ColIndex("MrkExp"))) = "", Null, val(.TextMatrix(i, .ColIndex("MrkExp"))))
                RsDevsub("MarkTotal").value = IIf((.TextMatrix(i, .ColIndex("MarkTotal"))) = "", Null, val(.TextMatrix(i, .ColIndex("MarkTotal"))))
                RsDevsub("PerstSup").value = IIf((.TextMatrix(i, .ColIndex("PerstSup"))) = "", Null, val(.TextMatrix(i, .ColIndex("PerstSup"))))
                RsDevsub("TotalSup").value = IIf((.TextMatrix(i, .ColIndex("TotalSup"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotalSup"))))
                RsDevsub("SuperID").value = IIf((.TextMatrix(i, .ColIndex("SuperID"))) = "", Null, val(.TextMatrix(i, .ColIndex("SuperID"))))
                RsDevsub("TypeTrns").value = 0
       RsDevsub.Update
      End If
     Next i
    End With

      Set RsDevsub = New ADODB.Recordset
   strSQL = "SELECT     *  from dbo.TblIssuBillStudentDet Where (1 = -1)"
   RsDevsub.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 0 To ListGroupSelected.ListCount - 1
             RsDevsub.AddNew
             RsDevsub("IsuBillID").value = val(Me.TxtSerial1.Text)
             RsDevsub("CompID").value = val(ListGroupSelected.ItemData(i))
             RsDevsub("TypeTrns").value = 1
             RsDevsub.Update
       Next i
       RsDevsub.Close
    Dim Msg As String
    createVoucher
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

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Function check(Optional cont As Double) As Boolean
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = " select * from  TblIssuBillStudentDet where IsuBillID<>" & val(TxtSerial1.Text) & " and YearID=" & val(CboYear.ListIndex) & " and MothID=" & val(CmbMonth.ListIndex) & " and ContNo=" & cont & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
check = True
Else
check = False
End If
End Function
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID").value), "", RsSavRec.Fields("BrnchID").value)
    Me.TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    Lbl(6).Caption = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, RsSavRec.Fields("TotalValue").value)
    txtnoteid.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
    TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
    CboYear.ListIndex = IIf(IsNull(RsSavRec.Fields("YearID").value), -1, RsSavRec.Fields("YearID").value)
    CmbMonth.ListIndex = IIf(IsNull(RsSavRec.Fields("MonthID").value), -1, RsSavRec.Fields("MonthID").value)
    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
Relin
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
    Dim EmpAccount As String
    Dim Account_Code_dynamic As String
    Dim i As Integer
    With FgItem
    For i = 1 To .Rows - 1
    If val(.TextMatrix(i, .ColIndex("TypeSuper"))) = 1 Then
    EmpAccount = GetAccountSuperVisor(val(.TextMatrix(i, .ColIndex("ContNo"))))
   If EmpAccount = "" Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·„‰œÊ» ›Ì «·⁄ﬁœ —ﬁ„" & (.TextMatrix(i, .ColIndex("Fullcode")))
   
   Else
   MsgBox "Please select Account of Supervisor in Contract " & (.TextMatrix(i, .ColIndex("Fullcode")))
   End If
   Exit Sub
   End If
   End If
Next i
   End With
   Account_Code_dynamic = get_account_code_branch(132, my_branch)

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
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄« ", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
               Account_Code_dynamic = get_account_code_branch(133, my_branch)

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
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·⁄„Ê·« ", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
                          Account_Code_dynamic = get_account_code_branch(134, my_branch)

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
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   „’«—Ì› «· ”ÊÌﬁ", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If

    
If val(DcbBranch.BoundText) = 0 Or DcbBranch.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbBranch.SetFocus
Exit Sub
End If
With Me.FgItem
For i = 1 To .Rows - 1
If check(val(.TextMatrix(i, .ColIndex("ContNo")))) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Â–Â «·› —… „”Ã·… „”»ﬁ« ··⁄ﬁœ—ﬁ„"
Msg = Msg & .TextMatrix(i, .ColIndex("Fullcode")) & ""
Msg = Msg & "··‘—ﬂ… "
Msg = Msg & .TextMatrix(i, .ColIndex("CusName")) & ""
MsgBox Msg
Else
MsgBox "Thi is Period Already Exists"
End If
Exit Sub
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
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblIssuBillStudent", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
     TxtSerial1.Text = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Sub GetConInformation(Optional ContID As Double, Optional ByRef CompID As Long, Optional ByRef ContValue As Double, Optional ByRef TypeSuper As Integer, Optional ByRef Comm As Double, Optional ByRef EmpID As String)
Dim sql As String
Dim strID As String
Dim i As Integer
Dim Rs1 As ADODB.Recordset
strID = ""
sql = " SELECT   *"
sql = sql & " From  TblContrStudent"
sql = sql & " where ID=" & ContID & ""
Set Rs1 = New ADODB.Recordset
Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs1.RecordCount > 0 Then
For i = 1 To Rs1.RecordCount
CompID = IIf(IsNull(Rs1("CompID").value), 0, Rs1("CompID").value)
ContValue = IIf(IsNull(Rs1("ContValue").value), 0, Rs1("ContValue").value)
TypeSuper = IIf(IsNull(Rs1("TypeSuper").value), 0, Rs1("TypeSuper").value)
Comm = IIf(IsNull(Rs1("Comm").value), 0, Rs1("Comm").value)
EmpID = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
Next i
Else
EmpID = 0
Comm = 0
TypeSuper = 0
ContValue = 0
CompID = 0
End If
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    FgItem.Clear flexClearScrollable, flexClearEverything
    ListGroupSelected.Clear
            FgItem.Rows = 1
sql = " SELECT     dbo.TblIssuBillStudentDet.IsuBillID, dbo.TblIssuBillStudentDet.TypeTrns, dbo.TblIssuBillStudentDet.CompID, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, "
sql = sql & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblIssuBillStudentDet.Price, dbo.TblIssuBillStudentDet.EmpID, dbo.TblStudent.Name,"
sql = sql & "                      dbo.TblStudent.NameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblIssuBillStudentDet.ContNo, dbo.TblContrStudent.Fullcode AS ContFullcode,"
sql = sql & "                      dbo.TblIssuBillStudentDet.TotalSup, dbo.TblIssuBillStudentDet.PerstSup, dbo.TblIssuBillStudentDet.MarkTotal, dbo.TblIssuBillStudentDet.MrkPers,"
sql = sql & "                      dbo.TblIssuBillStudentDet.MrkExp, dbo.TblIssuBillStudentDet.SuperName, dbo.TblIssuBillStudentDet.SuperID, dbo.TblEmployee.Emp_Name,"
sql = sql & "                      dbo.TblEmployee.Fullcode AS SupFullcode, dbo.TblEmployee.Emp_Namee ,dbo.TblIssuBillStudentDet.TypeSuper"
sql = sql & " FROM         dbo.TblIssuBillStudentDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblIssuBillStudentDet.SuperID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblContrStudent ON dbo.TblIssuBillStudentDet.ContNo = dbo.TblContrStudent.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStudent ON dbo.TblIssuBillStudentDet.EmpID = dbo.TblStudent.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblIssuBillStudentDet.CompID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.TblIssuBillStudentDet.IsuBillID = " & val(TxtSerial1.Text) & ") And (dbo.TblIssuBillStudentDet.TypeTrns = 0)"
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.FgItem
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("TypeSuper")) = IIf(IsNull(Rs1("TypeSuper").value), 0, Rs1("TypeSuper").value)
                   .TextMatrix(i, .ColIndex("CompID")) = IIf(IsNull(Rs1("CompID").value), 0, Rs1("CompID").value)
                   .TextMatrix(i, .ColIndex("ContNo")) = IIf(IsNull(Rs1("ContNo").value), 0, Rs1("ContNo").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("StudCode")) = IIf(IsNull(Rs1("StudFullCode").value), "", Rs1("StudFullCode").value)
                   .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs1("Price").value), 0, Rs1("Price").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("ContFullcode").value), "", Rs1("ContFullcode").value)
                   .TextMatrix(i, .ColIndex("TotalSup")) = IIf(IsNull(Rs1("TotalSup").value), 0, Rs1("TotalSup").value)
                   .TextMatrix(i, .ColIndex("PerstSup")) = IIf(IsNull(Rs1("PerstSup").value), 0, Rs1("PerstSup").value)
                   .TextMatrix(i, .ColIndex("MarkTotal")) = IIf(IsNull(Rs1("MarkTotal").value), 0, Rs1("MarkTotal").value)
                   .TextMatrix(i, .ColIndex("MrkPers")) = IIf(IsNull(Rs1("MrkPers").value), 0, Rs1("MrkPers").value)
                   .TextMatrix(i, .ColIndex("SuperID")) = IIf(IsNull(Rs1("SuperID").value), 0, Rs1("SuperID").value)
                   .TextMatrix(i, .ColIndex("SuperName")) = IIf(IsNull(Rs1("SuperName").value), "", Rs1("SuperName").value)
                   .TextMatrix(i, .ColIndex("MrkExp")) = IIf(IsNull(Rs1("MrkExp").value), "", Rs1("MrkExp").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   
                   .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                    End If

                   Rs1.MoveNext
             Next i
End With
   ''//////////////////

Dim RsDetails As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
strSQL = " SELECT     dbo.TblIssuBillStudentDet.IsuBillID, dbo.TblIssuBillStudentDet.TypeTrns, dbo.TblIssuBillStudentDet.CompID, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, "
strSQL = strSQL & "                      dbo.TblCustemers.CusNamee , dbo.TblCustemers.FullCode"
strSQL = strSQL & " FROM         dbo.TblIssuBillStudentDet LEFT OUTER JOIN"
strSQL = strSQL & "                      dbo.TblCustemers ON dbo.TblIssuBillStudentDet.CompID = dbo.TblCustemers.CusID"
strSQL = strSQL & " Where (dbo.TblIssuBillStudentDet.IsuBillID = " & val(TxtSerial1.Text) & ") And (dbo.TblIssuBillStudentDet.TypeTrns = 1)"

RsDetails.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 0 To RsDetails.RecordCount - 1
  If SystemOptions.UserInterface = ArabicInterface Then
  ListGroupSelected.AddItem IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
  Else
  ListGroupSelected.AddItem IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
  End If
  ListGroupSelected.ItemData(i) = IIf(IsNull(RsDetails("CompID").value), "", RsDetails("CompID").value)
   RsDetails.MoveNext
  Next i
  
        Exit Sub
ErrTrap:
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
 MySQL = " SELECT     dbo.TblIssuBillStudent.ID, dbo.TblIssuBillStudent.RecordDate, dbo.TblIssuBillStudent.RecordDateH, dbo.TblIssuBillStudent.BrnchID, "
 MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblIssuBillStudent.Remarks, dbo.TblIssuBillStudent.NoteID,"
 MySQL = MySQL & "                      dbo.TblIssuBillStudent.NoteSerial, dbo.TblIssuBillStudent.TotalValue, dbo.TblIssuBillStudent.YearID, dbo.TblIssuBillStudent.MonthID,"
 MySQL = MySQL & "                      dbo.TblIssuBillStudentDet.IsuBillID, dbo.TblIssuBillStudentDet.TypeTrns, dbo.TblIssuBillStudentDet.ContNo, dbo.TblIssuBillStudentDet.Price,"
 MySQL = MySQL & "                      dbo.TblIssuBillStudentDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblIssuBillStudentDet.CompID,"
 MySQL = MySQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS Expr1, dbo.TblContrStudent.Fullcode AS ContFullcode,"
 MySQL = MySQL & "                      dbo.TblIssuBillStudentDet.SuperName, dbo.TblIssuBillStudentDet.MrkExp, dbo.TblIssuBillStudentDet.MrkPers, dbo.TblIssuBillStudentDet.MarkTotal,"
 MySQL = MySQL & "                      dbo.TblIssuBillStudentDet.PerstSup, dbo.TblIssuBillStudentDet.TotalSup, dbo.TblIssuBillStudentDet.TypeSuper, dbo.TblIssuBillStudentDet.MothID,"
 MySQL = MySQL & "                      dbo.TblIssuBillStudentDet.YearID AS YearIDDet, dbo.TblStudent.Name, dbo.TblStudent.NameE, dbo.TblStudent.FullCode AS StudFullCode"
 MySQL = MySQL & "    FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblIssuBillStudentDet LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblStudent ON dbo.TblIssuBillStudentDet.EmpID = dbo.TblStudent.ID ON dbo.TblEmployee.Emp_ID = dbo.TblIssuBillStudentDet.SuperID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblContrStudent ON dbo.TblIssuBillStudentDet.ContNo = dbo.TblContrStudent.ID RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblIssuBillStudent ON dbo.TblIssuBillStudentDet.IsuBillID = dbo.TblIssuBillStudent.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblIssuBillStudent.BrnchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblIssuBillStudentDet.CompID = dbo.TblCustemers.CusID"
 MySQL = MySQL & "  Where (dbo.TblIssuBillStudent.ID = " & val(TxtSerial1.Text) & ") and dbo.TblIssuBillStudentDet.TypeTrns=0"
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptIssuBillStudent.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptIssuBillStudent.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
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
Private Sub ISButton2_Click()
    FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.Rows = 1
Fill_Grid
Relin
End Sub
Function GetCountStudent(Optional cont As Double) As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     COUNT(dbo.TblStuCandidacyDet.StuID) AS CuntStude"
sql = sql & " FROM         dbo.TblStudent RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacyDet ON dbo.TblStudent.ID = dbo.TblStuCandidacyDet.StuID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacy ON dbo.TblStuCandidacyDet.StudCandID = dbo.TblStuCandidacy.ID"
sql = sql & " Where (dbo.TblStuCandidacy.ContNoID = " & cont & ")"
sql = sql & " GROUP BY dbo.TblStudent.StutsID"
sql = sql & " Having (dbo.TblStudent.StutsID <> 1)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCountStudent = IIf(IsNull(rs2("CuntStude").value), 0, rs2("CuntStude").value)
Else
GetCountStudent = 0
End If
End Function
Sub GetStudent(Optional cont As Double, Optional ByRef StudID As Double, Optional ByRef StudentName As String, Optional ByRef fullcode As String)
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = " SELECT     dbo.TblStuCandidacyDet.StuID, dbo.TblStudent.Name, dbo.TblStudent.NameE, dbo.TblStudent.FullCode , dbo.TblStudent.StutsID"
sql = sql & " FROM         dbo.TblStudent RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacyDet ON dbo.TblStudent.ID = dbo.TblStuCandidacyDet.StuID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacy ON dbo.TblStuCandidacyDet.StudCandID = dbo.TblStuCandidacy.ID"
sql = sql & " Where (dbo.TblStuCandidacyDet.AccptedID = 1) And (dbo.TblStuCandidacy.ContNoID = " & cont & ") and  dbo.TblStudent.StutsID<>1"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
StudID = IIf(IsNull(Rs5("StuID").value), 0, Rs5("StuID").value)
fullcode = IIf(IsNull(Rs5("FullCode").value), "", Rs5("FullCode").value)
If SystemOptions.UserInterface = ArabicInterface Then
StudentName = IIf(IsNull(Rs5("Name").value), "", Rs5("Name").value)
Else
StudentName = IIf(IsNull(Rs5("NameE").value), "", Rs5("NameE").value)
End If
Else
fullcode = ""
StudID = 0
StudentName = ""
End If
End Sub
Sub Fill_Grid()
Dim sql As String
Dim k As Integer
Dim fullcode As String
Dim name As String
Dim StudID As Double
Dim NoStud As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim Compnies As String
Compnies = ""
    For k = 1 To ListGroupSelected.ListCount
        Compnies = Compnies & ListGroupSelected.ItemData(k - 1)
        Compnies = Compnies & ","
        Next k
         Compnies = Compnies & 0
sql = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CompFullcode, "
sql = sql & "                       QryGetStudent.FullCode AS StudFullCode, QryGetStudent.NameE, QryGetStudent.Name, QryGetStudent.StuID, dbo.TblEmployee.Emp_Name,"
sql = sql & "                      dbo.TblEmployee.Emp_Namee, dbo.TblStudent.StutsID, dbo.TblContrStudent.*"
sql = sql & " FROM         dbo.TblStudent RIGHT OUTER JOIN"
sql = sql & "                      dbo.QryGetStudent() QryGetStudent ON dbo.TblStudent.ID = QryGetStudent.StuID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblContrStudent LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblContrStudent.EmpID = dbo.TblEmployee.Emp_ID ON QryGetStudent.ContNoID = dbo.TblContrStudent.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblContrStudent.CompID = dbo.TblCustemers.CusID"
sql = sql & "  Where  dbo.TblContrStudent.CompID in (" & Compnies & ") AND (dbo.TblStudent.StutsID <> 1) "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.FgItem
.Rows = .Rows + Rs4.RecordCount
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
For k = 1 To .Rows - 1
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("TypeSuper")) = IIf(IsNull(Rs4("TypeSuper").value), 0, Rs4("TypeSuper").value)

.TextMatrix(k, .ColIndex("CompID")) = IIf(IsNull(Rs4("CompID").value), 0, Rs4("CompID").value)
.TextMatrix(k, .ColIndex("ContNo")) = IIf(IsNull(Rs4("ID").value), 0, Rs4("ID").value)
.TextMatrix(k, .ColIndex("Fullcode")) = IIf(IsNull(Rs4("StudFullCode").value), "", Rs4("StudFullCode").value)
NoStud = IIf(IsNull(Rs4("NoStud").value), 1, Rs4("NoStud").value)
If NoStud = 0 Then
NoStud = 1
End If
.TextMatrix(k, .ColIndex("Price")) = IIf(IsNull(Rs4("StudValue").value), 0, Rs4("StudValue").value)
.TextMatrix(k, .ColIndex("Price")) = Round(val(.TextMatrix(k, .ColIndex("Price"))), 2)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("CusName")) = IIf(IsNull(Rs4("CusName").value), "", Rs4("CusName").value)
Else
.TextMatrix(k, .ColIndex("CusName")) = IIf(IsNull(Rs4("CusNamee").value), "", Rs4("CusNamee").value)
End If
'GetStudent val(.TextMatrix(k, .ColIndex("ContNo"))), StudID, Name, fullcode
.TextMatrix(k, .ColIndex("EmpID")) = IIf(IsNull(Rs4("StuID").value), 0, Rs4("StuID").value)
.TextMatrix(k, .ColIndex("StudCode")) = IIf(IsNull(Rs4("StudFullCode").value), "", Rs4("StudFullCode").value)
.TextMatrix(k, .ColIndex("MrkExp")) = IIf(IsNull(Rs4("TypeDis").value), -1, Rs4("TypeDis").value) + 1
.TextMatrix(k, .ColIndex("MrkPers")) = IIf(IsNull(Rs4("Discount").value), 0, Rs4("Discount").value)
.TextMatrix(k, .ColIndex("PerstSup")) = IIf(IsNull(Rs4("Comm").value), 0, Rs4("Comm").value)
.TextMatrix(k, .ColIndex("SuperID")) = IIf(IsNull(Rs4("EmpID").value), 0, Rs4("EmpID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("Name")) = IIf(IsNull(Rs4("Name").value), "", Rs4("Name").value)
Else
.TextMatrix(k, .ColIndex("Name")) = IIf(IsNull(Rs4("NameE").value), "", Rs4("NameE").value)
End If
If val(.TextMatrix(k, .ColIndex("TypeSuper"))) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("SuperName")) = IIf(IsNull(Rs4("Emp_Name").value), "", Rs4("Emp_Name").value)
Else
.TextMatrix(k, .ColIndex("SuperName")) = IIf(IsNull(Rs4("Emp_Namee").value), "", Rs4("Emp_Namee").value)

End If
Else
.TextMatrix(k, .ColIndex("SuperName")) = IIf(IsNull(Rs4("MedalMan").value), "", Rs4("MedalMan").value)
End If


If val(.TextMatrix(k, .ColIndex("MrkExp"))) = 2 Then
.TextMatrix(k, .ColIndex("MarkTotal")) = Round((val(.TextMatrix(k, .ColIndex("MrkPers"))) * val(.TextMatrix(k, .ColIndex("Price")))) / 100, 2)
Else
.TextMatrix(k, .ColIndex("MarkTotal")) = val(.TextMatrix(k, .ColIndex("MrkPers")))

End If
.TextMatrix(k, .ColIndex("TotalSup")) = Round((val(.TextMatrix(k, .ColIndex("PerstSup"))) * (val(.TextMatrix(k, .ColIndex("Price"))) - val(.TextMatrix(k, .ColIndex("MarkTotal"))))) / 100, 2)
Rs4.MoveNext
Next k
End If
Rs4.Close
End With
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub ListGroupAll_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   FrmCustemerSearch.searchtype = 24
        FrmCustemerSearch.show vbModal
  End If
End Sub

Private Sub ListGroupSelected_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   FrmCustemerSearch.searchtype = 25
        FrmCustemerSearch.show vbModal
  End If
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(RecordDate.value)
           On Error Resume Next
               On Error Resume Next
'
    CmbMonth.ListIndex = Month(RecordDate.value) - 1
     CboYear.ListIndex = val(year(RecordDate.value) - 2006)
End If
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
  On Error Resume Next
    CboYear.Text = year(RecordDate.value)
    CmbMonth.ListIndex = Month(RecordDate.value) - 1
 RecordDate.value = ToGregorianDate(RecordDateH.value)
End If
End Sub

Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 3000
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

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
    Dim sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.name, True) = False Then
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
         strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.txtnoteid.Text)
          Cn.Execute strSQL, , adExecuteNoRecords
          strSQL = "Delete From TblIssuBillStudentDet Where   IsuBillID=" & val(Me.TxtSerial1.Text)
               Cn.Execute strSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            FgItem.Clear flexClearScrollable, flexClearEverything
            FgItem.Rows = 1
                Lbl(6).Caption = 0
    ListGroupSelected.Clear
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
    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        FgItem.Rows = FgItem.Rows + 1
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
    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    clear_all Me
    TxtModFlg.Text = "N"
    RecordDate_Change
    FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.Rows = 1
    Lbl(6).Caption = 0
    ListGroupSelected.Clear
    Me.DCboUserName.BoundText = user_id
    DcbBranch.BoundText = Current_branch
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
   Label1(2).Caption = "Billing Companies"
Lbl(4).Caption = "No"
Ele(3).Caption = "Period"
Lbl(1).Caption = "Date"
Lbl(0).Caption = "Branch"
Lbl(2).Caption = "Remarks"
Lbl(3).Caption = "Companies"
Lbl(7).Caption = "Year"
Lbl(9).Caption = "Month"
ISButton2.Caption = "Issuing"
    Cmd(3).Caption = "Delete"
    Cmd(4).Caption = "Delete All"
     Label1(35).Caption = "No.GL"
Command9.Caption = "Print GL"
Lbl(6).Caption = "Total"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"
C1Tab1.Caption = "Data"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.Lbl(8).Caption = "by"
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
  With Me.FgItem
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("StudCode")) = "Student Code"
  .TextMatrix(0, .ColIndex("Name")) = "Student Name"
  .TextMatrix(0, .ColIndex("Price")) = "Value "
  .TextMatrix(0, .ColIndex("CusName")) = "Company"
  .TextMatrix(0, .ColIndex("Fullcode")) = "Contract"
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblIssuBillStudent"
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

 If ListGroupAll.ListIndex > -1 Then
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
            End If
End Sub
Sub serch(Optional CompID As Long, Optional Typ As Integer = 0)
Dim i As Integer
If Typ = 0 Then
With ListGroupAll
For i = 0 To .ListCount - 1
If CompID = val(.ItemData(i)) Then
.Selected(i) = True
End If
Next i
End With
Else
With ListGroupSelected
For i = 0 To .ListCount - 1
If CompID = val(.ItemData(i)) Then
.Selected(i) = True
End If
Next i
End With
End If
End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
  sql = " SELECT * from  TblCustemers where type=55"
  If SystemOptions.UserInterface = ArabicInterface Then
  sql = sql & "Order by CusName"
  Else
  sql = sql & "Order by CusNamee "
  End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("CusID").value
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
    ListGroupSelected.Clear
    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i

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
