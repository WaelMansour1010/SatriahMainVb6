VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmImportShifts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmImportShifts.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   14550
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
      ItemData        =   "FrmImportShifts.frx":6852
      Left            =   15480
      List            =   "FrmImportShifts.frx":6862
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
      TabIndex        =   0
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
            Picture         =   "FrmImportShifts.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportShifts.frx":83B1
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
      ButtonImage     =   "FrmImportShifts.frx":874B
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
      ButtonImage     =   "FrmImportShifts.frx":EFAD
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
      ButtonImage     =   "FrmImportShifts.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9315
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   14550
      _cx             =   25665
      _cy             =   16431
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
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   14535
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
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
            ButtonImage     =   "FrmImportShifts.frx":15BA9
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
            ButtonImage     =   "FrmImportShifts.frx":15F43
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
            ButtonImage     =   "FrmImportShifts.frx":162DD
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
            ButtonImage     =   "FrmImportShifts.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«” Ì—«œ »Ì«‰«  «·Õ÷Ê— Ê«·«‰’—«›"
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
            TabIndex        =   20
            Top             =   240
            Width           =   4080
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmImportShifts.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   990
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   8325
         Width           =   14550
         _cx             =   25665
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
            Height          =   285
            Left            =   13095
            TabIndex        =   22
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   540
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   285
            Left            =   11280
            TabIndex        =   23
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   540
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   285
            Left            =   9660
            TabIndex        =   24
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   540
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   285
            Left            =   7950
            TabIndex        =   25
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   540
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   285
            Left            =   6210
            TabIndex        =   26
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   540
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   345
            Left            =   5205
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   540
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   609
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
            ButtonImage     =   "FrmImportShifts.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   285
            Left            =   1665
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   540
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   285
            Left            =   3435
            TabIndex        =   29
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   540
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
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
            ButtonImage     =   "FrmImportShifts.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10605
            TabIndex        =   37
            Top             =   105
            Width           =   2670
            _ExtentX        =   4710
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
            Height          =   180
            Left            =   240
            TabIndex        =   42
            Top             =   210
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1785
            TabIndex        =   41
            Top             =   225
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   180
            Index           =   1
            Left            =   810
            TabIndex        =   40
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   180
            Index           =   0
            Left            =   2505
            TabIndex        =   39
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   240
            Index           =   8
            Left            =   13575
            TabIndex        =   38
            Top             =   105
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6960
         Left            =   0
         TabIndex        =   30
         Top             =   1410
         Width           =   14535
         _cx             =   25638
         _cy             =   12277
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
         Caption         =   "»Ì«‰«  «”«”Ì…|ÌœÊÌ|«·„—›ﬁ« "
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
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   6540
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   11536
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
               Height          =   975
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   0
               Width           =   14445
               _cx             =   25479
               _cy             =   1720
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
               Begin XtremeSuiteControls.RadioButton Rd 
                  Height          =   345
                  Index           =   1
                  Left            =   12600
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1695
                  _Version        =   786432
                  _ExtentX        =   2990
                  _ExtentY        =   609
                  _StockProps     =   79
                  Caption         =   "’Ì€… «·Êﬁ  12"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdImport 
                  Height          =   510
                  Left            =   120
                  TabIndex        =   47
                  Top             =   165
                  Width           =   4770
                  _ExtentX        =   8414
                  _ExtentY        =   900
                  Caption         =   " Õ„Ì· «·„·›"
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
                  ButtonImage     =   "FrmImportShifts.frx":2CB3E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin ImpulseButton.ISButton CMDSelectFile 
                  Height          =   510
                  Left            =   4950
                  TabIndex        =   48
                  Top             =   165
                  Width           =   5145
                  _ExtentX        =   9075
                  _ExtentY        =   900
                  Caption         =   "Õœœ „”«— «·„·›"
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
                  ButtonImage     =   "FrmImportShifts.frx":333A0
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin MSComDlg.CommonDialog CD1 
                  Left            =   1080
                  Top             =   0
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin XtremeSuiteControls.RadioButton Rd 
                  Height          =   345
                  Index           =   0
                  Left            =   10335
                  TabIndex        =   53
                  Top             =   315
                  Width           =   1695
                  _Version        =   786432
                  _ExtentX        =   2990
                  _ExtentY        =   609
                  _StockProps     =   79
                  Caption         =   "’Ì€… «·Êﬁ  24"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   5535
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   960
               Width           =   14445
               _cx             =   25479
               _cy             =   9763
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
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   4740
                  Left            =   135
                  TabIndex        =   36
                  Top             =   180
                  Width           =   14220
                  _cx             =   25082
                  _cy             =   8361
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
                  Rows            =   12
                  Cols            =   22
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmImportShifts.frx":39C02
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
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   465
                  Index           =   3
                  Left            =   13080
                  TabIndex        =   43
                  Top             =   4950
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   820
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
                  ButtonImage     =   "FrmImportShifts.frx":39F5B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   465
                  Index           =   4
                  Left            =   11775
                  TabIndex        =   44
                  Top             =   4950
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   820
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
                  ButtonImage     =   "FrmImportShifts.frx":3A4F5
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   975
               Left            =   0
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   0
               Width           =   14445
               _cx             =   25479
               _cy             =   1720
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
                  Left            =   7545
                  MaxLength       =   50
                  TabIndex        =   73
                  Top             =   600
                  Width           =   750
               End
               Begin XtremeSuiteControls.CheckBox SelectBranch 
                  Height          =   240
                  Left            =   13155
                  TabIndex        =   72
                  Top             =   240
                  Width           =   1020
                  _Version        =   786432
                  _ExtentX        =   1799
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "›—⁄ „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAll 
                  Height          =   285
                  Left            =   12855
                  TabIndex        =   74
                  Top             =   585
                  Width           =   1380
                  _Version        =   786432
                  _ExtentX        =   2434
                  _ExtentY        =   503
                  _StockProps     =   79
                  Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdEmp 
                  Height          =   240
                  Left            =   8160
                  TabIndex        =   75
                  Top             =   615
                  Width           =   1380
                  _Version        =   786432
                  _ExtentX        =   2434
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "„ÊŸ› „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEmployee1 
                  Height          =   315
                  Left            =   3690
                  TabIndex        =   76
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   615
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbBranch1 
                  Height          =   315
                  Left            =   9690
                  TabIndex        =   77
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   240
                  Width           =   3495
                  _ExtentX        =   6165
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcpDept1 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   78
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   240
                  Width           =   3495
                  _ExtentX        =   6165
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   79
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   600
                  Width           =   3480
                  _ExtentX        =   6138
                  _ExtentY        =   556
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
                  ButtonImage     =   "FrmImportShifts.frx":3AA8F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin XtremeSuiteControls.CheckBox SelectDept 
                  Height          =   240
                  Left            =   8310
                  TabIndex        =   80
                  Top             =   240
                  Width           =   1230
                  _Version        =   786432
                  _ExtentX        =   2170
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "«œ«—… „Õœœ…"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbProject1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   81
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   240
                  Width           =   3495
                  _ExtentX        =   6165
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox SelectProject 
                  Height          =   240
                  Left            =   3510
                  TabIndex        =   82
                  Top             =   255
                  Width           =   1230
                  _Version        =   786432
                  _ExtentX        =   2170
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "„‘—Ê⁄ „Õœœ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   0
                  Left            =   12855
                  TabIndex        =   83
                  Top             =   0
                  Width           =   1560
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   5535
               Left            =   0
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   840
               Width           =   14445
               _cx             =   25479
               _cy             =   9763
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
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   4740
                  Left            =   135
                  TabIndex        =   85
                  Top             =   180
                  Width           =   14220
                  _cx             =   25082
                  _cy             =   8361
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
                  Rows            =   12
                  Cols            =   23
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmImportShifts.frx":412F1
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
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   465
                  Index           =   0
                  Left            =   13095
                  TabIndex        =   86
                  Top             =   4950
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   820
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
                  ButtonImage     =   "FrmImportShifts.frx":41667
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   465
                  Index           =   1
                  Left            =   11775
                  TabIndex        =   87
                  Top             =   4950
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   820
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
                  ButtonImage     =   "FrmImportShifts.frx":41C01
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6540
            Left            =   15480
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   11536
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
            Begin VB.TextBox TxtCode 
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
               Height          =   2010
               Left            =   10740
               MaxLength       =   50
               TabIndex        =   58
               Top             =   4125
               Width           =   765
            End
            Begin XtremeSuiteControls.CheckBox BranchSelect 
               Height          =   1605
               Left            =   11595
               TabIndex        =   57
               Top             =   1605
               Width           =   1005
               _Version        =   786432
               _ExtentX        =   1773
               _ExtentY        =   2831
               _StockProps     =   79
               Caption         =   "›—⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton SelectAll 
               Height          =   1935
               Left            =   12735
               TabIndex        =   59
               Top             =   1500
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   3413
               _StockProps     =   79
               Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton EmpSelect 
               Height          =   1605
               Left            =   12735
               TabIndex        =   60
               Top             =   4125
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   2831
               _StockProps     =   79
               Caption         =   "„ÊŸ› „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbEmployee 
               Height          =   315
               Left            =   6945
               TabIndex        =   61
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   4125
               Width           =   3840
               _ExtentX        =   6773
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch 
               Height          =   315
               Left            =   6945
               TabIndex        =   62
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1605
               Width           =   4560
               _ExtentX        =   8043
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDepatment 
               Height          =   315
               Left            =   1080
               TabIndex        =   63
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1605
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   5235
               Left            =   135
               TabIndex        =   64
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   900
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   9234
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
               ButtonImage     =   "FrmImportShifts.frx":4219B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin XtremeSuiteControls.CheckBox DeptSelect 
               Height          =   1605
               Left            =   5550
               TabIndex        =   65
               Top             =   1605
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   2831
               _StockProps     =   79
               Caption         =   "«œ«—… „Õœœ…"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbProject 
               Height          =   315
               Left            =   1080
               TabIndex        =   66
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   4125
               Width           =   4380
               _ExtentX        =   7726
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ProjSelect 
               Height          =   1605
               Left            =   5550
               TabIndex        =   67
               Top             =   4125
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   2831
               _StockProps     =   79
               Caption         =   "„‘—Ê⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
               ForeColor       =   &H00800000&
               Height          =   1815
               Index           =   3
               Left            =   12870
               TabIndex        =   68
               Top             =   0
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   6540
            Left            =   15180
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   11536
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
            Begin C1SizerLibCtl.C1Tab C1Tab2 
               Height          =   6900
               Left            =   0
               TabIndex        =   70
               Top             =   0
               Width           =   14535
               _cx             =   25638
               _cy             =   12171
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
               Caption         =   "»Ì«‰«  «”«”Ì…|ÌœÊÌ|«·„—›ﬁ« "
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
               Flags(2)        =   2
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   690
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   735
         Width           =   14565
         _cx             =   25691
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
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   315
            Left            =   1680
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            CustomFormat    =   "hh:mm:ss"
            Format          =   95617026
            CurrentDate     =   38784
         End
         Begin VB.TextBox txtFile 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   49
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11565
            TabIndex        =   1
            Top             =   240
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8160
            TabIndex        =   45
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   95617025
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   3720
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   95617025
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   5520
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   95617025
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   315
            Left            =   0
            TabIndex        =   55
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            CustomFormat    =   "hh:mm"
            Format          =   95617026
            CurrentDate     =   38784
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   9840
            TabIndex        =   46
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   13440
            TabIndex        =   33
            Top             =   240
            Width           =   975
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
Attribute VB_Name = "FrmImportShifts"
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
 Public Auto_Man As Integer
 Dim II As Long
 Public LngRow As Long
 Public LngCol As Long
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  
MySQL = "SELECT     dbo.TblImportShifts.ID, dbo.TblImportShifts.RecordDate, dbo.TblImportShiftsDet.MachinDate, dbo.TblImportShiftsDet.MachinCode, dbo.TblImportShiftsDet.TimInExsist, "
MySQL = MySQL & "                      dbo.TblImportShiftsDet.TimIn, dbo.TblImportShiftsDet.TimOutExsist, dbo.TblImportShiftsDet.TimOut, dbo.TblImportShiftsDet.EmpID, dbo.TblEmployee.Emp_Code,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee"
MySQL = MySQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblImportShiftsDet ON dbo.TblEmployee.Emp_ID = dbo.TblImportShiftsDet.EmpID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblImportShifts ON dbo.TblImportShiftsDet.ImportShiftID = dbo.TblImportShifts.ID"
MySQL = MySQL & " Where (dbo.TblImportShifts.ID = " & val(TxtSerial1.Text) & ") "
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportShiftsImports.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportShiftsImports.rpt"
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



Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 0
RemoveGridRow2
Case 1
 fg.Clear flexClearScrollable, flexClearEverything
 fg.Rows = 1
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
End Select
End Sub


Private Sub CmdImport_Click()
ExilSheet
txtFile.Text = ""
'FullGridData2
End Sub

Private Sub CMDSelectFile_Click()
If Rd(0).value = False And Rd(1).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Õœœ ’Ì€… «·Êﬁ "
Exit Sub
Else
MsgBox "Plese Select Time format"
Exit Sub
End If
End If

CD1.ShowOpen
txtFile.Text = CD1.filename
End Sub

Private Sub DcbEmployee1_Change()
DcbEmployee1_Click (0)
End Sub

Private Sub DcbEmployee1_Click(Area As Integer)
If val(DcbEmployee1.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , Me.DcbEmployee1.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Auto_Man = 1 Then
Cancel = True
Else
With Me.fg
Select Case .ColKey(Col)
Case "FullCode"
Cancel = True
Case "Emp_Name"
Cancel = True
Case "MachinDate"
.ComboList = ""
Case "RecTime"
.ComboList = ""
Case "ToTime"
.ComboList = ""
End Select
End With
End If
End Sub


Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg.Text <> "R" Then
With Me.fg
Select Case .ColKey(Col)
  Case "MachinDate"
       LngRow = Row
        LngCol = Col
        FrmDateOpProject.Index = 19
        Load FrmDateOpProject
        FrmDateOpProject.Index = 19
        FrmDateOpProject.show vbModal
  Case "RecTime"
       LngRow = Row
        LngCol = Col
        FrmDateOpProject.Index = 20
        Load FrmDateOpProject
        FrmDateOpProject.Index = 20
        FrmDateOpProject.show vbModal
 Case "ToTime"
       LngRow = Row
        LngCol = Col
        FrmDateOpProject.Index = 21
        Load FrmDateOpProject
        FrmDateOpProject.Index = 21
        FrmDateOpProject.show vbModal
        
End Select
End With
End If
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.fg
Select Case .ColKey(Col)
  Case "MachinDate"
  If Auto_Man = 0 Then
.ColComboList(.ColIndex("MachinDate")) = "..."
End If
  Case "RecTime"
  If Auto_Man = 0 Then
.ColComboList(.ColIndex("RecTime")) = "..."
End If
  Case "ToTime"
  If Auto_Man = 0 Then
.ColComboList(.ColIndex("ToTime")) = "..."
End If
End Select
End With
End Sub

 Private Sub Form_Load()
 '   On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblImportShifts where Auto_Man=" & Auto_Man & "  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetEmployees Me.DcbEmployee1
   Dcombos.GetBranches Me.DcbBranch1
   Dcombos.GetEmpDepartments Me.DcpDept1
   Dcombos.GetProjects Me.DcbProject1
    C1Elastic5.Visible = False
    C1Elastic7.Visible = False
    
   With Me.GridInstallments
   If Auto_Man = 1 Then
   C1Elastic5.Visible = True
    Label1(2).Caption = " «” Ì—«œ »Ì«‰«  «·Õ÷Ê— Ê«·«‰’—«›"
   C1Elastic6.Visible = False
   C1Elastic4.Visible = True
   .ColHidden(.ColIndex("MachinCode")) = False
   Else
   C1Elastic7.Visible = True
   Label1(2).Caption = " ”ÃÌ· »Ì«‰«  «·Õ÷Ê— Ê«·«‰’—«› "
   C1Elastic4.Visible = False
   C1Elastic6.Visible = True
   .ColHidden(.ColIndex("MachinCode")) = True
   End If
   End With
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
Function CheckBetwenPeriod(Optional EmpID1 As Double, Optional NoDay As Integer, Optional Timin As String) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TblShiftWorker.EmpID AS EmpID1, dbo.TbLSheft.*"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where dbo.TblShiftWorker.EmpID=" & EmpID1 & ""
'Sql = " SELECT     dbo.TbLSheft.*"
'Sql = Sql & " FROM         dbo.TbLSheft where 1=1 "
Select Case NoDay
Case 7
sql = sql & " and  ShfitFromW <= '" & Timin & "' "
sql = sql & " and ShfitToW >= '" & Timin & "' "
Case 6
sql = sql & " and  FromFriW <= '" & Timin & "' "
sql = sql & " and ToFriW >= '" & Timin & "' "
Case 5
sql = sql & " and  FromThruW <= '" & Timin & "' "
sql = sql & " and ToThruW >= '" & Timin & "' "
Case 4
sql = sql & " and  FromWedW <= '" & Timin & "' "
sql = sql & " and ToWedW >= '" & Timin & "' "
Case 3
sql = sql & " and  FromTuesW <= '" & Timin & "' "
sql = sql & " and ToTuesW >=  '" & Timin & "' "
Case 2
sql = sql & " and  FromMonW <= '" & Timin & "' "
sql = sql & " and ToMonW >= '" & Timin & "' "
Case 1
sql = sql & " and  FromSunW <= '" & Timin & "' "
sql = sql & " and ToSunW >= '" & Timin & "' "
End Select

Rs8.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckBetwenPeriod = IIf(IsNull(Rs8("SeftCode").value), 0, (Rs8("SeftCode").value))
Else
CheckBetwenPeriod = 0
End If
End Function


Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                  StrSQL = "Delete From TblImportShiftsDet Where ImportShiftID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                  StrSQL = "Delete From TblImportShiftsDet2 Where ImportShiftID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
              End If
  If Auto_Man = 0 Then
     FillGriToGir
  End If
   RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("DeptID").value = val(Me.DcpDept1.BoundText)
   RsSavRec.Fields("ProjID").value = val(Me.DcbProject1.BoundText)
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch1.BoundText)
   RsSavRec.Fields("EmpID").value = val(Me.DcbEmployee1.BoundText)
   If Me.RdAll.value = True Then
   RsSavRec.Fields("SelectAll").value = 1
   End If
   If RdEmp.value = True Then
   RsSavRec.Fields("SelectEmp").value = 1
   End If
   If Me.SelectBranch.value = vbChecked Then
   RsSavRec.Fields("SelectBranch").value = 1
   End If
   If Me.SelectDept.value = vbChecked Then
   RsSavRec.Fields("SelectDept").value = 1
   End If
   If Me.SelectProject.value = vbChecked Then
   RsSavRec.Fields("SelectProject").value = 1
   End If
   RsSavRec.Fields("Auto_Man").value = Auto_Man
   
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblImportShiftsDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ImportShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("MachinDate").value = IIf((.TextMatrix(i, .ColIndex("MachinDate"))) = "", Null, (.TextMatrix(i, .ColIndex("MachinDate"))))
                RsDevsub("MachinCode").value = IIf((.TextMatrix(i, .ColIndex("MachinCode"))) = "", Null, Trim(.TextMatrix(i, .ColIndex("MachinCode"))))
                RsDevsub("RecTime").value = IIf((.TextMatrix(i, .ColIndex("RecTime"))) = "", Null, (.TextMatrix(i, .ColIndex("RecTime"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("ProjID").value = IIf((.TextMatrix(i, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjID"))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
   Dim RecTime As String
   
         RecTime = IIf(IsNull(RsDevsub("RecTime").value), "", RsDevsub("RecTime").value)
 
 If Rd(0).value = True Then '24
 Dim Hour As String
' Hour = mId(RecTime, 1, 2)
 
         If val(Hour) > 12 Then
         '        Hour = Hour - 12
         
                    If val(Hour) < 10 Then
         '           Hour = "0" & Hour
                    End If
         
         
         End If
 
'RecTime = Hour & mId(RecTime, 3, Len(RecTime))
  
 End If
 If RecTime <> "" Then
RecTime = mId(RecTime, 1, Len(RecTime) - 3)
 End If
RsDevsub("ShiftID").value = CheckBetwenPeriod(RsDevsub("EmpID").value, Weekday(RsDevsub("MachinDate").value), RecTime)

          RsDevsub.update
      End If
     Next i
    End With
   ''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblImportShiftsDet2 Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.fg
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 And .TextMatrix(i, .ColIndex("MachinDate")) <> "" Then
       RsDevsub.AddNew
                RsDevsub("ImportShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("MachinDate").value = IIf((.TextMatrix(i, .ColIndex("MachinDate"))) = "", Null, (.TextMatrix(i, .ColIndex("MachinDate"))))
                RsDevsub("RecTime").value = IIf((.TextMatrix(i, .ColIndex("RecTime"))) = "", Null, (.TextMatrix(i, .ColIndex("RecTime"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("ProjID").value = IIf((.TextMatrix(i, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjID"))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("ToTime").value = IIf((.TextMatrix(i, .ColIndex("ToTime"))) = "", Null, (.TextMatrix(i, .ColIndex("ToTime"))))
       RsDevsub.update
      End If
     Next i
    End With
'SavAbcens
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
     SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    Dim Shifttime As Date
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcpDept1.BoundText = IIf(IsNull(RsSavRec.Fields("DeptID").value), "", RsSavRec.Fields("DeptID").value)
    Me.DcbProject1.BoundText = IIf(IsNull(RsSavRec.Fields("ProjID").value), "", RsSavRec.Fields("ProjID").value)
    Me.DcbBranch1.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DcbEmployee1.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    If Not (IsNull(RsSavRec.Fields("SelectAll").value)) Then
    If RsSavRec.Fields("SelectAll").value = 1 Then
    RdAll.value = True
    Else
    RdAll.value = False
    End If
    Else
    RdAll.value = False
    End If
    If Not (IsNull(RsSavRec.Fields("SelectEmp").value)) Then
    If RsSavRec.Fields("SelectEmp").value = 1 Then
    Me.EmpSelect.value = True
    Else
    EmpSelect.value = False
    End If
    Else
    EmpSelect.value = False
    End If
    If Not (IsNull(RsSavRec.Fields("SelectBranch").value)) Then
    If RsSavRec.Fields("SelectBranch").value = 1 Then
    Me.SelectBranch.value = vbChecked
    Else
    EmpSelect.value = vbUnchecked
    End If
    Else
    EmpSelect.value = vbUnchecked
    End If
     If Not (IsNull(RsSavRec.Fields("SelectDept").value)) Then
    If RsSavRec.Fields("SelectDept").value = 1 Then
    Me.SelectDept.value = vbChecked
    Else
    SelectDept.value = vbUnchecked
    End If
    Else
    SelectDept.value = vbUnchecked
    End If
    If Not (IsNull(RsSavRec.Fields("SelectProject").value)) Then
    If RsSavRec.Fields("SelectProject").value = 1 Then
    Me.SelectProject.value = vbChecked
    Else
    SelectProject.value = vbUnchecked
    End If
    Else
    SelectProject.value = vbUnchecked
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

    '---------------------- check if data Vaclete -----------------------
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
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblImportShifts", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Function GetShiftId(Optional EmpID As Double, Optional Timin As String, Optional NoDay As Integer) As Double
Dim sql As String
Dim i As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim temp As String
Dim swap As String
Dim ShiftIDTemp As Double
Dim ShiftID As Double
sql = " SELECT     ShiftID, EmpID"
sql = sql & " From dbo.TblShiftWorker"
sql = sql & " WHERE     (EmpID = " & EmpID & ") "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Rs7.MoveFirst
For i = 1 To Rs7.RecordCount
ShiftID = IIf(IsNull(Rs7("ShiftID").value), 0, Rs7("ShiftID").value)
If CheckPeriod(ShiftID, NoDay) <> "" Then
If i <> 1 Then
swap = temp
Else
ShiftIDTemp = ShiftID
End If
temp = CheckPeriod(ShiftID, NoDay)
If i <> 1 Then
If Abs(DateDiff("n", Timin, swap)) > Abs(DateDiff("n", Timin, temp)) Then
ShiftIDTemp = ShiftID
End If
End If
End If
Rs7.MoveNext
Next i
Else
GetShiftId = 0
Exit Function
End If
GetShiftId = ShiftIDTemp
End Function
Function CheckPeriod(Optional ShiftID As Double, Optional NoDay As Integer) As String
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TbLSheft.*"
sql = sql & " FROM         dbo.TbLSheft where SeftCode=" & ShiftID & " "
Rs8.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
Select Case NoDay
Case 7
CheckPeriod = IIf(IsNull(Rs8("Shiftfrom").value), 0, (Rs8("Shiftfrom").value))
Case 6
CheckPeriod = IIf(IsNull(Rs8("FromFri").value), 0, (Rs8("FromFri").value))
Case 5
CheckPeriod = IIf(IsNull(Rs8("FromThru").value), 0, (Rs8("FromThru").value))
Case 4
CheckPeriod = IIf(IsNull(Rs8("FromWed").value), 0, (Rs8("FromWed").value))
Case 3
CheckPeriod = IIf(IsNull(Rs8("FromTues").value), 0, (Rs8("FromTues").value))
Case 2
CheckPeriod = IIf(IsNull(Rs8("FromMon").value), 0, (Rs8("FromMon").value))
Case 1
CheckPeriod = IIf(IsNull(Rs8("FromSun").value), 0, (Rs8("FromSun").value))
End Select
Else
CheckPeriod = ""
End If
End Function
 Sub FullGridData2()
 Dim TimeinExists As String
 Dim TimeoutExists As String
 Dim TypeDay As Integer
 Dim i As Integer
 Dim NoDay As Double
 Dim ShiftID As Double
 Dim TypeTrans As Integer
 On Error GoTo ErrTrap
     With Me.GridInstallments
     
                    For i = .FixedRows To .Rows
                    If (.TextMatrix(i, .ColIndex("TimIn")) = "" And .TextMatrix(i, .ColIndex("TimOut")) <> "") Or (.TextMatrix(i, .ColIndex("TimIn")) <> "" And .TextMatrix(i, .ColIndex("TimOut")) = "") Then
                    If .TextMatrix(i, .ColIndex("TimIn")) = "" Then
                    .Cell(flexcpBackColor, i, 1, i, 17) = &HE0E0E0
                    ElseIf .TextMatrix(i, .ColIndex("TimOut")) = "" Then
                    .Cell(flexcpBackColor, i, 1, i, 17) = &H80FF&
                    End If
                    End If
                    If .TextMatrix(i, .ColIndex("TimIn")) <> "" Then
                    DTPicker1.value = IIf(IsDate(.TextMatrix(i, .ColIndex("MachinDate"))), .TextMatrix(i, .ColIndex("MachinDate")), Null)
                    ShiftID = GetShiftId(val(.TextMatrix(i, .ColIndex("EmpID"))), .TextMatrix(i, .ColIndex("TimIn")), Weekday(DTPicker1))
                    Else
                    If val(.TextMatrix(i, .ColIndex("Absence"))) = 0 Then
                    RetriveShifts ShiftID, TimeinExists, TimeoutExists, TypeDay, 0, Weekday(DTPicker1)
                   .TextMatrix(i, .ColIndex("TypeDay")) = TypeDay
                   .TextMatrix(i, .ColIndex("TimInExsist")) = TimeinExists
                   .TextMatrix(i, .ColIndex("TimOutExsist")) = TimeoutExists
                   
                   If Rd(0).value = True Then
                    .TextMatrix(i, .ColIndex("DiffIn")) = DateDiff("b", .TextMatrix(i, .ColIndex("TimInExsist")), .TextMatrix(i, .ColIndex("TimIn")))
                 Else
                   DTPicker3.value = Format(.TextMatrix(i, .ColIndex("TimInExsist")), "HH:MM:SS")
                  .TextMatrix(i, .ColIndex("TimInExsist")) = DTPicker3.value
                  .TextMatrix(i, .ColIndex("DiffIn")) = DateDiff("n", .TextMatrix(i, .ColIndex("TimInExsist")), .TextMatrix(i, .ColIndex("TimIn")))
                End If
                TypeTrans = 0
                
                If val(.TextMatrix(i, .ColIndex("DiffIn"))) > 0 Then
                TypeTrans = 5
                ElseIf val(.TextMatrix(i, .ColIndex("DiffIn"))) < 0 And val(.TextMatrix(i, .ColIndex("TypeDay"))) = 0 Then
                TypeTrans = 1
                 ElseIf val(.TextMatrix(i, .ColIndex("DiffIn"))) < 0 And val(.TextMatrix(i, .ColIndex("TypeDay"))) = 1 Then
                 TypeTrans = 2
                End If
                 RereiveSlice val(.TextMatrix(i, .ColIndex("EmpID"))), val(.TextMatrix(i, .ColIndex("DiffIn"))), TypeTrans, NoDay
                 
                .TextMatrix(i, .ColIndex("NetDiffIn")) = NoDay
                .TextMatrix(i, .ColIndex("TypeTransIn")) = TypeTrans
                   If Rd(0).value = True Then
                    .TextMatrix(i, .ColIndex("DiffOut")) = DateDiff("b", .TextMatrix(i, .ColIndex("TimOutExsist")), .TextMatrix(i, .ColIndex("TimOut")))
                 Else
                   DTPicker3.value = Format(.TextMatrix(i, .ColIndex("TimOutExsist")), "HH:MM:SS")
                  .TextMatrix(i, .ColIndex("TimOutExsist")) = DTPicker3.value
                  .TextMatrix(i, .ColIndex("DiffOut")) = DateDiff("n", .TextMatrix(i, .ColIndex("TimOutExsist")), .TextMatrix(i, .ColIndex("TimOut")))
                End If
                TypeTrans = 0
                If val(.TextMatrix(i, .ColIndex("DiffOut"))) < 0 Then
                TypeTrans = 6
                ElseIf val(.TextMatrix(i, .ColIndex("DiffOut"))) > 0 And val(.TextMatrix(i, .ColIndex("TypeDay"))) = 0 Then
                 TypeTrans = 1
                 ElseIf val(.TextMatrix(i, .ColIndex("DiffOut"))) > 0 And val(.TextMatrix(i, .ColIndex("TypeDay"))) = 1 Then
                  TypeTrans = 2
                End If
                RereiveSlice val(.TextMatrix(i, .ColIndex("EmpID"))), val(.TextMatrix(i, .ColIndex("DiffOut"))), TypeTrans, NoDay
                .TextMatrix(i, .ColIndex("NetDiffOut")) = NoDay
                .TextMatrix(i, .ColIndex("TypeTransOut")) = TypeTrans
               End If
               End If
             Next i
End With
        
        Exit Sub
ErrTrap:
    End Sub
Sub RetriveShifts(Optional ShiftID As Double = 0, Optional ByRef TimeinExists As String, Optional ByRef TimoutExists As String, Optional ByRef TypeDay As Integer = -1, Optional Typ As Integer = 0, Optional NoDay As Integer)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TblShiftWorker.ShiftID, dbo.TblShiftWorker.Typetrans, dbo.TblShiftWorker.FromMint, dbo.TblShiftWorker.ToMint, dbo.TblShiftWorker.AverageMaint, "
sql = sql & "                       dbo.TblShiftWorker.DeptID, dbo.TblShiftWorker.BranchID, dbo.TbLSheft.SheftName, dbo.TbLSheft.Remarks, dbo.TbLSheft.ShiftFrom, dbo.TbLSheft.ShiftTo,"
sql = sql & "                        dbo.TbLSheft.ShiftTime, dbo.TbLSheft.SheftNamee, dbo.TbLSheft.SatWoVo, dbo.TbLSheft.SunWoVo, dbo.TbLSheft.MonWoVo, dbo.TbLSheft.TuesWoVo,"
sql = sql & "                        dbo.TbLSheft.WedWoVo, dbo.TbLSheft.ThurWoVo, dbo.TbLSheft.FrirWoVo, dbo.TbLSheft.FromSun, dbo.TbLSheft.ToSun, dbo.TbLSheft.FromMon, dbo.TbLSheft.ToMon,"
sql = sql & "                        dbo.TbLSheft.FromTues, dbo.TbLSheft.ToTues, dbo.TbLSheft.FromWed, dbo.TbLSheft.ToWed, dbo.TbLSheft.FromThru, dbo.TbLSheft.ToThru, dbo.TbLSheft.FromFri,"
sql = sql & "                        dbo.TbLSheft.ToFri"
sql = sql & "   FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                        dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & "   WHERE     (dbo.TblShiftWorker.ShiftID = " & ShiftID & ")"
If Typ = 0 Then
sql = sql & "  and dbo.TblShiftWorker.Typetrans IS NULL"
End If
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
Select Case NoDay
Case 7
TimeinExists = IIf(IsNull(Rs8("Shiftfrom").value), 0, (Rs8("Shiftfrom").value))
TimoutExists = IIf(IsNull(Rs8("ShiftTo").value), 0, (Rs8("ShiftTo").value))
TypeDay = IIf(IsNull(Rs8("SatWoVo").value), -1, Rs8("SatWoVo").value)
Case 6
TimeinExists = IIf(IsNull(Rs8("FromFri").value), 0, (Rs8("FromFri").value))
TimoutExists = IIf(IsNull(Rs8("ToFri").value), 0, (Rs8("ToFri").value))
TypeDay = IIf(IsNull(Rs8("FrirWoVo").value), -1, Rs8("FrirWoVo").value)
Case 5
TimeinExists = IIf(IsNull(Rs8("FromThru").value), 0, (Rs8("FromThru").value))
TimoutExists = IIf(IsNull(Rs8("ToThru").value), 0, (Rs8("ToThru").value))
TypeDay = IIf(IsNull(Rs8("ThurWoVo").value), -1, Rs8("ThurWoVo").value)
Case 4
TimeinExists = IIf(IsNull(Rs8("FromWed").value), 0, (Rs8("FromWed").value))
TimoutExists = IIf(IsNull(Rs8("ToWed").value), 0, (Rs8("ToWed").value))
TypeDay = IIf(IsNull(Rs8("WedWoVo").value), -1, Rs8("WedWoVo").value)
Case 3
TimeinExists = IIf(IsNull(Rs8("FromTues").value), 0, (Rs8("FromTues").value))
TimoutExists = IIf(IsNull(Rs8("ToTues").value), 0, (Rs8("ToTues").value))
TypeDay = IIf(IsNull(Rs8("TuesWoVo").value), -1, Rs8("TuesWoVo").value)
Case 2
TimeinExists = IIf(IsNull(Rs8("FromMon").value), 0, (Rs8("FromMon").value))
TimoutExists = IIf(IsNull(Rs8("ToMon").value), 0, (Rs8("ToMon").value))
TypeDay = IIf(IsNull(Rs8("MonWoVo").value), -1, Rs8("MonWoVo").value)
Case 1
TimeinExists = IIf(IsNull(Rs8("FromSun").value), 0, (Rs8("FromSun").value))
TimoutExists = IIf(IsNull(Rs8("ToSun").value), 0, (Rs8("ToSun").value))
TypeDay = IIf(IsNull(Rs8("SunWoVo").value), -1, Rs8("SunWoVo").value)
End Select
Else
TimeinExists = 0
TimoutExists = 0
End If
  End Sub
Sub RereiveSlice(Optional EmpID As Double = 0, Optional TimDif As Double, Optional TypeTrans As Integer = -1, Optional ByRef NoDay As Double)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = "SELECT     dbo.TblShiftWorker.EmpID, dbo.TblShiftWorker.Typetrans, dbo.TblShiftWorker.FromMint, dbo.TblShiftWorker.ToMint, dbo.TblShiftWorker.AverageMaint"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where (dbo.TblShiftWorker.TypeTrans = " & TypeTrans & ")"
sql = sql & "  and dbo.TblShiftWorker.FromMint <=" & Abs(TimDif) & ""
sql = sql & "  and dbo.TblShiftWorker.ToMint >=" & Abs(TimDif) & ""

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
NoDay = IIf(IsNull(Rs8("AverageMaint").value), 0, Rs8("AverageMaint").value)
Else
NoDay = 0
End If
  End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
 If Auto_Man = 1 Then
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
sql = "SELECT     dbo.TblImportShiftsDet.ID, dbo.TblImportShiftsDet.MachinDate, dbo.TblImportShiftsDet.MachinCode, dbo.TblImportShiftsDet.TimInExsist, dbo.TblImportShiftsDet.TimIn,"
sql = sql & "                       dbo.TblImportShiftsDet.TimOutExsist, dbo.TblImportShiftsDet.TimOut, dbo.TblImportShiftsDet.ImportShiftID, dbo.TblImportShiftsDet.EmpID,"
sql = sql & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblImportShiftsDet.TypeTransOut, dbo.TblImportShiftsDet.TypeTransIn,"
sql = sql & "                      dbo.TblImportShiftsDet.NetDiffOut, dbo.TblImportShiftsDet.DiffOut, dbo.TblImportShiftsDet.NetDiffIn, dbo.TblImportShiftsDet.DiffIn,"
sql = sql & "                      dbo.TblImportShiftsDet.TypeDay ,dbo.TblImportShiftsDet.Absence, "
sql = sql & "                      dbo.TblImportShiftsDet.DeptID ,dbo.TblImportShiftsDet.ProjID, "
sql = sql & "                      dbo.TblImportShiftsDet.RecTime ,dbo.TblImportShiftsDet.BranchID "
sql = sql & " FROM         dbo.TblImportShiftsDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblImportShiftsDet.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblImportShiftsDet.ImportShiftID =" & val(TxtSerial1.Text) & ") "

  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjID").value), 0, Rs1("ProjID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("MachinDate")) = IIf(IsNull(Rs1("MachinDate").value), "", Rs1("MachinDate").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                    .TextMatrix(i, .ColIndex("RecTime")) = IIf(IsNull(Rs1("RecTime").value), "", Rs1("RecTime").value)
                   
                  ' .TextMatrix(i, .ColIndex("TimInExsist")) = IIf(IsNull(Rs1("TimInExsist").value), 0, Rs1("TimInExsist").value)
                  ' .TextMatrix(i, .ColIndex("TimIn")) = IIf(IsNull(Rs1("TimIn").value), 0, Rs1("TimIn").value)
                  ' .TextMatrix(i, .ColIndex("TimOutExsist")) = IIf(IsNull(Rs1("TimOutExsist").value), 0, Rs1("TimOutExsist").value)
                  ' .TextMatrix(i, .ColIndex("TimOut")) = IIf(IsNull(Rs1("TimOut").value), 0, Rs1("TimOut").value)
                   .TextMatrix(i, .ColIndex("MachinCode")) = IIf(IsNull(Rs1("MachinCode").value), "", Rs1("MachinCode").value)
                   .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                  ' .TextMatrix(i, .ColIndex("TypeDay")) = IIf(IsNull(Rs1("TypeDay").value), -1, Rs1("TypeDay").value)
                  ' .TextMatrix(i, .ColIndex("DiffIn")) = IIf(IsNull(Rs1("DiffIn").value), 0, Rs1("DiffIn").value)
                  ' .TextMatrix(i, .ColIndex("NetDiffIn")) = IIf(IsNull(Rs1("NetDiffIn").value), 0, Rs1("NetDiffIn").value)
                  ' .TextMatrix(i, .ColIndex("DiffOut")) = IIf(IsNull(Rs1("DiffOut").value), 0, Rs1("DiffOut").value)
                  ' .TextMatrix(i, .ColIndex("NetDiffOut")) = IIf(IsNull(Rs1("NetDiffOut").value), 0, Rs1("NetDiffOut").value)
                  ' .TextMatrix(i, .ColIndex("TypeTransIn")) = IIf(IsNull(Rs1("TypeTransIn").value), 0, Rs1("TypeTransIn").value)
                   '.TextMatrix(i, .ColIndex("TypeTransOut")) = IIf(IsNull(Rs1("TypeTransOut").value), 0, Rs1("TypeTransOut").value)
                  ' .TextMatrix(i, .ColIndex("Absence")) = IIf(IsNull(Rs1("Absence").value), 0, Rs1("Absence").value)
              
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   End If
                  '   If .TextMatrix(i, .ColIndex("TimIn")) = "0" And .TextMatrix(i, .ColIndex("TimOut")) <> "0" Then
                   ' .Cell(flexcpBackColor, i, 1, i, 17) = &HE0E0E0
                  '  ElseIf .TextMatrix(i, .ColIndex("TimIn")) <> "0" And .TextMatrix(i, .ColIndex("TimOut")) = "0" Then
                  '  .Cell(flexcpBackColor, i, 1, i, 17) = &H80FF&
                   ' End If
                   Rs1.MoveNext
             Next i
End With
Else
''''''''''''''''''''''
    fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 1
sql = "SELECT     dbo.TblImportShiftsDet2.ID, dbo.TblImportShiftsDet2.MachinDate, dbo.TblImportShiftsDet2.ImportShiftID, dbo.TblImportShiftsDet2.ToTime, "
sql = sql & "                      dbo.TblImportShiftsDet2.RecTime, dbo.TblImportShiftsDet2.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
sql = sql & "                      dbo.TblImportShiftsDet2.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblImportShiftsDet2.BranchID,"
sql = sql & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblImportShiftsDet2.ProjID, dbo.projects.Project_name,"
sql = sql & "                      dbo.Projects.Project_nameE"
sql = sql & " FROM         dbo.projects RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblImportShiftsDet2 ON dbo.projects.id = dbo.TblImportShiftsDet2.ProjID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblImportShiftsDet2.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblImportShiftsDet2.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblImportShiftsDet2.DeptID = dbo.TblEmpDepartments.DeparmentID"
sql = sql & " Where (dbo.TblImportShiftsDet2.ImportShiftID =" & val(TxtSerial1.Text) & ") "
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     
     With Me.fg
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjID").value), 0, Rs1("ProjID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("MachinDate")) = IIf(IsNull(Rs1("MachinDate").value), "", Rs1("MachinDate").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                    .TextMatrix(i, .ColIndex("RecTime")) = IIf(IsNull(Rs1("RecTime").value), "", Rs1("RecTime").value)
                   .TextMatrix(i, .ColIndex("ToTime")) = IIf(IsNull(Rs1("ToTime").value), "", Rs1("ToTime").value)
                   .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   End If
                   Rs1.MoveNext
             Next i
End With
End If
        Exit Sub
ErrTrap:
    End Sub

Sub FillGriToGir()
Dim i As Integer
Dim k As Integer
  GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
With GridInstallments
.Rows = (fg.Rows * 2) + 1
For i = 1 To fg.Rows - 1
If fg.TextMatrix(i, fg.ColIndex("EmpID")) <> 0 And fg.TextMatrix(i, fg.ColIndex("MachinDate")) <> "" Then
k = k + 1
.TextMatrix(k, .ColIndex("EmpID")) = fg.TextMatrix(i, fg.ColIndex("EmpID"))
.TextMatrix(k, .ColIndex("DeptID")) = fg.TextMatrix(i, fg.ColIndex("DeptID"))
.TextMatrix(k, .ColIndex("ProjID")) = fg.TextMatrix(i, fg.ColIndex("ProjID"))
.TextMatrix(k, .ColIndex("BranchID")) = fg.TextMatrix(i, fg.ColIndex("BranchID"))
.TextMatrix(k, .ColIndex("FullCode")) = fg.TextMatrix(i, fg.ColIndex("FullCode"))
.TextMatrix(k, .ColIndex("Emp_Name")) = fg.TextMatrix(i, fg.ColIndex("Emp_Name"))
.TextMatrix(k, .ColIndex("RecTime")) = fg.TextMatrix(i, fg.ColIndex("RecTime"))
.TextMatrix(k, .ColIndex("MachinDate")) = fg.TextMatrix(i, fg.ColIndex("MachinDate"))
k = k + 1
.TextMatrix(k, .ColIndex("EmpID")) = fg.TextMatrix(i, fg.ColIndex("EmpID"))
.TextMatrix(k, .ColIndex("DeptID")) = fg.TextMatrix(i, fg.ColIndex("DeptID"))
.TextMatrix(k, .ColIndex("ProjID")) = fg.TextMatrix(i, fg.ColIndex("ProjID"))
.TextMatrix(k, .ColIndex("BranchID")) = fg.TextMatrix(i, fg.ColIndex("BranchID"))
.TextMatrix(k, .ColIndex("FullCode")) = fg.TextMatrix(i, fg.ColIndex("FullCode"))
.TextMatrix(k, .ColIndex("Emp_Name")) = fg.TextMatrix(i, fg.ColIndex("Emp_Name"))
.TextMatrix(k, .ColIndex("RecTime")) = fg.TextMatrix(i, fg.ColIndex("ToTime"))
.TextMatrix(k, .ColIndex("MachinDate")) = fg.TextMatrix(i, fg.ColIndex("MachinDate"))
End If
Next i
End With
End Sub
Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim StrSQL As String
   Dim rs2 As ADODB.Recordset

    With GridInstallments
        Select Case .ColKey(Col)
        Case "MachinCode"
        Set rs2 = New ADODB.Recordset
        StrSQL = "Select * From  TblEmployee where MachinCode ='" & .TextMatrix(Row, .ColIndex("MachinCode")) & "'"
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       If rs2.RecordCount > 0 Then
       .TextMatrix(Row, .ColIndex("ProjID")) = IIf(IsNull(rs2("project_id").value), 0, rs2("project_id").value)
       .TextMatrix(Row, .ColIndex("BranchID")) = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)
       .TextMatrix(Row, .ColIndex("DeptID")) = IIf(IsNull(rs2("DepartmentID").value), 0, rs2("DepartmentID").value)
       
       .TextMatrix(Row, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
       .TextMatrix(Row, .ColIndex("EmpID")) = IIf(IsNull(rs2("Emp_ID").value), 0, rs2("Emp_ID").value)
       If SystemOptions.UserInterface = ArabicInterface Then
       .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
       Else
       .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
       End If
       Else
       .TextMatrix(Row, .ColIndex("Emp_Name")) = ""
       .TextMatrix(Row, .ColIndex("Fullcode")) = ""
       .TextMatrix(Row, .ColIndex("EmpID")) = 0
       End If
       End Select
         If Row = .Rows - 1 Then
                     .Rows = .Rows + 1
                     End If
  End With
  
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Auto_Man = 1 Then
Cancel = True
Else
With Me.GridInstallments
Select Case .ColKey(Col)
Case "FullCode"
Cancel = True
Case "Emp_Name"
Cancel = True
Case "MachinDate"
.ComboList = ""
Case "RecTime"
.ComboList = ""
End Select
End With
End If
End Sub

Sub ExilSheet()
CD1.ShowOpen
txtFile.Text = CD1.filename
If txtFile.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Õœœ «·„·› «Ê·«"
Exit Sub
Else
MsgBox "Select File"
Exit Sub
End If
End If
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String
Dim j As Integer
Dim bol As Boolean
Dim MachinCode As String
Dim MachinDate As Date
Dim absence  As Double
Dim FromTim As String
Dim ToTim As Date
Dim tp As Variant
Dim tp2 As Variant
Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
Dim absc As String
GridInstallments.Rows = 2

    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    Do Until .cells(i, 2) & "" = ""
   MachinCode = .cells(i, 1)
    MachinDate = .cells(i, 2)
    DTPicker2.value = Format(.cells(i, 2), "DD/MM/YYYY")

tp = IIf(IsNull(.cells(i, 3)), 0, .cells(i, 3))
            FromTim = Format(.cells(i, 3), "HH:mm:SS")
 With GridInstallments
   .TextMatrix(i - 1, .ColIndex("Ser")) = i - 1
   .TextMatrix(i - 1, .ColIndex("MachinCode")) = MachinCode
   .TextMatrix(i - 1, .ColIndex("MachinDate")) = DTPicker2.value
   If tp = 0 Then
    .TextMatrix(i - 1, .ColIndex("RecTime")) = ""
    Else
    .TextMatrix(i - 1, .ColIndex("RecTime")) = FromTim
    End If
    GridInstallments_AfterEdit i - 1, .ColIndex("MachinCode")
 End With
 GridInstallments.Rows = GridInstallments.Rows + 1
        i = i + 1
    Loop

    End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
End Sub
Sub SavAbcens()
Dim ID As Double
Dim i As Integer
Dim TempEmpID  As Double
Dim NoDay As Double
Dim count As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "SELECT     TOP 100 PERCENT EmpID, Absence, ID"
sql = sql & " From dbo.TblImportShiftsDet"
sql = sql & " Where (Absence = 1)and (ImportShiftID=" & val(TxtSerial1.Text) & ")"
sql = sql & " ORDER BY EmpID"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
count = 0
Rs3.MoveFirst
For i = 1 To Rs3.RecordCount
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
If i = 1 Then

TempEmpID = IIf(IsNull(Rs3("EmpID").value), 0, Rs3("EmpID").value)
count = 1
Else
If TempEmpID <> IIf(IsNull(Rs3("EmpID").value), 0, Rs3("EmpID").value) Then
TempEmpID = IIf(IsNull(Rs3("EmpID").value), 0, Rs3("EmpID").value)
count = 1
Else
count = count + 1
End If
End If
RereiveSlice TempEmpID, count, 4, NoDay
Cn.Execute "Update TblImportShiftsDet set NetDiffOut=" & NoDay & ", TypeTransOut=4 where id=" & ID & " "
Rs3.MoveNext
Next i
End If
End Sub

Private Sub ISButton2_Click()
If RdEmp.value = True Then
If val(Me.DcbEmployee1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
Else
MsgBox "Please Select Employee"
End If
DcbEmployee1.SetFocus
Exit Sub
End If
End If

If SelectDept.value = vbChecked Then
If val(Me.DcpDept1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·«œ«—…"
Else
MsgBox "Please Select Management"
End If
DcpDept1.SetFocus
Exit Sub
End If
End If
''//////////////
If Me.SelectProject.value = vbChecked Then
If val(Me.DcbProject1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
Else
MsgBox "Please Select Project"
End If
DcbProject1.SetFocus
Exit Sub
End If
End If

If SelectBranch.value = vbChecked Then
If val(Me.DcbBranch1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbBranch1.SetFocus
Exit Sub
End If
End If
filgrid1
End Sub
Sub filgrid1()
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim i, k As Integer
Dim sql As String
sql = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, "
sql = sql & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.MachinCode, dbo.TblEmployee.DepartmentID,"
sql = sql & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.project_id, dbo.projects.Project_name,"
sql = sql & "                      dbo.Projects.Project_nameE"
sql = sql & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  Where (1 <> -1)"
If val(Me.DcbProject1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.project_id  =" & val(DcbProject1.BoundText) & " "
End If

If val(DcbBranch1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.BranchId  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.DepartmentID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.Emp_ID  =" & val(DcbEmployee1.BoundText) & " "
End If
 Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With fg
k = .Rows
Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs8("project_id").value), 0, Rs8("project_id").value)
.TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
'.TextMatrix(i, .ColIndex("MachinCode")) = IIf(IsNull(Rs8("MachinCode").value), "", Rs8("MachinCode").value)
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs8("Emp_ID").value), 0, Rs8("Emp_ID").value)
.TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DepartmentID").value), 0, Rs8("DepartmentID").value)
If SystemOptions.UserInterface = ArabicInterface Then
'.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), "", Rs8("Project_name").value)
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
'.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
'.TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
Else
'.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), "", Rs8("Project_nameE").value)
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
'.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
'.TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
End If
Rs8.MoveNext
Next i
'.AutoSize 0, .Cols - 1, False
End With
End If
End Sub
Private Sub ISButton5_Click()
print_report
End Sub


Private Sub RdEmp_Click()
If Me.RdEmp.value = False Then
Me.DcbEmployee1.BoundText = ""
End If
End Sub

Private Sub SelectBranch_Click()
If Me.SelectBranch.value = vbUnchecked Then
Me.DcbBranch1.BoundText = ""
End If
End Sub

Private Sub SelectDept_Click()
If Me.SelectDept.value = vbUnchecked Then
DcpDept1.BoundText = ""
End If
End Sub

Private Sub SelectProject_Click()
If Me.SelectProject.value = vbUnchecked Then
Me.DcbProject1.BoundText = ""
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        Me.DcbEmployee1.BoundText = EmpID
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

          StrSQL = "Delete From TblImportShiftsDet Where   ImportShiftID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From TblImportShiftsDet2 Where   ImportShiftID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
             fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
              SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
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
       ' GridInstallments.Rows = GridInstallments.Rows + 1
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
    SelectDept.value = vbUnchecked
    SelectProject.value = vbUnchecked
    SelectBranch.value = vbUnchecked
    TxtModFlg.Text = "N"
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
    fg.Clear flexClearScrollable, flexClearEverything
    fg.Rows = 1
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
'On Error GoTo ErrTrap
   ' form name
   
      Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic


  C1Tab1.Caption = "Data"
lbl(4).Caption = "No"
lbl(1).Caption = "Date"

Rd(1).Caption = "AM/PM"
Rd(0).Caption = "24 Hour"
          
           If Me.Auto_Man = 1 Then
           Label1(2).Caption = " Manual Attendance "
           Else
           Label1(2).Caption = " Import  Shifts  Attendance"
           End If
           
'Label1(2).Caption = "Imports Shifts Data "
CMDSelectFile.Caption = "Select Path"
CmdImport.Caption = "Load File"
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

  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("MachinCode")) = "Machin Code"
  .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name "
  .TextMatrix(0, .ColIndex("MachinDate")) = "Date"
  .TextMatrix(0, .ColIndex("TimIn")) = "Time Entry"
  .TextMatrix(0, .ColIndex("TimInExsist")) = "Exact Time Entry"
  .TextMatrix(0, .ColIndex("TimOut")) = "Time Out"
  .TextMatrix(0, .ColIndex("TimOutExsist")) = "Exact Time Out"
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblImportShifts"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub RemoveGridAllRow()
 GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
'    ReLineGrid
End Sub
Private Sub RemoveGridRow2()
    With Me.fg
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub

Private Sub RemoveGridRow()
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub

