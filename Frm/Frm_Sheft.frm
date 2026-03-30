VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form frm_sheft 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10980
   ClientLeft      =   3855
   ClientTop       =   3390
   ClientWidth     =   14550
   Icon            =   "Frm_Sheft.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10980
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "Frm_Sheft.frx":6852
      Left            =   15480
      List            =   "Frm_Sheft.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   26
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
      RightToLeft     =   -1  'True
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   28
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
      TabIndex        =   29
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
            Picture         =   "Frm_Sheft.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Sheft.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   30
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
      ButtonImage     =   "Frm_Sheft.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   1800
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
      ButtonImage     =   "Frm_Sheft.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   2400
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
      ButtonImage     =   "Frm_Sheft.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   10980
      Left            =   0
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   14550
      _cx             =   25665
      _cy             =   19368
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
         Caption         =   "«·‘Ì› « "
         Height          =   795
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   0
         Width           =   14535
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   38
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
            ButtonImage     =   "Frm_Sheft.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   39
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
            ButtonImage     =   "Frm_Sheft.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   40
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
            ButtonImage     =   "Frm_Sheft.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   41
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
            ButtonImage     =   "Frm_Sheft.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "  „Ã„Ê⁄«  «·—«Õ«   /  ⁄—Ì› «·‘› « "
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
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   4080
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "Frm_Sheft.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1140
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   9840
         Width           =   14550
         _cx             =   25665
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
            Height          =   345
            Left            =   13095
            TabIndex        =   44
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   615
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   345
            Left            =   11280
            TabIndex        =   45
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   615
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   345
            Left            =   9660
            TabIndex        =   46
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   615
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   345
            Left            =   7950
            TabIndex        =   47
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   615
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   345
            Left            =   6210
            TabIndex        =   48
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   615
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   5205
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   615
            Width           =   1020
            _ExtentX        =   1799
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
            ButtonImage     =   "Frm_Sheft.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   345
            Left            =   1665
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   615
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   345
            Left            =   3435
            TabIndex        =   51
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   615
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
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
            ButtonImage     =   "Frm_Sheft.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10605
            TabIndex        =   84
            Top             =   120
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
            Height          =   225
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   255
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   1
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   0
            Left            =   2505
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   285
            Index           =   8
            Left            =   13575
            TabIndex        =   85
            Top             =   120
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8805
         Left            =   0
         TabIndex        =   52
         Top             =   1140
         Width           =   14535
         _cx             =   25638
         _cy             =   15531
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
         Caption         =   "»Ì«‰«  «”«”Ì…|«·‘—«∆Õ"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   8385
            Left            =   45
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   14790
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
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   225
               Index           =   1
               Left            =   8940
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   5520
               Width           =   1230
            End
            Begin VB.CommandButton cmdInsertEmpItems 
               Caption         =   "«œ—«Ã"
               Height          =   945
               Left            =   1410
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   4530
               Width           =   3465
            End
            Begin VB.ListBox ListProductLineAll 
               Height          =   1815
               ItemData        =   "Frm_Sheft.frx":2CB3E
               Left            =   10320
               List            =   "Frm_Sheft.frx":2CB45
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   4020
               Width           =   3825
            End
            Begin VB.ListBox ListProductLineSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1815
               ItemData        =   "Frm_Sheft.frx":2CB57
               Left            =   5010
               List            =   "Frm_Sheft.frx":2CB5E
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   4020
               Width           =   3765
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   3645
               Left            =   -90
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   -60
               Width           =   14445
               _cx             =   25479
               _cy             =   6429
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
               Begin VB.TextBox TxtNoHourManaula 
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
                  Height          =   345
                  Left            =   6270
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   155
                  Top             =   3195
                  Width           =   1350
               End
               Begin XtremeSuiteControls.RadioButton RDTypHour 
                  Height          =   390
                  Index           =   0
                  Left            =   11280
                  TabIndex        =   153
                  Top             =   3195
                  Width           =   3015
                  _Version        =   786432
                  _ExtentX        =   5318
                  _ExtentY        =   688
                  _StockProps     =   79
                  Caption         =   "«Õ ”«» ⁄œœ ”«⁄«  «·⁄„· „‰ «·‘› "
                  ForeColor       =   12582912
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtNoHFri 
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
                  Height          =   360
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   2820
                  Width           =   615
               End
               Begin VB.TextBox TxtNoHThru 
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
                  Height          =   420
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   2310
                  Width           =   615
               End
               Begin VB.TextBox TxtNoHWed 
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
                  Height          =   330
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   1950
                  Width           =   615
               End
               Begin VB.TextBox TxtNoHTues 
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
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   1545
                  Width           =   615
               End
               Begin VB.TextBox TxtNoMon 
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
                  Height          =   345
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   1170
                  Width           =   615
               End
               Begin VB.TextBox TxtNoHSun 
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
                  Height          =   390
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   750
                  Width           =   615
               End
               Begin VB.TextBox TxtNoHSat 
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
                  Height          =   330
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   390
                  Width           =   615
               End
               Begin MSComCtl2.DTPicker FromMonW 
                  Height          =   345
                  Left            =   6270
                  TabIndex        =   122
                  Top             =   1155
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   609
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin VB.TextBox XPMTxtRemark 
                  Alignment       =   2  'Center
                  Height          =   420
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   22
                  Top             =   4215
                  Width           =   5700
               End
               Begin VB.ComboBox DcbThurWoVo 
                  Height          =   315
                  Left            =   12465
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   2280
                  Width           =   1140
               End
               Begin VB.ComboBox DcbTuesWoVo 
                  Height          =   315
                  Left            =   12465
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   1515
                  Width           =   1140
               End
               Begin VB.ComboBox DcbSunWoVo 
                  Height          =   315
                  Left            =   12465
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   750
                  Width           =   1140
               End
               Begin VB.ComboBox DcbFrirWoVo 
                  Height          =   315
                  Left            =   12465
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   2820
                  Width           =   1140
               End
               Begin VB.ComboBox DcbWedWoVo 
                  Height          =   315
                  Left            =   12465
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   1920
                  Width           =   1140
               End
               Begin VB.ComboBox DcbMonWoVo 
                  Height          =   315
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   1155
                  Width           =   1140
               End
               Begin VB.ComboBox DcbSatWoVo 
                  Height          =   315
                  Left            =   12465
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   375
                  Width           =   1140
               End
               Begin MSComCtl2.DTPicker FromFri 
                  Height          =   330
                  Left            =   10665
                  TabIndex        =   21
                  Top             =   2820
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromWed 
                  Height          =   330
                  Left            =   10665
                  TabIndex        =   15
                  Top             =   1920
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ShfitFrom 
                  Height          =   330
                  Left            =   10665
                  TabIndex        =   5
                  Top             =   375
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ShfitTo 
                  Height          =   330
                  Left            =   8850
                  TabIndex        =   6
                  Top             =   375
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToMon 
                  Height          =   345
                  Left            =   8850
                  TabIndex        =   11
                  Top             =   1155
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   609
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToWed 
                  Height          =   330
                  Left            =   8850
                  TabIndex        =   16
                  Top             =   1920
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToFri 
                  Height          =   330
                  Left            =   8850
                  TabIndex        =   23
                  Top             =   2820
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromThru 
                  Height          =   360
                  Left            =   10665
                  TabIndex        =   18
                  Top             =   2280
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   635
                  _Version        =   393216
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromSun 
                  Height          =   390
                  Left            =   10665
                  TabIndex        =   8
                  Top             =   750
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   688
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToSun 
                  Height          =   390
                  Left            =   8850
                  TabIndex        =   9
                  Top             =   750
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   688
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToTues 
                  Height          =   375
                  Left            =   8850
                  TabIndex        =   13
                  Top             =   1515
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   661
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToThru 
                  Height          =   360
                  Left            =   8850
                  TabIndex        =   19
                  Top             =   2280
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   635
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromMon 
                  Height          =   345
                  Left            =   10665
                  TabIndex        =   106
                  Top             =   1155
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   609
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromTues 
                  Height          =   375
                  Left            =   10665
                  TabIndex        =   107
                  Top             =   1515
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   661
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromFriW 
                  Height          =   330
                  Left            =   6270
                  TabIndex        =   108
                  Top             =   2820
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromWedW 
                  Height          =   330
                  Left            =   6270
                  TabIndex        =   109
                  Top             =   1920
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ShfitFromW 
                  Height          =   330
                  Left            =   6270
                  TabIndex        =   110
                  Top             =   375
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ShfitToW 
                  Height          =   330
                  Left            =   3495
                  TabIndex        =   111
                  Top             =   375
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToMonW 
                  Height          =   345
                  Left            =   3495
                  TabIndex        =   112
                  Top             =   1155
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   609
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToWedW 
                  Height          =   330
                  Left            =   3495
                  TabIndex        =   113
                  Top             =   1920
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToFriW 
                  Height          =   330
                  Left            =   3495
                  TabIndex        =   114
                  Top             =   2820
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromThruW 
                  Height          =   360
                  Left            =   6270
                  TabIndex        =   115
                  Top             =   2280
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   635
                  _Version        =   393216
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromSunW 
                  Height          =   390
                  Left            =   6270
                  TabIndex        =   116
                  Top             =   750
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   688
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToSunW 
                  Height          =   390
                  Left            =   3495
                  TabIndex        =   117
                  Top             =   750
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   688
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToTuesW 
                  Height          =   375
                  Left            =   3495
                  TabIndex        =   118
                  Top             =   1515
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   661
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker ToThruW 
                  Height          =   360
                  Left            =   3495
                  TabIndex        =   119
                  Top             =   2280
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   635
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker FromTuesW 
                  Height          =   375
                  Left            =   6270
                  TabIndex        =   123
                  Top             =   1515
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   661
                  _Version        =   393216
                  CustomFormat    =   "hh:mm:ss"
                  Format          =   98959362
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin ALLButtonS.ALLButton Add 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   136
                  Top             =   390
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   582
                  BTYPE           =   3
                  TX              =   " ÿ»Ìﬁ «·ﬂ·"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   2
                  FOCUSR          =   -1  'True
                  BCOL            =   65280
                  BCOLO           =   65280
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   192
                  MPTR            =   1
                  MICON           =   "Frm_Sheft.frx":2CB75
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin XtremeSuiteControls.RadioButton RDTypHour 
                  Height          =   390
                  Index           =   1
                  Left            =   7680
                  TabIndex        =   154
                  Top             =   3195
                  Width           =   3375
                  _Version        =   786432
                  _ExtentX        =   5953
                  _ExtentY        =   688
                  _StockProps     =   79
                  Caption         =   "«œŒ«· ⁄œœ ”«⁄«  «·⁄„· ÌœÊÌ ·ﬂ· «·‘› « "
                  ForeColor       =   12582912
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ… ›Ì Õ«·… «œŒ«· ⁄œœ ”«⁄«  «·⁄„· ÌœÊÌ ›Ì ‘›  Ê«Õœ ·« ÌÕ «Ã  ⁄—Ì›Â« ⁄·Ï »ﬁÌ… «·‘› « "
                  ForeColor       =   &H000080FF&
                  Height          =   375
                  Index           =   44
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   156
                  Top             =   3195
                  Width           =   6045
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   420
                  Index           =   43
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   2820
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   510
                  Index           =   42
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   2310
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   390
                  Index           =   41
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   1950
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   420
                  Index           =   40
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   1545
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   390
                  Index           =   39
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   1170
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   435
                  Index           =   38
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ”«⁄«  «·⁄„·"
                  Height          =   390
                  Index           =   37
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   360
                  Index           =   35
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   2820
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   540
                  Index           =   34
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   2280
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   375
                  Index           =   33
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   1920
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   420
                  Index           =   31
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   1515
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   390
                  Index           =   30
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   1155
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   420
                  Index           =   29
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   240
                  Index           =   25
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   2820
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   285
                  Index           =   24
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   2280
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   270
                  Index           =   0
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   1920
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   270
                  Index           =   28
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   1515
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   255
                  Index           =   27
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   1155
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   300
                  Index           =   26
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   750
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  »œ«Ì… «·‘› "
                  Height          =   270
                  Index           =   36
                  Left            =   7620
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   375
                  Width           =   1200
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Êﬁ  «‰ Â«¡ «·‘› "
                  Height          =   375
                  Index           =   32
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   375
                  Width           =   1245
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   540
                  Index           =   21
                  Left            =   9585
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   2280
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   420
                  Index           =   17
                  Left            =   9585
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   1515
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   420
                  Index           =   13
                  Left            =   9585
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   750
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   540
                  Index           =   20
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   2280
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   420
                  Index           =   16
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   750
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   420
                  Index           =   10
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   1515
                  Width           =   960
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   420
                  Index           =   5
                  Left            =   5700
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   4215
                  Width           =   1545
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Œ„Ì”"
                  Height          =   315
                  Index           =   8
                  Left            =   13320
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   2280
                  Width           =   1545
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·À·«À«¡"
                  Height          =   330
                  Index           =   6
                  Left            =   13320
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1515
                  Width           =   1545
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Õœ"
                  Height          =   330
                  Index           =   1
                  Left            =   13320
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   750
                  Width           =   1545
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   360
                  Index           =   23
                  Left            =   9525
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   2820
                  Width           =   1020
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   375
                  Index           =   19
                  Left            =   9525
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   1920
                  Width           =   1020
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   390
                  Index           =   15
                  Left            =   9525
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1155
                  Width           =   1020
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   375
                  Index           =   9
                  Left            =   9525
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   375
                  Width           =   1020
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   360
                  Index           =   22
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   2820
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   375
                  Index           =   18
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   1920
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   390
                  Index           =   14
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   1155
                  Width           =   960
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   375
                  Index           =   7
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   375
                  Width           =   960
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ã„⁄…"
                  Height          =   180
                  Index           =   9
                  Left            =   13335
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   2820
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«—»⁄«¡"
                  Height          =   300
                  Index           =   7
                  Left            =   13335
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   1920
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«À‰Ì‰"
                  Height          =   285
                  Index           =   4
                  Left            =   13335
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1155
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”» "
                  Height          =   300
                  Index           =   0
                  Left            =   13335
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   375
                  Width           =   1530
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   2310
               Left            =   150
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   6000
               Width           =   14445
               _cx             =   25479
               _cy             =   4075
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
                  Height          =   1680
                  Left            =   60
                  TabIndex        =   83
                  Top             =   30
                  Width           =   14235
                  _cx             =   25109
                  _cy             =   2963
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
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"Frm_Sheft.frx":2CB91
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
                  Height          =   285
                  Index           =   3
                  Left            =   13095
                  TabIndex        =   104
                  Top             =   2295
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   503
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
                  ButtonImage     =   "Frm_Sheft.frx":2CD0A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   285
                  Index           =   4
                  Left            =   11775
                  TabIndex        =   105
                  Top             =   2295
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   503
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
                  ButtonImage     =   "Frm_Sheft.frx":2D2A4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   135
                  Index           =   5
                  Left            =   12810
                  TabIndex        =   137
                  Top             =   2100
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   238
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
                  ButtonImage     =   "Frm_Sheft.frx":2D83E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   135
                  Index           =   6
                  Left            =   11490
                  TabIndex        =   138
                  Top             =   2100
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   238
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
                  ButtonImage     =   "Frm_Sheft.frx":2DDD8
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin XtremeSuiteControls.CheckBox BranchSelect 
               Height          =   255
               Left            =   9285
               TabIndex        =   169
               Top             =   3645
               Width           =   1020
               _Version        =   786432
               _ExtentX        =   1799
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "›—⁄ „Õœœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch 
               Height          =   315
               Left            =   6600
               TabIndex        =   170
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3645
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDepatment 
               Height          =   315
               Left            =   300
               TabIndex        =   171
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   3630
               Width           =   4620
               _ExtentX        =   8149
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox DeptSelect 
               Height          =   270
               Left            =   5160
               TabIndex        =   172
               Top             =   3660
               Width           =   1230
               _Version        =   786432
               _ExtentX        =   2170
               _ExtentY        =   476
               _StockProps     =   79
               Caption         =   "«œ«—… „Õœœ…"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label28 
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
               Left            =   8850
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   4170
               Width           =   495
            End
            Begin VB.Label Label29 
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
               Left            =   8850
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   4410
               Width           =   495
            End
            Begin VB.Label Label30 
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
               Left            =   8850
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   4770
               Width           =   495
            End
            Begin VB.Label Label31 
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
               Left            =   8850
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   5010
               Width           =   495
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   8385
            Left            =   15180
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   45
            Width           =   14445
            _cx             =   25479
            _cy             =   14790
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   3120
               Left            =   0
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   -120
               Width           =   14445
               _cx             =   25479
               _cy             =   5503
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
               AutoSizeChildren=   0
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   2055
                  Left            =   7740
                  TabIndex        =   91
                  Top             =   360
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   3625
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"Frm_Sheft.frx":2E372
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   92
                  Top             =   360
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   3625
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"Frm_Sheft.frx":2E423
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
                  Height          =   390
                  Index           =   21
                  Left            =   13080
                  TabIndex        =   100
                  Top             =   2460
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   688
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
                  ButtonImage     =   "Frm_Sheft.frx":2E4D4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   0
                  Left            =   5205
                  TabIndex        =   101
                  Top             =   2460
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   688
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
                  ButtonImage     =   "Frm_Sheft.frx":2EA6E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—«∆Õ «·«÷«›Ì -«Ã«“…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   6
                  Left            =   1650
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   120
                  Width           =   2400
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—«∆Õ «·«÷«›Ì -«·⁄„·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   5
                  Left            =   9975
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   120
                  Width           =   2415
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   3270
               Left            =   0
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   2985
               Width           =   14445
               _cx             =   25479
               _cy             =   5768
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
               AutoSizeChildren=   0
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
                  Height          =   2055
                  Left            =   7740
                  TabIndex        =   94
                  Top             =   480
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   3625
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"Frm_Sheft.frx":2F008
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   95
                  Top             =   480
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   3625
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"Frm_Sheft.frx":2F0B9
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
                  Height          =   375
                  Index           =   1
                  Left            =   13080
                  TabIndex        =   102
                  Top             =   2595
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   661
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
                  ButtonImage     =   "Frm_Sheft.frx":2F16A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   2
                  Left            =   5205
                  TabIndex        =   103
                  Top             =   2595
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   661
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
                  ButtonImage     =   "Frm_Sheft.frx":2F704
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—«∆Õ «·€Ì«»"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   12
                  Left            =   1515
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   120
                  Width           =   2445
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—«∆Õ «· «ŒÌ—"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   11
                  Left            =   10215
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   120
                  Width           =   2415
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   3270
               Left            =   0
               TabIndex        =   157
               TabStop         =   0   'False
               Top             =   6105
               Width           =   14445
               _cx             =   25479
               _cy             =   5768
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
               AutoSizeChildren=   0
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
                  Height          =   1695
                  Left            =   4620
                  TabIndex        =   158
                  Top             =   240
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   2990
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"Frm_Sheft.frx":2FC9E
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
                  Height          =   375
                  Index           =   7
                  Left            =   11040
                  TabIndex        =   159
                  Top             =   1635
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   661
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
                  ButtonImage     =   "Frm_Sheft.frx":2FD4F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   8
                  Left            =   5205
                  TabIndex        =   160
                  Top             =   2595
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   661
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
                  ButtonImage     =   "Frm_Sheft.frx":302E9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—«∆Õ «·„€«œ—…  «·„»ﬂ—…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   46
                  Left            =   11280
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   600
                  Width           =   2415
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   720
         Left            =   0
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   750
         Width           =   14565
         _cx             =   25691
         _cy             =   1270
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
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   11925
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   45
            Width           =   1425
         End
         Begin VB.TextBox XPTxtsheftName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   330
            Left            =   6810
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   45
            Width           =   4035
         End
         Begin VB.TextBox XPTxtsheftNamee 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   330
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   4425
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘›  "
            Height          =   300
            Index           =   4
            Left            =   13440
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   45
            Width           =   975
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ ⁄—»Ì"
            Height          =   300
            Index           =   1
            Left            =   10515
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   45
            Width           =   1620
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «‰Ã·Ì“Ì"
            Height          =   300
            Index           =   2
            Left            =   5460
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   45
            Width           =   1530
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
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frm_sheft"
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


Private Sub cmdReloadList_Click(Index As Integer)
    FillList
End Sub

Private Sub cmdInsertEmpItems_Click()

'FG7.Rows = 1
Dim II As Long
Dim j As Long
    If GridInstallments.Rows <= 2 Then
        If Trim(GridInstallments.TextMatrix(GridInstallments.Rows - 1, GridInstallments.ColIndex("EmpID"))) = "" Then
            GridInstallments.Rows = GridInstallments.Rows - 1
        End If
    End If
    For II = 0 To ListProductLineSelected.ListCount - 1
        
            
            If chkEmpItem(val(ListProductLineSelected.ItemData(II))) Then
                filgrid1 val(ListProductLineSelected.ItemData(II))
            End If
       
    Next II

End Sub
Private Function chkEmpItem(ByVal mEmpId As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim mEmpID2 As Long
    Dim mItemId2 As Long
    For i = 1 To GridInstallments.Rows - 1
        mEmpID2 = val(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("EmpID")))
        
        If mEmpId = mEmpID2 And mEmpID2 <> 0 Then chkEmpItem = False: Exit Function
        

    Next
     chkEmpItem = True
End Function

Private Sub DcbBranch_Click(Area As Integer)
FillList
End Sub

Private Sub DcbDepatment_Click(Area As Integer)
FillList
End Sub

Private Sub Label28_Click()
    If ListProductLineAll.ListIndex = -1 Then Exit Sub
'    ListProductLineSelected.AddItem ListProductLineAll.List(ListProductLineAll.ListIndex)
'    ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
'
    
Dim i As Long

For i = 0 To ListProductLineAll.ListCount - 1
    If ListProductLineAll.Selected(i) Then
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(i)
        
    End If
Next
    
'    FG.Rows = ListProductLineSelected.ListCount + 1
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
End Sub

Private Sub Label29_Click()
    Dim i As Integer
    ListProductLineSelected.Clear
'    FG.Rows = 1
'    FG.Rows = ListProductLineSelected.ListCount + 1
    For i = 0 To ListProductLineAll.ListCount - 1
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(i) = ListProductLineAll.ItemData(i)
'        FG.TextMatrix(i + 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'        FG.TextMatrix(i + 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
        
    Next i

End Sub

Private Sub Label30_Click()
 ListProductLineSelected.Clear
' FG.Rows = 1
End Sub

Private Sub Label31_Click()
'    If ListProductLineSelected.ListIndex > -1 Then
'      ListProductLineSelected.RemoveItem ListProductLineSelected.ListIndex
'        'FG.RemoveItem
'    End If


Dim i As Long

For i = 0 To ListProductLineSelected.ListCount - 1
    If i > ListProductLineSelected.ListCount - 1 Then Exit For
    If ListProductLineSelected.Selected(i) Then
        ListProductLineSelected.RemoveItem i
        'ListProductLineSelected.ListIndex
        i = i - 1
    End If
Next
    
End Sub
 Function ChekRepeatShift(Optional EmpID As Double, Optional Timin As String, Optional OutTime As String, Optional dY As Integer) As Boolean
 Dim sql As String
 Dim Rs3 As ADODB.Recordset
 Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblShiftWorker.EmpID, dbo.TbLSheft.*"
sql = sql & " FROM         dbo.TblShiftWorker RIGHT OUTER JOIN"
sql = sql & "                      dbo.TbLSheft ON dbo.TblShiftWorker.ShiftID = dbo.TbLSheft.SeftCode"
sql = sql & " Where (dbo.TbLSheft.SeftCode <> " & val(TxtSerial1.Text) & ") And (dbo.TblShiftWorker.EmpID = " & EmpID & ")"
Select Case dY
Case 1
sql = sql & " and ( "
sql = sql & " (  ShfitFromW <= '" & Timin & "' "
sql = sql & " and ShfitToW >= '" & Timin & "' )"
sql = sql & " or (  ShfitFromW <= '" & OutTime & "' "
sql = sql & " and ShfitToW >= '" & OutTime & "' ))"
Case 2
sql = sql & " and ( "
sql = sql & " (  FromSunW <= '" & Timin & "' "
sql = sql & " and ToSunW >= '" & Timin & "' )"
sql = sql & " or (  FromSunW <= '" & OutTime & "' "
sql = sql & " and ToSunW >= '" & OutTime & "' ))"
Case 3
sql = sql & " and ( "
sql = sql & " (  FromMonW <= '" & Timin & "' "
sql = sql & " and ToMonW >= '" & Timin & "' )"
sql = sql & " or (  FromMonW <= '" & OutTime & "' "
sql = sql & " and ToMonW >= '" & OutTime & "' ))"
Case 4
sql = sql & " and ( "
sql = sql & " (  FromTuesW <= '" & Timin & "' "
sql = sql & " and ToTuesW >= '" & Timin & "' )"
sql = sql & " or (  FromTuesW <= '" & OutTime & "' "
sql = sql & " and ToTuesW >= '" & OutTime & "' ))"
Case 5
sql = sql & " and ( "
sql = sql & " (  FromWedW <= '" & Timin & "' "
sql = sql & " and ToWedW >= '" & Timin & "' )"
sql = sql & " or (  FromWedW <= '" & OutTime & "' "
sql = sql & " and ToWedW >= '" & OutTime & "' ))"
Case 6
sql = sql & " and ( "
sql = sql & " (  FromThruW <= '" & Timin & "' "
sql = sql & " and ToThruW >= '" & Timin & "' )"
sql = sql & " or (  FromThruW <= '" & OutTime & "' "
sql = sql & " and ToThruW >= '" & OutTime & "' ))"
Case 7
sql = sql & " and ( "
sql = sql & " (  FromFriW <= '" & Timin & "' "
sql = sql & " and ToFriW >= '" & Timin & "' )"
sql = sql & " or (  FromFriW <= '" & OutTime & "' "
sql = sql & " and ToFriW >= '" & OutTime & "' ))"
End Select
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ChekRepeatShift = True
Else
ChekRepeatShift = False
End If
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
  
MySQL = "SELECT     dbo.TbLSheft.SeftCode, dbo.TbLSheft.SheftName, dbo.TbLSheft.SheftNamee, dbo.TbLSheft.Remarks, dbo.TbLSheft.ShiftFrom, dbo.TbLSheft.ShiftTo, "
MySQL = MySQL & "                       dbo.TbLSheft.ShiftTime, dbo.TbLSheft.SatWoVo, dbo.TbLSheft.SunWoVo, dbo.TbLSheft.MonWoVo, dbo.TbLSheft.TuesWoVo, dbo.TbLSheft.WedWoVo,"
MySQL = MySQL & "                      dbo.TbLSheft.ThurWoVo, dbo.TbLSheft.FrirWoVo, dbo.TbLSheft.FromSun, dbo.TbLSheft.ToSun, dbo.TbLSheft.FromMon, dbo.TbLSheft.ToMon, dbo.TbLSheft.FromTues,"
MySQL = MySQL & "                      dbo.TbLSheft.ToTues, dbo.TbLSheft.FromWed, dbo.TbLSheft.ToWed, dbo.TbLSheft.FromThru, dbo.TbLSheft.ToThru, dbo.TbLSheft.FromFri, dbo.TbLSheft.ToFri,"
MySQL = MySQL & "                      dbo.TbLSheft.BranchSelect, dbo.TbLSheft.DeptSelect, dbo.TbLSheft.EmpSelect, dbo.TbLSheft.SelectAll, dbo.TbLSheft.BranchID, TblBranchesData_1.branch_name,"
MySQL = MySQL & "                      TblBranchesData_1.branch_namee, dbo.TbLSheft.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TbLSheft.DeptID, TblEmpDepartments_1.DepartmentName,"
MySQL = MySQL & "                      TblEmpDepartments_1.DepartmentNamee, dbo.TblShiftWorker.MachinCode, dbo.TblShiftWorker.Typetrans, dbo.TblShiftWorker.FromMint, dbo.TblShiftWorker.ToMint,"
MySQL = MySQL & "                      dbo.TblShiftWorker.AverageMaint, dbo.TblShiftWorker.EmpID AS DetEmpID, TblEmployee_1.Emp_Name AS DetEmp_Name,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name1 AS DetEmp_Name1, TblEmployee_1.Emp_Name2 AS DetEmp_Name2, TblEmployee_1.Emp_Name3 AS DetEmp_Name3,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name4 AS DetEmp_Name4, TblEmployee_1.Fullcode AS DetFullcode, TblEmployee_1.Emp_Namee4 AS DetEmp_Namee4,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee3 AS DetEmp_Namee3, TblEmployee_1.Emp_Namee2 AS DetEmp_Namee2, TblEmployee_1.Emp_Namee1 AS DetEmp_Namee1,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee AS DetEmp_Namee, dbo.TblShiftWorker.BranchID AS DetBranchID, TblBranchesData_1.branch_name AS Detbranch_name,"
MySQL = MySQL & "                      TblBranchesData_1.branch_namee AS Detbranch_namee, dbo.TblShiftWorker.DeptID AS DetDeptID, TblEmpDepartments_1.DepartmentName AS DetDepartmentName,"
MySQL = MySQL & "                      TblEmpDepartments_1.DepartmentNamee AS DetDepartmentNamee, dbo.TbLSheft.ShfitFromW, dbo.TbLSheft.ShfitToW, dbo.TbLSheft.FromSunW,"
MySQL = MySQL & "                      dbo.TbLSheft.ToSunW, dbo.TbLSheft.FromMonW, dbo.TbLSheft.ToMonW, dbo.TbLSheft.FromTuesW, dbo.TbLSheft.ToTuesW, dbo.TbLSheft.FromWedW,"
MySQL = MySQL & "                      dbo.TbLSheft.ToWedW, dbo.TbLSheft.FromThruW, dbo.TbLSheft.ToThruW, dbo.TbLSheft.FromFriW, dbo.TbLSheft.ToFriW, dbo.TbLSheft.TypHour,"
MySQL = MySQL & "                      dbo.TbLSheft.NoHourManaula, dbo.TbLSheft.NoHFri, dbo.TbLSheft.NoHThru, dbo.TbLSheft.NoHWed, dbo.TbLSheft.NoHTues, dbo.TbLSheft.NoMon,"
MySQL = MySQL & "                      dbo.TbLSheft.NoHSun , dbo.TbLSheft.NoHSat, dbo.TblShiftWorker.valuee, dbo.TblShiftWorker.ProjID, dbo.Projects.Project_name, dbo.Projects.Project_nameE"
MySQL = MySQL & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblShiftWorker ON dbo.projects.id = dbo.TblShiftWorker.ProjID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments TblEmpDepartments_1 ON dbo.TblShiftWorker.DeptID = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData TblBranchesData_1 ON dbo.TblShiftWorker.BranchID = TblBranchesData_1.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblShiftWorker.EmpID = TblEmployee_1.Emp_ID ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TbLSheft.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments TblEmpDepartments_2 ON dbo.TbLSheft.DeptID = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData TblBranchesData_2 ON dbo.TbLSheft.BranchID = TblBranchesData_2.branch_id"
MySQL = MySQL & " Where (dbo.TbLSheft.SeftCode = " & val(TxtSerial1.Text) & ") "
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportShifts.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportShifts.rpt"
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

Private Sub Add_Click()
If Me.TxtModFlg.Text <> "R" Then
FromSun.value = ShfitFrom.value
ToSun.value = ShfitTo.value
FromSunW.value = ShfitFromW.value
ToSunW.value = ShfitToW.value
FromMon.value = ShfitFrom.value
ToMon.value = ShfitTo.value
FromMonW.value = ShfitFromW.value
ToMonW.value = ShfitToW.value
FromTues.value = ShfitFrom.value
ToTues.value = ShfitTo.value
FromTuesW.value = ShfitFromW.value
ToTuesW.value = ShfitToW.value
FromWed.value = ShfitFrom.value
ToWed.value = ShfitTo.value
FromWedW.value = ShfitFromW.value
ToWedW.value = ShfitToW.value
FromThru.value = ShfitFrom.value
ToThru.value = ShfitTo.value
FromThruW.value = ShfitFromW.value
ToThruW.value = ShfitToW.value
FromFri.value = ShfitFrom.value
ToFri.value = ShfitTo.value
FromFriW.value = ShfitFromW.value
ToFriW.value = ShfitToW.value
ShfitFrom_Change
FromSun_Change
FromMon_Change
FromTues_Change
FromWed_Change
FromThru_Change
FromFri_Change
End If
End Sub

Private Sub BranchSelect_Click()
If Me.BranchSelect = vbChecked Then
DcbBranch.Enabled = True
Else
DcbBranch.Enabled = False
DcbBranch.BoundText = 0
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 21
RemoveGridRow1
Case 0
RemoveGridRow2
Case 1
RemoveGridRow3
Case 2
RemoveGridRow4
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
Case 5
RemoveGridRow6
Case 6
RemoveGridAllRow
Case 7
RemoveGridRow5


End Select
End Sub

 

Private Sub DeptSelect_Click()
If Me.DeptSelect.value = vbUnchecked Then
DcbDepatment.Enabled = False
DcbDepatment.BoundText = 0
Else
DcbDepatment.Enabled = True

End If
End Sub

Private Sub EmpSelect_Click()
'If Me.EmpSelect.value = True Then
'DcbEmployee.Enabled = True
'TxtCode.Enabled = True
'Else
'DcbEmployee.Enabled = False
'TxtCode.Enabled = False
'TxtCode.Text = 0
'DcbEmployee.BoundText = 0
'End If

End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TbLSheft  order by  SeftCode "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    'Resize_Form Me
      Me.Height = 10000
  Me.Width = 17595
  
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetUsers Me.DCboUserName
   'Dcombos.GetEmployees Me.DcbEmployee
   Dcombos.GetEmpDepartments Me.DcbDepatment
   If SystemOptions.UserInterface = ArabicInterface Then
   With Me.DcbSatWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
    With Me.DcbSunWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
   With Me.DcbMonWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
    With Me.DcbWedWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
    With Me.DcbTuesWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
    With Me.DcbThurWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
    With Me.DcbFrirWoVo
   .Clear
   .AddItem "⁄„·"
   .AddItem "«Ã«“…"
   End With
   Else
    With Me.DcbFrirWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
   End With
  With Me.DcbTuesWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
   End With
     With Me.DcbThurWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
   End With
   With Me.DcbWedWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
   End With
    With Me.DcbMonWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
   End With
    With Me.DcbSunWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
   End With
   With Me.DcbSatWoVo
   .Clear
   .AddItem "Work"
   .AddItem "Holiday"
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
   ListProductLineSelected.Clear
    FillList
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                  StrSQL = "Delete From TblShiftWorker Where ShiftID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
              End If
   RsSavRec.Fields("SheftName").value = Trim(XPTxtsheftName.Text)
   RsSavRec.Fields("SheftNamee").value = Trim(XPTxtsheftNamee.Text)
   RsSavRec.Fields("Remarks").value = IIf(XPMTxtRemark.Text = "", "", Trim(XPMTxtRemark.Text))
   RsSavRec.Fields("Shiftfrom").value = FormatDateTime(Me.ShfitFrom.value, vbShortTime)
   RsSavRec.Fields("ShiftTO").value = FormatDateTime(Me.ShfitTo.value, vbShortTime)
   RsSavRec.Fields("FromSun").value = FormatDateTime(Me.FromSun.value, vbShortTime)
   RsSavRec.Fields("ToSun").value = FormatDateTime(Me.ToSun.value, vbShortTime)
   RsSavRec.Fields("FromMon").value = FormatDateTime(Me.FromMon.value, vbShortTime)
   RsSavRec.Fields("ToMon").value = FormatDateTime(Me.ToMon.value, vbShortTime)
   RsSavRec.Fields("FromTues").value = FormatDateTime(Me.FromTues.value, vbShortTime)
   RsSavRec.Fields("ToTues").value = FormatDateTime(Me.ToTues.value, vbShortTime)
   RsSavRec.Fields("FromWed").value = FormatDateTime(Me.FromWed.value, vbShortTime)
   RsSavRec.Fields("ToWed").value = FormatDateTime(Me.ToWed.value, vbShortTime)
   RsSavRec.Fields("FromThru").value = FormatDateTime(Me.FromThru.value, vbShortTime)
   RsSavRec.Fields("ToThru").value = FormatDateTime(Me.ToThru.value, vbShortTime)
   RsSavRec.Fields("FromFri").value = FormatDateTime(Me.FromFri.value, vbShortTime)
   RsSavRec.Fields("ToFri").value = FormatDateTime(Me.ToFri.value, vbShortTime)
   RsSavRec.Fields("SatWoVo").value = val(DcbSatWoVo.ListIndex)
   RsSavRec.Fields("SunWoVo").value = val(DcbSunWoVo.ListIndex)
   RsSavRec.Fields("MonWoVo").value = val(DcbMonWoVo.ListIndex)
   RsSavRec.Fields("TuesWoVo").value = val(DcbTuesWoVo.ListIndex)
   RsSavRec.Fields("WedWoVo").value = val(DcbWedWoVo.ListIndex)
   RsSavRec.Fields("ThurWoVo").value = val(DcbThurWoVo.ListIndex)
   RsSavRec.Fields("FrirWoVo").value = val(DcbFrirWoVo.ListIndex)
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("DeptID").value = val(Me.DcbDepatment.BoundText)
  ' RsSavRec.Fields("EmpID").value = val(Me.DcbEmployee.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("NoHourManaula").value = val(TxtNoHourManaula.Text)
   If RDTypHour(0).value = True Then
   RsSavRec.Fields("TypHour").value = 0
   ElseIf RDTypHour(1).value = True Then
   RsSavRec.Fields("TypHour").value = 1
   End If
'   If SelectAll.value = True Then
'   RsSavRec.Fields("SelectAll").value = 1
'   End If
'  If EmpSelect.value = True Then
'   RsSavRec.Fields("EmpSelect").value = 1
'   End If
  If DeptSelect.value = vbChecked Then
   RsSavRec.Fields("DeptSelect").value = 1
   End If
  If BranchSelect.value = vbChecked Then
   RsSavRec.Fields("BranchSelect").value = 1
   End If
   RsSavRec.Fields("ShfitFromW").value = FormatDateTime(Me.ShfitFromW.value, vbShortTime)
   RsSavRec.Fields("ShfitToW").value = FormatDateTime(Me.ShfitToW.value, vbShortTime)
   RsSavRec.Fields("FromSunW").value = FormatDateTime(Me.FromSunW.value, vbShortTime)
   RsSavRec.Fields("ToSunW").value = FormatDateTime(Me.ToSunW.value, vbShortTime)
   RsSavRec.Fields("FromMonW").value = FormatDateTime(Me.FromMonW.value, vbShortTime)
   RsSavRec.Fields("ToMonW").value = FormatDateTime(Me.ToMonW.value, vbShortTime)
   RsSavRec.Fields("FromTuesW").value = FormatDateTime(Me.FromTuesW.value, vbShortTime)
   RsSavRec.Fields("ToTuesW").value = FormatDateTime(Me.ToTuesW.value, vbShortTime)
   RsSavRec.Fields("FromWedW").value = FormatDateTime(Me.FromWedW.value, vbShortTime)
   RsSavRec.Fields("ToWedW").value = FormatDateTime(Me.ToWedW.value, vbShortTime)
   RsSavRec.Fields("FromThruW").value = FormatDateTime(Me.FromThruW.value, vbShortTime)
   RsSavRec.Fields("ToThruW").value = FormatDateTime(Me.ToThruW.value, vbShortTime)
   RsSavRec.Fields("FromFriW").value = FormatDateTime(Me.FromFriW.value, vbShortTime)
   RsSavRec.Fields("ToFriW").value = FormatDateTime(Me.ToFriW.value, vbShortTime)
   RsSavRec.Fields("NoHSat").value = val(Me.TxtNoHSat.Text)
   RsSavRec.Fields("NoHSun").value = val(Me.TxtNoHSun.Text)
   RsSavRec.Fields("NoMon").value = val(Me.TxtNoMon.Text)
   RsSavRec.Fields("NoHTues").value = val(Me.TxtNoHTues.Text)
   RsSavRec.Fields("NoHWed").value = val(Me.TxtNoHWed.Text)
   RsSavRec.Fields("NoHThru").value = val(Me.TxtNoHThru.Text)
   RsSavRec.Fields("NoHFri").value = val(Me.TxtNoHFri.Text)
   
    
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblShiftWorker Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("MachinCode").value = IIf((.TextMatrix(i, .ColIndex("MachinCode"))) = "", Null, Trim(.TextMatrix(i, .ColIndex("MachinCode"))))
       RsDevsub.update
      End If
     Next i
    End With
 ''///////////////1
 Set RsDevsub = New ADODB.Recordset
     StrSQL = "SELECT  *  from TblShiftWorker Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    With Me.VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("valuee").value = IIf((.TextMatrix(i, .ColIndex("valuee"))) = "", 0, val(.TextMatrix(i, .ColIndex("valuee"))))
                RsDevsub("ShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("FromMint").value = IIf((.TextMatrix(i, .ColIndex("FromMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("FromMint"))))
                RsDevsub("ToMint").value = IIf((.TextMatrix(i, .ColIndex("ToMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("ToMint"))))
                RsDevsub("AverageMaint").value = IIf((.TextMatrix(i, .ColIndex("AverageMaint"))) = "", 0, val(.TextMatrix(i, .ColIndex("AverageMaint"))))
                RsDevsub("Typetrans").value = 1
       RsDevsub.update
      End If
     Next i
    End With
    Set RsDevsub = New ADODB.Recordset
     ''///////////////1
     StrSQL = "SELECT  *  from TblShiftWorker Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    With Me.VSFlexGrid2
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("valuee").value = IIf((.TextMatrix(i, .ColIndex("valuee"))) = "", 0, val(.TextMatrix(i, .ColIndex("valuee"))))
                RsDevsub("FromMint").value = IIf((.TextMatrix(i, .ColIndex("FromMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("FromMint"))))
                RsDevsub("ToMint").value = IIf((.TextMatrix(i, .ColIndex("ToMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("ToMint"))))
                RsDevsub("AverageMaint").value = IIf((.TextMatrix(i, .ColIndex("AverageMaint"))) = "", 0, val(.TextMatrix(i, .ColIndex("AverageMaint"))))
                RsDevsub("Typetrans").value = 2
       RsDevsub.update
      End If
     Next i
    End With
    Set RsDevsub = New ADODB.Recordset
   ''////////////
         StrSQL = "SELECT  *  from TblShiftWorker Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    With Me.VSFlexGrid3
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("valuee").value = IIf((.TextMatrix(i, .ColIndex("valuee"))) = "", 0, val(.TextMatrix(i, .ColIndex("valuee"))))
                RsDevsub("FromMint").value = IIf((.TextMatrix(i, .ColIndex("FromMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("FromMint"))))
                RsDevsub("ToMint").value = IIf((.TextMatrix(i, .ColIndex("ToMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("ToMint"))))
                RsDevsub("AverageMaint").value = IIf((.TextMatrix(i, .ColIndex("AverageMaint"))) = "", 0, val(.TextMatrix(i, .ColIndex("AverageMaint"))))
                RsDevsub("Typetrans").value = 3
       RsDevsub.update
      End If
     Next i
    End With
    Set RsDevsub = New ADODB.Recordset
    ''////////////
         StrSQL = "SELECT  *  from TblShiftWorker Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.VSFlexGrid4
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("valuee").value = IIf((.TextMatrix(i, .ColIndex("valuee"))) = "", 0, val(.TextMatrix(i, .ColIndex("valuee"))))
                RsDevsub("FromMint").value = IIf((.TextMatrix(i, .ColIndex("FromMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("FromMint"))))
                RsDevsub("ToMint").value = IIf((.TextMatrix(i, .ColIndex("ToMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("ToMint"))))
                RsDevsub("AverageMaint").value = IIf((.TextMatrix(i, .ColIndex("AverageMaint"))) = "", 0, val(.TextMatrix(i, .ColIndex("AverageMaint"))))
                RsDevsub("Typetrans").value = 4
       RsDevsub.update
      End If
     Next i
    End With
      Set RsDevsub = New ADODB.Recordset
    ''////////////
         StrSQL = "SELECT  *  from TblShiftWorker Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.VSFlexGrid5
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ShiftID").value = val(Me.TxtSerial1.Text)
                RsDevsub("valuee").value = IIf((.TextMatrix(i, .ColIndex("valuee"))) = "", 0, val(.TextMatrix(i, .ColIndex("valuee"))))
                RsDevsub("FromMint").value = IIf((.TextMatrix(i, .ColIndex("FromMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("FromMint"))))
                RsDevsub("ToMint").value = IIf((.TextMatrix(i, .ColIndex("ToMint"))) = "", 0, val(.TextMatrix(i, .ColIndex("ToMint"))))
                RsDevsub("AverageMaint").value = IIf((.TextMatrix(i, .ColIndex("AverageMaint"))) = "", 0, val(.TextMatrix(i, .ColIndex("AverageMaint"))))
                RsDevsub("Typetrans").value = 5
       RsDevsub.update
      End If
     Next i
    End With
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
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("SeftCode").value), "", RsSavRec.Fields("SeftCode").value)
    XPTxtsheftName.Text = IIf(IsNull(RsSavRec.Fields("SheftName").value), "", RsSavRec.Fields("SheftName").value)
    XPTxtsheftNamee.Text = IIf(IsNull(RsSavRec.Fields("SheftNamee").value), "", RsSavRec.Fields("SheftNamee").value)
    DcbSatWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("SatWoVo").value), -1, RsSavRec.Fields("SatWoVo").value)
    DcbSunWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("SunWoVo").value), -1, RsSavRec.Fields("SunWoVo").value)
    DcbMonWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("MonWoVo").value), -1, RsSavRec.Fields("MonWoVo").value)
    DcbTuesWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("TuesWoVo").value), -1, RsSavRec.Fields("TuesWoVo").value)
    DcbWedWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("WedWoVo").value), -1, RsSavRec.Fields("WedWoVo").value)
    DcbThurWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("ThurWoVo").value), -1, RsSavRec.Fields("ThurWoVo").value)
    DcbFrirWoVo.ListIndex = IIf(IsNull(RsSavRec.Fields("FrirWoVo").value), -1, RsSavRec.Fields("FrirWoVo").value)
    If Not IsNull(RsSavRec("TypHour").value) Then
    If (RsSavRec("TypHour").value) = 1 Then
    RDTypHour(1).value = True
    Else
    RDTypHour(0).value = True
    End If
    Else
    RDTypHour(0).value = True
    End If
    TxtNoHourManaula.Text = IIf(IsNull(RsSavRec.Fields("NoHourManaula").value), "", RsSavRec.Fields("NoHourManaula").value)
    If Not IsNull(RsSavRec("ShiftFrom").value) Then
        Shifttime = FormatDateTime(RsSavRec("ShiftFrom").value, vbShortTime)
        Me.ShfitFrom.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ShiftTo").value) Then
        Shifttime = FormatDateTime(RsSavRec("ShiftTo").value, vbShortTime)
        Me.ShfitTo.value = Shifttime
    End If
    If Not IsNull(RsSavRec("FromSun").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromSun").value, vbShortTime)
        Me.FromSun.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToSun").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToSun").value, vbShortTime)
        Me.ToSun.value = Shifttime
    End If
      If Not IsNull(RsSavRec("FromMon").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromMon").value, vbShortTime)
        Me.FromMon.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToMon").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToMon").value, vbShortTime)
        Me.ToMon.value = Shifttime
    End If
    If Not IsNull(RsSavRec("FromTues").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromTues").value, vbShortTime)
        Me.FromTues.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToTues").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToTues").value, vbShortTime)
        Me.ToTues.value = Shifttime
    End If
     If Not IsNull(RsSavRec("FromWed").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromWed").value, vbShortTime)
        Me.FromWed.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToWed").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToWed").value, vbShortTime)
        Me.ToWed.value = Shifttime
    End If
       If Not IsNull(RsSavRec("FromThru").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromThru").value, vbShortTime)
        Me.FromThru.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToThru").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToThru").value, vbShortTime)
        Me.ToThru.value = Shifttime
    End If
    If Not IsNull(RsSavRec("FromFri").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromFri").value, vbShortTime)
        Me.FromFri.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToFri").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToFri").value, vbShortTime)
        Me.ToFri.value = Shifttime
    End If
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbDepatment.BoundText = IIf(IsNull(RsSavRec.Fields("DeptID").value), "", RsSavRec.Fields("DeptID").value)
   ' Me.DcbEmployee.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").Value), "", RsSavRec.Fields("EmpID").Value)
'    If RsSavRec("SelectAll").Value = 1 Then
'    SelectAll.Value = True
'    End If
'    If RsSavRec("EmpSelect").Value = 1 Then
'    EmpSelect.Value = True
'    End If
'    If RsSavRec("DeptSelect").Value = 1 Then
'    DeptSelect.Value = vbChecked
'    Else
'    DeptSelect.Value = vbUnchecked
'    End If
    If RsSavRec("BranchSelect").value = 1 Then
    BranchSelect.value = vbChecked
    Else
    BranchSelect.value = vbUnchecked
    End If
   ''///////////////////////
       If Not IsNull(RsSavRec("ShfitFromW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ShfitFromW").value, vbShortTime)
        Me.ShfitFromW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ShfitToW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ShfitToW").value, vbShortTime)
        Me.ShfitToW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("FromSunW").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromSunW").value, vbShortTime)
        Me.FromSunW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToSunW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToSunW").value, vbShortTime)
        Me.ToSunW.value = Shifttime
    End If
      If Not IsNull(RsSavRec("FromMonW").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromMonW").value, vbShortTime)
        Me.FromMonW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToMonW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToMonW").value, vbShortTime)
        Me.ToMonW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("FromTuesW").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromTuesW").value, vbShortTime)
        Me.FromTuesW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToTuesW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToTuesW").value, vbShortTime)
        Me.ToTuesW.value = Shifttime
    End If
     If Not IsNull(RsSavRec("FromWedW").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromWedW").value, vbShortTime)
        Me.FromWedW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToWedW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToWedW").value, vbShortTime)
        Me.ToWedW.value = Shifttime
    End If
       If Not IsNull(RsSavRec("FromThruW").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromThruW").value, vbShortTime)
        Me.FromThruW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToThruW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToThruW").value, vbShortTime)
        Me.ToThruW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("FromFriW").value) Then
        Shifttime = FormatDateTime(RsSavRec("FromFriW").value, vbShortTime)
        Me.FromFriW.value = Shifttime
    End If
    If Not IsNull(RsSavRec("ToFriW").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToFriW").value, vbShortTime)
        Me.ToFriW.value = Shifttime
    End If
    TxtNoHSat.Text = IIf(IsNull(RsSavRec.Fields("NoHSat").value), 0, RsSavRec.Fields("NoHSat").value)
    TxtNoHSun.Text = IIf(IsNull(RsSavRec.Fields("NoHSun").value), 0, RsSavRec.Fields("NoHSun").value)
    TxtNoMon.Text = IIf(IsNull(RsSavRec.Fields("NoMon").value), 0, RsSavRec.Fields("NoMon").value)
    TxtNoHTues.Text = IIf(IsNull(RsSavRec.Fields("NoHTues").value), 0, RsSavRec.Fields("NoHTues").value)
    TxtNoHWed.Text = IIf(IsNull(RsSavRec.Fields("NoHWed").value), 0, RsSavRec.Fields("NoHWed").value)
    TxtNoHThru.Text = IIf(IsNull(RsSavRec.Fields("NoHThru").value), 0, RsSavRec.Fields("NoHThru").value)
    TxtNoHFri.Text = IIf(IsNull(RsSavRec.Fields("NoHFri").value), 0, RsSavRec.Fields("NoHFri").value)
    
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
    Dim i As Integer
    Dim CtrlTxt As Control
    Dim Sm As Double
If XPTxtsheftName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· «”„ «·‘› "
Else
    MsgBox "Please enter Name"
End If
 XPTxtsheftName.SetFocus
Exit Sub
End If
If RDTypHour(1).value = True Then
If val(TxtNoHourManaula.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· ”«⁄«  «·⁄„· "
Else
    MsgBox "Please Eneter Working Hours"
End If
Exit Sub
End If
End If

With Me.GridInstallments
For i = 1 To .Rows - 1
If ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.ShfitFromW.value, vbShortTime), FormatDateTime(Me.ShfitToW.value, vbShortTime), 1) = True Then
GoTo l
ElseIf ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.FromSunW.value, vbShortTime), FormatDateTime(Me.ToSunW.value, vbShortTime), 2) = True Then
GoTo l
ElseIf ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.FromMonW.value, vbShortTime), FormatDateTime(Me.ToMonW.value, vbShortTime), 3) = True Then
GoTo l
ElseIf ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.FromTuesW.value, vbShortTime), FormatDateTime(Me.ToTuesW.value, vbShortTime), 4) = True Then
GoTo l
ElseIf ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.FromWedW.value, vbShortTime), FormatDateTime(Me.ToWedW.value, vbShortTime), 5) = True Then
GoTo l
ElseIf ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.FromThruW.value, vbShortTime), FormatDateTime(Me.ToThruW.value, vbShortTime), 6) = True Then
GoTo l
ElseIf ChekRepeatShift(val(.TextMatrix(i, .ColIndex("EmpID"))), FormatDateTime(Me.FromFriW.value, vbShortTime), FormatDateTime(Me.ToFriW.value, vbShortTime), 7) = True Then
GoTo l
End If
GoTo m
l:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "·«Ì„ﬂ‰ «·Õ›Ÿ «·„ÊŸ› "
    Msg = Msg & .TextMatrix(i, .ColIndex("Emp_Name"))
    Msg = Msg & " „”Ã·  »«ﬂÀ— „‰ ‘›  „ œ«Œ·Ì‰ ›Ì «·Êﬁ "
Else
    Msg = "Emolyee"
    Msg = Msg & .TextMatrix(i, .ColIndex("Emp_Name"))
    Msg = Msg & " Registered more than one Shifts.  have the same time"
End If
MsgBox Msg
Exit Sub
Next i
End With
m:
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
    StrRecID = new_id("TbLSheft", "SeftCode", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("SeftCode").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
   VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
   VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = 1
   VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 1
    VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid4.Rows = 1
    VSFlexGrid5.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid5.Rows = 1
            
sql = "SELECT     dbo.TblShiftWorker.ShiftID, dbo.TblShiftWorker.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
sql = sql & "                      dbo.TblShiftWorker.MachinCode, dbo.TblShiftWorker.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                      dbo.TblShiftWorker.TypeTrans , dbo.TblShiftWorker.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee"
sql = sql & " FROM         dbo.TblShiftWorker LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblShiftWorker.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblShiftWorker.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblShiftWorker.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblShiftWorker.TypeTrans Is Null)"
sql = sql & "  and (dbo.TblShiftWorker.ShiftID =" & val(TxtSerial1.Text) & ") "

  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(i, .ColIndex("MachinCode")) = IIf(IsNull(Rs1("MachinCode").value), "", Rs1("MachinCode").value)
                   .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentName").value), "", Rs1("DepartmentName").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   Else
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentNamee").value), "", Rs1("DepartmentNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
  '///////////////////////
  sql = " SELECT     ShiftID, Typetrans, FromMint, ToMint, AverageMaint,Valuee"
  sql = sql & " From dbo.TblShiftWorker"
  sql = sql & " Where (TypeTrans = 1)"
  sql = sql & "  and (ShiftID =" & val(TxtSerial1.Text) & ") "
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With Me.VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("FromMint")) = IIf(IsNull(Rs1("FromMint").value), "", Rs1("FromMint").value)
                   .TextMatrix(i, .ColIndex("ToMint")) = IIf(IsNull(Rs1("ToMint").value), 0, Rs1("ToMint").value)
                   .TextMatrix(i, .ColIndex("AverageMaint")) = IIf(IsNull(Rs1("AverageMaint").value), 0, Rs1("AverageMaint").value)
                   .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(Rs1("valuee").value), 0, Rs1("valuee").value)
                   Rs1.MoveNext
             Next i
        End With
      '///////////////////////
  sql = " SELECT     ShiftID, Typetrans, FromMint, ToMint, AverageMaint,Valuee"
  sql = sql & " From dbo.TblShiftWorker"
  sql = sql & " Where (TypeTrans = 2)"
  sql = sql & "  and (ShiftID =" & val(TxtSerial1.Text) & ") "
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     
     With Me.VSFlexGrid2
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("FromMint")) = IIf(IsNull(Rs1("FromMint").value), "", Rs1("FromMint").value)
                   .TextMatrix(i, .ColIndex("ToMint")) = IIf(IsNull(Rs1("ToMint").value), 0, Rs1("ToMint").value)
                   .TextMatrix(i, .ColIndex("AverageMaint")) = IIf(IsNull(Rs1("AverageMaint").value), 0, Rs1("AverageMaint").value)
                   .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(Rs1("valuee").value), 0, Rs1("valuee").value)
                   Rs1.MoveNext
             Next i
        End With
        '///////////////////////
  sql = " SELECT     ShiftID, Typetrans, FromMint, ToMint, AverageMaint,Valuee"
  sql = sql & " From dbo.TblShiftWorker"
  sql = sql & " Where (TypeTrans = 3)"
  sql = sql & "  and (ShiftID =" & val(TxtSerial1.Text) & ") "
  Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     
     With Me.VSFlexGrid3
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("FromMint")) = IIf(IsNull(Rs1("FromMint").value), "", Rs1("FromMint").value)
                   .TextMatrix(i, .ColIndex("ToMint")) = IIf(IsNull(Rs1("ToMint").value), 0, Rs1("ToMint").value)
                   .TextMatrix(i, .ColIndex("AverageMaint")) = IIf(IsNull(Rs1("AverageMaint").value), 0, Rs1("AverageMaint").value)
                   .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(Rs1("valuee").value), 0, Rs1("valuee").value)
                   Rs1.MoveNext
             Next i
        End With
        '///////////////////////
  sql = " SELECT     ShiftID, Typetrans, FromMint, ToMint, AverageMaint,Valuee"
  sql = sql & " From dbo.TblShiftWorker"
  sql = sql & " Where (TypeTrans = 4)"
  sql = sql & "  and (ShiftID =" & val(TxtSerial1.Text) & ") "
  Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     With Me.VSFlexGrid4
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("FromMint")) = IIf(IsNull(Rs1("FromMint").value), "", Rs1("FromMint").value)
                   .TextMatrix(i, .ColIndex("ToMint")) = IIf(IsNull(Rs1("ToMint").value), 0, Rs1("ToMint").value)
                   .TextMatrix(i, .ColIndex("AverageMaint")) = IIf(IsNull(Rs1("AverageMaint").value), 0, Rs1("AverageMaint").value)
                   .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(Rs1("valuee").value), 0, Rs1("valuee").value)
                   Rs1.MoveNext
             Next i
        End With
           '///////////////////////
  sql = " SELECT     ShiftID, Typetrans, FromMint, ToMint, AverageMaint,Valuee"
  sql = sql & " From dbo.TblShiftWorker"
  sql = sql & " Where (TypeTrans = 5)"
  sql = sql & "  and (ShiftID =" & val(TxtSerial1.Text) & ") "
  Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     With Me.VSFlexGrid5
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("FromMint")) = IIf(IsNull(Rs1("FromMint").value), "", Rs1("FromMint").value)
                   .TextMatrix(i, .ColIndex("ToMint")) = IIf(IsNull(Rs1("ToMint").value), 0, Rs1("ToMint").value)
                   .TextMatrix(i, .ColIndex("AverageMaint")) = IIf(IsNull(Rs1("AverageMaint").value), 0, Rs1("AverageMaint").value)
                   .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(Rs1("valuee").value), 0, Rs1("valuee").value)
                   Rs1.MoveNext
             Next i
        End With
        
        
        Exit Sub
ErrTrap:
    End Sub


Private Sub FromFri_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHFri.Text = DateDiff("n", FromFri.value, ToFri.value)
TxtNoHFri.Text = Round(val(TxtNoHFri.Text) / 60, 2)
If val(TxtNoHFri.Text) < 0 Then
TxtNoHFri.Text = val(TxtNoHFri.Text) + 24
End If
End If
End Sub

Private Sub FromMon_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoMon.Text = DateDiff("n", FromMon.value, ToMon.value)
TxtNoMon.Text = Round(val(TxtNoMon.Text) / 60, 2)
If val(TxtNoMon.Text) < 0 Then
TxtNoMon.Text = val(TxtNoMon.Text) + 24
End If
End If
End Sub

Private Sub FromSun_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHSun.Text = DateDiff("n", FromSun.value, ToSun.value)
TxtNoHSun.Text = Round(val(TxtNoHSun.Text) / 60, 2)
If val(TxtNoHSun.Text) < 0 Then
TxtNoHSun.Text = val(TxtNoHSun.Text) + 24
End If
End If
End Sub

Private Sub FromThru_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHThru.Text = DateDiff("n", FromThru.value, ToThru.value)
TxtNoHThru.Text = Round(val(TxtNoHThru.Text) / 60, 2)
If val(TxtNoHThru.Text) < 0 Then
TxtNoHThru.Text = val(TxtNoHThru.Text) + 24
End If
End If
End Sub

Private Sub FromTues_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHTues.Text = DateDiff("n", FromTues.value, ToTues.value)
TxtNoHTues.Text = Round(val(TxtNoHTues.Text) / 60, 2)
If val(TxtNoHTues.Text) < 0 Then
TxtNoHTues.Text = val(TxtNoHTues.Text) + 24
End If
End If
End Sub

Private Sub FromWed_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHWed.Text = DateDiff("n", FromWed.value, ToWed.value)
TxtNoHWed.Text = Round(val(TxtNoHWed.Text) / 60, 2)
If val(TxtNoHWed.Text) < 0 Then
TxtNoHWed.Text = val(TxtNoHWed.Text) + 24
End If
End If
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "MachinCode"
.ComboList = ""
Case "FullCode"
Cancel = True
Case "Emp_Name"
Cancel = True
Case "branch_name"
Cancel = True
Case "DepartmentName"
Cancel = True
End Select
End With
End Sub

Private Sub ISButton3_Click()


'If EmpSelect.value = True Then
'If val(Me.DcbEmployee.BoundText) = 0 Then
'If SystemOptions.UserInterface = ArabicInterface Then
'    MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
'Else
'    MsgBox "Please Select Employee"
'End If
'DcbEmployee.SetFocus
'Exit Sub
'End If
'End If
'If DeptSelect.value = vbChecked Then
'If val(Me.DcbDepatment.BoundText) = 0 Then
'If SystemOptions.UserInterface = ArabicInterface Then
'    MsgBox "Ì—ÃÏ «Œ Ì«— «·«œ«—…"
'Else
'    MsgBox "Please Select Management"
'End If
'DcbDepatment.SetFocus
Exit Sub
 

If BranchSelect.value = vbChecked Then
If val(Me.DcbBranch.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
Else
    MsgBox "Please Select Branch"
End If
BranchSelect.SetFocus
Exit Sub
End If
End If
filgrid1
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub RDTypHour_Click(Index As Integer)
TxtNoHourManaula.Enabled = False
If RDTypHour(1).value = True Then
TxtNoHourManaula.Enabled = True
End If
End Sub

Private Sub SelectAll_Click()
'If SelectAll.value = True Then
'DcbEmployee.Enabled = False
'TxtCode.Enabled = False
'TxtCode.Text = 0
'DcbEmployee.BoundText = 0
'End If
End Sub

Private Sub ShfitFrom_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHSat.Text = DateDiff("n", ShfitFrom.value, ShfitTo.value)
TxtNoHSat.Text = Round(val(TxtNoHSat.Text) / 60, 2)
If val(TxtNoHSat.Text) < 0 Then
TxtNoHSat.Text = val(TxtNoHSat.Text) + 24
End If
End If
End Sub

Private Sub ShfitTo_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHSat.Text = DateDiff("n", ShfitFrom.value, ShfitTo.value)
TxtNoHSat.Text = Round(val(TxtNoHSat.Text) / 60, 2)
If val(TxtNoHSat.Text) < 0 Then
TxtNoHSat.Text = val(TxtNoHSat.Text) + 24
End If
End If
End Sub

Private Sub ToFri_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHFri.Text = DateDiff("n", FromFri.value, ToFri.value)
TxtNoHFri.Text = Round(val(TxtNoHFri.Text) / 60, 2)
If val(TxtNoHFri.Text) < 0 Then
TxtNoHFri.Text = val(TxtNoHFri.Text) + 24
End If
End If
End Sub

Private Sub ToMon_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoMon.Text = DateDiff("n", FromMon.value, ToMon.value)
TxtNoMon.Text = Round(val(TxtNoMon.Text) / 60, 2)
If val(TxtNoMon.Text) < 0 Then
TxtNoMon.Text = val(TxtNoMon.Text) + 24
End If
End If
End Sub

Private Sub ToSun_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHSun.Text = DateDiff("n", FromSun.value, ToSun.value)
TxtNoHSun.Text = Round(val(TxtNoHSun.Text) / 60, 2)
If val(TxtNoHSun.Text) < 0 Then
TxtNoHSun.Text = val(TxtNoHSun.Text) + 24
End If
End If
End Sub

Private Sub ToThru_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHThru.Text = DateDiff("n", FromThru.value, ToThru.value)
TxtNoHThru.Text = Round(val(TxtNoHThru.Text) / 60, 2)
If val(TxtNoHThru.Text) < 0 Then
TxtNoHThru.Text = val(TxtNoHThru.Text) + 24
End If
End If
End Sub

Private Sub ToTues_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHTues.Text = DateDiff("n", FromTues.value, ToTues.value)
TxtNoHTues.Text = Round(val(TxtNoHTues.Text) / 60, 2)
If val(TxtNoHTues.Text) < 0 Then
TxtNoHTues.Text = val(TxtNoHTues.Text) + 24
End If
End If
End Sub

Private Sub ToWed_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoHWed.Text = DateDiff("n", FromWed.value, ToWed.value)
TxtNoHWed.Text = Round(val(TxtNoHWed.Text) / 60, 2)
If val(TxtNoHWed.Text) < 0 Then
TxtNoHWed.Text = val(TxtNoHWed.Text) + 24
End If
End If
End Sub
 

Private Sub TxtNoHourManaula_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoHourManaula.Text, 0)
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
    RsSavRec.find "SeftCode=" & RecId, , adSearchForward, 1
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
    Dim x As Integer
    Dim i As Integer
    Dim ID As Double
    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If x = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else

          StrSQL = "Delete From TblShiftWorker  Where ShiftID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "SeftCode=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = 1
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 1
            VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid4.Rows = 1
            
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
             
               '''''''''''''''''''''''''''''''

                 If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
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
   ' XPDtbTrans.Enabled = True
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
  '   XPDtbTrans.Enabled = False
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
        ShfitFrom_Change
FromSun_Change
FromMon_Change
FromTues_Change
FromWed_Change
FromThru_Change
FromFri_Change
        GridInstallments.Rows = GridInstallments.Rows + 1
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
        VSFlexGrid3.Rows = VSFlexGrid3.Rows + 1
        VSFlexGrid4.Rows = VSFlexGrid4.Rows + 1
        VSFlexGrid5.Rows = VSFlexGrid5.Rows + 1
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
'    clear_all Me
XPTxtsheftName.Text = ""
XPTxtsheftNamee.Text = ""
TxtSerial1.Text = ""
    TxtModFlg.Text = "N"
    'SelectAll.value = True
    RDTypHour(0).value = True
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
 TxtNoHourManaula.Enabled = False
If VSFlexGrid1.Rows = 1 Then
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
  VSFlexGrid1.Rows = 2
End If
If VSFlexGrid2.Rows = 1 Then
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
 End If
If VSFlexGrid3.Rows = 1 Then
    VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid3.Rows = 2
End If
If VSFlexGrid4.Rows = 1 Then
    VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid4.Rows = 2
 End If
 If VSFlexGrid5.Rows = 1 Then
    VSFlexGrid5.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid5.Rows = 2
 End If
    BranchSelect_Click
    DeptSelect_Click
'    SelectAll_Click
    Me.DCboUserName.BoundText = user_id
If DcbSatWoVo.ListIndex = -1 Then
DcbSatWoVo.ListIndex = 0
DcbSunWoVo.ListIndex = 0
DcbMonWoVo.ListIndex = 0
DcbTuesWoVo.ListIndex = 0
DcbWedWoVo.ListIndex = 0
DcbThurWoVo.ListIndex = 0
DcbFrirWoVo.ListIndex = 0
End If
'TxtSerial1
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
   ' form name
     Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    
  RDTypHour(0).RightToLeft = False
  RDTypHour(1).RightToLeft = False
  RDTypHour(0).Caption = "Calculating No. of working hours of Shafat"
  RDTypHour(1).Caption = "Manual"
   Add.Caption = "Apply All"
   lbl(46).Caption = "Early Departure"
   lbl(44).Caption = "'"
lbl(4).Caption = "No"
lbl(1).Caption = "Name Arabic"
lbl(2).Caption = "Name English"
Label1(2).Caption = "Shifts"
Label1(0).Caption = "Saturday"
Label1(1).Caption = "Sunday"
Label1(4).Caption = "Monday"
Label1(6).Caption = "Tuesday"
Label1(7).Caption = "Wednesday"
Label1(8).Caption = "Thursday"
Label1(9).Caption = "Friday"
Cmd(5).Caption = "Delete"
Cmd(6).Caption = "Delete All"
Cmd(7).Caption = "Delete"
lbl(7).Caption = "From"
lbl(14).Caption = "From"
lbl(18).Caption = "From"
lbl(22).Caption = "From"
lbl(16).Caption = "From"
lbl(10).Caption = "From"
lbl(20).Caption = "From"
Label1(5).Caption = "Remarks"
lbl(9).Caption = "To"
lbl(15).Caption = "To"
lbl(19).Caption = "To"
lbl(23).Caption = "To"
lbl(13).Caption = "To"
lbl(17).Caption = "To"
lbl(21).Caption = "To"
lbl(5).Caption = "Overtime Work Section "
lbl(6).Caption = "Overtime Holidays  Section "
lbl(11).Caption = "Delay Section "
lbl(12).Caption = "Absence Section "
lbl(3).Caption = "Data of Employee"
'SelectAll.RightToLeft = False
'SelectAll.Caption = "Select All"
'EmpSelect.RightToLeft = False
'EmpSelect.Caption = "Select Employee"
BranchSelect.Caption = "Select Branch"
BranchSelect.RightToLeft = False
DeptSelect.Caption = "Management"
DeptSelect.RightToLeft = False
Cmd(21).Caption = "Delete"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete"
Cmd(2).Caption = "Delete"
Cmd(3).Caption = "Delete"
Cmd(4).Caption = "Delete All"
C1Tab1.TabCaption(0) = "Data"
C1Tab1.TabCaption(1) = "Sections"
    
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
   With VSFlexGrid1
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("FromMint")) = "From"
   .TextMatrix(0, .ColIndex("ToMint")) = "To"
   .TextMatrix(0, .ColIndex("AverageMaint")) = "Average"
   .TextMatrix(0, .ColIndex("valuee")) = "Value"
   End With
   With VSFlexGrid2
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("FromMint")) = "From"
   .TextMatrix(0, .ColIndex("ToMint")) = "To"
   .TextMatrix(0, .ColIndex("AverageMaint")) = "Average"
   .TextMatrix(0, .ColIndex("valuee")) = "Value"
   End With
   With VSFlexGrid3
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("FromMint")) = "From"
   .TextMatrix(0, .ColIndex("ToMint")) = "To"
   .TextMatrix(0, .ColIndex("AverageMaint")) = "Average"
   .TextMatrix(0, .ColIndex("valuee")) = "Value"
   End With
   With VSFlexGrid4
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("FromMint")) = "From"
   .TextMatrix(0, .ColIndex("ToMint")) = "To"
   .TextMatrix(0, .ColIndex("AverageMaint")) = "Average"
   .TextMatrix(0, .ColIndex("valuee")) = "Value"
   End With
      With VSFlexGrid5
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("FromMint")) = "From"
   .TextMatrix(0, .ColIndex("ToMint")) = "To"
   .TextMatrix(0, .ColIndex("AverageMaint")) = "Average"
   .TextMatrix(0, .ColIndex("valuee")) = "Value"
   End With
  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("MachinCode")) = "Machin Code"
  .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name No."
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
  .TextMatrix(0, .ColIndex("DepartmentName")) = "Management"
  End With
  '##################### Khaled ######################
  lbl(36).Caption = "Shift Starts at"
  lbl(26).Caption = "Shift Starts at"
  lbl(27).Caption = "Shift Starts at"
  lbl(28).Caption = "Shift Starts at"
  lbl(0).Caption = "Shift Starts at"
  lbl(24).Caption = "Shift Starts at"
  lbl(25).Caption = "Shift Starts at"

  lbl(32).Caption = "Shift Ends at"
  lbl(29).Caption = "Shift Ends at"
  lbl(30).Caption = "Shift Ends at"
  lbl(31).Caption = "Shift Ends at"
  lbl(33).Caption = "Shift Ends at"
  lbl(34).Caption = "Shift Ends at"
  lbl(35).Caption = "Shift Ends at"
  
  lbl(37).Caption = "No. of Work Hours"
  lbl(38).Caption = "No. of Work Hours"
  lbl(39).Caption = "No. of Work Hours"
  lbl(40).Caption = "No. of Work Hours"
  lbl(41).Caption = "No. of Work Hours"
  lbl(42).Caption = "No. of Work Hours"
  lbl(43).Caption = "No. of Work Hours"
  '###################################################
  
  
ErrTrap:
End Sub

Private Sub FillList()
Dim sql As String
'sql = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId,"
'sql = sql & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.MachinCode, dbo.TblEmployee.DepartmentID,"
'sql = sql & "                       dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee"
'sql = sql & "  FROM         dbo.TblEmployee LEFT OUTER JOIN"
'sql = sql & "                       dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
'sql = sql & "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
'

'    sql = sql & " RIGHT OUTER JOIN"
'    sql = sql & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID "
    
 
 


'sql = sql & " where 1<>-1 "
'sql = sql & " and (  TBLSalesRepData.BranchId=0 or  TBLSalesRepData.BranchId is null or   TBLSalesRepData.BranchId in (  " & Current_branchSql & "))"
'If val(DcbBranch.BoundText) <> 0 Then
'sql = sql & " and dbo.TblEmployee.BranchId  =" & val(DcbBranch.BoundText) & " "
'End If
'If val(DcbDepatment.BoundText) <> 0 Then
'sql = sql & "   and dbo.TblEmployee.DepartmentID  =" & val(DcbDepatment.BoundText) & " "
'End If

sql = "SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, "
sql = sql & "                        dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.MachinCode, dbo.TblEmployee.DepartmentID,"
sql = sql & "                        dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee"
sql = sql & "  FROM         dbo.TblEmployee LEFT OUTER JOIN"
sql = sql & "                        dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                        dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
                      
   Dim i As Long
   Dim rs As New ADODB.Recordset
   rs.Open sql, Cn, adOpenKeyset, adLockReadOnly
    ListProductLineAll.Clear
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            ListProductLineAll.AddItem IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)

            ListProductLineAll.ItemData(ListProductLineAll.NewIndex) = rs("Emp_ID").value
            rs.MoveNext
        Next i
    End If
    
End Sub
Sub filgrid1(Optional ByVal mEmpId As Long = 0)
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim i, k As Integer
Dim sql As String
sql = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId,"
sql = sql & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.MachinCode, dbo.TblEmployee.DepartmentID,"
sql = sql & "                       dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee"
sql = sql & "  FROM         dbo.TblEmployee LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & " where 1<>-1 "
If val(DcbBranch.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.BranchId  =" & val(DcbBranch.BoundText) & " "
End If
If val(DcbDepatment.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.DepartmentID  =" & val(DcbDepatment.BoundText) & " "
End If
If mEmpId <> 0 Then
sql = sql & " and dbo.TblEmployee.Emp_ID  =" & mEmpId
End If
 Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With GridInstallments
k = .Rows
Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
.TextMatrix(i, .ColIndex("MachinCode")) = IIf(IsNull(Rs8("MachinCode").value), "", Rs8("MachinCode").value)
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs8("Emp_ID").value), 0, Rs8("Emp_ID").value)
.TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DepartmentID").value), 0, Rs8("DepartmentID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
Else
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
End If
Rs8.MoveNext
Next i
'.AutoSize 0, .Cols - 1, False
End With
End If
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TbLSheft"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid1
     If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid2
     If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
   IntCounter = 0
     With VSFlexGrid1
        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If

        Next i
    End With
      IntCounter = 0
     With VSFlexGrid2
        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
          IntCounter = 0
     With VSFlexGrid3
        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
           IntCounter = 0
     With VSFlexGrid4
        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
    '/////
               IntCounter = 0
     With VSFlexGrid5
        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("FromMint"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
  End Sub

Private Sub RemoveGridRow1()
    With Me.VSFlexGrid1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow2()
    With Me.VSFlexGrid2
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow3()
    With Me.VSFlexGrid3
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow4()
    With Me.VSFlexGrid4
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow6()
     With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow5()
    With Me.VSFlexGrid5
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub

Private Sub RemoveGridAllRow()
 GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
    ReLineGrid
End Sub
Private Sub RemoveGridRow()
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub VSFlexGrid3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid3
     If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub

Private Sub VSFlexGrid4_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid4
     If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub

Private Sub VSFlexGrid5_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid5
     If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub

