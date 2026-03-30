VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmBeforeInventoryK 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10665
   ClientLeft      =   1815
   ClientTop       =   1890
   ClientWidth     =   16170
   Icon            =   "FrmBeforeInventoryK.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10665
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   16170
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBeforeInventoryK.frx":6852
      Left            =   16560
      List            =   "FrmBeforeInventoryK.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   16920
      TabIndex        =   5
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
      TabIndex        =   6
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
            Picture         =   "FrmBeforeInventoryK.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBeforeInventoryK.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   16680
      TabIndex        =   7
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
      ButtonImage     =   "FrmBeforeInventoryK.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   9
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
      ButtonImage     =   "FrmBeforeInventoryK.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   10
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
      ButtonImage     =   "FrmBeforeInventoryK.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   10665
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   16170
      _cx             =   28522
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
         TabIndex        =   12
         Top             =   0
         Width           =   16140
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   15
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«⁄œ«œ«  «·ﬁÌ„… «·„÷«›… "
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
            Height          =   615
            Index           =   2
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   120
            Width           =   5160
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   15120
            Picture         =   "FrmBeforeInventoryK.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   9555
         Width           =   16170
         _cx             =   28522
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
            Left            =   14550
            TabIndex        =   21
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   1350
            _ExtentX        =   2381
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":17E16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   12510
            TabIndex        =   22
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":1E678
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   10710
            TabIndex        =   23
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   1590
            _ExtentX        =   2805
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":24EDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8820
            TabIndex        =   24
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":25274
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   6930
            TabIndex        =   25
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   1485
            _ExtentX        =   2619
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":2560E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   5820
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   600
            Width           =   1125
            _ExtentX        =   1984
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":25BA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1080
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   720
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":2C40A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   3825
            TabIndex        =   28
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   11760
            TabIndex        =   31
            Top             =   120
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   330
            Left            =   2415
            TabIndex        =   98
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·…"
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":2CB3E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1995
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   915
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2775
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   8
            Left            =   15060
            TabIndex        =   32
            Top             =   120
            Width           =   1035
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   690
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   735
         Width           =   16170
         _cx             =   28522
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
            Left            =   270
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   240
            Width           =   5910
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   12855
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   120
            Width           =   1965
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9795
            TabIndex        =   37
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   222887937
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker XPDtbTransTo 
            Height          =   315
            Left            =   7530
            TabIndex        =   43
            Top             =   240
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   222887937
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   285
            Index           =   10
            Left            =   8955
            TabIndex        =   42
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   9
            Left            =   11160
            TabIndex        =   41
            Top             =   240
            Width           =   345
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   2
            Left            =   6300
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   11205
            TabIndex        =   38
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   14895
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   1770
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1440
         Width           =   16170
         _cx             =   28522
         _cy             =   3122
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
         Begin VB.TextBox AddedValueTxt 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   13095
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1365
            Width           =   1125
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   1200
            Left            =   16245
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   2610
            _cx             =   4604
            _cy             =   2117
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
               Height          =   270
               Index           =   0
               Left            =   1485
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   525
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton Auto_Manula 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ì"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   1
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   525
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
               ForeColor       =   &H00C00000&
               Height          =   240
               Index           =   3
               Left            =   270
               TabIndex        =   49
               Top             =   15
               Width           =   2055
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   1200
            Left            =   10215
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   0
            Width           =   5925
            _cx             =   10451
            _cy             =   2117
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
               Caption         =   "—»⁄ ”‰ÊÌ"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   450
               Width           =   1035
            End
            Begin VB.OptionButton Mont_Day 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—Ì"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   450
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì Õ«·… «·Ì"
               ForeColor       =   &H00C00000&
               Height          =   45
               Index           =   5
               Left            =   1890
               TabIndex        =   54
               Top             =   1650
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œ… «Õ ”«» «·ﬁÌ„… «·„÷«›… "
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   6
               Left            =   3450
               TabIndex        =   53
               Top             =   450
               Width           =   2310
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   1200
            Left            =   -30
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   0
            Width           =   5505
            _cx             =   9710
            _cy             =   2117
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
            Begin MSDataListLib.DataCombo traTypeCombo 
               Bindings        =   "FrmBeforeInventoryK.frx":333A0
               Height          =   315
               Left            =   240
               TabIndex        =   56
               Top             =   645
               Width           =   5010
               _ExtentX        =   8837
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ—ﬂ… "
               ForeColor       =   &H00C00000&
               Height          =   240
               Index           =   7
               Left            =   1995
               TabIndex        =   57
               Top             =   255
               Width           =   1680
            End
         End
         Begin MSDataListLib.DataCombo AccDibDC 
            Bindings        =   "FrmBeforeInventoryK.frx":333B5
            Height          =   315
            Left            =   6570
            TabIndex        =   58
            Top             =   1365
            Width           =   3420
            _ExtentX        =   6033
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
         Begin MSDataListLib.DataCombo AccCirDC 
            Bindings        =   "FrmBeforeInventoryK.frx":333CA
            Height          =   315
            Left            =   480
            TabIndex        =   59
            Top             =   1365
            Width           =   3450
            _ExtentX        =   6085
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic15 
            Height          =   1200
            Left            =   5400
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   0
            Width           =   4800
            _cx             =   8467
            _cy             =   2117
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
            Begin VB.OptionButton transactions 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ—ﬂ« "
               ForeColor       =   &H00000000&
               Height          =   270
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   255
               Width           =   1275
            End
            Begin VB.OptionButton Accounts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«»« "
               ForeColor       =   &H00000000&
               Height          =   270
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   255
               Value           =   -1  'True
               Width           =   1755
            End
            Begin XtremeSuiteControls.CheckBox CheckProjAccount 
               Height          =   240
               Left            =   2040
               TabIndex        =   153
               Top             =   840
               Width           =   2070
               _Version        =   786432
               _ExtentX        =   3651
               _ExtentY        =   423
               _StockProps     =   79
               Caption         =   "—»ÿ Õ”«»«  «·„‘«—Ì⁄"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   19
            Left            =   12600
            TabIndex        =   63
            Top             =   1365
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·ﬁÌ„… «·„÷«›… œ«∆‰"
            Height          =   285
            Index           =   13
            Left            =   4050
            TabIndex        =   62
            Top             =   1365
            Width           =   2355
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·ﬁÌ„… «·„÷«›… „œÌ‰ "
            Height          =   300
            Index           =   12
            Left            =   10095
            TabIndex        =   61
            Top             =   1365
            Width           =   2475
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·ﬁÌ„… «·„÷«›…"
            Height          =   225
            Index           =   11
            Left            =   14355
            TabIndex        =   60
            Top             =   1365
            Width           =   1740
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2535
         Left            =   7410
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3240
         Width           =   8760
         _cx             =   15452
         _cy             =   4471
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   375
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   1800
            Width           =   3975
            Begin VB.OptionButton ItemStust 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„⁄›Ì"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton ItemStust 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’›—Ì"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton ItemStust 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰”»…"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   0
               Value           =   -1  'True
               Width           =   1245
            End
         End
         Begin VB.TextBox MultiPercentTxt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   360
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   345
            Width           =   780
         End
         Begin VB.OptionButton XPOptShowType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã„Ê⁄«  „Õœœ…"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   2445
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   120
            Width           =   3300
         End
         Begin VB.OptionButton XPOptShowType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’‰› „Õœœ  ≈Œ «— «·’‰›"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   6255
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1815
            Width           =   2325
         End
         Begin VB.OptionButton XPOptShowType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ· «·«’‰«›"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   7275
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   120
            Width           =   1185
         End
         Begin VB.ListBox ListGroupAll 
            Height          =   1230
            ItemData        =   "FrmBeforeInventoryK.frx":333DF
            Left            =   5010
            List            =   "FrmBeforeInventoryK.frx":333E6
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   345
            Width           =   3465
         End
         Begin VB.ListBox ListGroupSelected 
            Height          =   1230
            ItemData        =   "FrmBeforeInventoryK.frx":333F8
            Left            =   1245
            List            =   "FrmBeforeInventoryK.frx":333FF
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   345
            Width           =   3285
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   360
            Left            =   255
            TabIndex        =   79
            Top             =   1830
            Visible         =   0   'False
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   635
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":33416
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo cmbEyeDet 
            Height          =   315
            Index           =   1
            Left            =   4920
            TabIndex        =   155
            Top             =   2160
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl«”„«·ÊÕœ… 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Type"
            Height          =   330
            Index           =   10
            Left            =   5775
            TabIndex        =   156
            Top             =   2085
            Width           =   4905
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   16
            Left            =   120
            TabIndex        =   85
            Top             =   345
            Width           =   180
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰”»…"
            Height          =   270
            Index           =   14
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   120
            Width           =   975
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
            Left            =   4515
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   1020
            Width           =   510
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
            Height          =   360
            Left            =   4515
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1365
            Width           =   510
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
            Height          =   360
            Left            =   4515
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   675
            Width           =   510
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
            Left            =   4515
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   330
            Width           =   510
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2295
         Left            =   0
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3240
         Width           =   7455
         _cx             =   13150
         _cy             =   4048
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
            Height          =   240
            Left            =   6180
            TabIndex        =   65
            Top             =   120
            Width           =   1110
            _Version        =   786432
            _ExtentX        =   1958
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "ﬂ· «·„Œ«“‰"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbStore 
            Bindings        =   "FrmBeforeInventoryK.frx":337B0
            Height          =   315
            Left            =   990
            TabIndex        =   66
            Top             =   120
            Width           =   4230
            _ExtentX        =   7461
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
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   0
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":337C5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid FgStore 
            Height          =   1320
            Left            =   120
            TabIndex        =   68
            Top             =   570
            Width           =   7275
            _cx             =   12832
            _cy             =   2328
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
            FormatString    =   $"FrmBeforeInventoryK.frx":33B5F
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
            Height          =   255
            Index           =   13
            Left            =   6180
            TabIndex        =   69
            Top             =   1935
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":33BF3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   255
            Index           =   15
            Left            =   3810
            TabIndex        =   70
            Top             =   1935
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   450
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":3418D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ“‰ „Õœœ"
            Height          =   270
            Index           =   17
            Left            =   4695
            TabIndex        =   71
            Top             =   120
            Width           =   1395
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic19 
         Height          =   3450
         Left            =   0
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   6120
         Width           =   16170
         _cx             =   28522
         _cy             =   6085
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
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   3240
            Left            =   0
            TabIndex        =   118
            Top             =   0
            Width           =   16140
            _cx             =   28469
            _cy             =   5715
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
            Caption         =   "«’‰«› «·ﬁÌ„… «·„÷«›…|√’‰«› „⁄›Ì…|Õ”«»« "
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic13 
               Height          =   2820
               Left            =   17085
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   45
               Width           =   16050
               _cx             =   28310
               _cy             =   4974
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic18 
                  Height          =   2820
                  Left            =   0
                  TabIndex        =   130
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   16050
                  _cx             =   28310
                  _cy             =   4974
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
                  Begin ImpulseButton.ISButton DelRowAccGrid 
                     Height          =   285
                     Left            =   14430
                     TabIndex        =   131
                     Top             =   2385
                     Width           =   1395
                     _ExtentX        =   2461
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
                     ButtonImage     =   "FrmBeforeInventoryK.frx":34727
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton DelAllAccGrid 
                     Height          =   285
                     Left            =   12720
                     TabIndex        =   132
                     Top             =   2385
                     Width           =   1350
                     _ExtentX        =   2381
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
                     ButtonImage     =   "FrmBeforeInventoryK.frx":34CC1
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VSFlex8Ctl.VSFlexGrid AccGrid 
                     Height          =   2205
                     Left            =   135
                     TabIndex        =   133
                     Top             =   120
                     Width           =   15780
                     _cx             =   27834
                     _cy             =   3889
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
                     Cols            =   8
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmBeforeInventoryK.frx":3525B
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
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   2820
               Left            =   16785
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   45
               Width           =   16050
               _cx             =   28310
               _cy             =   4974
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic14 
                  Height          =   2955
                  Left            =   0
                  TabIndex        =   126
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   16005
                  _cx             =   28231
                  _cy             =   5212
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
                     Height          =   300
                     Index           =   0
                     Left            =   4815
                     TabIndex        =   127
                     Top             =   2535
                     Width           =   390
                     _ExtentX        =   688
                     _ExtentY        =   529
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
                     ButtonImage     =   "FrmBeforeInventoryK.frx":35399
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   300
                     Index           =   1
                     Left            =   4335
                     TabIndex        =   128
                     Top             =   2535
                     Width           =   390
                     _ExtentX        =   688
                     _ExtentY        =   529
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
                     ButtonImage     =   "FrmBeforeInventoryK.frx":35933
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                     Height          =   2385
                     Left            =   45
                     TabIndex        =   129
                     Top             =   105
                     Width           =   5220
                     _cx             =   9208
                     _cy             =   4207
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
                     Cols            =   15
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmBeforeInventoryK.frx":35ECD
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
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   2820
               Left            =   45
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   45
               Width           =   16050
               _cx             =   28310
               _cy             =   4974
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic9 
                  Height          =   2955
                  Left            =   0
                  TabIndex        =   122
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   16050
                  _cx             =   28310
                  _cy             =   5212
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
                     Height          =   300
                     Index           =   3
                     Left            =   5805
                     TabIndex        =   123
                     Top             =   2445
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   529
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
                     ButtonImage     =   "FrmBeforeInventoryK.frx":3610A
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   300
                     Index           =   4
                     Left            =   3510
                     TabIndex        =   124
                     Top             =   2445
                     Width           =   1830
                     _ExtentX        =   3228
                     _ExtentY        =   529
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
                     ButtonImage     =   "FrmBeforeInventoryK.frx":366A4
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VSFlex8Ctl.VSFlexGrid FgItem 
                     Height          =   2325
                     Left            =   45
                     TabIndex        =   125
                     Top             =   120
                     Width           =   15915
                     _cx             =   28072
                     _cy             =   4101
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
                     Cols            =   15
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmBeforeInventoryK.frx":36C3E
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
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic17 
         Height          =   2295
         Left            =   0
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   16170
         _cx             =   28522
         _cy             =   4048
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
         Begin VB.ListBox SelectedBranchList 
            Height          =   1620
            ItemData        =   "FrmBeforeInventoryK.frx":36E7B
            Left            =   1800
            List            =   "FrmBeforeInventoryK.frx":36E82
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   435
            Width           =   6735
         End
         Begin VB.ListBox BranchList 
            Height          =   1620
            ItemData        =   "FrmBeforeInventoryK.frx":36E99
            Left            =   9180
            List            =   "FrmBeforeInventoryK.frx":36EA0
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   435
            Width           =   6735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·›—Ê⁄"
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   25
            Left            =   8025
            TabIndex        =   137
            Top             =   120
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Œ «— «·›—Ê⁄"
            ForeColor       =   &H00C00000&
            Height          =   165
            Index           =   22
            Left            =   7410
            TabIndex        =   116
            Top             =   120
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   300
            Left            =   8460
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   360
            Left            =   8460
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   795
            Width           =   720
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   360
            Left            =   8595
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   300
            Left            =   8595
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   1260
            Width           =   465
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic16 
         Height          =   690
         Left            =   0
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   5520
         Visible         =   0   'False
         Width           =   16170
         _cx             =   28522
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13005
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox ForcedFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈«·“«„Ì "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   240
            Width           =   1230
         End
         Begin VB.TextBox AccPerTxt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3390
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin MSDataListLib.DataCombo AccountsDC 
            Bindings        =   "FrmBeforeInventoryK.frx":36EB2
            Height          =   315
            Left            =   7935
            TabIndex        =   104
            Top             =   240
            Width           =   4125
            _ExtentX        =   7276
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
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   270
            Left            =   360
            TabIndex        =   105
            Top             =   240
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   476
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":36EC7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkAllExp 
            Height          =   240
            Left            =   5190
            TabIndex        =   154
            Top             =   300
            Width           =   2430
            _Version        =   786432
            _ExtentX        =   4286
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "«œ—«Ã ﬂ· Õ”«»«  «·„’—Ê›« "
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·Õ”«»"
            Height          =   285
            Index           =   24
            Left            =   14775
            TabIndex        =   135
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Õ”«»"
            Height          =   285
            Index           =   23
            Left            =   11940
            TabIndex        =   108
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰”»…"
            Height          =   285
            Index           =   21
            Left            =   4170
            TabIndex        =   107
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   3060
            TabIndex        =   106
            Top             =   240
            Width           =   150
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   690
         Left            =   0
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   5520
         Width           =   16170
         _cx             =   28522
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
         Begin VB.TextBox SinglePercentTxt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4050
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtQty 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1245
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox TxtCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13455
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   240
            Width           =   1560
         End
         Begin MSDataListLib.DataCombo DcbItem 
            Bindings        =   "FrmBeforeInventoryK.frx":37261
            Height          =   315
            Left            =   8400
            TabIndex        =   90
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
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
            Bindings        =   "FrmBeforeInventoryK.frx":37276
            Height          =   315
            Left            =   6015
            TabIndex        =   91
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
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
            Left            =   255
            TabIndex        =   92
            Top             =   120
            Width           =   2160
            _ExtentX        =   3810
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
            ButtonImage     =   "FrmBeforeInventoryK.frx":3728B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   3795
            TabIndex        =   97
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰”»…"
            Height          =   285
            Index           =   15
            Left            =   5250
            TabIndex        =   96
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊÕœ…"
            Height          =   285
            Index           =   0
            Left            =   7590
            TabIndex        =   95
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·’‰›"
            Height          =   285
            Index           =   51
            Left            =   14445
            TabIndex        =   94
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ„Ì… Õœ «·ÿ·»"
            Height          =   285
            Index           =   49
            Left            =   2430
            TabIndex        =   93
            Top             =   240
            Visible         =   0   'False
            Width           =   1170
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic20 
         Height          =   2175
         Left            =   7440
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3240
         Width           =   8760
         _cx             =   15452
         _cy             =   3836
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
         Begin VB.TextBox TxtMonthNotMove 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3240
            TabIndex        =   145
            TabStop         =   0   'False
            Top             =   1020
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox TxtMonthMove 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   3240
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   630
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox TxtYearNotMove 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5400
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox TxtYearMove 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   5400
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   630
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " €Ì— „‰ﬁÊ·    «·„œ…"
            Height          =   225
            Index           =   33
            Left            =   6600
            TabIndex        =   148
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ﬁÊ·    «·„œ…"
            Height          =   225
            Index           =   32
            Left            =   6600
            TabIndex        =   147
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   240
            Index           =   30
            Left            =   4440
            TabIndex        =   143
            Top             =   1020
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   240
            Index           =   29
            Left            =   4440
            TabIndex        =   141
            Top             =   630
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«’Ê· «·À«» …"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Index           =   28
            Left            =   6240
            TabIndex        =   139
            Top             =   120
            Width           =   2475
         End
      End
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»… «·ﬁÌ„… «·„÷«›…"
      Height          =   225
      Index           =   31
      Left            =   0
      TabIndex        =   146
      Top             =   0
      Width           =   1740
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
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmBeforeInventoryK"
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
 Dim ii As Long
Private Sub AccCirDC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 888
    End If
End Sub

Private Sub AccDibDC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 666
    End If
End Sub

Private Sub Accounts_Click()
    C1Elastic2.Visible = False
    C1Elastic4.Visible = False
    C1Tab1.TabVisible(2) = True
    C1Elastic17.Visible = True
    C1Elastic16.Visible = True
    C1Tab1.TabVisible(0) = False
    C1Tab1.TabVisible(1) = False
    C1Elastic10.Visible = False
    C1Tab1.CurrTab = 2
    lbl(12).Visible = True
    AccDibDC.Visible = True
    lbl(13).Visible = True
    AccCirDC.Visible = True
    ItemStust_Click (0)
    If Accounts.value = True Then
    CheckProjAccount.Visible = True
    Else
    CheckProjAccount.Visible = False
    End If
End Sub
Private Sub AccountsDC_Change()
    AccountsDC_Click (0)
End Sub
Private Sub AccountsDC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 999
    End If
End Sub
Private Sub AccountsDC_Click(Area As Integer)
    Dim ClsAcc As New ClsAccounts
    Set ClsAcc = New ClsAccounts
    Text1.text = ClsAcc.Get_Account_Serial(AccountsDC.BoundText)
End Sub

Private Sub AddedValueTxt_Change()
   If Me.TxtModFlg <> "r " Then
    MultiPercentTxt.text = val(AddedValueTxt.text)
    End If
End Sub

Private Sub AddedValueTxt_Validate(Cancel As Boolean)
    Dim i As Long
    If Trim(AddedValueTxt) <> "" Then
    MultiPercentTxt.text = val(AddedValueTxt.text)
        For i = 1 To AccGrid.rows - 1
            AccGrid.TextMatrix(i, AccGrid.ColIndex("AddedPer")) = AddedValueTxt.text
        Next
    End If
End Sub

Private Sub Auto_Manula_Click(Index As Integer)
    If Me.Auto_Manula(0).value Then
        C1Elastic2.Enabled = False
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

Private Sub C1Tab1_Click()
 XPOptShowType_Click (0)
 ItemStust_Click (0)
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            RemoveGridRow5
        Case 1
            RemoveGridAllRow5
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
            ID = .rows
            .rows = .rows + RsDetails.RecordCount
            For i = ID To .rows - 1
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

  Dim i As Double
  Dim j As Double
  Dim k As Double
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
  bool = True
  Dim IK As Double
  
  For IK = FgStore.FixedRows To FgStore.rows - 1
       If val(FgStore.TextMatrix(IK, FgStore.ColIndex("StoreID"))) <> 0 Then
            With FgItem
                If XPOptShowType(2).value = True Or Auto_Manula(0).value = True Then
                    If val(DcbItem.BoundText) = 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
                        Else
                            MsgBox "Please Choose the item"
                        End If
                        DcbItem.SetFocus
                        Exit Sub
                    End If
                End If
                j = .rows
                If XPOptShowType(2).value = True Or XPOptShowType(0).value = True Or Auto_Manula(0).value = True Then
                    Set Rs1 = New ADODB.Recordset
                    sql = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, "
                    sql = sql & "                  dbo.TblItems.fullcode , dbo.TblItemsUnits.unitid, dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee , dbo.TblItemsUnits.UnitFactor"
                    sql = sql & "      FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
                    sql = sql & "                  dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
                    sql = sql & "                  dbo.TblItems ON dbo.TblItemsUnits.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
                    sql = sql & "                  dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"
                 
                    If XPOptShowType(2).value = True Or Auto_Manula(0).value = True Then
                If cmbEyeDet(1).BoundText = "" Then
                         sql = sql & "  Where (dbo.TblItems.ItemID =" & val(Me.DcbItem.BoundText) & ")"
                Else
                         sql = sql & "  Where (dbo.TblItems.TypeItemsID =" & val(Me.cmbEyeDet(1).BoundText) & ")"
                End If
                
                    
               
                    End If
                    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If Rs1.RecordCount > 0 Then
                        .rows = .rows + Rs1.RecordCount
                        Rs1.MoveFirst
                        For i = j To .rows - 1
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
                            If XPOptShowType(0).value = True Then
                                .TextMatrix(i, .ColIndex("UnitFactor")) = IIf(IsNull(Rs1("UnitFactor").value), 0, Rs1("UnitFactor").value)
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                                Else
                                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                                End If
                                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
                            Else
                                .TextMatrix(i, .ColIndex("UnitName")) = Me.DcbUnitDit.text
                                .TextMatrix(i, .ColIndex("UnitID")) = val(Me.DcbUnitDit.BoundText)
                                .TextMatrix(i, .ColIndex("UnitFactor")) = GetUnitFuctor(val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("UnitID"))))
                            End If
             
                            If XPOptShowType(2).value = True Then
                                .TextMatrix(i, .ColIndex("percent")) = SinglePercentTxt.text
                            ElseIf XPOptShowType(0).value = True Then
                                .TextMatrix(i, .ColIndex("percent")) = MultiPercentTxt.text
                            End If
                            Rs1.MoveNext
                        Next i
                    End If
                End If
                
                If XPOptShowType(1).value = True Then
                    Dim GROUPIDS As String
                    For k = 1 To ListGroupSelected.ListCount
                        Set Rs1 = New ADODB.Recordset
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
                            sql = sql & "  where dbo.TblItems.GroupID IN ( " & GROUPIDS & ")"
                            Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            If Rs1.RecordCount > 0 Then
                                j = .rows
                                .rows = .rows + Rs1.RecordCount
                                For i = j To .rows - 1
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
                                    .TextMatrix(i, .ColIndex("percent")) = MultiPercentTxt.text
                                    Rs1.MoveNext
                                Next i
                            End If
                        Next k
                    End If
                End With
            End If
        Next IK
        DcbItem.text = ""
        TxtCode.text = ""
        Txtqty.text = ""
End Sub
Sub Retrivetitems2()

  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
  bool = True
  Dim IK As Integer
  
  For IK = FgStore.FixedRows To FgStore.rows - 1
        If val(FgStore.TextMatrix(IK, FgStore.ColIndex("StoreID"))) <> 0 Then
            With VSFlexGrid1
                If XPOptShowType(2).value = True Or Auto_Manula(0).value = True Then
                    If val(DcbItem.BoundText) = 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
                        Else
                            MsgBox "Please Choose the item"
                        End If
                        DcbItem.SetFocus
                        Exit Sub
                    End If
                End If
                j = .rows
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
                    End If
                    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If Rs1.RecordCount > 0 Then
                        .rows = .rows + Rs1.RecordCount
                        Rs1.MoveFirst
                        For i = j To .rows - 1
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
                            If XPOptShowType(0).value = True Then
                                .TextMatrix(i, .ColIndex("UnitFactor")) = IIf(IsNull(Rs1("UnitFactor").value), 0, Rs1("UnitFactor").value)
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                                Else
                                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                                End If
                                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
                            Else
                                .TextMatrix(i, .ColIndex("UnitName")) = Me.DcbUnitDit.text
                                .TextMatrix(i, .ColIndex("UnitID")) = val(Me.DcbUnitDit.BoundText)
                                .TextMatrix(i, .ColIndex("UnitFactor")) = GetUnitFuctor(val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("UnitID"))))
                            End If
'                            .TextMatrix(i, .ColIndex("Qty")) = val(Me.TxtQty.Text)
                            If Mont_Day(1).value = True Then
'                                .TextMatrix(i, .ColIndex("Mont_Day")) = 2
                            ElseIf Mont_Day(2).value = True Then
                                .TextMatrix(i, .ColIndex("Mont_Day")) = 3
                            Else
                                .TextMatrix(i, .ColIndex("Mont_Day")) = 1
                            End If
             
                            If XPOptShowType(2).value = True Then
                                .TextMatrix(i, .ColIndex("percent")) = SinglePercentTxt.text
                            ElseIf XPOptShowType(0).value = True Then
                                .TextMatrix(i, .ColIndex("percent")) = MultiPercentTxt.text
                            End If
                        Rs1.MoveNext
                    Next i
                End If
            End If
            If XPOptShowType(1).value = True Then
                Dim GROUPIDS As String
                For k = 1 To ListGroupSelected.ListCount
                    Set Rs1 = New ADODB.Recordset
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
                        sql = sql & "  where dbo.TblItems.GroupID IN ( " & GROUPIDS & ")"
                        Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                        If Rs1.RecordCount > 0 Then
                            j = .rows
                            .rows = .rows + Rs1.RecordCount
                            For i = j To .rows - 1
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
'                                .TextMatrix(i, .ColIndex("Qty")) = val(Me.TxtQty.Text)
                                If Mont_Day(1).value = True Then
'                                    .TextMatrix(i, .ColIndex("Mont_Day")) = 2
                                ElseIf Mont_Day(2).value = True Then
                                    .TextMatrix(i, .ColIndex("Mont_Day")) = 3
                                Else
                                    .TextMatrix(i, .ColIndex("Mont_Day")) = 1
                                End If
                                .TextMatrix(i, .ColIndex("percent")) = MultiPercentTxt.text
                            Rs1.MoveNext
                        Next i
                    End If
                Next k
            End If
        End With
    End If
Next IK
    DcbItem.text = ""
    TxtCode.text = ""
    Txtqty.text = ""
End Sub
Private Sub DcbItem_Change()
    DcbItem_Click (0)
End Sub
Private Sub DcbItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmItemSearch.show
        FrmItemSearch.RetrunType = 666
    End If
End Sub
Private Sub DcbItem_Click(Area As Integer)
    
    Dim UnitName As String
    Dim UnitID As Long
    Dim Dcombos As New ClsDataCombos

    Me.TxtCode.text = GetItemCode(val(Me.DcbItem.BoundText))
    Dcombos.GetItemsUnitsDetai DcbUnitDit, val(DcbItem.BoundText)
    GetDefaultItemUnit val(Me.DcbItem.BoundText), UnitID, UnitName
    DcbUnitDit.text = UnitName
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

Private Sub ISButton4_Click()
    TxtSerial1.text = ""
    TxtModFlg.text = "N"
    btnSave.Enabled = True
End Sub

Private Sub ItemStust_Click(Index As Integer)
VSFlexGrid1.ColHidden(VSFlexGrid1.ColIndex("percent")) = False
If Me.ItemStust(0).value = True And Accounts.value = False Then
C1Tab1.CurrTab = 1
VSFlexGrid1.ColHidden(VSFlexGrid1.ColIndex("percent")) = True
MultiPercentTxt.Visible = False
lbl(14).Visible = False
lbl(21).Visible = False
lbl(20).Visible = False
lbl(16).Visible = False
AccPerTxt.Visible = False
ElseIf Accounts.value = False Then
MultiPercentTxt.Visible = True
lbl(14).Visible = True
lbl(21).Visible = True
lbl(20).Visible = True
AccPerTxt.Visible = True
C1Tab1.CurrTab = 0
Else
C1Tab1.CurrTab = 2
MultiPercentTxt.Visible = True
lbl(14).Visible = True
lbl(21).Visible = True
lbl(20).Visible = True
AccPerTxt.Visible = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If Text1.text = "" Then
            Me.AccountsDC.BoundText = ""
        Else
            Me.AccountsDC.BoundText = GetAccountCode(Trim$(Me.Text1.text))
        End If
End If
End Sub
Private Sub transactions_Click()
    C1Tab1.TabVisible(2) = False
    C1Elastic17.Visible = False
    C1Elastic16.Visible = False
    C1Elastic4.Visible = True
    C1Elastic2.Visible = True
    C1Tab1.TabVisible(0) = True
    C1Tab1.TabVisible(1) = True
    C1Elastic10.Visible = True
    C1Tab1.CurrTab = 0
    traTypeCombo_Change
    ItemStust_Click (0)
     If Accounts.value = True Then
    CheckProjAccount.Visible = True
    Else
    CheckProjAccount.Visible = False
    End If
End Sub
Private Sub traTypeCombo_Change()
C1Elastic20.Visible = False
C1Elastic17.Visible = True
C1Elastic2.Visible = True
If val(traTypeCombo.BoundText) = 0 Then Exit Sub
    If transactions.value = True Then
     If val(traTypeCombo.BoundText) = 11 Then
      C1Elastic20.Visible = True
      C1Elastic17.Visible = False
      C1Elastic2.Visible = False
     End If
     
     
        If val(traTypeCombo.BoundText) = 52 Or val(traTypeCombo.BoundText) = 2 Or val(traTypeCombo.BoundText) = 7 Or val(traTypeCombo.BoundText) = 11 Or val(traTypeCombo.BoundText) = 17 Or val(traTypeCombo.BoundText) = 22 Or val(traTypeCombo.BoundText) = 9 Or val(traTypeCombo.BoundText) = 29 Or val(traTypeCombo.BoundText) = 30 Or val(traTypeCombo.BoundText) = 42 Or val(traTypeCombo.BoundText) = 14 Or val(traTypeCombo.BoundText) = 44 Or val(traTypeCombo.BoundText) = 47 Then
            lbl(12).Visible = True
            AccDibDC.Visible = True
            lbl(13).Visible = False
            AccCirDC.Visible = False
            AccCirDC.BoundText = ""
        ElseIf val(traTypeCombo.BoundText) = 23 Or val(traTypeCombo.BoundText) = 45 Or val(traTypeCombo.BoundText) = 46 Then
             lbl(12).Visible = True
            AccDibDC.Visible = True
            lbl(13).Visible = True
            AccCirDC.Visible = True
        Else
'            lbl(12).Visible = False
'            AccDibDC.Visible = False
            AccDibDC.BoundText = ""
            lbl(13).Visible = True
            AccCirDC.Visible = True
        End If
    Else
        lbl(12).Visible = True
        AccDibDC.Visible = True
        lbl(13).Visible = True
        AccCirDC.Visible = True
    End If
End Sub

Private Sub traTypeCombo_Click(Area As Integer)
    traTypeCombo_Change
End Sub

Private Sub TxtYearMove_Change()
If Me.TxtModFlg.text <> "R" Then
TxtMonthMove.text = val(Me.TxtYearMove.text) * 12
End If
End Sub

Private Sub TxtYearMove_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtYearMove.text, 0)
End Sub

Private Sub TxtYearNotMove_Change()
If Me.TxtModFlg.text <> "R" Then
TxtMonthNotMove.text = val(Me.TxtYearNotMove.text) * 12
End If
End Sub

Private Sub TxtYearNotMove_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtYearNotMove.text, 0)
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Auto_Manula(1).value = True Then
        If VSFlexGrid1.ColKey(Col) = "OperReQst" Then
            VSFlexGrid1.ComboList = ""
        Else
            Cancel = True
        End If
    Else
        If VSFlexGrid1.ColKey(Col) = "OperReQst" Then
            VSFlexGrid1.ComboList = ""
        End If
        Cancel = False
    End If
End Sub
Private Sub Form_Load()
    
    On Error GoTo ErrTrap
    
    Dim conection As String
    Dim My_SQL As String

    conection = "select * from  TblSettsReqLimK  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    FillMylist
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetStores Me.DcbStore
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetItemsNamesupdate Me.DcbItem
    Dim nstr As String
   nstr = " SELECT     ID, Namee from tblTypeItems"
fill_combo cmbEyeDet(1), nstr



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
    
 
    
    
    Dcombos.GetAccountingCodes AccDibDC, True
    Dcombos.GetAccountingCodes AccCirDC, True
    
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,VatTypeName From VatTypes "
        StrSQL = StrSQL + " Order By VatTypeName"
    Else
        StrSQL = "SELECT ID,(VatTypeNamee) From VatTypes "
        StrSQL = StrSQL + " Order By VatTypeNamee"
    End If
    
    fill_combo traTypeCombo, StrSQL
     
'############################################################# Accounts Part ##################################################
    C1Tab1.TabVisible(2) = False
    Dcombos.GetAccountingCodes AccountsDC, False
    FillBranchList
'##############################################################################################################################
    Me.Refresh
    FiLLTXT
  ' AccDibDC.BoundText = 0
ErrTrap:
End Sub
Public Sub FiLLRec()
    
      On Error GoTo ErrTrap
    
    Dim sql As String
    Dim ID As Double
    
    If Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete From TblSettsReqLimKDet Where SetReqLID=" & val(Me.TxtSerial1.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
   If CheckProjAccount.value = vbChecked Then
     RsSavRec.Fields("ProjAccount").value = 1
    Else
     RsSavRec.Fields("ProjAccount").value = 0
    End If
    RsSavRec.Fields("MultiPercentTxt").value = val(MultiPercentTxt.text)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecordDateTo").value = XPDtbTransTo.value
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("Remarks").value = txtRemarks.text
    RsSavRec.Fields("StoreID").value = val(Me.DcbStore.BoundText)
    RsSavRec.Fields("ItemID").value = val(Me.DcbItem.BoundText)
    RsSavRec.Fields("UnitID").value = val(Me.DcbUnitDit.BoundText)
    RsSavRec.Fields("Qty").value = val(Me.Txtqty.text)
    RsSavRec.Fields("YearMove").value = val(Me.TxtYearMove.text)
    RsSavRec.Fields("YearNotMove").value = val(Me.TxtYearNotMove.text)
    RsSavRec.Fields("MonthMove").value = val(Me.TxtMonthMove.text)
    RsSavRec.Fields("MonthNotMove").value = val(Me.TxtMonthNotMove.text)
    If Mont_Day(1).value = True Then
        RsSavRec.Fields("Mont_Day").value = 1
    Else
        RsSavRec.Fields("Mont_Day").value = 0
    End If
   
    If Auto_Manula(1).value = True Then
        RsSavRec.Fields("Auto_Manula").value = 1
    ElseIf Auto_Manula(2).value = True Then
        RsSavRec.Fields("Auto_Manula").value = 2
    Else
        RsSavRec.Fields("Auto_Manula").value = 0
    End If
        If ItemStust(1).value = True Then
        RsSavRec.Fields("ItemStust").value = 1
    ElseIf ItemStust(2).value = True Then
        RsSavRec.Fields("ItemStust").value = 2
    Else
        RsSavRec.Fields("ItemStust").value = 0
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
    ElseIf XPOptShowType(2).value = True Then
        RsSavRec.Fields("SelectType").value = 2
    End If
    
    If traTypeCombo.BoundText <> "" Then
        RsSavRec.Fields("TransType").value = traTypeCombo.BoundText
    Else
        RsSavRec.Fields("TransType").value = 0
    End If
  
    RsSavRec.Fields("PercentH").value = val(AddedValueTxt.text)
    RsSavRec.Fields("AccDep").value = AccDibDC.BoundText
    RsSavRec.Fields("AccCir").value = AccCirDC.BoundText
   
   '################################# Account Part ###############################
   If Accounts.value = True Then
        RsSavRec.Fields("AccOrTran").value = 0
   ElseIf transactions.value = True Then
        RsSavRec.Fields("AccOrTran").value = 1
   End If
   
   '##############################################################################

    RsSavRec.update
    
    Dim TransType As String
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSettsReqLimKDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Double
    With Me.FgItem
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("SetReqLID").value = val(Me.TxtSerial1.text)
                RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val((.TextMatrix(i, .ColIndex("StoreID")))))
                RsDevsub("UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
                RsDevsub("UnitFactor").value = IIf((.TextMatrix(i, .ColIndex("UnitFactor"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitFactor"))))
                RsDevsub("GroupID").value = IIf((.TextMatrix(i, .ColIndex("GroupID"))) = "", Null, val(.TextMatrix(i, .ColIndex("GroupID"))))
                RsDevsub("ConsuRateLowQty").value = val(.TextMatrix(i, .ColIndex("ConsuRateLowQty")))
                RsDevsub("PercentD").value = val(.TextMatrix(i, .ColIndex("percent")))
                RsDevsub("Typ").value = 0
                RsDevsub.update
            End If
        Next i
    End With

'##########################################################################################################################################################################################################
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSettsReqLimKDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("SetReqLID").value = val(Me.TxtSerial1.text)
                RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val((.TextMatrix(i, .ColIndex("StoreID")))))
                RsDevsub("UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
                RsDevsub("UnitFactor").value = IIf((.TextMatrix(i, .ColIndex("UnitFactor"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitFactor"))))
                RsDevsub("GroupID").value = IIf((.TextMatrix(i, .ColIndex("GroupID"))) = "", Null, val(.TextMatrix(i, .ColIndex("GroupID"))))
'                RsDevsub("Mont_Day").value = IIf((.TextMatrix(i, .ColIndex("Mont_Day"))) = "", Null, val(.TextMatrix(i, .ColIndex("Mont_Day"))))
'                RsDevsub("Priod").value = IIf((.TextMatrix(i, .ColIndex("Priod"))) = "", Null, val(.TextMatrix(i, .ColIndex("Priod"))))
'                RsDevsub("PriodType").value = IIf((.TextMatrix(i, .ColIndex("PriodType"))) = "", Null, val(.TextMatrix(i, .ColIndex("PriodType"))))
'                RsDevsub("SafetyRate").value = IIf((.TextMatrix(i, .ColIndex("SafetyRate"))) = "", Null, val(.TextMatrix(i, .ColIndex("SafetyRate"))))
                
                If Auto_Manula(1).value = True Then
'                    .TextMatrix(i, .ColIndex("ConsuRate")) = Round(val(.TextMatrix(i, .ColIndex("ConsuRateLowQty"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
'                    .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("Minimum"))) + (val(.TextMatrix(i, .ColIndex("ConsuRate"))) * val(.TextMatrix(i, .ColIndex("Priod"))))
'                    .TextMatrix(i, .ColIndex("Maximum")) = val(.TextMatrix(i, .ColIndex("Qty"))) + val(.TextMatrix(i, .ColIndex("Minimum")))
                End If
                
                If val(.TextMatrix(i, .ColIndex("UnitFactor"))) <> 0 Then
'                    RsDevsub("MaxLowQty").value = Round(val(.TextMatrix(i, .ColIndex("Maximum"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
'                    RsDevsub("MinLowQty").value = Round(val(.TextMatrix(i, .ColIndex("Minimum"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
                    RsDevsub("ConsuRate").value = Round(val(.TextMatrix(i, .ColIndex("ConsuRateLowQty"))) / val(.TextMatrix(i, .ColIndex("UnitFactor"))), 2)
                End If
                
                RsDevsub("ConsuRateLowQty").value = val(.TextMatrix(i, .ColIndex("ConsuRateLowQty")))
'                RsDevsub("Minimum").value = val(.TextMatrix(i, .ColIndex("Minimum")))
'                RsDevsub("Qty").value = val(.TextMatrix(i, .ColIndex("Qty")))
'                RsDevsub("Maximum").value = val(.TextMatrix(i, .ColIndex("Maximum")))
'                RsDevsub("OperReQst").value = val(.TextMatrix(i, .ColIndex("OperReQst")))
                RsDevsub("PercentD").value = val(.TextMatrix(i, .ColIndex("percent")))
                
                RsDevsub("Typ").value = 5
                RsDevsub.update
            End If
        Next i
    End With
'################################################################################################################################################################################################################
    
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSettsReqLimKDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    With Me.AccGrid
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("FullCode")) <> "" And val(.TextMatrix(i, .ColIndex("BranchID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("SetReqLID").value = val(Me.TxtSerial1.text)
                RsDevsub("Account_Code").value = .TextMatrix(i, .ColIndex("AccCode"))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val((.TextMatrix(i, .ColIndex("BranchID")))))
                RsDevsub("PercentD").value = IIf((.TextMatrix(i, .ColIndex("AddedPer"))) = "", Null, val(.TextMatrix(i, .ColIndex("AddedPer"))))
                
                If .cell(flexcpChecked, i, .ColIndex("ForcedFlg")) = flexChecked Then  '.TextMatrix(i, .ColIndex("ForcedFlg")) = flexChecked Then
                    RsDevsub("ForcedFlg").value = True
                ElseIf .cell(flexcpChecked, i, .ColIndex("ForcedFlg")) = flexUnchecked Then
                    RsDevsub("ForcedFlg").value = False
                End If
                
                RsDevsub("Typ").value = 9
                RsDevsub.update
            End If
        Next i
    End With
'################################################################################################################################################################################################################

    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSettsReqLimKDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.FgStore
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("StoreID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("SetReqLID").value = val(Me.TxtSerial1.text)
                RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val((.TextMatrix(i, .ColIndex("StoreID")))))
                RsDevsub("Typ").value = 1
                RsDevsub.update
            End If
        Next i
    End With
    
'################################################################################################################################################################################################################

    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblSettsReqLimKDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 0 To ListGroupSelected.ListCount - 1
        RsDevsub.AddNew
        RsDevsub("SetReqLID").value = val(Me.TxtSerial1.text)
        RsDevsub("GroupID").value = val(ListGroupSelected.ItemData(i))
        RsDevsub("Typ").value = 2
        RsDevsub.update
    Next i
    
'#################################################################################################################################################################################################################
    
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT * from dbo.TblSettsReqLimKDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 0 To SelectedBranchList.ListCount - 1
        RsDevsub.AddNew
        RsDevsub("SetReqLID").value = val(Me.TxtSerial1.text)
        RsDevsub("BranchID").value = val(SelectedBranchList.ItemData(i))
        RsDevsub("Typ").value = 8
        RsDevsub.update
    Next i

'#################################################################################################################################################################################################################
    RsDevsub.Close
    Dim Msg As String
    Select Case Me.TxtModFlg.text
        Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                If SystemOptions.UserInterface = ArabicInterface Then
                Else
                    Me.Refresh
                    FiLLTXT
                    TxtModFlg = "R"
                    MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
        Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
Public Sub FiLLTXT()
    
   ' On Error GoTo ErrTrap
    
    Dim i As Integer
    Dim Shifttime As Date
    
    If RsSavRec.RecordCount = 0 Then Exit Sub
    
    MultiPercentTxt.text = IIf(IsNull(RsSavRec.Fields("MultiPercentTxt").value), 0, RsSavRec.Fields("MultiPercentTxt").value)
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    XPDtbTransTo.value = IIf(IsNull(RsSavRec.Fields("RecordDateTo").value), Date, RsSavRec.Fields("RecordDateTo").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbStore.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), "", RsSavRec.Fields("StoreID").value)
    Me.DcbItem.BoundText = IIf(IsNull(RsSavRec.Fields("ItemID").value), "", RsSavRec.Fields("ItemID").value)
    Me.DcbUnitDit.BoundText = IIf(IsNull(RsSavRec.Fields("UnitID").value), "", RsSavRec.Fields("UnitID").value)
    Txtqty.text = IIf(IsNull(RsSavRec.Fields("Qty").value), 0, RsSavRec.Fields("Qty").value)
    traTypeCombo.BoundText = IIf(IsNull(RsSavRec.Fields("TransType").value), "", RsSavRec.Fields("TransType").value)
    AddedValueTxt.text = IIf(IsNull(RsSavRec.Fields("PercentH").value), "", RsSavRec.Fields("PercentH").value)
    AccDibDC.BoundText = IIf(IsNull(RsSavRec.Fields("AccDep").value), "", RsSavRec.Fields("AccDep").value)
    AccCirDC.BoundText = IIf(IsNull(RsSavRec.Fields("AccCir").value), "", RsSavRec.Fields("AccCir").value)
    TxtYearMove.text = IIf(IsNull(RsSavRec.Fields("YearMove").value), "", RsSavRec.Fields("YearMove").value)
    TxtYearNotMove.text = IIf(IsNull(RsSavRec.Fields("YearNotMove").value), 0, RsSavRec.Fields("YearNotMove").value)
    TxtMonthMove.text = IIf(IsNull(RsSavRec.Fields("MonthMove").value), 0, RsSavRec.Fields("MonthMove").value)
    TxtMonthNotMove.text = IIf(IsNull(RsSavRec.Fields("MonthNotMove").value), 0, RsSavRec.Fields("MonthNotMove").value)


    
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
        If RsSavRec.Fields("Mont_Day").value = 0 Then
         
            Mont_Day(0).value = True
            Else
                   Mont_Day(1).value = True
        End If
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
    
   If Not IsNull(RsSavRec.Fields("ProjAccount").value) Then
        If RsSavRec.Fields("ProjAccount").value = 1 Then
            CheckProjAccount.value = vbChecked
        Else
            CheckProjAccount.value = vbUnchecked
        End If
    Else
        CheckProjAccount.value = vbUnchecked
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
        If Not IsNull(RsSavRec.Fields("ItemStust").value) Then
        If RsSavRec.Fields("ItemStust").value = 2 Then
            ItemStust(2).value = True
        ElseIf RsSavRec.Fields("ItemStust").value = 1 Then
            ItemStust(1).value = True
        ElseIf RsSavRec.Fields("ItemStust").value = 0 Then
            ItemStust(0).value = True
        End If
    End If
    
'###################################### Account Part ########################################
    If Not IsNull(RsSavRec.Fields("AccOrTran").value) Then
        If RsSavRec.Fields("AccOrTran").value = 1 Then
                     
                        transactions.value = True
                        transactions_Click
                        
                      
        Else
        
        
            Accounts.value = True
            Accounts_Click
            
            
            
        End If
    Else
        transactions.value = True
        transactions_Click
    End If
    
    
'############################################################################################

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    
    FullGridData
ItemStust_Click (0)
ErrTrap:
End Sub
Private Sub btnSave_Click()
    
    On Error GoTo ErrTrap
    
    Dim total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
   
    If val(traTypeCombo.BoundText) = 0 And transactions.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·—Ã«¡ «œŒ«· ‰Ê⁄ «·Õ—ﬂ…", vbOKOnly + vbMsgBoxRight, App.Title
        Else
            MsgBox "Please choose the transaction type", vbOKOnly + vbMsgBoxRight, App.Title
        End If
        Exit Sub
    End If
    If val(traTypeCombo.BoundText) <> 29 And val(traTypeCombo.BoundText) <> 30 And val(traTypeCombo.BoundText) <> 42 Then
    If AccDibDC.BoundText = "" And AccCirDC.BoundText = "" And transactions.value = True And Accounts.value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·—Ã«¡ «œŒ«· «·Õ”«»", vbOKOnly + vbMsgBoxRight, App.Title
        Else
            MsgBox "Please choose an account", vbOKOnly + vbMsgBoxRight, App.Title
        End If
        Exit Sub
    ElseIf AccDibDC.BoundText = "" And AccCirDC.BoundText = "" And Accounts.value = True And Accounts.value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·—Ã«¡ «œŒ«· «·Õ”«»", vbOKOnly + vbMsgBoxRight, App.Title
        Else
            MsgBox "Please choose an account", vbOKOnly + vbMsgBoxRight, App.Title
        End If
        Exit Sub
    End If
     
    End If
    If Accounts.value = True And AccGrid.rows = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·—Ã«¡ «· √ﬂœ „‰ «÷«›… Õ”«» Ê«Õœ ⁄·Ï «·«ﬁ· ", vbOKOnly + vbMsgBoxRight, App.Title
        Else
            MsgBox "Please make sure to include at least one account", vbOKOnly + vbMsgBoxRight, App.Title
        End If
        Exit Sub
    ElseIf transactions.value = True And XPOptShowType(2).value = True Then
        If VSFlexGrid1.rows = 1 And FgItem.rows = 1 And C1Tab1.CurrTab = 0 And (val(traTypeCombo.BoundText) = 21 Or val(traTypeCombo.BoundText) = 22 Or val(traTypeCombo.BoundText) = 9 Or val(traTypeCombo.BoundText) = 5 And ((C1Tab1.CurrTab = 0 And XPOptShowType(2).value = True) Or (C1Tab1.CurrTab = 1 And XPOptShowType(2).value = True))) Then
            If VSFlexGrid1.rows = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·—Ã«¡ œŒ«· ’‰› Ê«Õœ ⁄·Ï «·√ﬁ·", vbOKOnly + vbMsgBoxRight, App.Title
                Else
                    MsgBox "Please make sure to include at least one item", vbOKOnly + vbMsgBoxRight, App.Title
                End If
                Exit Sub
            End If
        
            If FgItem.rows = 1 And (val(traTypeCombo.BoundText) = 21 Or val(traTypeCombo.BoundText) = 5 Or val(traTypeCombo.BoundText) = 9 Or val(traTypeCombo.BoundText) = 22) And (C1Tab1.CurrTab = 1 Or (C1Tab1.CurrTab = 0 And XPOptShowType(2).value = True)) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·—Ã«¡ œŒ«· ’‰› Ê«Õœ ⁄·Ï «·√ﬁ·", vbOKOnly + vbMsgBoxRight, App.Title
                Else
                    MsgBox "Please make sure to include at least one item", vbOKOnly + vbMsgBoxRight, App.Title
                End If
                Exit Sub
            End If
        End If
    End If
    
    
    Select Case Me.TxtModFlg.text
        Case "N"
            AddNewRec
        Case "E"
            FiLLRec
    End Select
    Exit Sub
ErrTrap:

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
Public Sub AddNewRec()
    
'    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    StrRecID = new_id("TblSettsReqLimK", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.text = IIf(StrRecID <> "", StrRecID, Null)
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Function GetQtyMonthaly(Optional ItemID As Double, Optional TransType As String) As Double
    
    Dim sql As String
    Dim CounQty As Double
    Dim i As Integer
    Dim SumQty As Double
    
    Dim Rs7 As ADODB.Recordset
    Set Rs7 = New ADODB.Recordset
    
    CounQty = 0
    SumQty = 0
    
    sql = " SELECT MONTH(dbo.Transactions.Transaction_Date) as Monthaly , dbo.Transaction_Details.Item_ID, SUM(dbo.Transaction_Details.Quantity) AS SumQty"
    sql = sql & " FROM dbo.Transactions LEFT OUTER JOIN"
    sql = sql & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    sql = sql & " WHERE (dbo.Transaction_Details.Item_ID = " & ItemID & ") AND (dbo.Transactions.Transaction_Type in (" & TransType & " ))"
    sql = sql & " GROUP BY MONTH(dbo.Transactions.Transaction_Date), dbo.Transaction_Details.Item_ID"
    
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
    
    Dim Rs7 As ADODB.Recordset
    Set Rs7 = New ADODB.Recordset
    
    CounQty = 0
    SumQty = 0

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
    FgItem.rows = 1
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 1
    ListGroupSelected.Clear
    
    sql = " SELECT dbo.TblSettsReqLimKDet.Qty, dbo.TblSettsReqLimKDet.Typ, dbo.TblSettsReqLimKDet.SetReqLID, dbo.TblSettsReqLimKDet.ID, "
    sql = sql & " dbo.TblSettsReqLimKDet.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblSettsReqLimKDet.ItemID, dbo.TblItems.ItemName,"
    sql = sql & " dbo.TblItems.ItemNamee, dbo.TblSettsReqLimKDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblSettsReqLimKDet.StoreID,"
    sql = sql & " dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.TblSettsReqLimKDet.Minimum, dbo.TblSettsReqLimKDet.Maximum,"
    sql = sql & " dbo.TblSettsReqLimKDet.SafetyRate, dbo.TblSettsReqLimKDet.Priod, dbo.TblSettsReqLimKDet.PriodType, dbo.TblSettsReqLimKDet.Mont_Day,"
    sql = sql & " dbo.TblSettsReqLimKDet.ConsuRate, dbo.TblItems.Fullcode, dbo.TblSettsReqLimKDet.UnitFactor, dbo.TblSettsReqLimKDet.ConsuRateLowQty,"
    sql = sql & " dbo.TblSettsReqLimKDet.MinLowQty , dbo.TblSettsReqLimKDet.MaxLowQty,dbo.TblSettsReqLimKDet.OperReQst ,TblSettsReqLimKDet.PercentD"
    sql = sql & " FROM dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
    sql = sql & " dbo.TblStore ON dbo.TblSettsReqLimKDet.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    sql = sql & " dbo.TblUnites ON dbo.TblSettsReqLimKDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    sql = sql & " dbo.TblItems ON dbo.TblSettsReqLimKDet.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    sql = sql & " dbo.Groups ON dbo.TblSettsReqLimKDet.GroupID = dbo.Groups.GroupID"
    sql = sql & " Where (dbo.TblSettsReqLimKDet.SetReqLID = " & val(TxtSerial1.text) & ") And (dbo.TblSettsReqLimKDet.typ = 0)"
    
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
    
    Dim i As Double
    With Me.FgItem
        For i = .FixedRows To Rs1.RecordCount
            .rows = .FixedRows + Rs1.RecordCount
            .TextMatrix(i, .ColIndex("Ser")) = i
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
            
            .TextMatrix(i, .ColIndex("MaxLowQty")) = IIf(IsNull(Rs1("MaxLowQty").value), 0, Rs1("MaxLowQty").value)
            .TextMatrix(i, .ColIndex("MinLowQty")) = IIf(IsNull(Rs1("MinLowQty").value), 0, Rs1("MinLowQty").value)
            .TextMatrix(i, .ColIndex("ConsuRateLowQty")) = IIf(IsNull(Rs1("ConsuRateLowQty").value), 0, Rs1("ConsuRateLowQty").value)
            .TextMatrix(i, .ColIndex("percent")) = IIf(IsNull(Rs1("percentD").value), 0, Rs1("percentD").value)
            
            Rs1.MoveNext
        Next i
    End With
    
'###################################################################################################################################################################################################

    sql = " SELECT dbo.TblSettsReqLimKDet.Qty, dbo.TblSettsReqLimKDet.Typ, dbo.TblSettsReqLimKDet.SetReqLID, dbo.TblSettsReqLimKDet.ID, "
    sql = sql & " dbo.TblSettsReqLimKDet.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblSettsReqLimKDet.ItemID, dbo.TblItems.ItemName,"
    sql = sql & " dbo.TblItems.ItemNamee, dbo.TblSettsReqLimKDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblSettsReqLimKDet.StoreID,"
    sql = sql & " dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.TblSettsReqLimKDet.Minimum, dbo.TblSettsReqLimKDet.Maximum,"
    sql = sql & " dbo.TblSettsReqLimKDet.SafetyRate, dbo.TblSettsReqLimKDet.Priod, dbo.TblSettsReqLimKDet.PriodType, dbo.TblSettsReqLimKDet.Mont_Day,"
    sql = sql & " dbo.TblSettsReqLimKDet.ConsuRate, dbo.TblItems.Fullcode, dbo.TblSettsReqLimKDet.UnitFactor, dbo.TblSettsReqLimKDet.ConsuRateLowQty,"
    sql = sql & " dbo.TblSettsReqLimKDet.MinLowQty , dbo.TblSettsReqLimKDet.MaxLowQty,dbo.TblSettsReqLimKDet.OperReQst ,TblSettsReqLimKDet.PercentD"
    sql = sql & " FROM dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
    sql = sql & " dbo.TblStore ON dbo.TblSettsReqLimKDet.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    sql = sql & " dbo.TblUnites ON dbo.TblSettsReqLimKDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    sql = sql & " dbo.TblItems ON dbo.TblSettsReqLimKDet.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    sql = sql & " dbo.Groups ON dbo.TblSettsReqLimKDet.GroupID = dbo.Groups.GroupID"
    sql = sql & " Where (dbo.TblSettsReqLimKDet.SetReqLID = " & val(TxtSerial1.text) & ") And (dbo.TblSettsReqLimKDet.typ = 5)"

    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
    
    With Me.VSFlexGrid1
        For i = .FixedRows To Rs1.RecordCount
            .rows = .FixedRows + Rs1.RecordCount
            .TextMatrix(i, .ColIndex("Ser")) = i
'            .TextMatrix(i, .ColIndex("OperReQst")) = IIf(IsNull(Rs1("OperReQst").value), 0, Rs1("OperReQst").value)
'            .TextMatrix(i, .ColIndex("ConsuRate")) = IIf(IsNull(Rs1("ConsuRate").value), 0, Rs1("ConsuRate").value)
'            .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(Rs1("Qty").value), 0, Rs1("Qty").value)
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
            
'            .TextMatrix(i, .ColIndex("Mont_Day")) = IIf(IsNull(Rs1("Mont_Day").value), "", Rs1("Mont_Day").value)
'            .TextMatrix(i, .ColIndex("PriodType")) = IIf(IsNull(Rs1("PriodType").value), "", Rs1("PriodType").value)
'            .TextMatrix(i, .ColIndex("Priod")) = IIf(IsNull(Rs1("Priod").value), 0, Rs1("Priod").value)
'            .TextMatrix(i, .ColIndex("SafetyRate")) = IIf(IsNull(Rs1("SafetyRate").value), 0, Rs1("SafetyRate").value)
            .TextMatrix(i, .ColIndex("MaxLowQty")) = IIf(IsNull(Rs1("MaxLowQty").value), 0, Rs1("MaxLowQty").value)
            .TextMatrix(i, .ColIndex("MinLowQty")) = IIf(IsNull(Rs1("MinLowQty").value), 0, Rs1("MinLowQty").value)
            .TextMatrix(i, .ColIndex("ConsuRateLowQty")) = IIf(IsNull(Rs1("ConsuRateLowQty").value), 0, Rs1("ConsuRateLowQty").value)
'            .TextMatrix(i, .ColIndex("Maximum")) = IIf(IsNull(Rs1("Maximum").value), 0, Rs1("Maximum").value)
'            .TextMatrix(i, .ColIndex("Minimum")) = IIf(IsNull(Rs1("Minimum").value), 0, Rs1("Minimum").value)
            .TextMatrix(i, .ColIndex("percent")) = IIf(IsNull(Rs1("percentD").value), 0, Rs1("percentD").value)
            
            Rs1.MoveNext
        Next i
    End With

'###############################################################################################################################################################
    
    sql = " SELECT ACCOUNTS.Account_Name, ACCOUNTS.Account_NameEng, TblBranchesData.branch_name, TblBranchesData.branch_namee, ACCOUNTS.Account_Serial, TblSettsReqLimKDet.ForcedFlg, TblSettsReqLimKDet.PercentD, "
    sql = sql & " TblSettsReqLimKDet.BranchID , TblSettsReqLimKDet.Account_Code "
    sql = sql & " FROM TblBranchesData RIGHT OUTER JOIN "
    sql = sql & " TblSettsReqLimKDet ON TblBranchesData.branch_id = TblSettsReqLimKDet.BranchID LEFT OUTER JOIN "
    sql = sql & " ACCOUNTS ON TblSettsReqLimKDet.Account_Code = ACCOUNTS.Account_Code FULL OUTER JOIN "
    sql = sql & " TblSettsReqLimK ON TblSettsReqLimKDet.SetReqLID = TblSettsReqLimK.ID "
    sql = sql & " Where (TblSettsReqLimKDet.SetReqLID = " & val(TxtSerial1.text) & ") And (TblSettsReqLimKDet.typ = 9) "
    
    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
    
    With Me.AccGrid
        For i = .FixedRows To Rs1.RecordCount
            .rows = .FixedRows + Rs1.RecordCount
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs1("Account_Serial").value), "", Rs1("Account_Serial").value)
            .TextMatrix(i, .ColIndex("AccCode")) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("AccName")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
            Else
                .TextMatrix(i, .ColIndex("AccName")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
            End If
            .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
            Else
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
            End If
            
            .TextMatrix(i, .ColIndex("AddedPer")) = IIf(IsNull(Rs1("PercentD").value), 0, Rs1("PercentD").value)
            
            If Not IsNull(Rs1("ForcedFlg").value) Then
                If Rs1("ForcedFlg").value = True Then
                    .cell(flexcpChecked, i, .ColIndex("ForcedFlg")) = flexChecked
                Else
                    .cell(flexcpChecked, i, .ColIndex("ForcedFlg")) = flexUnchecked
                End If
            Else
                .cell(flexcpChecked, i, .ColIndex("ForcedFlg")) = flexUnchecked
            End If
            Rs1.MoveNext
        Next i
    End With

'###############################################################################################################################################################
    FgStore.Clear flexClearScrollable, flexClearEverything
    FgStore.rows = 1
    
    sql = "SELECT     dbo.TblSettsReqLimKDet.Typ, dbo.TblSettsReqLimKDet.SetReqLID, dbo.TblSettsReqLimKDet.ID, dbo.TblSettsReqLimKDet.StoreID, "
    sql = sql & "                      dbo.TblStore.STORENAME , dbo.TblStore.StoreNamee, dbo.TblStore.code"
    sql = sql & "  FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
    sql = sql & "                       dbo.TblStore ON dbo.TblSettsReqLimKDet.StoreID = dbo.TblStore.StoreID"
    sql = sql & " Where (dbo.TblSettsReqLimKDet.SetReqLID = " & val(TxtSerial1.text) & ") And (dbo.TblSettsReqLimKDet.typ = 1)"
    
    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
     
    With Me.FgStore
        For i = .FixedRows To Rs1.RecordCount
            .rows = .FixedRows + Rs1.RecordCount
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
    
'#############################################################################################################################################################

    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    
    StrSQL = " SELECT     dbo.TblSettsReqLimKDet.Typ, dbo.TblSettsReqLimKDet.SetReqLID, dbo.TblSettsReqLimKDet.ID, dbo.TblSettsReqLimKDet.GroupID,"
    StrSQL = StrSQL & "                       dbo.Groups.GroupName , dbo.Groups.fullcode, dbo.Groups.GroupNamee"
    StrSQL = StrSQL & " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblSettsReqLimKDet.GroupID = dbo.Groups.GroupID"
    StrSQL = StrSQL & " Where (dbo.TblSettsReqLimKDet.SetReqLID = " & val(TxtSerial1.text) & ") And (dbo.TblSettsReqLimKDet.typ = 2)"
    
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
'#############################################################################################################################################################

    Set RsDetails = New ADODB.Recordset
    
    StrSQL = "SELECT TblBranchesData.branch_name, TblBranchesData.branch_namee, TblSettsReqLimKDet.BranchID "
    StrSQL = StrSQL & " FROM TblBranchesData RIGHT OUTER JOIN "
    StrSQL = StrSQL & " TblSettsReqLimKDet ON TblBranchesData.branch_id = TblSettsReqLimKDet.BranchID "
    StrSQL = StrSQL & " Where (dbo.TblSettsReqLimKDet.SetReqLID = " & val(TxtSerial1.text) & ") And (dbo.TblSettsReqLimKDet.typ = 8) "
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    SelectedBranchList.Clear
    For i = 0 To RsDetails.RecordCount - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            SelectedBranchList.AddItem IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
        Else
            SelectedBranchList.AddItem IIf(IsNull(RsDetails("branch_namee").value), "", RsDetails("branch_namee").value)
        End If
        SelectedBranchList.ItemData(i) = IIf(IsNull(RsDetails("BranchID").value), "", RsDetails("BranchID").value)
        RsDetails.MoveNext
    Next i
    
'#############################################################################################################################################################
Exit Sub
ErrTrap:
End Sub
Private Sub ISButton2_Click()
    If Me.TxtModFlg.text <> "R" Then
        If FgStore.rows = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «·„Œ“‰ «Ê·«"
            Else
                MsgBox "Please Select Store"
            End If
            FgStore.SetFocus
            Exit Sub
        End If
        
        If C1Tab1.CurrTab = 0 Then
          '  Retrivetitems
        Else
         '   Retrivetitems2
        End If
    End If
End Sub
Private Sub ISButton3_Click()
    If Me.TxtModFlg.text <> "R" Then
        If FgStore.rows = 1 Then
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
                MsgBox "Please select unit first"
            End If
            Exit Sub
        End If
        If C1Tab1.CurrTab = 0 Then
            Retrivetitems
        Else
            Retrivetitems2
        End If
    End If
End Sub
Private Sub ISButton5_Click()
    print_report
End Sub
Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtCode.text = "" Then
            Me.DcbItem.BoundText = ""
        Else
            Me.DcbItem.BoundText = GetItemID(Trim$(Me.TxtCode.text))
        End If
    End If
End Sub
Private Sub TxtSerial1_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
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
Private Sub BtnCancel_Click()
    Unload Me
End Sub
 Private Sub BtnUndo_Click()
 
 If Me.TxtModFlg.text = "E" Then
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
    FiLLTXT
    ElseIf Me.TxtModFlg.text = "N" Then
    Me.TxtModFlg.text = "R"
     BtnLast_Click
    End If
End Sub
Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim sql As String
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
    
    On Error GoTo ErrTrap
    
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    
    If X = vbNo Then Exit Sub
        If TxtSerial1.text = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
            Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
            End If
        Else
            StrSQL = "Delete From TblSettsReqLimKDet Where   SetReqLID=" & val(Me.TxtSerial1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
            RsSavRec.delete
            FgItem.Clear flexClearScrollable, flexClearEverything
            FgItem.rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1
            
            AccGrid.Clear flexClearScrollable, flexClearEverything
            AccGrid.rows = 1
            
            SelectedBranchList.Clear
            
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            
            If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
            Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
            End If
            
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
        End If
        Me.Refresh
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           Cn.Errors.Clear
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
Public Sub EditRec(StrTable As String, RecId As String)
    FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.text = "N" Then
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.text = "R" Then
        btnModify.Enabled = False
        btnDelete.Enabled = False
        
        If TxtSerial1.text <> "" Then
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
    ElseIf TxtModFlg.text = "E" Then
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub
Private Sub BtnFirst_Click()

    Dim Msg As String
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()

   On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()

    Dim Msg As String
    
    On Error GoTo ErrTrap
    
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
        FgStore.rows = FgStore.rows + 1
        Me.DCboUserName.BoundText = user_id
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry.." & CHR(13)
                Msg = Msg & " You can not edit this the record now" & CHR(13)
                Msg = Msg & "It was being edited by another user on the network"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    
    On Error GoTo ErrTrap
    
    clear_all Me
    TxtModFlg.text = "N"
    XPOptShowType_Click (0)
    Auto_Manula(1).value = True
    Auto_Manula_Click (1)
    ChAllStore.value = vbUnchecked
    Mont_Day(1).value = True
    Label6_Click
    FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.rows = 1
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 1
     FgStore.Clear flexClearScrollable, flexClearEverything
    FgStore.rows = 1
    transactions.value = True
    transactions_Click
    ListGroupSelected.Clear
    SelectedBranchList.Clear
     AccGrid.Clear flexClearScrollable, flexClearEverything
    AccGrid.rows = 1
    
    AddNewRecored
    Me.DCboUserName.BoundText = user_id
    
ErrTrap:
End Sub
Private Sub BtnNext_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()

    Dim Msg As String
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub ShowTip()

    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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
Private Sub ChangeLang()

    On Error GoTo ErrTrap
    CheckProjAccount.Caption = "Link Project Accounts"
    CheckProjAccount.RightToLeft = False
    Dim XPic As IPictureDisp
    Set XPic = btnFirst.ButtonImage
    Set btnFirst.ButtonImage = btnLast.ButtonImage
    Set btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = btnNext.ButtonImage
    Set btnNext.ButtonImage = XPic
    ForcedFlg.RightToLeft = False
    ForcedFlg.Caption = "Mandatory / Optional"
    lbl(28).Caption = "Fixed Assest"
    lbl(32).Caption = "Movable       Period"
    lbl(33).Caption = "Not Movable   Period"
    lbl(29).Caption = "Year"
    lbl(30).Caption = "Year"
    lbl(24).Caption = "Code"
    lbl(25).Caption = "Select Branch"
    ISButton4.Caption = "Same Copy"
    Label1(2).Caption = "VAT Settings"
    lbl(4).Caption = "No"
    lbl(1).Caption = "Date"
    lbl(2).Caption = "Remarks"
    lbl(5).Caption = "Auto"
    lbl(49).Caption = "Order Qty limit"
    ChAllStore.Caption = "All Store"
    lbl(17).Caption = "Store"
    lbl(51).Caption = "Item"
    lbl(0).Caption = "Unit"
    C1Tab1.Caption = "«VAT|Free Items|Accounts"
    Cmd(0).Caption = "Delete"
    Cmd(1).Caption = "Delete All"
    lbl(7).Caption = "Transaction Type"
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
    lbl(3).Caption = "Calculation Method"
    Mont_Day(0).Caption = "Monthly"
    Mont_Day(1).Caption = "Daily"
    Auto_Manula(0).Caption = "Manual"
    Auto_Manula(1).Caption = "Auto"
    lbl(9).Caption = "From"
    lbl(10).Caption = "To"
    ISButton2.Caption = "Add"
    BtonAdd.Caption = "Add"
    ISButton3.Caption = "Add"
    Cmd(13).Caption = "Delete"
    Cmd(15).Caption = "Delete All"
    Cmd(3).Caption = "Delete"
    Cmd(4).Caption = "Delete All"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.lbl(8).Caption = "by"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    lbl(6).Caption = "Period"
    
    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name "
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit"
       .TextMatrix(0, .ColIndex("percent")) = "Percent"
        .TextMatrix(0, .ColIndex("storename")) = "Store"
    End With
  
    With Me.FgItem
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name "
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit"
        .TextMatrix(0, .ColIndex("percent")) = "Percent"
        .TextMatrix(0, .ColIndex("storename")) = "Store"
    End With
    
    lbl(11).Caption = "Add percent Value"
    
    With Me.FgStore
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Code")) = "Code "
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name "
    End With
    
    lbl(12).Caption = "Debit account for add value"
    lbl(13).Caption = "Credit account for add value"
    lbl(14).Caption = "Percent"
    lbl(15).Caption = "Percent"
    Accounts.Caption = "Accounts"
    transactions.Caption = "Transactions"
    lbl(23).Caption = "Account Name"
    lbl(21).Caption = "VAT Percentage"
    ISButton6.Caption = "Add"
    DelRowAccGrid.Caption = "Delete Row"
    DelAllAccGrid.Caption = "Delete All"
    
    With AccGrid
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("FullCode")) = "Account Code"
        .TextMatrix(0, .ColIndex("AccName")) = "Account Name"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("AddedPer")) = "VAT Percentage"
        .TextMatrix(0, .ColIndex("ForcedFlg")) = "Mandatory/Optional"
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
    
    'MySQL = "SELECT dbo.TblSettsReqLimK.ID, dbo.TblSettsReqLimK.RecordDate, dbo.TblSettsReqLimK.AllStore, dbo.TblSettsReqLimK.Remarks, dbo.TblSettsReqLimK.SelectType, dbo.TblSettsReqLimK.Qty, dbo.TblSettsReqLimK.Mont_Day, "
    'MySQL = MySQL & " dbo.TblSettsReqLimK.Auto_Manula, dbo.TblSettsReqLimK.TransType, dbo.TblSettsReqLimKDet.Typ, dbo.TblSettsReqLimKDet.Qty AS DetQty, dbo.TblSettsReqLimKDet.ConsuRate, "
    'MySQL = MySQL & " dbo.TblSettsReqLimKDet.Mont_Day AS DetMont_Day, dbo.TblSettsReqLimKDet.PriodType AS DetPriodType, dbo.TblSettsReqLimKDet.Priod AS DetPriod, dbo.TblSettsReqLimKDet.SafetyRate AS DetSafetyRate, "
    'MySQL = MySQL & " dbo.TblSettsReqLimKDet.Maximum, dbo.TblSettsReqLimKDet.Minimum, dbo.TblSettsReqLimKDet.UnitFactor, dbo.TblSettsReqLimKDet.MaxLowQty, dbo.TblSettsReqLimKDet.MinLowQty, "
    'MySQL = MySQL & " dbo.TblSettsReqLimKDet.ConsuRateLowQty, dbo.TblSettsReqLimKDet.StoreID AS DetStoreID, dbo.TblSettsReqLimKDet.ItemID AS DetItemID, dbo.TblSettsReqLimKDet.UnitID AS DetUnitID, dbo.TblSettsReqLimKDet.GroupID, "
    'MySQL = MySQL & " dbo.TblSettsReqLimK.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblSettsReqLimK.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblSettsReqLimK.UnitID, "
    'MySQL = MySQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, TblStore_1.StoreName AS DetStoreName, TblStore_1.StoreNamee AS DetStoreNameE, TblStore_1.Code, TblItems_1.ItemName AS DetItemName, "
    'MySQL = MySQL & " TblItems_1.ItemNamee AS DetItemNameE, TblItems_1.Fullcode AS DetFullcode, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.Groups.Fullcode AS GropFullcode, TblUnites_1.UnitName AS DetUnitName, "
    'MySQL = MySQL & " TblUnites_1.UnitNamee AS DetUnitNameE, dbo.TblSettsReqLimK.PercentH, dbo.TblSettsReqLimKDet.PercentD, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_NameEng, dbo.ACCOUNTS.Account_Name AS AccNameDebit, "
    'MySQL = MySQL & " dbo.ACCOUNTS.Account_NameEng AS AccNameEDebit, dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.TblSettsReqLimK.RecordDateTo "
    'MySQL = MySQL & " FROM dbo.ACCOUNTS INNER JOIN "
    'MySQL = MySQL & " dbo.ACCOUNTS AS ACCOUNTS_1 INNER JOIN "
    'MySQL = MySQL & " dbo.TblSettsReqLimK ON ACCOUNTS_1.Account_Code = dbo.TblSettsReqLimK.AccDep ON dbo.ACCOUNTS.Account_Code = dbo.TblSettsReqLimK.AccCir INNER JOIN "
    'MySQL = MySQL & " dbo.TransactionTypes ON dbo.TblSettsReqLimK.TransType = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblUnites ON dbo.TblSettsReqLimK.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblItems ON dbo.TblSettsReqLimK.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblStore AS TblStore_1 RIGHT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblSettsReqLimKDet LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblUnites AS TblUnites_1 ON dbo.TblSettsReqLimKDet.UnitID = TblUnites_1.UnitID LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.Groups ON dbo.TblSettsReqLimKDet.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblItems AS TblItems_1 ON dbo.TblSettsReqLimKDet.ItemID = TblItems_1.ItemID ON TblStore_1.StoreID = dbo.TblSettsReqLimKDet.StoreID ON "
    'MySQL = MySQL & " dbo.TblSettsReqLimK.ID = dbo.TblSettsReqLimKDet.SetReqLID LEFT OUTER JOIN "
    'MySQL = MySQL & " dbo.TblStore ON dbo.TblSettsReqLimK.StoreID = dbo.TblStore.StoreID "

    'MySQL = "SELECT TblSettsReqLimK.ID, TblSettsReqLimK.RecordDate, TblSettsReqLimK.AllStore, TblSettsReqLimK.Remarks, TblSettsReqLimK.SelectType, TblSettsReqLimK.Qty, TblSettsReqLimK.Mont_Day, TblSettsReqLimK.Auto_Manula, "
    'MySQL = MySQL & " TblSettsReqLimK.TransType, TblSettsReqLimKDet.Typ, TblSettsReqLimKDet.Qty AS DetQty, TblSettsReqLimKDet.ConsuRate, TblSettsReqLimKDet.Mont_Day AS DetMont_Day, TblSettsReqLimKDet.PriodType AS DetPriodType, "
    'MySQL = MySQL & " TblSettsReqLimKDet.Priod AS DetPriod, TblSettsReqLimKDet.SafetyRate AS DetSafetyRate, TblSettsReqLimKDet.Maximum, TblSettsReqLimKDet.Minimum, TblSettsReqLimKDet.UnitFactor, TblSettsReqLimKDet.MaxLowQty, "
    'MySQL = MySQL & " TblSettsReqLimKDet.MinLowQty, TblSettsReqLimKDet.ConsuRateLowQty, TblSettsReqLimKDet.StoreID AS DetStoreID, TblSettsReqLimKDet.ItemID AS DetItemID, TblSettsReqLimKDet.UnitID AS DetUnitID, "
    'MySQL = MySQL & " TblSettsReqLimKDet.GroupID, TblSettsReqLimK.StoreID, TblStore.StoreName, TblStore.StoreNamee, TblSettsReqLimK.ItemID, TblItems.ItemName, TblItems.ItemNamee, TblItems.Fullcode, TblSettsReqLimK.UnitID, "
    'MySQL = MySQL & " TblUnites.UnitName, TblUnites.UnitNamee, TblStore_1.StoreName AS DetStoreName, TblStore_1.StoreNamee AS DetStoreNameE, TblStore_1.Code, TblItems_1.ItemName AS DetItemName, "
    'MySQL = MySQL & " TblItems_1.ItemNamee AS DetItemNameE, TblItems_1.Fullcode AS DetFullcode, Groups.GroupName, Groups.GroupNamee, Groups.Fullcode AS GropFullcode, TblUnites_1.UnitName AS DetUnitName, "
    'MySQL = MySQL & " TblUnites_1.UnitNamee AS DetUnitNameE, TblSettsReqLimK.PercentH, TblSettsReqLimKDet.PercentD, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_NameEng, ACCOUNTS.Account_Name AS AccNameDebit, "
    'MySQL = MySQL & " ACCOUNTS.Account_NameEng AS AccNameEDebit, TransactionTypes.TransactionTypeName, TransactionTypes.TransactionEnglishName, TblSettsReqLimK.RecordDateTo "
    'MySQL = MySQL & " FROM ACCOUNTS AS ACCOUNTS_1 RIGHT OUTER JOIN "
    'MySQL = MySQL & " ACCOUNTS RIGHT OUTER JOIN "
    'MySQL = MySQL & " TblSettsReqLimK ON ACCOUNTS.Account_Code = TblSettsReqLimK.AccCir ON ACCOUNTS_1.Account_Code = TblSettsReqLimK.AccDep LEFT OUTER JOIN "
    'MySQL = MySQL & " TransactionTypes ON TblSettsReqLimK.TransType = TransactionTypes.Transaction_Type LEFT OUTER JOIN "
    'MySQL = MySQL & " TblUnites ON TblSettsReqLimK.UnitID = TblUnites.UnitID LEFT OUTER JOIN "
    'MySQL = MySQL & " TblItems ON TblSettsReqLimK.ItemID = TblItems.ItemID LEFT OUTER JOIN "
    'MySQL = MySQL & " TblStore ON TblSettsReqLimK.StoreID = TblStore.StoreID FULL OUTER JOIN "
    'MySQL = MySQL & " TblStore AS TblStore_1 RIGHT OUTER JOIN "
    'MySQL = MySQL & " TblSettsReqLimKDet LEFT OUTER JOIN "
    'MySQL = MySQL & " TblUnites AS TblUnites_1 ON TblSettsReqLimKDet.UnitID = TblUnites_1.UnitID LEFT OUTER JOIN "
    'MySQL = MySQL & " Groups ON TblSettsReqLimKDet.GroupID = Groups.GroupID LEFT OUTER JOIN "
    'MySQL = MySQL & " TblItems AS TblItems_1 ON TblSettsReqLimKDet.ItemID = TblItems_1.ItemID ON TblStore_1.StoreID = TblSettsReqLimKDet.StoreID ON TblSettsReqLimK.ID = TblSettsReqLimKDet.SetReqLID "
    'MySQL = MySQL & " Where (dbo.TblSettsReqLimK.ID = " & val(TxtSerial1.Text) & ") "

    MySQL = "SELECT TblSettsReqLimK.ID, TblSettsReqLimK.RecordDate, TblSettsReqLimK.AllStore, TblSettsReqLimK.Remarks, TblSettsReqLimK.SelectType, TblSettsReqLimK.Qty, TblSettsReqLimK.Mont_Day, TblSettsReqLimK.Auto_Manula,"
    MySQL = MySQL & " TblSettsReqLimK.TransType, TblSettsReqLimKDet.Typ, TblSettsReqLimKDet.Qty AS DetQty, TblSettsReqLimKDet.ConsuRate, TblSettsReqLimKDet.Mont_Day AS DetMont_Day, TblSettsReqLimKDet.PriodType AS DetPriodType,"
    MySQL = MySQL & " TblSettsReqLimKDet.Priod AS DetPriod, TblSettsReqLimKDet.SafetyRate AS DetSafetyRate, TblSettsReqLimKDet.Maximum, TblSettsReqLimKDet.Minimum, TblSettsReqLimKDet.UnitFactor, TblSettsReqLimKDet.MaxLowQty,"
    MySQL = MySQL & " TblSettsReqLimKDet.MinLowQty, TblSettsReqLimKDet.ConsuRateLowQty, TblSettsReqLimKDet.StoreID AS DetStoreID, TblSettsReqLimKDet.ItemID AS DetItemID, TblSettsReqLimKDet.UnitID AS DetUnitID,"
    MySQL = MySQL & " TblSettsReqLimKDet.GroupID, TblSettsReqLimK.StoreID, TblStore.StoreName, TblStore.StoreNamee, TblSettsReqLimK.ItemID, TblItems.ItemName, TblItems.ItemNamee, TblItems.Fullcode, TblSettsReqLimK.UnitID,"
    MySQL = MySQL & " TblUnites.UnitName, TblUnites.UnitNamee, TblStore_1.StoreName AS DetStoreName, TblStore_1.StoreNamee AS DetStoreNameE, TblStore_1.Code, TblItems_1.ItemName AS DetItemName,"
    MySQL = MySQL & " TblItems_1.ItemNamee AS DetItemNameE, TblItems_1.Fullcode AS DetFullcode, Groups.GroupName, Groups.GroupNamee, Groups.Fullcode AS GropFullcode, TblUnites_1.UnitName AS DetUnitName,"
    MySQL = MySQL & " TblUnites_1.UnitNamee AS DetUnitNameE, TblSettsReqLimK.PercentH, TblSettsReqLimKDet.PercentD, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_NameEng, ACCOUNTS.Account_Name AS AccNameDebit,"
    MySQL = MySQL & " ACCOUNTS.Account_NameEng AS AccNameEDebit, VatTypes.VatTypeName, VatTypes.VatTypeNamee, TblSettsReqLimK.RecordDateTo"
    MySQL = MySQL & " FROM ACCOUNTS AS ACCOUNTS_1 RIGHT OUTER JOIN"
    MySQL = MySQL & " ACCOUNTS RIGHT OUTER JOIN"
    MySQL = MySQL & " TblSettsReqLimK ON ACCOUNTS.Account_Code = TblSettsReqLimK.AccCir ON ACCOUNTS_1.Account_Code = TblSettsReqLimK.AccDep LEFT OUTER JOIN"
    MySQL = MySQL & " VatTypes ON TblSettsReqLimK.TransType = VatTypes.ID LEFT OUTER JOIN"
    MySQL = MySQL & " TblUnites ON TblSettsReqLimK.UnitID = TblUnites.UnitID LEFT OUTER JOIN"
    MySQL = MySQL & " TblItems ON TblSettsReqLimK.ItemID = TblItems.ItemID LEFT OUTER JOIN"
    MySQL = MySQL & " TblStore ON TblSettsReqLimK.StoreID = TblStore.StoreID FULL OUTER JOIN"
    MySQL = MySQL & " TblStore AS TblStore_1 RIGHT OUTER JOIN"
    MySQL = MySQL & " TblSettsReqLimKDet LEFT OUTER JOIN"
    MySQL = MySQL & " TblUnites AS TblUnites_1 ON TblSettsReqLimKDet.UnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    MySQL = MySQL & " Groups ON TblSettsReqLimKDet.GroupID = Groups.GroupID LEFT OUTER JOIN"
    MySQL = MySQL & " TblItems AS TblItems_1 ON TblSettsReqLimKDet.ItemID = TblItems_1.ItemID ON TblStore_1.StoreID = TblSettsReqLimKDet.StoreID ON TblSettsReqLimK.ID = TblSettsReqLimKDet.SetReqLID"
    MySQL = MySQL & " Where (dbo.TblSettsReqLimK.ID = " & val(TxtSerial1.text) & ") "

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAt\" & "RepAddValueSett.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAt\" & "RepAddValueSettE.rpt"
    End If
    If Dir(StrFileName) = "" Then
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
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
Private Sub AddNewRecored()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
       
    On Error GoTo ErrTrap

    My_SQL = "TblSettsReqLimK"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
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
    FgStore.rows = 1
End Sub
Private Sub RemoveGridAllRow()
    FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.rows = 1
End Sub
Private Sub RemoveGridAllRow5()
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 1
End Sub
Private Sub RemoveGridRow5()
    With Me.VSFlexGrid1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub RemoveGridRow()
    With Me.FgItem
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub RemoveGridStoreRow()
    With Me.FgStore
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub XPOptShowType_Click(Index As Integer)
    DcbUnitDit.BoundText = 0
    DcbItem.BoundText = 0
    Txtqty.text = ""
    TxtCode.text = ""
    If XPOptShowType(1).value = True Or XPOptShowType(0).value = True Then
        C1Elastic3.Enabled = False
    Else
        C1Elastic3.Enabled = True
    End If
End Sub
'########################################################### Accounts Part ##########################################################
Function FillBranchList()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    
    sql = " SELECT * from  TblBranchesData"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    BranchList.Clear
    SelectedBranchList.Clear
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                BranchList.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                BranchList.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            End If
            BranchList.ItemData(BranchList.NewIndex) = rs("branch_id").value
            rs.MoveNext
        Next i
    End If
    rs.Close
End Function
Private Sub Label10_Click()
    Dim sql As String
    Dim Rs1  As ADODB.Recordset
    Dim i, k As Integer
    If BranchList.ListIndex > -1 Then
        SelectedBranchList.AddItem BranchList.List(BranchList.ListIndex)
        SelectedBranchList.ItemData(SelectedBranchList.NewIndex) = BranchList.ItemData(BranchList.ListIndex)
    End If
End Sub
Private Sub Label9_Click()
    Dim i As Integer
    SelectedBranchList.Clear
    For i = 0 To BranchList.ListCount - 1
        SelectedBranchList.AddItem BranchList.List(i)
        SelectedBranchList.ItemData(i) = BranchList.ItemData(i)
    Next i
End Sub
Private Sub Label3_Click()
    If SelectedBranchList.ListIndex > -1 Then
        SelectedBranchList.RemoveItem SelectedBranchList.ListIndex
    End If
End Sub
Private Sub Label4_Click()
    SelectedBranchList.Clear
End Sub
Private Sub ISButton6_Click()
If Me.TxtModFlg.text <> "R" Then
        If SelectedBranchList.ListCount < 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— ›—⁄ Ê«Õœ ⁄·Ï «·√ﬁ· "
            Else
                MsgBox "Please Select Branch"
            End If
            Exit Sub
        End If
        If AccountsDC.BoundText = "" And chkAllExp.value = xtpUnchecked Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "—ÃÏ «Œ Ì«— Õ”«»"
            Else
                MsgBox "Please select Account first"
            End If
            Exit Sub
        End If
        If AccPerTxt.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «œŒ«· ‰”»… «·ﬁÌ„… «·„÷«›…"
            Else
                MsgBox "Please Enter VAT First"
            End If
            Exit Sub
        End If
        If C1Tab1.CurrTab = 2 Then
            FillAccGrid
        End If
    End If
End Sub
Sub FillAccGrid()
    Dim i As Long
    Dim j As Long
    Dim Rs1 As ADODB.Recordset
    Dim sql As String
    Dim s As String
    For i = 0 To SelectedBranchList.ListCount - 1
        Set Rs1 = New ADODB.Recordset
        'sql = "select * from ACCOUNTS where ACCOUNTS.Account_Code = '" & AccountsDC.BoundText & "'"
        
    If chkAllExp.value = xtpChecked And Trim(AccountsDC.text) = "" Then
        s = " SELECT * FROM ACCOUNTS AS a WHERE a.Account_Code IN (SELECT et.Account_Code FROM ExpensesType AS et)"
    Else
            s = " SELECT a.*"
        s = s & "          FROM   ACCOUNTS             AS a"
        s = s & "                 INNER JOIN ACCOUNTS  AS a2"
        s = s & "                      ON     a.last_account = 1 and a.Parent_Account_Code = a2.Account_Code"
        s = s & "          WHERE  a.Account_Code IN (SELECT Code"
        s = s & "                                    FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(AccountsDC.BoundText) & "', '" & Trim(AccountsDC.BoundText) & "', 1))"
        s = s & " OR (a.Account_Code = '" & Trim(AccountsDC.BoundText) & "' AND a.last_account = 1)"
        s = s & " AND a.last_account = 1"
        s = s & "          Order By"
        s = s & "                 a.Parent_Account_Code,"
        s = s & "                 a.last_account"
         
   End If
 
        
        
        Rs1.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not Rs1.EOF
            With AccGrid
           
                
                .rows = .rows + 1
                j = .rows - 1
                
                .TextMatrix(j, .ColIndex("Serial")) = j
                .TextMatrix(j, .ColIndex("FullCode")) = Rs1("Account_Serial").value
                .TextMatrix(j, .ColIndex("AccCode")) = Rs1("Account_Code").value
                .TextMatrix(j, .ColIndex("AccName")) = Rs1("Account_Name").value
                .TextMatrix(j, .ColIndex("BranchID")) = SelectedBranchList.ItemData(i)
                .TextMatrix(j, .ColIndex("Branch")) = SelectedBranchList.List(i)
                .TextMatrix(j, .ColIndex("AddedPer")) = AccPerTxt.text
                If ForcedFlg = vbChecked Then
                    .cell(flexcpChecked, j, .ColIndex("ForcedFlg")) = flexChecked
                Else
                    .cell(flexcpChecked, j, .ColIndex("ForcedFlg")) = flexUnchecked
                End If
                Rs1.MoveNext
            
            End With
            
        
        Loop
    Next i
End Sub
Private Sub DelRowAccGrid_Click()
    With Me.AccGrid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub DelAllAccGrid_Click()
    AccGrid.Clear flexClearScrollable, flexClearEverything
    AccGrid.rows = 1
End Sub
Public Function GetAccountCode(StrAccSerial As String) As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    If Trim(StrAccSerial) <> "" Then
        If Trim(StrAccSerial) = "" Then Exit Function
        StrSQL = "Select Account_Code From ACCOUNTS Where Account_Serial ='" & Trim(StrAccSerial) & "'"
        
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetAccountCode = rs("Account_Code").value
        Else
        End If

        rs.Close
        Set rs = Nothing
    End If
End Function
