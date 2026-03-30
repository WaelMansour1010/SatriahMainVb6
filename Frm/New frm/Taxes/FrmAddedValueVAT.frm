VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmAddedValueVAT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   Caption         =   " ”ÃÌ· «·Õ—ﬂ« "
   ClientHeight    =   9000
   ClientLeft      =   3780
   ClientTop       =   4740
   ClientWidth     =   16155
   Icon            =   "FrmAddedValueVAT.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   16560
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1560
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   16560
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1080
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmAddedValueVAT.frx":000C
      Left            =   16440
      List            =   "FrmAddedValueVAT.frx":001C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   3000
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   16560
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Text            =   "modflag"
      Top             =   4080
      Visible         =   0   'False
      Width           =   465
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   16800
      TabIndex        =   41
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   840
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
      Left            =   16440
      TabIndex        =   42
      Top             =   2160
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
      Left            =   16560
      Top             =   3600
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
            Picture         =   "FrmAddedValueVAT.frx":0035
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":03CF
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":0769
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":0B03
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":0E9D
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":1237
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":15D1
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddedValueVAT.frx":1B6B
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   16560
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   4920
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
      ButtonImage     =   "FrmAddedValueVAT.frx":1F05
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   17880
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   0
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
      ButtonImage     =   "FrmAddedValueVAT.frx":8767
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic7 
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16155
      _cx             =   28496
      _cy             =   15875
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
      Caption         =   " "
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
         Height          =   735
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   -15
         Width           =   16290
         Begin VB.TextBox TransTypeTblID 
            Height          =   285
            Left            =   3000
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   3
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
            ButtonImage     =   "FrmAddedValueVAT.frx":8B01
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   4
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
            ButtonImage     =   "FrmAddedValueVAT.frx":8E9B
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   5
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
            ButtonImage     =   "FrmAddedValueVAT.frx":9235
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
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
            ButtonImage     =   "FrmAddedValueVAT.frx":95CF
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   5400
            Picture         =   "FrmAddedValueVAT.frx":9969
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ”ÃÌ· «·Õ—ﬂ« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   7320
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1035
         Left            =   -90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   750
         Width           =   16350
         _cx             =   28840
         _cy             =   1826
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
            Height          =   480
            Left            =   8040
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   570
            Width           =   3045
            _cx             =   5371
            _cy             =   847
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
            Begin XtremeSuiteControls.RadioButton ServiceRd 
               Height          =   285
               Left            =   1440
               TabIndex        =   52
               Top             =   120
               Width           =   1380
               _Version        =   786432
               _ExtentX        =   2434
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Œ‹‹‹‹œ„‹‹‹‹‹‹‹‹‹‹…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton ProuductRd 
               Height          =   285
               Left            =   120
               TabIndex        =   53
               Top             =   120
               Width           =   1260
               _Version        =   786432
               _ExtentX        =   2222
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "”‹‹‹‹‹·‹‹‹‹‹⁄‹‹‹‹‹…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1755
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   300
            Left            =   9150
            TabIndex        =   10
            Top             =   270
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            _Version        =   393216
            Format          =   236191745
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmAddedValueVAT.frx":D5D1
            Height          =   315
            Left            =   645
            TabIndex        =   11
            Top             =   240
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   480
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   570
            Width           =   7965
            _cx             =   14049
            _cy             =   847
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
            Begin VB.TextBox txtFile 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -1200
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   0
               Visible         =   0   'False
               Width           =   1365
            End
            Begin XtremeSuiteControls.RadioButton RdTyp 
               Height          =   315
               Index           =   0
               Left            =   6720
               TabIndex        =   55
               Top             =   120
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "ÌœÊÌ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTyp 
               Height          =   315
               Index           =   1
               Left            =   5130
               TabIndex        =   56
               Top             =   120
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "„‰ „·›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   315
               Left            =   720
               TabIndex        =   57
               Top             =   120
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               Caption         =   "«” Ì—«œ «·„·›"
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
               ButtonImage     =   "FrmAddedValueVAT.frx":D5E6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton4 
               Height          =   315
               Left            =   2970
               TabIndex        =   58
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   120
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               Caption         =   "Õœœ «·„”«—"
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
               ButtonImage     =   "FrmAddedValueVAT.frx":13E48
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   120
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin MSDataListLib.DataCombo VATTypeTransDC 
            Bindings        =   "FrmAddedValueVAT.frx":1A6AA
            Height          =   315
            Left            =   11280
            TabIndex        =   62
            Top             =   690
            Width           =   3210
            _ExtentX        =   5662
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
         Begin VB.ComboBox TransTypeCBox 
            Height          =   315
            ItemData        =   "FrmAddedValueVAT.frx":1A6BF
            Left            =   11310
            List            =   "FrmAddedValueVAT.frx":1A6C1
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   690
            Width           =   3150
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
            Height          =   240
            Index           =   5
            Left            =   14895
            TabIndex        =   24
            Top             =   690
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   270
            Index           =   4
            Left            =   14760
            TabIndex        =   23
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   0
            Left            =   8265
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   270
            Index           =   2
            Left            =   11280
            TabIndex        =   12
            Top             =   240
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid 
         Height          =   2955
         Left            =   0
         TabIndex        =   1
         Top             =   4185
         Width           =   16215
         _cx             =   28601
         _cy             =   5212
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   4
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   28
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmAddedValueVAT.frx":1A6C3
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2400
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1740
         Width           =   16215
         _cx             =   28601
         _cy             =   4233
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
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   13260
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   1440
            Width           =   1230
         End
         Begin VB.TextBox TxtVATNO 
            Height          =   345
            Left            =   5400
            TabIndex        =   49
            Top             =   675
            Width           =   3675
         End
         Begin VB.TextBox VATPerTxt 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   5865
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Text            =   "5"
            Top             =   1050
            Width           =   3210
         End
         Begin VB.TextBox NotesTxt 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   5400
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1785
            Width           =   9090
         End
         Begin VB.TextBox VatValueTxt 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1050
            Width           =   3675
         End
         Begin VB.TextBox ValueTxt 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Text            =   "0"
            Top             =   1050
            Width           =   3690
         End
         Begin VB.TextBox DocNo 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   255
            Width           =   3690
         End
         Begin MSComCtl2.DTPicker DocDate 
            Height          =   345
            Left            =   5400
            TabIndex        =   18
            Top             =   255
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   609
            _Version        =   393216
            Format          =   236191745
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton AddToGrid 
            Height          =   465
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   1785
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   820
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
            ButtonImage     =   "FrmAddedValueVAT.frx":1AAED
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo CusDC 
            Bindings        =   "FrmAddedValueVAT.frx":2134F
            Height          =   315
            Left            =   10800
            TabIndex        =   31
            Top             =   675
            Width           =   3690
            _ExtentX        =   6509
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
         Begin MSDataListLib.DataCombo BranchDCDet 
            Bindings        =   "FrmAddedValueVAT.frx":21364
            Height          =   315
            Left            =   270
            TabIndex        =   32
            Top             =   255
            Width           =   3675
            _ExtentX        =   6482
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
         Begin MSDataListLib.DataCombo DcbStore 
            Bindings        =   "FrmAddedValueVAT.frx":21379
            Height          =   315
            Left            =   270
            TabIndex        =   60
            Top             =   660
            Width           =   3675
            _ExtentX        =   6482
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
         Begin MSDataListLib.DataCombo DcbPayment 
            Bindings        =   "FrmAddedValueVAT.frx":2138E
            Height          =   315
            Left            =   10800
            TabIndex        =   91
            Top             =   1440
            Width           =   1530
            _ExtentX        =   2699
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
         Begin MSDataListLib.DataCombo DcbAccount 
            Bindings        =   "FrmAddedValueVAT.frx":213A3
            Height          =   315
            Left            =   5400
            TabIndex        =   93
            Top             =   1440
            Width           =   3675
            _ExtentX        =   6482
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
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   270
            TabIndex        =   95
            Top             =   1455
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰œÊﬁ"
            Height          =   255
            Index           =   24
            Left            =   3945
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—ﬂ“ «· ﬂ·›…"
            Height          =   360
            Index           =   23
            Left            =   9285
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   1440
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   22
            Left            =   12165
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·”œ«œ"
            Height          =   255
            Index           =   12
            Left            =   14685
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰"
            Height          =   300
            Index           =   21
            Left            =   4095
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· ”ÃÌ·"
            Height          =   240
            Index           =   15
            Left            =   9225
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   675
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Index           =   11
            Left            =   14685
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1905
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   360
            Index           =   10
            Left            =   4125
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   255
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÃÂ…"
            Height          =   285
            Index           =   9
            Left            =   14685
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   675
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·‹ VAT"
            Height          =   255
            Index           =   8
            Left            =   3945
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·‹ VAT"
            Height          =   255
            Index           =   7
            Left            =   9285
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   255
            Index           =   6
            Left            =   14685
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1080
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„” ‰œ"
            Height          =   480
            Index           =   3
            Left            =   9255
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   255
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„” ‰œ"
            Height          =   480
            Index           =   1
            Left            =   14700
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   255
            Width           =   1380
         End
      End
      Begin ImpulseButton.ISButton DelRow 
         Height          =   360
         Left            =   13455
         TabIndex        =   20
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   7260
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·”ÿ— «·Õ«·Ì"
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
         ButtonImage     =   "FrmAddedValueVAT.frx":213B8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton DelAll 
         Height          =   360
         Left            =   11805
         TabIndex        =   21
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   7260
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·ﬂ· "
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
         ButtonImage     =   "FrmAddedValueVAT.frx":27C1A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   630
         Left            =   0
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   8370
         Width           =   16155
         _cx             =   28496
         _cy             =   1111
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
            Left            =   9585
            TabIndex        =   64
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":2E47C
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   285
            Left            =   8610
            TabIndex        =   65
            Top             =   90
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":34CDE
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   285
            Left            =   7635
            TabIndex        =   66
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":3B540
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   285
            Left            =   6615
            TabIndex        =   67
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":41DA2
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   285
            Left            =   5610
            TabIndex        =   68
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":48604
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   285
            Left            =   2355
            TabIndex        =   69
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":4EE66
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   285
            Left            =   4620
            TabIndex        =   70
            Top             =   90
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":78A88
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   285
            Left            =   3555
            TabIndex        =   71
            Top             =   90
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
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
            ButtonImage     =   "FrmAddedValueVAT.frx":7F2EA
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   11520
            TabIndex        =   104
            Top             =   120
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ…"
            Height          =   345
            Index           =   13
            Left            =   15075
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   120
            Width           =   990
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   1095
         Left            =   0
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   8400
         Visible         =   0   'False
         Width           =   16215
         _cx             =   28601
         _cy             =   1931
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
         Begin VB.TextBox TxtCheckNo 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   10155
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   150
            Width           =   1485
         End
         Begin VB.TextBox TxtTransferNO 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   10170
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   630
            Width           =   1440
         End
         Begin MSComCtl2.DTPicker TransferDate 
            Height          =   405
            Left            =   8175
            TabIndex        =   75
            Top             =   510
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   714
            _Version        =   393216
            Format          =   244973569
            CurrentDate     =   38784
         End
         Begin XtremeSuiteControls.RadioButton RDPatmentType 
            Height          =   390
            Index           =   0
            Left            =   12900
            TabIndex        =   76
            Top             =   150
            Width           =   1125
            _Version        =   786432
            _ExtentX        =   1984
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "‘Ì‹‹ﬂ"
            ForeColor       =   12582912
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RDPatmentType 
            Height          =   390
            Index           =   1
            Left            =   12900
            TabIndex        =   77
            Top             =   510
            Width           =   1125
            _Version        =   786432
            _ExtentX        =   1984
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Õ‹Ê«·…"
            ForeColor       =   12582912
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker CheckDate 
            Height          =   405
            Left            =   8160
            TabIndex        =   78
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   714
            _Version        =   393216
            Format          =   244973569
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·”œ«œ"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   14985
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   120
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   375
            Index           =   17
            Left            =   11595
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   150
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Index           =   18
            Left            =   9585
            TabIndex        =   81
            Top             =   150
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÕÊ«·…"
            Height          =   375
            Index           =   19
            Left            =   11655
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Index           =   20
            Left            =   9645
            TabIndex        =   79
            Top             =   510
            Width           =   465
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic9 
         Height          =   675
         Left            =   0
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   7680
         Width           =   5535
         _cx             =   9763
         _cy             =   1191
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
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   225
            Width           =   900
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   3405
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   300
            Index           =   1
            Left            =   1530
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   225
            Width           =   1290
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   300
            Index           =   0
            Left            =   4035
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   225
            Width           =   1320
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   675
         Index           =   11
         Left            =   5640
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   7680
         Width           =   10455
         _cx             =   18441
         _cy             =   1191
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
         Begin VB.CommandButton Command2 
            Caption         =   "Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
            Height          =   465
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   135
            Width           =   1845
         End
         Begin VB.CommandButton Command5 
            Caption         =   "≈‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
            Height          =   465
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   135
            Width           =   1845
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8055
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   105
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   465
            Left            =   2070
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   135
            Width           =   3375
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   465
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   135
            Width           =   1845
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   390
            Index           =   35
            Left            =   3135
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   255
            Width           =   3165
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   240
         Index           =   26
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«Ã„«·Ì"
         Height          =   240
         Index           =   25
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   7200
         Width           =   1455
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
      Left            =   16440
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmAddedValueVAT"
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
Public LonRow As Double
Public LngCol As Double
Dim dstore As Integer
Dim dBox As Integer
Dim usertype As Integer
Dim EmpID As Integer
Dim userbranchid As Integer
Public CtranIndex As Integer

Private Sub AddToGrid_Click()
    
    If val(BranchDCDet.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ»  ÕœÌœ «·›—⁄"
       Else
            MsgBox "Branch is a must"
       End If
       Exit Sub
    End If
    If DocNo.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·„” ‰œ"
       Else
            MsgBox "document number is a must"
       End If
       Exit Sub
    End If
 If ProuductRd.value = False And ServiceRd.value = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ  ÕœÌœ Œœ„… /”·⁄…"
    Else
      MsgBox "Please select Service /Commodity"
    End If
    Exit Sub
End If

If TxtVATNO.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ» «œŒ«·    ”ÃÌ· «·›« "
       Else
            MsgBox "VAT  number is a must"
       End If
       Exit Sub
    End If
    

   ' CusDC_Click 0
   ' BranchDCDet_Click 0
    
    FillGrid2
    
End Sub
Function AddCus(Optional Customer As String, Optional ByRef VATNO As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim StrRecID As String
    Dim sql As String
    Dim SqlQu As String
    StrRecID = new_id("TblCustemers", "CusID", "")
    sql = "Select * from TblCustemers where 1 = -1"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    RsTemp.AddNew
    RsTemp("CusID").value = IIf(StrRecID <> "", StrRecID, Null)
    RsTemp("Type") = 1
    RsTemp("CusName").value = Customer
    RsTemp("CusNamee").value = Customer
    RsTemp("VATNO").value = VATNO
    RsTemp.update
    AddCus = StrRecID
End Function

Private Sub Command2_Click()
If Me.TxtModFlg.text = "R" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update TblVATSettings set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1.text) & " "
        RsSavRec.Requery
         FindRec val(TxtSerial1.text)
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If
End Sub

Private Sub Command5_Click()
If TxtNoteSerial.text = "" Then
createVoucher
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
        Else
            MsgBox "Done"
        End If
End If
End Sub
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "  ”ÃÌ· Õ—ﬂ«  «·›«   " & TxtSerial1.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblVATSettings"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1)
Notevalue = 0
notytype = 9086
Notevalue = val(lbl(26).Caption)
BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
        
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial

CREATE_VOUCHER_GE val(TXTNoteID.text), BranchID, user_id, NoteDate
'rs.Resync adAffectCurrent
 

     End If

End Function
Function CheckAccount(Optional ByRef str As String) As Boolean
Dim i As Integer
Dim StrAccountCodeCridet As String
   With VSFlexGrid
   CheckAccount = False
   For i = 1 To .rows - 1
   If val(.TextMatrix(i, .ColIndex("Value"))) Then
        If val(.TextMatrix(i, .ColIndex("BranchID"))) = 0 Then
          If SystemOptions.UserInterface = ArabicInterface Then
          str = "    «·›—⁄ €Ì— „ÊÃÊœ ›Ì «·”ÿ—    " & i
          Else
          str = "Please Select Branch in Line" & i
          End If
          CheckAccount = True
          Exit Function
        End If
        If val(.TextMatrix(i, .ColIndex("StoreID"))) = 0 Then
          If SystemOptions.UserInterface = ArabicInterface Then
          str = "    «·„Œ“‰  €Ì— „ÊÃÊœ ›Ì «·”ÿ—    " & i
          Else
          str = "Please Select Store in Line" & i
          End If
          CheckAccount = True
          Exit Function
        End If
        
          If val(.TextMatrix(i, .ColIndex("PayedType"))) = 1 Then
          If val(.TextMatrix(i, .ColIndex("BankId"))) = 0 And val(.TextMatrix(i, .ColIndex("BoxID"))) = 0 Then
          If SystemOptions.UserInterface = ArabicInterface Then
          str = "Ì—ÃÏ «Œ Ì«— «·’‰œÊﬁ ›Ì «·”ÿ— " & i
          Else
          str = "Please select Box in Line" & i
          End If
          CheckAccount = True
          Exit Function
          End If
          Else
          If .TextMatrix(i, .ColIndex("Account_Code")) = "" Then
          If SystemOptions.UserInterface = ArabicInterface Then
          str = "Ì—ÃÏ «Œ Ì«— «·Õ”«» ›Ì «·”ÿ— " & i
          Else
          str = "Please Select Account in Line" & i
          End If
          CheckAccount = True
          Exit Function
           End If
          End If
        If val(.TextMatrix(i, .ColIndex("VATValue"))) > 0 Then
             GetValueAddedAccount XPDtbTrans.value, , StrAccountCodeCridet, 1, 21
              If StrAccountCodeCridet = "" Then
          If SystemOptions.UserInterface = ArabicInterface Then
          str = "Ì—ÃÏ «Œ Ì«— Õ”«» «·ﬁÌ„… «·„÷«›…    "
          Else
          str = "Please Select Account of VAT"
          End If
          CheckAccount = True
          Exit Function
           End If
        End If
       End If
    Next i
   End With
End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "       ”ÃÌ· Õ—ﬂ«  «·›« " & TxtSerial1.text
    notes_id = general_noteid
    my_branch = val(dcBranch.BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Double
    Dim Msg2 As String
   If lbl(26).Caption > 0 Then
    line_no = 1
   With VSFlexGrid
   For i = 1 To .rows - 1
   Notevalue = val(.TextMatrix(i, .ColIndex("Value")))
    If Notevalue > 0 Then
          If val(.TextMatrix(i, .ColIndex("PayedType"))) = 1 Then
          If val(.TextMatrix(i, .ColIndex("BankId"))) <> 0 Then
           StrAccountCodeDebt = GetMyAccountCode("BanksData", "BankiD", val(.TextMatrix(i, .ColIndex("BankId"))))
            Msg2 = "Õ”«» «·»‰ﬂ"
           Else
           StrAccountCodeDebt = GetMyAccountCode("TblBoxesData", "BoxID", val(.TextMatrix(i, .ColIndex("BoxID"))))
            Msg2 = "Õ”«» «·Œ“‰…"
           End If
           Else
           StrAccountCodeDebt = .TextMatrix(i, .ColIndex("Account_Code"))
           End If
                             
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue + val(.TextMatrix(i, .ColIndex("VATValue"))), 0, Msg & "    Õ”«»  «·Œ“‰…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                            StrAccountCodeCridet = get_account_code_branch(2, val(.TextMatrix(i, .ColIndex("BranchID"))))
                            line_no = line_no + 1
                                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«»  «·„Œ“Ê‰  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
                        If val(.TextMatrix(i, .ColIndex("VATValue"))) > 0 Then
                           GetValueAddedAccount XPDtbTrans.value, , StrAccountCodeCridet, 1, 21
                           If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(.TextMatrix(i, .ColIndex("VATValue"))), 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ··„»Ì⁄«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                           End If
                           line_no = line_no + 1
                       End If
            End If
    
      Next i
      End With
      End If

    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
  End Function

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub ISButton3_Click()

    On Error Resume Next

    Dim astrSplit2tems2() As String
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Long
    Dim VATNO As String
    Dim DocNo As String
    Dim DocDate As String
    Dim Branch As String
    Dim BranchID As Double
    Dim CusID As Double
    Dim cus As String
    Dim value As String
    Dim VATPer As String
    Dim notes As String
    Dim TypeService As String
    Dim store As String
    Dim StoreID As Double
    Dim PayedTyp As Integer
    Dim Msg As String
    Dim PaymentNam As String
    Dim Account_Nam As String
    Dim BoxNam As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    If Me.TxtModFlg.text <> "R" Then
        Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
        VSFlexGrid.rows = 1
        If txtFile.text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
        Set ExcelObj = CreateObject("Excel.Application")
        Set ExcelSheet = CreateObject("Excel.Sheet")
        ExcelObj.Workbooks.Open txtFile.text
        DoEvents
        Set ExcelBook = ExcelObj.Workbooks(1)
        Set ExcelSheet = ExcelBook.Worksheets(1)
        
        With ExcelSheet
            i = 2
            Do Until .cells(i, 1) & "" = ""
            DocNo = .cells(i, 1)
            DocDate = .cells(i, 2)
            Branch = .cells(i, 3)
            
            If Branch = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«ÌÊÃœ ›—⁄ ›Ì «·”ÿ— " & i & CHR(13)
                Msg = Msg & "«·—Ã«¡ «œŒ«· «·›—⁄ Ê«⁄«œ… « Ì—«œ «·„·› "
            Else
                Msg = "The Row number " & i & "doesn't specify a branch " & CHR(13)
                Msg = Msg & "Please make sure all recoreds specify a branch "
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            Exit Sub
            End If
            store = .cells(i, 4)
            cus = .cells(i, 5)
            VATNO = .cells(i, 6)
            TypeService = .cells(i, 7)
            value = .cells(i, 8)
            VATPer = .cells(i, 9)
            PayedTyp = .cells(i, 10)
            Account_Nam = .cells(i, 11)
            PaymentNam = .cells(i, 12)
            BoxNam = .cells(i, 13)
            notes = .cells(i, 14)
            With VSFlexGrid
                .rows = .rows + 1
                If TypeService <> "" And TypeService <> "0" Then
                    .TextMatrix(i - 1, .ColIndex("TypeService")) = 2
                Else
                    .TextMatrix(i - 1, .ColIndex("TypeService")) = 1
                End If
                .TextMatrix(i - 1, .ColIndex("DocNo")) = DocNo
                .TextMatrix(i - 1, .ColIndex("DocDate")) = DocDate
                .TextMatrix(i - 1, .ColIndex("Branch")) = Branch
                BranchID = ChcekBranch(Branch)
                .TextMatrix(i - 1, .ColIndex("BranchID")) = BranchID
                .TextMatrix(i - 1, .ColIndex("StoreName")) = store
                StoreID = ChcekStore(store)
                .TextMatrix(i - 1, .ColIndex("StoreID")) = StoreID
                CusID = CheckCustomer(cus, VATNO)
                .TextMatrix(i - 1, .ColIndex("VATNO")) = VATNO
                .TextMatrix(i - 1, .ColIndex("Cus")) = cus
                .TextMatrix(i - 1, .ColIndex("CusID")) = CusID
                .TextMatrix(i - 1, .ColIndex("Value")) = value
                .TextMatrix(i - 1, .ColIndex("VATPer")) = VATPer
                .TextMatrix(i - 1, .ColIndex("PayedType")) = PayedTyp
                .TextMatrix(i - 1, .ColIndex("Account_Name")) = Account_Nam
                .TextMatrix(i - 1, .ColIndex("PaymentName")) = PaymentNam
                .TextMatrix(i - 1, .ColIndex("BoxName")) = BoxNam
                If Account_Nam <> "" Then
                .TextMatrix(i - 1, .ColIndex("Account_Code")) = ChcekAccount(Account_Nam)
                End If
                If PaymentNam <> "" Then
                .TextMatrix(i - 1, .ColIndex("PaymentID")) = ChcekPaymentType(PaymentNam)
                End If
                If BoxNam <> "" Then
                .TextMatrix(i - 1, .ColIndex("BoxID")) = ChcekBoxes(BoxNam)
                End If
                .TextMatrix(i - 1, .ColIndex("VATValue")) = Round((val(.TextMatrix(i - 1, .ColIndex("VATPer"))) * val(.TextMatrix(i - 1, .ColIndex("Value")))) / 100, 2)
                .TextMatrix(i - 1, .ColIndex("Notes")) = notes
                
                If val(.TextMatrix(i - 1, .ColIndex("PaymentID"))) > 0 Then
                StrSQL = " SELECT     dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.MaxValue,"
                StrSQL = StrSQL & "       dbo.BanksData.Account_Code"
                StrSQL = StrSQL & "      FROM         dbo.TblPaymentType LEFT OUTER JOIN"
                StrSQL = StrSQL & "        dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
                StrSQL = StrSQL & "  where  dbo.TblPaymentType.PaymentID =" & val(.TextMatrix(i - 1, .ColIndex("PaymentID"))) & ""
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If rs.RecordCount > 0 Then
               .TextMatrix(i - 1, .ColIndex("BankId")) = IIf(IsNull(rs("BankId").value), "", rs("BankId").value)
               .TextMatrix(i - 1, .ColIndex("Accountsus")) = IIf(IsNull(rs("Accountsus").value), "", rs("Accountsus").value)
               .TextMatrix(i - 1, .ColIndex("Accountcom")) = IIf(IsNull(rs("Accountcom").value), "", rs("Accountcom").value)
               .TextMatrix(i - 1, .ColIndex("commision")) = IIf(IsNull(rs("commision").value), 0, rs("commision").value)
               .TextMatrix(i - 1, .ColIndex("MaxValue")) = IIf(IsNull(rs("MaxValue").value), 0, rs("MaxValue").value)
               .TextMatrix(i - 1, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
               Else
               .TextMatrix(i - 1, .ColIndex("BankId")) = 0
               .TextMatrix(i - 1, .ColIndex("bankAccount_Code")) = ""
               .TextMatrix(i - 1, .ColIndex("Accountsus")) = ""
               .TextMatrix(i - 1, .ColIndex("Accountcom")) = ""
               .TextMatrix(i - 1, .ColIndex("commision")) = 0
               .TextMatrix(i - 1, .ColIndex("MaxValue")) = 0
               End If
              Else
                .TextMatrix(i - 1, .ColIndex("BankId")) = 0
               .TextMatrix(i - 1, .ColIndex("bankAccount_Code")) = ""
               .TextMatrix(i - 1, .ColIndex("Accountsus")) = ""
               .TextMatrix(i - 1, .ColIndex("Accountcom")) = ""
               .TextMatrix(i - 1, .ColIndex("commision")) = 0
               .TextMatrix(i - 1, .ColIndex("MaxValue")) = 0
               End If
               
                Branch = ""
            End With
            If .cells(i, 1) & "" = "" Then Exit Sub
                i = i + 1
                Loop
        End With
        'ReLineGrid
        Me.VSFlexGrid.SetFocus
        ExcelObj.Workbooks.Close

        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelObj = Nothing
    End If
End Sub
Private Sub ISButton4_Click()
    VSFlexGrid.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid.rows = 1
    CD1.ShowOpen
    txtFile.text = CD1.FileName
End Sub

Private Sub RDPatmentType_Click(Index As Integer)
    TxtCheckNo.Enabled = False
    CheckDate.Enabled = False
    TxtTransferNO.Enabled = False
    TransferDate.Enabled = False
    If RDPatmentType(0).value = True Then
        TxtCheckNo.Enabled = True
        CheckDate.Enabled = True
    End If
    If RDPatmentType(1).value = True Then
        TxtTransferNO.Enabled = True
        TransferDate.Enabled = True
    End If
End Sub

Private Sub RdTyp_Click(Index As Integer)

    ISButton4.Enabled = False
    ISButton3.Enabled = False
    If RdTyp(0).value = True Then
        C1Elastic2.Enabled = True
    Else
        ISButton4.Enabled = True
        ISButton3.Enabled = True
        C1Elastic2.Enabled = False
    End If
End Sub

Function CheckCustomer(Optional CusName As String, Optional ByRef VATNO As String) As Double

    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim sql As String
    
        sql = "Select * from TblCustemers where VATNO = '" & VATNO & "' or CusNamee = '" & CusName & "' or CusName = '" & CusName & "'"
        RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        If RsTemp.RecordCount > 0 Then
            VATNO = IIf(IsNull(RsTemp("VATNO").value), "", RsTemp("VATNO").value)
            CheckCustomer = IIf(IsNull(RsTemp("CusID").value), 0, RsTemp("CusID").value)
        Else
           CheckCustomer = AddCus(CusName, VATNO)
        End If
End Function
Function AddStore(Optional StoreName As String) As Double

    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim StrRecID As String
    Dim sql As String
    
    StrRecID = new_id("TblStore", "StoreID", "")
    sql = "Select * from TblStore where 1 = -1"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    RsTemp.AddNew
    RsTemp("StoreID").value = IIf(StrRecID <> "", StrRecID, Null)
    RsTemp("StoreName").value = StoreName
    RsTemp("StoreNamee").value = StoreName
    RsTemp.update
    AddStore = StrRecID
End Function
Function AddBranch(Optional Branch As String) As Double

    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim StrRecID As String
    Dim sql As String
    
    StrRecID = new_id("TblBranchesData", "branch_id", "")
    sql = "Select * from TblBranchesData where 1 = -1"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    RsTemp.AddNew
    RsTemp("branch_id").value = IIf(StrRecID <> "", StrRecID, Null)
    RsTemp("branch_name").value = Branch
    RsTemp("branch_namee").value = Branch
    RsTemp.update
    AddBranch = StrRecID
End Function
Function ChcekBoxes(Optional BoxName As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from TblBoxesData where BoxName like '%" & BoxName & "%' or BoxNameE like '%" & BoxName & "%'"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If RsTemp.RecordCount > 0 Then
        ChcekBoxes = IIf(IsNull(RsTemp("BoxID").value), 0, RsTemp("BoxID").value)
    Else
        ChcekBoxes = ""
    End If
End Function
Function ChcekPaymentType(Optional PaymentName As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from TblPaymentType where PaymentName like '%" & PaymentName & "%' Or PaymentNamee like '%" & PaymentName & "%'"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If RsTemp.RecordCount > 0 Then
        ChcekPaymentType = IIf(IsNull(RsTemp("PaymentID").value), 0, RsTemp("PaymentID").value)
    Else
        ChcekPaymentType = ""
    End If
End Function
Function ChcekAccount(Optional account_name As String) As String
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Name like '%" & account_name & "%' or Account_NameEng like '%" & account_name & "%'"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If RsTemp.RecordCount > 0 Then
        ChcekAccount = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
    Else
        ChcekAccount = ""
    End If
End Function
Function ChcekBranch(Optional Branch As String) As Double

    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from TblBranchesData where branch_name like '%" & Branch & "%' or branch_namee like '%" & Branch & "%'"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    If RsTemp.RecordCount > 0 Then
        ChcekBranch = IIf(IsNull(RsTemp("branch_id").value), 0, RsTemp("branch_id").value)
    Else
       ' ChcekBranch = AddBranch(Branch)
       ChcekBranch = 0
    End If
End Function
Function ChcekStore(Optional StoreName As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from TblStore where StoreName like '%" & StoreName & "%' or StoreNamee like '%" & StoreName & "%'"
    RsTemp.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    If RsTemp.RecordCount > 0 Then
        ChcekStore = IIf(IsNull(RsTemp("StoreID").value), 0, RsTemp("StoreID").value)
    Else
       ' ChcekStore = AddStore(StoreName)
       ChcekStore = 0
    End If
End Function
Private Sub ValueTxt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, ValueTxt.text, 0)
End Sub
Private Sub VATPerTxt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, VATPerTxt.text, 0)
End Sub
Private Sub Form_Load()
    Dim conection As String
    Dim My_SQL As String
    Dim Dcombos As New ClsDataCombos
    Dim SqlQu As String
    
    On Error GoTo ErrTrap
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.VSFlexGrid.ColComboList(VSFlexGrid.ColIndex("TypeService")) = "#1;”·⁄…|#2; Œœ„…"
        VSFlexGrid.ColComboList(VSFlexGrid.ColIndex("PayedType")) = "#1;‰ﬁœ« |#2;«Ã·"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        VSFlexGrid.ColComboList(VSFlexGrid.ColIndex("TypeService")) = "#1;Commodity |#2;Service"
        VSFlexGrid.ColComboList(VSFlexGrid.ColIndex("PayedType")) = "#1;Cash |#2;Credit"
    End If
    
    Dcombos.GetBranches dcBranch
    Dcombos.GetBranches BranchDCDet
    Dcombos.GetStores DcbStore
    Dcombos.GetUsers DCboUserName
    Dcombos.GetBoxes DcboBox
    Dcombos.GetPaymentType Me.DcbPayment, -1
    Dcombos.GetAccountingCodes Me.DcbAccount, True
    If SystemOptions.UserInterface = ArabicInterface Then
        With CboPayMentType
        .AddItem "‰ﬁœ«"
        .AddItem "¬Ã·"
    End With
        SqlQu = "select CusID , CusName from TblCustemers where TblCustemers.Type in (1,2)"
    Else
        CboPayMentType.Clear
    CboPayMentType.AddItem "Cash"
    CboPayMentType.AddItem "Credit"
        SqlQu = "select CusID , CusNamee from TblCustemers where TblCustemers.Type in (1,2)"
    End If
    fill_combo CusDC, SqlQu
    
    If SystemOptions.UserInterface = ArabicInterface Then
        SqlQu = "select ID , VatTypeName from VatTypes "
    Else
        SqlQu = "select ID , VatTypeNamee from VatTypes "
    End If
    fill_combo VATTypeTransDC, SqlQu
    conection = "select * from TblVATSettings order by  ID "
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    
    BtnLast_Click
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
    Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        'btnNew_Click
    End If
   Me.Refresh
   
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
    ISButton4.Caption = "Select Path"
    lbl(16).Caption = "Payment Data"
    lbl(17).Caption = "Check No."
    lbl(18).Caption = "Date"
    lbl(20).Caption = "Date"
    lbl(21).Caption = "Store"
    lbl(19).Caption = "Transfer No. "
    lbl(15).Caption = "VAT No."
 '   RDPatmentType(0).RightToLeft = False
 '   RDPatmentType(1).RightToLeft = False
    RDPatmentType(0).Caption = "Cheque"
    RDPatmentType(1).Caption = "Transfer  "
    RdTyp(0).RightToLeft = False
    RdTyp(1).RightToLeft = False
    RdTyp(0).Caption = "Manual"
    ISButton3.Caption = "Import"
    RdTyp(1).Caption = "From a File"
    Me.Caption = "VAT Transactions"
    Label1(2).Caption = Me.Caption
    
    lbl(4).Caption = "No."
    lbl(2).Caption = "Date"
    lbl(0).Caption = "Branch"
    lbl(5).Caption = "Transaction Type"
    ServiceRd.Caption = "Service"
    ProuductRd.Caption = "Product"
    
    lbl(1).Caption = "Document No."
    lbl(3).Caption = "Document Date"
    lbl(9).Caption = "Organization"
    lbl(10).Caption = "Branch"
    lbl(6).Caption = "Value"
    lbl(7).Caption = "VAT percentage"
    lbl(8).Caption = "VAT Value"
    lbl(11).Caption = "Notes"
    AddToGrid.Caption = "Add"
    
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    
    With Me.VSFlexGrid
        .TextMatrix(0, .ColIndex("Ser")) = "No."
        .TextMatrix(0, .ColIndex("DocNo")) = "Document No."
        .TextMatrix(0, .ColIndex("DocDate")) = "Document Date"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("Cus")) = "Organization"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("VATPer")) = "VAT percentage"
        .TextMatrix(0, .ColIndex("VATValue")) = "VAT Value"
        .TextMatrix(0, .ColIndex("Notes")) = "Notes"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store"
        .TextMatrix(0, .ColIndex("TypeService")) = "Type"
        .TextMatrix(0, .ColIndex("VATNO")) = "VAT No."
    End With
    
    DelRow.Caption = "Delete current row "
    DelAll.Caption = "Delete All"
    
    lbl(13).Caption = "By"
    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "Records count"
Exit Sub
ErrTrap:
End Sub
Private Sub cleargriid()
    Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
    Me.VSFlexGrid.rows = 2
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    cleargriid
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
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    cleargriid
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
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnNext_Click()
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      cleargriid
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
    cleargriid
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
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    cleargriid
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
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub AddNewRecored()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    My_SQL = "select Max(ID) as MaxID from TblVATSettings"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rs.RecordCount > 0 Then
        TxtSerial1.text = IIf(IsNull(rs("MaxID").value), 0, rs("MaxID").value) + 1
    Else
        TxtSerial1.text = 1
    End If
    
   rs.Close
   
Exit Sub
ErrTrap:
End Sub
Public Sub AddNewRec()
    Dim StrRecID As String
    
    On Error GoTo ErrTrap

    StrRecID = new_id("TblVATSettings", "ID", "")
    TxtSerial1.text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
    
Exit Sub
ErrTrap:
End Sub
Public Sub FiLLRec()
    Dim sql As String
        
    On Error GoTo ErrTrap

    If TxtModFlg = "E" Then
        StrSQL = "Delete From TblVATSettingsDet Where VATSettingsID = " & val(TxtSerial1.text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("TransIndx").value = val(Me.VATTypeTransDC.BoundText)
    RsSavRec.Fields("TransType").value = val(Me.TransTypeTblID.text)
    RsSavRec.Fields("VATNO").value = TxtVATNO.text
    RsSavRec.Fields("CheckDate").value = CheckDate.value
    RsSavRec.Fields("TransferDate").value = TransferDate.value
    RsSavRec.Fields("CheckNo").value = TxtCheckNo.text
    RsSavRec.Fields("TransferNO").value = TxtTransferNO.text
    RsSavRec.Fields("BoxID").value = val(Me.DcboBox.BoundText)
    RsSavRec.Fields("PaymentID").value = val(Me.DcbPayment.BoundText)
    RsSavRec.Fields("PayedType").value = val(Me.CboPayMentType.ListIndex)
    RsSavRec.Fields("Account_Code").value = DcbAccount.BoundText
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("Total").value = val(lbl(26).Caption)
    If Me.RdTyp(0).value = True Then
    RsSavRec.Fields("TpyImport").value = 0
    Else
    RsSavRec.Fields("TpyImport").value = 1
    End If
    If Me.RDPatmentType(0).value = True Then
    RsSavRec.Fields("PatmentType").value = 0
    Else
    RsSavRec.Fields("PatmentType").value = 1
    End If
    If ServiceRd.value = True Then
        RsSavRec.Fields("SerOrProud").value = 1
    ElseIf ProuductRd.value = True Then
        RsSavRec.Fields("SerOrProud").value = 0
    End If
    RsSavRec.update
    
    '#################################### Det Part #################################
    Dim i As Integer
    Dim str2 As String
    
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblVATSettingsDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With VSFlexGrid
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("DocNo")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("VATSettingsID").value = val(Me.TxtSerial1.text)
                RsDevsub("TypeService").value = IIf((.TextMatrix(i, .ColIndex("TypeService"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeService"))))
                RsDevsub("DocNo").value = IIf((.TextMatrix(i, .ColIndex("DocNo"))) = "", Null, .TextMatrix(i, .ColIndex("DocNo")))
                RsDevsub("DocDate").value = IIf((.TextMatrix(i, .ColIndex("DocDate"))) = "", Null, .TextMatrix(i, .ColIndex("DocDate")))
                RsDevsub("Value").value = IIf((.TextMatrix(i, .ColIndex("Value"))) = "", Null, ((.TextMatrix(i, .ColIndex("Value")))))
                RsDevsub("VATPer").value = IIf((.TextMatrix(i, .ColIndex("VATPer"))) = "", Null, ((.TextMatrix(i, .ColIndex("VATPer")))))
                RsDevsub("VATValue").value = IIf((.TextMatrix(i, .ColIndex("VATValue"))) = "", Null, ((.TextMatrix(i, .ColIndex("VATValue")))))
                RsDevsub("CusID").value = IIf((.TextMatrix(i, .ColIndex("CusID"))) = "", Null, ((.TextMatrix(i, .ColIndex("CusID")))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, ((.TextMatrix(i, .ColIndex("BranchID")))))
                RsDevsub("Notes").value = IIf((.TextMatrix(i, .ColIndex("Notes"))) = "", Null, ((.TextMatrix(i, .ColIndex("Notes")))))
                RsDevsub("VATNO").value = IIf((.TextMatrix(i, .ColIndex("VATNO"))) = "", Null, ((.TextMatrix(i, .ColIndex("VATNO")))))
                RsDevsub("StoreId").value = IIf((.TextMatrix(i, .ColIndex("StoreId"))) = "", Null, ((.TextMatrix(i, .ColIndex("StoreId")))))
                RsDevsub("Account_Code").value = IIf((.TextMatrix(i, .ColIndex("Account_Code"))) = "", Null, ((.TextMatrix(i, .ColIndex("Account_Code")))))
                RsDevsub("PayedType").value = IIf((.TextMatrix(i, .ColIndex("PayedType"))) = "", Null, val(((.TextMatrix(i, .ColIndex("PayedType"))))))
                RsDevsub("PaymentID").value = IIf((.TextMatrix(i, .ColIndex("PaymentID"))) = "", Null, val((.TextMatrix(i, .ColIndex("PaymentID")))))
                RsDevsub("BankId").value = IIf((.TextMatrix(i, .ColIndex("BankId"))) = "", Null, val((.TextMatrix(i, .ColIndex("BankId")))))
                RsDevsub("Accountcom").value = IIf((.TextMatrix(i, .ColIndex("Accountcom"))) = "", Null, ((.TextMatrix(i, .ColIndex("Accountcom")))))
                RsDevsub("MaxValue").value = IIf((.TextMatrix(i, .ColIndex("MaxValue"))) = "", Null, val((.TextMatrix(i, .ColIndex("MaxValue")))))
                RsDevsub("commision").value = IIf((.TextMatrix(i, .ColIndex("commision"))) = "", Null, val((.TextMatrix(i, .ColIndex("commision")))))
                RsDevsub("BoxID").value = IIf((.TextMatrix(i, .ColIndex("BoxID"))) = "", Null, val((.TextMatrix(i, .ColIndex("BoxID")))))
                RsDevsub("bankAccount_Code").value = IIf((.TextMatrix(i, .ColIndex("bankAccount_Code"))) = "", Null, ((.TextMatrix(i, .ColIndex("bankAccount_Code")))))
                RsDevsub.update
            End If
        Next i
    End With
    Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
                Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                If SystemOptions.UserInterface = ArabicInterface Then
                Else
                    Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
                    Me.Refresh
                    FiLLTXT
                    TxtModFlg = "R"
                    MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
            Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
        Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
        End Select
       RsSavRec.Resync adAffectCurrent
       
Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT()
    Dim i As Integer
    Dim ContactTime  As Date
    On Error GoTo ErrTrap
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), 0, RsSavRec.Fields("BranchID").value)
    VATTypeTransDC.BoundText = IIf(IsNull(RsSavRec.Fields("TransIndx").value), -1, RsSavRec.Fields("TransIndx").value)
    TransTypeTblID.text = IIf(IsNull(RsSavRec.Fields("TransType").value), "", RsSavRec.Fields("TransType").value)
    TxtVATNO.text = IIf(IsNull(RsSavRec.Fields("VATNO").value), "", RsSavRec.Fields("VATNO").value)
    TxtCheckNo.text = IIf(IsNull(RsSavRec.Fields("CheckNo").value), "", RsSavRec.Fields("CheckNo").value)
    TxtTransferNO.text = IIf(IsNull(RsSavRec.Fields("TransferNO").value), "", RsSavRec.Fields("TransferNO").value)
    CheckDate.value = IIf(IsNull(RsSavRec.Fields("CheckDate").value), Date, RsSavRec.Fields("CheckDate").value)
    TransferDate.value = IIf(IsNull(RsSavRec.Fields("TransferDate").value), Date, RsSavRec.Fields("TransferDate").value)
    Me.DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID").value), "", RsSavRec.Fields("BoxID").value)
    Me.DcbAccount.BoundText = IIf(IsNull(RsSavRec.Fields("Account_Code").value), "", RsSavRec.Fields("Account_Code").value)
    Me.CboPayMentType.ListIndex = IIf(IsNull(RsSavRec.Fields("PayedType").value), -1, RsSavRec.Fields("PayedType").value)
    DcbPayment.BoundText = IIf(IsNull(RsSavRec.Fields("PaymentID").value), "", RsSavRec.Fields("PaymentID").value)
   Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
   lbl(26).Caption = IIf(IsNull(RsSavRec.Fields("Total").value), "", RsSavRec.Fields("Total").value)
   TXTNoteID.text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
   TxtNoteSerial.text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
 
    If Not (IsNull(RsSavRec.Fields("TpyImport").value)) Then
    If (RsSavRec.Fields("TpyImport").value) = 1 Then
    RdTyp(1).value = True
    Else
    RdTyp(0).value = True
    End If
    Else
    RdTyp(0).value = True
    End If
    If Not (IsNull(RsSavRec.Fields("PatmentType").value)) Then
    If (RsSavRec.Fields("PatmentType").value) = 1 Then
    RDPatmentType(1).value = True
    Else
    RDPatmentType(0).value = True
    End If
    Else
    RDPatmentType(0).value = True
    End If
    If Not (IsNull(RsSavRec.Fields("SerOrProud").value)) Then
        If RsSavRec.Fields("SerOrProud").value = 0 Then
            ServiceRd.value = True
            ProuductRd.value = False
        ElseIf RsSavRec.Fields("SerOrProud").value = 1 Then
            ServiceRd.value = False
            ProuductRd.value = True
        End If
    End If
    
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    
    FullGrid
    
Exit Sub
ErrTrap:

End Sub
Sub FullGrid()
    Dim sql As String
    Dim i As Integer
    
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    
    On Error GoTo ErrTrap
    
    sql = " SELECT     dbo.TblVATSettingsDet.ID, dbo.TblVATSettingsDet.VATSettingsID, dbo.TblVATSettingsDet.DocNo, dbo.TblVATSettingsDet.DocDate, dbo.TblVATSettingsDet.[Value], "
    sql = sql & "                  dbo.TblVATSettingsDet.VATPer, dbo.TblVATSettingsDet.VATValue, dbo.TblVATSettingsDet.CusID, dbo.TblVATSettingsDet.BranchID, dbo.TblVATSettingsDet.Notes,"
    sql = sql & "                  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblVATSettingsDet.VATNO,"
    sql = sql & "                   dbo.TblVATSettingsDet.TypeService, dbo.TblVATSettingsDet.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblVATSettingsDet.BoxID,"
    sql = sql & "                  dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.TblVATSettingsDet.PayedType, dbo.TblVATSettingsDet.PaymentID,"
    sql = sql & "                  dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.TblVATSettingsDet.BankId, dbo.TblVATSettingsDet.Accountcom,"
    sql = sql & "                  dbo.TblVATSettingsDet.bankAccount_Code, dbo.TblVATSettingsDet.MaxValue, dbo.TblVATSettingsDet.commision, dbo.TblVATSettingsDet.Account_Code,"
    sql = sql & "                  dbo.Accounts.account_name , dbo.Accounts.Account_NameEng"
    sql = sql & "    FROM         dbo.TblVATSettingsDet LEFT OUTER JOIN"
    sql = sql & "                  dbo.ACCOUNTS ON dbo.TblVATSettingsDet.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblPaymentType ON dbo.TblVATSettingsDet.PaymentID = dbo.TblPaymentType.PaymentID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBoxesData ON dbo.TblVATSettingsDet.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStore ON dbo.TblVATSettingsDet.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblVATSettingsDet.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblVATSettingsDet.BranchID = dbo.TblBranchesData.branch_id"
    sql = sql & " Where TblVATSettingsDet.VATSettingsID = " & val(TxtSerial1.text)
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
    With Me.VSFlexGrid
        For i = .FixedRows To Rs1.RecordCount
            .rows = .FixedRows + Rs1.RecordCount
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(Rs1("BoxID").value), "", Rs1("BoxID").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(Rs1("Accountcom").value), "", Rs1("Accountcom").value)
            .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(Rs1("MaxValue").value), "", Rs1("MaxValue").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(Rs1("commision").value), "", Rs1("commision").value)
            .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(Rs1("bankAccount_Code").value), "", Rs1("bankAccount_Code").value)
            .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
            .TextMatrix(i, .ColIndex("PayedType")) = IIf(IsNull(Rs1("PayedType").value), 1, Rs1("PayedType").value)
            .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(Rs1("PaymentID").value), "", Rs1("PaymentID").value)
            .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(Rs1("BankId").value), "", Rs1("BankId").value)
            .TextMatrix(i, .ColIndex("TypeService")) = IIf(IsNull(Rs1("TypeService").value), 1, Rs1("TypeService").value)
            .TextMatrix(i, .ColIndex("VATNO")) = IIf(IsNull(Rs1("VATNO").value), "", Rs1("VATNO").value)
            .TextMatrix(i, .ColIndex("DocNo")) = IIf(IsNull(Rs1("DocNo").value), "", Rs1("DocNo").value)
            .TextMatrix(i, .ColIndex("DocDate")) = IIf(IsNull(Rs1("DocDate").value), "", Rs1("DocDate").value)
            .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(Rs1("Value").value), "", Rs1("Value").value)
            .TextMatrix(i, .ColIndex("VATPer")) = IIf(IsNull(Rs1("VATPer").value), "", Rs1("VATPer").value)
            .TextMatrix(i, .ColIndex("VATValue")) = IIf(IsNull(Rs1("VATValue").value), "", Rs1("VATValue").value)
            .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)
            .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs1("CusID").value), "", Rs1("CusID").value)
            .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs1("StoreID").value), "", Rs1("StoreID").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxName").value), "", Rs1("BoxName").value)
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(Rs1("PaymentName").value), "", Rs1("PaymentName").value)
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreName").value), "", Rs1("StoreName").value)
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                .TextMatrix(i, .ColIndex("Cus")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
            Else
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxNameE").value), "", Rs1("BoxNameE").value)
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(Rs1("PaymentNamee").value), "", Rs1("PaymentNamee").value)
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreNamee").value), "", Rs1("StoreNamee").value)
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                .TextMatrix(i, .ColIndex("Cus")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
            End If
            .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(Rs1("Notes").value), "", Rs1("Notes").value)
            
            Rs1.MoveNext
        Next i
    End With
    
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
    
    MySQL = "SELECT TblVATSettingsDet.ID, TblVATSettingsDet.VATSettingsID, TblVATSettingsDet.DocNo, TblVATSettingsDet.DocDate, TblVATSettingsDet.Value, TblVATSettingsDet.VATPer, TblVATSettingsDet.VATValue, "
    MySQL = MySQL & " TblVATSettingsDet.CusID, TblVATSettingsDet.BranchID, TblVATSettingsDet.Notes, TblCustemers.CusName, TblCustemers.CusNamee, TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, "
    MySQL = MySQL & " TblVATSettings.RecordDate, dbo.TblVATSettings.BranchID AS BranchH, TblVATSettings.TransIndx, TblVATSettings.TransType, TblVATSettings.SerOrProud, TblBranchesData.branch_name AS BranchNameH, "
    MySQL = MySQL & " TblBranchesData.branch_namee AS BranchNameeH "
    MySQL = MySQL & " FROM TblBranchesData RIGHT OUTER JOIN "
    MySQL = MySQL & " TblVATSettings ON TblBranchesData.branch_id = TblVATSettings.BranchID FULL OUTER JOIN "
    MySQL = MySQL & " TblVATSettingsDet LEFT OUTER JOIN "
    MySQL = MySQL & " TblCustemers ON TblVATSettingsDet.CusID = TblCustemers.CusID LEFT OUTER JOIN "
    MySQL = MySQL & " TblBranchesData AS TblBranchesData_1 ON TblVATSettingsDet.BranchID = TblBranchesData_1.branch_id ON TblVATSettings.ID = TblVATSettingsDet.VATSettingsID "
    MySQL = MySQL & " Where (dbo.TblVATSettings.id =" & val(TxtSerial1.text) & ")"
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAT\" & "RepVATSettings.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAT\" & "RepVATSettingsE.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
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

    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
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
Public Sub EditRec(StrTable As String, RecId As String)
    FiLLRec
End Sub
Public Function FindRec(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec.Find "ID = " & RecId, , adSearchForward, 1
    
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
Private Sub TransTypeCBox_Change()
    TransTypeCBox_Click
End Sub
Private Sub TxtSerial1_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
Private Sub TransTypeCBox_Click()
    Select Case TransTypeCBox.ListIndex
        Case 0
            'ServiceRd.value = False
            'ServiceRd.Enabled = True
            'ProuductRd.value = True
            'ProuductRd.Enabled = True
            
            TransTypeTblID.text = "1"
        Case 1
            'ServiceRd.value = False
            'ServiceRd.Enabled = True
            'ProuductRd.value = True
            'ProuductRd.Enabled = True
            
            TransTypeTblID.text = "5"
        Case 2
           ' ServiceRd.value = False
           ' ServiceRd.Enabled = True
           ' ProuductRd.value = True
           ' ProuductRd.Enabled = True
            
            TransTypeTblID.text = "2"
        Case 3
           ' ServiceRd.value = False
           ' ServiceRd.Enabled = True
           ' ProuductRd.value = True
           ' ProuductRd.Enabled = True
            
            TransTypeTblID.text = "9"
        Case 4
           ' ServiceRd.value = True
           ' ServiceRd.Enabled = False
           ' ProuductRd.value = False
           ' ProuductRd.Enabled = False
            
            TransTypeTblID = ""
        Case 5
           ' ServiceRd.value = True
           ' ServiceRd.Enabled = False
           ' ProuductRd.value = False
           ' ProuductRd.Enabled = False
            
           TransTypeTblID = ""
        End Select
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.text = "N" Then
        TransTypeCBox.Enabled = True
        XPDtbTrans.Enabled = True
        
        DocNo.Enabled = True
        DocDate.Enabled = True
        CusDC.Enabled = True
        BranchDCDet.Enabled = True
        ValueTxt.Enabled = True
        VATPerTxt.Enabled = True
        NotesTxt.Enabled = True
        AddToGrid.Enabled = True
        
        VSFlexGrid.Enabled = True
        DelRow.Enabled = True
        DelAll.Enabled = True
        
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.ISButton8.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        

        
    ElseIf TxtModFlg.text = "R" Then
    
        TransTypeCBox.Enabled = False
        XPDtbTrans.Enabled = False
        
        DocNo.Enabled = False
        DocDate.Enabled = False
        CusDC.Enabled = False
        BranchDCDet.Enabled = False
        ValueTxt.Enabled = False
        VATPerTxt.Enabled = False
        NotesTxt.Enabled = False
        AddToGrid.Enabled = False
        
        VSFlexGrid.Enabled = False
        DelRow.Enabled = False
        DelAll.Enabled = False
    
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
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
        
   ElseIf TxtModFlg.text = "E" Then
        TransTypeCBox.Enabled = True
        XPDtbTrans.Enabled = True
        
        DocNo.Enabled = True
        DocDate.Enabled = True
        CusDC.Enabled = True
        BranchDCDet.Enabled = True
        ValueTxt.Enabled = True
        VATPerTxt.Enabled = True
        NotesTxt.Enabled = True
        AddToGrid.Enabled = True
        
        VSFlexGrid.Enabled = True
        DelRow.Enabled = True
        DelAll.Enabled = True
        
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
Sub FillGrid2()
Dim VATNO As String
Dim k As Integer
Dim rs As ADODB.Recordset
Dim StrSQL As String
    With VSFlexGrid
        k = .rows - 1
        .rows = .rows + 1
        Do While k < (.rows - 1)
            .TextMatrix(k, .ColIndex("Ser")) = k
            If ServiceRd.value = True Then
            .TextMatrix(k, .ColIndex("TypeService")) = 2
            Else
            .TextMatrix(k, .ColIndex("TypeService")) = 1
            End If
            .TextMatrix(k, .ColIndex("DocNo")) = DocNo.text
            .TextMatrix(k, .ColIndex("DocDate")) = FormatDateTime(Me.DocDate.value, vbShortDate)
            VATNO = TxtVATNO.text
            .TextMatrix(k, .ColIndex("CusID")) = CheckCustomer(CusDC.text, VATNO)
            .TextMatrix(k, .ColIndex("Cus")) = CusDC.text
            .TextMatrix(k, .ColIndex("BranchID")) = ChcekBranch(BranchDCDet.text)
            .TextMatrix(k, .ColIndex("Branch")) = BranchDCDet.text
            .TextMatrix(k, .ColIndex("Value")) = ValueTxt.text
            .TextMatrix(k, .ColIndex("VATPer")) = VATPerTxt.text
            .TextMatrix(k, .ColIndex("VatValue")) = VatValueTxt.text
            .TextMatrix(k, .ColIndex("Notes")) = NotesTxt.text
            .TextMatrix(k, .ColIndex("StoreID")) = ChcekStore(DcbStore.text)
            .TextMatrix(k, .ColIndex("StoreName")) = DcbStore.text
            .TextMatrix(k, .ColIndex("VATNO")) = TxtVATNO.text
            .TextMatrix(k, .ColIndex("PayedType")) = val(CboPayMentType.ListIndex) + 1
            .TextMatrix(k, .ColIndex("PaymentID")) = val(DcbPayment.BoundText)
            .TextMatrix(k, .ColIndex("PaymentName")) = DcbPayment.text
            .TextMatrix(k, .ColIndex("Account_Name")) = DcbAccount.text
            .TextMatrix(k, .ColIndex("Account_Code")) = DcbAccount.BoundText
            .TextMatrix(k, .ColIndex("BoxID")) = val(DcboBox.BoundText)
            .TextMatrix(k, .ColIndex("BoxName")) = DcboBox.text
        
               If val(.TextMatrix(k, .ColIndex("PaymentID"))) > 0 Then
                StrSQL = " SELECT     dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.MaxValue,"
                StrSQL = StrSQL & "       dbo.BanksData.Account_Code"
                StrSQL = StrSQL & "      FROM         dbo.TblPaymentType LEFT OUTER JOIN"
                StrSQL = StrSQL & "        dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
                StrSQL = StrSQL & "  where  dbo.TblPaymentType.PaymentID =" & val(.TextMatrix(k, .ColIndex("PaymentID"))) & ""
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If rs.RecordCount > 0 Then
               .TextMatrix(k, .ColIndex("BankId")) = IIf(IsNull(rs("BankId").value), "", rs("BankId").value)
               .TextMatrix(k, .ColIndex("Accountsus")) = IIf(IsNull(rs("Accountsus").value), "", rs("Accountsus").value)
               .TextMatrix(k, .ColIndex("Accountcom")) = IIf(IsNull(rs("Accountcom").value), "", rs("Accountcom").value)
               .TextMatrix(k, .ColIndex("commision")) = IIf(IsNull(rs("commision").value), 0, rs("commision").value)
               .TextMatrix(k, .ColIndex("MaxValue")) = IIf(IsNull(rs("MaxValue").value), 0, rs("MaxValue").value)
               .TextMatrix(k, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
               Else
               .TextMatrix(k, .ColIndex("BankId")) = 0
               .TextMatrix(k, .ColIndex("bankAccount_Code")) = ""
               .TextMatrix(k, .ColIndex("Accountsus")) = ""
               .TextMatrix(k, .ColIndex("Accountcom")) = ""
               .TextMatrix(k, .ColIndex("commision")) = 0
               .TextMatrix(k, .ColIndex("MaxValue")) = 0
               End If
              Else
                .TextMatrix(k, .ColIndex("BankId")) = 0
               .TextMatrix(k, .ColIndex("bankAccount_Code")) = ""
               .TextMatrix(k, .ColIndex("Accountsus")) = ""
               .TextMatrix(k, .ColIndex("Accountcom")) = ""
               .TextMatrix(k, .ColIndex("commision")) = 0
               .TextMatrix(k, .ColIndex("MaxValue")) = 0
               End If

        k = k + 1
        Loop
        '.AutoSize 0, .Cols - 1, False
    End With
    ReLineGrid
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
    
Exit Sub
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Sub calcVAT()
    VatValueTxt.text = Round((IIf(ValueTxt.text = "", 0, val(ValueTxt.text)) * IIf(VATPerTxt.text = "", 0, val(VATPerTxt.text)) / 100), 2)
End Sub
Private Sub ValueTxt_Change()
    calcVAT
End Sub
Private Sub VATPerTxt_Change()
    calcVAT
End Sub
Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub
Private Sub DelRow_Click()
    RemoveMyGridRow
End Sub
Private Sub DelAll_Click()
    clearMyGrid
End Sub
Private Sub RemoveMyGridRow()
    With Me.VSFlexGrid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub clearMyGrid()
    VSFlexGrid.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid.rows = 2
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrTrap
   
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    clear_all Me
    cleargriid
    
    TxtModFlg.text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    RdTyp(0).value = True
     RdTyp_Click (0)
    GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
    Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid.rows = 2
    RDPatmentType_Click (0)
    VATPerTxt.text = 5
VATTypeTransDC.BoundText = Me.CtranIndex

Exit Sub
ErrTrap:
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    
    On Error GoTo ErrTrap
        If TxtNoteSerial.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
    Else
    MsgBox "Please Delete Voucher"
    End If
    Exit Sub
    End If
    GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
    
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
        VSFlexGrid.rows = VSFlexGrid.rows + 1
        Me.DCboUserName.BoundText = user_id
        Me.dcBranch.BoundText = branch_id
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
                Msg = Msg & "This recored can't be edited now" & CHR(13)
                Msg = Msg & "it's under modification by other user on the network"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnSave_Click()

    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    On Error GoTo ErrTrap
    
    'If Dcbranch.Text = "" Or val(Dcbranch.BoundText) = 0 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '    Else
    '        MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '    End If
    '    Dcbranch.SetFocus
    '        Exit Sub
    ' End If

If val(VATTypeTransDC.BoundText) = 0 Or VATTypeTransDC.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·Õ—ﬂ…"
Else
MsgBox "Please select Type"
End If
TransTypeCBox.SetFocus
Exit Sub
End If
Dim str As String
If CheckAccount(str) = True Then
MsgBox str
Exit Sub
End If
    Select Case Me.TxtModFlg.text
        Case "N"
            AddNewRecored
            AddNewRec
            BtnLast_Click
        Case "E"
            FiLLRec
    End Select
Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Error while saving data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
    FiLLTXT
    BtnLast_Click
End Sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim i As Integer
    Dim Msg As String
    Dim X As Integer
    On Error GoTo ErrTrap
    If TxtNoteSerial.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
    Else
    MsgBox "Please Delete Voucher"
    End If
    Exit Sub
    End If
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
            StrSQL = "Delete From TblVATSettingsDet Where VATSettingsID = " & val(TxtSerial1.text) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
            RsSavRec.Find "ID = " & val(TxtSerial1.text), , adSearchForward, 1
            RsSavRec.delete
            
            If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
            Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
            End If
            cleargriid
        End If
        Me.Refresh
        BtnNext_Click
Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub ISButton5_Click()
    print_report
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub
'##########################################################################################################
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ End My Code $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'##########################################################################################################

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
Private Sub VSFlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    With VSFlexGrid
        Select Case .ColKey(Col)
           Case "Branch"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BranchID"), False, True)
                .TextMatrix(Row, .ColIndex("BranchID")) = StrAccountCode
           Case "StoreName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("StoreID"), False, True)
                .TextMatrix(Row, .ColIndex("StoreID")) = StrAccountCode
           Case "BoxName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BoxID"), False, True)
                .TextMatrix(Row, .ColIndex("BoxID")) = StrAccountCode
                
            Case "Value"
                .TextMatrix(Row, .ColIndex("VATValue")) = (.TextMatrix(Row, .ColIndex("VATPer")) / 100) * .TextMatrix(Row, .ColIndex("value"))
            Case "VATPer"
                .TextMatrix(Row, .ColIndex("VATValue")) = (.TextMatrix(Row, .ColIndex("VATPer")) / 100) * .TextMatrix(Row, .ColIndex("value"))
            Case "BoxName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BoxID"), False, True)
                .TextMatrix(Row, .ColIndex("BoxID")) = StrAccountCode
            Case "Account_Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Account_Code"), False, True)
               .TextMatrix(Row, .ColIndex("Account_Code")) = StrAccountCode
            Case "PaymentName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PaymentID"), False, True)
               .TextMatrix(Row, .ColIndex("PaymentID")) = StrAccountCode
               If val(.TextMatrix(Row, .ColIndex("PaymentID"))) > 0 Then
                StrSQL = " SELECT     dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.MaxValue,"
                StrSQL = StrSQL & "       dbo.BanksData.Account_Code"
                StrSQL = StrSQL & "      FROM         dbo.TblPaymentType LEFT OUTER JOIN"
                StrSQL = StrSQL & "        dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
                StrSQL = StrSQL & "  where  dbo.TblPaymentType.PaymentID =" & val(.TextMatrix(Row, .ColIndex("PaymentID"))) & ""
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If rs.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("BankId")) = IIf(IsNull(rs("BankId").value), "", rs("BankId").value)
               .TextMatrix(Row, .ColIndex("Accountsus")) = IIf(IsNull(rs("Accountsus").value), "", rs("Accountsus").value)
               .TextMatrix(Row, .ColIndex("Accountcom")) = IIf(IsNull(rs("Accountcom").value), "", rs("Accountcom").value)
               .TextMatrix(Row, .ColIndex("commision")) = IIf(IsNull(rs("commision").value), 0, rs("commision").value)
               .TextMatrix(Row, .ColIndex("MaxValue")) = IIf(IsNull(rs("MaxValue").value), 0, rs("MaxValue").value)
               .TextMatrix(Row, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
               Else
               .TextMatrix(Row, .ColIndex("BankId")) = 0
               .TextMatrix(Row, .ColIndex("bankAccount_Code")) = ""
               .TextMatrix(Row, .ColIndex("Accountsus")) = ""
               .TextMatrix(Row, .ColIndex("Accountcom")) = ""
               .TextMatrix(Row, .ColIndex("commision")) = 0
               .TextMatrix(Row, .ColIndex("MaxValue")) = 0
               End If
              Else
                .TextMatrix(Row, .ColIndex("BankId")) = 0
               .TextMatrix(Row, .ColIndex("bankAccount_Code")) = ""
               .TextMatrix(Row, .ColIndex("Accountsus")) = ""
               .TextMatrix(Row, .ColIndex("Accountcom")) = ""
               .TextMatrix(Row, .ColIndex("commision")) = 0
               .TextMatrix(Row, .ColIndex("MaxValue")) = 0
               End If
               
        End Select
    End With
    ReLineGrid
End Sub
Private Sub ReLineGrid()
    Dim SumValu As Double
    SumValu = 0
    Dim i As Integer
    With VSFlexGrid
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
               SumValu = SumValu + val(.TextMatrix(i, .ColIndex("Value")))
            End If
        Next i
    End With
    lbl(26).Caption = SumValu
End Sub
Private Sub VSFlexGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Dim StrAccountCode As String
    Dim Msg As String
    Dim StrSQL As String
    Dim MyStrList As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    With VSFlexGrid
        Select Case .ColKey(Col)
        Case "Branch"
                  If SystemOptions.UserInterface = ArabicInterface Then
                       StrSQL = " SELECT     branch_name, branch_id"
                  Else
                       StrSQL = " SELECT     branch_namee, branch_id"
                  End If
                StrSQL = StrSQL + " from dbo.TblBranchesData"
                StrSQL = StrSQL & "  where    (  branch_id=0  or      branch_id in(" & Current_branchSql & "))"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                 MyStrList = .BuildComboList(rs, "branch_name", "branch_id")
                Else
                MyStrList = .BuildComboList(rs, "branch_namee", "branch_id")
                End If
                .ColComboList(.ColIndex("Branch")) = "|" & MyStrList
                End If
       Case "StoreName"
                  If SystemOptions.UserInterface = ArabicInterface Then
                       StrSQL = " SELECT     StoreName, StoreID"
                  Else
                       StrSQL = " SELECT     StoreNamee, StoreID"
                  End If
                StrSQL = StrSQL + " from dbo.TblStore"
                StrSQL = StrSQL & "  where    (  BranchId=0  or      BranchId in(" & Current_branchSql & "))"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                 MyStrList = .BuildComboList(rs, "StoreName", "StoreID")
                Else
                MyStrList = .BuildComboList(rs, "StoreNamee", "StoreID")
                End If
                .ColComboList(.ColIndex("StoreName")) = "|" & MyStrList
                End If
                
        Case "BoxName"
                  If SystemOptions.UserInterface = ArabicInterface Then
                       StrSQL = " SELECT     BoxName, BoxID"
                  Else
                       StrSQL = " SELECT     BoxNamee, BoxID"
                  End If
                StrSQL = StrSQL + " from dbo.tblBoxesData"
                StrSQL = StrSQL & "  where    (  BranchId=0  or      BranchId in(" & Current_branchSql & "))"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                 MyStrList = .BuildComboList(rs, "BoxName", "BoxID")
                Else
                MyStrList = .BuildComboList(rs, "BoxNamee", "BoxID")
                End If
                .ColComboList(.ColIndex("BoxName")) = "|" & MyStrList
                End If
         Case "Account_Name"
                If SystemOptions.UserInterface = ArabicInterface Then
                   StrSQL = "SELECT Account_Code,Account_Name From ACCOUNTS where 1=1 "
                Else
                   StrSQL = "SELECT Account_Code,Account_Nameeng From ACCOUNTS  where 1=1 "
                End If
                  StrSQL = StrSQL + "  and last_account=1 "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                 MyStrList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                MyStrList = .BuildComboList(rs, "Account_Nameeng", "Account_Code")
                End If
                .ColComboList(.ColIndex("Account_Name")) = "|" & MyStrList
                End If
           Case "PaymentName"
           
                   StrSQL = "SELECT PaymentID,PaymentName From TblPaymentType where 1=1 "
               If SystemOptions.LinkUsersWithPayment = True Then
                   StrSQL = StrSQL & " and PaymentID in (SELECT     PaynetID"
                   StrSQL = StrSQL & " From dbo.TblPaymentUser"
                   StrSQL = StrSQL & " Where (UserID = " & user_id & "))"
               End If
                   StrSQL = StrSQL & " order by PaymentID"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not (rs.BOF Or rs.EOF) Then
                 MyStrList = .BuildComboList(rs, "PaymentName", "PaymentID")
                .ColComboList(.ColIndex("PaymentName")) = "|" & MyStrList
                End If
        End Select

    End With
End Sub
