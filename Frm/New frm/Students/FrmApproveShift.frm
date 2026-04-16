VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmApproveShift 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmApproveShift.frx":0000
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
      ItemData        =   "FrmApproveShift.frx":6852
      Left            =   15480
      List            =   "FrmApproveShift.frx":6862
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
            Picture         =   "FrmApproveShift.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveShift.frx":83B1
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
      ButtonImage     =   "FrmApproveShift.frx":874B
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
      ButtonImage     =   "FrmApproveShift.frx":EFAD
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
      ButtonImage     =   "FrmApproveShift.frx":1580F
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
            ForeColor       =   &H000000FF&
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
            ForeColor       =   &H0000C000&
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
            ButtonImage     =   "FrmApproveShift.frx":15BA9
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
            ButtonImage     =   "FrmApproveShift.frx":15F43
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
            ButtonImage     =   "FrmApproveShift.frx":162DD
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
            ButtonImage     =   "FrmApproveShift.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   585
            Index           =   3
            Left            =   2160
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
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
            AutoSizeChildren=   0
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
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3480
               TabIndex        =   81
               Top             =   240
               Width           =   1590
            End
            Begin VB.ComboBox CmbMonth 
               Height          =   315
               Left            =   75
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   180
               Width           =   1485
            End
            Begin VB.ComboBox CboYear 
               Height          =   315
               Left            =   2355
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   165
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   195
               Index           =   7
               Left            =   1425
               TabIndex        =   80
               Top             =   270
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   240
               Index           =   6
               Left            =   2955
               TabIndex        =   79
               Top             =   180
               Width           =   1020
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«⁄ „«œ «·Õ÷Ê— Ê«·«‰’—«›"
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
            Picture         =   "FrmApproveShift.frx":16A11
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
            ButtonImage     =   "FrmApproveShift.frx":17E16
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
            ButtonImage     =   "FrmApproveShift.frx":1E678
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
            ButtonImage     =   "FrmApproveShift.frx":24EDA
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
            ButtonImage     =   "FrmApproveShift.frx":25274
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
            ButtonImage     =   "FrmApproveShift.frx":2560E
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
            ButtonImage     =   "FrmApproveShift.frx":25BA8
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
            ButtonImage     =   "FrmApproveShift.frx":2C40A
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
            ButtonImage     =   "FrmApproveShift.frx":2C7A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10605
            TabIndex        =   36
            Top             =   105
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   300
            Left            =   8985
            TabIndex        =   94
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   529
            Caption         =   "«Õ ”«»"
            BackColor       =   16761024
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmApproveShift.frx":2CB3E
            ColorButton     =   16761024
            ColorHoverText  =   8388608
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledText=   8388608
            ColorToggledHoverText=   8388608
            LowerToggledContent=   0   'False
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   240
            TabIndex        =   41
            Top             =   210
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1785
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   210
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   240
            Index           =   8
            Left            =   13575
            TabIndex        =   37
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
         Caption         =   "»Ì«‰«  «”«”Ì…|New Tab|»Ì«‰«  »œÊ‰ ‘› « "
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   5535
               Left            =   0
               TabIndex        =   34
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
                  Height          =   4980
                  Left            =   135
                  TabIndex        =   35
                  Top             =   180
                  Width           =   14220
                  _cx             =   25082
                  _cy             =   8784
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
                  BackColorAlternate=   16777152
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
                  Rows            =   2
                  Cols            =   41
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmApproveShift.frx":333A0
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   -1  'True
                  AutoSizeMode    =   1
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
                  Height          =   225
                  Index           =   3
                  Left            =   13080
                  TabIndex        =   42
                  Top             =   5190
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmApproveShift.frx":339BA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   4
                  Left            =   11775
                  TabIndex        =   43
                  Top             =   5190
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmApproveShift.frx":33F54
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   450
                  Left            =   0
                  TabIndex        =   84
                  TabStop         =   0   'False
                  Top             =   5160
                  Width           =   11685
                  _cx             =   20611
                  _cy             =   794
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00808080&
                     Caption         =   "·Ì” ·Â »’„…"
                     ForeColor       =   &H8000000B&
                     Height          =   285
                     Index           =   17
                     Left            =   6000
                     TabIndex        =   96
                     Top             =   120
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFC0C0&
                     Caption         =   "·Ì” ·Â ﬂÊœ «·„ﬂ‰…"
                     ForeColor       =   &H8000000B&
                     Height          =   285
                     Index           =   16
                     Left            =   7440
                     TabIndex        =   95
                     Top             =   120
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H0000C000&
                     Caption         =   "·Ì” ·Â «÷«›Ì"
                     ForeColor       =   &H80000007&
                     Height          =   285
                     Index           =   15
                     Left            =   120
                     TabIndex        =   90
                     Top             =   120
                     Width           =   1020
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FF8080&
                     Caption         =   "·„ Ì”Ã· Œ—ÊÕ"
                     ForeColor       =   &H80000007&
                     Height          =   285
                     Index           =   14
                     Left            =   1215
                     TabIndex        =   89
                     Top             =   120
                     Width           =   1140
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H000000FF&
                     Caption         =   "·„ Ì”Ã· ›Ì «⁄œ«œ  «·‘› « "
                     ForeColor       =   &H8000000B&
                     Height          =   285
                     Index           =   13
                     Left            =   2460
                     TabIndex        =   88
                     Top             =   120
                     Width           =   2055
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00404080&
                     Caption         =   "€Ì— „— »ÿ »«·ﬁ”„"
                     ForeColor       =   &H8000000B&
                     Height          =   285
                     Index           =   12
                     Left            =   4635
                     TabIndex        =   87
                     Top             =   120
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H0080C0FF&
                     Caption         =   "·«ÌÊÃœ ›Ì „·› «·„ÊŸ›Ì‰"
                     Height          =   285
                     Index           =   10
                     Left            =   8715
                     TabIndex        =   86
                     Top             =   120
                     Width           =   1725
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "œ·«·«  «·«·Ê«‰"
                     Height          =   285
                     Index           =   9
                     Left            =   10275
                     TabIndex        =   85
                     Top             =   120
                     Width           =   1140
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   975
               Left            =   0
               TabIndex        =   48
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
                  Left            =   10755
                  MaxLength       =   50
                  TabIndex        =   50
                  Top             =   615
                  Width           =   750
               End
               Begin XtremeSuiteControls.CheckBox SelectBranch 
                  Height          =   240
                  Left            =   11595
                  TabIndex        =   49
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
                  Left            =   12735
                  TabIndex        =   51
                  Top             =   225
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
                  Left            =   11775
                  TabIndex        =   52
                  Top             =   615
                  Width           =   2340
                  _Version        =   786432
                  _ExtentX        =   4128
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
                  Left            =   6930
                  TabIndex        =   53
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
                  Left            =   6930
                  TabIndex        =   54
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   240
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcpDept1 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   55
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   240
                  Width           =   4380
                  _ExtentX        =   7726
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   780
                  Left            =   120
                  TabIndex        =   56
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   120
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   1376
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
                  ButtonImage     =   "FrmApproveShift.frx":344EE
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin XtremeSuiteControls.CheckBox SelectDept 
                  Height          =   240
                  Left            =   5550
                  TabIndex        =   57
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
                  Left            =   1080
                  TabIndex        =   58
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   615
                  Width           =   4380
                  _ExtentX        =   7726
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox SelectProject 
                  Height          =   240
                  Left            =   5550
                  TabIndex        =   59
                  Top             =   615
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰ "
                  ForeColor       =   &H00800000&
                  Height          =   270
                  Index           =   0
                  Left            =   12855
                  TabIndex        =   60
                  Top             =   0
                  Width           =   1560
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   6540
            Left            =   15180
            TabIndex        =   46
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
               TabIndex        =   47
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                  Height          =   6480
                  Left            =   45
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   14445
                  _cx             =   25479
                  _cy             =   11430
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
                     Height          =   5550
                     Left            =   135
                     TabIndex        =   92
                     Top             =   210
                     Width           =   14220
                     _cx             =   25082
                     _cy             =   9790
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
                     BackColorAlternate=   16777152
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
                     Rows            =   2
                     Cols            =   3
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmApproveShift.frx":3AD50
                     ScrollTrack     =   0   'False
                     ScrollBars      =   3
                     ScrollTips      =   0   'False
                     MergeCells      =   0
                     MergeCompare    =   0
                     AutoResize      =   -1  'True
                     AutoSizeMode    =   1
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
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6540
            Left            =   15480
            TabIndex        =   103
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   5220
               Left            =   135
               TabIndex        =   104
               Top             =   540
               Width           =   14220
               _cx             =   25082
               _cy             =   9208
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
               BackColorAlternate=   16777152
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
               Rows            =   2
               Cols            =   41
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmApproveShift.frx":3ADBA
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   450
               Left            =   0
               TabIndex        =   105
               TabStop         =   0   'False
               Top             =   6000
               Width           =   11685
               _cx             =   20611
               _cy             =   794
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ·«·«  «·«·Ê«‰"
                  Height          =   285
                  Index           =   25
                  Left            =   10275
                  TabIndex        =   113
                  Top             =   120
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0080C0FF&
                  Caption         =   "·«ÌÊÃœ ›Ì „·› «·„ÊŸ›Ì‰"
                  Height          =   285
                  Index           =   24
                  Left            =   8715
                  TabIndex        =   112
                  Top             =   120
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00404080&
                  Caption         =   "€Ì— „— »ÿ »«·ﬁ”„"
                  ForeColor       =   &H8000000B&
                  Height          =   285
                  Index           =   23
                  Left            =   4635
                  TabIndex        =   111
                  Top             =   120
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H000000FF&
                  Caption         =   "·„ Ì”Ã· ›Ì «⁄œ«œ  «·‘› « "
                  ForeColor       =   &H8000000B&
                  Height          =   285
                  Index           =   22
                  Left            =   2460
                  TabIndex        =   110
                  Top             =   120
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FF8080&
                  Caption         =   "·„ Ì”Ã· Œ—ÊÕ"
                  ForeColor       =   &H80000007&
                  Height          =   285
                  Index           =   21
                  Left            =   1215
                  TabIndex        =   109
                  Top             =   120
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000C000&
                  Caption         =   "·Ì” ·Â «÷«›Ì"
                  ForeColor       =   &H80000007&
                  Height          =   285
                  Index           =   20
                  Left            =   120
                  TabIndex        =   108
                  Top             =   120
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "·Ì” ·Â ﬂÊœ «·„ﬂ‰…"
                  ForeColor       =   &H8000000B&
                  Height          =   285
                  Index           =   18
                  Left            =   7440
                  TabIndex        =   107
                  Top             =   120
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Caption         =   "·Ì” ·Â »’„…"
                  ForeColor       =   &H8000000B&
                  Height          =   285
                  Index           =   3
                  Left            =   6000
                  TabIndex        =   106
                  Top             =   120
                  Width           =   1230
               End
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „ÊŸ›Ì‰ ·„  ÿ«»ﬁ »’«„« Â„ «⁄œ«œ«  «·‘› « "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   405
               Index           =   26
               Left            =   3240
               TabIndex        =   114
               Top             =   120
               Width           =   7980
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
         Begin VB.Frame Frame2 
            Caption         =   "ÊÕœ… «·„›—œ"
            Enabled         =   0   'False
            Height          =   495
            Left            =   7440
            TabIndex        =   67
            Top             =   -360
            Visible         =   0   'False
            Width           =   7425
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1200
               TabIndex        =   75
               Text            =   "0"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox TxtValue1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   74
               Text            =   "0"
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "”«⁄« "
               Height          =   195
               Index           =   2
               Left            =   3960
               TabIndex        =   70
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ì«„"
               Height          =   195
               Index           =   1
               Left            =   5040
               TabIndex        =   69
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌ„…"
               Height          =   195
               Index           =   0
               Left            =   6120
               TabIndex        =   68
               Top             =   240
               Width           =   855
            End
            Begin VB.Label LBLWhereSTR 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               Height          =   255
               Left            =   360
               TabIndex        =   82
               Top             =   240
               Width           =   735
            End
            Begin VB.Label LBLhOURrATE 
               Alignment       =   2  'Center
               Caption         =   "1.5"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   600
               TabIndex        =   73
               Top             =   120
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„⁄œ· "
               Height          =   465
               Index           =   11
               Left            =   1440
               TabIndex        =   72
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LBLavg 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               Height          =   255
               Left            =   2880
               TabIndex        =   71
               Top             =   120
               Width           =   735
            End
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
            Left            =   8640
            TabIndex        =   44
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   315
            Left            =   5640
            TabIndex        =   61
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   315
            Left            =   3360
            TabIndex        =   63
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker TempDate 
            Height          =   315
            Left            =   3360
            TabIndex        =   65
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   3480
            TabIndex        =   66
            Top             =   120
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ReDta 
            Height          =   315
            Left            =   4440
            TabIndex        =   83
            Top             =   0
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker NoDate 
            Height          =   315
            Left            =   4440
            TabIndex        =   93
            Top             =   0
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3600
            TabIndex        =   97
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
            Format          =   95551489
            CurrentDate     =   38784
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   585
            Index           =   0
            Left            =   120
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   0
            Width           =   3375
            _cx             =   5953
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
            Caption         =   "  — Ì» ÿ»ﬁ«"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
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
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   100
               Top             =   240
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "··„ÊŸ›"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   101
               Top             =   240
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«· «—ÌŒ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·ﬂ·"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   0
               TabIndex        =   115
               Top             =   0
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
               Format          =   95551489
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   15
               Index           =   19
               Left            =   90
               TabIndex        =   99
               Top             =   555
               Width           =   3195
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   285
            Index           =   5
            Left            =   4680
            TabIndex        =   64
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   7440
            TabIndex        =   62
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   10200
            TabIndex        =   45
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbl 
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
Attribute VB_Name = "FrmApproveShift"
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
  Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.TblApproveShift.ID, dbo.TblApproveShift.RecordDate, dbo.TblApproveShift.FromDate, dbo.TblApproveShift.ToDate, dbo.TblApproveShift.DeptID, "
 MySQL = MySQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblApproveShift.ProjID, dbo.projects.Project_name,"
 MySQL = MySQL & "                      dbo.projects.Project_nameE, dbo.TblApproveShift.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblApproveShift.EmpID,"
 MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblApproveShift.SelectAll, dbo.TblApproveShift.SelectEmp,"
 MySQL = MySQL & "                      dbo.TblApproveShift.SelectBranch, dbo.TblApproveShift.SelectDept, dbo.TblApproveShift.SelectProject, dbo.TblApproveShiftDet.ApprovID,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.RecordDate AS DetRecordDate, dbo.TblApproveShiftDet.ShiftID, dbo.TblApproveShiftDet.EnterTime, dbo.TblApproveShiftDet.OutTime,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.FromTime, dbo.TblApproveShiftDet.ToTime, dbo.TblApproveShiftDet.DelayID, dbo.TblApproveShiftDet.Absence,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.EarExit, dbo.TblApproveShiftDet.Additional, dbo.TblApproveShiftDet.NoRegOut, dbo.TblApproveShiftDet.MachinCode,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.DelayType, dbo.TblApproveShiftDet.AbsenType, dbo.TblApproveShiftDet.EarExitType, dbo.TblApproveShiftDet.AddiType,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.NoRegOutType, dbo.TblApproveShiftDet.AbsenceTime, dbo.TblApproveShiftDet.EarExitTime, dbo.TblApproveShiftDet.DelayTimeTime,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.AdditioTime, dbo.TblApproveShiftDet.TypeDay, dbo.TblApproveShiftDet.IDImport, dbo.TblApproveShiftDet.AbsenceTimeVal,"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet.NotFingPrin, dbo.TblApproveShiftDet.selected, dbo.TblApproveShiftDet.EmpID AS DetEmpID, TblEmployee_1.Emp_Name AS DetEmp_Name,"
 MySQL = MySQL & "                      TblEmployee_1.Fullcode AS DetFullcode, TblEmployee_1.Emp_Namee AS DetEmp_NameE, dbo.TblApproveShiftDet.ProjID AS DetProjID,"
 MySQL = MySQL & "                      projects_1.Project_name AS DetProject_name, projects_1.Project_nameE AS DetProject_nameE, dbo.TblApproveShiftDet.BranchID AS DetBranchID,"
 MySQL = MySQL & "                      TblBranchesData_1.branch_name AS Detbranch_name, TblBranchesData_1.branch_namee AS Detbranch_nameE, dbo.TblApproveShiftDet.DeptID AS DetDeptID,"
 MySQL = MySQL & "                      TblEmpDepartments_1.DepartmentName AS DetDepartmentName, TblEmpDepartments_1.DepartmentNamee AS DetDepartmentNameE ,dbo.TblApproveShiftDet.TypeTrans"
 MySQL = MySQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblEmpDepartments TblEmpDepartments_1 RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblApproveShiftDet ON TblEmpDepartments_1.DeparmentID = dbo.TblApproveShiftDet.DeptID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBranchesData TblBranchesData_1 ON dbo.TblApproveShiftDet.BranchID = TblBranchesData_1.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.projects projects_1 ON dbo.TblApproveShiftDet.ProjID = projects_1.id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblApproveShiftDet.EmpID = TblEmployee_1.Emp_ID RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblApproveShift ON dbo.TblApproveShiftDet.ApprovID = dbo.TblApproveShift.ID ON dbo.TblEmployee.Emp_ID = dbo.TblApproveShift.EmpID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblApproveShift.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.projects ON dbo.TblApproveShift.ProjID = dbo.projects.id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.TblApproveShift.DeptID = dbo.TblEmpDepartments.DeparmentID"
 MySQL = MySQL & " Where (dbo.TblApproveShift.ID = " & val(TxtSerial1.Text) & ") and (dbo.TblApproveShiftDet.TypeTrans IS NULL)"
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_ApproveShift.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_ApproveShift.rpt"
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
 Function CheckHolidaies(Optional RecDate As Date) As Integer
 Dim sql As String
 Dim Rs7 As ADODB.Recordset
 Set Rs7 = New ADODB.Recordset
 sql = "Select * from dbo.TblVacationschedule22 where ISVac=1 and Date =" & SQLDate(RecDate, True) & ""
 Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 CheckHolidaies = 1
 Else
 CheckHolidaies = 0
 End If
 End Function

Function CheckShiftHolidaies(Optional EmpID As Double, Optional NoDay As Integer) As Integer
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TblShiftWorker.EmpID AS EmpID1, dbo.TbLSheft.*"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where dbo.TblShiftWorker.EmpID=" & EmpID & ""
Select Case NoDay
Case 7
sql = sql & " and  SatWoVo=1"
Case 6
sql = sql & " and  FrirWoVo=1"
Case 5
sql = sql & " and  ThurWoVo=1"
Case 4
sql = sql & " and WedWoVo=1"
Case 3
sql = sql & " and TuesWoVo=1"
Case 2
sql = sql & " and MonWoVo=1"
Case 1
sql = sql & " and  SunWoVo=1"
End Select
Rs8.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckShiftHolidaies = 1
Else
CheckShiftHolidaies = 0
End If
End Function
Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
End Select
End If
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

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    With GridInstallments
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("DelayType")) = "#1;ÌŒ’„ „‰ «·—« »  |#2;«–‰ "
            .ColComboList(.ColIndex("AbsenType")) = "#1;ÌŒ’„ „‰ «·—« » |#2;ÌŒ’„ „‰ „” Õﬁ« Â |#3;«–‰ "
            .ColComboList(.ColIndex("EarExitType")) = "#1;ÌŒ’„ „‰ «·—« »  |#2;«–‰ "
            .ColComboList(.ColIndex("AddiType")) = "#1;Ì÷«› ··—« » |#2;·«Ì÷«› "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("DelayType")) = "#1;Discount From Salary |#2;Permission"
           .ColComboList(.ColIndex("AbsenType")) = "#1;Discount From Salary |#2;Discount From Entitlements |#3;Permission"
           .ColComboList(.ColIndex("EarExitType")) = "#1;Discount From Salary |#2;Permission"
           .ColComboList(.ColIndex("AddiType")) = "#1;Addition To Salary |#2;No Addition"
            End If
    End With
        With VSFlexGrid1
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("DelayType")) = "#1;ÌŒ’„ „‰ «·—« »  |#2;«–‰ "
            .ColComboList(.ColIndex("AbsenType")) = "#1;ÌŒ’„ „‰ «·—« » |#2;ÌŒ’„ „‰ „” Õﬁ« Â |#3;«–‰ "
            .ColComboList(.ColIndex("EarExitType")) = "#1;ÌŒ’„ „‰ «·—« »  |#2;«–‰ "
            .ColComboList(.ColIndex("AddiType")) = "#1;Ì÷«› ··—« » |#2;·«Ì÷«› "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("DelayType")) = "#1;Discount From Salary |#2;Permission"
           .ColComboList(.ColIndex("AbsenType")) = "#1;Discount From Salary |#2;Discount From Entitlements |#3;Permission"
           .ColComboList(.ColIndex("EarExitType")) = "#1;Discount From Salary |#2;Permission"
           .ColComboList(.ColIndex("AddiType")) = "#1;Addition To Salary |#2;No Addition"
            End If
    End With
    
    YearMonth
    conection = "select * from  TblApproveShift  order by  ID "
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
Public Sub FiLLRec2()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double

                  StrSQL = "Delete From TblApproveShiftDet Where ApprovID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
      StrSQL = "Delete From TblChangedComponentRegister Where ApproveShiftID=" & val(Me.TxtSerial1.Text) & " and Finger =1"
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
          StrSQL = "Delete From TblInforVacatiom Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
   RsSavRec.Fields("FromDate").value = Fromdate.value
   RsSavRec.Fields("ToDate").value = ToDate.value
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
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblApproveShiftDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
       RsDevsub.AddNew
       RsDevsub("NotFingPrin").value = IIf((.TextMatrix(i, .ColIndex("NotFingPrin"))) = "", Null, val(.TextMatrix(i, .ColIndex("NotFingPrin"))))
                RsDevsub("ApprovID").value = val(Me.TxtSerial1.Text)
                RsDevsub("selected").value = 1
                RsDevsub("sortdate").value = IIf((.TextMatrix(i, .ColIndex("sortdate"))) = "", Null, (.TextMatrix(i, .ColIndex("sortdate"))))
                RsDevsub("IDImport2").value = IIf((.TextMatrix(i, .ColIndex("IDImport2"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDImport2"))))
                RsDevsub("IDImport").value = IIf((.TextMatrix(i, .ColIndex("IDImport"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDImport"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                RsDevsub("MachinCode").value = IIf((.TextMatrix(i, .ColIndex("MachinCode"))) = "", Null, Trim(.TextMatrix(i, .ColIndex("MachinCode"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("ProjID").value = IIf((.TextMatrix(i, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjID"))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("ShiftID").value = IIf((.TextMatrix(i, .ColIndex("ShiftID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ShiftID"))))
                RsDevsub("EnterTime").value = IIf((.TextMatrix(i, .ColIndex("EnterTime"))) = "", Null, (.TextMatrix(i, .ColIndex("EnterTime"))))
                RsDevsub("OutTime").value = IIf((.TextMatrix(i, .ColIndex("OutTime"))) = "", Null, (.TextMatrix(i, .ColIndex("OutTime"))))
                RsDevsub("FromTime").value = IIf((.TextMatrix(i, .ColIndex("FromTime"))) = "", Null, (.TextMatrix(i, .ColIndex("FromTime"))))
                RsDevsub("ToTime").value = IIf((.TextMatrix(i, .ColIndex("ToTime"))) = "", Null, (.TextMatrix(i, .ColIndex("ToTime"))))
                RsDevsub("DelayID").value = IIf((.TextMatrix(i, .ColIndex("DelayID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DelayID"))))
                RsDevsub("Absence").value = IIf((.TextMatrix(i, .ColIndex("Absence"))) = "", Null, val(.TextMatrix(i, .ColIndex("Absence"))))
                RsDevsub("EarExit").value = IIf((.TextMatrix(i, .ColIndex("EarExit"))) = "", Null, val(.TextMatrix(i, .ColIndex("EarExit"))))
                RsDevsub("Additional").value = IIf((.TextMatrix(i, .ColIndex("Additional"))) = "", Null, val(.TextMatrix(i, .ColIndex("Additional"))))
                RsDevsub("TypeDay").value = IIf((.TextMatrix(i, .ColIndex("TypeDay"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeDay"))))
                RsDevsub("DelayType").value = IIf((.TextMatrix(i, .ColIndex("DelayType"))) = "", Null, val(.TextMatrix(i, .ColIndex("DelayType"))))
                RsDevsub("AbsenType").value = IIf((.TextMatrix(i, .ColIndex("AbsenType"))) = "", Null, val(.TextMatrix(i, .ColIndex("AbsenType"))))
                RsDevsub("EarExitType").value = IIf((.TextMatrix(i, .ColIndex("EarExitType"))) = "", Null, val(.TextMatrix(i, .ColIndex("EarExitType"))))
                RsDevsub("AddiType").value = IIf((.TextMatrix(i, .ColIndex("AddiType"))) = "", Null, val(.TextMatrix(i, .ColIndex("AddiType"))))
                RsDevsub("NoRegOutType").value = IIf((.TextMatrix(i, .ColIndex("NoRegOutType"))) = "", Null, val(.TextMatrix(i, .ColIndex("NoRegOutType"))))
                RsDevsub("AbsenceTime").value = IIf((.TextMatrix(i, .ColIndex("AbsenceTime"))) = "", Null, (.TextMatrix(i, .ColIndex("AbsenceTime"))))
                RsDevsub("AbsenceTimeVal").value = IIf((.TextMatrix(i, .ColIndex("AbsenceTimeVal"))) = "", Null, val(.TextMatrix(i, .ColIndex("AbsenceTimeVal"))))
                RsDevsub("EarExitTime").value = IIf((.TextMatrix(i, .ColIndex("EarExitTime"))) = "", Null, (.TextMatrix(i, .ColIndex("EarExitTime"))))
                RsDevsub("DelayTimeTime").value = IIf((.TextMatrix(i, .ColIndex("DelayTimeTime"))) = "", Null, (.TextMatrix(i, .ColIndex("DelayTimeTime"))))
                RsDevsub("AdditioTime").value = IIf((.TextMatrix(i, .ColIndex("AdditioTime"))) = "", Null, (.TextMatrix(i, .ColIndex("AdditioTime"))))
                If .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexChecked Then
                RsDevsub("NoRegOut").value = 1
                Else
                RsDevsub("NoRegOut").value = 0
                End If
                Cn.Execute "update TblImportShiftsDet set IsSele=1 where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
                If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=1 where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
                End If
           
       RsDevsub.update
        Else
            Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
               If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
                End If
            End If

     Next i
    End With
  ''//////////////////
        Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblApproveShiftDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       RsDevsub.AddNew
       RsDevsub("NotFingPrin").value = IIf((.TextMatrix(i, .ColIndex("NotFingPrin"))) = "", Null, val(.TextMatrix(i, .ColIndex("NotFingPrin"))))
                RsDevsub("ApprovID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TypeTrans").value = 1
                RsDevsub("sortdate").value = IIf((.TextMatrix(i, .ColIndex("sortdate"))) = "", Null, (.TextMatrix(i, .ColIndex("sortdate"))))
                RsDevsub("IDImport2").value = IIf((.TextMatrix(i, .ColIndex("IDImport2"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDImport2"))))
                RsDevsub("IDImport").value = IIf((.TextMatrix(i, .ColIndex("IDImport"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDImport"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                RsDevsub("MachinCode").value = IIf((.TextMatrix(i, .ColIndex("MachinCode"))) = "", Null, Trim(.TextMatrix(i, .ColIndex("MachinCode"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("ProjID").value = IIf((.TextMatrix(i, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjID"))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("ShiftID").value = IIf((.TextMatrix(i, .ColIndex("ShiftID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ShiftID"))))
                RsDevsub("EnterTime").value = IIf((.TextMatrix(i, .ColIndex("EnterTime"))) = "", Null, (.TextMatrix(i, .ColIndex("EnterTime"))))
                RsDevsub("OutTime").value = IIf((.TextMatrix(i, .ColIndex("OutTime"))) = "", Null, (.TextMatrix(i, .ColIndex("OutTime"))))
                RsDevsub("FromTime").value = IIf((.TextMatrix(i, .ColIndex("FromTime"))) = "", Null, (.TextMatrix(i, .ColIndex("FromTime"))))
                RsDevsub("ToTime").value = IIf((.TextMatrix(i, .ColIndex("ToTime"))) = "", Null, (.TextMatrix(i, .ColIndex("ToTime"))))
                RsDevsub("DelayID").value = IIf((.TextMatrix(i, .ColIndex("DelayID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DelayID"))))
                RsDevsub("Absence").value = IIf((.TextMatrix(i, .ColIndex("Absence"))) = "", Null, val(.TextMatrix(i, .ColIndex("Absence"))))
                RsDevsub("EarExit").value = IIf((.TextMatrix(i, .ColIndex("EarExit"))) = "", Null, val(.TextMatrix(i, .ColIndex("EarExit"))))
                RsDevsub("Additional").value = IIf((.TextMatrix(i, .ColIndex("Additional"))) = "", Null, val(.TextMatrix(i, .ColIndex("Additional"))))
                RsDevsub("TypeDay").value = IIf((.TextMatrix(i, .ColIndex("TypeDay"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeDay"))))
                RsDevsub("DelayType").value = IIf((.TextMatrix(i, .ColIndex("DelayType"))) = "", Null, val(.TextMatrix(i, .ColIndex("DelayType"))))
                RsDevsub("AbsenType").value = IIf((.TextMatrix(i, .ColIndex("AbsenType"))) = "", Null, val(.TextMatrix(i, .ColIndex("AbsenType"))))
                RsDevsub("EarExitType").value = IIf((.TextMatrix(i, .ColIndex("EarExitType"))) = "", Null, val(.TextMatrix(i, .ColIndex("EarExitType"))))
                RsDevsub("AddiType").value = IIf((.TextMatrix(i, .ColIndex("AddiType"))) = "", Null, val(.TextMatrix(i, .ColIndex("AddiType"))))
                RsDevsub("NoRegOutType").value = IIf((.TextMatrix(i, .ColIndex("NoRegOutType"))) = "", Null, val(.TextMatrix(i, .ColIndex("NoRegOutType"))))
                RsDevsub("AbsenceTime").value = IIf((.TextMatrix(i, .ColIndex("AbsenceTime"))) = "", Null, (.TextMatrix(i, .ColIndex("AbsenceTime"))))
                RsDevsub("AbsenceTimeVal").value = IIf((.TextMatrix(i, .ColIndex("AbsenceTimeVal"))) = "", Null, val(.TextMatrix(i, .ColIndex("AbsenceTimeVal"))))
                RsDevsub("EarExitTime").value = IIf((.TextMatrix(i, .ColIndex("EarExitTime"))) = "", Null, (.TextMatrix(i, .ColIndex("EarExitTime"))))
                RsDevsub("DelayTimeTime").value = IIf((.TextMatrix(i, .ColIndex("DelayTimeTime"))) = "", Null, (.TextMatrix(i, .ColIndex("DelayTimeTime"))))
                RsDevsub("AdditioTime").value = IIf((.TextMatrix(i, .ColIndex("AdditioTime"))) = "", Null, (.TextMatrix(i, .ColIndex("AdditioTime"))))
                If .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexChecked Then
                RsDevsub("NoRegOut").value = 1
                Else
                RsDevsub("NoRegOut").value = 0
                End If
                Cn.Execute "update TblImportShiftsDet set IsSele=1 where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
                Cn.Execute "update TblImportShiftsDet set IsSele=1 where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
       RsDevsub.update
  
         
     Next i
    End With
    MofrdComponan
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
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             
                  StrSQL = "Delete From TblApproveShiftDet Where ApprovID=" & val(Me.TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
      StrSQL = "Delete From TblChangedComponentRegister Where ApproveShiftID=" & val(Me.TxtSerial1.Text) & " and Finger =1"
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
          StrSQL = "Delete From TblInforVacatiom Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
   RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
   RsSavRec.Fields("FromDate").value = Fromdate.value
   RsSavRec.Fields("ToDate").value = ToDate.value
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
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblApproveShiftDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If .TextMatrix(i, .ColIndex("Emp_Name")) <> " " Then
       RsDevsub.AddNew
       If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
       RsDevsub("selected").value = 1
       Else
       RsDevsub("selected").value = 0
       End If
                RsDevsub("ApprovID").value = val(Me.TxtSerial1.Text)
                RsDevsub("IDImport2").value = IIf((.TextMatrix(i, .ColIndex("IDImport2"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDImport2"))))
                RsDevsub("IDImport").value = IIf((.TextMatrix(i, .ColIndex("IDImport"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDImport"))))
                RsDevsub("NotFingPrin").value = IIf((.TextMatrix(i, .ColIndex("NotFingPrin"))) = "", Null, val(.TextMatrix(i, .ColIndex("NotFingPrin"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                RsDevsub("sortdate").value = IIf((.TextMatrix(i, .ColIndex("sortdate"))) = "", Null, (.TextMatrix(i, .ColIndex("sortdate"))))
                RsDevsub("MachinCode").value = IIf((.TextMatrix(i, .ColIndex("MachinCode"))) = "", Null, Trim(.TextMatrix(i, .ColIndex("MachinCode"))))
                RsDevsub("DeptID").value = IIf((.TextMatrix(i, .ColIndex("DeptID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DeptID"))))
                RsDevsub("ProjID").value = IIf((.TextMatrix(i, .ColIndex("ProjID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjID"))))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("ShiftID").value = IIf((.TextMatrix(i, .ColIndex("ShiftID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ShiftID"))))
                RsDevsub("EnterTime").value = IIf((.TextMatrix(i, .ColIndex("EnterTime"))) = "", Null, (.TextMatrix(i, .ColIndex("EnterTime"))))
                RsDevsub("OutTime").value = IIf((.TextMatrix(i, .ColIndex("OutTime"))) = "", Null, (.TextMatrix(i, .ColIndex("OutTime"))))
                RsDevsub("FromTime").value = IIf((.TextMatrix(i, .ColIndex("FromTime"))) = "", Null, (.TextMatrix(i, .ColIndex("FromTime"))))
                RsDevsub("ToTime").value = IIf((.TextMatrix(i, .ColIndex("ToTime"))) = "", Null, (.TextMatrix(i, .ColIndex("ToTime"))))
                RsDevsub("DelayID").value = IIf((.TextMatrix(i, .ColIndex("DelayID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DelayID"))))
                RsDevsub("Absence").value = IIf((.TextMatrix(i, .ColIndex("Absence"))) = "", Null, val(.TextMatrix(i, .ColIndex("Absence"))))
                RsDevsub("EarExit").value = IIf((.TextMatrix(i, .ColIndex("EarExit"))) = "", Null, val(.TextMatrix(i, .ColIndex("EarExit"))))
                RsDevsub("Additional").value = IIf((.TextMatrix(i, .ColIndex("Additional"))) = "", Null, val(.TextMatrix(i, .ColIndex("Additional"))))
                RsDevsub("TypeDay").value = IIf((.TextMatrix(i, .ColIndex("TypeDay"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeDay"))))
                RsDevsub("DelayType").value = IIf((.TextMatrix(i, .ColIndex("DelayType"))) = "", Null, val(.TextMatrix(i, .ColIndex("DelayType"))))
                RsDevsub("AbsenType").value = IIf((.TextMatrix(i, .ColIndex("AbsenType"))) = "", Null, val(.TextMatrix(i, .ColIndex("AbsenType"))))
                RsDevsub("EarExitType").value = IIf((.TextMatrix(i, .ColIndex("EarExitType"))) = "", Null, val(.TextMatrix(i, .ColIndex("EarExitType"))))
                RsDevsub("AddiType").value = IIf((.TextMatrix(i, .ColIndex("AddiType"))) = "", Null, val(.TextMatrix(i, .ColIndex("AddiType"))))
                RsDevsub("NoRegOutType").value = IIf((.TextMatrix(i, .ColIndex("NoRegOutType"))) = "", Null, val(.TextMatrix(i, .ColIndex("NoRegOutType"))))
                RsDevsub("AbsenceTime").value = IIf((.TextMatrix(i, .ColIndex("AbsenceTime"))) = "", Null, (.TextMatrix(i, .ColIndex("AbsenceTime"))))
                RsDevsub("AbsenceTimeVal").value = IIf((.TextMatrix(i, .ColIndex("AbsenceTimeVal"))) = "", Null, val(.TextMatrix(i, .ColIndex("AbsenceTimeVal"))))
                RsDevsub("EarExitTime").value = IIf((.TextMatrix(i, .ColIndex("EarExitTime"))) = "", Null, (.TextMatrix(i, .ColIndex("EarExitTime"))))
                RsDevsub("DelayTimeTime").value = IIf((.TextMatrix(i, .ColIndex("DelayTimeTime"))) = "", Null, (.TextMatrix(i, .ColIndex("DelayTimeTime"))))
                RsDevsub("AdditioTime").value = IIf((.TextMatrix(i, .ColIndex("AdditioTime"))) = "", Null, (.TextMatrix(i, .ColIndex("AdditioTime"))))
                If .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexChecked Then
                RsDevsub("NoRegOut").value = 1
                Else
                RsDevsub("NoRegOut").value = 0
                End If
                Cn.Execute "update TblImportShiftsDet set IsSele=1 where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
              If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=1 where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
                End If
       RsDevsub.update
        Else
           If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
                End If
            Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
            End If

     Next i
    End With
  CalcalteAbsence
  FiLLTXT
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
    Fromdate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value)
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcpDept1.BoundText = IIf(IsNull(RsSavRec.Fields("DeptID").value), "", RsSavRec.Fields("DeptID").value)
    Me.DcbProject1.BoundText = IIf(IsNull(RsSavRec.Fields("ProjID").value), "", RsSavRec.Fields("ProjID").value)
    Me.DcbBranch1.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DcbEmployee1.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
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
    Me.RdEmp.value = True
    Else
    RdEmp.value = False
    End If
    Else
    RdEmp.value = False
    End If
    If Not (IsNull(RsSavRec.Fields("SelectBranch").value)) Then
    If RsSavRec.Fields("SelectBranch").value = 1 Then
    Me.SelectBranch.value = vbChecked
    Else
    RdEmp.value = vbUnchecked
    End If
    Else
    RdEmp.value = vbUnchecked
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

     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
If Me.TxtModFlg.Text = "R" Then
FullGridData2
End If
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
    Dim i As Integer
   With GridInstallments
   For i = 1 To .Rows - 1
   If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
   If val(.TextMatrix(i, .ColIndex("EmpID"))) = 0 Or val(.TextMatrix(i, .ColIndex("ShiftID"))) = 0 Or val(.TextMatrix(i, .ColIndex("DeptID"))) = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ  «ﬂœ „‰ œ·«·«  «·«·Ê«‰"
   Else
   MsgBox "Can not save. Make sure the data"
   End If
   Exit Sub
   End If
   End If
   Next i
   End With
   If Me.TxtModFlg.Text = "E" Then
ISButton4_Click
End If
            '----------------------------- save edit -------------------------------
            FiLLRec2
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
    StrRecID = new_id("TblApproveShift", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Function ChecPeriodSalary(Optional TemDate As Date, Optional ByRef MonthID As Integer, Optional ByRef YearID As Integer) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblDurations2Salary where  FromDate<=" & SQLDate(TemDate, True) & ""
sql = sql & " and ToDate>=" & SQLDate(TemDate, True) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ChecPeriodSalary = True
MonthID = IIf(IsNull(Rs3("MonthID").value), 0, Rs3("MonthID").value)
YearID = IIf(IsNull(Rs3("YearID").value), 0, Rs3("YearID").value)
Else
ChecPeriodSalary = False
End If
End Function
Private Sub SaveData(Optional Emp_id As Integer, Optional BranchID As Integer, Optional MofrdID As Integer, Optional NoofDays As Double, Optional NoOfHours As Double, Optional RecDate As Date)
    Dim Msg As String
    Dim BasicSalary As Double
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim HourRate As Double
    Dim value As Double
    Dim Rs3 As ADODB.Recordset
    Dim SmuHourofShift As Double
    Dim NetHour As Double
    Dim SmuHourofVoca As Double
    Dim NoHourWork As Double
    Dim MonthID As Integer
    Dim YearID As Integer
    '-------------------------------------------------------------------------------------------
    ReDta.value = RecDate
    If ChecPeriodSalary(ReDta.value, MonthID, YearID) = True Then
        CboYear.Text = YearID
    CmbMonth.ListIndex = MonthID - 1
    Else
        CboYear.Text = year(ReDta.value)
    CmbMonth.ListIndex = Month(ReDta.value) - 1
    End If
    Dim EmployeeSalary As Double
    Dim NoDayMonth As Integer
  '  Dim NoOfHours As Double
    Dim i As Long

                If Opt(1).value = True Then
                    EmployeeSalary = GetEmployeeSalaryAccordingToComponent(Emp_id, LBLWhereSTR)
                    '«Ì«„
                    If SystemOptions.MonthIs30days = True Then
                    
                        HourRate = (EmployeeSalary / 30)
                    Else
                        HourRate = (EmployeeSalary * 12 / 365)
                    End If

If val(LBLavg.Caption) > 0 Then
HourRate = val(LBLavg.Caption)
LBLhOURrATE.Caption = 1
End If
                  ' value = Round(HourRate * val(LBLhOURrATE) * NoofDays, SystemOptions.EmpComponentDigts)
                    value = Round(HourRate * NoofDays, SystemOptions.EmpComponentDigts)
                     BasicSalary = EmployeeSalary
                    NoOfHours = 0
                ElseIf Opt(2).value = True Then '”«⁄« 
               ' If SystemOptions.MonthIs30days = True Then
                NoDayMonth = 30
               ' Else
               ' NoDayMonth = Day(DateSerial(year(RecDate), Month(RecDate) + 1, 1) - 1)
               ' End If
               If NoHourInShift(NoHourWork) = True Then
               SmuHourofShift = NoHourWork
                If SmuHourofShift = 0 Then
                 SmuHourofShift = 8
                 End If
               Else
                 SmuHourofShift = SumHour(Emp_id, Weekday(RecDate))
                 If SmuHourofShift = 0 Then
                 SmuHourofShift = 8
                 End If
               End If
               '  SmuHourofVoca = SumDayVaction(Emp_id, Weekday(RecDate))
               '  If SystemOptions.AllWorkdays = False Then
               '  NetHour = SmuHourofShift * NoDayMonth - SmuHourofVoca
               '  Else
                 NetHour = SmuHourofShift * NoDayMonth
               '  End If
                    EmployeeSalary = GetEmployeeSalaryAccordingToComponent(Emp_id, LBLWhereSTR)

                    If GetNoOfHourPerMonth > 0 Then
                       HourRate = (EmployeeSalary / NetHour)
                        
                    Else
                        HourRate = 0
                        
                       HourRate = (EmployeeSalary / NetHour)
                    End If
                    NoOfHours = NoOfHours / 60
                  
If val(LBLavg.Caption) > 0 Then
NoOfHours = val(LBLavg.Caption)
LBLhOURrATE.Caption = 1
End If

                   ' value = Round((NoOfHours) * HourRate * val(LBLhOURrATE), SystemOptions.EmpComponentDigts)
    value = Round((NoOfHours) * HourRate, SystemOptions.EmpComponentDigts)
                    NoofDays = 0
                    
                    BasicSalary = EmployeeSalary
                    
                End If

    Set Rs3 = New ADODB.Recordset
    StrSQL = "select * from TblChangedComponentRegister where 1=-1"
    Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        Me.TxtId.Text = CStr(new_id("TblChangedComponentRegister", "ChangedComponentid", "", True))
        Rs3.AddNew
        Rs3("ChangedComponentid").value = val(Me.TxtId.Text)

    Rs3("RecordDate").value = ReDta.value
    Rs3("year").value = val(CboYear.ListIndex)
    Rs3("month").value = CmbMonth.ListIndex
    Rs3("Actualyear").value = val(CboYear.Text)
    Rs3("Actualmonth").value = val(CmbMonth.ListIndex) + 1
    Rs3("ComponentID").value = MofrdID
    Rs3("BranchId").value = BranchID
    Rs3("Finger").value = 1
    Rs3("ApproveShiftID").value = val(TxtSerial1.Text)
    Rs3.update
 Dim xx As Integer
    Set RsDev = New ADODB.Recordset
'If ChekMofrdAbcence(MofrdID) = True Then
'xx = 1
'Else
'xx = 0
'End If
    StrSQL = " SELECT     * FROM         dbo.TblChangedComponentRegisterDetails WHERE     (Emp_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Emp_id <> 0 Then
                RsDev.AddNew
                RsDev("ApproveShiftID").value = val(TxtSerial1.Text)
                RsDev("ChangedComponentid").value = val(TxtId.Text)
                RsDev("Emp_ID").value = Emp_id
                RsDev("NoofDays").value = NoofDays
               ' RsDev("NoOfMinutes").Value = NoOfMinutes
                RsDev("HourRate").value = HourRate
                RsDev("NoOfHour").value = NoOfHours
                RsDev("Salary").value = BasicSalary
                RsDev("Value").value = value
             '   If xx = 1 And Opt(1).value = True Then
              '  SaveInformationVacation (val(CboYear.ListIndex) + 2006), val(CmbMonth.ListIndex) + 1, Emp_id, Noofdays
               ' End If
                RsDev.update
            End If
End Sub
Function ChekMofrdAbcence(Optional MofrID As Integer) As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     MofrdAbcen, id"
sql = sql & " From dbo.MOFRAD"
sql = sql & " Where (ID =" & MofrID & ") And (MofrdAbcen = 1)"
Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekMofrdAbcence = True
Else
ChekMofrdAbcence = False
End If

End Function
Sub SaveInformationVacation(Optional Yearr As Integer, Optional Monthh As Integer, Optional EmpID As Integer = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
str = "  «·„›—œ«  «·„ €Ì—…„‰ «⁄ „«œ «·»’„…"
Else
str = "Components Changing From Approve Shift"
End If
sql = "select * from TblInforVacatiom where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("ApproveShiftID").value = val(Me.TxtSerial1.Text)
      Rs7("AbcenceID").value = val(Me.TxtId.Text)
      Rs7("EmpID").value = EmpID
      Rs7("NoDay").value = (NoDay)
      Rs7("RecordDate").value = XPDtbTrans.value
      Rs7("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
      Rs7("TypeVacation").value = 1
      Rs7("Remarks").value = str
      Rs7("Yearr").value = Yearr
      Rs7("Monthh").value = Monthh
      Rs7.update
End Sub
Sub MofrdPr(Optional EmpID As Integer, Optional BranchID As Integer, Optional MofrdID As Integer, Optional NoofDay As Double, Optional NoofHour As Double, Optional RecDate As Date)
If MofrdID <> 0 Then
Dim Equation As Double
Dim avg As Double
Dim componentUnit As Integer
    componentUnit = GetMofradUnit(MofrdID, avg)
LBLavg.Caption = avg
    Opt(componentUnit).value = True
   ' ChangeGridView componentUnit
    LBLWhereSTR.Caption = GetSpecificComponentIncalculations(MofrdID, Equation)
LBLhOURrATE.Caption = Equation
If Opt(1).value = True Or Opt(2).value = True Then
SaveData EmpID, BranchID, MofrdID, NoofDay, NoofHour, RecDate
End If
End If
End Sub
Function MofrdComponan()
Dim i As Integer
With GridInstallments
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 And val(.TextMatrix(i, .ColIndex("DeptID"))) <> 0 Then
If val(.TextMatrix(i, .ColIndex("DelayID"))) <> 0 And val(.TextMatrix(i, .ColIndex("DelayType"))) = 1 Then
MofrdPr val(.TextMatrix(i, .ColIndex("EmpID"))), val(.TextMatrix(i, .ColIndex("BranchID"))), GetMofrad(val(.TextMatrix(i, .ColIndex("DeptID"))), 0), 0, val(.TextMatrix(i, .ColIndex("DelayID"))), .TextMatrix(i, .ColIndex("RecordDate"))
End If
If val(.TextMatrix(i, .ColIndex("EarExit"))) <> 0 And val(.TextMatrix(i, .ColIndex("EarExitType"))) = 1 Then
MofrdPr val(.TextMatrix(i, .ColIndex("EmpID"))), val(.TextMatrix(i, .ColIndex("BranchID"))), GetMofrad(val(.TextMatrix(i, .ColIndex("DeptID"))), 1), 0, val(.TextMatrix(i, .ColIndex("EarExit"))), .TextMatrix(i, .ColIndex("RecordDate"))
End If
If val(.TextMatrix(i, .ColIndex("Absence"))) <> 0 And val(.TextMatrix(i, .ColIndex("AbsenType"))) = 1 Then
MofrdPr val(.TextMatrix(i, .ColIndex("EmpID"))), val(.TextMatrix(i, .ColIndex("BranchID"))), GetMofrad(val(.TextMatrix(i, .ColIndex("DeptID"))), 2), val(.TextMatrix(i, .ColIndex("Absence"))), , .TextMatrix(i, .ColIndex("RecordDate"))
End If
If val(.TextMatrix(i, .ColIndex("Additional"))) <> 0 And val(.TextMatrix(i, .ColIndex("AddiType"))) = 1 Then
MofrdPr val(.TextMatrix(i, .ColIndex("EmpID"))), val(.TextMatrix(i, .ColIndex("BranchID"))), GetMofrad(val(.TextMatrix(i, .ColIndex("DeptID"))), 3), 0, val(.TextMatrix(i, .ColIndex("Additional"))), .TextMatrix(i, .ColIndex("RecordDate"))
End If
End If
End If
Next i
End With
End Function
Function GetMofrad(Optional DeparmentID As Integer, Optional typeid As Integer = 0) As Integer
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = "Select * from TblEmpDepartments where DeparmentID=" & DeparmentID & ""
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
If typeid = 0 Then
GetMofrad = IIf(IsNull(Rs6("DelayID").value), 0, Rs6("DelayID").value)
ElseIf typeid = 1 Then
GetMofrad = IIf(IsNull(Rs6("EarlyexitID").value), 0, Rs6("EarlyexitID").value)
ElseIf typeid = 2 Then
GetMofrad = IIf(IsNull(Rs6("AbscenID").value), 0, Rs6("AbscenID").value)
ElseIf typeid = 3 Then
GetMofrad = IIf(IsNull(Rs6("AddID").value), 0, Rs6("AddID").value)
End If
Else
GetMofrad = 0
End If
End Function
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
sql = " SELECT  TblEmployee.fullcode,   dbo.TblApproveShiftDet.ID, dbo.TblApproveShiftDet.RecordDate, dbo.TblApproveShiftDet.ApprovID, dbo.TblApproveShiftDet.EnterTime, "
sql = sql & "                      dbo.TblApproveShiftDet.OutTime, dbo.TblApproveShiftDet.FromTime, dbo.TblApproveShiftDet.ToTime, dbo.TblApproveShiftDet.Absence,"
sql = sql & "                      dbo.TblApproveShiftDet.EarExit, dbo.TblApproveShiftDet.Additional, dbo.TblApproveShiftDet.NoRegOut, dbo.TblApproveShiftDet.MachinCode,"
sql = sql & "                      dbo.TblApproveShiftDet.DelayType, dbo.TblApproveShiftDet.AbsenType, dbo.TblApproveShiftDet.EarExitType, dbo.TblApproveShiftDet.AddiType,"
sql = sql & "                      dbo.TblApproveShiftDet.NoRegOutType, dbo.TblApproveShiftDet.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
sql = sql & "                      dbo.TblApproveShiftDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblApproveShiftDet.ProjID,"
sql = sql & "                      dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblApproveShiftDet.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                      dbo.TblApproveShiftDet.ShiftID , dbo.TbLSheft.SheftName, dbo.TbLSheft.SheftNamee, dbo.TblApproveShiftDet.TypeDay,"
sql = sql & "                      dbo.TblApproveShiftDet.EarExitTime,dbo.TblApproveShiftDet.DelayTimeTime, dbo.TblApproveShiftDet.AdditioTime, dbo.TblApproveShiftDet.AbsenceTime ,dbo.TblApproveShiftDet.IDImport ,"
sql = sql & " dbo.TblApproveShiftDet.DelayID ,dbo.TblApproveShiftDet.AbsenceTimeVal,dbo.TblApproveShiftDet.NotFingPrin ,dbo.TblApproveShiftDet.selected ,dbo.TblApproveShiftDet.IDImport2  ,dbo.TblApproveShiftDet.sortdate ,dbo.TblApproveShiftDet.TypeTrans"
sql = sql & " FROM         dbo.TblApproveShiftDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TbLSheft ON dbo.TblApproveShiftDet.ShiftID = dbo.TbLSheft.SeftCode LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblApproveShiftDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.TblApproveShiftDet.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblApproveShiftDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblApproveShiftDet.DeptID = dbo.TblEmpDepartments.DeparmentID"
sql = sql & "  Where (dbo.TblApproveShiftDet.ApprovID = " & val(TxtSerial1.Text) & ")and dbo.TblApproveShiftDet.TypeTrans is null"
If Rd(2).value = True Then
sql = sql & " order by   dbo.TblApproveShiftDet.EmpID, dbo.TblApproveShiftDet.RecordDate"
ElseIf Rd(0).value = True Then
sql = sql & " order by   dbo.TblApproveShiftDet.EmpID"
ElseIf Rd(1).value = True Then
sql = sql & " order by   dbo.TblApproveShiftDet.RecordDate"
End If
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .Cell(flexcpChecked, i, .ColIndex("selected")) = IIf(IsNull(Rs1("selected").value), 0, Rs1("selected").value)
                   .TextMatrix(i, .ColIndex("sortdate")) = IIf(IsNull(Rs1("sortdate").value), "", Rs1("sortdate").value)
                   .TextMatrix(i, .ColIndex("DelayID")) = IIf(IsNull(Rs1("DelayID").value), 0, Rs1("DelayID").value)
                   .TextMatrix(i, .ColIndex("IDImport")) = IIf(IsNull(Rs1("IDImport").value), 0, Rs1("IDImport").value)
                   .TextMatrix(i, .ColIndex("IDImport2")) = IIf(IsNull(Rs1("IDImport2").value), 0, Rs1("IDImport2").value)
                   .TextMatrix(i, .ColIndex("AbsenceTime")) = IIf(IsNull(Rs1("AbsenceTime").value), "", Rs1("AbsenceTime").value)
                   .TextMatrix(i, .ColIndex("AbsenceTimeVal")) = IIf(IsNull(Rs1("AbsenceTimeVal").value), "", Rs1("AbsenceTimeVal").value)
                   .TextMatrix(i, .ColIndex("EarExitTime")) = IIf(IsNull(Rs1("EarExitTime").value), "", Rs1("EarExitTime").value)
                   .TextMatrix(i, .ColIndex("DelayTimeTime")) = IIf(IsNull(Rs1("DelayTimeTime").value), "", Rs1("DelayTimeTime").value)
                   .TextMatrix(i, .ColIndex("AdditioTime")) = IIf(IsNull(Rs1("AdditioTime").value), "", Rs1("AdditioTime").value)
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("EnterTime")) = IIf(IsNull(Rs1("EnterTime").value), "", Rs1("EnterTime").value)
                   .TextMatrix(i, .ColIndex("OutTime")) = IIf(IsNull(Rs1("OutTime").value), "", Rs1("OutTime").value)
                   .TextMatrix(i, .ColIndex("FromTime")) = IIf(IsNull(Rs1("FromTime").value), "", Rs1("FromTime").value)
                   .TextMatrix(i, .ColIndex("ToTime")) = IIf(IsNull(Rs1("ToTime").value), "", Rs1("ToTime").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjID").value), 0, Rs1("ProjID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("TypeDay")) = IIf(IsNull(Rs1("TypeDay").value), -1, Rs1("TypeDay").value)
                   .TextMatrix(i, .ColIndex("ShiftID")) = IIf(IsNull(Rs1("ShiftID").value), 0, Rs1("ShiftID").value)
                   .TextMatrix(i, .ColIndex("Absence")) = IIf(IsNull(Rs1("Absence").value), "", Rs1("Absence").value)
                   .TextMatrix(i, .ColIndex("EarExit")) = IIf(IsNull(Rs1("EarExit").value), 0, Rs1("EarExit").value)
                   .TextMatrix(i, .ColIndex("Additional")) = IIf(IsNull(Rs1("Additional").value), 0, Rs1("Additional").value)
                 '  .TextMatrix(i, .ColIndex("NoRegOut")) = IIf(IsNull(Rs1("NoRegOut").value), 0, Rs1("NoRegOut").value)
                   .TextMatrix(i, .ColIndex("MachinCode")) = IIf(IsNull(Rs1("MachinCode").value), "", Rs1("MachinCode").value)
                   If .TextMatrix(i, .ColIndex("MachinCode")) = "" Then
                   .Cell(flexcpBackColor, i, 1, i, 34) = &HFFC0C0
                   End If
                   If Not IsNull(Rs1("NotFingPrin").value) Then
                   If (Rs1("NotFingPrin").value) = 1 Then
                   .Cell(flexcpBackColor, i, 1, i, 34) = &H808080
                   End If
                   End If
                   .TextMatrix(i, .ColIndex("NotFingPrin")) = IIf(IsNull(Rs1("NotFingPrin").value), "", Rs1("NotFingPrin").value)
                   .TextMatrix(i, .ColIndex("DelayType")) = IIf(IsNull(Rs1("DelayType").value), "", Rs1("DelayType").value)
                   .TextMatrix(i, .ColIndex("AbsenType")) = IIf(IsNull(Rs1("AbsenType").value), "", Rs1("AbsenType").value)
                   .TextMatrix(i, .ColIndex("EarExitType")) = IIf(IsNull(Rs1("EarExitType").value), "", Rs1("EarExitType").value)
                   .TextMatrix(i, .ColIndex("AddiType")) = IIf(IsNull(Rs1("AddiType").value), "", Rs1("AddiType").value)
                   .TextMatrix(i, .ColIndex("NoRegOutType")) = IIf(IsNull(Rs1("NoRegOutType").value), "", Rs1("NoRegOutType").value)
              
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentName").value), "", Rs1("DepartmentName").value)
                   .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_name").value), "", Rs1("Project_name").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   .TextMatrix(i, .ColIndex("SheftName")) = IIf(IsNull(Rs1("SheftName").value), "", Rs1("SheftName").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("SheftName")) = IIf(IsNull(Rs1("SheftNamee").value), "", Rs1("SheftNamee").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_nameE").value), "", Rs1("Project_nameE").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentNamee").value), "", Rs1("DepartmentNamee").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   End If
                        If Rs1("NoRegOut").value = 1 Then
                     .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexChecked
                    Else
                    .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexUnchecked
                  End If
                  .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                  
                   Rs1.MoveNext
             Next i
End With
        
        Exit Sub
ErrTrap:
    End Sub
 Sub FullGridData2()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
sql = " SELECT  TblEmployee.fullcode,   dbo.TblApproveShiftDet.ID, dbo.TblApproveShiftDet.RecordDate, dbo.TblApproveShiftDet.ApprovID, dbo.TblApproveShiftDet.EnterTime, "
sql = sql & "                      dbo.TblApproveShiftDet.OutTime, dbo.TblApproveShiftDet.FromTime, dbo.TblApproveShiftDet.ToTime, dbo.TblApproveShiftDet.Absence,"
sql = sql & "                      dbo.TblApproveShiftDet.EarExit, dbo.TblApproveShiftDet.Additional, dbo.TblApproveShiftDet.NoRegOut, dbo.TblApproveShiftDet.MachinCode,"
sql = sql & "                      dbo.TblApproveShiftDet.DelayType, dbo.TblApproveShiftDet.AbsenType, dbo.TblApproveShiftDet.EarExitType, dbo.TblApproveShiftDet.AddiType,"
sql = sql & "                      dbo.TblApproveShiftDet.NoRegOutType, dbo.TblApproveShiftDet.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
sql = sql & "                      dbo.TblApproveShiftDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblApproveShiftDet.ProjID,"
sql = sql & "                      dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblApproveShiftDet.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                      dbo.TblApproveShiftDet.ShiftID , dbo.TbLSheft.SheftName, dbo.TbLSheft.SheftNamee, dbo.TblApproveShiftDet.TypeDay,"
sql = sql & "                      dbo.TblApproveShiftDet.EarExitTime,dbo.TblApproveShiftDet.DelayTimeTime, dbo.TblApproveShiftDet.AdditioTime, dbo.TblApproveShiftDet.AbsenceTime ,dbo.TblApproveShiftDet.IDImport ,"
sql = sql & " dbo.TblApproveShiftDet.DelayID ,dbo.TblApproveShiftDet.AbsenceTimeVal,dbo.TblApproveShiftDet.NotFingPrin ,dbo.TblApproveShiftDet.selected ,dbo.TblApproveShiftDet.IDImport2  ,dbo.TblApproveShiftDet.sortdate ,dbo.TblApproveShiftDet.TypeTrans"
sql = sql & " FROM         dbo.TblApproveShiftDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TbLSheft ON dbo.TblApproveShiftDet.ShiftID = dbo.TbLSheft.SeftCode LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblApproveShiftDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.TblApproveShiftDet.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblApproveShiftDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblApproveShiftDet.DeptID = dbo.TblEmpDepartments.DeparmentID"
sql = sql & "  Where (dbo.TblApproveShiftDet.ApprovID = " & val(TxtSerial1.Text) & ")and dbo.TblApproveShiftDet.TypeTrans=1"
If Rd(2).value = True Then
sql = sql & " order by   dbo.TblApproveShiftDet.EmpID, dbo.TblApproveShiftDet.RecordDate"
ElseIf Rd(0).value = True Then
sql = sql & " order by   dbo.TblApproveShiftDet.EmpID"
ElseIf Rd(1).value = True Then
sql = sql & " order by   dbo.TblApproveShiftDet.RecordDate"
End If
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     
     With Me.VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .Cell(flexcpChecked, i, .ColIndex("selected")) = IIf(IsNull(Rs1("selected").value), 0, Rs1("selected").value)
                   .TextMatrix(i, .ColIndex("sortdate")) = IIf(IsNull(Rs1("sortdate").value), "", Rs1("sortdate").value)
                   .TextMatrix(i, .ColIndex("DelayID")) = IIf(IsNull(Rs1("DelayID").value), 0, Rs1("DelayID").value)
                   .TextMatrix(i, .ColIndex("IDImport")) = IIf(IsNull(Rs1("IDImport").value), 0, Rs1("IDImport").value)
                   .TextMatrix(i, .ColIndex("IDImport2")) = IIf(IsNull(Rs1("IDImport2").value), 0, Rs1("IDImport2").value)
                   .TextMatrix(i, .ColIndex("AbsenceTime")) = IIf(IsNull(Rs1("AbsenceTime").value), "", Rs1("AbsenceTime").value)
                   .TextMatrix(i, .ColIndex("AbsenceTimeVal")) = IIf(IsNull(Rs1("AbsenceTimeVal").value), "", Rs1("AbsenceTimeVal").value)
                   .TextMatrix(i, .ColIndex("EarExitTime")) = IIf(IsNull(Rs1("EarExitTime").value), "", Rs1("EarExitTime").value)
                   .TextMatrix(i, .ColIndex("DelayTimeTime")) = IIf(IsNull(Rs1("DelayTimeTime").value), "", Rs1("DelayTimeTime").value)
                   .TextMatrix(i, .ColIndex("AdditioTime")) = IIf(IsNull(Rs1("AdditioTime").value), "", Rs1("AdditioTime").value)
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("EnterTime")) = IIf(IsNull(Rs1("EnterTime").value), "", Rs1("EnterTime").value)
                   .TextMatrix(i, .ColIndex("OutTime")) = IIf(IsNull(Rs1("OutTime").value), "", Rs1("OutTime").value)
                   .TextMatrix(i, .ColIndex("FromTime")) = IIf(IsNull(Rs1("FromTime").value), "", Rs1("FromTime").value)
                   .TextMatrix(i, .ColIndex("ToTime")) = IIf(IsNull(Rs1("ToTime").value), "", Rs1("ToTime").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("DeptID")) = IIf(IsNull(Rs1("DeptID").value), 0, Rs1("DeptID").value)
                   .TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(Rs1("ProjID").value), 0, Rs1("ProjID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("TypeDay")) = IIf(IsNull(Rs1("TypeDay").value), -1, Rs1("TypeDay").value)
                   .TextMatrix(i, .ColIndex("ShiftID")) = IIf(IsNull(Rs1("ShiftID").value), 0, Rs1("ShiftID").value)
                   .TextMatrix(i, .ColIndex("Absence")) = IIf(IsNull(Rs1("Absence").value), "", Rs1("Absence").value)
                   .TextMatrix(i, .ColIndex("EarExit")) = IIf(IsNull(Rs1("EarExit").value), 0, Rs1("EarExit").value)
                   .TextMatrix(i, .ColIndex("Additional")) = IIf(IsNull(Rs1("Additional").value), 0, Rs1("Additional").value)
                 '  .TextMatrix(i, .ColIndex("NoRegOut")) = IIf(IsNull(Rs1("NoRegOut").value), 0, Rs1("NoRegOut").value)
                   .TextMatrix(i, .ColIndex("MachinCode")) = IIf(IsNull(Rs1("MachinCode").value), "", Rs1("MachinCode").value)
                   If .TextMatrix(i, .ColIndex("MachinCode")) = "" Then
                   .Cell(flexcpBackColor, i, 1, i, 34) = &HFFC0C0
                   End If
                   If Not IsNull(Rs1("NotFingPrin").value) Then
                   If (Rs1("NotFingPrin").value) = 1 Then
                   .Cell(flexcpBackColor, i, 1, i, 34) = &H808080
                   End If
                   End If
                   .TextMatrix(i, .ColIndex("NotFingPrin")) = IIf(IsNull(Rs1("NotFingPrin").value), "", Rs1("NotFingPrin").value)
                   .TextMatrix(i, .ColIndex("DelayType")) = IIf(IsNull(Rs1("DelayType").value), "", Rs1("DelayType").value)
                   .TextMatrix(i, .ColIndex("AbsenType")) = IIf(IsNull(Rs1("AbsenType").value), "", Rs1("AbsenType").value)
                   .TextMatrix(i, .ColIndex("EarExitType")) = IIf(IsNull(Rs1("EarExitType").value), "", Rs1("EarExitType").value)
                   .TextMatrix(i, .ColIndex("AddiType")) = IIf(IsNull(Rs1("AddiType").value), "", Rs1("AddiType").value)
                   .TextMatrix(i, .ColIndex("NoRegOutType")) = IIf(IsNull(Rs1("NoRegOutType").value), "", Rs1("NoRegOutType").value)
              
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentName").value), "", Rs1("DepartmentName").value)
                   .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_name").value), "", Rs1("Project_name").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   .TextMatrix(i, .ColIndex("SheftName")) = IIf(IsNull(Rs1("SheftName").value), "", Rs1("SheftName").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("SheftName")) = IIf(IsNull(Rs1("SheftNamee").value), "", Rs1("SheftNamee").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs1("Project_nameE").value), "", Rs1("Project_nameE").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs1("DepartmentNamee").value), "", Rs1("DepartmentNamee").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   End If
                        If Rs1("NoRegOut").value = 1 Then
                     .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexChecked
                    Else
                    .Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexUnchecked
                  End If
                  .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                  
                   Rs1.MoveNext
             Next i
End With
        
        Exit Sub
ErrTrap:
    End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "Emp_Name"
Cancel = True
Case "RecordDate"
Cancel = True
Case "EnterTime"
Cancel = True
Case "OutTime"
Cancel = True
Case "SheftName"
Cancel = True
Case "FromTime"
Cancel = True
Case "ToTime"
Cancel = True
Case "DelayTimeTime"
Cancel = True
Case "DelayType"
.ComboList = ""
Case "EarExitTime"
Cancel = True
Case "EarExitType"
.ComboList = ""
Case "Absence"
Cancel = True
Case "AbsenType"
.ComboList = ""
Case "AdditioTime"
Cancel = True
Case "AddiType"
.ComboList = ""
Case "NoRegOut"
Cancel = True
Case "NoRegOutType"
.ComboList = ""

End Select
End With
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
Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
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
 If SystemOptions.AllowTowShift = True Then
DeleRepatit
 End If
filgrid1
filgrid2
End If
End Sub
Sub FilGridDate()
Dim DayNO As Integer
Dim i As Integer
DayNO = DateDiff("d", Fromdate.value, ToDate.value) + 1
With fg
.Rows = DayNO + 1
For i = 1 To DayNO
If i = 1 Then
NoDate.value = Fromdate.value
Else
NoDate.value = DateAdd("d", i - 1, Me.Fromdate.value)
End If
.TextMatrix(i, .ColIndex("RecDate")) = NoDate.value
'.TextMatrix(i, .ColIndex("Dy")) = CheckHolidaies(NoDate.value)
Next i
End With

End Sub
Sub DeleRepatit()
Dim i As Integer
Dim sql As String
Dim Flag As Integer
Dim TempID As Double
Dim MachinCode As String
Dim RecTime As String
Dim MachinDate As Date
Dim TepMachinCode As String
Dim TempMachinDate As Date
Dim ShiftID As Double
Dim TempShiftID As Double
Dim ID As Double

DTPicker2.value = "01/01/1999"
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT BranchID,ProjID,BranchID,DeptID,EmpID, ID,MachinCode, MachinDate, Sortdate,RecTime ,ShiftID"
sql = sql & " From dbo.TblImportShiftsDet"
sql = sql & " WHERE   TblImportShiftsDet.EmpID<>0 and   (IsSele is null) and   not(ShiftID is null)and ShiftID<>0 "
sql = sql & " AND( dbo.TblImportShiftsDet.MachinDate >=" & SQLDate(Fromdate.value, True) & " )"
'Sql = Sql & "  AND (CONVERT(varchar, dbo.TblImportShiftsDet.MachinDate, 103) >= (CONVERT(varchar, " & Me.Fromdate.value & ", 103)"
sql = sql & " AND (dbo.TblImportShiftsDet.MachinDate <= " & SQLDate(ToDate.value, True) & ")"
'Sql = Sql & "  AND (CONVERT(varchar, dbo.TblImportShiftsDet.MachinDate, 103) <= (CONVERT(varchar, " & Me.toDate.value & ", 103)"
If val(Me.DcbProject1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.ProjID  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.BranchID  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.DeptID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.EmpID  =" & val(DcbEmployee1.BoundText) & " "
End If
sql = sql & " ORDER BY MachinCode,MachinDate,sortdate"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
Rs5.MoveFirst
For i = 1 To Rs5.RecordCount
MachinCode = IIf(IsNull(Rs5("MachinCode").value), "", Rs5("MachinCode").value)
MachinDate = IIf(IsNull(Rs5("MachinDate").value), "01/01/1999", Rs5("MachinDate").value)
ShiftID = IIf(IsNull(Rs5("ShiftID").value), 0, Rs5("ShiftID").value)
ID = IIf(IsNull(Rs5("ID").value), 0, Rs5("ID").value)

If i = 1 Then
TepMachinCode = MachinCode
TempMachinDate = MachinDate
TempShiftID = ShiftID
TempID = ID
Flag = 1
Else
If TepMachinCode = MachinCode And TempMachinDate = MachinDate And ShiftID = TempShiftID And Flag = 3 Then

 If Flag = 3 Then
 Cn.Execute " Update TblImportShiftsDet set FlgDeleted=1  where ID=" & TempID & " "
 End If
 Flag = 0
 TempID = IIf(IsNull(Rs5("ID").value), 0, Rs5("ID").value)
 ElseIf TepMachinCode = MachinCode And TempMachinDate = MachinDate And ShiftID = TempShiftID And Flag = 1 Then
Flag = 2
TepMachinCode = MachinCode
TempMachinDate = MachinDate
TempShiftID = ShiftID
TempID = ID
 ElseIf TepMachinCode = MachinCode And TempMachinDate = MachinDate And ShiftID = TempShiftID And (Flag = 0 Or Flag = 2) Then
Flag = 3
 Cn.Execute " Update TblImportShiftsDet set FlgDeleted=1  where ID=" & TempID & " "
TepMachinCode = MachinCode
TempMachinDate = MachinDate
TempShiftID = ShiftID
TempID = ID
Else
TepMachinCode = MachinCode
TempMachinDate = MachinDate
TempShiftID = ShiftID
TempID = ID
Flag = 1
End If
End If
Rs5.MoveNext
Next i
End If
End Sub
Sub filgrid2()
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim cont As Integer
Dim i, k As Integer
Dim sql As String
Dim EmpID As Double
Dim RecTime As Date
Dim ShiftID As Double
Dim RecTime2 As String
TempDate.value = "01/01/1999"
  VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
  VSFlexGrid1.Rows = 1
sql = "seLECT    dbo.TblImportShiftsDet.sortdate, CONVERT(varchar, dbo.TblImportShiftsDet.MachinDate, 103) AS MachinDate1, dbo.TblImportShiftsDet.ImportShiftID, dbo.TblImportShiftsDet.ID, dbo.TblImportShiftsDet.EmpID, "
sql = sql & "                      dbo.TblImportShiftsDet.MachinCode, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblImportShiftsDet.RecTime, dbo.TblImportShiftsDet.BranchID,"
sql = sql & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblImportShiftsDet.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblImportShiftsDet.DeptID,"
sql = sql & "                      dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee, dbo.TblImportShiftsDet.ShiftID"
sql = sql & " FROM         dbo.TblImportShiftsDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblImportShiftsDet.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.TblImportShiftsDet.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblImportShiftsDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblImportShiftsDet.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " WHERE   TblImportShiftsDet.EmpID<>0 and   (IsSele is null)and dbo.TblEmployee.workstate=1 "
sql = sql & " AND( dbo.TblImportShiftsDet.MachinDate >=" & SQLDate(Fromdate.value, True) & " )and dbo.TblImportShiftsDet.ShiftID=0 "

If val(Me.DcbProject1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.ProjID  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.BranchID  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.DeptID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.EmpID  =" & val(DcbEmployee1.BoundText) & " "
End If
If Rd(0).value = True Then
sql = sql & " ORDER BY dbo.TblImportShiftsDet.EmpID, dbo.TblImportShiftsDet.MachinDate, dbo.TblImportShiftsDet.SortDate ,dbo.TblImportShiftsDet.ID"
ElseIf Rd(1).value = True Then
sql = sql & " ORDER BY  dbo.TblImportShiftsDet.MachinDate, dbo.TblImportShiftsDet.SortDate ,dbo.TblImportShiftsDet.EmpID,dbo.TblImportShiftsDet.ID"
End If
 Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With VSFlexGrid1
k = .Rows
cont = k
Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For i = k To .Rows - 1
.Cell(flexcpChecked, i, .ColIndex("selected")) = True
RecTime2 = FormatDateTime(Rs8("RecTime").value, vbShortTime)
.TextMatrix(cont, .ColIndex("sortdate")) = IIf(IsNull(Rs8("sortdate").value), "01/01/1999", Rs8("sortdate").value)
EmpID = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
ShiftID = IIf(IsNull(Rs8("ShiftID").value), 0, Rs8("ShiftID").value)
TempDate.value = IIf(IsNull(Rs8("MachinDate1").value), TempDate.value, Rs8("MachinDate1").value)
 
If val(.TextMatrix(cont - 1, .ColIndex("Flg"))) = 1 And cont <> 1 And EmpID <> 0 And ShiftID = val(.TextMatrix(cont - 1, .ColIndex("ShiftID"))) And EmpID = val(.TextMatrix(cont - 1, .ColIndex("EmpID"))) And TempDate.value = (.TextMatrix(cont - 1, .ColIndex("RecordDate"))) Then
    If Not IsNull(Rs8("RecTime").value) Then
        RecTime = FormatDateTime(Rs8("RecTime").value, vbShortTime)
        .TextMatrix(cont - 1, .ColIndex("OutTime")) = RecTime
        .TextMatrix(cont - 1, .ColIndex("IDImport2")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
        
    End If
     .TextMatrix(cont - 1, .ColIndex("Flg")) = 0
'.TextMatrix(cont - 1, .ColIndex("OutTime")) = IIf(IsNull(Rs8("RecTime").value), "", Rs8("RecTime").value)
Else
.TextMatrix(cont, .ColIndex("RecordDate")) = IIf(IsNull(Rs8("MachinDate1").value), TempDate.value, Rs8("MachinDate1").value)
.TextMatrix(cont, .ColIndex("EmpID")) = EmpID
.TextMatrix(cont, .ColIndex("ShiftID")) = ShiftID
.TextMatrix(cont, .ColIndex("Ser")) = cont
    If Not IsNull(Rs8("RecTime").value) Then
        RecTime = FormatDateTime(Rs8("RecTime").value, vbShortTime)
        .TextMatrix(cont, .ColIndex("EnterTime")) = RecTime
        .TextMatrix(cont, .ColIndex("IDImport")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
    End If
    .TextMatrix(cont, .ColIndex("Flg")) = 1
'.TextMatrix(cont, .ColIndex("EnterTime")) = IIf(IsNull(Rs8("RecTime").value), "", Rs8("RecTime").value)
.TextMatrix(cont, .ColIndex("ProjID")) = IIf(IsNull(Rs8("ProjID").value), 0, Rs8("ProjID").value)
.TextMatrix(cont, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchID").value), 0, Rs8("BranchID").value)
.TextMatrix(cont, .ColIndex("MachinCode")) = IIf(IsNull(Rs8("MachinCode").value), "", Rs8("MachinCode").value)
.TextMatrix(cont, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(cont, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DeptID").value), 0, Rs8("DeptID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(cont, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), "", Rs8("Project_name").value)
.TextMatrix(cont, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
.TextMatrix(cont, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(cont, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
Else
.TextMatrix(cont, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), "", Rs8("Project_nameE").value)
.TextMatrix(cont, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
.TextMatrix(cont, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(cont, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
End If

cont = cont + 1
End If
H:
Rs8.MoveNext
Next i
'.AutoSize 0, .Cols - 1, False
.Rows = cont
End With
End If

End Sub
Sub filgrid1()
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim cont As Integer
Dim i, k As Integer
Dim sql As String
Dim EmpID As Double
Dim RecTime As Date
Dim ShiftID As Double
Dim RecTime2 As String
  GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
TempDate.value = "01/01/1999"
sql = "seLECT    dbo.TblImportShiftsDet.sortdate, CONVERT(varchar, dbo.TblImportShiftsDet.MachinDate, 103) AS MachinDate1, dbo.TblImportShiftsDet.ImportShiftID, dbo.TblImportShiftsDet.ID, dbo.TblImportShiftsDet.EmpID, "
sql = sql & "                      dbo.TblImportShiftsDet.MachinCode, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblImportShiftsDet.RecTime, dbo.TblImportShiftsDet.BranchID,"
sql = sql & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblImportShiftsDet.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblImportShiftsDet.DeptID,"
sql = sql & "                      dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee, dbo.TblImportShiftsDet.ShiftID"
sql = sql & " FROM         dbo.TblImportShiftsDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblImportShiftsDet.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.TblImportShiftsDet.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblImportShiftsDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblImportShiftsDet.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " WHERE   TblImportShiftsDet.EmpID<>0 and   (IsSele is null)and dbo.TblEmployee.workstate=1 "
sql = sql & " AND( dbo.TblImportShiftsDet.MachinDate >=" & SQLDate(Fromdate.value, True) & " )and dbo.TblImportShiftsDet.ShiftID<>0 and not(dbo.TblImportShiftsDet.ShiftID is null)"
sql = sql & " AND (dbo.TblImportShiftsDet.MachinDate <= " & SQLDate(ToDate.value, True) & ")"
 If SystemOptions.AllowTowShift = True Then
sql = sql & " and dbo.TblImportShiftsDet.FlgDeleted  is null"
 End If
If val(Me.DcbProject1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.ProjID  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.BranchID  =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.DeptID  =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 Then
sql = sql & " and dbo.TblImportShiftsDet.EmpID  =" & val(DcbEmployee1.BoundText) & " "
End If
If Rd(0).value = True Then
sql = sql & " ORDER BY dbo.TblImportShiftsDet.EmpID, dbo.TblImportShiftsDet.MachinDate, dbo.TblImportShiftsDet.SortDate ,dbo.TblImportShiftsDet.ID"
ElseIf Rd(1).value = True Then
sql = sql & " ORDER BY  dbo.TblImportShiftsDet.MachinDate, dbo.TblImportShiftsDet.SortDate ,dbo.TblImportShiftsDet.EmpID,dbo.TblImportShiftsDet.ID"
End If
 Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs8.RecordCount > 0 Then
With GridInstallments
k = .Rows
cont = k
Rs8.MoveFirst
.Rows = .Rows + Rs8.RecordCount
For i = k To .Rows - 1
.Cell(flexcpChecked, i, .ColIndex("selected")) = True
RecTime2 = FormatDateTime(Rs8("RecTime").value, vbShortTime)
.TextMatrix(cont, .ColIndex("sortdate")) = IIf(IsNull(Rs8("sortdate").value), "01/01/1999", Rs8("sortdate").value)
EmpID = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
ShiftID = IIf(IsNull(Rs8("ShiftID").value), 0, Rs8("ShiftID").value)
TempDate.value = IIf(IsNull(Rs8("MachinDate1").value), TempDate.value, Rs8("MachinDate1").value)
 
  DTPicker1.value = IIf(IsNull(Rs8("MachinDate1").value), "01/01/1999", Rs8("MachinDate1").value)
' If CheckBetwenPeriod(EmpID, Weekday(DTPicker1), RecTime2) = 0 Then
' If count <> 1 Then
' If .TextMatrix(cont - 1, .ColIndex("EnterTime")) = "" Then
' ShiftID = GetShiftId(EmpID, .TextMatrix(cont, .ColIndex("EnterTime")), Weekday(DTPicker1))
' ElseIf cont <> 1 And EmpID <> 0 And EmpID = val(.TextMatrix(cont - 1, .ColIndex("EmpID"))) And TempDate.value = (.TextMatrix(cont - 1, .ColIndex("RecordDate"))) Then
' ShiftID = GetShiftId2(EmpID, .TextMatrix(cont, .ColIndex("OutTime")), Weekday(DTPicker1))
 '   If Not IsNull(Rs8("RecTime").value) Then
 '       RecTime = FormatDateTime(Rs8("RecTime").value, vbShortTime)
 '       .TextMatrix(cont - 1, .ColIndex("OutTime")) = RecTime
 '       .TextMatrix(cont - 1, .ColIndex("IDImport2")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
 '      GoTo H
 '   End If
 'End If
 'Else
' ShiftID = GetShiftId(EmpID, .TextMatrix(cont, .ColIndex("EnterTime")), Weekday(DTPicker1))
' End If
'Else
'ShiftID = CheckBetwenPeriod(EmpID, Weekday(DTPicker1), RecTime2)
'End If

If val(.TextMatrix(cont - 1, .ColIndex("Flg"))) = 1 And cont <> 1 And EmpID <> 0 And ShiftID = val(.TextMatrix(cont - 1, .ColIndex("ShiftID"))) And EmpID = val(.TextMatrix(cont - 1, .ColIndex("EmpID"))) And TempDate.value = (.TextMatrix(cont - 1, .ColIndex("RecordDate"))) Then
    If Not IsNull(Rs8("RecTime").value) Then
        RecTime = FormatDateTime(Rs8("RecTime").value, vbShortTime)
        .TextMatrix(cont - 1, .ColIndex("OutTime")) = RecTime
        .TextMatrix(cont - 1, .ColIndex("IDImport2")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
        
    End If
     .TextMatrix(cont - 1, .ColIndex("Flg")) = 0
'.TextMatrix(cont - 1, .ColIndex("OutTime")) = IIf(IsNull(Rs8("RecTime").value), "", Rs8("RecTime").value)
Else
.TextMatrix(cont, .ColIndex("RecordDate")) = IIf(IsNull(Rs8("MachinDate1").value), TempDate.value, Rs8("MachinDate1").value)
.TextMatrix(cont, .ColIndex("EmpID")) = EmpID
.TextMatrix(cont, .ColIndex("ShiftID")) = ShiftID
.TextMatrix(cont, .ColIndex("Ser")) = cont
    If Not IsNull(Rs8("RecTime").value) Then
        RecTime = FormatDateTime(Rs8("RecTime").value, vbShortTime)
        .TextMatrix(cont, .ColIndex("EnterTime")) = RecTime
        .TextMatrix(cont, .ColIndex("IDImport")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
    End If
    .TextMatrix(cont, .ColIndex("Flg")) = 1
'.TextMatrix(cont, .ColIndex("EnterTime")) = IIf(IsNull(Rs8("RecTime").value), "", Rs8("RecTime").value)
.TextMatrix(cont, .ColIndex("ProjID")) = IIf(IsNull(Rs8("ProjID").value), 0, Rs8("ProjID").value)
.TextMatrix(cont, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchID").value), 0, Rs8("BranchID").value)
.TextMatrix(cont, .ColIndex("MachinCode")) = IIf(IsNull(Rs8("MachinCode").value), "", Rs8("MachinCode").value)
.TextMatrix(cont, .ColIndex("FullCode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(cont, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DeptID").value), 0, Rs8("DeptID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(cont, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), "", Rs8("Project_name").value)
.TextMatrix(cont, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
.TextMatrix(cont, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(cont, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentName").value), "", Rs8("DepartmentName").value)
Else
.TextMatrix(cont, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), "", Rs8("Project_nameE").value)
.TextMatrix(cont, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
.TextMatrix(cont, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(cont, .ColIndex("DepartmentName")) = IIf(IsNull(Rs8("DepartmentNamee").value), "", Rs8("DepartmentNamee").value)
End If

cont = cont + 1
End If
H:
Rs8.MoveNext
Next i
'.AutoSize 0, .Cols - 1, False
.Rows = cont
End With
End If
FillGrid2
End Sub
'Function GetShiftId2(Optional EmpID As Double, Optional Timin As String, Optional NoDay As Integer) As Double
'Dim Sql As String
'Dim I As Integer
'Dim Rs7 As ADODB.Recordset
'Set Rs7 = New ADODB.Recordset
'Dim temp As String
'Dim swap As String
'Dim ShiftIDTemp As Double
'Dim ShiftID As Double
'Sql = " SELECT     ShiftID, EmpID"
'Sql = Sql & " From dbo.TblShiftWorker"
'Sql = Sql & " WHERE     (EmpID = " & EmpID & ") "
'Rs7.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If Rs7.RecordCount > 0 Then
'Rs7.MoveFirst
'For I = 1 To Rs7.RecordCount
'ShiftID = IIf(IsNull(Rs7("ShiftID").value), 0, Rs7("ShiftID").value)
'If CheckPeriodout(ShiftID, NoDay) <> "" Then
'If I <> 1 Then
'swap = temp
'Else
''ShiftIDTemp = ShiftID
'End If
'temp = CheckPeriod(ShiftID, NoDay)
'If I <> 1 Then
'If Timin <> "" Then
'If Abs(DateDiff("n", Timin, swap)) > Abs(DateDiff("n", Timin, temp)) Then
'ShiftIDTemp = ShiftID
'End If
'End If
'End If
'End If
'Rs7.MoveNext
'Next I
'Else
'GetShiftId2 = 0
'Exit Function
'End If
'GetShiftId2 = ShiftIDTemp
'End Function
'Function GetShiftId(Optional EmpID As Double, Optional Timin As String, Optional NoDay As Integer) As Double
'Dim Sql As String
'Dim I As Integer
'Dim Rs7 As ADODB.Recordset
'Set Rs7 = New ADODB.Recordset
'Dim temp As String
'Dim swap As String
'Dim ShiftIDTemp As Double
'Dim ShiftID As Double
'Sql = " SELECT     ShiftID, EmpID"
'Sql = Sql & " From dbo.TblShiftWorker"
'Sql = Sql & " WHERE     (EmpID = " & EmpID & ") "
'Rs7.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If Rs7.RecordCount > 0 Then
'Rs7.MoveFirst
'For I = 1 To Rs7.RecordCount
'ShiftID = IIf(IsNull(Rs7("ShiftID").value), 0, Rs7("ShiftID").value)
'If CheckPeriod(ShiftID, NoDay) <> "" Then
'If I <> 1 Then
'swap = temp
'Else
'ShiftIDTemp = ShiftID
'End If
'temp = CheckPeriod(ShiftID, NoDay)
'If I <> 1 Then
'If Timin <> "" Then
'If Abs(DateDiff("n", Timin, swap)) > Abs(DateDiff("n", Timin, temp)) Then
'ShiftIDTemp = ShiftID
'End If
'End If
'End If
'End If
'Rs7.MoveNext
'Next I
'Else
'GetShiftId = 0
'Exit Function
'End If
'GetShiftId = ShiftIDTemp
'End Function
'Function CheckBetwenPeriod(Optional empID1 As Double, Optional NoDay As Integer, Optional Timin As String) As Double
'Dim Sql As String
'Dim Rs8 As ADODB.Recordset
'Set Rs8 = New ADODB.Recordset

'Sql = " SELECT     dbo.TblShiftWorker.EmpID AS EmpID1, dbo.TbLSheft.*"
'Sql = Sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
'Sql = Sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
'Sql = Sql & " Where dbo.TblShiftWorker.EmpID=" & empID1 & ""
'Select Case NoDay
'Case 7
'Sql = Sql & " and  Shiftfrom <= '" & Timin & "' "
'Sql = Sql & " and ShiftTo >= '" & Timin & "' "
'Case 6
'Sql = Sql & " and  FromFri <= '" & Timin & "' "
'Sql = Sql & " and ToFri >= '" & Timin & "' "
'Case 5
'Sql = Sql & " and  FromThru <= '" & Timin & "' "
'Sql = Sql & " and ToThru >= '" & Timin & "' "
'Case 4
'Sql = Sql & " and  FromWed <= '" & Timin & "' "
'Sql = Sql & " and ToWed >= '" & Timin & "' "
'Case 3
'Sql = Sql & " and  FromTues <= '" & Timin & "' "
'Sql = Sql & " and ToTues >=  '" & Timin & "' "
'Case 2
'Sql = Sql & " and  FromMon <= '" & Timin & "' "
'Sql = Sql & " and ToMon >= '" & Timin & "' "
'Case 1
'Sql = Sql & " and  FromSun <= '" & Timin & "' "
'Sql = Sql & " and ToSun >= '" & Timin & "' "
'End Select

'Rs8.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'If Rs8.RecordCount > 0 Then
'CheckBetwenPeriod = IIf(IsNull(Rs8("SeftCode").value), 0, (Rs8("SeftCode").value))
'Else
'CheckBetwenPeriod = 0
'End If
'End Function
Sub RetriveShifts(Optional ShiftID As Double = 0, Optional ByRef TimeinExists As String, Optional ByRef TimoutExists As String, Optional ByRef TypeDay As Integer = -1, Optional NoDay As Integer)
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
TimeinExists = FormatDateTime(IIf(IsNull(Rs8("FromWed").value), 0, (Rs8("FromWed").value)), vbShortTime)
TimoutExists = FormatDateTime(IIf(IsNull(Rs8("ToWed").value), 0, (Rs8("ToWed").value)), vbShortTime)
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
'  Function CheckPeriodout(Optional ShiftID As Double, Optional NoDay As Integer) As String
'Dim Sql As String
'Dim Rs8 As ADODB.Recordset
'Set Rs8 = New ADODB.Recordset
'Sql = " SELECT     dbo.TbLSheft.*"
'Sql = Sql & " FROM         dbo.TbLSheft where SeftCode=" & ShiftID & " "
'Rs8.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'If Rs8.RecordCount > 0 Then
'Select Case NoDay
'Case 7
'CheckPeriodout = IIf(IsNull(Rs8("ShiftTo").value), 0, (Rs8("ShiftTo").value))
'Case 6
'CheckPeriodout = IIf(IsNull(Rs8("ToFri").value), 0, (Rs8("ToFri").value))
'Case 5
'CheckPeriodout = IIf(IsNull(Rs8("ToThru").value), 0, (Rs8("ToThru").value))
'Case 4
'CheckPeriodout = IIf(IsNull(Rs8("ToWed").value), 0, (Rs8("ToWed").value))
'Case 3
'CheckPeriodout = IIf(IsNull(Rs8("ToTues").value), 0, (Rs8("ToTues").value))
'Case 2
'CheckPeriodout = IIf(IsNull(Rs8("ToMon").value), 0, (Rs8("ToMon").value))
'Case 1
'CheckPeriodout = IIf(IsNull(Rs8("ToSun").value), 0, (Rs8("ToSun").value))
'End Select
'Else
'CheckPeriodout = ""
'End If
'End Function
'Function CheckPeriod(Optional ShiftID As Double, Optional NoDay As Integer) As String
'Dim Sql As String
'Dim Rs8 As ADODB.Recordset
'Set Rs8 = New ADODB.Recordset
'Sql = " SELECT     dbo.TbLSheft.*"
'Sql = Sql & " FROM         dbo.TbLSheft where SeftCode=" & ShiftID & " "
'Rs8.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'If Rs8.RecordCount > 0 Then
'Select Case NoDay
'Case 7
'CheckPeriod = IIf(IsNull(Rs8("Shiftfrom").value), 0, (Rs8("Shiftfrom").value))
'Case 6
'CheckPeriod = IIf(IsNull(Rs8("FromFri").value), 0, (Rs8("FromFri").value))
'Case 5
'CheckPeriod = IIf(IsNull(Rs8("FromThru").value), 0, (Rs8("FromThru").value))
'Case 4
'CheckPeriod = IIf(IsNull(Rs8("FromWed").value), 0, (Rs8("FromWed").value))
'Case 3
'CheckPeriod = IIf(IsNull(Rs8("FromTues").value), 0, (Rs8("FromTues").value))
'Case 2
'CheckPeriod = IIf(IsNull(Rs8("FromMon").value), 0, (Rs8("FromMon").value))
'Case 1
'CheckPeriod = IIf(IsNull(Rs8("FromSun").value), 0, (Rs8("FromSun").value))
'End Select
'Else
'CheckPeriod = ""
'End If
'End Function
'Function GetCountAbsecn(Optional EmpID As Double, Optional Row As Integer, Optional RecDate As Date) As Integer
'Dim I As Integer
'Dim cnt As Integer
'cnt = 0
'With GridInstallments
'For I = 1 To Row
'If val(.TextMatrix(I, .ColIndex("AbsenceTime"))) <> 0 Then
'If DateDiff("m", RecDate, .TextMatrix(I, .ColIndex("RecordDate"))) = 0 Then
'cnt = cnt + 1
'End If
'End If
'Next I
'End With
'GetCountAbsecn = cnt
'End Function
Function CheckEmployeeNoAddition(Optional Emp_id As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from  TblEmployee where Emp_ID=" & Emp_id & " and NoAdded=1"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckEmployeeNoAddition = True
Else
CheckEmployeeNoAddition = False
End If
End Function

Function CheckEmployee(Optional Emp_id As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from  TblEmployee where Emp_ID=" & Emp_id & " and workstate=1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckEmployee = True
Else
CheckEmployee = False
End If
End Function
Sub FillGrid2()
Dim i As Integer
Dim CouAbsece As Integer
Dim ShiftID As Double
Dim FromTime As String
Dim ToTime As String
Dim TypeDay As Integer
Dim TimDiff As Double
Dim NoDay As Double
Dim X As Double
Dim RecTime As String
With GridInstallments

For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
   
                               .Row = i
                             .Col = .ColIndex("SheftName")
                             .ShowCell i, .ColIndex("SheftName")
                            
                             .SetFocus

RecTime = FormatDateTime(.TextMatrix(i, .ColIndex("sortdate")), vbShortTime)
     
 DTPicker1.value = IIf(IsDate(.TextMatrix(i, .ColIndex("RecordDate"))), .TextMatrix(i, .ColIndex("RecordDate")), Null)
' If CheckBetwenPeriod(val(.TextMatrix(I, .ColIndex("EmpID"))), Weekday(DTPicker1), RecTime) = 0 Then
'
' ShiftID = GetShiftId(val(.TextMatrix(I, .ColIndex("EmpID"))), .TextMatrix(I, .ColIndex("EnterTime")), Weekday(DTPicker1))
'Else
'ShiftID = CheckBetwenPeriod(val(.TextMatrix(I, .ColIndex("EmpID"))), Weekday(DTPicker1), RecTime)
'End If
ShiftID = val(.TextMatrix(i, .ColIndex("ShiftID")))
If CheckEmployee(val(.TextMatrix(i, .ColIndex("EmpID")))) = False Then
.Cell(flexcpBackColor, i, 1, i, 34) = &H80C0FF
 ElseIf val(.TextMatrix(i, .ColIndex("DeptID"))) = 0 Then
 .Cell(flexcpBackColor, i, 1, i, 34) = &H404080
 ElseIf ShiftID = 0 Then
   .Cell(flexcpBackColor, i, 1, i, 34) = &HFF&
    ElseIf .TextMatrix(i, .ColIndex("OutTime")) = "" Then
  .Cell(flexcpBackColor, i, 1, i, 34) = &HFF8080     'red
 End If
.TextMatrix(i, .ColIndex("MachinCode")) = GetMahineCode(val(.TextMatrix(i, .ColIndex("EmpID"))))
 ShiftID = val(.TextMatrix(i, .ColIndex("ShiftID")))
 .TextMatrix(i, .ColIndex("SheftName")) = GetNamShift(ShiftID)
  RetriveShifts ShiftID, FromTime, ToTime, TypeDay, Weekday(DTPicker1)
  If CheckShiftHolidaies(val(.TextMatrix(i, .ColIndex("EmpID"))), Weekday(DTPicker1.value)) = 1 Then
 .TextMatrix(i, .ColIndex("TypeDay")) = 1
 Else
 .TextMatrix(i, .ColIndex("TypeDay")) = CheckHolidaies(DTPicker1.value)
 End If
 .TextMatrix(i, .ColIndex("FromTime")) = FromTime
 .TextMatrix(i, .ColIndex("ToTime")) = ToTime
If val(.TextMatrix(i, .ColIndex("TypeDay"))) = 1 Then
If .TextMatrix(i, .ColIndex("EnterTime")) <> "" And ShiftID <> 0 Then
If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
TimDiff = DateDiff("n", .TextMatrix(i, .ColIndex("EnterTime")), .TextMatrix(i, .ColIndex("OutTime")))
End If
If RereiveSliceValue(ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 2, NoDay) = False Then
  RereiveSlice ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 2, NoDay
  End If
  If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
      X = Abs(DateDiff("n", .TextMatrix(i, .ColIndex("EnterTime")), .TextMatrix(i, .ColIndex("OutTime"))) / 60)
      End If
       X = Fix(X)
If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
   .TextMatrix(i, .ColIndex("AdditioTime")) = X & ":" & Abs(DateDiff("n", .TextMatrix(i, .ColIndex("EnterTime")), .TextMatrix(i, .ColIndex("OutTime"))) Mod 60)
   End If
  .TextMatrix(i, .ColIndex("Additional")) = NoDay
  .TextMatrix(i, .ColIndex("AddiType")) = 1
  
  '''''
  If CheckEmployeeNoAddition(val(.TextMatrix(i, .ColIndex("EmpID")))) = True Then
  .TextMatrix(i, .ColIndex("AddiType")) = 2
  .Cell(flexcpBackColor, i, 1, i, 36) = &HC000&
  End If
  .TextMatrix(i, .ColIndex("EarExit")) = ""
  .TextMatrix(i, .ColIndex("EarExitType")) = ""
  .TextMatrix(i, .ColIndex("EarExitTime")) = ""
   .TextMatrix(i, .ColIndex("DelayID")) = ""
  .TextMatrix(i, .ColIndex("DelayType")) = ""
  End If
Else
  If .TextMatrix(i, .ColIndex("EnterTime")) = "" Then
  If GetShiftNo(val(.TextMatrix(i, .ColIndex("EmpID")))) <> 0 Then
     .TextMatrix(i, .ColIndex("AbsenceTimeVal")) = 1 / GetShiftNo(val(.TextMatrix(i, .ColIndex("EmpID"))))
     .TextMatrix(i, .ColIndex("AbsenceTime")) = 1 / GetShiftNo(val(.TextMatrix(i, .ColIndex("EmpID"))))
  End If
  .TextMatrix(i, .ColIndex("AbsenType")) = 1
     
    ' .TextMatrix(I, .ColIndex("CouAbsece")) = GetCountAbsecn(val(.TextMatrix(I, .ColIndex("EmpID"))), I, DTPicker1.value)
  Else
  TimDiff = DateDiff("n", .TextMatrix(i, .ColIndex("EnterTime")), .TextMatrix(i, .ColIndex("FromTime")))
  If TimDiff < 0 Then
  If RereiveSliceValue(ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 3, NoDay) = False Then
   RereiveSlice ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 3, NoDay
   End If
   X = Abs(DateDiff("n", .TextMatrix(i, .ColIndex("EnterTime")), .TextMatrix(i, .ColIndex("FromTime"))) / 60)
   X = Fix(X)
   .TextMatrix(i, .ColIndex("DelayTimeTime")) = X & ":" & Abs(DateDiff("n", .TextMatrix(i, .ColIndex("EnterTime")), .TextMatrix(i, .ColIndex("FromTime"))) Mod 60)
  .TextMatrix(i, .ColIndex("DelayID")) = NoDay
  .TextMatrix(i, .ColIndex("DelayType")) = 1
   Else
  .TextMatrix(i, .ColIndex("DelayID")) = ""
  .TextMatrix(i, .ColIndex("DelayType")) = ""
End If
If .TextMatrix(i, .ColIndex("OutTime")) = "" Then
.Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexChecked
.TextMatrix(i, .ColIndex("NoRegOut")) = 1
'.TextMatrix(i, .ColIndex("NoRegOutType")) = 1
Else
.Cell(flexcpChecked, i, .ColIndex("NoRegOut")) = flexUnchecked
TimDiff = DateDiff("n", .TextMatrix(i, .ColIndex("OutTime")), .TextMatrix(i, .ColIndex("ToTime")))
  If TimDiff > 0 Then
  If RereiveSliceValue(ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 5, NoDay) = False Then
   RereiveSlice ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 5, NoDay
   End If
    X = Abs(DateDiff("n", .TextMatrix(i, .ColIndex("OutTime")), .TextMatrix(i, .ColIndex("ToTime"))) / 60)
     X = Fix(X)
   .TextMatrix(i, .ColIndex("EarExitTime")) = X & ":" & Abs(DateDiff("n", .TextMatrix(i, .ColIndex("OutTime")), .TextMatrix(i, .ColIndex("ToTime"))) Mod 60)
  .TextMatrix(i, .ColIndex("EarExit")) = NoDay
  .TextMatrix(i, .ColIndex("EarExitType")) = 1
  .TextMatrix(i, .ColIndex("Additional")) = ""
  .TextMatrix(i, .ColIndex("AddiType")) = ""
  .TextMatrix(i, .ColIndex("AdditioTime")) = ""
   ElseIf TimDiff < 0 Then
   If val(.TextMatrix(i, .ColIndex("TypeDay"))) = 0 Then
   If RereiveSliceValue(ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 1, NoDay) = False Then
   RereiveSlice ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 1, NoDay
   End If
   Else
   If RereiveSliceValue(ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 2, NoDay) = False Then
   RereiveSlice ShiftID, val(.TextMatrix(i, .ColIndex("EmpID"))), TimDiff, 2, NoDay
   End If
   End If
       X = Abs(DateDiff("n", .TextMatrix(i, .ColIndex("OutTime")), .TextMatrix(i, .ColIndex("ToTime"))) / 60)
       X = Fix(X)
   .TextMatrix(i, .ColIndex("AdditioTime")) = X & ":" & Abs(DateDiff("n", .TextMatrix(i, .ColIndex("OutTime")), .TextMatrix(i, .ColIndex("ToTime"))) Mod 60)
  .TextMatrix(i, .ColIndex("Additional")) = NoDay
  .TextMatrix(i, .ColIndex("AddiType")) = 1
   If CheckEmployeeNoAddition(val(.TextMatrix(i, .ColIndex("EmpID")))) = True Then
  .TextMatrix(i, .ColIndex("AddiType")) = 2
  .Cell(flexcpBackColor, i, 1, i, 36) = &HC000&
  End If
  .TextMatrix(i, .ColIndex("EarExit")) = ""
  .TextMatrix(i, .ColIndex("EarExitType")) = ""
  .TextMatrix(i, .ColIndex("EarExitTime")) = ""

End If
End If
End If
End If
End If
 If (.TextMatrix(i, .ColIndex("OutTime"))) = "" Then
    ' .Cell(flexcpBackColor, i, 1, i, 34) = .Cell(flexcpBackColor, i, 1, i, 34) = &HC00000
 End If
Next i
End With
End Sub
Function GetMahineCode(Optional Emp_id As Double) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select MachinCode from TblEmployee  where Emp_ID=" & Emp_id & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMahineCode = IIf(IsNull(rs2("MachinCode").value), "", rs2("MachinCode").value)
Else
GetMahineCode = ""
End If
End Function
Function CheckVacotionSick(Optional EmpID As Double, Optional HolDate As Date) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     EmpID, FrmDate, ToDate"
sql = sql & " From dbo.TblRegsterSickleave"
sql = sql & " WHERE     (EmpID = " & EmpID & ") "
sql = sql & " and     (FrmDate <= " & SQLDate(HolDate, True) & ") "
sql = sql & " and     (ToDate >= " & SQLDate(HolDate, True) & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckVacotionSick = True
Else
CheckVacotionSick = False
End If
End Function

Function CheckVacotionAretha(Optional EmpID As Double, Optional HolDate As Date) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     Emp_id, FromDate, ToDate"
sql = sql & " From dbo.TblEmpPassOver"
sql = sql & " WHERE     (Emp_id = " & EmpID & ") "
sql = sql & " and     (FromDate <= " & SQLDate(HolDate, True) & ") "
sql = sql & " and     (ToDate >= " & SQLDate(HolDate, True) & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckVacotionAretha = True
Else
CheckVacotionAretha = False
End If
End Function
Function CheckVacotion(Optional EmpID As Double, Optional HolDate As Date) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     EmpID, stratDate, EndDate"
sql = sql & " From dbo.TblVocationEntitlements"
sql = sql & " WHERE     (EmpID = " & EmpID & ") "
sql = sql & " and     (stratDate <= " & SQLDate(HolDate, True) & ") "
sql = sql & " and     (EndDate >= " & SQLDate(HolDate, True) & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckVacotion = True
Else
CheckVacotion = False
End If
End Function
Function CalcalteAbsence()
Dim i As Integer
Dim k As Integer
Dim j As Integer
FilGridDate
Dim Rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim RsDevsub As ADODB.Recordset
Dim sql As String
With fg
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("RecDate")) <> "" Then
sql = " SELECT     dbo.TblShiftWorker.EmpID"
sql = sql & " FROM         dbo.TblShiftWorker LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblShiftWorker.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " where  (dbo.TblEmployee.WorkState = 1)"
If val(Me.DcbProject1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.project_id  =" & val(DcbProject1.BoundText) & " "
End If
If val(DcbBranch1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.BranchId   =" & val(DcbBranch1.BoundText) & " "
End If
If val(DcpDept1.BoundText) <> 0 Then
sql = sql & " and dbo.TblEmployee.DepartmentID   =" & val(DcpDept1.BoundText) & " "
End If
If val(DcbEmployee1.BoundText) <> 0 Then
sql = sql & " and  dbo.TblShiftWorker.EmpID =" & val(DcbEmployee1.BoundText) & " "
End If

sql = sql & " GROUP BY dbo.TblShiftWorker.EmpID, dbo.TblEmployee.workstate"

Set Rs1 = New ADODB.Recordset
Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs1.RecordCount > 0 Then
Rs1.MoveFirst
For k = 1 To Rs1.RecordCount
sql = " SELECT     dbo.TblShiftWorker.ShiftID, dbo.TblShiftWorker.EmpID, dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID,"
sql = sql & "                       dbo.TblEmployee.MachinCode"
sql = sql & " FROM         dbo.TblShiftWorker LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee ON dbo.TblShiftWorker.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblShiftWorker.EmpID = " & IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value) & ")"
sql = sql & " and dbo.TblShiftWorker.ShiftID not in(select ShiftID from  TblApproveShiftDet where EmpID<>0 and EmpID=" & Rs1("EmpID").value & " and ApprovID=" & val(TxtSerial1.Text) & " and RecordDate=" & SQLDate(IIf((fg.TextMatrix(i, fg.ColIndex("RecDate"))) = "", Null, (fg.TextMatrix(i, fg.ColIndex("RecDate")))), True) & "  )"
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

If rs2.RecordCount > 0 Then
rs2.MoveFirst
For j = 1 To rs2.RecordCount
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblApproveShiftDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not IsNull(rs2("EmpID").value) Then
    If IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value) <> 0 Then
   DTPicker3.value = IIf(fg.TextMatrix(i, fg.ColIndex("RecDate")) = "", "01/01/1999", fg.TextMatrix(i, fg.ColIndex("RecDate")))
    If CheckShiftHolidaies(IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value), Weekday(DTPicker3.value)) = 0 Then
    If CheckHolidaies(DTPicker3.value) = 0 Then
   If CheckVacotion(IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value), DTPicker3.value) = False Then
   If CheckVacotionAretha(IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value), DTPicker3.value) = False Then
   If CheckVacotionSick(IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value), DTPicker3.value) = False Then
                RsDevsub.AddNew
                RsDevsub("ApprovID").value = val(Me.TxtSerial1.Text)
                RsDevsub("EmpID").value = rs2("EmpID").value
                RsDevsub("ShiftID").value = rs2("ShiftID").value
                RsDevsub("RecordDate").value = IIf((fg.TextMatrix(i, fg.ColIndex("RecDate"))) = "", Null, (fg.TextMatrix(i, fg.ColIndex("RecDate"))))
                RsDevsub("Absence").value = 1
                RsDevsub("AbsenType").value = 1
                If GetShiftNo(rs2("EmpID").value) <> 0 Then
                RsDevsub("AbsenceTimeVal").value = 1 / GetShiftNo(rs2("EmpID").value)
                RsDevsub("AbsenceTime").value = 1 / GetShiftNo(rs2("EmpID").value)
                End If
                RsDevsub("MachinCode").value = rs2("MachinCode").value
                RsDevsub("DeptID").value = rs2("DepartmentID").value
                RsDevsub("ProjID").value = rs2("project_id").value
                RsDevsub("BranchID").value = rs2("BranchId").value
                RsDevsub("NotFingPrin").value = 1
                
       RsDevsub.update
       End If
       End If
       End If
       End If
       End If
       End If
       End If
       rs2.MoveNext
    Next j
       
End If
Rs1.MoveNext
Next k
End If
End If
Next i
End With
FillgridAbsence
End Function
Sub FillgridAbsence()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim EmpID As Double
Dim count1 As Double
Dim i As Integer
Dim NoDay As Double
Dim netval As Double
Dim MothID As Integer
Dim ShiftID As Double
Dim sql As String
Dim MonthRec As Double
Dim RecordDate As Date

'Sql = " SELECT     Sum(AbsenceTimeVal) AS SumAbcence, EmpID, MONTH(RecordDate) AS MonthRec"
'Sql = Sql & " From dbo.TblApproveShiftDet"
'Sql = Sql & " WHERE     ((NOT (AbsenceTimeVal IS NULL)) OR"
'Sql = Sql & "                      (AbsenceTimeVal <> 0))and empId<>0"
'Sql = Sql & " GROUP BY EmpID, MONTH(RecordDate)"
sql = " SELECT      EmpID, MONTH(RecordDate) AS MonthRec ,RecordDate"
sql = sql & " From dbo.TblApproveShiftDet"
sql = sql & " WHERE     ((NOT (AbsenceTimeVal IS NULL)) OR"
sql = sql & "                      (AbsenceTimeVal <> 0))and empId<>0"
If val(DcbEmployee1.BoundText) <> 0 And DcbEmployee1.Text <> "" Then
sql = sql & " and EmpID =" & val(DcbEmployee1.BoundText) & ""

End If
sql = sql & " ORDER BY EmpID, MONTH(RecordDate)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With GridInstallments
Rs3.MoveFirst
For i = 1 To Rs3.RecordCount
EmpID = IIf(IsNull(Rs3("EmpID").value), 0, Rs3("EmpID").value)
If i = 1 Then
MonthRec = IIf(IsNull(Rs3("MonthRec").value), 0, Rs3("MonthRec").value)

count1 = 1
Else
If MonthRec = IIf(IsNull(Rs3("MonthRec").value), 0, Rs3("MonthRec").value) And EmpID = IIf(IsNull(Rs3("EmpID").value), 0, Rs3("EmpID").value) Then
count1 = count1 + 1
Else
count1 = 1
End If
End If
'count1 = IIf(IsNull(Rs3("SumAbcence").value), 0, Rs3("SumAbcence").value)
MothID = IIf(IsNull(Rs3("MonthRec").value), 0, Rs3("MonthRec").value)
RecordDate = IIf(IsNull(Rs3("RecordDate").value), Date, Rs3("RecordDate").value)
ShiftID = CheckBetwenPeriod(EmpID)
If RereiveSliceValue(ShiftID, EmpID, count1, 4, NoDay) = False Then
   RereiveSlice ShiftID, EmpID, count1, 4, NoDay
   End If
  ' If GetShiftNo(EmpID) <> 0 Then
  'netval = NoDay / GetShiftNo(EmpID)
  'End If
  If EmpID <> 0 Then
 Cn.Execute "update TblApproveShiftDet set Absence=" & NoDay & " where EmpID=" & EmpID & " and MONTH(RecordDate)=" & MothID & " and not(AbsenceTime is null)and RecordDate=" & SQLDate(RecordDate, True) & " "
  End If
  Rs3.MoveNext
Next i
End With
End If
End Sub
Function GetShiftNo(Optional EmpID As Double) As Double
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "select Count(ShiftID)as contShift from TblShiftWorker where EmpID=" & EmpID & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetShiftNo = IIf(IsNull(Rs3("contShift").value), 0, Rs3("contShift").value)
Else
GetShiftNo = 0
End If
End Function
Function GetNamShift(Optional ShiftID As Double) As String
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     SeftCode, SheftName, SheftNamee"
sql = sql & " From dbo.TbLSheft"
sql = sql & " Where (SeftCode = " & ShiftID & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
GetNamShift = ""
If Rs8.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
GetNamShift = IIf(IsNull(Rs8("SheftName").value), "«ÌÊÃœ ‘› ", Rs8("SheftName").value)
Else
GetNamShift = IIf(IsNull(Rs8("SheftNamee").value), "No Shift", Rs8("SheftNamee").value)
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
GetNamShift = "·«ÌÊÃœ ‘› "
Else
GetNamShift = "No Shift"
End If
End If
End Function
Private Sub ISButton4_Click()
   ' On Error GoTo ErrTrap
 If GridInstallments.Rows = 1 Then Exit Sub
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
    Dim i As Integer
   With GridInstallments
   For i = 1 To .Rows - 1
   If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
   If val(.TextMatrix(i, .ColIndex("EmpID"))) = 0 Or val(.TextMatrix(i, .ColIndex("ShiftID"))) = 0 Or val(.TextMatrix(i, .ColIndex("DeptID"))) = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «· «ﬂÌœ  «ﬂœ „‰ œ·«·«  «·«·Ê«‰"
   Else
   MsgBox "Can not save. Make sure the data"
   End If
   Exit Sub
   End If
   End If
   Next i
   End With
   Dim MonthID As Integer
   Dim YearID As Integer
   With GridInstallments
   For i = 1 To .Rows - 1
   If .TextMatrix(i, .ColIndex("RecordDate")) <> "" Then
       ReDta.value = .TextMatrix(i, .ColIndex("RecordDate"))
    If ChecPeriodSalary(ReDta.value, MonthID, YearID) = True Then
        CboYear.Text = YearID
    CmbMonth.ListIndex = MonthID - 1
    Else
        CboYear.Text = year(ReDta.value)
    CmbMonth.ListIndex = Month(ReDta.value) - 1
    End If
           If ChekPayedSalary(val(CboYear.Text), val(CmbMonth.ListIndex) + 1, val(.TextMatrix(i, .ColIndex("BranchID")))) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ Õ–› ﬁÌœ «·—Ê« »  ··‘Â—  "
            Else
            MsgBox "Delete Salary Allocation JL"
            End If
            Exit Sub
            End If
         End If
     Next i
    End With
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
    ISButton4.Enabled = False
    btnSave.Enabled = True
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

Private Sub Rd_Click(Index As Integer)
If ISButton4.Enabled = False Then
FullGridData
End If

End Sub

Private Sub RdEmp_Click()
If Me.RdEmp.value = False Then
Me.DcbEmployee1.BoundText = ""
End If
End Sub
Function GetMaxAverg(Optional TimDif As Double, Optional TypeTrans As Integer = -1) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "SELECT     MAX(dbo.TblShiftWorker.AverageMaint) AS MaxAverageMaint"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where (dbo.TblShiftWorker.TypeTrans = TypeTrans) And (dbo.TblShiftWorker.FromMint <= " & Abs(TimDif) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetMaxAverg = IIf(IsNull(Rs3("MaxAverageMaint").value), 0, Rs3("MaxAverageMaint").value)
Else
GetMaxAverg = 0
End If
End Function
Function RereiveSliceValue(ShiftID As Double, Optional EmpID As Double = 0, Optional TimDif As Double, Optional TypeTrans As Integer = -1, Optional ByRef NoDay As Double) As Boolean
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = "SELECT     dbo.TblShiftWorker.EmpID, dbo.TblShiftWorker.Typetrans, dbo.TblShiftWorker.FromMint, dbo.TblShiftWorker.ToMint, dbo.TblShiftWorker.AverageMaint ,dbo.TblShiftWorker.Valuee"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where (dbo.TblShiftWorker.TypeTrans = " & TypeTrans & ")"
sql = sql & "  and dbo.TblShiftWorker.FromMint <=" & Abs(TimDif) & ""
sql = sql & "  and dbo.TblShiftWorker.ToMint >=" & Abs(TimDif) & ""
If ShiftID <> 0 Then
sql = sql & " and dbo.TblShiftWorker.ShiftID=" & ShiftID & " and not(dbo.TblShiftWorker.Valuee is null) and dbo.TblShiftWorker.Valuee<>0"
End If
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
NoDay = IIf(IsNull(Rs8("Valuee").value), 0, Rs8("Valuee").value)
RereiveSliceValue = True
Else
RereiveSliceValue = False
NoDay = 0
End If
  End Function
Sub RereiveSlice(ShiftID As Double, Optional EmpID As Double = 0, Optional TimDif As Double, Optional TypeTrans As Integer = -1, Optional ByRef NoDay As Double)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = "SELECT     dbo.TblShiftWorker.EmpID, dbo.TblShiftWorker.Typetrans, dbo.TblShiftWorker.FromMint, dbo.TblShiftWorker.ToMint, dbo.TblShiftWorker.AverageMaint"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where (dbo.TblShiftWorker.TypeTrans = " & TypeTrans & ")"
sql = sql & "  and dbo.TblShiftWorker.FromMint <=" & Abs(TimDif) & ""
sql = sql & "  and dbo.TblShiftWorker.ToMint >=" & Abs(TimDif) & ""
If ShiftID <> 0 Then
sql = sql & " and dbo.TblShiftWorker.ShiftID=" & ShiftID & ""
End If
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
NoDay = IIf(IsNull(Rs8("AverageMaint").value), GetMaxAverg(TimDif, TypeTrans), Rs8("AverageMaint").value)
Else
NoDay = GetMaxAverg(TimDif, TypeTrans)
End If
If TypeTrans <> 4 Then
NoDay = NoDay * Abs(TimDif)
Else
NoDay = NoDay
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
 Dim i As Integer
 If Me.TxtModFlg.Text = "N" Then
 If val(TxtSerial1.Text) <> 0 Then
            With GridInstallments
              For i = 1 To .Rows - 1
                Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
               If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
               End If
             Next i
           End With
         With VSFlexGrid1
              For i = 1 To .Rows - 1
                Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
                Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
             Next i
           End With
           
       StrSQL = "Delete From TblApproveShiftDet Where ApprovID=" & val(Me.TxtSerial1.Text)
             Cn.Execute StrSQL, , adExecuteNoRecords
      StrSQL = "Delete From TblChangedComponentRegister Where ApproveShiftID=" & val(Me.TxtSerial1.Text) & " and Finger =1"
             Cn.Execute StrSQL, , adExecuteNoRecords
      StrSQL = "Delete From TblChangedComponentRegisterDetails Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
             Cn.Execute StrSQL, , adExecuteNoRecords
      StrSQL = "Delete From TblInforVacatiom Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
             Cn.Execute StrSQL, , adExecuteNoRecords
      RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
                
 End If
 End If
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
Function CheckBetwenPeriod(Optional EmpID1 As Double) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TblShiftWorker.EmpID AS EmpID1, dbo.TbLSheft.*"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " Where dbo.TblShiftWorker.EmpID=" & EmpID1 & ""

Rs8.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckBetwenPeriod = IIf(IsNull(Rs8("SeftCode").value), 0, (Rs8("SeftCode").value))
Else
CheckBetwenPeriod = 0
End If
End Function
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
 With GridInstallments
For i = 1 To .Rows - 1
Cn.Execute " update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
  Cn.Execute " update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
Next i
End With
With VSFlexGrid1
For i = 1 To .Rows - 1
 Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
 Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
Next i
End With
    StrSQL = "Delete From TblChangedComponentRegister Where ApproveShiftID=" & val(Me.TxtSerial1.Text) & " and Finger =1"
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
          StrSQL = "Delete From TblInforVacatiom Where ApproveShiftID=" & val(Me.TxtSerial1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
          StrSQL = "Delete From TblApproveShiftDet Where   ApprovID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
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
btnSave.Enabled = False
ISButton4.Enabled = False
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
       ' Me.btnSave.Enabled = True
       ISButton4.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
    ISButton4.Enabled = False
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
       ' Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
   ISButton4.Enabled = True
  ' XPDtbTrans.Enabled = True
  '     Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
     '   Me.btnSave.Enabled = True
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
        ISButton4.Enabled = False
        btnSave.Enabled = True
        GridInstallments.Rows = GridInstallments.Rows + 1
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
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
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
    Me.DCboUserName.BoundText = user_id
    XPDtbTrans_Change
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
  C1Tab1.Caption = "Data"
lbl(4).Caption = "No"
lbl(1).Caption = "Date"
Label1(2).Caption = "Approve Data"
Cmd(3).Caption = "Delete"
Cmd(4).Caption = "Delete All"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"
ELe(0).Caption = "Sort By"
Rd(0).Caption = "Employee"
Rd(1).Caption = "Date"
Rd(2).Caption = "All"

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
lbl(1).Caption = "Date"
lbl(2).Caption = "From Date"
lbl(5).Caption = "To Date"
SelectDept.RightToLeft = False
SelectDept.Caption = "Management"
SelectProject.RightToLeft = False
SelectProject.Caption = "Project"
ISButton2.Caption = "Add"
SelectBranch.RightToLeft = False
SelectBranch.Caption = "Branch"
RdAll.RightToLeft = False
RdAll.Caption = "All"
RdEmp.RightToLeft = False
lbl(9).Caption = "Color"
lbl(10).Caption = "No In File Employee"
lbl(12).Caption = "No In Department"
lbl(13).Caption = "No In Shifts"
lbl(14).Caption = "No Log Out"
lbl(15).Caption = "No Additional"
RdEmp.Caption = "Select Employee"
lbl(0).Caption = "Data"
  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("selected")) = "Select"
  .TextMatrix(0, .ColIndex("MachinCode")) = "Machin Code"
  .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("EnterTime")) = "Time Entry"
  .TextMatrix(0, .ColIndex("Emp_Name")) = " Name "
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("SheftName")) = "Sheft Name"
  .TextMatrix(0, .ColIndex("FromTime")) = "From Time"
  .TextMatrix(0, .ColIndex("OutTime")) = "Time Out"
  .TextMatrix(0, .ColIndex("ToTime")) = "To Time"
  .TextMatrix(0, .ColIndex("DelayTimeTime")) = "Delay"
  .TextMatrix(0, .ColIndex("DelayType")) = "Procedure"
  .TextMatrix(0, .ColIndex("EarExitTime")) = "Early Exit"
  .TextMatrix(0, .ColIndex("EarExitType")) = "Procedure"
  .TextMatrix(0, .ColIndex("Absence")) = "Absence "
  .TextMatrix(0, .ColIndex("AbsenType")) = "Procedure"
  .TextMatrix(0, .ColIndex("AdditioTime")) = "Additional "
  .TextMatrix(0, .ColIndex("AddiType")) = "Procedure"
   .TextMatrix(0, .ColIndex("NoRegOut")) = "Additional "
  .TextMatrix(0, .ColIndex("NoRegOutType")) = "Not Logged Out"
  End With
    With Me.VSFlexGrid1
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("selected")) = "Select"
  .TextMatrix(0, .ColIndex("MachinCode")) = "Machin Code"
  .TextMatrix(0, .ColIndex("FullCode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("EnterTime")) = "Time Entry"
  .TextMatrix(0, .ColIndex("Emp_Name")) = " Name "
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("SheftName")) = "Sheft Name"
  .TextMatrix(0, .ColIndex("FromTime")) = "From Time"
  .TextMatrix(0, .ColIndex("OutTime")) = "Time Out"
  .TextMatrix(0, .ColIndex("ToTime")) = "To Time"
  .TextMatrix(0, .ColIndex("DelayTimeTime")) = "Delay"
  .TextMatrix(0, .ColIndex("DelayType")) = "Procedure"
  .TextMatrix(0, .ColIndex("EarExitTime")) = "Early Exit"
  .TextMatrix(0, .ColIndex("EarExitType")) = "Procedure"
  .TextMatrix(0, .ColIndex("Absence")) = "Absence "
  .TextMatrix(0, .ColIndex("AbsenType")) = "Procedure"
  .TextMatrix(0, .ColIndex("AdditioTime")) = "Additional "
  .TextMatrix(0, .ColIndex("AddiType")) = "Procedure"
   .TextMatrix(0, .ColIndex("NoRegOut")) = "Additional "
  .TextMatrix(0, .ColIndex("NoRegOutType")) = "Not Logged Out"
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblApproveShift"
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
Dim i As Integer
With GridInstallments
For i = 1 To .Rows - 1
Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(i, .ColIndex("IDImport"))) & "  "
   If .TextMatrix(i, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(i, .ColIndex("IDImport2"))) & "  "
                End If
Next i
End With
 GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
'    ReLineGrid
End Sub
Private Sub RemoveGridRow()
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        Cn.Execute "update TblImportShiftsDet set IsSele=null where id =" & val(.TextMatrix(.Row, .ColIndex("IDImport"))) & "  "
           If .TextMatrix(.Row, .ColIndex("OutTime")) <> "" Then
                Cn.Execute "update TblImportShiftsDet set IsSele=Null where id =" & val(.TextMatrix(.Row, .ColIndex("IDImport2"))) & "  "
                End If
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub


Private Sub XPDtbTrans_Change()
    On Error Resume Next

End Sub
