VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmAccidentReport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   Icon            =   "FrmAccidentReport.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9825
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13500
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
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmAccidentReport.frx":6852
      Left            =   15480
      List            =   "FrmAccidentReport.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   17
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
      TabIndex        =   18
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
            Picture         =   "FrmAccidentReport.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccidentReport.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   19
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
      ButtonImage     =   "FrmAccidentReport.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   21
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
      ButtonImage     =   "FrmAccidentReport.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   22
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
      ButtonImage     =   "FrmAccidentReport.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9825
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   13500
      _cx             =   23813
      _cy             =   17330
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
         Left            =   13545
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   11865
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   465
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   13560
         _cx             =   23918
         _cy             =   820
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
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   10710
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   105
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   270
            Left            =   8085
            TabIndex        =   9
            Top             =   105
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   476
            _Version        =   393216
            Format          =   238485505
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   105
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   240
            Index           =   11
            Left            =   6225
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   105
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   240
            Index           =   25
            Left            =   9585
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   105
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   225
            Index           =   4
            Left            =   12615
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   105
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   870
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   9000
         Width           =   13470
         _cx             =   23760
         _cy             =   1535
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
            Height          =   270
            Left            =   11955
            TabIndex        =   27
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   465
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAccidentReport.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   270
            Left            =   10395
            TabIndex        =   28
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   465
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAccidentReport.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   270
            Left            =   8730
            TabIndex        =   11
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAccidentReport.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   270
            Left            =   6975
            TabIndex        =   29
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAccidentReport.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   270
            Left            =   5505
            TabIndex        =   30
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   465
            Width           =   1455
            _ExtentX        =   2566
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
            ButtonImage     =   "FrmAccidentReport.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   330
            Left            =   4590
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   465
            Width           =   1020
            _ExtentX        =   1799
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
            ButtonImage     =   "FrmAccidentReport.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   270
            Left            =   1560
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   465
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAccidentReport.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   270
            Left            =   0
            TabIndex        =   33
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   465
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   476
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
            ButtonImage     =   "FrmAccidentReport.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8430
            TabIndex        =   34
            Top             =   75
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   330
            Left            =   2880
            TabIndex        =   170
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   465
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·„’—Ê›« "
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
            ButtonImage     =   "FrmAccidentReport.frx":2A8D1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   315
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   75
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   2370
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   75
            Width           =   780
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   270
            Index           =   1
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   75
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   270
            Index           =   0
            Left            =   3255
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   75
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   14
            Left            =   12375
            TabIndex        =   35
            Top             =   75
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   585
         Index           =   18
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   0
         Width           =   13545
         _cx             =   23892
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
            Height          =   240
            Left            =   135
            TabIndex        =   44
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   423
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
            ButtonImage     =   "FrmAccidentReport.frx":31133
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   240
            Left            =   675
            TabIndex        =   45
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   423
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
            ButtonImage     =   "FrmAccidentReport.frx":314CD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   240
            Left            =   1365
            TabIndex        =   46
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   423
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
            ButtonImage     =   "FrmAccidentReport.frx":31867
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   240
            Left            =   1965
            TabIndex        =   47
            Top             =   180
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   423
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
            ButtonImage     =   "FrmAccidentReport.frx":31C01
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   167
            Top             =   0
            Width           =   3855
         End
         Begin VB.Image Image1 
            Height          =   465
            Left            =   12465
            Picture         =   "FrmAccidentReport.frx":31F9B
            Stretch         =   -1  'True
            Top             =   90
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰„Ê–Ã  ﬁ—Ì— Õ«œÀ"
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
            Height          =   285
            Index           =   2
            Left            =   7620
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   60
            Width           =   4695
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8025
         Left            =   0
         TabIndex        =   51
         Top             =   960
         Width           =   13665
         _cx             =   24104
         _cy             =   14155
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
         Caption         =   "»Ì«‰«  «”«”Ì…|«·„’—Ê›« |«·„—›ﬁ« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   7605
            Left            =   45
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   45
            Width           =   13575
            _cx             =   23945
            _cy             =   13414
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   465
               Left            =   0
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   120
               Width           =   13455
               _cx             =   23733
               _cy             =   820
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
               Begin VB.TextBox TxtPlace 
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
                  Height          =   270
                  Left            =   5115
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   54
                  Top             =   105
                  Width           =   7200
               End
               Begin MSComCtl2.DTPicker AccDate 
                  Height          =   270
                  Left            =   120
                  TabIndex        =   55
                  Top             =   105
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   476
                  _Version        =   393216
                  Format          =   244252673
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker AccTime 
                  Height          =   270
                  Left            =   2715
                  TabIndex        =   56
                  Top             =   105
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   476
                  _Version        =   393216
                  Format          =   197787650
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Êﬁ "
                  Height          =   255
                  Index           =   1
                  Left            =   4140
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   75
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   255
                  Index           =   0
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   75
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   165
                  Index           =   1
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   825
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Êﬁ⁄"
                  Height          =   285
                  Index           =   0
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   75
                  Width           =   1245
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   1905
               Left            =   0
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   840
               Width           =   13455
               _cx             =   23733
               _cy             =   3360
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
               Begin VB.TextBox TxtRepairCost 
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
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1230
                  Width           =   2130
               End
               Begin VB.TextBox TxtThidPayment 
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
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   1230
                  Width           =   1890
               End
               Begin VB.TextBox TxtDimagVehicle 
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
                  Left            =   9915
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   71
                  Top             =   1200
                  Width           =   2130
               End
               Begin VB.TextBox TxtDimThiParInvol 
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
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   70
                  Top             =   840
                  Width           =   4170
               End
               Begin VB.TextBox TxtYearMake 
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
                  Left            =   7200
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1200
                  Width           =   1410
               End
               Begin VB.TextBox TxtModel 
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
                  Left            =   7200
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   870
                  Width           =   1410
               End
               Begin VB.TextBox TxtVehicleType 
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
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   510
                  Width           =   2130
               End
               Begin VB.TextBox TxtFleetNo 
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
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   510
                  Width           =   1890
               End
               Begin VB.TextBox TxtPlateNo 
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
                  Left            =   7200
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   510
                  Width           =   1410
               End
               Begin VB.TextBox TxtNameDriver 
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
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   150
                  Width           =   5490
               End
               Begin VB.TextBox TxtStaffNo 
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
                  Left            =   7200
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   120
                  Width           =   1410
               End
               Begin VB.TextBox TxtJobTitle 
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
                  Left            =   9915
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   120
                  Width           =   2130
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   315
                  Left            =   4320
                  TabIndex        =   74
                  TabStop         =   0   'False
                  Top             =   840
                  Width           =   1410
                  _cx             =   2487
                  _cy             =   556
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
                  Appearance      =   0
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
                  Begin XtremeSuiteControls.RadioButton ThirdPartyInvol 
                     Height          =   255
                     Index           =   0
                     Left            =   720
                     TabIndex        =   75
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·«"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ThirdPartyInvol 
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   76
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "‰⁄„"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   195
                     Index           =   4
                     Left            =   12120
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   960
                     Width           =   1245
                  End
               End
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   195
                  Left            =   9720
                  TabIndex        =   78
                  ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
                  Top             =   1560
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   344
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "«·’Ê—"
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
                  ButtonImage     =   "FrmAccidentReport.frx":333A0
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic5 
                  Height          =   315
                  Left            =   10680
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1410
                  _cx             =   2487
                  _cy             =   556
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
                  Appearance      =   0
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
                  Begin XtremeSuiteControls.RadioButton IsPhoto 
                     Height          =   255
                     Index           =   0
                     Left            =   840
                     TabIndex        =   80
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·«"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton IsPhoto 
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   81
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "‰⁄„"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   195
                     Index           =   19
                     Left            =   12120
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   960
                     Width           =   1245
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic6 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   83
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1290
                  _cx             =   2275
                  _cy             =   556
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
                  Appearance      =   0
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
                  Begin XtremeSuiteControls.RadioButton IsNdliab 
                     Height          =   255
                     Index           =   0
                     Left            =   720
                     TabIndex        =   84
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·«"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton IsNdliab 
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   85
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "‰⁄„"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   195
                     Index           =   21
                     Left            =   12120
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   960
                     Width           =   1245
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   87
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1290
                  _cx             =   2275
                  _cy             =   556
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
                  Appearance      =   0
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
                  Begin XtremeSuiteControls.RadioButton IsPolice 
                     Height          =   255
                     Index           =   0
                     Left            =   720
                     TabIndex        =   88
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·«"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton IsPolice 
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   89
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "‰⁄„"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   195
                     Index           =   23
                     Left            =   12120
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Top             =   960
                     Width           =   1245
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic13 
                  Height          =   315
                  Left            =   8700
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   3345
                  _cx             =   5900
                  _cy             =   556
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Begin VB.TextBox txtNum4 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   0
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   8
                     Top             =   0
                     Width           =   450
                  End
                  Begin VB.TextBox txtLetter4 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   1740
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   4
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.TextBox txtNum3 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   420
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   7
                     Top             =   0
                     Width           =   435
                  End
                  Begin VB.TextBox txtNum2 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   750
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   0
                     Width           =   495
                  End
                  Begin VB.TextBox txtNum1 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   1215
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   5
                     Top             =   0
                     Width           =   525
                  End
                  Begin VB.TextBox txtLetter3 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   2160
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   3
                     Top             =   0
                     Width           =   375
                  End
                  Begin VB.TextBox txtLetter2 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   2565
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   2
                     Top             =   0
                     Width           =   405
                  End
                  Begin VB.TextBox txtLetter1 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   2925
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   1
                     Top             =   0
                     Width           =   450
                  End
               End
               Begin MSDataListLib.DataCombo DcbCrID 
                  Height          =   315
                  Left            =   9915
                  TabIndex        =   168
                  Top             =   840
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   195
                  Index           =   47
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   840
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â· ÌÊÃœ  ﬁ—Ì— ··‘—ÿ…"
                  Height          =   195
                  Index           =   22
                  Left            =   5595
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   1560
                  Width           =   1605
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " Õ„· „”ƒÊ·Ì…"
                  Height          =   195
                  Index           =   20
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   1560
                  Width           =   1125
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â· Â‰«ﬂ ’Ê—"
                  Height          =   195
                  Index           =   18
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   1560
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " ﬂ·›… «·«’·«Õ"
                  Height          =   195
                  Index           =   17
                  Left            =   2235
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   1200
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„»·€  ﬁœÌ—Ì ··ÿ—›"
                  Height          =   195
                  Index           =   16
                  Left            =   5595
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   1200
                  Width           =   1485
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ﬁœ«— «·÷—— ··”Ì«—…"
                  Height          =   195
                  Index           =   15
                  Left            =   12030
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰… «·’‰⁄"
                  Height          =   195
                  Index           =   12
                  Left            =   8595
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   1200
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â· ÌÊÃœÿ—› «Œ—"
                  Height          =   315
                  Index           =   11
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   840
                  Width           =   1410
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ÊœÌ·"
                  Height          =   195
                  Index           =   10
                  Left            =   8595
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   840
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   195
                  Index           =   9
                  Left            =   2235
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   480
                  Width           =   1485
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·«”ÿÊ·"
                  Height          =   195
                  Index           =   8
                  Left            =   5595
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   480
                  Width           =   1485
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   195
                  Index           =   7
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   480
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ ﬁ«∆œ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   195
                  Index           =   6
                  Left            =   5595
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   120
                  Width           =   1485
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—ﬁ„ «·ÊŸÌ›Ì"
                  Height          =   195
                  Index           =   3
                  Left            =   8595
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   90
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„”„Ï «·ÊŸÌ›Ì"
                  Height          =   195
                  Index           =   5
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   90
                  Width           =   1245
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   1620
               Left            =   0
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   3000
               Width           =   13455
               _cx             =   23733
               _cy             =   2858
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
               Begin VB.TextBox TxtDesInjuried 
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
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   113
                  Top             =   1200
                  Width           =   8805
               End
               Begin VB.TextBox TxtStaffNo2 
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
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   112
                  Top             =   840
                  Width           =   2130
               End
               Begin VB.TextBox TxtJobTitle2 
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
                  Left            =   3240
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   111
                  Top             =   840
                  Width           =   2130
               End
               Begin VB.TextBox TxtTheirName 
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
                  Height          =   555
                  Left            =   6795
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   110
                  Top             =   120
                  Width           =   5370
               End
               Begin VB.TextBox TxtDepatment 
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
                  Left            =   6795
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   109
                  Top             =   840
                  Width           =   2130
               End
               Begin VB.TextBox TxtHowConected 
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
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   108
                  Top             =   120
                  Width           =   5250
               End
               Begin XtremeSuiteControls.CheckBox IsEmployee 
                  Height          =   255
                  Left            =   10680
                  TabIndex        =   114
                  Top             =   840
                  Width           =   735
                  _Version        =   786432
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„ÊŸ›"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox IsOther 
                  Height          =   255
                  Left            =   9840
                  TabIndex        =   115
                  Top             =   840
                  Width           =   735
                  _Version        =   786432
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«Œ—Ì‰"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox IsContrct 
                  Height          =   255
                  Left            =   11430
                  TabIndex        =   116
                  Top             =   840
                  Width           =   735
                  _Version        =   786432
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„ﬁ«Ê·"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DateComm 
                  Height          =   315
                  Left            =   10710
                  TabIndex        =   117
                  Top             =   1200
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   244514817
                  CurrentDate     =   38784
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ê’› «·«’«»…"
                  Height          =   195
                  Index           =   31
                  Left            =   9120
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   1200
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «·«»·«€"
                  Height          =   195
                  Index           =   30
                  Left            =   12240
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   1200
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—ﬁ„ «·ÊŸÌ›Ì"
                  Height          =   195
                  Index           =   29
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   840
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„”„Ï «·ÊŸÌ›Ì"
                  Height          =   195
                  Index           =   28
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   840
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«œ«—…"
                  Height          =   195
                  Index           =   24
                  Left            =   8715
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   840
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â· Â„"
                  Height          =   195
                  Index           =   27
                  Left            =   12360
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   840
                  Width           =   1125
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„«¡ «·„’«»Ì‰"
                  Height          =   195
                  Index           =   26
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÿ—Ìﬁ… «· Ê«’·"
                  Height          =   195
                  Index           =   25
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   240
                  Width           =   1485
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic10 
               Height          =   1905
               Left            =   0
               TabIndex        =   126
               TabStop         =   0   'False
               Top             =   4920
               Width           =   13455
               _cx             =   23733
               _cy             =   3360
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
               Begin VB.TextBox TxtActionTaken 
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
                  Height          =   495
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   132
                  Top             =   1320
                  Width           =   4770
               End
               Begin VB.TextBox TxtDesHospital 
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
                  Height          =   495
                  Left            =   6720
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   131
                  Top             =   1320
                  Width           =   5250
               End
               Begin VB.TextBox TxtFirstAid 
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
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   130
                  Top             =   720
                  Width           =   3210
               End
               Begin VB.TextBox TxtNameAddress 
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
                  Height          =   555
                  Left            =   6720
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   129
                  Top             =   720
                  Width           =   5250
               End
               Begin VB.TextBox TxtDesWhyHappen 
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
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   128
                  Top             =   120
                  Width           =   4770
               End
               Begin VB.TextBox TxtDesIncident 
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
                  Height          =   555
                  Left            =   6720
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   127
                  Top             =   120
                  Width           =   5250
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic11 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   133
                  TabStop         =   0   'False
                  Top             =   840
                  Width           =   1410
                  _cx             =   2487
                  _cy             =   556
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
                  Appearance      =   0
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
                  Begin XtremeSuiteControls.RadioButton IsFirstAid 
                     Height          =   255
                     Index           =   0
                     Left            =   840
                     TabIndex        =   134
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·«"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton IsFirstAid 
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   135
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "‰⁄„"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   195
                     Index           =   36
                     Left            =   12120
                     RightToLeft     =   -1  'True
                     TabIndex        =   136
                     Top             =   960
                     Width           =   1245
                  End
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã—« ¡ «·„ Œ– ·„‰⁄  ﬂ—«— «·Õ«œÀ „” ﬁ»·«"
                  Height          =   555
                  Index           =   38
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   1320
                  Width           =   1845
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " ›«’Ì· «·„‘›Ï"
                  Height          =   195
                  Index           =   37
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   1440
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â·  „ ⁄„· «”⁄«›«  «Ê·Ì…"
                  Height          =   195
                  Index           =   35
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   840
                  Width           =   1845
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„⁄·Ê„«  «·‘ÂÊœ"
                  Height          =   195
                  Index           =   34
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   810
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—«Ìﬂ ›Ì ”»» ÊﬁÊ⁄ «·Õ«œÀ"
                  Height          =   435
                  Index           =   32
                  Left            =   4950
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   90
                  Width           =   1725
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„⁄·Ê„«  «·Õ«œÀ"
                  Height          =   195
                  Index           =   33
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   210
                  Width           =   1245
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic14 
               Height          =   540
               Left            =   0
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   7080
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
               Begin VB.TextBox TxtDepartment2 
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   150
                  Width           =   2610
               End
               Begin VB.TextBox TxtPosition 
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
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   150
                  Width           =   2610
               End
               Begin VB.TextBox TxtName1 
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
                  Left            =   9120
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   120
                  Width           =   3330
               End
               Begin MSComCtl2.DTPicker DateMaking 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   147
                  Top             =   150
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   244514817
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   195
                  Index           =   2
                  Left            =   1500
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   210
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«œ«—…"
                  Height          =   195
                  Index           =   42
                  Left            =   4830
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   210
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Êﬁ⁄"
                  Height          =   195
                  Index           =   41
                  Left            =   8310
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   210
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«”„"
                  Height          =   195
                  Index           =   40
                  Left            =   12270
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   210
                  Width           =   1125
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   195
                  Index           =   39
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   960
                  Width           =   1245
               End
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„⁄·Ê„«  Õ«œÀ ”Ì«—…"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   46
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   600
               Width           =   2565
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„⁄·Ê„«  «·«‘Œ«’ «·„’«»Ì‰"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   45
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   2760
               Width           =   2565
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„⁄·Ê„«  ⁄‰ «·Õ«œÀ Ê«·«Ã—«¡ «·„ »⁄"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   44
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   4680
               Width           =   2565
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·‘Œ’ «·–Ì «⁄œ «· ﬁ—Ì—"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   43
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   6840
               Width           =   2565
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic15 
            Height          =   7605
            Left            =   14310
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   45
            Width           =   13575
            _cx             =   23945
            _cy             =   13414
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic17 
               Height          =   4545
               Left            =   0
               TabIndex        =   158
               TabStop         =   0   'False
               Top             =   120
               Width           =   13455
               _cx             =   23733
               _cy             =   8017
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
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3570
                  Left            =   120
                  TabIndex        =   160
                  Top             =   600
                  Width           =   13275
                  _cx             =   23416
                  _cy             =   6297
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
                  Rows            =   2
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAccidentReport.frx":39C02
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
               Begin ImpulseButton.ISButton Cmd1 
                  Height          =   270
                  Index           =   0
                  Left            =   12480
                  TabIndex        =   161
                  Top             =   4200
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
                  ButtonImage     =   "FrmAccidentReport.frx":39CD7
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd1 
                  Height          =   270
                  Index           =   1
                  Left            =   10680
                  TabIndex        =   162
                  Top             =   4200
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
                  ButtonImage     =   "FrmAccidentReport.frx":3A271
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label CompValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  Height          =   195
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   4200
                  Width           =   2325
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì «·⁄«„"
                  Height          =   195
                  Index           =   48
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   4200
                  Width           =   1485
               End
               Begin VB.Label TotalValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  Height          =   195
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   4200
                  Width           =   2325
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì  Õ„· «·‘—ﬂ…"
                  Height          =   195
                  Index           =   14
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   4200
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„’—Ê›« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   360
                  Index           =   3
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   120
                  Width           =   2160
               End
            End
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
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmAccidentReport"
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

Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblAccidentReport", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Sub FullGrid()
Dim i As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
 Grid.Clear flexClearScrollable, flexClearEverything
 Grid.rows = Grid.FixedRows + 1
sql = " SELECT     dbo.TblAccidentReportDet.ID, dbo.TblAccidentReportDet.AccID, dbo.TblAccidentReportDet.Valuee, dbo.TblAccidentReportDet.Typ, dbo.TblAccidentReportDet.Remarks,"
sql = sql & "                       dbo.TblAccidentReportDet.ExpID , dbo.TblDataTypeExchange.name, dbo.TblDataTypeExchange.NameE"
sql = sql & "  FROM         dbo.TblAccidentReportDet LEFT OUTER JOIN"
sql = sql & "                       dbo.TblDataTypeExchange ON dbo.TblAccidentReportDet.ExpID = dbo.TblDataTypeExchange.Id"
sql = sql & "  Where (dbo.TblAccidentReportDet.AccID = " & val(TxtSerial1.text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Grid
.rows = Rs3.RecordCount + 1
Rs3.MoveFirst
For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("ExpID")) = IIf(IsNull(Rs3("ExpID").value), 0, Rs3("ExpID").value)
.TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(Rs3("Valuee").value), 0, Rs3("Valuee").value)
.TextMatrix(i, .ColIndex("Typ")) = IIf(IsNull(Rs3("Typ").value), "", Rs3("Typ").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs3("Remarks").value), "", Rs3("Remarks").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("name").value), "", Rs3("name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("namee").value), "", Rs3("namee").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub

Sub GetCarInformation()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     dbo.TblCarsData.id, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Job, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
sql = sql & "                       dbo.TblCarsData.Model AS YearModel, dbo.TblCarsData.VModel, dbo.TblCarModels.Model, dbo.TblCarModels.ModelE, dbo.TblCarsData.Emp_id,"
sql = sql & "                       dbo.TblEmployee.emp_name , dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee"
sql = sql & "  FROM         dbo.TblCarsData LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCarModels ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
sql = sql & "                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id"
' Sql = Sql & "  WHERE     (dbo.TblCarsData.BoardNO = N'" & TxtPlateNo.Text & "')"
sql = sql & " WHERE REPLACE(dbo.TblCarsData.BoardNO, ' ', '')= '" & Replace(TxtPlateNo, " ", "") & "'"

Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
TxtStaffNo.text = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
TxtYearMake.text = IIf(IsNull(Rs3("YearModel").value), "", Rs3("YearModel").value)
TXTJobTitle.text = IIf(IsNull(Rs3("Job").value), "", Rs3("Job").value)
If SystemOptions.UserInterface = ArabicInterface Then
TxtVehicleType.text = IIf(IsNull(Rs3("name").value), "", Rs3("name").value)
TxtModel.text = IIf(IsNull(Rs3("Model").value), "", Rs3("Model").value)
TxtNameDriver.text = IIf(IsNull(Rs3("Emp_Name").value), "", Rs3("Emp_Name").value)
Else
TxtVehicleType.text = IIf(IsNull(Rs3("namee").value), "", Rs3("namee").value)
TxtModel.text = IIf(IsNull(Rs3("ModelE").value), "", Rs3("ModelE").value)
TxtNameDriver.text = IIf(IsNull(Rs3("Emp_Namee").value), "", Rs3("Emp_Namee").value)
End If
Else
TxtNameDriver.text = ""
TxtModel.text = ""
TxtVehicleType.text = ""
TxtStaffNo.text = ""
TxtYearMake.text = ""
TXTJobTitle.text = ""
End If
End Sub


Private Sub RemoveGridRow()
    With Me.Grid
        If .Row <= 0 Then
                .rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub

Private Sub Cmd1_Click(Index As Integer)
If Me.TxtModFlg.text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
Case 1
Grid.Clear flexClearScrollable, flexClearEverything
      Grid.rows = 2
End Select
End If
Relin
End Sub

Private Sub DcbCrID_Change()
DcbCrID_Click (0)
End Sub

Private Sub DcbCrID_Click(Area As Integer)
TxtPlateNo.text = Me.DcbCrID.text
End Sub

Private Sub DcbCrID_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "FrmAccidentReport2"
        FrmCasrShearches.show vbModal
    End If
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
        With Grid
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("Typ")) = "#1;«·‘—ﬂ…|#2;«·„ÊŸ› |#3;«·ÿ—› «·À«·À "
             ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("Typ")) = "#1;Company|#2;Employee|#3;Third"
            End If
    End With
    
    conection = "select * from  TblAccidentReport  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
     Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetBordNo Me.DcbCrID
   
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
    Dim i As Integer
    Dim k As Integer
    If Me.TxtModFlg.text = "E" Then
    Cn.Execute "Delete from TblAccidentReportDet where AccID=" & val(TxtSerial1.text) & " "
    End If
   RsSavRec.Fields("RecordDate").value = RecordDate.value
    RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("CarID").value = val(Me.DcbCrID.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("Place").value = Me.TxtPlace.text
   RsSavRec.Fields("AccDate").value = AccDate.value
   RsSavRec.Fields("AccTime").value = FormatDateTime(AccTime.value, vbShortTime)
   RsSavRec.Fields("JobTitle").value = Me.TXTJobTitle.text
   RsSavRec.Fields("StaffNo").value = Me.TxtStaffNo.text
   RsSavRec.Fields("NameDriver").value = Me.TxtNameDriver.text
   RsSavRec.Fields("PlateNo").value = Me.TxtPlateNo.text
   RsSavRec.Fields("FleetNo").value = Me.TxtFleetNo.text
   RsSavRec.Fields("VehicleType").value = Me.TxtVehicleType.text
   RsSavRec.Fields("Model").value = Me.TxtModel.text
   RsSavRec.Fields("YearMake").value = Me.TxtYearMake.text
   If ThirdPartyInvol(1).value = True Then
   RsSavRec.Fields("ThirdPartyInvol").value = 1
   Else
   RsSavRec.Fields("ThirdPartyInvol").value = 0
   End If
   RsSavRec.Fields("DimThiParInvol").value = Me.TxtDimThiParInvol.text
   RsSavRec.Fields("DimagVehicle").value = Me.TxtDimagVehicle.text
   RsSavRec.Fields("ThidPayment").value = val(TxtThidPayment.text)
   RsSavRec.Fields("RepairCost").value = val(TxtRepairCost.text)
   If IsPhoto(1).value = True Then
   RsSavRec.Fields("IsPhoto").value = 1
   Else
   RsSavRec.Fields("IsPhoto").value = 0
   End If
   If IsNdliab(1).value = True Then
   RsSavRec.Fields("IsNdliab").value = 1
   Else
   RsSavRec.Fields("IsNdliab").value = 0
   End If
   If IsPolice(1).value = True Then
   RsSavRec.Fields("IsPolice").value = 1
   Else
   RsSavRec.Fields("IsPolice").value = 0
   End If
   RsSavRec.Fields("HowConected").value = Me.TxtHowConected.text
   RsSavRec.Fields("TheirName").value = Me.TxtTheirName.text
   If IsContrct.value = vbChecked Then
   RsSavRec.Fields("IsContrct").value = 1
   Else
   RsSavRec.Fields("IsContrct").value = 0
   End If
   If IsEmployee.value = vbChecked Then
   RsSavRec.Fields("IsEmployee").value = 1
   Else
   RsSavRec.Fields("IsEmployee").value = 0
   End If
   If IsOther.value = vbChecked Then
   RsSavRec.Fields("IsOther").value = 1
   Else
   RsSavRec.Fields("IsOther").value = 0
   End If
   RsSavRec.Fields("Depatment").value = Me.TxtDepatment.text
   RsSavRec.Fields("JobTitle2").value = Me.TxtJobTitle2.text
   RsSavRec.Fields("StaffNo2").value = Me.TxtStaffNo2.text
   RsSavRec.Fields("DateComm").value = DateComm.value
   RsSavRec.Fields("DesInjuried").value = Me.TxtDesInjuried.text
   RsSavRec.Fields("DesIncident").value = Me.TxtDesIncident.text
   RsSavRec.Fields("DesWhyHappen").value = Me.TxtDesWhyHappen.text
   RsSavRec.Fields("NameAddress").value = Me.TxtDepatment.text
   RsSavRec.Fields("NameAddress").value = Me.TxtNameAddress.text
    If IsFirstAid(1).value = True Then
   RsSavRec.Fields("IsFirstAid").value = 1
   Else
   RsSavRec.Fields("IsFirstAid").value = 0
   End If
   RsSavRec.Fields("FirstAid").value = Me.TxtFirstAid.text
   RsSavRec.Fields("DesHospital").value = Me.TxtDesHospital.text
   RsSavRec.Fields("ActionTaken").value = Me.TxtActionTaken.text
   RsSavRec.Fields("DateMaking").value = Me.DateMaking.value
   RsSavRec.Fields("Name").value = Me.TxtName1.text
   RsSavRec.Fields("Position").value = Me.TxtPosition.text
   RsSavRec.Fields("Department2").value = Me.TxtDepartment2.text
   RsSavRec.Fields("TotalValue").value = val(TotalValue.Caption)
   RsSavRec.Fields("CompValue").value = val(CompValue.Caption)
   RsSavRec.update
  ''//////////////////////////
   
        Dim RsDet As ADODB.Recordset
        Set RsDet = New ADODB.Recordset
        StrSQL = " select * from TblAccidentReportDet  where 1 = -1 "
        RsDet.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Grid
        For i = 1 To .rows - 1
           If val(.TextMatrix(i, .ColIndex("ExpID"))) <> 0 Then
                    RsDet.AddNew
                    RsDet("AccID") = val(TxtSerial1.text)
                    RsDet("ExpID") = val(.TextMatrix(i, .ColIndex("ExpID")))
                    RsDet("Valuee") = val(.TextMatrix(i, .ColIndex("Valuee")))
                    RsDet("Remarks") = .TextMatrix(i, .ColIndex("Remarks"))
                    RsDet("Typ") = IIf((val(.TextMatrix(i, .ColIndex("Typ")))) = 0, Null, val(.TextMatrix(i, .ColIndex("Typ"))))
                    RsDet.update
                 End If
           Next
        End With
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

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    TxtPlace.text = IIf(IsNull(RsSavRec.Fields("Place").value), "", RsSavRec.Fields("Place").value)
    Dim TimD As Date
    If Not IsNull(RsSavRec.Fields("AccTime").value) Then
    TimD = FormatDateTime(RsSavRec.Fields("AccTime").value, vbShortTime)
    AccTime.value = TimD
    End If
    Me.DcbCrID.BoundText = IIf(IsNull(RsSavRec.Fields("CarID").value), "", RsSavRec.Fields("CarID").value)
    
    TxtHowConected.text = IIf(IsNull(RsSavRec.Fields("HowConected").value), "", RsSavRec.Fields("HowConected").value)
    TxtTheirName.text = IIf(IsNull(RsSavRec.Fields("TheirName").value), "", RsSavRec.Fields("TheirName").value)
    AccDate.value = IIf(IsNull(RsSavRec.Fields("AccDate").value), Date, RsSavRec.Fields("AccDate").value)
    TXTJobTitle.text = IIf(IsNull(RsSavRec.Fields("JobTitle").value), "", RsSavRec.Fields("JobTitle").value)
    TxtStaffNo.text = IIf(IsNull(RsSavRec.Fields("StaffNo").value), "", RsSavRec.Fields("StaffNo").value)
    TxtNameDriver.text = IIf(IsNull(RsSavRec.Fields("NameDriver").value), "", RsSavRec.Fields("NameDriver").value)
    TxtPlateNo.text = IIf(IsNull(RsSavRec.Fields("PlateNo").value), "", RsSavRec.Fields("PlateNo").value)
    TxtFleetNo.text = IIf(IsNull(RsSavRec.Fields("FleetNo").value), "", RsSavRec.Fields("FleetNo").value)
    TxtYearMake.text = IIf(IsNull(RsSavRec.Fields("YearMake").value), "", RsSavRec.Fields("YearMake").value)
    TxtModel.text = IIf(IsNull(RsSavRec.Fields("Model").value), "", RsSavRec.Fields("Model").value)
    TxtVehicleType.text = IIf(IsNull(RsSavRec.Fields("VehicleType").value), "", RsSavRec.Fields("VehicleType").value)
    TxtDimThiParInvol.text = IIf(IsNull(RsSavRec.Fields("DimThiParInvol").value), "", RsSavRec.Fields("DimThiParInvol").value)
    If Not IsNull(RsSavRec.Fields("ThirdPartyInvol").value) Then
    If (RsSavRec.Fields("ThirdPartyInvol").value) = 1 Then
    ThirdPartyInvol(1).value = True
    Else
    ThirdPartyInvol(0).value = True
    End If
    Else
    ThirdPartyInvol(0).value = True
    End If
    TxtDimagVehicle.text = IIf(IsNull(RsSavRec.Fields("DimagVehicle").value), "", RsSavRec.Fields("DimagVehicle").value)
    TxtThidPayment.text = IIf(IsNull(RsSavRec.Fields("ThidPayment").value), "", RsSavRec.Fields("ThidPayment").value)
    TxtRepairCost.text = IIf(IsNull(RsSavRec.Fields("RepairCost").value), "", RsSavRec.Fields("RepairCost").value)
    If Not IsNull(RsSavRec.Fields("IsPhoto").value) Then
    If (RsSavRec.Fields("IsPhoto").value) = 1 Then
    IsPhoto(1).value = True
    Else
    IsPhoto(0).value = True
    End If
    Else
    IsPhoto(0).value = True
    End If
    If Not IsNull(RsSavRec.Fields("IsNdliab").value) Then
    If (RsSavRec.Fields("IsNdliab").value) = 1 Then
    IsNdliab(1).value = True
    Else
    IsNdliab(0).value = True
    End If
    Else
    IsNdliab(0).value = True
    End If
    If Not IsNull(RsSavRec.Fields("IsPolice").value) Then
    If (RsSavRec.Fields("IsPolice").value) = 1 Then
    IsPolice(1).value = True
    Else
    IsPolice(0).value = True
    End If
    Else
    IsPolice(0).value = True
    End If
    If Not IsNull(RsSavRec.Fields("IsContrct").value) Then
    If (RsSavRec.Fields("IsContrct").value) = 1 Then
    IsContrct.value = vbChecked
    Else
    IsContrct.value = vbUnchecked
    End If
    Else
    IsContrct.value = vbUnchecked
    End If
    If Not IsNull(RsSavRec.Fields("IsEmployee").value) Then
    If (RsSavRec.Fields("IsEmployee").value) = 1 Then
    IsEmployee.value = vbChecked
    Else
    IsEmployee.value = vbUnchecked
    End If
    Else
    IsEmployee.value = vbUnchecked
    End If
    If Not IsNull(RsSavRec.Fields("IsOther").value) Then
    If (RsSavRec.Fields("IsOther").value) = 1 Then
    IsOther.value = vbChecked
    Else
    IsOther.value = vbUnchecked
    End If
    Else
    IsOther.value = vbUnchecked
    End If
    TxtDepatment.text = IIf(IsNull(RsSavRec.Fields("Depatment").value), "", RsSavRec.Fields("Depatment").value)
    TxtJobTitle2.text = IIf(IsNull(RsSavRec.Fields("JobTitle2").value), "", RsSavRec.Fields("JobTitle2").value)
    TxtStaffNo2.text = IIf(IsNull(RsSavRec.Fields("StaffNo2").value), "", RsSavRec.Fields("StaffNo2").value)
    DateComm.value = IIf(IsNull(RsSavRec.Fields("DateComm").value), Date, RsSavRec.Fields("DateComm").value)
    TxtDesInjuried.text = IIf(IsNull(RsSavRec.Fields("DesInjuried").value), "", RsSavRec.Fields("DesInjuried").value)
    TxtDesIncident.text = IIf(IsNull(RsSavRec.Fields("DesIncident").value), "", RsSavRec.Fields("DesIncident").value)
    TxtDesWhyHappen.text = IIf(IsNull(RsSavRec.Fields("DesWhyHappen").value), "", RsSavRec.Fields("DesWhyHappen").value)
    TxtNameAddress.text = IIf(IsNull(RsSavRec.Fields("NameAddress").value), "", RsSavRec.Fields("NameAddress").value)
    TxtFirstAid.text = IIf(IsNull(RsSavRec.Fields("FirstAid").value), "", RsSavRec.Fields("FirstAid").value)
    If Not IsNull(RsSavRec.Fields("IsFirstAid").value) Then
    If (RsSavRec.Fields("IsFirstAid").value) = 1 Then
    IsFirstAid(1).value = True
    Else
    IsFirstAid(0).value = True
    End If
    Else
    IsFirstAid(0).value = True
    End If
    TxtDesHospital.text = IIf(IsNull(RsSavRec.Fields("DesHospital").value), "", RsSavRec.Fields("DesHospital").value)
    TxtActionTaken.text = IIf(IsNull(RsSavRec.Fields("ActionTaken").value), "", RsSavRec.Fields("ActionTaken").value)
    TxtName1.text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    TxtPosition.text = IIf(IsNull(RsSavRec.Fields("Position").value), "", RsSavRec.Fields("Position").value)
    TxtDepartment2.text = IIf(IsNull(RsSavRec.Fields("Department2").value), "", RsSavRec.Fields("Department2").value)
    DateMaking.value = IIf(IsNull(RsSavRec.Fields("DateMaking").value), Date, RsSavRec.Fields("DateMaking").value)
    TotalValue.Caption = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, RsSavRec.Fields("TotalValue").value)
    CompValue.Caption = IIf(IsNull(RsSavRec.Fields("CompValue").value), 0, RsSavRec.Fields("CompValue").value)
    FullGrid
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
ErrTrap:
End Sub
Sub Relin()
Dim i As Integer
Dim Cour As Integer
Dim SmCompVal As Double
Dim total As Double
Cour = 0
SmCompVal = 0
total = 0
With Grid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("ExpID"))) <> 0 Then
Cour = Cour + 1
.TextMatrix(i, .ColIndex("Serial")) = Cour
total = total + val(.TextMatrix(i, .ColIndex("Valuee")))
If val(.TextMatrix(i, .ColIndex("Typ"))) = 1 Then
SmCompVal = SmCompVal + val(.TextMatrix(i, .ColIndex("Valuee")))
End If
End If
Next i
End With
TotalValue.Caption = total
CompValue.Caption = SmCompVal
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
    If val(DcbBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
    Else
    MsgBox "Please Select Branch"
    End If
    DcbBranch.SetFocus
    Exit Sub
    End If


    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub


Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim LngRow As Long
    With Grid

     Select Case .ColKey(Col)
 Case "Name"
                        StrAccountCode = .ComboData
                       LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ExpID"), False, True)
                      .TextMatrix(Row, .ColIndex("ExpID")) = StrAccountCode
     End Select
     If .Row = .rows - 1 Then
     .rows = .rows + 1
     End If
End With
Relin
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With Grid
     Select Case .ColKey(Col)
     Case "Valuee"
     .ComboList = ""
     Case "Remarks"
     .ComboList = ""
     End Select
    End With
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Rs_Temp As ADODB.Recordset
    Dim Msg As String
    With Grid
    Select Case .ColKey(Col)
    Case "Name"
        Set Rs_Temp = New ADODB.Recordset
          StrSQL = " Select * From TblDataTypeExchange  "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          If SystemOptions.UserInterface = ArabicInterface Then
          StrComboList = .BuildComboList(Rs_Temp, "Name", "ID")
          Else
          StrComboList = .BuildComboList(Rs_Temp, "NameE", "ID")
          End If
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList

     End Select
   End With
End Sub

Private Sub ISButton2_Click()
print_report2
End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
ShowAttachments TxtSerial1.text, "16112016123"
ErrTrap:
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  
    MySQL = " SELECT     dbo.TblAccidentReport.ID, dbo.TblAccidentReport.RecordDate, dbo.TblAccidentReport.BranchID, dbo.TblBranchesData.branch_name, "
    MySQL = MySQL & "                  dbo.TblBranchesData.branch_namee, dbo.TblAccidentReport.Place, dbo.TblAccidentReport.AccTime, dbo.TblAccidentReport.AccDate, dbo.TblAccidentReport.JobTitle,"
    MySQL = MySQL & "                  dbo.TblAccidentReport.StaffNo, dbo.TblAccidentReport.NameDriver, dbo.TblAccidentReport.PlateNo, dbo.TblAccidentReport.VehicleType,"
    MySQL = MySQL & "                  dbo.TblAccidentReport.FleetNo, dbo.TblAccidentReport.Model, dbo.TblAccidentReport.YearMake, dbo.TblAccidentReport.ThirdPartyInvol,"
    MySQL = MySQL & "                  dbo.TblAccidentReport.DimThiParInvol, dbo.TblAccidentReport.DimagVehicle, dbo.TblAccidentReport.ThidPayment, dbo.TblAccidentReport.RepairCost,"
    MySQL = MySQL & "                  dbo.TblAccidentReport.IsPhoto, dbo.TblAccidentReport.IsNdliab, dbo.TblAccidentReport.IsPolice, dbo.TblAccidentReport.HowConected,"
    MySQL = MySQL & "                  dbo.TblAccidentReport.TheirName, dbo.TblAccidentReport.IsContrct, dbo.TblAccidentReport.IsEmployee, dbo.TblAccidentReport.IsOther,"
    MySQL = MySQL & "                  dbo.TblAccidentReport.Depatment, dbo.TblAccidentReport.JobTitle2, dbo.TblAccidentReport.CarID, dbo.TblCarsData.Fullcode, dbo.TblCarsData.BoardNO,"
    MySQL = MySQL & "                  dbo.TblAccidentReportDet.Valuee, dbo.TblAccidentReportDet.Typ, dbo.TblAccidentReportDet.Remarks, dbo.TblAccidentReportDet.ExpID,"
    MySQL = MySQL & "                  dbo.TblDataTypeExchange.Name , dbo.TblDataTypeExchange.NameE"
    MySQL = MySQL & "  FROM         dbo.TblDataTypeExchange RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAccidentReportDet ON dbo.TblDataTypeExchange.Id = dbo.TblAccidentReportDet.ExpID RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAccidentReport ON dbo.TblAccidentReportDet.AccID = dbo.TblAccidentReport.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCarsData ON dbo.TblAccidentReport.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData ON dbo.TblAccidentReport.BranchID = dbo.TblBranchesData.branch_id"
    MySQL = MySQL & " Where (dbo.TblAccidentReport.ID = " & val(TxtSerial1.text) & ")"

       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAccedentReporttExpen.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAccedentReporttExpen.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
     
  '      xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  
    MySQL = " SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAccidentReport.*"
    MySQL = MySQL & "        FROM         dbo.TblAccidentReport LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData ON dbo.TblAccidentReport.BranchID = dbo.TblBranchesData.branch_id"
    MySQL = MySQL & " Where (dbo.TblAccidentReport.ID = " & val(TxtSerial1.text) & ")"
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAccedentReportt.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAccedentReportt.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
     
  '      xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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

Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.text = ""
If Len(txtLetter1.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub

Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.text = ""
If Len(txtLetter2.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.text = ""
If Len(txtLetter3.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.text = ""
If Len(txtLetter4.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.text = ""
If Len(txtNum1.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.text = ""
If Len(txtNum2.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.text = ""
If Len(txtNum3.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.text = ""
If Len(txtNum4.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If

Cal_Board
End Sub


Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub Cal_Board()
    Me.TxtPlateNo.text = txtLetter1.text & " " & txtLetter2.text & " " & txtLetter3.text & " " & txtLetter4.text & " " & txtNum1.text & " " & txtNum2.text & " " & txtNum3.text & " " & txtNum4.text
End Sub

Private Sub TxtPlateNo_Change()
Label3.Caption = TxtPlateNo.text
If Me.TxtModFlg.text <> "R" Then
If TxtPlateNo.text <> "" Then
GetCarInformation
End If
End If
End Sub

Private Sub TxtPlateNo_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "FrmAccidentReport"
        FrmCasrShearches.show vbModal
    End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

' search for select id
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
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
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
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
               Cn.Execute "Delete from TblAccidentReportDet where AccID=" & val(TxtSerial1.text) & " "
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                RsSavRec.delete
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            TotalValue.Caption = 0
            CompValue.Caption = 0
           Grid.Clear flexClearScrollable, flexClearEverything
           Grid.rows = Grid.FixedRows + 1
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           Cn.Errors.Clear
    End Select

End Sub

' exit without save sub
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
    If TxtModFlg.text = "N" Then
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
    ElseIf TxtModFlg.text = "R" Then
   ' XPDtbTrans.Enabled = False
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
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
           Grid.rows = Grid.rows + 1
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    TxtModFlg.text = "N"
    Grid.Clear flexClearScrollable, flexClearEverything
           Grid.rows = Grid.FixedRows + 1
    Me.DcbBranch.BoundText = Current_branch
    Me.DCboUserName.BoundText = user_id
RecordDate.value = Date
TotalValue.Caption = 0
CompValue.Caption = 0
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

'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
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
' short cut for keys
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
 
  Label1(2).Caption = "Record Attendance "
lbl(4).Caption = "No"
lbl(11).Caption = "Branch"
lbl(25).Caption = "Date"
Label1(3).Caption = "Group"
Label1(0).Caption = "Instractor"
Label1(1).Caption = "Remarks"
Label1(11).Caption = "Subject"
Label1(9).Caption = "Hall"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"
lbl(14).Caption = "By"
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

   
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblAccidentReport"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
