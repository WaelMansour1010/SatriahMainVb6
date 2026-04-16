VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_TradingContract 
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14430
   Icon            =   "Frm_TradingContract.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   14430
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
      TabIndex        =   3
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   288
      ItemData        =   "Frm_TradingContract.frx":6852
      Left            =   15480
      List            =   "Frm_TradingContract.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   4
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
      TabIndex        =   5
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
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1665
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7710
      Width           =   14430
      _cx             =   25453
      _cy             =   2937
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
      Align           =   2
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   1
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
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   465
         Left            =   90
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   600
         Width           =   3045
      End
      Begin VB.TextBox TxtNoteID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   630
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.CommandButton CmdCreateV 
         Caption         =   "≈‰‘«¡ «·ﬁÌœ "
         Height          =   465
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   120
         Width           =   2430
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ–› «·ﬁÌœ "
         Height          =   465
         Left            =   3990
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   120
         Width           =   2790
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
         Height          =   465
         Left            =   7890
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   600
         Width           =   2385
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   312
         Left            =   10680
         TabIndex        =   12
         Top             =   360
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   -840
         Width           =   13965
         _cx             =   24633
         _cy             =   1005
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   636
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1032
         Width           =   14436
         _cx             =   25453
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
            Height          =   336
            Left            =   12912
            TabIndex        =   21
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
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
            ButtonImage     =   "Frm_TradingContract.frx":687B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   336
            Left            =   9612
            TabIndex        =   22
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
            Width           =   1356
            _ExtentX        =   2381
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
            ButtonImage     =   "Frm_TradingContract.frx":D0DD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   336
            Left            =   11328
            TabIndex        =   23
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
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
            ButtonImage     =   "Frm_TradingContract.frx":D477
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   336
            Left            =   7788
            TabIndex        =   24
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
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
            ButtonImage     =   "Frm_TradingContract.frx":13CD9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   336
            Left            =   5712
            TabIndex        =   25
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   120
            Width           =   1356
            _ExtentX        =   2381
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
            ButtonImage     =   "Frm_TradingContract.frx":14073
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   336
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   120
            Width           =   1236
            _ExtentX        =   2170
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
            ButtonImage     =   "Frm_TradingContract.frx":1460D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   408
            Left            =   4248
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Width           =   1032
            _ExtentX        =   1826
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
            ButtonImage     =   "Frm_TradingContract.frx":149A7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   336
            Left            =   1824
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Width           =   972
            _ExtentX        =   1720
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
            ButtonImage     =   "Frm_TradingContract.frx":1B209
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   408
            Left            =   3240
            TabIndex        =   67
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Visible         =   0   'False
            Width           =   1152
            _ExtentX        =   2037
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·„⁄«ÌÌ—"
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
            ButtonImage     =   "Frm_TradingContract.frx":1B5A3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         Height          =   390
         Index           =   35
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   276
         Index           =   8
         Left            =   13560
         TabIndex        =   13
         Top             =   360
         Width           =   900
      End
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
            Picture         =   "Frm_TradingContract.frx":21E05
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":2219F
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":22539
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":228D3
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":22C6D
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":23007
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":233A1
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_TradingContract.frx":2393B
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   14
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
      ButtonImage     =   "Frm_TradingContract.frx":23CD5
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
      ButtonImage     =   "Frm_TradingContract.frx":2A537
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
      ButtonImage     =   "Frm_TradingContract.frx":30D99
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic100 
      Height          =   7710
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   14430
      _cx             =   25453
      _cy             =   13600
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
      Begin VB.TextBox txtPercentAlarm 
         Alignment       =   2  'Center
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
         Left            =   7740
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   7170
         Width           =   1035
      End
      Begin VB.TextBox txtTotalInstall 
         Alignment       =   2  'Center
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
         Left            =   11340
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   6660
         Width           =   1245
      End
      Begin VB.TextBox txtTotalPrice 
         Alignment       =   2  'Center
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
         Left            =   8580
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   6660
         Width           =   1245
      End
      Begin VB.TextBox txtProjectTotal2 
         Alignment       =   2  'Center
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
         Left            =   5130
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   6660
         Width           =   1245
      End
      Begin VB.TextBox txtTotalWithVat2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   210
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   6660
         Width           =   1335
      End
      Begin VB.TextBox TxtVAt22 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   6660
         Width           =   795
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   0
         Width           =   14484
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   32
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
            ButtonImage     =   "Frm_TradingContract.frx":31133
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   33
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
            ButtonImage     =   "Frm_TradingContract.frx":314CD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   34
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
            ButtonImage     =   "Frm_TradingContract.frx":31867
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   35
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
            ButtonImage     =   "Frm_TradingContract.frx":31C01
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "Frm_TradingContract.frx":31F9B
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "√ ›«ﬁÌ…"
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
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   4080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   585
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   795
         Width           =   14190
         _cx             =   25030
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
         Begin VB.CheckBox chkIsSupply 
            Alignment       =   1  'Right Justify
            Caption         =   " Ê—Ìœ ›ﬁÿ"
            Height          =   255
            Left            =   4830
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   120
            Width           =   1845
         End
         Begin VB.CheckBox chkIsCanceld 
            Alignment       =   1  'Right Justify
            Caption         =   "«·€«¡ «·« ›«ﬁÌ…"
            Height          =   315
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   90
            Width           =   1335
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   11952
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   108
            Width           =   792
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   105
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   503
            _Version        =   393216
            Format          =   105250817
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal RecorddateH 
            Height          =   240
            Left            =   8610
            TabIndex        =   40
            Top             =   105
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   423
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ «·ÂÃ—Ï"
            Height          =   270
            Index           =   7
            Left            =   10275
            TabIndex        =   43
            Top             =   105
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„"
            Height          =   270
            Index           =   4
            Left            =   12870
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   105
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ «·„Ì·«œÌ"
            Height          =   270
            Index           =   2
            Left            =   2970
            TabIndex        =   41
            Top             =   120
            Width           =   1890
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
         Height          =   2325
         Left            =   120
         TabIndex        =   44
         Top             =   4080
         Width           =   14190
         _cx             =   25030
         _cy             =   4101
         Appearance      =   2
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   12
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"Frm_TradingContract.frx":333A0
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
         WallPaperAlignment=   0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   2880
         Left            =   120
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1500
         Width           =   14220
         _cx             =   25083
         _cy             =   5080
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
            Height          =   525
            Left            =   120
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   915
            Width           =   13980
            _cx             =   24659
            _cy             =   926
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
            Begin VB.TextBox TxtAddress 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   300
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   165
               Width           =   4668
            End
            Begin VB.TextBox TxtPhone 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   300
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   165
               Width           =   1548
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   12000
               TabIndex        =   47
               Top             =   165
               Width           =   1035
            End
            Begin MSDataListLib.DataCombo DcbCus 
               Height          =   315
               Left            =   8520
               TabIndex        =   50
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   165
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄‰Ê«‰"
               Height          =   195
               Index           =   16
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   165
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Â« ›"
               Height          =   195
               Index           =   17
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   165
               Width           =   435
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   195
               Index           =   1
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   165
               Width           =   915
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   1530
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1440
            Width           =   13980
            _cx             =   24659
            _cy             =   2699
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
            Begin VB.TextBox txtPeriod2 
               Alignment       =   2  'Center
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
               Left            =   4530
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   480
               Width           =   948
            End
            Begin VB.TextBox Txt_specific_Ar 
               Alignment       =   2  'Center
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
               Left            =   8280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   480
               Width           =   2748
            End
            Begin VB.TextBox Txt_Qun 
               Alignment       =   2  'Center
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
               Left            =   5640
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   480
               Width           =   948
            End
            Begin VB.TextBox Txt_SalPrice 
               Alignment       =   2  'Center
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
               Left            =   3330
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   480
               Width           =   948
            End
            Begin VB.TextBox Txt_Install 
               Alignment       =   2  'Center
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
               Left            =   2130
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   480
               Width           =   948
            End
            Begin VB.TextBox Txt_Value 
               Alignment       =   2  'Center
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
               Left            =   1110
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   480
               Width           =   948
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   360
               Left            =   30
               TabIndex        =   59
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   635
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
               ButtonImage     =   "Frm_TradingContract.frx":33601
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin MSDataListLib.DataCombo DcboUnits 
               Height          =   315
               Left            =   6840
               TabIndex        =   60
               Top             =   480
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DC_specific 
               Height          =   315
               Left            =   11040
               TabIndex        =   72
               Top             =   480
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
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
               Caption         =   "«·„œ… «·“„‰Ì…"
               Height          =   405
               Index           =   22
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·‘—Õ"
               Height          =   405
               Index           =   7
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   120
               Width           =   795
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÊÕœ… «·ﬁÌ«”"
               Height          =   405
               Index           =   11
               Left            =   7080
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   120
               Width           =   795
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   405
               Index           =   12
               Left            =   5670
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   120
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Ê«’›« "
               Height          =   405
               Index           =   3
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   120
               Width           =   795
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄— «· Ê—Ìœ"
               Height          =   405
               Index           =   5
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   120
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄— «· —ﬂÌ»"
               Height          =   405
               Index           =   0
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÌ„…"
               Height          =   405
               Index           =   4
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   120
               Width           =   915
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   930
            Left            =   150
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   -60
            Width           =   13950
            _cx             =   24606
            _cy             =   1640
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
            Begin VB.TextBox txtNewMeasureNo 
               Alignment       =   2  'Center
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
               Left            =   4950
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   60
               Width           =   1035
            End
            Begin VB.TextBox txtNetBVat 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   5970
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   510
               Width           =   1185
            End
            Begin VB.TextBox txtTotalDisc 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   60
               Width           =   1215
            End
            Begin VB.TextBox TxtVATValue 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   -600
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   0
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.TextBox TxtVAt2 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   4080
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   525
               Width           =   795
            End
            Begin VB.TextBox txtTotalWithVat 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   150
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   510
               Width           =   1185
            End
            Begin VB.TextBox txtRespWorker 
               Alignment       =   2  'Center
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
               Left            =   9210
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   525
               Width           =   1365
            End
            Begin VB.TextBox txtResp 
               Alignment       =   2  'Center
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
               Left            =   11700
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   525
               Width           =   1515
            End
            Begin VB.TextBox txtProjectTotal 
               Alignment       =   2  'Center
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
               Left            =   2070
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   45
               Width           =   1245
            End
            Begin VB.TextBox txtPeriod 
               Alignment       =   2  'Center
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
               Left            =   6930
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   60
               Width           =   948
            End
            Begin VB.TextBox txtLocation 
               Alignment       =   2  'Center
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
               Height          =   300
               Left            =   11700
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   60
               Width           =   1515
            End
            Begin MSComCtl2.DTPicker XPDtProjStart 
               Height          =   345
               Left            =   8820
               TabIndex        =   74
               Top             =   45
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   609
               _Version        =   393216
               Format          =   105250817
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌ«”"
               Height          =   270
               Index           =   23
               Left            =   5970
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   90
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ï »⁄œ ﬁ „÷«›…"
               Height          =   300
               Index           =   21
               Left            =   1530
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   600
               Width           =   1545
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ï ﬁ»· ﬁ „÷«›…"
               Height          =   300
               Index           =   20
               Left            =   7380
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   540
               Width           =   1545
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ’„"
               Height          =   300
               Index           =   19
               Left            =   1230
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   150
               Width           =   795
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… «·„÷«›…"
               Height          =   300
               Index           =   98
               Left            =   5010
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   555
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—ﬂ» + ⁄„«·"
               Height          =   300
               Index           =   9
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   555
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—ﬂ» «·„”∆Ê·"
               Height          =   420
               Index           =   0
               Left            =   13080
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   435
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€ «·«Ã„«·Ï ··„‘—Ê⁄"
               Height          =   270
               Index           =   6
               Left            =   3180
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   105
               Width           =   1785
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œ… «·“„‰Ì…"
               Height          =   270
               Index           =   5
               Left            =   7860
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   90
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ »œ«Ì… «·„‘—Ê⁄"
               Height          =   285
               Index           =   3
               Left            =   10200
               TabIndex        =   76
               Top             =   90
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Êﬁ⁄"
               Height          =   600
               Index           =   1
               Left            =   13260
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   105
               Width           =   630
            End
         End
      End
      Begin ImpulseButton.ISButton Cmd_DeleteRow 
         Height          =   285
         Left            =   12705
         TabIndex        =   70
         Top             =   7245
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " Õ–› ”ÿ—"
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
         ButtonImage     =   "Frm_TradingContract.frx":39E63
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd_DeleteAll 
         Height          =   285
         Left            =   11040
         TabIndex        =   71
         Top             =   7245
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " Õ–› «·ﬂ·"
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
         ButtonImage     =   "Frm_TradingContract.frx":3A3FD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «· ‰»ÌÂ"
         Height          =   270
         Index           =   24
         Left            =   8850
         RightToLeft     =   -1  'True
         TabIndex        =   119
         Top             =   7200
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ã„«·Ï «· —ﬂÌ»"
         Height          =   270
         Index           =   18
         Left            =   12390
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   6690
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ã„«·Ï «· Ê—Ìœ"
         Height          =   270
         Index           =   15
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   6690
         Width           =   1785
      End
      Begin VB.Label lblTotalNet 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   7110
         Width           =   6885
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€ «·«Ã„«·Ï ··„‘—Ê⁄"
         Height          =   270
         Index           =   14
         Left            =   6450
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   6690
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«Ã„«·Ì »⁄œ «·÷—Ì»…"
         Height          =   300
         Index           =   13
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   6675
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„… «·„÷«›…"
         Height          =   300
         Index           =   10
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   6675
         Width           =   945
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
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Frm_TradingContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
                Dim RootAccount1 As String
                        Dim RootAccount2 As String
                        Dim RootAccount3 As String
Public LngRow As Long
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double


Private Sub chkIsSupply_Click()

    If Me.TxtModFlg.Text <> "R" Then
        If chkIsSupply Then
            lbl(9).Visible = False
            txtRespWorker = ""
            txtRespWorker.Visible = False
            Label1(0).Visible = False
            Txt_Install.Visible = False
            Txt_Install.Tag = Txt_Install
            Txt_Install = ""
            lbl(18).Visible = False
            txtTotalInstall.Tag = txtTotalInstall
            txtTotalInstall = ""
            txtTotalInstall.Visible = False
            txtResp.Visible = False
            lbl(0).Visible = False
            GridInstallments.ColHidden(GridInstallments.ColIndex("InstallPrice")) = True
            GridInstallments.ColHidden(GridInstallments.ColIndex("TotalInstall")) = True
            
            Dim i As Long
            For i = 1 To GridInstallments.Rows - 1
                GridInstallments.TextMatrix(i, GridInstallments.ColIndex("InstallPrice")) = 0
                GridInstallments.TextMatrix(i, GridInstallments.ColIndex("TotalInstall")) = 0
                GridInstallments_AfterEdit i, GridInstallments.ColIndex("InstallPrice")
            Next
        Else
            lbl(9).Visible = True
            txtRespWorker.Visible = True
            Label1(0).Visible = True
            Txt_Install.Visible = True
            txtResp.Visible = True
            lbl(0).Visible = True
            'Txt_Install = Txt_Install.Tag
            lbl(18).Visible = True
            'txtTotalInstall = txtTotalInstall.Tag
            txtTotalInstall.Visible = True
            GridInstallments.ColHidden(GridInstallments.ColIndex("InstallPrice")) = False
            GridInstallments.ColHidden(GridInstallments.ColIndex("TotalInstall")) = False
        End If
    End If
End Sub

Private Sub CmdCreateV_Click()

If val(TxtNoteSerial.Text) = 0 Then
If createVoucher Then
       'FindRec val(TXTLCNO.Text)
       
        FindRec val(TxtSerial1.Text)
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
            If val(TxtNoteID) <> 0 Then
                CmdCreateV.Enabled = False
                Command9.Enabled = True
                Command2.Enabled = True
                btnSave.Enabled = False
            Else
                CmdCreateV.Enabled = True
                Command9.Enabled = False
                Command2.Enabled = False
            End If
        Else
            MsgBox "Done"
        End If
    Else
        CmdCreateV.Enabled = True
        Command9.Enabled = False
        Command2.Enabled = False
    End If
End If
End Sub

Private Sub Command2_Click()
If Me.TxtModFlg.Text = "R" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› «·ﬁÌœ "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update Tbl_TradingContract set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1.Text)
       
        
        RsSavRec.Requery
         FindRec val(TxtSerial1.Text)
         TxtModFlg.Text = ""
         TxtNoteSerial = ""
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› «·ﬁÌœ   "
            
           
            If val(TxtNoteID) <> 0 Then
                CmdCreateV.Enabled = False
                Command9.Enabled = True
                Command2.Enabled = True
                btnSave.Enabled = False
                btnModify.Enabled = False
             Else
                CmdCreateV.Enabled = True
                Command9.Enabled = False
                Command2.Enabled = False
            End If
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If
  
 
End Sub


Function createVoucher() As Boolean
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·" & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
tablename = "Tbl_TradingContract"

Filedname = "ID"
NoteSerial1 = val(TxtSerial1)

BranchID = val(branch_id)
mRate = 1

'


' ⁄‰ „ﬂ«‰ Ê÷⁄ «·ÀÊ«»  ÊﬂÌ›Ì… «· —ﬁÌ„   Õ «Ã  Ê÷ÌÕ
' «” ›”«— Ê«∆·
' ·Œ»ÿ… ⁄‰œÏ ›Ï «·„”„Ì«  Ê«·‰Ê   «Ì»
notytype = 3335
Notevalue = val(txtTotalWithVat)

'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des           ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial

    If Not CREATE_VOUCHER_GE(val(TxtNoteID.Text), BranchID, val(DCboUserName.BoundText), NoteDate) Then createVoucher = False Else createVoucher = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(TxtNoteSerial.Text), Format(txtTotalWithVat2.Text, "###.00")
'
'
'    StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"
'
'    StrSQL = StrSQL & " Where " & Filedname & " = " & NoteSerial1 & ""
'    Cn.Execute StrSQL
     
     
 
End If
End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date) As Boolean

     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtSerial1.Text
    notes_id = general_noteid
    my_branch = val(branch_id)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
    
    Dim s As String
    Dim mRate As Double
    mRate = 1
    ' „‰ Õ”«» «·⁄„Ì·
    StrAccountCodeDebt = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCus.BoundText))
    
   
    Notevalue = val(txtTotalWithVat.Text)
    If Notevalue > 0 Then
        
       ' StrAccountCodeDebt = Trim(DboParentAccount.BoundText)
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·⁄„Ì·  ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , DcbCus.BoundText) = False Then
            GoTo ErrTrap
        End If
       ' «·Ï Õ”«» «·ﬁÌ„… «·„÷«›…
        GetValueAddedAccount XPDtbTrans.value, , StrAccountCodeCridet, 1, 43
        
        line_no = line_no + 1
        
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(TxtVAt2), 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(branch_id)) = False Then
            GoTo ErrTrap
        End If
        line_no = line_no + 1
    End If

    
    ' «·«ÿ—«›
    
     ' «·Ï Õ”«» «Ì—«œ«  «· —ﬂÌÌ« 
    Notevalue = val(txtNetBVat.Text)
    If Notevalue > 0 Then
    
                StrAccountCodeCridet = get_account_code_branch(159, my_branch)
        
                If StrAccountCodeCridet = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "No Branch Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If StrAccountCodeCridet = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·«Ì—«œ«  Ê«· —ﬂÌ»« ", vbCritical
                    Else
                        MsgBox "Please Select Account VAT ", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If

        
        
 
        
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«» «Ì—«œ«  Ê —ﬂÌ»«   ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(branch_id)) = False Then
            GoTo ErrTrap
        End If

        line_no = line_no + 1
    End If
    

    updateNotesValueAndNobytext (val(notes_id))
    CREATE_VOUCHER_GE = True
    Exit Function
ErrTrap:
CREATE_VOUCHER_GE = False
  End Function



Private Sub Command9_Click()

'add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3335 ,'«·« ›«ﬁÌ«     ' ,'      Trading Contract' ", "NotesType", 3335
'add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", "«·« ›«ﬁÌ«  ", "NotesType", 3335
ShowGL_cc Me.TxtNoteSerial.Text, , 3335

End Sub

 
 
Private Sub DcbInvise_Change()
'DcbInvise_Click (0)
End Sub
 Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
End Sub
 


  
Private Sub Cmd_Click(Index As Integer)

End Sub

 

Private Sub Cmd_DeleteAll_Click()
If Me.TxtModFlg.Text <> "R" Then


 GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 2
CalcTotal
End If
End Sub

Private Sub Cmd_DeleteRow_Click()
If Me.TxtModFlg.Text <> "R" Then

RemoveGridRow
CalcTotal

End If
End Sub
Private Sub RemoveGridRow()

    With Me.GridInstallments
'MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub



Private Sub DC_specific_Click(Area As Integer)

If DC_specific.Text = "" Then Exit Sub
 Dim UnitID As Integer
 Dim ProcessName As String
 
' GetTblProcessDEF DC_specific.BoundText, , ProcessName, UnitID
 DcboUnits.BoundText = UnitID
 Txt_specific_Ar.Text = ProcessName

End Sub

Private Sub DcbCus_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 104
        FrmCustemerSearch.show vbModal
    End If
End Sub

Private Sub DcboUnits_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txt_Qun.SetFocus
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    

    
    conection = "select * from Tbl_TradingContract order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    SetDtpickerDate Me.XPDtProjStart


If SystemOptions.UserInterface = ArabicInterface Then

End If
  If SystemOptions.UserInterface = ArabicInterface Then
    
End If
     
     

     
     
     If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=20 and Flg=1  order by CusName"
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=20 and Flg=1  order by CusNamee"
    End If
    fill_combo DcbCus, My_SQL
   
    Dim Dcombos As New ClsDataCombos
    'Dcombos.GetCustomersSuppliers 1, Me.DcbSales, True
    Dcombos.GetCustomersSuppliers 1, Me.DcbCus
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetIProcessUnits Me.DcboUnits
    Dcombos.GetProcess Me.DC_specific
    
   ' Dcombos.GetInvestmentActive Me.DcbInvise, 1
  '  Dcombos.GetCustomerType Me.DcCustomerType
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
  ' FiLLTXT
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
               
               'StrSQL = "Delete From TblTransactionInvest Where BuyBilID =" & val(TxtSerial1.Text) & ""
                  'Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From Tbl_TradingContractDet Where TContractDet_TContractID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
               
               StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
                
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                TxtNoteSerial = 0
                TxtNoteID = 0
               
              End If
              
              
              
    RsSavRec.Fields("TContract_DateH").value = RecorddateH.value
    RsSavRec.Fields("TContract_Date").value = XPDtbTrans.value
    RsSavRec.Fields("TContract_CustID").value = val(Me.DcbCus.BoundText)
    RsSavRec.Fields("TOrder_Address").value = (Me.TxtAddress.Text)
    RsSavRec.Fields("TOrder_Phone").value = (Me.TxtPhone.Text)
    RsSavRec.Fields("UserID").value = (Me.DCboUserName.BoundText)
    RsSavRec.Fields("NewMeasureNo").value = val(txtNewMeasureNo.Text)
    
    RsSavRec.Fields("Period").value = val(Me.txtPeriod.Text)
    RsSavRec.Fields("Location").value = (Me.txtLocation.Text)
    RsSavRec.Fields("Resp").value = (Me.txtResp.Text)
    RsSavRec.Fields("RespWorker").value = (Me.txtRespWorker.Text)
    RsSavRec.Fields("DtProjStart").value = XPDtProjStart.value
    RsSavRec.Fields("ProjectTotal").value = val(Me.txtProjectTotal.Text)
    RsSavRec.Fields("PercentAlarm").value = val(Me.txtPercentAlarm.Text)
    
    RsSavRec.Fields("NetBVat").value = val(Me.txtNetBVat.Text)
    RsSavRec.Fields("TotalDisc").value = val(Me.txtTotalDisc.Text)
    
    RsSavRec.Fields("VAt2").value = val(Me.TxtVAt2.Text)
    If Me.chkIsCanceld.value = vbChecked Then
        RsSavRec.Fields("IsCanceld").value = 1
    Else
       RsSavRec.Fields("IsCanceld").value = 0
    End If
    
    If Me.chkIsSupply.value = vbChecked Then
        RsSavRec.Fields("IsSupply").value = 1
    Else
       RsSavRec.Fields("IsSupply").value = 0
    End If

    RsSavRec!NoteSerial = Null
    RsSavRec!NoteID = Null
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from Tbl_TradingContractDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = Msg & " ‰«“·"
    Else
    Msg = Msg & "Waiver/Sale of Shares"
    End If
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If (.TextMatrix(i, .ColIndex("specification"))) <> "" Then
       RsDevsub.AddNew
                RsDevsub("TContractDet_TContractID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TContractDet_specification").value = IIf((.TextMatrix(i, .ColIndex("specification"))) = "", Null, (.TextMatrix(i, .ColIndex("specification"))))
                RsDevsub("TContractDet_UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, (.TextMatrix(i, .ColIndex("UnitID"))))
                RsDevsub("TContractDet_Qun").value = IIf((.TextMatrix(i, .ColIndex("Qun"))) = "", Null, val(.TextMatrix(i, .ColIndex("Qun"))))
                RsDevsub("TContractDet_SalPrice").value = IIf((.TextMatrix(i, .ColIndex("SalPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("SalPrice"))))
                RsDevsub("TContractDet_InstallPrice").value = IIf((.TextMatrix(i, .ColIndex("InstallPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("InstallPrice"))))
                RsDevsub("TContractDet_TotalSalPrice").value = IIf((.TextMatrix(i, .ColIndex("TotalPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotalPrice"))))
                RsDevsub("TContractDet_TotalInstallPrice").value = IIf((.TextMatrix(i, .ColIndex("TotalInstall"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotalInstall"))))
                RsDevsub("ProcessDEFID").value = IIf((.TextMatrix(i, .ColIndex("ProcessID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProcessID"))))
                RsDevsub("TContractDet_specificationEn").value = IIf((.TextMatrix(i, .ColIndex("specificationAr"))) = "", Null, (.TextMatrix(i, .ColIndex("specificationAr"))))
                RsDevsub("TContractDet_Value").value = IIf((.TextMatrix(i, .ColIndex("Value"))) = "", Null, val(.TextMatrix(i, .ColIndex("Value"))))
                RsDevsub("FinishDate").value = IIf((.TextMatrix(i, .ColIndex("FinishDate"))) = "", Null, val(.TextMatrix(i, .ColIndex("FinishDate"))))
                RsDevsub("Periods").value = IIf((.TextMatrix(i, .ColIndex("Periods"))) = "", Null, val(.TextMatrix(i, .ColIndex("Periods"))))
       
       RsDevsub.update
       
      End If
     Next i
    End With

    
'''///////////////
  
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
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    RecorddateH.value = IIf(IsNull(RsSavRec.Fields("TContract_DateH").value), Date, RsSavRec.Fields("TContract_DateH").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("TContract_Date").value), "", RsSavRec.Fields("TContract_Date").value)
    DcbCus.BoundText = IIf(IsNull(RsSavRec.Fields("TContract_CustID").value), 0, RsSavRec.Fields("TContract_CustID").value)
    TxtAddress.Text = IIf(IsNull(RsSavRec.Fields("TOrder_Address").value), "", RsSavRec.Fields("TOrder_Address").value)
    TxtPhone.Text = IIf(IsNull(RsSavRec.Fields("TOrder_Phone").value), "", RsSavRec.Fields("TOrder_Phone").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    
    txtNewMeasureNo.Text = IIf(IsNull(RsSavRec.Fields("NewMeasureNo").value), "", RsSavRec.Fields("NewMeasureNo").value)
    
    txtPeriod.Text = IIf(IsNull(RsSavRec.Fields("Period").value), "", RsSavRec.Fields("Period").value)
    txtLocation.Text = IIf(IsNull(RsSavRec.Fields("Location").value), "", RsSavRec.Fields("Location").value)
    txtResp.Text = IIf(IsNull(RsSavRec.Fields("Resp").value), "", RsSavRec.Fields("Resp").value)
    txtRespWorker.Text = IIf(IsNull(RsSavRec.Fields("RespWorker").value), "", RsSavRec.Fields("RespWorker").value)
    txtPeriod.Text = IIf(IsNull(RsSavRec.Fields("Period").value), "", RsSavRec.Fields("Period").value)
    XPDtProjStart.value = IIf(IsNull(RsSavRec.Fields("DtProjStart").value), Date, RsSavRec.Fields("DtProjStart").value)
    txtPercentAlarm.Text = IIf(IsNull(RsSavRec.Fields("PercentAlarm").value), 0, RsSavRec.Fields("PercentAlarm").value)
    txtProjectTotal.Text = IIf(IsNull(RsSavRec.Fields("ProjectTotal").value), "", RsSavRec.Fields("ProjectTotal").value)
    
    txtNetBVat.Text = IIf(IsNull(RsSavRec.Fields("NetBVat").value), "", RsSavRec.Fields("NetBVat").value)
    txtTotalDisc.Text = IIf(IsNull(RsSavRec.Fields("TotalDisc").value), "", RsSavRec.Fields("TotalDisc").value)


    
    If RsSavRec("IsCanceld").value = vbTrue Then
        Me.chkIsCanceld.value = vbChecked
    Else
        Me.chkIsCanceld.value = vbUnchecked
    End If
    
   
    If RsSavRec("IsSupply").value = vbTrue Then
        Me.chkIsSupply.value = vbChecked
    Else
        Me.chkIsSupply.value = vbUnchecked
    End If
    chkIsSupply_Click
    
    TxtVAt2.Text = IIf(IsNull(RsSavRec.Fields("VAt2").value), "", RsSavRec.Fields("VAt2").value)
    
    TxtNoteID = RsSavRec!NoteID & ""
    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
 

     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60

FullGridData

CalcTotal
RelinGrid

 
    If val(TxtNoteID) <> 0 Then
        CmdCreateV.Enabled = False
        Command9.Enabled = True
        Command2.Enabled = True

     Else
        CmdCreateV.Enabled = True
        Command9.Enabled = False
        Command2.Enabled = False

    End If
ErrTrap:
End Sub

 Sub GetInformationCustomer(Optional Cus_ID As Double)
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
If Cus_ID <> 0 Then
sql = "select CusID,CusName ,CusNamee,Cus_mobile,Address from TblCustemers where CusID =" & Cus_ID & " "
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
TxtAddress.Text = IIf(IsNull(Rs6("Address").value), "", Rs6("Address").value)
TxtPhone.Text = IIf(IsNull(Rs6("Cus_mobile").value), "", Rs6("Cus_mobile").value)
'TxtCusID.Text = IIf(IsNull(Rs6("CusID").Value), "", Rs6("CusID").Value)
Else
'TxtCusID = ""
'TxtRecordNo = ""
'DcCustomerType.BoundText = ""
End If
End If
End Sub


Private Sub DcbCus_Change()
DcbCus_Click (0)
End Sub

Private Sub DcbCus_Click(Area As Integer)
  If val(DcbCus.BoundText) = 0 Then Exit Sub
  

    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCus.BoundText, EmpCode
    Me.Text9.Text = EmpCode
 
If Me.TxtModFlg.Text <> "R" Then
If val(DcbCus.BoundText) <> 0 Then

GetInformationCustomer (DcbCus.BoundText)

End If
End If
End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim Frm As New FrmDateOpProject

Select Case GridInstallments.ColKey(Col)
Case "InstallPrice", "SalPrice", "Qun"
    Dim mQty As Double, mSalPrice  As Double, mInstallPrice As Double
    mQty = val(GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("Qun")))
    mSalPrice = val(GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("SalPrice")))
    mInstallPrice = val(GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("InstallPrice")))
    GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("TotalInstall")) = val(mQty) * val(mInstallPrice)
    GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("TotalPrice")) = val(mQty) * val((mSalPrice))
    GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("value")) = val(GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("TotalInstall"))) + val(GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("TotalPrice")))
    
    CalcTotal
End Select

End Sub

Private Sub GridInstallments_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
 
 Select Case GridInstallments.ColKey(Col)
    Case "FinishDate"
         
        Dim Frm As New FrmDateOpProject
        Frm.Index = 35
        Me.LngRow = Row
        Frm.show 1
    End Select
        
End Sub

Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case GridInstallments.ColKey(Col)
    Case "UintMeas", "Value", "TotalPrice", "TotalInstall"
        Cancel = True
    Case "InstallPrice", "SalPrice", "Qun"
        GridInstallments.EditMaxLength = 10
    Case "specificationAr"
        GridInstallments.EditMaxLength = 3000
    End Select
End Sub

Private Sub ISButton2_Click()
print_report2
End Sub

Private Sub ISButton3_Click()

If val(DcboUnits.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ÊÕœ… «·ﬁÌ«” "
Else
MsgBox "Please Enter Unit Measure"
End If
DcboUnits.SetFocus
Exit Sub
End If

If (Txt_Qun.Text) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·ﬂ„Ì… "
Else
MsgBox "Please Enter Quantity "
End If
Txt_Qun.SetFocus
Exit Sub
End If

If (Txt_SalPrice.Text) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ”⁄— «· Ê—Ìœ "
Else
MsgBox "Please Enter Sales Price "
End If
Txt_SalPrice.SetFocus
Exit Sub
End If
If chkIsSupply = vbUnchecked Then
    If (Txt_Install.Text) = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ ≈œŒ«· ”⁄— «· —ﬂÌ» "
    Else
    MsgBox "Please Enter Price of the installation "
    End If
    Txt_Install.SetFocus
    Exit Sub
    End If
End If
If (Txt_Value.Text) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·ﬁÌ„… "
Else
MsgBox "Please Enter Value "
End If
Txt_Value.SetFocus
Exit Sub
End If


If (DC_specific.Text) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·„Ê«’›«  "
Else
MsgBox "Please Enter Specification"
End If
DC_specific.SetFocus
Exit Sub
End If
filgrid1
RelinGrid
CalcTotal
DC_specific.Text = ""
DcboUnits.Text = ""
Txt_Qun.Text = ""
Txt_SalPrice.Text = ""
Txt_Install.Text = ""
Txt_Value.Text = ""
Txt_specific_Ar.Text = ""

DC_specific.SetFocus

End Sub

 

Private Sub ISButton5_Click()
print_report
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
      
           
     If val(DcbCus.BoundText) = 0 Or DcbCus.Text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·⁄„Ì·  "
     Else
     MsgBox "Please Select Customer"
     End If
     DcbCus.SetFocus
     Exit Sub
     End If

    
    'With Me.GridInstallments
     '      If .Rows >= 2 Then
      '     If val(.TextMatrix(1, .ColIndex("ID"))) = 0 Then
       '    If SystemOptions.UserInterface = ArabicInterface Then
        '      MsgBox "Ì—ÃÏ «Œ Ì«— «·»Ì«‰«  "
         '  Else
          ' MsgBox "Please Enter Data"
           'End If
           'Exit Sub
           'End If
           'End If
           'If .Rows < 2 Then
           'If SystemOptions.UserInterface = ArabicInterface Then
           '  MsgBox "Ì—ÃÏ «Œ Ì«— «·»Ì«‰«  "
           'Else
           'MsgBox "Please Enter Data"
           'End If
           'Exit Sub
           'End If
   ' End With
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
    StrRecID = new_id("Tbl_TradingContract", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


 Sub FullGridData()
 'On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
        GridInstallments.Rows = 1



sql = "SELECT dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.FinishDate,dbo.Tbl_TradingContractDet.Periods, dbo.Tbl_TradingContractDet.TContractDet_TContractID, dbo.Tbl_TradingContractDet.TContractDet_specification,"
  sql = sql & "              dbo.Tbl_TradingContractDet.ProcessDEFID, dbo.Tbl_TradingContractDet.TContractDet_specificationEn, dbo.Tbl_TradingContractDet.TContractDet_UnitID,"
   sql = sql & "             dbo.Tbl_TradingContractDet.TContractDet_Qun, dbo.Tbl_TradingContractDet.TContractDet_SalPrice, dbo.Tbl_TradingContractDet.TContractDet_InstallPrice,"
    sql = sql & "            dbo.Tbl_TradingContractDet.TContractDet_Value, dbo.Tbl_TradingContractDet.TContractDet_TotalSalPrice,"
  sql = sql & "              dbo.Tbl_TradingContractDet.TContractDet_TotalInstallPrice, dbo.Tbl_TradingContractDet.TContractDet_DayMeter, dbo.TblProcessUnites.UnitName,"
    sql = sql & "            dbo.TblProcessUnites.UnitNamee , dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
sql = sql & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
 sql = sql & "               dbo.TblProcessUnites ON dbo.Tbl_TradingContractDet.TContractDet_UnitID = dbo.TblProcessUnites.UnitID lEFT Outer JOIN"
  sql = sql & "              dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"


sql = sql & "   Where (dbo.Tbl_TradingContractDet.TContractDet_TContractID =" & val(TxtSerial1.Text) & ")"

Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     GridInstallments.Rows = GridInstallments.Rows + Rs1.RecordCount
     Dim i As Integer
     With Me.GridInstallments
                For i = .FixedRows To Rs1.RecordCount
                .TextMatrix(i, .ColIndex("Ser")) = i
                              .TextMatrix(i, .ColIndex("ProcessID")) = IIf(IsNull(Rs1("ProcessDEFID").value), 0, Rs1("ProcessDEFID").value)
                .TextMatrix(i, .ColIndex("specification")) = IIf(IsNull(Rs1("ProcessName").value), 0, Rs1("ProcessName").value)
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("TContractDet_UnitID").value), 0, Rs1("TContractDet_UnitID").value)
                .TextMatrix(i, .ColIndex("UintMeas")) = IIf(IsNull(Rs1("UnitName").value), 0, Rs1("UnitName").value)
                .TextMatrix(i, .ColIndex("Qun")) = IIf(IsNull(Rs1("TContractDet_Qun").value), 0, Rs1("TContractDet_Qun").value)
                .TextMatrix(i, .ColIndex("SalPrice")) = IIf(IsNull(Rs1("TContractDet_SalPrice").value), 0, Rs1("TContractDet_SalPrice").value)
                .TextMatrix(i, .ColIndex("InstallPrice")) = IIf(IsNull(Rs1("TContractDet_InstallPrice").value), 0, Rs1("TContractDet_InstallPrice").value)
                .TextMatrix(i, .ColIndex("TotalPrice")) = IIf(IsNull(Rs1("TContractDet_TotalSalPrice").value), 0, Rs1("TContractDet_TotalSalPrice").value)
                .TextMatrix(i, .ColIndex("TotalInstall")) = IIf(IsNull(Rs1("TContractDet_TotalInstallPrice").value), 0, Rs1("TContractDet_TotalInstallPrice").value)
                .TextMatrix(i, .ColIndex("specificationAr")) = IIf(IsNull(Rs1("TContractDet_specificationEn").value), 0, Rs1("TContractDet_specificationEn").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(Rs1("TContractDet_Value").value), 0, Rs1("TContractDet_Value").value)
                .TextMatrix(i, .ColIndex("FinishDate")) = IIf(IsNull(Rs1("FinishDate").value), "", Rs1("FinishDate").value)
                .TextMatrix(i, .ColIndex("Periods")) = IIf(IsNull(Rs1("Periods").value), 0, Rs1("Periods").value)
                          
                
                 Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub ISButton6_Click()
If GridInstallments.Rows <= 1 Then Exit Sub
GridInstallments.RemoveItem GridInstallments.Row

End Sub

Private Sub ISButton8_Click()

FrmProjectSearch.C1Tab1.CurrTab = 4
FrmProjectSearch.Caption = "»ÕÀ ⁄‰ « ›«ﬁÌ…"
FrmProjectSearch.show vbModal


End Sub





Private Sub ISButton9_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.Text, "170420168"
ErrTrap:
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
  End Sub

Private Sub txtNewMeasureNo_KeyPress(KeyAscii As Integer)

If txtNewMeasureNo <> "" Then
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    s = "Select * from TBL_measureMent Where Id = " & val(txtNewMeasureNo)
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
        txtNewMeasureNo = ""
        MsgBox "Â–« «·ﬁÌ«” €Ì— „”Ã· ›Ï Õ—ﬂ… —›⁄ «·„ﬁ«”« "
    End If
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text9.Text, EmpID
        DcbCus.BoundText = EmpID
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

End Sub

   
Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 104
        FrmCustemerSearch.show vbModal
    End If
End Sub

Private Sub Txt_DayMeter_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Txt_Install_Change()
Txt_Value.Text = (val(Txt_Qun.Text) * val(Txt_SalPrice)) + (val(Txt_Qun.Text) * val(Txt_Install.Text))
End Sub

Private Sub Txt_Install_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ISButton3.SetFocus
KeyAscii = KeyAscii_Num(KeyAscii, Me.Txt_Install.Text)

End Sub

Private Sub Txt_Qun_Change()
Txt_Value.Text = (val(Txt_Qun.Text) * val(Txt_SalPrice)) + (val(Txt_Qun.Text) * val(Txt_Install.Text))
End Sub

Private Sub Txt_Qun_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txt_SalPrice.SetFocus
KeyAscii = KeyAscii_Num(KeyAscii, Me.Txt_Qun.Text)
End Sub

Private Sub Txt_SalPrice_Change()
Txt_Value.Text = (val(Txt_Qun.Text) * val(Txt_SalPrice)) + (val(Txt_Qun.Text) * val(Txt_Install.Text))
End Sub

Private Sub Txt_SalPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txt_Install.SetFocus
KeyAscii = KeyAscii_Num(KeyAscii, Me.Txt_SalPrice.Text)

End Sub

Private Sub Txt_specific_Ar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DcboUnits.SetFocus
End Sub

Private Sub Txt_specific_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txt_specific_Ar.SetFocus
End Sub

Private Sub Txt_UnitM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txt_Qun.SetFocus
End Sub

Private Sub Txt_Value_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ISButton3.SetFocus
KeyAscii = KeyAscii_Num(KeyAscii, Me.Txt_Value.Text)
End Sub

Private Sub txtProjectTotal_Change()
CalcTotal
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
               With Me.GridInstallments
      
    End With
          
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
                 
                 StrSQL = "Delete From Tbl_TradingContractDet Where TContractDet_TContractID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          
                                          
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute " Update Tbl_TradingContract set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1.Text)
            
                                                        
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
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
   XPDtbTrans.Enabled = True
       
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
            GridInstallments.Rows = GridInstallments.Rows + 1
'        Me.DCboUserName.BoundText = user_id
GridInstallments.Enabled = True
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
    GridInstallments.Clear flexClearScrollable, flexClearEverything
      GridInstallments.Rows = 1
    Me.DCboUserName.BoundText = user_id
  AddNewRecored
  RecorddateH.value = ToHijriDate(Date)
chkIsSupply = vbUnchecked
chkIsSupply_Click
XPDtbTrans.value = Date
RecorddateH.SetFocus
  
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
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
        
 
 MySQL = "SELECT dbo.Tbl_TradingContract.ID, dbo.TblCustemers.CusName,NetBVat,TotalDisc , dbo.Tbl_TradingContract.TContract_DateH, dbo.Tbl_TradingContract.TContract_Date,dbo.Tbl_TradingContractDet.TContractDet_specificationEn Discr,"
 MySQL = MySQL & "                dbo.Tbl_TradingContract.TOrder_Address, dbo.Tbl_TradingContract.TOrder_Phone, dbo.Tbl_TradingContractDet.TContractDet_Qun,"
 MySQL = MySQL & "                dbo.Tbl_TradingContractDet.TContractDet_SalPrice, dbo.Tbl_TradingContractDet.TContractDet_InstallPrice, dbo.Tbl_TradingContractDet.TContractDet_Value,"
 MySQL = MySQL & "                dbo.Tbl_TradingContractDet.TContractDet_TotalSalPrice, dbo.Tbl_TradingContractDet.TContractDet_TotalInstallPrice,"
  MySQL = MySQL & "         TblProcessDEF.Interval TContractDet_DayMeter, dbo.TblProcessUnites.UnitID, dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee,"
  MySQL = MySQL & "               dbo.Tbl_TradingContractDet.ProcessDEFID , dbo.TblProcessDEF.ProcessName TContractDet_specification, dbo.TblProcessDEF.ProcessNameE,Tbl_TradingContract.VAt2 +Tbl_TradingContract.ProjectTotal - TotalDisc as  TotalWithVat,Tbl_TradingContract.VAt2"
 MySQL = MySQL & "  FROM  dbo.Tbl_TradingContract INNER JOIN"
 MySQL = MySQL & "                 dbo.Tbl_TradingContractDet ON dbo.Tbl_TradingContract.ID = dbo.Tbl_TradingContractDet.TContractDet_TContractID INNER JOIN"
 MySQL = MySQL & "                dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID INNER JOIN"
 MySQL = MySQL & "                dbo.TblProcessUnites ON dbo.Tbl_TradingContractDet.TContractDet_UnitID = dbo.TblProcessUnites.UnitID "

MySQL = MySQL & "                        Left Outer JOIN dbo.TblProcessDEF"
MySQL = MySQL & "                       ON Tbl_TradingContractDet.ProcessDEFID =TblProcessDEF.TblProcessDEFID"
 
MySQL = MySQL & "  Where (dbo.Tbl_TradingContract.ID =" & val(TxtSerial1.Text) & ")"
  
        
        If chkIsSupply Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TradingContact2.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TradingContact2.rpt"
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TradingContact.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TradingContact.rpt"
            End If
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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

    Dim i As Integer
    For i = 1 To xReport.FormulaFields.count
         Select Case xReport.FormulaFields.Item(i).Name
         Case "{@TotalValue}"
                xReport.FormulaFields.Item(i).Text = "'" & lblTotalNet & "'"
         End Select
     Next i

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

Function print_report2(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  MySQL = "SELECT     dbo.Tbl_TradingContract.ID, dbo.TblCustemers.CusName, dbo.Tbl_TradingContract.TContract_DateH, dbo.Tbl_TradingContract.TContract_Date, "

  MySQL = MySQL & "                     dbo.Tbl_TradingContract.TOrder_Address, dbo.Tbl_TradingContract.TOrder_Phone, dbo.Tbl_TradingContractDet.TContractDet_specification,dbo.Tbl_TradingContractDet.TContractDet_specificationEn, "

  MySQL = MySQL & "                     dbo.Tbl_TradingContractDet.TContractDet_UnitID,dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Tbl_TradingContractDet.TContractDet_Qun, dbo.Tbl_TradingContractDet.TContractDet_SalPrice, "

  MySQL = MySQL & "                     dbo.Tbl_TradingContractDet.TContractDet_InstallPrice, dbo.Tbl_TradingContractDet.TContractDet_Value, dbo.Tbl_TradingContractDet.TContractDet_TotalSalPrice, dbo.Tbl_TradingContractDet.TContractDet_TotalInstallPrice,dbo.Tbl_TradingContractDet.TContractDet_DayMeter"

  MySQL = MySQL & " FROM         dbo.Tbl_TradingContract INNER JOIN"
  MySQL = MySQL & "                    dbo.Tbl_TradingContractDet ON dbo.Tbl_TradingContract.ID = dbo.Tbl_TradingContractDet.TContractDet_TContractID INNER JOIN"
  MySQL = MySQL & "  dbo.TblUnites ON dbo.Tbl_TradingContractDet.TContractDet_UnitID = dbo.TblUnites.UnitID INNER JOIN"
  MySQL = MySQL & "                     dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID"
 
  MySQL = MySQL & "  Where (dbo.Tbl_TradingContract.ID =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_ProductionS.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_ProductionS.rpt"
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
  Dim XPic As IPictureDisp
  
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Me.Caption = "Conventions"
         'FraHeader.Caption = Me.Caption
         Label1(2).Caption = Me.Caption


    Me.btnNew.Caption = "New"
    Me.btnModify.Caption = "Edit"
    Me.btnSave.Caption = "Save"
    Me.BtnUndo.Caption = "Undo"
    Me.btnDelete.Caption = "Delete"
    ISButton8.Caption = "Search"
    Me.btnCancel.Caption = "Exit"
    Me.ISButton5.Caption = "Print"
    
    Label2(0) = "Cuurent Record"
    Label2(1) = "Record Count"


    
    Cmd_DeleteRow.Caption = "Delete a line"
    Cmd_DeleteAll.Caption = "Delete all"
lbl(8) = "Edited By"
         
lbl(4).Caption = "Order No"
lbl(7) = "Hijri date"
lbl(2) = "date"
lbl(1) = "Location"
lbl(0) = "Responsible compound"
lbl(3) = "Start Date"
lbl(5) = "Period"
lbl(6) = "Total"
lbl(9) = "Workers"
Label1(1) = "Customer"
lbl(16) = "Address"
lbl(17) = "Tel"
Label1(3) = "Specifications"
Label1(7) = "Description"
lbl(11) = "Unit"
lbl(12) = "Qty"
Label1(5) = "Instalation Price"
Label1(0) = "Price"
Label1(4) = "Value"
ISButton3.Caption = "Add"

With GridInstallments
.TextMatrix(0, .ColIndex("specification")) = "specification"
.TextMatrix(0, .ColIndex("specificationAr")) = "Description"
.TextMatrix(0, .ColIndex("UintMeas")) = "Uint Meas"
.TextMatrix(0, .ColIndex("Qun")) = "Quntity"
.TextMatrix(0, .ColIndex("SalPrice")) = "Sal Price"
.TextMatrix(0, .ColIndex("InstallPrice")) = "Install Price"
.TextMatrix(0, .ColIndex("TotalInstall")) = "Total Install"
.TextMatrix(0, .ColIndex("TotalPrice")) = "Total Price"
.TextMatrix(0, .ColIndex("Value")) = "Value"


End With
End Sub
Sub filgrid1()
Dim i, k As Integer

Dim Shareval As Double
With GridInstallments
k = .Rows
.Rows = .Rows + 1
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("ID")) = (TxtSerial1.Text)
.TextMatrix(i, .ColIndex("ProcessID")) = DC_specific.BoundText
.TextMatrix(i, .ColIndex("specification")) = DC_specific.Text
.TextMatrix(i, .ColIndex("UnitID")) = (DcboUnits.BoundText)
.TextMatrix(i, .ColIndex("UintMeas")) = (DcboUnits.Text)
.TextMatrix(i, .ColIndex("Qun")) = (Txt_Qun.Text)
.TextMatrix(i, .ColIndex("SalPrice")) = (Txt_SalPrice.Text)
.TextMatrix(i, .ColIndex("InstallPrice")) = Txt_Install.Text
.TextMatrix(i, .ColIndex("TotalInstall")) = val(Txt_Qun.Text) * val(Txt_Install.Text)
.TextMatrix(i, .ColIndex("TotalPrice")) = val(Txt_Qun.Text) * val((Txt_SalPrice.Text))
.TextMatrix(i, .ColIndex("Periods")) = val(txtPeriod2)

.TextMatrix(i, .ColIndex("specificationAr")) = (Txt_specific_Ar.Text)

.TextMatrix(i, .ColIndex("Value")) = Txt_Value.Text

Next i
'.AutoSize 0, .Cols - 1, False
End With
End Sub
Sub RelinGrid()
Dim Sm, summation, AllTotal As Double
Dim Counter As Integer
Dim i As Integer
Counter = 0
Sm = 0
summation = 0

With Me.GridInstallments
For i = 1 To .Rows - 1
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("SalPrice")))
Sm = Sm + val(.TextMatrix(i, .ColIndex("InstallPrice")))

Next i
End With
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
   Set rs = New ADODB.Recordset
   My_SQL = "Tbl_TradingContract"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub txtTotalDisc_Change()
    CalCulteVAT 3
End Sub

 Private Sub XPDtbTrans_Change()

If Me.TxtModFlg.Text <> "R" Then
     
         RecorddateH.value = ToHijriDate(XPDtbTrans.value)
         XPDtbTrans.value = (XPDtbTrans.value)
  
    If ChekSanNumber(Current_branch, 60) = True Then
          TxtSerial1.Text = ""
      End If
      TxtSerial1.Text = ""
      CalcTotal
End If

End Sub


Private Sub CalcTotal()
    With GridInstallments
    .IsSubtotal(.Rows - 1) = True
    Dim SngTotal As Single
    If .Rows > 1 Then
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        txtTotalInstall = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalInstall"), .Rows - 1, .ColIndex("TotalInstall"))
        txtTotalPrice = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalPrice"), .Rows - 1, .ColIndex("TotalPrice"))
        
    End If
    
    txtProjectTotal = SngTotal
    txtProjectTotal2 = SngTotal
    CalCulteVAT 3
    End With
End Sub
Sub CalCulteVAT(Optional Ind As Integer = 0)
Dim AccountVATCreit As String
Dim Percetage As Double

Dim mVal As Double
    txtNetBVat = val(txtProjectTotal) - val(txtTotalDisc)
    If Ind = 3 Then
        PercentgValueAddedAccount_Transec XPDtbTrans.value, 43, 0, AccountVATCreit, Percetage
        TxtVAt2.Text = val(Format((txtNetBVat.Text), "###.00")) * Percetage / 100
         
         TxtVATValue.Text = val(Format((txtNetBVat.Text), "###.00")) * Percetage / 100
         TxtVAt2.Text = TxtVATValue.Text
         
         
         mVal = val(Format((txtNetBVat.Text), "###.00"))
         TxtVATValue.Text = val(Format((mVal), "###.00")) * Percetage / 100
         txtTotalWithVat.Text = Round(val(Format((mVal), "###.00")) + val(TxtVATValue.Text), 2)
         txtTotalWithVat2.Text = txtTotalWithVat.Text
         
'         Exit Sub
    End If
    'XPDtbTrans.value = 100
    'XPTxtVal = 100
    
     txtTotalWithVat.Text = Round(val(Format((mVal), "###.00")) + val(TxtVATValue.Text), 2)
    txtTotalWithVat2.Text = txtTotalWithVat.Text
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lblTotalNet.Caption = WriteNo(txtTotalWithVat2.Text, 0, True, ".", , 0)
    Else
        Me.lblTotalNet.Caption = WriteNo(txtTotalWithVat2.Text, 0, True, ".", , 1)
    End If
TxtVAt2.Text = TxtVATValue.Text
TxtVAt22.Text = TxtVATValue.Text
End Sub

