VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmContStudent 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13380
   Icon            =   "FrmContStudent.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13380
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
      TabIndex        =   35
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmContStudent.frx":6852
      Left            =   15480
      List            =   "FrmContStudent.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14400
      TabIndex        =   31
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   36
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
      TabIndex        =   37
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
            Picture         =   "FrmContStudent.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContStudent.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   38
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
      ButtonImage     =   "FrmContStudent.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   40
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
      ButtonImage     =   "FrmContStudent.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   41
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
      ButtonImage     =   "FrmContStudent.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9975
      Left            =   0
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   0
      Width           =   13380
      _cx             =   23601
      _cy             =   17595
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
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   540
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   720
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
         Begin VB.TextBox txtid 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11190
            MaxLength       =   50
            TabIndex        =   99
            Top             =   120
            Width           =   1125
         End
         Begin XtremeSuiteControls.RadioButton ContType 
            Height          =   255
            Index           =   0
            Left            =   1335
            TabIndex        =   4
            Top             =   120
            Width           =   1050
            _Version        =   786432
            _ExtentX        =   1852
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "⁄ﬁœ „ œ—»"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5925
            TabIndex        =   0
            Top             =   -240
            Visible         =   0   'False
            Width           =   1725
         End
         Begin XtremeSuiteControls.RadioButton ContType 
            Height          =   255
            Index           =   1
            Left            =   45
            TabIndex        =   5
            Top             =   120
            Width           =   1170
            _Version        =   786432
            _ExtentX        =   2064
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " ⁄ﬁœ ‘—ﬂ« "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   6360
            TabIndex        =   2
            Top             =   120
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   7815
            TabIndex        =   1
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Format          =   93454337
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   2520
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCPreFix1 
            Height          =   315
            Left            =   9960
            TabIndex        =   100
            Top             =   120
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   11
            Left            =   5430
            TabIndex        =   73
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   25
            Left            =   8955
            TabIndex        =   72
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·⁄ﬁœ "
            Height          =   255
            Index           =   4
            Left            =   12510
            TabIndex        =   44
            Top             =   120
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   8880
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
            Left            =   11490
            TabIndex        =   46
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   1590
            _ExtentX        =   2805
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
            ButtonImage     =   "FrmContStudent.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   9765
            TabIndex        =   47
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   1605
            _ExtentX        =   2831
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
            ButtonImage     =   "FrmContStudent.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   7980
            TabIndex        =   30
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   1365
            _ExtentX        =   2408
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
            ButtonImage     =   "FrmContStudent.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   5940
            TabIndex        =   48
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   1890
            _ExtentX        =   3334
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
            ButtonImage     =   "FrmContStudent.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   3990
            TabIndex        =   49
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
            ButtonImage     =   "FrmContStudent.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   2550
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   600
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
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
            ButtonImage     =   "FrmContStudent.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1560
            TabIndex        =   51
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
            Width           =   975
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
            ButtonImage     =   "FrmContStudent.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   0
            TabIndex        =   52
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   600
            Width           =   1395
            _ExtentX        =   2461
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
            ButtonImage     =   "FrmContStudent.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   53
            Top             =   90
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   330
            Left            =   5760
            TabIndex        =   68
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«·„—›ﬁ« "
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
            ButtonImage     =   "FrmContStudent.frx":2A8D1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   315
            TabIndex        =   58
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   62
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
            TabIndex        =   63
            Top             =   240
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
            ButtonImage     =   "FrmContStudent.frx":31133
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   675
            TabIndex        =   64
            Top             =   240
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
            ButtonImage     =   "FrmContStudent.frx":314CD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1350
            TabIndex        =   65
            Top             =   240
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
            ButtonImage     =   "FrmContStudent.frx":31867
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1950
            TabIndex        =   66
            Top             =   240
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
            ButtonImage     =   "FrmContStudent.frx":31C01
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12375
            Picture         =   "FrmContStudent.frx":31F9B
            Stretch         =   -1  'True
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·⁄ﬁÊœ"
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
            TabIndex        =   67
            Top             =   240
            Width           =   4665
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1380
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1200
         Width           =   13455
         _cx             =   23733
         _cy             =   2434
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
         Begin VB.TextBox TxtNoStud 
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
            Left            =   7710
            MaxLength       =   50
            TabIndex        =   116
            Top             =   555
            Width           =   1500
         End
         Begin VB.TextBox TxtDiscount 
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
            Left            =   10470
            MaxLength       =   50
            TabIndex        =   115
            Top             =   960
            Width           =   660
         End
         Begin VB.TextBox TxtContValue 
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
            Left            =   7710
            MaxLength       =   50
            TabIndex        =   113
            Top             =   960
            Width           =   1500
         End
         Begin VB.TextBox TxtStudValue 
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
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   101
            Top             =   555
            Width           =   1140
         End
         Begin VB.TextBox TxtDiscountNet 
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
            MaxLength       =   50
            TabIndex        =   95
            Top             =   1440
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox TxtMidShare 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   19
            Top             =   1440
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ComboBox DcbDiscount 
            Height          =   315
            Left            =   11190
            TabIndex        =   18
            Top             =   960
            Width           =   780
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   17
            Top             =   960
            Width           =   1140
         End
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
            Height          =   315
            Left            =   10830
            MaxLength       =   50
            TabIndex        =   14
            Top             =   135
            Width           =   1140
         End
         Begin MSDataListLib.DataCombo DcbCompany 
            Height          =   315
            Left            =   5520
            TabIndex        =   15
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   135
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbTypeContract 
            Height          =   315
            Left            =   10440
            TabIndex        =   16
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   555
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   1215
            Left            =   0
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   120
            Width           =   5490
            _cx             =   9684
            _cy             =   2143
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
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   615
               Left            =   0
               TabIndex        =   128
               Top             =   120
               Width           =   735
               _Version        =   786432
               _ExtentX        =   1296
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   " ÕœÌÀ «·Õ”«»"
               BackColor       =   12640511
               UseVisualStyle  =   -1  'True
            End
            Begin VB.TextBox TxtAccount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   840
               Width           =   585
            End
            Begin VB.TextBox TxtComm 
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
               Left            =   4215
               MaxLength       =   50
               TabIndex        =   111
               Top             =   840
               Width           =   660
            End
            Begin VB.TextBox TxtMedalMan 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   720
               TabIndex        =   109
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   120
               Width           =   915
            End
            Begin XtremeSuiteControls.RadioButton ChSuperv 
               Height          =   240
               Index           =   0
               Left            =   4140
               TabIndex        =   105
               Top             =   240
               Width           =   1260
               _Version        =   786432
               _ExtentX        =   2222
               _ExtentY        =   423
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton ChSuperv 
               Height          =   255
               Index           =   1
               Left            =   4575
               TabIndex        =   106
               Top             =   480
               Width           =   825
               _Version        =   786432
               _ExtentX        =   1455
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ ⁄«Ê‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmployee 
               Height          =   315
               Left            =   720
               TabIndex        =   107
               Top             =   120
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount 
               Height          =   315
               Left            =   120
               TabIndex        =   126
               Top             =   840
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» «·„‰œÊ»"
               Height          =   285
               Index           =   6
               Left            =   3120
               TabIndex        =   127
               Top             =   840
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "(Ì‘ —ÿ ”⁄ÊœÌ)"
               Height          =   195
               Index           =   19
               Left            =   3360
               TabIndex        =   117
               Top             =   480
               Width           =   1125
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·‰”»…"
               Height          =   195
               Index           =   3
               Left            =   4890
               TabIndex        =   112
               Top             =   885
               Width           =   525
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„‰œÊ»"
               Height          =   195
               Index           =   17
               Left            =   4320
               TabIndex        =   110
               Top             =   0
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁ«∆œ «·„⁄œ…"
               Height          =   270
               Index           =   29
               Left            =   5535
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ì  "
            Height          =   195
            Index           =   18
            Left            =   9210
            TabIndex        =   114
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁÌ„… ··„—‘Õ"
            Height          =   195
            Index           =   8
            Left            =   6690
            TabIndex        =   102
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄«œ·…"
            Height          =   195
            Index           =   6
            Left            =   2160
            TabIndex        =   78
            Top             =   1320
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„’«—Ì›  ”ÊÌﬁ"
            Height          =   195
            Index           =   5
            Left            =   12030
            TabIndex        =   77
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·⁄ﬁœ"
            Height          =   195
            Index           =   4
            Left            =   6690
            TabIndex        =   76
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·„—‘ÕÌ‰"
            Height          =   195
            Index           =   1
            Left            =   9270
            TabIndex        =   75
            Top             =   555
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄ﬁœ"
            Height          =   195
            Index           =   0
            Left            =   12030
            TabIndex        =   74
            Top             =   570
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ…"
            Height          =   195
            Index           =   0
            Left            =   12030
            TabIndex        =   71
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Ê”Ìÿ"
            Height          =   195
            Index           =   22
            Left            =   3150
            TabIndex        =   70
            Top             =   1395
            Visible         =   0   'False
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1380
         Left            =   0
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1200
         Width           =   13455
         _cx             =   23733
         _cy             =   2434
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
         Begin VB.TextBox TxtDiscountNet1 
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
            Left            =   4560
            TabIndex        =   124
            Top             =   960
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox TxtValTrin 
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
            Left            =   9960
            TabIndex        =   122
            Top             =   960
            Width           =   2010
         End
         Begin VB.ComboBox DcbDiscount1 
            Height          =   315
            Left            =   7485
            TabIndex        =   119
            Top             =   960
            Width           =   780
         End
         Begin VB.TextBox TxtDiscount1 
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
            Left            =   6720
            MaxLength       =   50
            TabIndex        =   118
            Top             =   960
            Width           =   660
         End
         Begin VB.TextBox TxtUQama 
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
            TabIndex        =   8
            Top             =   120
            Width           =   3090
         End
         Begin VB.TextBox TxtSudCode 
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
            Left            =   9960
            TabIndex        =   6
            Top             =   120
            Width           =   2010
         End
         Begin VB.ComboBox Cash_Insta 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   3090
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            TabIndex        =   9
            Top             =   555
            Width           =   2010
         End
         Begin VB.TextBox TxtPrice1 
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
            Left            =   4560
            TabIndex        =   12
            Top             =   960
            Width           =   1170
         End
         Begin MSDataListLib.DataCombo DcbQuali 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   555
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCurs 
            Height          =   315
            Left            =   4560
            TabIndex        =   10
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   555
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbStudent 
            Height          =   315
            Left            =   4560
            TabIndex        =   7
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁÌ„…"
            Height          =   195
            Index           =   22
            Left            =   12030
            TabIndex        =   123
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "(„Ã„Ê⁄« )"
            Height          =   195
            Index           =   21
            Left            =   8280
            TabIndex        =   121
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’„ „ œ—»"
            Height          =   195
            Index           =   20
            Left            =   9120
            TabIndex        =   120
            Top             =   960
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÂÊÌ…"
            Height          =   285
            Index           =   1
            Left            =   3390
            TabIndex        =   94
            Top             =   165
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ—Ìﬁ…«·œ›⁄"
            Height          =   195
            Index           =   7
            Left            =   3120
            TabIndex        =   85
            Top             =   1005
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ƒÂ·« "
            Height          =   360
            Index           =   3
            Left            =   3360
            TabIndex        =   84
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·›Ê‰"
            Height          =   360
            Index           =   2
            Left            =   12030
            TabIndex        =   83
            Top             =   555
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ œ—»"
            Height          =   195
            Index           =   12
            Left            =   12030
            TabIndex        =   82
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„«œ…"
            Height          =   195
            Index           =   11
            Left            =   8370
            TabIndex        =   81
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ì"
            Height          =   195
            Index           =   9
            Left            =   5550
            TabIndex        =   80
            Top             =   960
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   3300
         Left            =   0
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   2520
         Width           =   13455
         _cx             =   23733
         _cy             =   5821
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
         Begin VSFlex8Ctl.VSFlexGrid fg1 
            Height          =   2955
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   13185
            _cx             =   23257
            _cy             =   5212
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmContStudent.frx":333A0
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   3060
         Left            =   0
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   5760
         Width           =   13455
         _cx             =   23733
         _cy             =   5398
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
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   360
            Width           =   3345
         End
         Begin VB.ComboBox TypePeriod 
            Height          =   315
            ItemData        =   "FrmContStudent.frx":334D8
            Left            =   9720
            List            =   "FrmContStudent.frx":334E5
            TabIndex        =   24
            Top             =   465
            Width           =   1095
         End
         Begin VB.TextBox TxtPeriod 
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
            Height          =   300
            Left            =   10920
            MaxLength       =   50
            TabIndex        =   23
            Top             =   465
            Width           =   1065
         End
         Begin VB.TextBox TxtNoInstal 
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
            Height          =   300
            Left            =   10920
            MaxLength       =   50
            TabIndex        =   20
            Top             =   120
            Width           =   1065
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   240
            Index           =   2
            Left            =   4800
            TabIndex        =   27
            Top             =   465
            Width           =   735
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ— ﬁ”ÿ"
            Height          =   240
            Index           =   3
            Left            =   5640
            TabIndex        =   26
            Top             =   465
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ê· ﬁ”ÿ"
            Height          =   240
            Index           =   4
            Left            =   6840
            TabIndex        =   25
            Top             =   465
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker FirstDate 
            Height          =   255
            Left            =   8280
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   93454339
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal FirstDateH 
            Height          =   240
            Left            =   6840
            TabIndex        =   22
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   1845
            Left            =   0
            TabIndex        =   96
            Top             =   840
            Width           =   13245
            _cx             =   23363
            _cy             =   3254
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmContStudent.frx":334F8
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   420
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
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
            ButtonImage     =   "FrmContStudent.frx":335E6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker TempDate 
            Height          =   255
            Left            =   120
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   93454339
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   97
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   270
            Index           =   16
            Left            =   2640
            TabIndex        =   93
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
            Height          =   270
            Index           =   15
            Left            =   11880
            TabIndex        =   92
            Top             =   465
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
            Height          =   270
            Index           =   14
            Left            =   9600
            TabIndex        =   91
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·œ›⁄« "
            Height          =   270
            Index           =   10
            Left            =   12240
            TabIndex        =   90
            Top             =   120
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   37
            Left            =   8040
            TabIndex        =   89
            Top             =   465
            Width           =   1455
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
      TabIndex        =   39
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmContStudent"
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


Private Sub Cash_Insta_Change()
If Me.TxtModFlg.Text <> "R" Then
If ContType(0).value = True Then
C1Elastic5.Enabled = False
If val(Cash_Insta.ListIndex) = 1 Then
C1Elastic5.Enabled = True
End If
End If
End If
End Sub

Private Sub Cash_Insta_Click()
Cash_Insta_Change
End Sub



Private Sub ChSuperv_Click(Index As Integer)
If ChSuperv(0).value = True Then
Text6.Enabled = True
PushButton1.Enabled = False
DcEmployee.Enabled = True
DcbAccount.Enabled = False
DcbAccount.BoundText = 0
TxtMedalMan.Enabled = False
TxtAccount.Enabled = False
TxtAccount.Text = ""
TxtMedalMan.Text = ""
ElseIf ChSuperv(1).value = True Then
Text6.Enabled = False
PushButton1.Enabled = True
DcbAccount.Enabled = True
DcEmployee.Enabled = False
TxtMedalMan.Enabled = True
TxtAccount.Enabled = True
DcEmployee.BoundText = 0
Text6.Text = ""
End If
End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 29123

    End If
End Sub

Private Sub DcbCompany_Change()
DcbCompany_Click (0)
End Sub
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblContrStudent", "ID", "")
    RsSavRec.AddNew
     Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    TxtSerial1.Text = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Private Sub DcbCompany_Click(Area As Integer)
  If val(DcbCompany.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany.BoundText, EmpCode
    Me.txtCode.Text = EmpCode
End Sub

Private Sub ContType_Click(Index As Integer)
C1Elastic2.Visible = False
C1Elastic3.Visible = False
If Me.ContType(1).value = True Then
C1Elastic2.Visible = True
C1Elastic5.Enabled = True
ElseIf Me.ContType(0).value = True Then
C1Elastic3.Visible = True
End If

End Sub

Private Sub DcbCompany_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   FrmCustemerSearch.SearchType = 19
        FrmCustemerSearch.show vbModal
  End If
End Sub

Private Sub DcbDiscount_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbDiscount.ListIndex) = 1 Then
TxtDiscountNet.Text = val(TxtStudValue.Text) * (val(TxtDiscount.Text) / 100)
Else
TxtDiscountNet.Text = val(TxtDiscount.Text)
End If
TxtMidShare.Text = Round((val(TxtComm.Text) * (val(TxtContValue.Text) - val(TxtDiscountNet.Text))) / 100, 2)
TxtContValue.Text = val(TxtStudValue.Text) - val(TxtDiscountNet.Text)
txtPrice.Text = val(TxtContValue.Text) * val(TxtNoStud.Text)
End If
End Sub

Private Sub DcbDiscount_Click()
DcbDiscount_Change
End Sub

Private Sub DcbDiscount1_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbDiscount1.ListIndex) = 1 Then
TxtDiscountNet1.Text = val(TxtValTrin.Text) * (val(TxtDiscount1.Text) / 100)
Else
TxtDiscountNet1.Text = val(TxtDiscount1.Text)
End If
TxtPrice1.Text = val(TxtValTrin.Text) - val(TxtDiscountNet1.Text)
End If
End Sub

Private Sub DcbDiscount1_Click()
DcbDiscount1_Change
End Sub

Private Sub DcbStudent_Change()
DcbStudent_Click (0)
End Sub

Private Sub DcbStudent_Click(Area As Integer)
Dim QuliID As Double
Dim UQama As String
Dim phone As String
  If val(DcbStudent.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode DcbStudent.BoundText, EmpCode
    Me.TxtSudCode.Text = EmpCode
   If Me.TxtModFlg.Text <> "R" Then
   GetInformationofStudent val(DcbStudent.BoundText), UQama, phone, QuliID
   TxtUQama.Text = UQama
   TxtPhone.Text = phone
   DcbQuali.BoundText = QuliID
   End If
End Sub

Sub FallGridStudent()
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim i As Integer

Dim sql As String
sql = "SELECT     dbo.TblStuCandidacyDet.ID, dbo.TblStuCandidacyDet.AccptedID, dbo.TblStuCandidacy.ContNoID, dbo.TblStuCandidacyDet.StuID, dbo.TblStudent.Name, "
sql = sql & "                      dbo.TblStudent.NameE, dbo.TblStudent.FullCode, dbo.TblStudent.UQama, dbo.TblStudent.SuperPhone, dbo.TblStudent.SuperVisorName,"
sql = sql & "                      dbo.TblStudent.StudentAddres, dbo.TblStudent.DateBrith, dbo.TblStudent.DateBrithH, dbo.TblStudent.StudentPhone, dbo.TblStudent.StudentEmail,"
sql = sql & "                      dbo.TblStudent.Remarks, dbo.TblStudent.SexID, dbo.TblStudentQualification.Name AS QName, dbo.TblStudentQualification.NameE AS QNameE"
sql = sql & " FROM         dbo.TblStudentQualification RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStudent ON dbo.TblStudentQualification.ID = dbo.TblStudent.DcbQualiID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacyDet ON dbo.TblStudent.ID = dbo.TblStuCandidacyDet.StuID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacy ON dbo.TblStuCandidacyDet.StudCandID = dbo.TblStuCandidacy.ID"
sql = sql & "   WHERE     (dbo.TblStuCandidacyDet.AccptedID = 1) AND (dbo.TblStuCandidacy.ContNoID = " & val(TxtSerial1.Text) & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Rs7.MoveFirst
FG1.Rows = 1
With FG1
.Rows = .Rows + Rs7.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs7("FullCode").value), "", Rs7("FullCode").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Qualification")) = IIf(IsNull(Rs7("QName").value), "", Rs7("QName").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs7("Name").value), "", Rs7("Name").value)
Else
.TextMatrix(i, .ColIndex("Qualification")) = IIf(IsNull(Rs7("QNameE").value), "", Rs7("QNameE").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs7("NameE").value), "", Rs7("NameE").value)
End If

.TextMatrix(i, .ColIndex("Supervisor")) = IIf(IsNull(Rs7("SuperVisorName").value), "", Rs7("SuperVisorName").value)
.TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(Rs7("StudentPhone").value), "", Rs7("StudentPhone").value)
.TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(Rs7("UQama").value), "", Rs7("UQama").value)
Rs7.MoveNext
Next i
End With
End If
End Sub

Private Sub DcbStudent_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 101
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub DcEmployee_Change()
DcEmployee_Click (0)
End Sub

Private Sub DcEmployee_Click(Area As Integer)
    If val(DcEmployee.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcEmployee.BoundText, EmpCode
      Text6.Text = EmpCode
End Sub

Private Sub ISButton8_Click()
FrmSearStudent.inde = 4
Load FrmSearStudent
FrmSearStudent.show vbModal
End Sub

Private Sub PushButton1_Click()
If Me.TxtModFlg.Text = "R" Then
If DcbAccount.BoundText <> "" Then
Cn.Execute "Update TblContrStudent set TypeSuper=1, MedalMan='" & TxtMedalMan.Text & "', AccountCode='" & DcbAccount.BoundText & "' where  ID =" & val(TxtSerial1.Text) & ""
RsSavRec.Resync
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «· ÕœÌÀ"
Else
MsgBox "Update Successfully"
End If
End If
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DcEmployee.BoundText = EmpID
    End If
End Sub
Private Sub FirstDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         FirstDateH.value = ToHijriDate(FirstDate.value)
End If
End Sub
Sub filgrid1()
Dim Price As Double
Dim i As Integer
 GridInstallments.Clear flexClearScrollable, flexClearEverything
GridInstallments.Rows = 1
With GridInstallments
If ContType(0).value Then
Price = val(TxtPrice1.Text)
Else
Price = val(txtPrice.Text)
End If
.Rows = .Rows + val(TxtNoInstal.Text)
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i

.TextMatrix(i, .ColIndex("InstalValue")) = Price / val(TxtNoInstal.Text)
.TextMatrix(i, .ColIndex("InstalValue")) = Round(val(.TextMatrix(i, .ColIndex("InstalValue"))), 2)
.TextMatrix(i, .ColIndex("InstalNo")) = i
If i = 1 Then
.TextMatrix(i, .ColIndex("InstalDate")) = FirstDate.value
.TextMatrix(i, .ColIndex("InstalDateH")) = FirstDateH.value
TempDate.value = FirstDate.value
Else
If val(Me.TypePeriod.ListIndex) = 0 Then
TempDate.value = DateAdd("d", val(Me.txtPeriod.Text), TempDate.value)
ElseIf val(Me.TypePeriod.ListIndex) = 1 Then
TempDate.value = DateAdd("M", val(Me.txtPeriod.Text), TempDate.value)
ElseIf val(Me.TypePeriod.ListIndex) = 1 Then
TempDate.value = DateAdd("YYYY", val(Me.txtPeriod.Text), TempDate.value)
End If
.TextMatrix(i, .ColIndex("InstalDate")) = TempDate.value
.TextMatrix(i, .ColIndex("InstalDateH")) = ToHijriDate(TempDate.value)
End If
.TextMatrix(i, .ColIndex("Remarks")) = TxtRemarks.Text
Next i

End With
End Sub

Private Sub FirstDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 FirstDate.value = ToGregorianDate(FirstDateH.value)
End If
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblContrStudent  "
      conection = conection & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    conection = conection & " Order By ID"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
     Dim Dcombos As New ClsDataCombos
     Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Dcombos.GetCodeing Me.DCPreFix1, 10
   Dcombos.GetStudentContract Me.DcbTypeContract
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetStudentCurs Me.DcbCurs
   Dcombos.GetEmployees Me.DcEmployee
   Dcombos.GetStudentQualification Me.DcbQuali
   Dcombos.GetStudent Me.DcbStudent, 1
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany
   If SystemOptions.UserInterface = ArabicInterface Then
   With DcbDiscount
   .Clear
   .AddItem "ﬁÌ„…"
   .AddItem "‰”»…"
   End With
      With DcbDiscount1
   .Clear
   .AddItem "ﬁÌ„…"
   .AddItem "‰”»…"
   End With
   With Cash_Insta
   .Clear
   .AddItem "‰ﬁœ«"
   .AddItem "œ›⁄« "
   End With
   Else
    With DcbDiscount1
   .Clear
   .AddItem "Value"
   .AddItem "percentage"
   End With
    With DcbDiscount
   .Clear
   .AddItem "Value"
   .AddItem "percentage"
   End With
      With Cash_Insta
   .Clear
   .AddItem "Cash"
   .AddItem "Installments"
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
Sub relod()
    Dim Dcombos As New ClsDataCombos
    Dcombos.ClearMyDataCombo Me.DcbStudent
   Dcombos.GetStudent Me.DcbStudent
End Sub
Sub relod1()
    Dim Dcombos As New ClsDataCombos
    Dcombos.ClearMyDataCombo Me.DcbStudent
   Dcombos.GetStudent Me.DcbStudent, 1
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    Dim Msg As String
                RsSavRec.Fields("code").value = txtid.Text
                RsSavRec.Fields("prifix").value = IIf(DCPreFix1.Text = "", Null, DCPreFix1.Text)
               RsSavRec.Fields("Fullcode").value = IIf(DCPreFix1.BoundText = "", Null, DCPreFix1.Text) & IIf(Trim(txtid.Text) = "", Null, txtid.Text)
   RsSavRec.Fields("RecordDateH").value = RecorddateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("AccountCode").value = DcbAccount.BoundText
   RsSavRec.Fields("CompID").value = val(DcbCompany.BoundText)
   RsSavRec.Fields("StudeID").value = val(DcbStudent.BoundText)
   RsSavRec.Fields("CursID").value = val(Me.DcbCurs.BoundText)
   RsSavRec.Fields("ValTrin").value = val(Me.TxtValTrin.Text)
   If ContType(0).value = True Then
   RsSavRec.Fields("Price").value = val(Me.TxtPrice1.Text)
   RsSavRec.Fields("DiscountNet").value = val(Me.TxtDiscountNet1.Text)
   RsSavRec.Fields("Discount").value = val(Me.TxtDiscount1.Text)
   RsSavRec.Fields("TypeDis").value = val(Me.DcbDiscount1.ListIndex)
   RsSavRec.Fields("ContType").value = 0
   Else
     RsSavRec.Fields("DiscountNet").value = val(Me.TxtDiscountNet.Text)
   RsSavRec.Fields("Discount").value = val(Me.TxtDiscount.Text)
   RsSavRec.Fields("TypeDis").value = val(Me.DcbDiscount.ListIndex)
   RsSavRec.Fields("Price").value = val(Me.txtPrice.Text)
   RsSavRec.Fields("ContType").value = 1
   End If
   RsSavRec.Fields("ContValue").value = val(Me.TxtContValue.Text)
   RsSavRec.Fields("Cash_Insta").value = val(Me.Cash_Insta.ListIndex)
   RsSavRec.Fields("TypeContID").value = val(Me.DcbTypeContract.BoundText)
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("QualiID").value = val(Me.DcbQuali.BoundText)
   RsSavRec.Fields("Phone").value = (Me.TxtPhone.Text)
   RsSavRec.Fields("UQama").value = TxtUQama.Text
   If Opt(4).value = True Then
   RsSavRec.Fields("TypePayment").value = 4
   ElseIf Opt(3).value = True Then
   RsSavRec.Fields("TypePayment").value = 3
   ElseIf Opt(2).value = True Then
   RsSavRec.Fields("TypePayment").value = 2
   Else
   RsSavRec.Fields("TypePayment").value = Null
   End If
   RsSavRec.Fields("NoStud").value = val(Me.TxtNoStud.Text)
   RsSavRec.Fields("Comm").value = val(Me.TxtComm.Text)
   RsSavRec.Fields("MidShare").value = val(Me.TxtMidShare.Text)
   RsSavRec.Fields("Remarks").value = (Me.TxtRemarks.Text)
   RsSavRec.Fields("NoInstal").value = val(Me.TxtNoInstal.Text)
   RsSavRec.Fields("Period").value = val(Me.txtPeriod.Text)
   RsSavRec.Fields("TypePeriod").value = val(Me.TypePeriod.ListIndex)
   RsSavRec.Fields("FirstDateH").value = FirstDateH.value
   RsSavRec.Fields("FirstDate").value = FirstDate.value
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("MedalMan").value = (Me.TxtMedalMan.Text)
   RsSavRec.Fields("TotalPayment").value = val(Me.lbl(5).Caption)
  RsSavRec.Fields("StudValue").value = val(TxtStudValue.Text)
   RsSavRec.Fields("EmpID").value = val(Me.DcEmployee.BoundText)
   If Me.ChSuperv(1).value = True Then
   RsSavRec.Fields("TypeSuper").value = 1
   End If
   RsSavRec.update
  ''//////////////////////////
  If Me.TxtModFlg.Text = "E" Then
 Cn.Execute "delete from TblContrStudentDet where ContStudID=" & val(TxtSerial1.Text) & " "
 End If
  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblContrStudentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("InstalNo"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ContStudID").value = val(Me.TxtSerial1.Text)
                RsDevsub("InstalDateH").value = IIf((.TextMatrix(i, .ColIndex("InstalDateH"))) = "", Null, .TextMatrix(i, .ColIndex("InstalDateH")))
                RsDevsub("InstalDate").value = IIf((.TextMatrix(i, .ColIndex("InstalDate"))) = "", Null, .TextMatrix(i, .ColIndex("InstalDate")))
                If ContType(1).value = True Then
                If Opt(3).value = True And i = (.Rows - 1) Then
               .TextMatrix(i, .ColIndex("InstalValue")) = val(.TextMatrix(i, .ColIndex("InstalValue"))) + val(txtPrice.Text) - val(lbl(5).Caption)
                ElseIf Opt(4).value = True And i = 1 Then
               .TextMatrix(i, .ColIndex("InstalValue")) = val(.TextMatrix(i, .ColIndex("InstalValue"))) + val(txtPrice.Text) - val(lbl(5).Caption)
                End If
                Else
                If Opt(3).value = True And i = (.Rows - 1) Then
               .TextMatrix(i, .ColIndex("InstalValue")) = val(.TextMatrix(i, .ColIndex("InstalValue"))) + val(TxtPrice1.Text) - val(lbl(5).Caption)
                ElseIf Opt(4).value = True And i = 1 Then
               .TextMatrix(i, .ColIndex("InstalValue")) = val(.TextMatrix(i, .ColIndex("InstalValue"))) + val(TxtPrice1.Text) - val(lbl(5).Caption)
                End If
                End If
                RsDevsub("InstalValue").value = IIf((.TextMatrix(i, .ColIndex("InstalValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("InstalValue"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, .TextMatrix(i, .ColIndex("Remarks")))
                RsDevsub("InstalNo").value = IIf((.TextMatrix(i, .ColIndex("InstalNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("InstalNo"))))
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
Sub FullGri()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 2
sql = " SELECT     * "
sql = sql & " From dbo.TblContrStudentDet"
sql = sql & " Where (ContStudID =" & val(TxtSerial1.Text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With GridInstallments
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows
.TextMatrix(i, .ColIndex("InstalNo")) = IIf(IsNull(Rs3("InstalNo").value), 0, Rs3("InstalNo").value)
.TextMatrix(i, .ColIndex("InstalDate")) = IIf(IsNull(Rs3("InstalDate").value), Date, Rs3("InstalDate").value)
.TextMatrix(i, .ColIndex("InstalDateH")) = IIf(IsNull(Rs3("InstalDateH").value), ToHijriDate(Date), Rs3("InstalDateH").value)
.TextMatrix(i, .ColIndex("InstalValue")) = IIf(IsNull(Rs3("InstalValue").value), 0, Rs3("InstalValue").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs3("Remarks").value), "", Rs3("Remarks").value)
Rs3.MoveNext
Next i
End With
End If
End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()

   On Error GoTo ErrTrap
     FG1.Clear flexClearScrollable, flexClearEverything
     FG1.Rows = 2
     GridInstallments.Clear flexClearScrollable, flexClearEverything
     GridInstallments.Rows = 2
     
    Dim i As Integer
    Dim Shifttime As Date
    DcbAccount.BoundText = IIf(IsNull(RsSavRec.Fields("AccountCode").value), "", RsSavRec.Fields("AccountCode").value)
    TxtValTrin.Text = IIf(IsNull(RsSavRec.Fields("ValTrin").value), 0, RsSavRec.Fields("ValTrin").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbCompany.BoundText = IIf(IsNull(RsSavRec.Fields("CompID").value), "", RsSavRec.Fields("CompID").value)
    RecorddateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    DCPreFix1.Text = IIf(IsNull(RsSavRec("prifix").value), "", RsSavRec("prifix").value)
    Me.txtid.Text = IIf(IsNull(RsSavRec("code").value), "", RsSavRec("code").value)
    Me.DcbStudent.BoundText = IIf(IsNull(RsSavRec.Fields("StudeID").value), "", RsSavRec.Fields("StudeID").value)
    Me.DcbCurs.BoundText = IIf(IsNull(RsSavRec.Fields("CursID").value), "", RsSavRec.Fields("CursID").value)
    TxtContValue.Text = IIf(IsNull(RsSavRec.Fields("ContValue").value), 0, RsSavRec.Fields("ContValue").value)
    If RsSavRec.Fields("ContType").value = 1 Then
    ContType(1).value = True
    Me.TxtDiscountNet.Text = IIf(IsNull(RsSavRec.Fields("DiscountNet").value), 0, RsSavRec.Fields("DiscountNet").value)
    Me.txtPrice.Text = IIf(IsNull(RsSavRec.Fields("Price").value), 0, RsSavRec.Fields("Price").value)
    Me.TxtDiscount.Text = IIf(IsNull(RsSavRec.Fields("Discount").value), 0, RsSavRec.Fields("Discount").value)
    Me.DcbDiscount.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeDis").value), -1, RsSavRec.Fields("TypeDis").value)
    Else
    ContType(0).value = True
    Me.TxtDiscountNet1.Text = IIf(IsNull(RsSavRec.Fields("DiscountNet").value), 0, RsSavRec.Fields("DiscountNet").value)
    Me.TxtDiscount1.Text = IIf(IsNull(RsSavRec.Fields("Discount").value), 0, RsSavRec.Fields("Discount").value)
    Me.DcbDiscount1.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeDis").value), -1, RsSavRec.Fields("TypeDis").value)
    Me.TxtPrice1.Text = IIf(IsNull(RsSavRec.Fields("Price").value), 0, RsSavRec.Fields("Price").value)
    End If
    Cash_Insta.ListIndex = IIf(IsNull(RsSavRec.Fields("Cash_Insta").value), -1, RsSavRec.Fields("Cash_Insta").value)
     Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
     Me.DcbQuali.BoundText = IIf(IsNull(RsSavRec.Fields("QualiID").value), "", RsSavRec.Fields("QualiID").value)
     Me.DcbTypeContract.BoundText = IIf(IsNull(RsSavRec.Fields("TypeContID").value), "", RsSavRec.Fields("TypeContID").value)
     
    Me.TxtPhone.Text = IIf(IsNull(RsSavRec.Fields("Phone").value), "", RsSavRec.Fields("Phone").value)
    Me.TxtUQama.Text = IIf(IsNull(RsSavRec.Fields("UQama").value), "", RsSavRec.Fields("UQama").value)
    If RsSavRec.Fields("TypePayment").value = 4 Then
    Opt(4).value = True
    ElseIf RsSavRec.Fields("TypePayment").value = 3 Then
     Opt(3).value = True
     ElseIf RsSavRec.Fields("TypePayment").value = 2 Then
     Opt(2).value = True
    End If
    TxtNoStud.Text = IIf(IsNull(RsSavRec.Fields("NoStud").value), 0, RsSavRec.Fields("NoStud").value)
    TxtComm.Text = IIf(IsNull(RsSavRec.Fields("Comm").value), 0, RsSavRec.Fields("Comm").value)
    Me.TxtMidShare.Text = IIf(IsNull(RsSavRec.Fields("MidShare").value), 0, RsSavRec.Fields("MidShare").value)
    TxtMedalMan.Text = IIf(IsNull(RsSavRec.Fields("MedalMan").value), "", RsSavRec.Fields("MedalMan").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtNoInstal.Text = IIf(IsNull(RsSavRec.Fields("NoInstal").value), 0, RsSavRec.Fields("NoInstal").value)
    TypePeriod.ListIndex = IIf(IsNull(RsSavRec.Fields("TypePeriod").value), -1, RsSavRec.Fields("TypePeriod").value)
    txtPeriod.Text = IIf(IsNull(RsSavRec.Fields("Period").value), 0, RsSavRec.Fields("Period").value)
    FirstDateH.value = IIf(IsNull(RsSavRec.Fields("FirstDateH").value), ToHijriDate(Date), RsSavRec.Fields("FirstDateH").value)
    FirstDate.value = IIf(IsNull(RsSavRec.Fields("FirstDate").value), Date, RsSavRec.Fields("FirstDate").value)
    lbl(5).Caption = IIf(IsNull(RsSavRec.Fields("TotalPayment").value), 0, RsSavRec.Fields("TotalPayment").value)
    TxtStudValue.Text = IIf(IsNull(RsSavRec.Fields("StudValue").value), 0, RsSavRec.Fields("StudValue").value)
    Me.DcEmployee.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    If Not IsNull(RsSavRec.Fields("TypeSuper").value) Then
    If (RsSavRec.Fields("TypeSuper").value) = 1 Then
    Me.ChSuperv(1).value = True
    Else
    Me.ChSuperv(0).value = True
    End If
    Else
    Me.ChSuperv(0).value = True
    End If
    ''//////////
    If ContType(1).value = True Then
    FallGridStudent
    Else
      FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = 2
    End If
    
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
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
    If val(DcbBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
    Else
    MsgBox "Please Select Branch"
    End If
    Exit Sub
    End If
    If ContType(0).value = False And ContType(1).value = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄ «·⁄ﬁœ «Ê·«"
    Else
    MsgBox "Please Select Type of Contract"
    End If
    Exit Sub
    End If
    If ContType(0).value = True Then
    If val(DcbStudent.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„ œ—»"
    Else
    MsgBox "Please Select Student"
    End If
    DcbStudent.SetFocus
    Exit Sub
    End If
    If val(TxtPrice1.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ  «œŒ«· «·ﬁÌ„…"
    Else
    MsgBox "Please Enter Value"
    End If
    TxtPrice1.SetFocus
    Exit Sub
    End If
    If val(DcbCurs.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„«œ… "
    Else
    MsgBox "Please Select A course"
    End If
    DcbCurs.SetFocus
    Exit Sub
    End If
    If val(Cash_Insta.ListIndex) = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ "
    Else
    MsgBox "Please Select Payment"
    End If
    Cash_Insta.SetFocus
    Exit Sub
    End If
    End If
    ''///////////
    If ContType(1).value = True Then
    If val(DcbCompany.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·‘—ﬂ…"
    Else
    MsgBox "Please Select Company"
    End If
    DcbCompany.SetFocus
    Exit Sub
    End If
    If val(txtPrice.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ  «œŒ«· «·ﬁÌ„…"
    Else
    MsgBox "Please Enter Value"
    End If
    txtPrice.SetFocus
    Exit Sub
    End If
    If val(DcbTypeContract.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·⁄ﬁœ "
    Else
    MsgBox "Please Select Type Contract"
    End If
    DcbTypeContract.SetFocus
    Exit Sub
    End If
    If val(TxtNoStud.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· ⁄œœ «·„ œ—»Ì‰ "
    Else
    MsgBox "Please Eneter No. Student"
    End If
    TxtNoStud.SetFocus
    Exit Sub
    End If
    End If
    
               Dim currentcode As String

            If txtid.Text = "" Then
                currentcode = get_coding(Current_branch, "TblContrStudent", 10, Me.DCPreFix1.Text)

                If currentcode = "miniError" Then
                    MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Exit Sub
            
                ElseIf currentcode = "Manual" Then
                    MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Exit Sub
                Else
                    txtid = currentcode
                End If
                End If
         Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
              StrSQL = "Select * From TblContrStudent where  fullcode='" & Trim(DCPreFix1.Text & txtid.Text) & "'and id <>" & val(TxtSerial1.Text) & ""
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ ⁄ﬁœ  „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ " & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    Else
                     Msg = "This Contract Already Exist" & CHR(13)
                     
                    End If

                   
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                  '  XPTxtCusName.SetFocus
                    Exit Sub
                End If
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

Sub RelinGrid()
Dim Sm, summation As Double
Dim Counter As Integer
Dim i As Integer
Counter = 0
Sm = 0
summation = 0
lbl(5).Caption = 0
With Me.GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("InstalNo"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("InstalValue")))
End If
Next i
lbl(5).Caption = summation

End With
End Sub


Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.GridInstallments
Cancel = True
Select Case .ColKey(Col)
Case "InstalValue"
If Opt(2).value = True Then
Cancel = False
.ComboList = ""
Else
Cancel = True
End If
End Select
End With
End Sub

Private Sub ISButton2_Click()
 If Opt(4).value = False And Opt(3).value = False And Opt(2).value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
        Else
        MsgBox "Please Select Method Number of decimal"
        End If
        Exit Sub
        End If
 If ContType(0).value = True Then
If val(TxtPrice1.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·ﬁÌ„… "
Else
MsgBox "Please Enter  Value"
End If
TxtPrice1.SetFocus
Exit Sub
End If
Else
If val(txtPrice.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ﬁÌ„… «·⁄ﬁœ «·«Ã„«·Ì…"
Else
MsgBox "Please Enter  Value"
End If
txtPrice.SetFocus
Exit Sub
End If
End If
If val(Me.TxtNoInstal.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ⁄œœ «·œ›⁄«  "
Else
MsgBox "Please Enter No of  Payments"
End If
TxtNoInstal.SetFocus
Exit Sub
End If
If val(txtPeriod.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·› —… »Ì‰ «·œ›⁄«  "
Else
MsgBox "Please Enter No of  Period"
End If
txtPeriod.SetFocus
Exit Sub
End If
If val(TypePeriod.ListIndex) = -1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«—    ‰Ê⁄ «·› —… "
Else
MsgBox "Please Enter Type of  Period"
End If
TypePeriod.SetFocus
Exit Sub
End If
filgrid1
RelinGrid
End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
ShowAttachments TxtSerial1.Text, "0109201611"
ErrTrap:
End Sub
Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecorddateH.value = ToHijriDate(RecordDate.value)
End If
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecorddateH.value)
End If
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCode.Text, EmpID
        DcbCompany.BoundText = EmpID
    End If
End Sub

Private Sub TxtComm_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtMidShare.Text = Round((val(TxtComm.Text) * (val(txtPrice.Text) - val(TxtDiscountNet.Text))) / 100, 2)
End If
End Sub

Private Sub TxtComm_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtComm.Text, 0)
End Sub

Private Sub txtDiscount_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbDiscount.ListIndex) = 1 Then
TxtDiscountNet.Text = val(TxtStudValue.Text) * (val(TxtDiscount.Text) / 100)
Else
TxtDiscountNet.Text = val(TxtDiscount.Text)
End If
TxtMidShare.Text = Round((val(TxtComm.Text) * (val(TxtStudValue.Text) - val(TxtDiscountNet.Text))) / 100, 2)
TxtContValue.Text = val(TxtStudValue.Text) - val(TxtDiscountNet.Text)
txtPrice.Text = val(TxtContValue.Text) * val(TxtNoStud.Text)
End If
End Sub

Private Sub TxtDiscount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDiscount.Text, 0)
End Sub

Private Sub TxtDiscount1_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbDiscount1.ListIndex) = 1 Then
TxtDiscountNet1.Text = val(TxtValTrin.Text) * (val(TxtDiscount1.Text) / 100)
Else
TxtDiscountNet1.Text = val(TxtDiscount1.Text)
End If
TxtPrice1.Text = val(TxtValTrin.Text) - val(TxtDiscountNet1.Text)
End If
End Sub

Private Sub TxtMidShare_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMidShare.Text, 0)
End Sub
Private Sub TxtNoInstal_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoInstal.Text, 0)
End Sub

Private Sub TxtNoStud_Change()
If Me.TxtModFlg.Text <> "R" Then
txtDiscount_Change
txtPrice.Text = val(TxtNoStud.Text) * val(TxtContValue.Text)
End If
End Sub

Private Sub TxtNoStud_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoStud.Text, 0)
End Sub
Private Sub TxtPeriod_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.txtPeriod.Text, 0)
End Sub

Private Sub txtPrice_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtMidShare.Text = val(TxtComm.Text) * (val(txtPrice.Text) - val(TxtDiscountNet.Text))
End If
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.txtPrice.Text, 0)
End Sub
Private Sub TxtPrice1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPrice1.Text, 0)
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
Function CheckNomination() As Boolean
Dim Rs2 As ADODB.Recordset
Dim sql As String
Set Rs2 = New ADODB.Recordset
sql = "Select * from TblStuCandidacy where ContNoID=" & val(TxtSerial1.Text) & " "
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
CheckNomination = True
Else
CheckNomination = False
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
    Dim x As Integer
    Dim i As Integer
    Dim ID As Double
   If CheckNomination() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »«· —‘ÌÕ"
    Else
    MsgBox "Can not  delete this is process  linked to the process nomination"
    End If
    Exit Sub
    End If
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
         Cn.Execute "Delete from TblContrStudentDet where ContStudID=" & val(TxtSerial1.Text) & " "
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
                FG1.Clear flexClearScrollable, flexClearEverything
                FG1.Rows = 2
               GridInstallments.Clear flexClearScrollable, flexClearEverything
               GridInstallments.Rows = 2
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
    ElseIf TxtModFlg.Text = "R" Then
    relod1
   ' XPDtbTrans.Enabled = False
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
     If CheckNomination() = True Then
 '   If SystemOptions.UserInterface = ArabicInterface Then
 '   MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »«· —‘ÌÕ"
 ''   Else
  '  MsgBox "Can not  edit this is process  linked to the process nomination"
  '  End If
  '  Exit Sub
    End If
    GridInstallments.Rows = GridInstallments.Rows + 1
        TxtModFlg = "E"
        
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
    Cash_Insta_Change
    ContType(0).value = True
    relod
    Me.DcbBranch.BoundText = Current_branch
      FG1.Clear flexClearScrollable, flexClearEverything
     FG1.Rows = 2
      GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 2
    Me.DCboUserName.BoundText = user_id
 ChSuperv(0).value = True
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
   Label1(17).Caption = "SuperVisor"
  Label1(2).Caption = "Data of Contracts"
  Label1(22).Caption = "Value"
  Label1(21).Caption = "(Groups)"
  Label1(20).Caption = "Discount"
  Label1(19).Caption = "Saudi"
lbl(4).Caption = "No"
Label1(18).Caption = "Net"
ChSuperv(0).RightToLeft = False
ChSuperv(1).RightToLeft = False
ChSuperv(0).Caption = "Employee"
ChSuperv(1).Caption = "Not"
Label1(8).Caption = "Student Value"
lbl(11).Caption = "Branch"
ISButton3.Caption = "Attachments"
lbl(25).Caption = "Date"
ContType(0).Caption = "Student"
ContType(1).Caption = "Company"
Label1(0).Caption = "Company"
lbl(22).Caption = "Middleman"
Label1(3).Caption = "Commission"
Label1(1).Caption = "No.Student"
Label1(4).Caption = "Value"
Label1(5).Caption = "Discount"
Label1(6).Caption = "Mid.share"
Label1(12).Caption = "Student"
lbl(1).Caption = "ID No."
lbl(2).Caption = "Phone No."
Label1(11).Caption = "A course"
lbl(3).Caption = "Qualification"
Label1(9).Caption = "Net"
Label1(7).Caption = "Payment Method"
Label1(10).Caption = "No.Payments"
Label1(14).Caption = "First Payment"
Label1(16).Caption = "Remarks"
Label1(15).Caption = "Period"
lbl(37).Caption = "Method"
lbl(0).Caption = "Type Contract"
Opt(4).RightToLeft = False
Opt(4).Caption = "First Pay."
Opt(3).RightToLeft = False
Opt(3).Caption = "Last Pay."
Opt(2).RightToLeft = False
Opt(2).Caption = "Manual"
ISButton2.Caption = "Add"
lbl(14).Caption = "By"

    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"

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
      With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("InstalNo")) = "Payment No."
  .TextMatrix(0, .ColIndex("InstalDate")) = "Date"
  .TextMatrix(0, .ColIndex("InstalDateH")) = "Date"
  .TextMatrix(0, .ColIndex("InstalValue")) = "Value"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  End With
  
  With FG1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("FullCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Student Name"
  .TextMatrix(0, .ColIndex("Supervisor")) = "Supervisor"
  .TextMatrix(0, .ColIndex("Qualification")) = "Qualification"
  .TextMatrix(0, .ColIndex("Phone")) = "Phone"
  .TextMatrix(0, .ColIndex("UQama")) = "ID No."
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblContrStudent"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub TxtStudValue_Change()
If Me.TxtModFlg.Text <> "R" Then
txtDiscount_Change
txtPrice.Text = val(TxtNoStud.Text) * val(TxtContValue.Text)
End If
End Sub

Private Sub TxtSudCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
 Dim UQama As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, TxtSudCode.Text, 1, UQama
        DcbStudent.BoundText = EmpID
        TxtUQama.Text = UQama
    End If
End Sub

Private Sub TxtUQama_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
 Dim code As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, code, 2, TxtUQama.Text
        DcbStudent.BoundText = EmpID
        TxtSudCode.Text = code
    End If
End Sub

Private Sub TxtValTrin_Change()
DcbDiscount1_Change
End Sub
