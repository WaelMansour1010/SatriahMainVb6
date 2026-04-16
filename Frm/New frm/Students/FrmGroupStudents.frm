VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmGroupStudents 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13365
   Icon            =   "FrmGroupStudents.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13365
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
      TabIndex        =   33
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmGroupStudents.frx":6852
      Left            =   15480
      List            =   "FrmGroupStudents.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   34
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
      TabIndex        =   35
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
            Picture         =   "FrmGroupStudents.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGroupStudents.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   36
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
      ButtonImage     =   "FrmGroupStudents.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   38
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
      ButtonImage     =   "FrmGroupStudents.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   39
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
      ButtonImage     =   "FrmGroupStudents.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9825
      Left            =   0
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Width           =   13365
      _cx             =   23574
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
         Left            =   13410
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   11730
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   59
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   540
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   720
         Width           =   13425
         _cx             =   23680
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
            Left            =   11070
            MaxLength       =   50
            TabIndex        =   91
            Top             =   120
            Width           =   1125
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10575
            TabIndex        =   0
            Top             =   -240
            Visible         =   0   'False
            Width           =   1725
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   6120
            TabIndex        =   2
            Top             =   120
            Width           =   1320
            _extentx        =   2328
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   7545
            TabIndex        =   1
            Top             =   120
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193665
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCPreFix1 
            Height          =   315
            Left            =   9810
            TabIndex        =   92
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
            Left            =   4950
            TabIndex        =   68
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   25
            Left            =   8925
            TabIndex        =   67
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
            Height          =   255
            Index           =   4
            Left            =   12240
            TabIndex        =   42
            Top             =   120
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   8760
         Width           =   13335
         _cx             =   23521
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
            Left            =   11460
            TabIndex        =   44
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
            ButtonImage     =   "FrmGroupStudents.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   9735
            TabIndex        =   45
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
            ButtonImage     =   "FrmGroupStudents.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8310
            TabIndex        =   28
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
            ButtonImage     =   "FrmGroupStudents.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   6780
            TabIndex        =   46
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
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
            ButtonImage     =   "FrmGroupStudents.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5310
            TabIndex        =   47
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
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
            ButtonImage     =   "FrmGroupStudents.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   4080
            TabIndex        =   48
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
            ButtonImage     =   "FrmGroupStudents.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   2520
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
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
            ButtonImage     =   "FrmGroupStudents.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   480
            TabIndex        =   50
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   600
            Width           =   1515
            _ExtentX        =   2672
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
            ButtonImage     =   "FrmGroupStudents.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8370
            TabIndex        =   51
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
            TabIndex        =   66
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1635
            _ExtentX        =   2884
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
            ButtonImage     =   "FrmGroupStudents.frx":2A8D1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   315
            TabIndex        =   56
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   345
            Index           =   14
            Left            =   12240
            TabIndex        =   52
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   0
         Width           =   13410
         _cx             =   23654
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
            TabIndex        =   61
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
            ButtonImage     =   "FrmGroupStudents.frx":31133
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   675
            TabIndex        =   62
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
            ButtonImage     =   "FrmGroupStudents.frx":314CD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1350
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
            ButtonImage     =   "FrmGroupStudents.frx":31867
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1950
            TabIndex        =   64
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
            ButtonImage     =   "FrmGroupStudents.frx":31C01
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   12345
            Picture         =   "FrmGroupStudents.frx":31F9B
            Stretch         =   -1  'True
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  „Ã„Ê⁄«  «·„ œ—»Ì‰"
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
            TabIndex        =   65
            Top             =   240
            Width           =   4635
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   2100
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1320
         Width           =   13425
         _cx             =   23680
         _cy             =   3704
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
            Left            =   3960
            TabIndex        =   7
            Top             =   480
            Width           =   1290
         End
         Begin VB.TextBox TxtNameE 
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
            TabIndex        =   6
            Top             =   120
            Width           =   5130
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
            Left            =   3960
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.TextBox TxtNam 
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
            Left            =   6810
            TabIndex        =   5
            Top             =   120
            Width           =   5130
         End
         Begin MSDataListLib.DataCombo DcbInstrucor 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   480
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCurs 
            Height          =   315
            Left            =   6810
            TabIndex        =   9
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   840
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox Sa 
            Height          =   315
            Left            =   11085
            TabIndex        =   10
            Top             =   1320
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·”» "
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox Su 
            Height          =   315
            Left            =   10170
            TabIndex        =   11
            Top             =   1320
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·«Õœ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox Mo 
            Height          =   315
            Left            =   9330
            TabIndex        =   12
            Top             =   1320
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·«À‰Ì‰"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox Tu 
            Height          =   315
            Left            =   8250
            TabIndex        =   13
            Top             =   1320
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·À·«À«¡"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox We 
            Height          =   315
            Left            =   7170
            TabIndex        =   14
            Top             =   1320
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·«—»⁄«¡"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox Th 
            Height          =   315
            Left            =   5760
            TabIndex        =   15
            Top             =   1320
            Width           =   1065
            _Version        =   786432
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·Œ„Ì”"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox Fr 
            Height          =   315
            Left            =   4920
            TabIndex        =   16
            Top             =   1320
            Width           =   735
            _Version        =   786432
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "«·Ã„⁄…"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal StartDateH 
            Height          =   315
            Left            =   9210
            TabIndex        =   20
            Top             =   1680
            Width           =   1350
            _extentx        =   2381
            _extenty        =   556
         End
         Begin Dynamic_Byte.NourHijriCal EndDateH 
            Height          =   315
            Left            =   5040
            TabIndex        =   22
            Top             =   1680
            Width           =   1350
            _extentx        =   2355
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker EndDate 
            Height          =   315
            Left            =   6495
            TabIndex        =   21
            Top             =   1680
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193665
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker StartDate 
            Height          =   315
            Left            =   10590
            TabIndex        =   19
            Top             =   1680
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193665
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker FrmTime 
            Height          =   315
            Left            =   2460
            TabIndex        =   17
            Top             =   1320
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193666
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToTime 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193666
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
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
            ButtonImage     =   "FrmGroupStudents.frx":333A0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker DateTimp 
            Height          =   315
            Left            =   3120
            TabIndex        =   88
            Top             =   1680
            Visible         =   0   'False
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193665
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbDoplom 
            Height          =   315
            Left            =   6810
            TabIndex        =   104
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   480
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbHall 
            Height          =   315
            Left            =   120
            TabIndex        =   106
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   840
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁ«⁄…"
            Height          =   285
            Index           =   9
            Left            =   5430
            TabIndex        =   107
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·œÊ—…"
            Height          =   285
            Index           =   8
            Left            =   12120
            TabIndex        =   105
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï «·”«⁄…"
            Height          =   195
            Index           =   5
            Left            =   1320
            TabIndex        =   83
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰ «·”«⁄…"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   82
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ‰ ÂÌ „‰  «—ÌŒ"
            Height          =   285
            Index           =   5
            Left            =   7965
            TabIndex        =   81
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " »œ« „‰  «—ÌŒ"
            Height          =   285
            Index           =   3
            Left            =   12165
            TabIndex        =   80
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«Ì«„ «·„«œ…"
            Height          =   195
            Index           =   3
            Left            =   12090
            TabIndex        =   79
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„«œ…"
            Height          =   285
            Index           =   11
            Left            =   12120
            TabIndex        =   76
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ—»"
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   75
            Top             =   450
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ì"
            Height          =   285
            Index           =   2
            Left            =   5400
            TabIndex        =   74
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   285
            Index           =   0
            Left            =   12090
            TabIndex        =   73
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
            Height          =   285
            Index           =   1
            Left            =   12000
            TabIndex        =   72
            Top             =   165
            Visible         =   0   'False
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2700
         Left            =   0
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   6000
         Width           =   13425
         _cx             =   23680
         _cy             =   4763
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
            Left            =   10290
            TabIndex        =   24
            Top             =   45
            Width           =   2010
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
            Left            =   7530
            TabIndex        =   25
            Top             =   45
            Width           =   1410
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg 
            Height          =   1875
            Left            =   120
            TabIndex        =   71
            Top             =   465
            Width           =   13155
            _cx             =   23204
            _cy             =   3307
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmGroupStudents.frx":39C02
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
         Begin MSDataListLib.DataCombo DcbStudent 
            Height          =   315
            Left            =   3000
            TabIndex        =   26
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   45
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   45
            Width           =   2655
            _ExtentX        =   4683
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
            ButtonImage     =   "FrmGroupStudents.frx":39D0D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   0
            Left            =   12450
            TabIndex        =   84
            Top             =   2400
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
            ButtonImage     =   "FrmGroupStudents.frx":4056F
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   10650
            TabIndex        =   85
            Top             =   2400
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
            ButtonImage     =   "FrmGroupStudents.frx":40B09
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ œ—»"
            Height          =   240
            Index           =   1
            Left            =   9000
            TabIndex        =   78
            Top             =   105
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÂÊÌ…"
            Height          =   240
            Index           =   12
            Left            =   12240
            TabIndex        =   77
            Top             =   105
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2580
         Left            =   3600
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   3360
         Width           =   9705
         _cx             =   17119
         _cy             =   4551
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2115
            Left            =   120
            TabIndex        =   87
            Top             =   105
            Width           =   9585
            _cx             =   16907
            _cy             =   3731
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmGroupStudents.frx":410A3
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   2
            Left            =   7920
            TabIndex        =   89
            Top             =   2280
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
            ButtonImage     =   "FrmGroupStudents.frx":4125D
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   3
            Left            =   6120
            TabIndex        =   90
            Top             =   2280
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
            ButtonImage     =   "FrmGroupStudents.frx":417F7
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   1260
         Left            =   0
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   3360
         Width           =   3495
         _cx             =   6165
         _cy             =   2223
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
            Height          =   795
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   102
            Top             =   360
            Width           =   3165
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   7
            Left            =   1200
            TabIndex        =   103
            Top             =   120
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   1260
         Left            =   0
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   4680
         Width           =   3495
         _cx             =   6165
         _cy             =   2223
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
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
         Begin Dynamic_Byte.NourHijriCal NEndDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   1350
            _extentx        =   2355
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker NEndDate 
            Height          =   315
            Left            =   1515
            TabIndex        =   96
            Top             =   480
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62193665
            CurrentDate     =   38784
         End
         Begin XtremeSuiteControls.CheckBox ChFlgEnd 
            Height          =   315
            Left            =   2280
            TabIndex        =   98
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   " „ «·«‰ Â«¡"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUser2 
            Height          =   315
            Left            =   120
            TabIndex        =   99
            Top             =   840
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «‰ Â«¡ «·„Ã„Ê⁄…"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   6
            Left            =   -120
            TabIndex        =   101
            Top             =   120
            Width           =   2805
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "»Ê«”ÿ…  "
            Height          =   285
            Index           =   7
            Left            =   2715
            TabIndex        =   100
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "» «—ÌŒ"
            Height          =   285
            Index           =   6
            Left            =   2595
            TabIndex        =   97
            Top             =   480
            Width           =   1095
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
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmGroupStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim II As Long

Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblStuGroup", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
Case 1
Fg.Clear flexClearScrollable, flexClearEverything
      Fg.Rows = 1
Case 2
RemoveGridRow2
Case 3
VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
      VSFlexGrid1.Rows = 1
End Select
End If
End Sub
Private Sub RemoveGridRow2()

    With Me.VSFlexGrid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub

Private Sub RemoveGridRow()

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub

Private Sub DcbDoplom_Change()
DcbDoplom_Click (0)
End Sub

Private Sub DcbDoplom_Click(Area As Integer)
     Dim Dcombos As New ClsDataCombos
     Set Dcombos = New ClsDataCombos
     If val(DcbDoplom.BoundText) <> 0 Then
     Dcombos.GetStudentCurs Me.DcbCurs, val(DcbDoplom.BoundText)
     End If
End Sub

Private Sub DcbInstrucor_Change()
DcbInstrucor_Click (0)
End Sub

Private Sub DcbInstrucor_Click(Area As Integer)
Dim UQama As String
  If val(DcbInstrucor.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetInstructorCode val(DcbInstrucor.BoundText), EmpCode, 0
    Me.Text1.Text = EmpCode
End Sub

Private Sub DcbInstrucor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 201
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub DcbStudent_Change()
DcbStudent_Click (0)
End Sub
Private Sub DcbStudent_Click(Area As Integer)
Dim UQama As String
  If val(DcbStudent.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode val(DcbStudent.BoundText), EmpCode, 0, UQama
    TxtUQama.Text = UQama
    Me.TxtSudCode.Text = EmpCode
End Sub


Private Sub ENDDATE_Change()
If Me.TxtModFlg.Text <> "R" Then
         EndDateH.value = ToHijriDate(EndDate.value)
End If
End Sub

Private Sub ENDDATEH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 EndDate.value = ToGregorianDate(EndDateH.value)
End If
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblStuGroup   "
      conection = conection & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    conection = conection & " Order By ID"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
     Dim Dcombos As New ClsDataCombos
   Dcombos.GetCodeing Me.DCPreFix1, 14
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   'Dcombos.GetStudentCurs Me.DcbCurs
   Dcombos.GetStudentClassRooms Me.DcbHall
   Dcombos.GetStudent Me.DcbStudent
   Dcombos.GeInstructor Me.DcbInstrucor
   Dcombos.GetStudentDeploma Me.DcbDoplom
   
   Dcombos.GetUsers Me.DcbUser2
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
    Dim Sql As String
    Dim ID As Double
    Dim I As Integer
    Dim k As Integer
    
      If Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete from TblStuGroupDet where StudGrouID=" & val(TxtSerial1.Text) & " "
        Cn.Execute "delete from TblStuFingerprint where GroupID=" & val(TxtSerial1.Text) & " "
 End If
         Dim Msg As String

                    RsSavRec.Fields("code").value = txtid.Text
                RsSavRec.Fields("prifix").value = IIf(DCPreFix1.Text = "", Null, DCPreFix1.Text)
               RsSavRec.Fields("Fullcode").value = IIf(DCPreFix1.BoundText = "", Null, DCPreFix1.Text) & IIf(Trim(txtid.Text) = "", Null, txtid.Text)
   RsSavRec.Fields("RecordDateH").value = RecordDateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("Remarks").value = TxtRemarks.Text
   RsSavRec.Fields("Name").value = TxtNam.Text
   RsSavRec.Fields("UQama").value = TxtUQama.Text
   RsSavRec.Fields("NameE").value = TxtNameE.Text
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("InstrcID").value = val(Me.DcbInstrucor.BoundText)
   RsSavRec.Fields("CursID").value = val(Me.DcbCurs.BoundText)
   RsSavRec.Fields("HallID").value = val(Me.DcbHall.BoundText)
   RsSavRec.Fields("DoplomID").value = val(Me.DcbDoplom.BoundText)
   RsSavRec.Fields("StudID").value = val(Me.DcbStudent.BoundText)
   RsSavRec.Fields("StartDate").value = StartDate.value
   RsSavRec.Fields("StartDateH").value = StartDateH.value
   RsSavRec.Fields("EndDateH").value = EndDateH.value
   RsSavRec.Fields("EndDate").value = EndDate.value
   RsSavRec.Fields("HallID").value = val(DcbHall.BoundText)
   If Sa.value = vbChecked Then
   RsSavRec.Fields("Sa").value = 1
   Else
   RsSavRec.Fields("Sa").value = 0
   End If
   If Su.value = vbChecked Then
   RsSavRec.Fields("Su").value = 1
   Else
   RsSavRec.Fields("Su").value = 0
   End If
   If Mo.value = vbChecked Then
   RsSavRec.Fields("Mo").value = 1
   Else
   RsSavRec.Fields("Mo").value = 0
   End If
   If Tu.value = vbChecked Then
   RsSavRec.Fields("Tu").value = 1
   Else
   RsSavRec.Fields("Tu").value = 0
   End If
    If We.value = vbChecked Then
   RsSavRec.Fields("We").value = 1
   Else
   RsSavRec.Fields("We").value = 0
   End If
    If Th.value = vbChecked Then
   RsSavRec.Fields("Th").value = 1
   Else
   RsSavRec.Fields("Th").value = 0
   End If
   If Fr.value = vbChecked Then
   RsSavRec.Fields("Fr").value = 1
   Else
   RsSavRec.Fields("Fr").value = 0
   End If
   RsSavRec.Fields("FrmTime").value = FormatDateTime(Me.FrmTime.value, vbShortTime)
   RsSavRec.Fields("ToTime").value = FormatDateTime(Me.ToTime.value, vbShortTime)
   RsSavRec.Update
  ''//////////////////////////

  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblStuGroupDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With Me.Fg
       For I = .FixedRows To .Rows - 1
       If val(.TextMatrix(I, .ColIndex("StudID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("StudGrouID").value = val(Me.TxtSerial1.Text)
                RsDevsub("StudID").value = IIf((.TextMatrix(I, .ColIndex("StudID"))) = "", Null, .TextMatrix(I, .ColIndex("StudID")))
                RsDevsub("CompID").value = val(.TextMatrix(I, .ColIndex("CompID")))
                RsDevsub("TypeTr").value = 0
       RsDevsub.Update
      End If
     Next I
    End With
'''///////////////
 
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblStuGroupDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.VSFlexGrid1
       For I = .FixedRows To .Rows - 1
       If (.TextMatrix(I, .ColIndex("GDate"))) <> "" Then
       RsDevsub.AddNew
                RsDevsub("StudGrouID").value = val(Me.TxtSerial1.Text)
                RsDevsub("FrmTime").value = IIf((.TextMatrix(I, .ColIndex("FrmTime"))) = "", Null, FormatDateTime(.TextMatrix(I, .ColIndex("FrmTime")), vbShortTime))
                RsDevsub("ToTime").value = IIf((.TextMatrix(I, .ColIndex("ToTime"))) = "", Null, FormatDateTime(.TextMatrix(I, .ColIndex("ToTime")), vbShortTime))
                RsDevsub("GDate").value = IIf((IsDate(.TextMatrix(I, .ColIndex("GDate")))), .TextMatrix(I, .ColIndex("GDate")), Null)
                RsDevsub("GDateH").value = IIf(((.TextMatrix(I, .ColIndex("GDateH")))) <> "", ToHijriDate((.TextMatrix(I, .ColIndex("GDateH")))), Null)
                RsDevsub("HallID").value = val(.TextMatrix(I, .ColIndex("HallID")))
                RsDevsub("InstructID").value = val(.TextMatrix(I, .ColIndex("InstructID")))
                RsDevsub("CursID").value = val(.TextMatrix(I, .ColIndex("CursID")))
                RsDevsub("TypeTr").value = 1
       RsDevsub.Update
      End If
     Next I
    End With
    '''/////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblStuFingerprint Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.VSFlexGrid1
    For k = 1 To Fg.Rows - 1
       For I = .FixedRows To .Rows - 1
       If val(Fg.TextMatrix(k, Fg.ColIndex("StudID"))) <> 0 Then
       If .TextMatrix(I, .ColIndex("GDate")) <> "" Then
       RsDevsub.AddNew
                RsDevsub("GroupID").value = val(Me.TxtSerial1.Text)
                RsDevsub("StudID").value = val(Fg.TextMatrix(k, Fg.ColIndex("StudID")))
                RsDevsub("CompID").value = val(Fg.TextMatrix(k, Fg.ColIndex("CompID")))
                RsDevsub("HallID").value = val(.TextMatrix(I, .ColIndex("HallID")))
                RsDevsub("InstructID").value = val(.TextMatrix(I, .ColIndex("InstructID")))
                RsDevsub("DoplomID").value = val(Me.DcbDoplom.BoundText)
                RsDevsub("CursID").value = val(.TextMatrix(I, .ColIndex("CursID")))
                RsDevsub("brnchid").value = val(Me.DcbBranch.BoundText)
                RsDevsub("FrmTime").value = IIf((.TextMatrix(I, .ColIndex("FrmTime"))) = "", Null, FormatDateTime(.TextMatrix(I, .ColIndex("FrmTime")), vbShortTime))
                RsDevsub("ToTime").value = IIf((.TextMatrix(I, .ColIndex("ToTime"))) = "", Null, FormatDateTime(.TextMatrix(I, .ColIndex("ToTime")), vbShortTime))
                RsDevsub("GDate").value = IIf((IsDate(.TextMatrix(I, .ColIndex("GDate")))), .TextMatrix(I, .ColIndex("GDate")), Null)
                RsDevsub("GDateH").value = IIf(((.TextMatrix(I, .ColIndex("GDateH")))) <> "", ToHijriDate((.TextMatrix(I, .ColIndex("GDateH")))), Null)
       RsDevsub.Update
       End If
      End If
     Next I
     Next k
    End With
    
   
    
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
   Function ISAtted() As Boolean
   Dim Sql As String
   Dim Rs3 As ADODB.Recordset
   Set Rs3 = New ADODB.Recordset
   Sql = "Select * from TblAttendance where GroupID=" & val(TxtSerial1.Text) & ""
   Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs3.RecordCount > 0 Then
   ISAtted = True
   Else
   ISAtted = False
   End If
   End Function
Sub FullGri()
Dim Rs3 As ADODB.Recordset
Dim I As Integer
Dim Sql As String
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = 1
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
 Set Rs3 = New ADODB.Recordset
Sql = " SELECT     dbo.TblStuGroupDet.StudID, dbo.TblStuGroupDet.StudGrouID, dbo.TblStudent.Name, dbo.TblStudent.NameE, dbo.TblStudent.FullCode, dbo.TblStudent.UQama, "
Sql = Sql & "                       dbo.TblStudent.CompID , dbo.TblStuGroupDet.TypeTr ,dbo.TblStuGroupDet.Description"
Sql = Sql & " FROM         dbo.TblStuGroupDet LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblStudent ON dbo.TblStuGroupDet.StudID = dbo.TblStudent.ID"
Sql = Sql & " WHERE     (dbo.TblStuGroupDet.TypeTr = 0) AND (dbo.TblStuGroupDet.StudGrouID = " & val(TxtSerial1.Text) & ")"
Sql = Sql & " and     (dbo.TblStuGroupDet.FlgGrpuoUpdae = 0 or dbo.TblStuGroupDet.FlgGrpuoUpdae = 1 or dbo.TblStuGroupDet.FlgGrpuoUpdae is null)"
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Fg
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("FullCode")) = IIf(IsNull(Rs3("FullCode").value), "", Rs3("FullCode").value)
.TextMatrix(I, .ColIndex("UQama")) = IIf(IsNull(Rs3("UQama").value), "", Rs3("UQama").value)
.TextMatrix(I, .ColIndex("CompID")) = IIf(IsNull(Rs3("CompID").value), 0, Rs3("CompID").value)
.TextMatrix(I, .ColIndex("StudID")) = IIf(IsNull(Rs3("StudID").value), 0, Rs3("StudID").value)
.TextMatrix(I, .ColIndex("Description")) = IIf(IsNull(Rs3("Description").value), "", Rs3("Description").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
End If
Rs3.MoveNext
Next I
End With
End If
''/////////////
Dim Shifttime As Date
 Set Rs3 = New ADODB.Recordset
Sql = "SELECT     dbo.TblInstructors.Name AS InstruName, dbo.TblInstructors.NameE AS InstruNameE, dbo.TblStudentCurs.Name AS CursName, "
Sql = Sql & "                       dbo.TblStudentCurs.NameE AS CursNameE, dbo.TblStudentClassRooms.Name AS RoomName, dbo.TblStudentClassRooms.NameE AS RoomNameE,"
Sql = Sql & "                      dbo.TblStuGroupDet.*"
Sql = Sql & " FROM         dbo.TblStuGroupDet LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblStudentClassRooms ON dbo.TblStuGroupDet.HallID = dbo.TblStudentClassRooms.ID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblStudentCurs ON dbo.TblStuGroupDet.CursID = dbo.TblStudentCurs.ID LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblInstructors ON dbo.TblStuGroupDet.InstructID = dbo.TblInstructors.ID"
Sql = Sql & " Where (dbo.TblStuGroupDet.TypeTr = 1) And (dbo.TblStuGroupDet.StudGrouID = " & val(TxtSerial1.Text) & ")"

Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With VSFlexGrid1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
    If Not IsNull(Rs3("FrmTime").value) Then
        Shifttime = FormatDateTime(Rs3("FrmTime").value, vbShortTime)
        .TextMatrix(I, .ColIndex("FrmTime")) = Shifttime
    End If
    If Not IsNull(Rs3("ToTime").value) Then
        Shifttime = FormatDateTime(Rs3("ToTime").value, vbShortTime)
        .TextMatrix(I, .ColIndex("ToTime")) = Shifttime
    End If
    .TextMatrix(I, .ColIndex("GDate")) = IIf(IsNull(Rs3("GDate").value), "", Rs3("GDate").value)
    .TextMatrix(I, .ColIndex("GDateH")) = IIf(IsNull(Rs3("GDateH").value), "", Rs3("GDateH").value)
    .TextMatrix(I, .ColIndex("HallID")) = IIf(IsNull(Rs3("HallID").value), "", Rs3("HallID").value)
    .TextMatrix(I, .ColIndex("InstructID")) = IIf(IsNull(Rs3("InstructID").value), "", Rs3("InstructID").value)
    .TextMatrix(I, .ColIndex("CursID")) = IIf(IsNull(Rs3("CursID").value), "", Rs3("CursID").value)
    If SystemOptions.UserInterface = ArabicInterface Then
    .TextMatrix(I, .ColIndex("InstruName")) = IIf(IsNull(Rs3("InstruName").value), "", Rs3("InstruName").value)
    .TextMatrix(I, .ColIndex("CursName")) = IIf(IsNull(Rs3("CursName").value), "", Rs3("CursName").value)
    .TextMatrix(I, .ColIndex("RoomName")) = IIf(IsNull(Rs3("RoomName").value), "", Rs3("RoomName").value)
    Else

    
     .TextMatrix(I, .ColIndex("InstruName")) = IIf(IsNull(Rs3("InstruNameE").value), "", Rs3("InstruNameE").value)
    .TextMatrix(I, .ColIndex("CursName")) = IIf(IsNull(Rs3("CursNameE").value), "", Rs3("CursNameE").value)
    .TextMatrix(I, .ColIndex("RoomName")) = IIf(IsNull(Rs3("RoomNameE").value), "", Rs3("RoomNameE").value)
    End If
    
Rs3.MoveNext
Next I
End With
End If
End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim I As Integer
    Dim Shifttime As Date
     DCPreFix1.Text = IIf(IsNull(RsSavRec("prifix").value), "", RsSavRec("prifix").value)
     Me.txtid.Text = IIf(IsNull(RsSavRec("code").value), "", RsSavRec("code").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtNam.Text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    TxtNameE.Text = IIf(IsNull(RsSavRec.Fields("NameE").value), "", RsSavRec.Fields("NameE").value)
    Me.DcbInstrucor.BoundText = IIf(IsNull(RsSavRec.Fields("InstrcID").value), "", RsSavRec.Fields("InstrcID").value)
    Me.DcbStudent.BoundText = IIf(IsNull(RsSavRec.Fields("StudID").value), "", RsSavRec.Fields("StudID").value)
    TxtUQama.Text = IIf(IsNull(RsSavRec.Fields("UQama").value), "", RsSavRec.Fields("UQama").value)
    StartDateH.value = IIf(IsNull(RsSavRec.Fields("StartDateH").value), ToHijriDate(Date), RsSavRec.Fields("StartDateH").value)
    StartDate.value = IIf(IsNull(RsSavRec.Fields("StartDate").value), Date, RsSavRec.Fields("StartDate").value)
    EndDateH.value = IIf(IsNull(RsSavRec.Fields("EndDateH").value), ToHijriDate(Date), RsSavRec.Fields("EndDateH").value)
    EndDate.value = IIf(IsNull(RsSavRec.Fields("EndDate").value), Date, RsSavRec.Fields("EndDate").value)
    NEndDate.value = IIf(IsNull(RsSavRec.Fields("NEndDate").value), Date, RsSavRec.Fields("NEndDate").value)
    NEndDateH.value = IIf(IsNull(RsSavRec.Fields("NEndDateH").value), ToHijriDate(Date), RsSavRec.Fields("NEndDateH").value)
    DcbUser2.BoundText = IIf(IsNull(RsSavRec.Fields("UsrtID2").value), "", RsSavRec.Fields("UsrtID2").value)
    Me.DcbDoplom.BoundText = IIf(IsNull(RsSavRec.Fields("DoplomID").value), "", RsSavRec.Fields("DoplomID").value)
     Me.DcbCurs.BoundText = IIf(IsNull(RsSavRec.Fields("CursID").value), "", RsSavRec.Fields("CursID").value)
     Me.DcbHall.BoundText = IIf(IsNull(RsSavRec("HallID").value), "", RsSavRec("HallID").value)
                '''''''''''''''''
    If Not IsNull(RsSavRec.Fields("FlgEnd").value) Then
    If RsSavRec.Fields("FlgEnd").value = 1 Then
    ChFlgEnd.value = vbChecked
    Else
    ChFlgEnd.value = vbUnchecked
    End If
    Else
    ChFlgEnd.value = vbUnchecked
    End If
    
    If Not IsNull(RsSavRec.Fields("Sa").value) Then
    If RsSavRec.Fields("Sa").value = True Then
    Sa.value = vbChecked
    Else
    Sa.value = vbUnchecked
    End If
    Else
    Sa.value = vbUnchecked
    End If
    ''''
    If Not IsNull(RsSavRec.Fields("Su").value) Then
    If RsSavRec.Fields("Su").value = True Then
    Su.value = vbChecked
    Else
    Su.value = vbUnchecked
    End If
    Else
    Su.value = vbUnchecked
    End If
   ''//
    If Not IsNull(RsSavRec.Fields("Mo").value) Then
    If RsSavRec.Fields("Mo").value = True Then
    Mo.value = vbChecked
    Else
    Mo.value = vbUnchecked
    End If
    Else
    Mo.value = vbUnchecked
    End If
  '///
    If Not IsNull(RsSavRec.Fields("Tu").value) Then
    If RsSavRec.Fields("Tu").value = True Then
    Tu.value = vbChecked
    Else
    Tu.value = vbUnchecked
    End If
    Else
    Tu.value = vbUnchecked
    End If
 ''//
    If Not IsNull(RsSavRec.Fields("We").value) Then
    If RsSavRec.Fields("We").value = True Then
    We.value = vbChecked
    Else
    We.value = vbUnchecked
    End If
    Else
    We.value = vbUnchecked
    End If
  ''//
    If Not IsNull(RsSavRec.Fields("Th").value) Then
    If RsSavRec.Fields("Th").value = True Then
    Th.value = vbChecked
    Else
    Th.value = vbUnchecked
    End If
    Else
    Th.value = vbUnchecked
    End If
      ''//
    If Not IsNull(RsSavRec.Fields("Fr").value) Then
    If RsSavRec.Fields("Fr").value = True Then
    Fr.value = vbChecked
    Else
    Fr.value = vbUnchecked
    End If
    Else
    Fr.value = vbUnchecked
    End If
    
       If Not IsNull(RsSavRec("FrmTime").value) Then
        Shifttime = FormatDateTime(RsSavRec("FrmTime").value, vbShortTime)
        Me.FrmTime.value = Shifttime
    End If
       If Not IsNull(RsSavRec("ToTime").value) Then
        Shifttime = FormatDateTime(RsSavRec("ToTime").value, vbShortTime)
        Me.ToTime.value = Shifttime
    End If
    'FallGridStudent
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
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
    If val(DcbBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
    Else
    MsgBox "Please Select Branch"
    End If
    DcbBranch.SetFocus
    Exit Sub
    End If
    If TxtNam.Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· «·«”„ "
    Else
    MsgBox "Please Enter Name"
    End If
    TxtNam.SetFocus
    Exit Sub
    End If
    
    If val(DcbInstrucor.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„œ—»"
    Else
    MsgBox "Please Select Instructor"
    End If
    DcbInstrucor.SetFocus
    Exit Sub
    End If
    If val(DcbCurs.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·ﬂÊ—”"
    Else
    MsgBox "Please Select A course"
    End If
    DcbCurs.SetFocus
    Exit Sub
    End If
    Dim currentcode As String

            If txtid.Text = "" Then
                currentcode = get_coding(Current_branch, "TblStuGroup", 14, Me.DCPreFix1.Text)

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
              StrSQL = "Select * From TblStuGroup where  Fullcode ='" & Trim(DCPreFix1.Text & txtid.Text) & "'and id<>" & val(Me.TxtSerial1.Text) & ""
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ „Ã„Ê⁄…  „”Ã·… „”»ﬁ« »Â–« «·ﬂÊœ " & Chr(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
                                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    Else
                     Msg = "This Group Already Exist" & Chr(13)
                     
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


Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
If val(DcbStudent.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·ÿ«·»"
Else
MsgBox "Please Select Student "
End If
Exit Sub
End If
If CheckRepeat() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ﬂ—«—"
Else
MsgBox "Can not Repetition"
End If
Exit Sub
End If
FullGridStudent
End If
End Sub
Function CheckRepeat() As Boolean
Dim I As Integer
With Fg
For I = 1 To .Rows - 1
If val(DcbStudent.BoundText) = val(.TextMatrix(I, .ColIndex("StudID"))) Then
CheckRepeat = True
Exit Function
End If
Next I
End With
CheckRepeat = False
End Function
Sub FullGridStudent()
Dim I As Integer
Dim k As Integer
With Fg
k = .Rows
.Rows = .Rows + 1
For I = k To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("StudID")) = val(DcbStudent.BoundText)
.TextMatrix(I, .ColIndex("FullCode")) = TxtSudCode.Text
.TextMatrix(I, .ColIndex("Name")) = DcbStudent.Text
.TextMatrix(I, .ColIndex("UQama")) = TxtUQama.Text
.TextMatrix(I, .ColIndex("CompID")) = GetComID(val(DcbStudent.BoundText))
Next I
End With
DcbStudent.BoundText = 0
TxtUQama.Text = ""
TxtSudCode.Text = ""
End Sub
Private Sub ISButton3_Click()
            On Error Resume Next
ShowAttachments TxtSerial1.Text, "060920160011"
ErrTrap:
End Sub

Private Sub ISButton4_Click()
If Me.TxtModFlg.Text <> "R" Then
If Sa.value = vbUnchecked And Su.value = vbUnchecked And Mo.value = vbUnchecked And Tu.value = vbUnchecked And We.value = vbUnchecked And Th.value = vbUnchecked And Fr.value = vbUnchecked Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·«Ì«„"
Else
MsgBox "Please Select Days"
End If
Exit Sub
End If
If val(DcbInstrucor.BoundText) = 0 Or DcbInstrucor.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„œ—»"
Else
MsgBox "Please Select Instructor"
End If
DcbInstrucor.SetFocus
Exit Sub
End If

If val(DcbCurs.BoundText) = 0 Or DcbCurs.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„«œ…"
Else
MsgBox "Please Select Subject"
End If
DcbCurs.SetFocus
Exit Sub
End If

If val(DcbHall.BoundText) = 0 Or DcbHall.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·ﬁ«⁄…"
Else
MsgBox "Please Select Hall"
End If
DcbHall.SetFocus
Exit Sub
End If
FillGridDate
End If
End Sub
Function GetComID(Optional ID As Double) As Double
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim Sql As String
Sql = "Select CompID from  TblStudent where ID=" & ID & " "
Rs5.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetComID = IIf(IsNull(Rs5("CompID").value), 0, Rs5("CompID").value)
Else
GetComID = 0
End If
End Function
Sub FillGridDate()
Dim I As Integer
Dim NoDay As Integer
Dim k As Integer
Dim DifDate As Integer
Dim cout As Integer
DateTimp.value = StartDate.value
DifDate = DateDiff("D", Me.StartDate.value, Me.EndDate.value)

With VSFlexGrid1
cout = .Rows
k = .Rows
.Rows = .Rows + DifDate
For I = k To .Rows - 1
If I <> k Then
DateTimp.value = DateAdd("d", 1, DateTimp.value)
End If
NoDay = Weekday(DateTimp.value)
If (Sa.value = vbChecked And NoDay = 7) Or (Su.value = vbChecked And NoDay = 1) Or (Mo.value = vbChecked And NoDay = 2) Or (Tu.value = vbChecked And NoDay = 3) _
Or (We.value = vbChecked And NoDay = 4) Or (Th.value = vbChecked And NoDay = 5) Or (Fr.value = vbChecked And NoDay = 6) Then
.TextMatrix(cout, .ColIndex("Serial")) = cout
.TextMatrix(cout, .ColIndex("GDate")) = DateTimp.value
.TextMatrix(cout, .ColIndex("GDateH")) = ToHijriDate(DateTimp.value)
.TextMatrix(cout, .ColIndex("FrmTime")) = FormatDateTime(Me.FrmTime.value, vbShortTime)
.TextMatrix(cout, .ColIndex("ToTime")) = FormatDateTime(Me.ToTime.value, vbShortTime)
.TextMatrix(cout, .ColIndex("InstruName")) = DcbInstrucor.Text
.TextMatrix(cout, .ColIndex("CursName")) = DcbCurs.Text
.TextMatrix(cout, .ColIndex("RoomName")) = DcbHall.Text
.TextMatrix(cout, .ColIndex("HallID")) = val(DcbHall.BoundText)
.TextMatrix(cout, .ColIndex("InstructID")) = val(DcbInstrucor.BoundText)
.TextMatrix(cout, .ColIndex("CursID")) = val(DcbCurs.BoundText)
cout = cout + 1
Else
.Rows = .Rows - 1
End If
Next I
.Rows = cout
End With
End Sub

Private Sub ISButton8_Click()
FrmSearStudent.inde = 7
Load FrmSearStudent
FrmSearStudent.show vbModal
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(RecordDate.value)
End If
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecordDateH.value)
End If
End Sub



Private Sub STARTDATE_Change()
If Me.TxtModFlg.Text <> "R" Then
         StartDateH.value = ToHijriDate(StartDate.value)
End If
End Sub

Private Sub StartDateH_GotFocus()
If Me.TxtModFlg.Text <> "R" Then
 StartDate.value = ToGregorianDate(StartDateH.value)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetInstructorCode EmpID, Text1.Text, 1
        DcbInstrucor.BoundText = EmpID
    End If
End Sub

Private Sub TxtName_GotFocus(Index As Integer)
SwitchKeyboardLang LANG_ARABIC
End Sub



Private Sub TxtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
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
    Dim Sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim I As Integer
    Dim ID As Double
    If ISAtted() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–›. Â–Â «·„Ã„Ê⁄… „— »ÿ… »«·Õ÷Ê— "
    Else
    MsgBox "Can not be delete all linked to attend "
    End If
    Exit Sub
    End If
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
                Cn.Execute "delete from TblStuGroupDet where StudGrouID=" & val(TxtSerial1.Text) & " "
               Cn.Execute "delete from TblStuFingerprint where GroupID=" & val(TxtSerial1.Text) & " "
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            uncheck
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 1
               VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
               VSFlexGrid1.Rows = 1
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
    If ISAtted() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·„Ã„Ê⁄… „— »ÿ… »«·Õ÷Ê— "
    Else
    MsgBox "Can not be edited all linked to attend "
    End If
    Exit Sub
    End If
    ' VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
    Fg.Rows = Fg.Rows + 1
        TxtModFlg = "E"
        
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
  
    Me.DcbBranch.BoundText = Current_branch
      Fg.Clear flexClearScrollable, flexClearEverything
      Fg.Rows = 1
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
    Me.DCboUserName.BoundText = user_id
 uncheck
ErrTrap:
End Sub
Sub uncheck()
Me.Sa.value = vbUnchecked
Me.Su.value = vbUnchecked
Me.Mo.value = vbUnchecked
Me.Tu.value = vbUnchecked
Me.We.value = vbUnchecked
Me.Th.value = vbUnchecked
Me.Fr.value = vbUnchecked
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
   ' form name
   lbl(6).Caption = "Date"
   lbl(7).Caption = "By"
   ChFlgEnd.Caption = "Terminate"
   ChFlgEnd.RightToLeft = False
   Label1(6).Caption = "Data of Expiration Group"
   Label1(7).Caption = "Remarks"
   Label1(11).Caption = "Subject"
   Label1(1).Caption = "Student"
  Label1(2).Caption = "Groups Data"
lbl(4).Caption = "No"
lbl(11).Caption = "Branch"
ISButton3.Caption = "Attachments"
lbl(25).Caption = "Date"
lbl(1).Caption = "Code"
lbl(0).Caption = "Name Arabic"
lbl(2).Caption = "Name English"
Label1(0).Caption = "Instractor"
Label1(3).Caption = "Days"
Sa.RightToLeft = False
Sa.Caption = "Saturday"
Su.RightToLeft = False
Su.Caption = "Sunday"
Mo.RightToLeft = False
Mo.Caption = "Monday"
Tu.RightToLeft = False
Tu.Caption = "Tuesday"
We.RightToLeft = False
We.Caption = "Wednesday"
Th.RightToLeft = False
Th.Caption = "Thursday"
Fr.RightToLeft = False
Fr.Caption = "Friday"
Label1(4).Caption = "From"
Label1(5).Caption = "To"
lbl(3).Caption = "Start"
lbl(5).Caption = "End"
ISButton4.Caption = "Add"
ISButton2.Caption = "Add"
lbl(14).Caption = "By"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete All"
Cmd(2).Caption = "Delete"
Cmd(3).Caption = "Delete All"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"
Label1(12).Caption = "ID No."
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
   Label1(8).Caption = "Diploma"
   Label1(9).Caption = "Hall"
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
      
      .TextMatrix(0, .ColIndex("InstruName")) = "Instructor"
      .TextMatrix(0, .ColIndex("RoomName")) = "Hall"
      .TextMatrix(0, .ColIndex("CursName")) = "Subject"
    .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("GDate")) = "Date"
  .TextMatrix(0, .ColIndex("GDateH")) = "Date"
  .TextMatrix(0, .ColIndex("FrmTime")) = "From"
  .TextMatrix(0, .ColIndex("ToTime")) = "To"
  End With
  
  With Fg
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("FullCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Student Name"
  .TextMatrix(0, .ColIndex("UQama")) = "ID No."
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblStuGroup"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
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
 Dim fullcode As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, fullcode, 2, TxtUQama.Text
        DcbStudent.BoundText = EmpID
        TxtSudCode.Text = fullcode
    End If
End Sub

