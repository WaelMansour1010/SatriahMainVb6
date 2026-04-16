VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form frmDetailsAdoption 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   Icon            =   "frmDetailsAdoption.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13395
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
      TabIndex        =   7
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "frmDetailsAdoption.frx":6852
      Left            =   15480
      List            =   "frmDetailsAdoption.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   8
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
      TabIndex        =   9
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
            Picture         =   "frmDetailsAdoption.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailsAdoption.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   10
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
      ButtonImage     =   "frmDetailsAdoption.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   12
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
      ButtonImage     =   "frmDetailsAdoption.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   13
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
      ButtonImage     =   "frmDetailsAdoption.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9120
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      _cx             =   23627
      _cy             =   16087
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
         Height          =   750
         Left            =   13440
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   540
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   690
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
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10965
            TabIndex        =   0
            Top             =   120
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   315
            Left            =   8415
            TabIndex        =   1
            Top             =   120
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   78118913
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   6960
            TabIndex        =   51
            Top             =   120
            Width           =   1350
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   3600
            TabIndex        =   70
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCbSeason 
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·„«·Ì…"
            Height          =   180
            Index           =   0
            Left            =   2550
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   11
            Left            =   5910
            TabIndex        =   43
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   25
            Left            =   9795
            TabIndex        =   42
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   255
            Index           =   4
            Left            =   12510
            TabIndex        =   16
            Top             =   120
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   960
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8175
         Width           =   13365
         _cx             =   23574
         _cy             =   1693
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
            Left            =   12330
            TabIndex        =   18
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   570
            Width           =   870
            _ExtentX        =   1535
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
            ButtonImage     =   "frmDetailsAdoption.frx":15BA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   10965
            TabIndex        =   19
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   570
            Width           =   1125
            _ExtentX        =   1984
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
            ButtonImage     =   "frmDetailsAdoption.frx":1C40B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   10020
            TabIndex        =   2
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   570
            Width           =   885
            _ExtentX        =   1561
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
            ButtonImage     =   "frmDetailsAdoption.frx":22C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8820
            TabIndex        =   20
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   570
            Width           =   1035
            _ExtentX        =   1826
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
            ButtonImage     =   "frmDetailsAdoption.frx":23007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7830
            TabIndex        =   21
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
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
            ButtonImage     =   "frmDetailsAdoption.frx":233A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   330
            Left            =   5190
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… Œÿ«» «· —‘ÌÕ"
            Top             =   570
            Width           =   1185
            _ExtentX        =   2090
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
            ButtonImage     =   "frmDetailsAdoption.frx":2393B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   6480
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   570
            Visible         =   0   'False
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
            ButtonImage     =   "frmDetailsAdoption.frx":2A19D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   570
            Width           =   915
            _ExtentX        =   1614
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
            ButtonImage     =   "frmDetailsAdoption.frx":2A537
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   8400
            TabIndex        =   25
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
            Left            =   1200
            TabIndex        =   40
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   570
            Width           =   1065
            _ExtentX        =   1879
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
            ButtonImage     =   "frmDetailsAdoption.frx":2A8D1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   330
            Left            =   3960
            TabIndex        =   93
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   570
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«‰‘«¡ «·ﬁÌœ"
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
            ButtonImage     =   "frmDetailsAdoption.frx":31133
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   330
            Left            =   2640
            TabIndex        =   94
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   570
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬁÌœ"
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
            ButtonImage     =   "frmDetailsAdoption.frx":37995
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   315
            TabIndex        =   30
            Top             =   240
            Width           =   630
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   90
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   750
         Index           =   18
         Left            =   0
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Width           =   13440
         _cx             =   23707
         _cy             =   1323
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
            Height          =   285
            Left            =   135
            TabIndex        =   35
            Top             =   240
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
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
            ButtonImage     =   "frmDetailsAdoption.frx":37F2F
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   285
            Left            =   675
            TabIndex        =   36
            Top             =   240
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
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
            ButtonImage     =   "frmDetailsAdoption.frx":382C9
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   285
            Left            =   1350
            TabIndex        =   37
            Top             =   240
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
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
            ButtonImage     =   "frmDetailsAdoption.frx":38663
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   285
            Left            =   1950
            TabIndex        =   38
            Top             =   240
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
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
            ButtonImage     =   "frmDetailsAdoption.frx":389FD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   585
            Left            =   12375
            Picture         =   "frmDetailsAdoption.frx":38D97
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ÿ«·»…"
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
            Height          =   345
            Index           =   2
            Left            =   5655
            TabIndex        =   39
            Top             =   240
            Width           =   4665
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2070
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1170
         Width           =   13455
         _cx             =   23733
         _cy             =   3651
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
         Begin VB.TextBox TxtTotalValue 
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
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   1680
            Width           =   3855
         End
         Begin VB.TextBox TxtFATValue 
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
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   1680
            Width           =   1290
         End
         Begin VB.TextBox TxtFATYou 
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
            Height          =   285
            Left            =   8160
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1680
            Width           =   1290
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   10560
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   1320
            Width           =   1290
         End
         Begin VB.ComboBox DcbTypeClim 
            Height          =   315
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   570
            Width           =   4860
         End
         Begin VB.TextBox TxtNetValue 
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
            Height          =   285
            Left            =   10560
            MaxLength       =   50
            TabIndex        =   67
            Top             =   1650
            Width           =   1290
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
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   65
            Top             =   1290
            Width           =   1290
         End
         Begin VB.TextBox TxtTotal 
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
            Height          =   285
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   63
            Top             =   1290
            Width           =   1290
         End
         Begin VB.TextBox TxtDescription 
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
            Left            =   150
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   930
            Width           =   11700
         End
         Begin VB.TextBox TxtComanyNo 
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
            Left            =   10350
            MaxLength       =   50
            TabIndex        =   56
            Top             =   510
            Width           =   1500
         End
         Begin VB.TextBox TxtComanyName 
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
            Left            =   150
            MaxLength       =   50
            TabIndex        =   50
            Top             =   180
            Width           =   4860
         End
         Begin VB.TextBox TxtClaimNo 
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
            Left            =   10350
            MaxLength       =   50
            TabIndex        =   49
            Top             =   180
            Width           =   1500
         End
         Begin MSComCtl2.DTPicker ClaimDate 
            Height          =   285
            Left            =   7695
            TabIndex        =   54
            Top             =   180
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
            _Version        =   393216
            Format          =   78118913
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal ClaimDateH 
            Height          =   285
            Left            =   6315
            TabIndex        =   55
            Top             =   180
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
         End
         Begin MSDataListLib.DataCombo DcbStatClim 
            Height          =   315
            Left            =   6315
            TabIndex        =   58
            Top             =   570
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   5640
            TabIndex        =   96
            Top             =   1320
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "frmDetailsAdoption.frx":3A19C
            Height          =   315
            Left            =   0
            TabIndex        =   104
            Top             =   1440
            Visible         =   0   'False
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   68
            Left            =   4605
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›« "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   67
            Left            =   7125
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»…«·›« "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   66
            Left            =   9645
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   285
            Index           =   27
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·„»·€"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   3
            Left            =   12255
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   1680
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·Õ”„"
            Height          =   285
            Index           =   2
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì „»·€ «·„ÿ«·»…"
            Height          =   285
            Index           =   8
            Left            =   4200
            TabIndex        =   64
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·»Ì«‰"
            Height          =   285
            Index           =   3
            Left            =   12000
            TabIndex        =   61
            Top             =   990
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„ÿ«·»…"
            Height          =   285
            Index           =   1
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   570
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·„ÿ«·»…"
            Height          =   180
            Index           =   10
            Left            =   9270
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   570
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ „ÿ«·»… «·‘—ﬂ…"
            Height          =   285
            Index           =   1
            Left            =   12000
            TabIndex        =   57
            Top             =   570
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·„ÿ«·»…"
            Height          =   255
            Index           =   0
            Left            =   9120
            TabIndex        =   53
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·„ÿ«·»…"
            Height          =   255
            Index           =   5
            Left            =   12000
            TabIndex        =   52
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‘—ﬂ…  «·‰ﬁ·"
            Height          =   255
            Index           =   10
            Left            =   5040
            TabIndex        =   46
            Top             =   240
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2925
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   4800
         Width           =   13455
         _cx             =   23733
         _cy             =   5159
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
         ForeColor       =   8388608
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
         Begin XtremeSuiteControls.CheckBox ChekAll 
            Height          =   255
            Left            =   11760
            TabIndex        =   80
            Top             =   0
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " ÕœÌœ «·ﬂ·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   2235
            Left            =   120
            TabIndex        =   45
            Top             =   270
            Width           =   13185
            _cx             =   23257
            _cy             =   3942
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmDetailsAdoption.frx":3A1B1
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   0
            Left            =   12360
            TabIndex        =   47
            Top             =   2610
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
            ButtonImage     =   "frmDetailsAdoption.frx":3A33B
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   10200
            TabIndex        =   48
            Top             =   2610
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
            ButtonImage     =   "frmDetailsAdoption.frx":3A8D5
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   12
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   2550
            Width           =   1830
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«⁄ „«œ"
            Height          =   285
            Index           =   9
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   2550
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   8
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   0
            Width           =   1830
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«⁄ „«œ"
            Height          =   285
            Index           =   7
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ›«’Ì· «·«⁄ „«œ« "
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   6
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   0
            Width           =   1935
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1545
         Left            =   0
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3270
         Width           =   13455
         _cx             =   23733
         _cy             =   2725
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
         Begin VSFlex8Ctl.VSFlexGrid Fg2 
            Height          =   1185
            Left            =   120
            TabIndex        =   73
            Top             =   300
            Width           =   13185
            _cx             =   23257
            _cy             =   2090
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmDetailsAdoption.frx":3AE6F
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   285
            Index           =   2
            Left            =   12360
            TabIndex        =   74
            Top             =   4395
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   503
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
            ButtonImage     =   "frmDetailsAdoption.frx":3AFD5
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   285
            Index           =   3
            Left            =   10200
            TabIndex        =   75
            Top             =   4395
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
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
            ButtonImage     =   "frmDetailsAdoption.frx":3B56F
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ·«’… «·«⁄ „«œ« "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» «·„ﬁ»Ê·Ì‰ Õ«·Ì«"
            Height          =   210
            Index           =   9
            Left            =   6000
            TabIndex        =   77
            Top             =   4395
            Width           =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Height          =   210
            Index           =   4
            Left            =   4200
            TabIndex        =   76
            Top             =   4395
            Width           =   1725
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   11
         Left            =   0
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   7680
         Width           =   13455
         _cx             =   23733
         _cy             =   767
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
            Height          =   225
            Left            =   9270
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   90
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   9150
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   90
            Width           =   2955
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   300
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   90
            Width           =   1410
         End
         Begin MSDataListLib.DataCombo DcbUserVoucher 
            Height          =   315
            Left            =   225
            TabIndex        =   90
            Top             =   90
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   210
            Index           =   35
            Left            =   11970
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   90
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " „ «‰‘«¡ «·ﬁÌœ »Ê«”ÿ…"
            Height          =   210
            Index           =   28
            Left            =   4530
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   90
            Width           =   2460
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
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmDetailsAdoption"
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
Dim StrDepID As String
Function print_report(Optional NoteSerial As String)
On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.TblDetailsAdoption.ID, dbo.TblDetailsAdoption.RecordDate, dbo.TblDetailsAdoption.RecordDateH, dbo.TblDetailsAdoption.ClaimNo, "
 MySQL = MySQL & "                      dbo.TblDetailsAdoption.ClaimDate, dbo.TblDetailsAdoption.ClaimDateH, dbo.TblDetailsAdoption.ComanyName, dbo.TblDetailsAdoption.ComanyNo,"
 MySQL = MySQL & "                     dbo.TblDetailsAdoption.TypeClimID, dbo.TblDetailsAdoption.Description, dbo.TblDetailsAdoption.Total, dbo.TblDetailsAdoption.Discount,"
 MySQL = MySQL & "                      dbo.TblDetailsAdoption.NetValue, dbo.TblDetailsAdoption.TotalG, dbo.TblDetailsAdoption.TotalDet, dbo.TblDetailsAdoption.BranchID,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblDetailsAdoption.SeasonID, dbo.TblCompaniesGroup.Name,"
 MySQL = MySQL & "                      dbo.TblCompaniesGroup.NameE, dbo.TblDetailsAdoption.StatClimID, dbo.TblTypeClaim.Name AS ClaimName, dbo.TblTypeClaim.NameE AS ClaimNameE,"
 MySQL = MySQL & "                     dbo.TblDetailsAdoptionDet.TypeTrans, dbo.TblDetailsAdoptionDet.DepandNo, dbo.TblDetailsAdoptionDet.LargNo, dbo.TblDetailsAdoptionDet.SmalNo,"
 MySQL = MySQL & "                     dbo.TblDetailsAdoptionDet.TotalvALUE, dbo.TblDetailsAdoptionDet.PathID, dbo.TblShrines.Name AS PATHName, dbo.TblShrines.NameE AS PATHNameE,"
 MySQL = MySQL & "                     dbo.TblDetailsAdoptionDet.VehicleTypeID, dbo.TblVehicleType.Name AS ClassName, dbo.TblVehicleType.NameE AS ClassNameE, dbo.TblDetailsAdoption.CusID,"
 MySQL = MySQL & "                     dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode , dbo.TblDetailsAdoption.FATYou, dbo.TblDetailsAdoption.FATValue, "
 MySQL = MySQL & "                     dbo.TblDetailsAdoption.TotalValue AS TotalValueAll"
 MySQL = MySQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblDetailsAdoption ON dbo.TblCustemers.CusID = dbo.TblDetailsAdoption.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblVehicleType RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblDetailsAdoptionDet ON dbo.TblVehicleType.ID = dbo.TblDetailsAdoptionDet.VehicleTypeID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblShrines ON dbo.TblDetailsAdoptionDet.PathID = dbo.TblShrines.ID ON dbo.TblDetailsAdoption.ID = dbo.TblDetailsAdoptionDet.DetAdoID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblTypeClaim ON dbo.TblDetailsAdoption.StatClimID = dbo.TblTypeClaim.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCompaniesGroup ON dbo.TblDetailsAdoption.SeasonID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblDetailsAdoption.BranchID = dbo.TblBranchesData.branch_id"
 MySQL = MySQL & " Where (dbo.TblDetailsAdoption.ID = " & val(TxtSerial1.Text) & ")"
   
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDetailsAdoption.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDetailsAdoption.rpt"
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
    Dim DisCont As Double
    DisCont = GetDisValue()
    xReport.ParameterFields(4).AddCurrentValue DisCont
    xReport.ParameterFields(3).AddCurrentValue user_name
     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(11).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(11).AddCurrentValue GetRegVATNo(val(DcbBranch.BoundText))
    End If
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

 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "   „ÿ«·»… —ﬁ„ " & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblDetailsAdoption"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.Text)
Notevalue = 0
 notytype = 9069
Notevalue = val(txtTotal.Text)
BranchID = val(Me.DcbBranch.BoundText)
NoteDate = (RecordDate.value)
 
If Notevalue > 0 Then
                              
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, RecorddateH.value         ',
                                              TxtNoteID.Text = NoteID
                                                    TxtNoteSerial.Text = NoteSerial
                            
Cn.Execute " update TblDetailsAdoption set UserVouchID=" & user_id & " where ID=" & val(TxtSerial1.Text) & " "
Me.DcbUserVoucher.BoundText = user_id
CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent

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
    Dim I As Integer
    Dim valuee As Double
    Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
  LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID
    valuee = val(txtTotal.Text)
       StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.OutClientID.BoundText)) 'get_account_code_branch(137, my_branch)
      StrTempDes = " „ÿ«·»…  —ﬁ„  " & TxtSerial1.Text

             LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val(TxtNetValue.Text) + val(TxtFATValue.Text), 0, StrTempDes & "       Õ”«» ‰ﬁ«»… «·ÕÃ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = get_account_code_branch(144, my_branch)
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val(TxtDiscount.Text) / 1.05, 0, StrTempDes & "       Õ”«»  «·Õ”„Ì«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
      StrTempAccountCode = get_account_code_branch(138, my_branch)
          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, (txtTotal.Text / 1.05), 1, StrTempDes & "    Õ”«» «Ì—«œ«  «·ÕÃ  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            If val(TxtFATValue.Text) > 0 Then
               LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, Me.AccountVat.BoundText, val(TxtFATValue.Text), 1, "  Õ”«» «·ﬁÌ„… «·„÷«›… " & StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            End If
ErrTrap:
End Function

Private Sub ChekAll_Click()
  Dim I As Integer

    If ChekAll.value = vbChecked Then

        With Me.FG1
 
            For I = 1 To .Rows - 1
        
                .TextMatrix(I, .ColIndex("selected")) = True
            Next I

        End With

    Else

        With Me.FG1

            For I = 1 To .Rows - 1
        
                .TextMatrix(I, .ColIndex("selected")) = False
            Next I

        End With

    End If
    ReLineGrid
End Sub

Private Sub ClaimDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         ClaimDateH.value = ToHijriDate(ClaimDate.value)
End If
End Sub

Private Sub ClaimDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 ClaimDate.value = ToGregorianDate(ClaimDateH.value)
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
Dim I As Integer
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
ReLineGrid
Case 1
With FG1
For I = 1 To .Rows - 1
If val(DcbTypeClim.ListIndex) = 1 Then
Cn.Execute "update  TblEndorseTrans set FlagDepand=null  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
Else
Cn.Execute "update  TblEndorseTransMashar set FlagDepand=null  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
End If
Next I
     .Clear flexClearScrollable, flexClearEverything
     .Rows = 1
   ReLineGrid
 End With
End Select
End If
End Sub
Private Sub RemoveGridRow()
    With Me.FG1

        If .Row <= 0 Then Exit Sub
        If val(DcbTypeClim.ListIndex) = 1 Then
        Cn.Execute "update  TblEndorseTrans set FlagDepand=null  where ID=" & val(.TextMatrix(.Row, .ColIndex("DepandNo"))) & " "
        Else
        Cn.Execute "update  TblEndorseTransMashar set FlagDepand=null  where ID=" & val(.TextMatrix(.Row, .ColIndex("DepandNo"))) & " "
        End If
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Sub ReLineGrid()
Dim I As Integer
Dim Conter As Integer
Conter = 0
Dim SumValue As Double
SumValue = 0
StrDepID = ""
With FG1
For I = 1 To .Rows - 1
If .TextMatrix(I, .ColIndex("Name")) <> "" Then
Conter = Conter + 1
.TextMatrix(I, .ColIndex("Serial")) = Conter
 If .Cell(flexcpChecked, I, .ColIndex("selected")) = flexChecked Then
 SumValue = SumValue + val(.TextMatrix(I, .ColIndex("TotalvALUE")))
 StrDepID = StrDepID & val(.TextMatrix(I, .ColIndex("DepandNo"))) & ","
 End If
End If
Next I
 StrDepID = StrDepID & 0
lbl(12).Caption = SumValue
FullGrDataDepandGroup val(DcbTypeClim.ListIndex)
txtTotal.Text = SumValue
TxtTotal_Change
End With
''///////////
SumValue = 0
Conter = 0
With FG2
For I = 1 To .Rows - 1
If .TextMatrix(I, .ColIndex("Name")) <> "" Then
Conter = Conter + 1
.TextMatrix(I, .ColIndex("Serial")) = Conter
If val(.TextMatrix(I, .ColIndex("TotalvALUE"))) <> 0 Then
 SumValue = SumValue + val(.TextMatrix(I, .ColIndex("TotalvALUE")))
 End If
End If
Next I
lbl(8).Caption = SumValue
End With
ClculteVAT
End Sub
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblDetailsAdoption", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Private Sub Fill_Combos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetUsers Me.DcbUserVoucher
   Dcombos.GetAccountingCodes AccountVat
   Dcombos.GetTypeClaim DcbStatClim
   Dcombos.GetCompany OutClientID, 2, 0
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
 End If
 str = str & " where Omra_Hajj=1"
   fill_combo DCbSeason, str
 If SystemOptions.UserInterface = ArabicInterface Then
With DcbTypeClim
.Clear
.AddItem "„‘«⁄—"
.AddItem "ÕÃ"
End With
Else
With DcbTypeClim
.Clear
.AddItem "Mashare"
.AddItem "hajj"
End With
End If
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DCbSeason_Click(Area As Integer)
DcbTypeClim_Change
End Sub

Private Sub DcbTypeClim_Change()
If Me.TxtModFlg.Text = "E" Then
MsgBox "·«Ì„ﬂ‰  €Ì— ‰Ê⁄ «·„ÿ«·»… «Ê «·”‰… «·„«·Ì… Ì—ÃÏ Õ–› «·„ÿ«·»… ﬂ«„·«"
Exit Sub
End If
If Me.TxtModFlg.Text <> "R" Then
If val(DCbSeason.BoundText) <> 0 Then
If val(DcbTypeClim.ListIndex) <> -1 Then
If val(DcbTypeClim.ListIndex) = 1 Then
FullGrDataDepand 1, val(DCbSeason.BoundText)
Else
FullGrDataDepand 0, val(DCbSeason.BoundText)
End If
End If
End If
End If
End Sub

Private Sub DcbTypeClim_Click()
If Me.TxtModFlg.Text = "E" Then
MsgBox "·«Ì„ﬂ‰  €Ì— ‰Ê⁄ «·„ÿ«·»… «Ê «·”‰… «·„«·Ì… Ì—ÃÏ Õ–› «·„ÿ«·»… ﬂ«„·«"
Exit Sub
End If
DcbTypeClim_Change
End Sub

Private Sub Fg1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub

Private Sub Fg1_Click()
ReLineGrid
End Sub

Private Sub FG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub

Private Sub Fg2_Click()
ReLineGrid
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
      If SystemOptions.AllowCreateHajomraVoucher = True Then
        ISButton2.Enabled = True
        ISButton4.Enabled = True
      Ele(11).Enabled = True
     Else
        ISButton2.Enabled = False
        ISButton4.Enabled = False
        Ele(11).Enabled = False
 End If
 
    conection = "select * from  TblDetailsAdoption  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
Fill_Combos
   
   
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
Function CheCode() As Boolean
    Dim Rs6 As ADODB.Recordset
    Set Rs6 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from TblDetailsAdoption where id<>" & val(TxtSerial1.Text) & " and ClaimNo='" & TxtClaimNo.Text & "' "
    Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs6.RecordCount > 0 Then
    CheCode = True
    Else
    CheCode = False
    End If
End Function
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
   RsSavRec.Fields("FATYou").value = val(TxtFATYou.Text)
   RsSavRec.Fields("FATValue").value = val(TxtFATValue.Text)
   RsSavRec.Fields("TotalValue").value = val(TxtTotalValue.Text)
   RsSavRec.Fields("AccountCodeVat").value = Me.AccountVat.BoundText
   RsSavRec.Fields("RecordDateH").value = RecorddateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("SeasonID").value = val(Me.DCbSeason.BoundText)
   RsSavRec.Fields("ClaimNo").value = TxtClaimNo.Text
   RsSavRec.Fields("ClaimDate").value = ClaimDate.value
   RsSavRec.Fields("ClaimDateH").value = ClaimDateH.value
   RsSavRec.Fields("ComanyName").value = TxtComanyName.Text
   RsSavRec.Fields("ComanyNo").value = TxtComanyNo.Text
   RsSavRec.Fields("StatClimID").value = val(Me.DcbStatClim.BoundText)
   RsSavRec.Fields("TypeClimID").value = val(Me.DcbTypeClim.ListIndex)
   RsSavRec.Fields("Description").value = TxtDescription.Text
   RsSavRec.Fields("Total").value = val(txtTotal.Text)
   RsSavRec.Fields("Discount").value = val(TxtDiscount.Text)
   RsSavRec.Fields("NetValue").value = val(TxtNetValue.Text)
   RsSavRec.Fields("TotalG").value = val(lbl(8).Caption)
   RsSavRec.Fields("TotalDet").value = val(lbl(12).Caption)
   RsSavRec.Fields("StrDepID").value = StrDepID
   RsSavRec.Fields("CusID").value = val(Me.OutClientID.BoundText)
   RsSavRec.update
  ''//////////////////////////
  If Me.TxtModFlg.Text = "E" Then
 Cn.Execute "delete from TblDetailsAdoptionDet where DetAdoID=" & val(TxtSerial1.Text) & " "
 End If
  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblDetailsAdoptionDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim I As Integer
    With Me.FG1
       For I = .FixedRows To .Rows - 1
       If .Cell(flexcpChecked, I, .ColIndex("selected")) = flexChecked Then
       RsDevsub.AddNew
                RsDevsub("DetAdoID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TypeTrans").value = 0
                RsDevsub("DepandNo").value = IIf((.TextMatrix(I, .ColIndex("DepandNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("DepandNo"))))
                RsDevsub("PathID").value = IIf((.TextMatrix(I, .ColIndex("PathID"))) = "", Null, val(.TextMatrix(I, .ColIndex("PathID"))))
                RsDevsub("VehicleTypeID").value = IIf((.TextMatrix(I, .ColIndex("VehicleTypeID"))) = "", Null, val(.TextMatrix(I, .ColIndex("VehicleTypeID"))))
                RsDevsub("LargNo").value = IIf((.TextMatrix(I, .ColIndex("LargNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("LargNo"))))
                RsDevsub("SmalNo").value = IIf((.TextMatrix(I, .ColIndex("SmalNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("SmalNo"))))
                RsDevsub("TotalvALUE").value = IIf((.TextMatrix(I, .ColIndex("TotalvALUE"))) = "", Null, val(.TextMatrix(I, .ColIndex("TotalvALUE"))))
          RsDevsub.update
              If val(DcbTypeClim.ListIndex) = 1 Then
           Cn.Execute " update  TblEndorseTrans set FlagDepand=1  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
           Else
           Cn.Execute " update  TblEndorseTransMashar set FlagDepand=1  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
           End If
          Else
              If val(DcbTypeClim.ListIndex) = 1 Then
           Cn.Execute " update  TblEndorseTrans set FlagDepand=null  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
           Else
           Cn.Execute " update  TblEndorseTransMashar set FlagDepand=null  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
           End If
      End If
     Next I
    End With
  ''/////////////////
        Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblDetailsAdoptionDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    With Me.FG2
       For I = .FixedRows To .Rows - 1
       If val(.TextMatrix(I, .ColIndex("DepandNo"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("DetAdoID").value = val(Me.TxtSerial1.Text)
                RsDevsub("TypeTrans").value = 1
                RsDevsub("PathID").value = IIf((.TextMatrix(I, .ColIndex("PathID"))) = "", Null, val(.TextMatrix(I, .ColIndex("PathID"))))
                RsDevsub("DepandNo").value = IIf((.TextMatrix(I, .ColIndex("DepandNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("DepandNo"))))
                RsDevsub("VehicleTypeID").value = IIf((.TextMatrix(I, .ColIndex("VehicleTypeID"))) = "", Null, val(.TextMatrix(I, .ColIndex("VehicleTypeID"))))
                RsDevsub("LargNo").value = IIf((.TextMatrix(I, .ColIndex("LargNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("LargNo"))))
                RsDevsub("SmalNo").value = IIf((.TextMatrix(I, .ColIndex("SmalNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("SmalNo"))))
                RsDevsub("TotalvALUE").value = IIf((.TextMatrix(I, .ColIndex("TotalvALUE"))) = "", Null, val(.TextMatrix(I, .ColIndex("TotalvALUE"))))
       End If
      RsDevsub.update
     Next I
    End With
    FiLLTXT
'''///////////////
   
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
  Sub FullGri()
Dim Rs3 As ADODB.Recordset
Dim I As Integer
Dim sql As String
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = 1
    FG2.Clear flexClearScrollable, flexClearEverything
    FG2.Rows = 1
 Set Rs3 = New ADODB.Recordset
sql = "SELECT     dbo.TblDetailsAdoptionDet.ID, dbo.TblDetailsAdoptionDet.TypeTrans, dbo.TblDetailsAdoptionDet.DetAdoID, dbo.TblDetailsAdoptionDet.DepandNo, "
sql = sql & "                      dbo.TblDetailsAdoptionDet.LargNo, dbo.TblDetailsAdoptionDet.SmalNo, dbo.TblDetailsAdoptionDet.TotalvALUE, dbo.TblShrines.Name, dbo.TblShrines.NameE,"
sql = sql & "                      dbo.TblDetailsAdoptionDet.PathID, dbo.TblDetailsAdoptionDet.VehicleTypeID, dbo.TblVehicleType.Name AS VehicName,"
sql = sql & "                      dbo.TblVehicleType.NameE AS VehicNameE"
sql = sql & " FROM         dbo.TblDetailsAdoptionDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblVehicleType ON dbo.TblDetailsAdoptionDet.VehicleTypeID = dbo.TblVehicleType.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShrines ON dbo.TblDetailsAdoptionDet.PathID = dbo.TblShrines.ID"
sql = sql & " Where (dbo.TblDetailsAdoptionDet.TypeTrans = 0) And (dbo.TblDetailsAdoptionDet.DetAdoID = " & val(TxtSerial1.Text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("selected")) = 1
.TextMatrix(I, .ColIndex("DepandNo")) = IIf(IsNull(Rs3("DepandNo").value), "", Rs3("DepandNo").value)
.TextMatrix(I, .ColIndex("LargNo")) = IIf(IsNull(Rs3("LargNo").value), "", Rs3("LargNo").value)
.TextMatrix(I, .ColIndex("SmalNo")) = IIf(IsNull(Rs3("SmalNo").value), "", Rs3("SmalNo").value)
.TextMatrix(I, .ColIndex("TotalvALUE")) = IIf(IsNull(Rs3("TotalvALUE").value), "", Rs3("TotalvALUE").value)
.TextMatrix(I, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(I, .ColIndex("VehicleTypeID")) = IIf(IsNull(Rs3("VehicleTypeID").value), 0, Rs3("VehicleTypeID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicName").value), "", Rs3("VehicName").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicNameE").value), "", Rs3("VehicNameE").value)
End If
Rs3.MoveNext
Next I
End With
End If
''////////////////////////
 Set Rs3 = New ADODB.Recordset
sql = "SELECT     dbo.TblDetailsAdoptionDet.ID, dbo.TblDetailsAdoptionDet.TypeTrans, dbo.TblDetailsAdoptionDet.DetAdoID, dbo.TblDetailsAdoptionDet.DepandNo, "
sql = sql & "                      dbo.TblDetailsAdoptionDet.LargNo, dbo.TblDetailsAdoptionDet.SmalNo, dbo.TblDetailsAdoptionDet.TotalvALUE, dbo.TblShrines.Name, dbo.TblShrines.NameE,"
sql = sql & "                      dbo.TblDetailsAdoptionDet.PathID, dbo.TblDetailsAdoptionDet.VehicleTypeID, dbo.TblVehicleType.Name AS VehicName,"
sql = sql & "                      dbo.TblVehicleType.NameE AS VehicNameE"
sql = sql & " FROM         dbo.TblDetailsAdoptionDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblVehicleType ON dbo.TblDetailsAdoptionDet.VehicleTypeID = dbo.TblVehicleType.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShrines ON dbo.TblDetailsAdoptionDet.PathID = dbo.TblShrines.ID"
sql = sql & " Where (dbo.TblDetailsAdoptionDet.TypeTrans = 1) And (dbo.TblDetailsAdoptionDet.DetAdoID = " & val(TxtSerial1.Text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG2
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("DepandNo")) = IIf(IsNull(Rs3("DepandNo").value), "", Rs3("DepandNo").value)
.TextMatrix(I, .ColIndex("LargNo")) = IIf(IsNull(Rs3("LargNo").value), "", Rs3("LargNo").value)
.TextMatrix(I, .ColIndex("SmalNo")) = IIf(IsNull(Rs3("SmalNo").value), "", Rs3("SmalNo").value)
.TextMatrix(I, .ColIndex("TotalvALUE")) = IIf(IsNull(Rs3("TotalvALUE").value), "", Rs3("TotalvALUE").value)
.TextMatrix(I, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(I, .ColIndex("VehicleTypeID")) = IIf(IsNull(Rs3("VehicleTypeID").value), 0, Rs3("VehicleTypeID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicName").value), "", Rs3("VehicName").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicNameE").value), "", Rs3("VehicNameE").value)
End If
Rs3.MoveNext
Next I
End With
End If
End Sub
Sub FullGrDataDepand(Optional Typ As Integer = 0, Optional SesonID As Double)
Dim Rs3 As ADODB.Recordset
Dim I As Integer
Dim sql As String
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = 1
If Typ = 0 Then
 Set Rs3 = New ADODB.Recordset
sql = "SELECT     dbo.TblEndorseTransMashar.ID, dbo.TblEndorseTransMashar.SeasonsID, dbo.TblEndorseTransMashar.SDate, dbo.TblEndorseTransMashar.PathID, "
sql = sql & "                       dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblEndorseTransMashar.TotalPrice, dbo.TblEndorseTransMashar.TotOlds,"
sql = sql & "                       dbo.TblEndorseTransMashar.TotYoungs, dbo.TblEndorseTransMashar.VehicleType, dbo.TblVehicleType.Name AS VehicName,"
sql = sql & "                       dbo.TblVehicleType.NameE AS VehicNameE ,dbo.TblEndorseTransMashar.FlagDepand"
sql = sql & "  FROM         dbo.TblEndorseTransMashar LEFT OUTER JOIN"
sql = sql & "                       dbo.TblVehicleType ON dbo.TblEndorseTransMashar.VehicleType = dbo.TblVehicleType.ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblShrines ON dbo.TblEndorseTransMashar.PathID = dbo.TblShrines.ID"
If Me.TxtModFlg.Text = "N" Then
sql = sql & "  Where (dbo.TblEndorseTransMashar.SeasonsID = " & SesonID & ") and (dbo.TblEndorseTransMashar.FlagDepand is null)"
End If
If Me.TxtModFlg.Text = "E" Then
sql = sql & "  Where ((dbo.TblEndorseTransMashar.SeasonsID = " & SesonID & ") and (dbo.TblEndorseTransMashar.FlagDepand is null)) or (dbo.TblEndorseTransMashar.ID in(" & StrDepID & "))"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("DepandNo")) = IIf(IsNull(Rs3("ID").value), "", Rs3("ID").value)
.TextMatrix(I, .ColIndex("LargNo")) = IIf(IsNull(Rs3("TotOlds").value), "", Rs3("TotOlds").value)
.TextMatrix(I, .ColIndex("SmalNo")) = IIf(IsNull(Rs3("TotYoungs").value), "", Rs3("TotYoungs").value)
.TextMatrix(I, .ColIndex("TotalvALUE")) = IIf(IsNull(Rs3("TotalPrice").value), "", Rs3("TotalPrice").value)
.TextMatrix(I, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(I, .ColIndex("VehicleTypeID")) = IIf(IsNull(Rs3("VehicleType").value), 0, Rs3("VehicleType").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicName").value), "", Rs3("VehicName").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicNameE").value), "", Rs3("VehicNameE").value)
End If
Rs3.MoveNext
Next I
End With
End If
''////////////////////////
Else
 Set Rs3 = New ADODB.Recordset
sql = "SELECT     dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblVehicleType.Name AS VehicName, dbo.TblVehicleType.NameE AS VehicNameE, dbo.TblEndorseTrans.ID, "
sql = sql & "                      dbo.TblEndorseTrans.SDate, dbo.TblEndorseTrans.SeasonsID, dbo.TblEndorseTrans.TotalPrice, dbo.TblEndorseTrans.TotOlds, dbo.TblEndorseTrans.TotYoungs,"
sql = sql & "                      dbo.TblEndorseTrans.PathID , dbo.TblEndorseTrans.VehicleType ,dbo.TblEndorseTrans.FlagDepand"
sql = sql & " FROM         dbo.TblVehicleType RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblEndorseTrans ON dbo.TblVehicleType.ID = dbo.TblEndorseTrans.VehicleType LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShrines ON dbo.TblEndorseTrans.PathID = dbo.TblShrines.ID"
If Me.TxtModFlg.Text = "N" Then
sql = sql & " Where (dbo.TblEndorseTrans.SeasonsID = " & SesonID & ") and (dbo.TblEndorseTrans.FlagDepand is null)"
End If
If Me.TxtModFlg.Text = "E" Then
sql = sql & " Where ((dbo.TblEndorseTrans.SeasonsID = " & SesonID & ") and (dbo.TblEndorseTrans.FlagDepand is null)) or (dbo.TblEndorseTrans.ID in (" & StrDepID & ")) "
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("DepandNo")) = IIf(IsNull(Rs3("ID").value), "", Rs3("ID").value)
.TextMatrix(I, .ColIndex("LargNo")) = IIf(IsNull(Rs3("TotOlds").value), "", Rs3("TotOlds").value)
.TextMatrix(I, .ColIndex("SmalNo")) = IIf(IsNull(Rs3("TotYoungs").value), "", Rs3("TotYoungs").value)
.TextMatrix(I, .ColIndex("TotalvALUE")) = IIf(IsNull(Rs3("TotalPrice").value), "", Rs3("TotalPrice").value)
.TextMatrix(I, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(I, .ColIndex("VehicleTypeID")) = IIf(IsNull(Rs3("VehicleType").value), 0, Rs3("VehicleType").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicName").value), "", Rs3("VehicName").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicNameE").value), "", Rs3("VehicNameE").value)
End If
Rs3.MoveNext
Next I
End With
End If
End If
ReLineGrid
End Sub
Sub FullGrDataDepandGroup(Optional Typ As Integer = 0)
Dim Rs3 As ADODB.Recordset
Dim I As Integer
Dim sql As String
    FG2.Clear flexClearScrollable, flexClearEverything
    FG2.Rows = 1
If Typ = 0 Then
 Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblVehicleType.Name AS VehicName, dbo.TblVehicleType.NameE AS VehicNameE, "
sql = sql & "                       dbo.TblEndorseTransMashar.SeasonsID, SUM(dbo.TblEndorseTransMashar.TotalPrice) AS SmTotalPrice, SUM(dbo.TblEndorseTransMashar.TotOlds) AS SmTotOlds,"
sql = sql & "                      SUM(dbo.TblEndorseTransMashar.TotYoungs) AS SmTotYoungs, dbo.TblEndorseTransMashar.PathID, dbo.TblEndorseTransMashar.VehicleType, COUNT(dbo.TblEndorseTransMashar.ID)"
sql = sql & "                      AS CuntID"
sql = sql & " FROM         dbo.TblVehicleType RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblEndorseTransMashar ON dbo.TblVehicleType.ID = dbo.TblEndorseTransMashar.VehicleType LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShrines ON dbo.TblEndorseTransMashar.PathID = dbo.TblShrines.ID"
sql = sql & " WHERE     (dbo.TblEndorseTransMashar.ID IN (" & StrDepID & " ))"
sql = sql & " GROUP BY dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblEndorseTransMashar.SeasonsID, dbo.TblEndorseTransMashar.PathID, dbo.TblEndorseTransMashar.VehicleType,"
sql = sql & "                       dbo.TblVehicleType.name , dbo.TblVehicleType.NameE"

Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG2
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("DepandNo")) = IIf(IsNull(Rs3("CuntID").value), "", Rs3("CuntID").value)
.TextMatrix(I, .ColIndex("LargNo")) = IIf(IsNull(Rs3("SmTotOlds").value), "", Rs3("SmTotOlds").value)
.TextMatrix(I, .ColIndex("SmalNo")) = IIf(IsNull(Rs3("SmTotYoungs").value), "", Rs3("SmTotYoungs").value)
.TextMatrix(I, .ColIndex("TotalvALUE")) = IIf(IsNull(Rs3("SmTotalPrice").value), "", Rs3("SmTotalPrice").value)
.TextMatrix(I, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(I, .ColIndex("VehicleTypeID")) = IIf(IsNull(Rs3("VehicleType").value), 0, Rs3("VehicleType").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicName").value), "", Rs3("VehicName").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicNameE").value), "", Rs3("VehicNameE").value)
End If
Rs3.MoveNext
Next I
End With
End If

''////////////////////////
Else
 Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblVehicleType.Name AS VehicName, dbo.TblVehicleType.NameE AS VehicNameE, "
sql = sql & "                       dbo.TblEndorseTrans.SeasonsID, SUM(dbo.TblEndorseTrans.TotalPrice) AS SmTotalPrice, SUM(dbo.TblEndorseTrans.TotOlds) AS SmTotOlds,"
sql = sql & "                      SUM(dbo.TblEndorseTrans.TotYoungs) AS SmTotYoungs, dbo.TblEndorseTrans.PathID, dbo.TblEndorseTrans.VehicleType, COUNT(dbo.TblEndorseTrans.ID)"
sql = sql & "                      AS CuntID"
sql = sql & " FROM         dbo.TblVehicleType RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblEndorseTrans ON dbo.TblVehicleType.ID = dbo.TblEndorseTrans.VehicleType LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShrines ON dbo.TblEndorseTrans.PathID = dbo.TblShrines.ID"
sql = sql & " WHERE     (dbo.TblEndorseTrans.ID IN (" & StrDepID & " ))"
sql = sql & " GROUP BY dbo.TblShrines.Name, dbo.TblShrines.NameE, dbo.TblEndorseTrans.SeasonsID, dbo.TblEndorseTrans.PathID, dbo.TblEndorseTrans.VehicleType,"
sql = sql & "                       dbo.TblVehicleType.name , dbo.TblVehicleType.NameE"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG2
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For I = 1 To .Rows - 1
.TextMatrix(I, .ColIndex("Serial")) = I
.TextMatrix(I, .ColIndex("DepandNo")) = IIf(IsNull(Rs3("CuntID").value), "", Rs3("CuntID").value)
.TextMatrix(I, .ColIndex("LargNo")) = IIf(IsNull(Rs3("SmTotOlds").value), "", Rs3("SmTotOlds").value)
.TextMatrix(I, .ColIndex("SmalNo")) = IIf(IsNull(Rs3("SmTotYoungs").value), "", Rs3("SmTotYoungs").value)
.TextMatrix(I, .ColIndex("TotalvALUE")) = IIf(IsNull(Rs3("SmTotalPrice").value), "", Rs3("SmTotalPrice").value)
.TextMatrix(I, .ColIndex("PathID")) = IIf(IsNull(Rs3("PathID").value), 0, Rs3("PathID").value)
.TextMatrix(I, .ColIndex("VehicleTypeID")) = IIf(IsNull(Rs3("VehicleType").value), 0, Rs3("VehicleType").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicName").value), "", Rs3("VehicName").value)
Else
.TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(I, .ColIndex("VehicName")) = IIf(IsNull(Rs3("VehicNameE").value), "", Rs3("VehicNameE").value)
End If
Rs3.MoveNext
Next I
End With
End If
End If

End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()

   On Error GoTo ErrTrap
    Dim I As Integer
    Dim Shifttime As Date
     TxtFATYou.Text = IIf(IsNull(RsSavRec.Fields("FATYou").value), 0, (RsSavRec.Fields("FATYou").value))
     TxtFATValue.Text = IIf(IsNull(RsSavRec.Fields("FATValue").value), 0, (RsSavRec.Fields("FATValue").value))
     TxtTotalValue.Text = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, (RsSavRec.Fields("TotalValue").value))
     Me.AccountVat.BoundText = IIf(IsNull(RsSavRec.Fields("AccountCodeVat").value), "", (RsSavRec.Fields("AccountCodeVat").value))
    TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
    TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
    DcbUserVoucher.BoundText = IIf(IsNull(RsSavRec.Fields("UserVouchID").value), "", RsSavRec.Fields("UserVouchID").value)
    Me.OutClientID.BoundText = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    RecorddateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DCbSeason.BoundText = IIf(IsNull(RsSavRec.Fields("SeasonID").value), "", RsSavRec.Fields("SeasonID").value)
    TxtClaimNo.Text = IIf(IsNull(RsSavRec.Fields("ClaimNo").value), "", RsSavRec.Fields("ClaimNo").value)
    ClaimDate.value = IIf(IsNull(RsSavRec.Fields("ClaimDate").value), Date, RsSavRec.Fields("ClaimDate").value)
    ClaimDateH.value = IIf(IsNull(RsSavRec.Fields("ClaimDateH").value), ToHijriDate(Date), RsSavRec.Fields("ClaimDateH").value)
    TxtComanyName.Text = IIf(IsNull(RsSavRec.Fields("ComanyName").value), "", RsSavRec.Fields("ComanyName").value)
    TxtComanyNo.Text = IIf(IsNull(RsSavRec.Fields("ComanyNo").value), "", RsSavRec.Fields("ComanyNo").value)
    Me.DcbStatClim.BoundText = IIf(IsNull(RsSavRec.Fields("StatClimID").value), "", RsSavRec.Fields("StatClimID").value)
    Me.DcbTypeClim.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeClimID").value), -1, RsSavRec.Fields("TypeClimID").value)
    TxtDescription.Text = IIf(IsNull(RsSavRec.Fields("Description").value), "", RsSavRec.Fields("Description").value)
    txtTotal.Text = IIf(IsNull(RsSavRec.Fields("Total").value), "", RsSavRec.Fields("Total").value)
    TxtDiscount.Text = IIf(IsNull(RsSavRec.Fields("Discount").value), "", RsSavRec.Fields("Discount").value)
    TxtNetValue.Text = IIf(IsNull(RsSavRec.Fields("NetValue").value), "", RsSavRec.Fields("NetValue").value)
    lbl(8).Caption = IIf(IsNull(RsSavRec.Fields("TotalG").value), "", RsSavRec.Fields("TotalG").value)
    lbl(12).Caption = IIf(IsNull(RsSavRec.Fields("TotalDet").value), "", RsSavRec.Fields("TotalDet").value)
    StrDepID = IIf(IsNull(RsSavRec.Fields("StrDepID").value), "", RsSavRec.Fields("StrDepID").value)
 TxtDiscount.Text = GetDisValue()
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
TxtNetValue.Text = Round((val(Me.txtTotal.Text) - val(TxtDiscount.Text)) / 1.05, 2)
ErrTrap:
End Sub
Function ChekDeD() As Boolean
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblDetailsAdoption where id=" & val(TxtSerial1.Text) & " and FlagDeduc =1 "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
ChekDeD = True
Else
ChekDeD = False
End If
End Function
Function GetDisValue() As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblDeduction where ClaimNo=" & val(TxtClaimNo.Text) & " "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetDisValue = IIf(IsNull(Rs4("Discount").value), 0, Rs4("Discount").value)
Else
GetDisValue = 0
End If
End Function
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
    DcbBranch.SetFocus
    Exit Sub
    End If
If val(DCbSeason.BoundText) = 0 Or DCbSeason.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·”‰… «·„«·Ì…"
Else
MsgBox "Please select the financial year "
End If
DCbSeason.SetFocus
Exit Sub
End If
If val(DcbTypeClim.ListIndex) = -1 Or DcbTypeClim.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·„ÿ«·»…"
Else
MsgBox "Please select type "
End If
DcbTypeClim.SetFocus
Exit Sub
End If
If CheCode() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "—ﬁ„ «·„ÿ«·»… „ÊÃÊœ „”»ﬁ«"
Else
MsgBox "This is No already exists"
End If
Exit Sub
End If
If val(OutClientID.BoundText) = 0 Or OutClientID.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
Else
MsgBox "Please Select Customer"
End If
OutClientID.SetFocus
Exit Sub
End If
Dim AccountVATDept As String
If AccountVat.BoundText = "" And True = True And CheckAnyVAT = True Then
MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
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
   
   If ChekClodePeriod(RecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox " This is Period is Closed"
              End If
              Exit Sub
  End If
Cn.Execute "update TblDetailsAdoption set FATValue=" & val(TxtFATValue) & ",NetValue=" & val(TxtNetValue) & " where id=" & TxtSerial1 & " "
If TxtNoteSerial.Text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If
If val(OutClientID.BoundText) = 0 Or OutClientID.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ Ì—ÃÏ «Œ Ì«— «·⁄„Ì· «Ê·«"
Else
MsgBox "Please Select Customer"
End If
OutClientID.SetFocus
Exit Sub
End If
If TxtNoteSerial.Text = "" Then
    Dim Account_Code_dynamic As String
'   Account_Code_dynamic = get_account_code_branch(137, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'         Else
'                MsgBox "Please Create Branch"
'        End If
'                Exit Sub
'            Else
'
'                If Account_Code_dynamic = "NO account" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ‰ﬁ«»… «·ÕÃ", vbCritical
'                 Else
'                 MsgBox "Please Select Account"
'                End If
'                   Exit Sub
'                End If
  '          End If
      Account_Code_dynamic = get_account_code_branch(138, my_branch)

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
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «Ì—«œ«  «·ÕÃ", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
       Account_Code_dynamic = get_account_code_branch(144, my_branch)

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
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   Õ”„Ì«  «·ÕÃ", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
        Calculte
createVoucher
updateNotesValueAndNobytext (val(TxtNoteID.Text))
MsgBox " „ «‰‘«¡ «·ﬁÌœ"
Else
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
End If
End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.Text, "251020162016"
ErrTrap:



End Sub

Private Sub ISButton4_Click()
Dim StrSQL As String
        If ChekClodePeriod(RecordDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «·Õ–›   ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox " This is Period is Closed"
           End If
              Exit Sub
        End If
        
       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
       Cn.Execute "Update TblDetailsAdoption set NoteSerial=null,UserVouchID=null ,NoteID=null where id=" & val(TxtSerial1.Text) & ""
       DcbUserVoucher.BoundText = ""
TxtNoteSerial.Text = ""
TxtNoteID.Text = 0
MsgBox " „ Õ–› «·ﬁÌœ"
RsSavRec.Resync adAffectCurrent
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub OutClientID_Change()
OutClientID_Click (0)
End Sub

Private Sub OutClientID_Click(Area As Integer)
   Dim Fullcode As String
    GetCustomersDetail val(OutClientID.BoundText), , Fullcode, 1
    Text2.Text = Fullcode
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecorddateH.value = ToHijriDate(RecordDate.value)
End If
ClculteVAT
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecorddateH.value)
End If
ClculteVAT
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.Text, 1
        OutClientID.BoundText = CUSTID
    End If
End Sub

Private Sub txtDiscount_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNetValue.Text = (val(Me.txtTotal.Text) - val(TxtDiscount.Text)) / 1.05
End If
Calculte
End Sub
Sub ClculteVAT()
If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim account As String
PercentgValueAddedAccount_Transec ClaimDate.value, 4, 1, account, Percetage
TxtFATYou.Text = Percetage
AccountVat.BoundText = account
Calculte
End If
End Sub
Sub Calculte()

'If Me.TxtModFlg.Text <> "R" Then
If val(TxtFATYou.Text) > 0 Then
TxtFATValue.Text = (val(TxtNetValue.Text) * val(TxtFATYou.Text)) / 100
Else
TxtFATValue.Text = 0
End If


TxtTotalValue.Text = val(TxtNetValue.Text) + val(TxtFATValue.Text)
'End If
End Sub

Private Sub TxtNetValue_Change()
Calculte
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
    Dim I As Integer
    Dim ID As Double

    If ChekDeD() = True Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·„ÿ«·»… „— »ÿ… »«·Õ”„Ì« "
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
         Cn.Execute "Delete from TblDetailsAdoptionDet where DetAdoID=" & val(TxtSerial1.Text) & " "
         With FG1
         For I = 1 To .Rows - 1
         If val(DcbTypeClim.ListIndex) = 1 Then
           Cn.Execute " update  TblEndorseTrans set FlagDepand=null  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
           Else
           Cn.Execute " update  TblEndorseTransMashar set FlagDepand=null  where ID=" & val(.TextMatrix(I, .ColIndex("DepandNo"))) & " "
           End If
          Next I
          End With
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
                  FG1.Clear flexClearScrollable, flexClearEverything
                  FG1.Rows = 1
                  FG2.Clear flexClearScrollable, flexClearEverything
                  FG2.Rows = 1
                  lbl(8).Caption = 0
                  lbl(12).Caption = 0
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            
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
DcbTypeClim.locked = False
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
   DcbTypeClim.locked = True
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
    FG1.Rows = FG1.Rows + 1
      If ChekDeD() = True Then
    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·„ÿ«·»… „— »ÿ… »«·Õ”„Ì« "
    Exit Sub
    End If
If TxtNoteSerial.Text <> "" Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Exit Sub
End If
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
        If Me.TxtModFlg.Text <> "R" Then
If val(DCbSeason.BoundText) <> 0 Then
If val(DcbTypeClim.ListIndex) <> -1 Then
If val(DcbTypeClim.ListIndex) = 1 Then
FullGrDataDepand 1, val(DCbSeason.BoundText)
Else
FullGrDataDepand 0, val(DCbSeason.BoundText)
End If
End If
End If
End If
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
    Me.DcbBranch.BoundText = Current_branch
      FG1.Clear flexClearScrollable, flexClearEverything
     FG1.Rows = 1
     FG2.Clear flexClearScrollable, flexClearEverything
     FG2.Rows = 1
    Me.DCboUserName.BoundText = user_id
        Dim cCompanyInfo As New ClsCompanyInfo
        TxtComanyName.Text = cCompanyInfo.ArabCompanyName
    lbl(8).Caption = 0
    lbl(12).Caption = 0
    ClculteVAT
 DCbSeason.BoundText = GetMosim(1)
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
  Label1(2).Caption = "Approve the Nomination"
lbl(4).Caption = "No"
lbl(11).Caption = "Branch"
ISButton3.Caption = "Attachments"
lbl(25).Caption = "Date"
Label1(0).Caption = "No.Nomination"
Label1(5).Caption = "Companies"
Label1(1).Caption = "No. Student "
Label1(3).Caption = "Remarks"
Label1(14).Caption = "No. Remaining"
Label1(10).Caption = "No. Accepted"
Label1(4).Caption = "Supervisor"
Label1(6).Caption = "No. Accepted"
lbl(0).Caption = "Date"
 
 
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

  
  With FG1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("selected")) = "Approve"
  .TextMatrix(0, .ColIndex("Name")) = "Student Name"
  .TextMatrix(0, .ColIndex("Supervisor")) = "Supervisor"
  .TextMatrix(0, .ColIndex("QName")) = "Course"
  .TextMatrix(0, .ColIndex("CursNoHour")) = "Number of Hours"
  .TextMatrix(0, .ColIndex("CursValue")) = "Value"
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblDetailsAdoption"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub TxtTotal_Change()
'If Me.TxtModFlg.Text <> "R" Then
TxtNetValue.Text = (val(Me.txtTotal.Text) - val(TxtDiscount.Text)) / 1.05

'End If
Calculte
End Sub
