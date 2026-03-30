VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_NewMeasure 
   ClientHeight    =   10575
   ClientLeft      =   1425
   ClientTop       =   2985
   ClientWidth     =   16380
   Icon            =   "Frm_NewMeasure.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10575
   ScaleWidth      =   16380
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   288
      ItemData        =   "Frm_NewMeasure.frx":6852
      Left            =   16680
      List            =   "Frm_NewMeasure.frx":6862
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
      Left            =   16800
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
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16920
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   312
      Left            =   17040
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
      Height          =   312
      Left            =   16680
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   16800
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
            Picture         =   "Frm_NewMeasure.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_NewMeasure.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   336
      Left            =   16800
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
      ButtonImage     =   "Frm_NewMeasure.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   408
      Left            =   17160
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Visible         =   0   'False
      Width           =   1248
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
      ButtonImage     =   "Frm_NewMeasure.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   336
      Left            =   18120
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
      ButtonImage     =   "Frm_NewMeasure.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   10575
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   16380
      _cx             =   28893
      _cy             =   18653
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   10515
         Left            =   0
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   16290
         _cx             =   28734
         _cy             =   18547
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
            Height          =   705
            Left            =   0
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   9075
            Width           =   16290
            _cx             =   28734
            _cy             =   1244
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
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   630
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   0
               Width           =   4830
               Begin VB.Label LabCurrRec 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H00800000&
                  Height          =   276
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ã· «·Õ«·Ì:"
                  Height          =   270
                  Index           =   0
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·”Ã·« :"
                  Height          =   276
                  Index           =   1
                  Left            =   1164
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   240
                  Width           =   960
               End
               Begin VB.Label LabCountRec 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H00C00000&
                  Height          =   276
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   240
                  Width           =   768
               End
            End
            Begin MSDataListLib.DataCombo DCboUserName 
               Height          =   315
               Left            =   7305
               TabIndex        =   19
               Top             =   120
               Width           =   5430
               _ExtentX        =   9578
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Õ—— »Ê«”ÿ…  "
               Height          =   225
               Index           =   8
               Left            =   13035
               TabIndex        =   20
               Top             =   135
               Width           =   1380
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   735
            Left            =   0
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   9780
            Width           =   16290
            _cx             =   28734
            _cy             =   1296
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
               Height          =   480
               Left            =   14310
               TabIndex        =   22
               ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":15BA9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   480
               Left            =   10755
               TabIndex        =   23
               ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   135
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":1C40B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   480
               Left            =   12510
               TabIndex        =   24
               ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
               Top             =   135
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":1C7A5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   480
               Left            =   8700
               TabIndex        =   25
               ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
               Top             =   135
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":23007
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   480
               Left            =   6945
               TabIndex        =   26
               ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
               Top             =   135
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":233A1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   480
               Left            =   1185
               TabIndex        =   27
               ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
               Top             =   135
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":2393B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton5 
               Height          =   555
               Left            =   5385
               TabIndex        =   28
               TabStop         =   0   'False
               ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
               Top             =   105
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   979
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
               ButtonImage     =   "Frm_NewMeasure.frx":23CD5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton8 
               Height          =   480
               Left            =   3705
               TabIndex        =   29
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   135
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   847
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
               ButtonImage     =   "Frm_NewMeasure.frx":2A537
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic100 
            Height          =   900
            Left            =   0
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   765
            Width           =   16290
            _cx             =   28734
            _cy             =   1588
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
            Align           =   1
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
            Begin VB.TextBox txtRowNo 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   5040
               TabIndex        =   59
               Text            =   "1"
               Top             =   465
               Width           =   870
            End
            Begin VB.TextBox TxtSerial1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   13950
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Text            =   "0"
               Top             =   60
               Width           =   945
            End
            Begin VB.TextBox txtCustomerName 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   11340
               TabIndex        =   35
               Top             =   435
               Width           =   3555
            End
            Begin VB.TextBox TXT_District 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   3345
               TabIndex        =   34
               Text            =   "0"
               Top             =   450
               Width           =   975
            End
            Begin VB.TextBox TXT_Time 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   3345
               TabIndex        =   33
               Text            =   "0"
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox TXT_City 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   7005
               TabIndex        =   32
               Text            =   "0"
               Top             =   450
               Width           =   870
            End
            Begin VB.TextBox TXT_Mobile 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   8460
               TabIndex        =   31
               Text            =   "0"
               Top             =   450
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTP_Order 
               Height          =   270
               Left            =   1140
               TabIndex        =   36
               Top             =   150
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   476
               _Version        =   393216
               Format          =   170065921
               CurrentDate     =   43122
            End
            Begin MSComCtl2.DTPicker DTP_Measure 
               Height          =   270
               Left            =   1140
               TabIndex        =   37
               Top             =   495
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   476
               _Version        =   393216
               Format          =   170000385
               CurrentDate     =   43122
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·«œÊ«—"
               Height          =   210
               Left            =   5955
               TabIndex        =   60
               Top             =   510
               Width           =   735
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÿ·»"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   15075
               TabIndex        =   58
               Top             =   60
               Width           =   1095
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·«„—"
               Height          =   210
               Left            =   2490
               TabIndex        =   44
               Top             =   120
               Width           =   660
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—»Œ «·ﬁÌ«”"
               Height          =   210
               Left            =   2430
               TabIndex        =   43
               Top             =   495
               Width           =   810
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÕÏ"
               Height          =   210
               Left            =   4425
               TabIndex        =   42
               Top             =   495
               Width           =   300
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Êﬁ "
               Height          =   210
               Left            =   4440
               TabIndex        =   41
               Top             =   120
               Width           =   345
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œÌ‰…"
               Height          =   210
               Left            =   7920
               TabIndex        =   40
               Top             =   495
               Width           =   375
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÃÊ«·"
               Height          =   210
               Left            =   10920
               TabIndex        =   39
               Top             =   495
               Width           =   375
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ì·"
               Height          =   195
               Index           =   2
               Left            =   15480
               TabIndex        =   38
               Top             =   435
               Width           =   660
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   765
            Left            =   0
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Width           =   16290
            _cx             =   28734
            _cy             =   1349
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
            Align           =   1
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
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   0
               Width           =   17730
               Begin VB.TextBox tXTRootAccount 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   852
               End
               Begin VB.TextBox TxtName 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   0
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1572
               End
               Begin ImpulseButton.ISButton btnLast 
                  Height          =   312
                  Left            =   336
                  TabIndex        =   49
                  Top             =   120
                  Width           =   408
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
                  ButtonImage     =   "Frm_NewMeasure.frx":2A8D1
                  ColorButton     =   16777215
                  AcclimateGrayTones=   -1  'True
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnNext 
                  Height          =   312
                  Left            =   792
                  TabIndex        =   50
                  Top             =   120
                  Width           =   408
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
                  ButtonImage     =   "Frm_NewMeasure.frx":2AC6B
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnPrevious 
                  Height          =   312
                  Left            =   1392
                  TabIndex        =   51
                  Top             =   120
                  Width           =   408
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
                  ButtonImage     =   "Frm_NewMeasure.frx":2B005
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnFirst 
                  Height          =   312
                  Left            =   1920
                  TabIndex        =   52
                  Top             =   120
                  Width           =   408
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
                  ButtonImage     =   "Frm_NewMeasure.frx":2B39F
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin VB.Image Image1 
                  Height          =   615
                  Left            =   13320
                  Picture         =   "Frm_NewMeasure.frx":2B739
                  Stretch         =   -1  'True
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÿ·» —›⁄ «·ﬁÌ«”"
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
                  Height          =   495
                  Index           =   2
                  Left            =   9120
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   120
                  Width           =   3480
               End
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   6465
            Left            =   90
            TabIndex        =   54
            Top             =   2115
            Width           =   16080
            _cx             =   28363
            _cy             =   11404
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
            Cols            =   73
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Frm_NewMeasure.frx":2CB3E
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
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   435
            Left            =   14475
            TabIndex        =   55
            Top             =   8670
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   767
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
            ButtonImage     =   "Frm_NewMeasure.frx":2D7C8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   435
            Left            =   12810
            TabIndex        =   56
            Top             =   8670
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   767
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
            ButtonImage     =   "Frm_NewMeasure.frx":2DD62
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
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
      Height          =   276
      Index           =   13
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   852
   End
End
Attribute VB_Name = "Frm_NewMeasure"
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
 Dim Account_Code_dynamic As String
 Dim RevenueAccount As String
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double


 

 

Private Sub Cmd_DeleteAll_Click()
If Me.TxtModFlg.text <> "R" Then


 
            GridInstallments.Rows = 2
            GridInstallments.Rows = 3

End If
End Sub

Private Sub Cmd_DeleteRow_Click()
If Me.TxtModFlg.text <> "R" Then

RemoveGridRow

End If
End Sub
Private Sub RemoveGridRow()

    With Me.GridInstallments
'MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        If .Row > .FixedRows - 1 Then
        
            .RemoveItem .Row
        End If
        If .Rows = .FixedRows Then
            .Rows = .Rows + 1
        End If
        
        End If
    End With
End Sub

Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
Case 1
   End Select
End If
End Sub

Private Sub Command1_Click()


End Sub

Private Sub Command2_Click()
 End Sub

 
 



    Private Sub Form_Load()
   ' On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    'If SystemOptions.UserInterface = ArabicInterface Then
   
   ' With DcbConstStatus
   ' .Clear
   ' .AddItem "·„ Ì „ «·»‰«¡"
   ' .AddItem " Õ  «·«‰‘«¡"
   ' .AddItem " „ «·»‰«¡"
   ' .AddItem "„ Êﬁ›"
   ' End With
   'Else
   '    With DcbConstStatus
   ' .Clear
   ' .AddItem "Not Built"
   ' .AddItem "Under Construction"
   ' .AddItem "Built"
   ' .AddItem "Stopped"
   ' End With
   'With DcbTypePrstg
   ' .Clear
   ' .AddItem "Value"
   ' .AddItem "Percentage"
   ' End With
 '  End If
   '   If SystemOptions.UserInterface = ArabicInterface Then
   '             Grid.ColComboList(Grid.ColIndex("ConstStatus")) = "#1;·„ Ì „ «·»‰«¡|#2; Õ  «·≈‰‘«¡|#3; „ «·»‰«¡|#4;„ Êﬁ›"
   '         ElseIf SystemOptions.UserInterface = EnglishInterface Then
   '            Grid.ColComboList(Grid.ColIndex("ConstStatus")) = "#1;Not Built |#2;Under Construction|#3; Built |#4;Stopped"
   '         End If
   

 
 
   
    conection = "select * from TBL_measureMent order by ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
'    Dcombos.GetCustomersSuppliers 1, Me.DcbCustmer
    Dcombos.GetUsers Me.DCboUserName
    TXT_Time.text = Format(Time, "hh : mm AM/PM")
    Me.DCboUserName.BoundText = user_id
    BtnLast_Click
    GridInstallments.FixedRows = 2
   GridInstallments.MergeCells = flexMergeFixedOnly
    GridInstallments.MergeRow(0) = True
     GridInstallments.MergeRow(1) = True
     Dim i As Long
    For i = 1 To GridInstallments.Cols - 1
        GridInstallments.MergeCol(i) = True
    Next
'
'    ShowTip
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
'   Me.Refresh
'ErrTrap:

End Sub


Public Sub FiLLRec()
  
  ' On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    If Me.TxtModFlg.text = "E" Then
    'StrSQL = "Delete From TBL_measureMent Where ID =" & val(TxtSerial1.Text) & ""
   ' Cn.Execute StrSQL, , adExecuteNoRecords
    End If
   
    RsSavRec.Fields("CustomerName").value = txtCustomerName.text
    RsSavRec.Fields("Cust_Mobile").value = Txt_Mobile.text

    RsSavRec.Fields("Cust_Time").value = TXT_Time.text
    RsSavRec.Fields("Cust_District").value = TXT_District
    RsSavRec.Fields("Date_Order").value = DTP_Order.value
    RsSavRec.Fields("Date_measureMent").value = DTP_Measure.value
     RsSavRec.Fields("Cust_City").value = Txt_City
     
   
   
    RsSavRec.Fields("UserID").value = DCboUserName.BoundText
       
    RsSavRec.update
  
''//////////////////////////
   ' Set RsDevsub = New ADODB.Recordset
   ' StrSQL = "SELECT  *  from TblBookingBondsInvsDet Where (1 = -1)"
   ' RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
   ' Dim RsMost As ADODB.Recordset
   ' Set RsMost = New ADODB.Recordset
    
   ' StrSQL = "SELECT  *  from Land_Planner Where Land_Name = '" & CommonDialog1.FileTitle & "'"
   ' RsMost.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Dim i As Integer
    Dim str2 As String
    'With Me.Grid
    '   For I = .FixedRows To .Rows - 1
    '   If .TextMatrix(I, .ColIndex("BlockNo")) <> "" Then
    '   RsDevsub.AddNew
    '            RsDevsub("TypeTrans").value = 0
    '            RsDevsub("BokID").value = val(Me.TxtSerial1.Text)
    '            RsDevsub("SquareID").value = IIf((.TextMatrix(I, .ColIndex("SquareID"))) = "", Null, val(.TextMatrix(I, .ColIndex("SquareID"))))
    '            RsDevsub("BlockNo").value = IIf((.TextMatrix(I, .ColIndex("BlockNo"))) = "", Null, (.TextMatrix(I, .ColIndex("BlockNo"))))
    '            RsDevsub("PartID").value = IIf((.TextMatrix(I, .ColIndex("PartID"))) = "", Null, val(.TextMatrix(I, .ColIndex("PartID"))))
    '            RsDevsub("ModelID").value = IIf((.TextMatrix(I, .ColIndex("ModelID"))) = "", Null, val(.TextMatrix(I, .ColIndex("ModelID"))))
    '            RsDevsub("Name").value = IIf((.TextMatrix(I, .ColIndex("Name"))) = "", Null, (.TextMatrix(I, .ColIndex("Name"))))
    '            RsDevsub("BedroomsNo").value = IIf((.TextMatrix(I, .ColIndex("BedroomsNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("BedroomsNo"))))
    '            RsDevsub("ConstStatus").value = IIf((.TextMatrix(I, .ColIndex("ConstStatus"))) = "", Null, val(.TextMatrix(I, .ColIndex("ConstStatus"))))
    '            RsDevsub("MOHPrice").value = IIf((.TextMatrix(I, .ColIndex("MOHPrice"))) = "", Null, val(.TextMatrix(I, .ColIndex("MOHPrice"))))
    '            RsDevsub("LandArea").value = IIf((.TextMatrix(I, .ColIndex("LandArea"))) = "", Null, val(.TextMatrix(I, .ColIndex("LandArea"))))
    '            RsDevsub("Remarks").value = IIf((.TextMatrix(I, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(I, .ColIndex("Remarks"))))
    '            RsDevsub("HouseArea").value = IIf((.TextMatrix(I, .ColIndex("HouseArea"))) = "", Null, val(.TextMatrix(I, .ColIndex("HouseArea"))))
    '            RsDevsub("Total").value = IIf((.TextMatrix(I, .ColIndex("Total"))) = "", 0, .TextMatrix(I, .ColIndex("Total")))
    '            RsDevsub("AddValue").value = IIf((.TextMatrix(I, .ColIndex("AddValue"))) = "", 0, .TextMatrix(I, .ColIndex("AddValue")))
    '            RsDevsub("ValueOffice").value = IIf((.TextMatrix(I, .ColIndex("ValueOffice"))) = "", 0, val(.TextMatrix(I, .ColIndex("ValueOffice"))))
    '   RsDevsub.update
    '  End If
    ' Next I
    'End With
    
    
            
            
            If Me.TxtModFlg.text = "E" Then
            
                'StrSQL = "Delete From TblTransactionInvest Where BuyBilID =" & val(TxtSerial1.Text) & ""
                'Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TBL_measureMent2 Where BDet_BD_ID=" & val(Me.TxtSerial1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
              End If
    
            Set RsDevsub = New ADODB.Recordset
            StrSQL = "SELECT  *  from TBL_measureMent2 Where (1 = -1)"
            RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
            
'            Dim Msg As String
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Msg = Msg & " ‰«“·"
'            Else
'                Msg = Msg & "Waiver/Sale of Shares"
'            End If
            Dim mCol As Long
            
            With Me.GridInstallments
               For i = .FixedRows To .Rows - 1
               If (.TextMatrix(i, .ColIndex("Level"))) <> "" Then
                    RsDevsub.AddNew
                    RsDevsub("BDet_BD_ID").value = val(Me.TxtSerial1.text)
                    
                   
                            
                            For mCol = 1 To .Cols - 1
                                If UCase(.ColKey(mCol)) <> UCase("ChkAll") Then
                                    If UCase(.ColKey(mCol)) = UCase(RsDevsub(.ColKey(mCol)).Name) Then
                                        
                                        If .ColDataType(mCol) = flexDTCurrency Or .ColDataType(mCol) = flexDTDecimal Or .ColDataType(mCol) = flexDTDouble Or .ColDataType(mCol) = flexDTSingle Or .ColDataType(mCol) = flexDTLong Or RsDevsub(.ColKey(mCol)).type = adInteger Then
                                                RsDevsub(.ColKey(mCol)) = val(.TextMatrix(i, mCol) & "")
                                          
                                        ElseIf .ColDataType(mCol) = flexDTBoolean Then
                                            RsDevsub(.ColKey(mCol)) = IIf(.TextMatrix(i, mCol) = "", False, .TextMatrix(i, mCol))
                                        ElseIf .ColDataType(mCol) <> flexDTEmpty Then
                                            RsDevsub(.ColKey(mCol)) = Trim(.TextMatrix(i, mCol) & "")
                                        ElseIf .ColDataType(mCol) = adVarWChar Or .ColDataType(mCol) = adEmpty Then
                                            RsDevsub(.ColKey(mCol)) = Trim(.TextMatrix(i, mCol) & "")
                                        End If
                                    End If
                                End If
                            Next
                
              ' RsDevsub!ID = i
               RsDevsub.update
               
              End If
             Next i
            End With
              
              
'''///////////////
  ''///////////////////
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                BtnLast_Click
                
                'FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
               ' FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                
               ' FiLLTXT
                TxtModFlg = "R"
            End If
            BtnLast_Click
            
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
 '  On Error GoTo ErrTrap
    If RsSavRec.EOF Then Exit Sub
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    Me.txtCustomerName = IIf(IsNull(RsSavRec.Fields("CustomerName").value), "", RsSavRec.Fields("CustomerName").value)
    
    Txt_Mobile.text = IIf(IsNull(RsSavRec.Fields("Cust_Mobile").value), "", RsSavRec.Fields("Cust_Mobile").value)
    Txt_City.text = IIf(IsNull(RsSavRec.Fields("Cust_City").value), "False", RsSavRec.Fields("Cust_City"))
    TXT_Time.text = IIf(IsNull(RsSavRec.Fields("Cust_Time").value), "", RsSavRec.Fields("Cust_Time").value)
    TXT_District = IIf(IsNull(RsSavRec.Fields("Cust_District").value), "", RsSavRec.Fields("Cust_District").value)
    DTP_Order.value = IIf(IsNull(RsSavRec.Fields("Date_Order").value), "", RsSavRec.Fields("Date_Order").value)
    DTP_Measure.value = IIf(IsNull(RsSavRec.Fields("Date_measureMent").value), "", RsSavRec.Fields("Date_measureMent").value)
    
'   If Not IsNull(RsSavRec.Fields("level1").value) Then
'   If RsSavRec.Fields("level1").value = True Then Check_Level1.value = vbChecked Else Check_Level1.value = vbUnchecked
'   Else
'   Check_Level1.value = vbUnchecked
'   End If

'   If Not IsNull(RsSavRec.Fields("WCMen1").value) Then
'   If RsSavRec.Fields("WCMen1").value = True Then Check5(0).value = vbChecked Else Check5(0).value = vbUnchecked
'   Else
'   Check5(0).value = vbUnchecked
'   End If
    
    
'   If Not IsNull(RsSavRec.Fields("WCWomen1").value) Then
'   If RsSavRec.Fields("WCWomen1").value = True Then Check5(3).value = vbChecked Else Check5(3).value = vbUnchecked
'   Else
'   Check5(3).value = vbUnchecked
'   End If
     
'   If Not IsNull(RsSavRec.Fields("WCChildren1").value) Then
'   If RsSavRec.Fields("WCChildren1").value = True Then Check5(6).value = vbChecked Else Check5(6).value = vbUnchecked
'   Else
'   Check5(6).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("WCGirls1").value) Then
'   If RsSavRec.Fields("WCGirls1").value = True Then Check5(9).value = vbChecked Else Check5(9).value = vbUnchecked
'   Else
'   Check5(9).value = vbUnchecked
'   End If
''
'    Txt_WCCount1.Text = IIf(IsNull(RsSavRec.Fields("WCCount1").value), "0", RsSavRec.Fields("WCCount1").value)
'    Txt_WCNote1.Text = IIf(IsNull(RsSavRec.Fields("WCNote1").value), "", RsSavRec.Fields("WCNote1").value)
'
'
'   If Not IsNull(RsSavRec.Fields("laundryMen1").value) Then
'   If RsSavRec.Fields("laundryMen1").value = True Then Check5(1).value = vbChecked Else Check5(1).value = vbUnchecked
'   Else
'   Check5(1).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("laundryWomen1").value) Then
'   If RsSavRec.Fields("laundryWomen1").value = True Then Check5(4).value = vbChecked Else Check5(4).value = vbUnchecked
'   Else
'   Check5(4).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("laundryChildren1").value) Then
'   If RsSavRec.Fields("laundryChildren1").value = True Then Check5(7).value = vbChecked Else Check5(7).value = vbUnchecked
'   Else
'   Check5(7).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("laundryGirls1").value) Then
'   If RsSavRec.Fields("laundryGirls1").value = True Then Check5(10).value = vbChecked Else Check5(10).value = vbUnchecked
'   Else
'   Check5(10).value = vbUnchecked
'   End If
'    Txt_laundryCount1.Text = IIf(IsNull(RsSavRec.Fields("laundryCount1").value), "", RsSavRec.Fields("laundryCount1").value)
'    Txt_laundryNote1.Text = IIf(IsNull(RsSavRec.Fields("laundryNote1").value), "", RsSavRec.Fields("laundryNote1").value)
'
'   If Not IsNull(RsSavRec.Fields("laundryareaMen1").value) Then
'   If RsSavRec.Fields("laundryareaMen1").value = True Then Check5(2).value = vbChecked Else Check5(2).value = vbUnchecked
'   Else
'   Check5(2).value = vbUnchecked
'   End If
'
'
'   If Not IsNull(RsSavRec.Fields("laundryareaWomen1").value) Then
'   If RsSavRec.Fields("laundryareaWomen1").value = True Then Check5(5).value = vbChecked Else Check5(5).value = vbUnchecked
'   Else
'   Check5(5).value = vbUnchecked
'   End If
'
'
'   If Not IsNull(RsSavRec.Fields("laundryareaChildren1").value) Then
'   If RsSavRec.Fields("laundryareaChildren1").value = True Then Check5(8).value = vbChecked Else Check5(8).value = vbUnchecked
'   Else
'   Check5(8).value = vbUnchecked
'   End If
'
'
'   If Not IsNull(RsSavRec.Fields("laundryareaGirls1").value) Then
'   If RsSavRec.Fields("laundryareaGirls1").value = True Then Check5(11).value = vbChecked Else Check5(11).value = vbUnchecked
'   Else
'   Check5(11).value = vbUnchecked
'   End If
'    Txt_laundryareaCount1.Text = IIf(IsNull(RsSavRec.Fields("laundryareaCount1").value), "0", RsSavRec.Fields("laundryareaCount1").value)
'    Txt_laundryareaNote1.Text = IIf(IsNull(RsSavRec.Fields("laundryareaNote1").value), "", RsSavRec.Fields("laundryareaNote1").value)
'
'   If Not IsNull(RsSavRec.Fields("MainHall1").value) Then
'   If RsSavRec.Fields("MainHall1").value = True Then Check5(12).value = vbChecked Else Check5(12).value = vbUnchecked
'   Else
'   Check5(12).value = vbUnchecked
'   End If
'    Txt_MainHallCount1.Text = IIf(IsNull(RsSavRec.Fields("MainHallCount1").value), "0", RsSavRec.Fields("MainHallCount1").value)
'    Txt_MainHallNote1.Text = IIf(IsNull(RsSavRec.Fields("MainHallNote1").value), "", RsSavRec.Fields("MainHallNote1").value)
'
'   If Not IsNull(RsSavRec.Fields("kitchen1").value) Then
'   If RsSavRec.Fields("kitchen1").value = True Then Check5(35).value = vbChecked Else Check5(35).value = vbUnchecked
'   Else
'   Check5(35).value = vbUnchecked
'   End If
'    Txt_kitchenCount1.Text = IIf(IsNull(RsSavRec.Fields("kitchenCount1").value), "0", RsSavRec.Fields("kitchenCount1").value)
'    Txt_kitchenNote1.Text = IIf(IsNull(RsSavRec.Fields("kitchenNote1").value), "", RsSavRec.Fields("kitchenNote1").value)
'
'   If Not IsNull(RsSavRec.Fields("BoardMen1").value) Then
'   If RsSavRec.Fields("BoardMen1").value = True Then Check5(13).value = vbChecked Else Check5(13).value = vbUnchecked
'   Else
'   Check5(13).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("BoardWomen1").value) Then
'   If RsSavRec.Fields("BoardWomen1").value = True Then Check5(14).value = vbChecked Else Check5(14).value = vbUnchecked
'   Else
'   Check5(14).value = vbUnchecked
'   End If
'    Txt_BoardCount1.Text = IIf(IsNull(RsSavRec.Fields("BoardCount1").value), "0", RsSavRec.Fields("BoardCount1").value)
'    Txt_BoardNote1.Text = IIf(IsNull(RsSavRec.Fields("BoardNote1").value), "", RsSavRec.Fields("BoardNote1").value)
'
'   If Not IsNull(RsSavRec.Fields("MaklatMen1").value) Then
'   If RsSavRec.Fields("MaklatMen1").value = True Then Check5(15).value = vbChecked Else Check5(15).value = vbUnchecked
'   Else
'   Check5(15).value = vbUnchecked
'   End If
'
'
'   If Not IsNull(RsSavRec.Fields("MaklatWomen1").value) Then
'   If RsSavRec.Fields("MaklatWomen1").value = True Then Check5(16).value = vbChecked Else Check5(16).value = vbUnchecked
'   Else
'   Check5(16).value = vbUnchecked
'   End If
'    Txt_MaklatCount1.Text = IIf(IsNull(RsSavRec.Fields("MaklatCount1").value), "0", RsSavRec.Fields("MaklatCount1").value)
'    Txt_MaklatNote1.Text = IIf(IsNull(RsSavRec.Fields("MaklatNote1").value), "", RsSavRec.Fields("MaklatNote1").value)
'
'
'   If Not IsNull(RsSavRec.Fields("EntranceMen1").value) Then
'   If RsSavRec.Fields("EntranceMen1").value = True Then Check5(17).value = vbChecked Else Check5(17).value = vbUnchecked
'   Else
'   Check5(17).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("EntranceWomen1").value) Then
'   If RsSavRec.Fields("EntranceWomen1").value = True Then Check5(18).value = vbChecked Else Check5(18).value = vbUnchecked
'   Else
'   Check5(18).value = vbUnchecked
'   End If
'    Txt_EntranceCount1.Text = IIf(IsNull(RsSavRec.Fields("EntranceCount1").value), "0", RsSavRec.Fields("EntranceCount1").value)
'    Txt_EntranceNote1.Text = IIf(IsNull(RsSavRec.Fields("EntranceNote1").value), "", RsSavRec.Fields("EntranceNote1").value)
'
'
'   If Not IsNull(RsSavRec.Fields("Dorginside1").value) Then
'   If RsSavRec.Fields("Dorginside1").value = True Then Check5(19).value = vbChecked Else Check5(19).value = vbUnchecked
'   Else
'   Check5(19).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("DorgOutside1").value) Then
'   If RsSavRec.Fields("DorgOutside1").value = True Then Check5(20).value = vbChecked Else Check5(20).value = vbUnchecked
'   Else
'   Check5(20).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("DorgTwacheh1").value) Then
'   If RsSavRec.Fields("DorgTwacheh1").value = True Then Check5(73).value = vbChecked Else Check5(73).value = vbUnchecked
'   Else
'   Check5(73).value = vbUnchecked
'   End If
'    Txt_DorgCount1.Text = IIf(IsNull(RsSavRec.Fields("DorgCount1").value), "0", RsSavRec.Fields("DorgCount1").value)
'    Txt_DorgNote1.Text = IIf(IsNull(RsSavRec.Fields("DorgNote1").value), "", RsSavRec.Fields("DorgNote1").value)
'
'   If Not IsNull(RsSavRec.Fields("ElevatorInside1").value) Then
'   If RsSavRec.Fields("ElevatorInside1").value = True Then Check5(21).value = vbChecked Else Check5(21).value = vbUnchecked
'   Else
'   Check5(21).value = vbUnchecked
'   End If
'
'
'    If Not IsNull(RsSavRec.Fields("ElevatorOutSide1").value) Then
'   If RsSavRec.Fields("ElevatorOutSide1").value = True Then Check5(22).value = vbChecked Else Check5(22).value = vbUnchecked
'   Else
'   Check5(22).value = vbUnchecked
'   End If
'    Txt_ElevatorCount1.Text = IIf(IsNull(RsSavRec.Fields("ElevatorCount1").value), "0", RsSavRec.Fields("ElevatorCount1").value)
'    Txt_ElevatorNote1.Text = IIf(IsNull(RsSavRec.Fields("ElevatorNote1").value), "", RsSavRec.Fields("ElevatorNote1").value)
'
'
'    If Not IsNull(RsSavRec.Fields("HoshInside1").value) Then
'   If RsSavRec.Fields("HoshInside1").value = True Then Check5(23).value = vbChecked Else Check5(23).value = vbUnchecked
'   Else
'   Check5(23).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("HoshOutSide1").value) Then
'   If RsSavRec.Fields("HoshOutSide1").value = True Then Check5(24).value = vbChecked Else Check5(24).value = vbUnchecked
'   Else
'   Check5(24).value = vbUnchecked
'   End If
'    Txt_HoshCount1.Text = IIf(IsNull(RsSavRec.Fields("HoshCount1").value), "0", RsSavRec.Fields("HoshCount1").value)
'    Txt_HoshNote1.Text = IIf(IsNull(RsSavRec.Fields("HoshNote1").value), "", RsSavRec.Fields("HoshNote1").value)
'
'
'    If Not IsNull(RsSavRec.Fields("MainRoom1").value) Then
'   If RsSavRec.Fields("MainRoom1").value = True Then Check5(25).value = vbChecked Else Check5(25).value = vbUnchecked
'   Else
'   Check5(25).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("MRoom1").value) Then
'   If RsSavRec.Fields("MRoom1").value = True Then Check5(26).value = vbChecked Else Check5(26).value = vbUnchecked
'   Else
'   Check5(26).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("LRoom1").value) Then
'   If RsSavRec.Fields("LRoom1").value = True Then Check5(27).value = vbChecked Else Check5(27).value = vbUnchecked
'   Else
'   Check5(27).value = vbUnchecked
'   End If
'    Txt_MainRoomCount1.Text = IIf(IsNull(RsSavRec.Fields("MainRoomCount1").value), "0", RsSavRec.Fields("MainRoomCount1").value)
'    Txt_MainRoomNote1.Text = IIf(IsNull(RsSavRec.Fields("MainRoomNote1").value), "", RsSavRec.Fields("MainRoomNote1").value)
'
'
'    If Not IsNull(RsSavRec.Fields("Na3laNormal1").value) Then
'   If RsSavRec.Fields("Na3laNormal1").value = True Then Check5(28).value = vbChecked Else Check5(28).value = vbUnchecked
'   Else
'   Check5(28).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("Na3laDorg1").value) Then
'   If RsSavRec.Fields("Na3laDorg1").value = True Then Check5(29).value = vbChecked Else Check5(29).value = vbUnchecked
'   Else
'   Check5(29).value = vbUnchecked
'   End If
'    Txt_Na3laCount1.Text = IIf(IsNull(RsSavRec.Fields("Na3laCount1").value), "0", RsSavRec.Fields("Na3laCount1").value)
'    Txt_Na3laNote1.Text = IIf(IsNull(RsSavRec.Fields("Na3laNote1").value), "", RsSavRec.Fields("Na3laNote1").value)
'
'
'    If Not IsNull(RsSavRec.Fields("ClothesInside1").value) Then
'   If RsSavRec.Fields("ClothesInside1").value = True Then Check5(30).value = vbChecked Else Check5(30).value = vbUnchecked
'   Else
'   Check5(30).value = vbUnchecked
'   End If
'
'
'    If Not IsNull(RsSavRec.Fields("ClothesOutInside1").value) Then
'   If RsSavRec.Fields("ClothesOutInside1").value = True Then Check5(31).value = vbChecked Else Check5(31).value = vbUnchecked
'   Else
'   Check5(31).value = vbUnchecked
'   End If
'    Txt_ClothesCount1.Text = IIf(IsNull(RsSavRec.Fields("ClothesCount1").value), "0", RsSavRec.Fields("ClothesCount1").value)
'    Txt_ClothesNote1.Text = IIf(IsNull(RsSavRec.Fields("ClothesNote1").value), "", RsSavRec.Fields("ClothesNote1").value)
'
'
'    If Not IsNull(RsSavRec.Fields("ParkingGround1").value) Then
'   If RsSavRec.Fields("ParkingGround1").value = True Then Check5(32).value = vbChecked Else Check5(32).value = vbUnchecked
'   Else
'   Check5(32).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("ParkingGdar1").value) Then
'   If RsSavRec.Fields("ParkingGdar1").value = True Then Check5(33).value = vbChecked Else Check5(33).value = vbUnchecked
'   Else
'   Check5(33).value = vbUnchecked
'   End If
'    Txt_ParkingCount1.Text = IIf(IsNull(RsSavRec.Fields("ParkingCount1").value), "0", RsSavRec.Fields("ParkingCount1").value)
'    Txt_ParkingNote1.Text = IIf(IsNull(RsSavRec.Fields("ParkingNote1").value), "", RsSavRec.Fields("ParkingNote1").value)
'
'    If Not IsNull(RsSavRec.Fields("Office1").value) Then
'   If RsSavRec.Fields("Office1").value = True Then Check5(34).value = vbChecked Else Check5(34).value = vbUnchecked
'   Else
'   Check5(34).value = vbUnchecked
'   End If
'    Txt_OfficeCount1.Text = IIf(IsNull(RsSavRec.Fields("OfficeCount1").value), "0", RsSavRec.Fields("OfficeCount1").value)
'    Txt_OfficeNote1.Text = IIf(IsNull(RsSavRec.Fields("OfficeNote1").value), "", RsSavRec.Fields("OfficeNote1").value)
'
'    TXT_Mobile2.Text = IIf(IsNull(RsSavRec.Fields("Cust_Mobile2").value), "", RsSavRec.Fields("Cust_Mobile2").value)
'    TXT_City2.Text = IIf(IsNull(RsSavRec.Fields("Cust_City2").value), "", RsSavRec.Fields("Cust_City2").value)
'    TXT_Time2.Text = IIf(IsNull(RsSavRec.Fields("Cust_Time2").value), "", RsSavRec.Fields("Cust_Time2").value)
'    TXT_District2.Text = IIf(IsNull(RsSavRec.Fields("Cust_District2").value), "", RsSavRec.Fields("Cust_District2").value)
'    DTP_Order2.value = IIf(IsNull(RsSavRec.Fields("Date_Order2").value), "", RsSavRec.Fields("Date_Order2").value)
'    DTP_Measure2.value = IIf(IsNull(RsSavRec.Fields("Date_measureMent2").value), "", RsSavRec.Fields("Date_measureMent2").value)
'
'
'
'    If Not IsNull(RsSavRec.Fields("level2").value) Then
'   If RsSavRec.Fields("level2").value = True Then Check_Level2.value = vbChecked Else Check_Level2.value = vbUnchecked
'   Else
'  Check_Level2.value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("WCMen2").value) Then
'   If RsSavRec.Fields("WCMen2").value = True Then Check5(36).value = vbChecked Else Check5(36).value = vbUnchecked
'   Else
'   Check5(36).value = vbUnchecked
'   End If
'
'
'    If Not IsNull(RsSavRec.Fields("WCWomen2").value) Then
'   If RsSavRec.Fields("WCWomen2").value = True Then Check5(39).value = vbChecked Else Check5(39).value = vbUnchecked
'   Else
'   Check5(39).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("WCChildren2").value) Then
'   If RsSavRec.Fields("WCChildren2").value = True Then Check5(42).value = vbChecked Else Check5(42).value = vbUnchecked
'   Else
'   Check5(42).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("WCGirls2").value) Then
'   If RsSavRec.Fields("WCGirls2").value = True Then Check5(45).value = vbChecked Else Check5(45).value = vbUnchecked
'   Else
'   Check5(45).value = vbUnchecked
'   End If
'    Txt_WCCount2.Text = IIf(IsNull(RsSavRec.Fields("WCCount2").value), "0", RsSavRec.Fields("WCCount2").value)
'    Txt_WCNote2.Text = IIf(IsNull(RsSavRec.Fields("WCNote2").value), "", RsSavRec.Fields("WCNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("laundryMen2").value) Then
'   If RsSavRec.Fields("laundryMen2").value = True Then Check5(37).value = vbChecked Else Check5(37).value = vbUnchecked
'   Else
'   Check5(37).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("laundryWomen2").value) Then
'   If RsSavRec.Fields("laundryWomen2").value = True Then Check5(40).value = vbChecked Else Check5(40).value = vbUnchecked
'   Else
'   Check5(40).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("laundryChildren2").value) Then
'   If RsSavRec.Fields("laundryChildren2").value = True Then Check5(43).value = vbChecked Else Check5(43).value = vbUnchecked
'   Else
'   Check5(43).value = vbUnchecked
'   End If
'
'
'    If Not IsNull(RsSavRec.Fields("laundryGirls2").value) Then
'   If RsSavRec.Fields("laundryGirls2").value = True Then Check5(46).value = vbChecked Else Check5(46).value = vbUnchecked
'   Else
'   Check5(46).value = vbUnchecked
'   End If
'    Txt_laundryCount2.Text = IIf(IsNull(RsSavRec.Fields("laundryCount2").value), "0", RsSavRec.Fields("laundryCount2").value)
'    Txt_laundryNote2.Text = IIf(IsNull(RsSavRec.Fields("laundryNote2").value), "", RsSavRec.Fields("laundryNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("laundryareaMen2").value) Then
'   If RsSavRec.Fields("laundryareaMen2").value = True Then Check5(38).value = vbChecked Else Check5(38).value = vbUnchecked
'   Else
'   Check5(38).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("laundryareaWomen2").value) Then
'   If RsSavRec.Fields("laundryareaWomen2").value = True Then Check5(41).value = vbChecked Else Check5(41).value = vbUnchecked
'   Else
'   Check5(41).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("laundryareaChildren2").value) Then
'   If RsSavRec.Fields("laundryareaChildren2").value = True Then Check5(44).value = vbChecked Else Check5(44).value = vbUnchecked
'   Else
'   Check5(44).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("laundryareaGirls2").value) Then
'   If RsSavRec.Fields("laundryareaGirls2").value = True Then Check5(47).value = vbChecked Else Check5(47).value = vbUnchecked
'   Else
'   Check5(47).value = vbUnchecked
'   End If
'    Txt_laundryareaCount2.Text = IIf(IsNull(RsSavRec.Fields("laundryareaCount2").value), "0", RsSavRec.Fields("laundryareaCount2").value)
'    Txt_laundryareaNote2.Text = IIf(IsNull(RsSavRec.Fields("laundryareaNote2").value), "", RsSavRec.Fields("laundryareaNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("MainHall2").value) Then
'   If RsSavRec.Fields("MainHall2").value = True Then Check5(48).value = vbChecked Else Check5(48).value = vbUnchecked
'   Else
'   Check5(48).value = vbUnchecked
'   End If
'    Txt_MainHallCount2.Text = IIf(IsNull(RsSavRec.Fields("MainHallCount2").value), "0", RsSavRec.Fields("MainHallCount2").value)
'    Txt_MainHallNote2.Text = IIf(IsNull(RsSavRec.Fields("MainHallNote2").value), "", RsSavRec.Fields("MainHallNote2").value)
'
'
'   If Not IsNull(RsSavRec.Fields("kitchen2").value) Then
'   If RsSavRec.Fields("kitchen2").value = True Then Check5(49).value = vbChecked Else Check5(49).value = vbUnchecked
'   Else
'   Check5(49).value = vbUnchecked
'   End If
'    Txt_kitchenCount2.Text = IIf(IsNull(RsSavRec.Fields("kitchenCount2").value), "0", RsSavRec.Fields("kitchenCount2").value)
'    Txt_kitchenNote2 = IIf(IsNull(RsSavRec.Fields("kitchenNote2").value), "", RsSavRec.Fields("kitchenNote2").value)
'
'   If Not IsNull(RsSavRec.Fields("BoardMen2").value) Then
'   If RsSavRec.Fields("BoardMen2").value = True Then Check5(50).value = vbChecked Else Check5(50).value = vbUnchecked
'   Else
'   Check5(50).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("BoardWomen2").value) Then
'   If RsSavRec.Fields("BoardWomen2").value = True Then Check5(51).value = vbChecked Else Check5(51).value = vbUnchecked
'   Else
'   Check5(51).value = vbUnchecked
'   End If
'    Txt_BoardCount2.Text = IIf(IsNull(RsSavRec.Fields("BoardCount2").value), "0", RsSavRec.Fields("BoardCount2").value)
'    Txt_BoardNote2.Text = IIf(IsNull(RsSavRec.Fields("BoardNote2").value), "", RsSavRec.Fields("BoardNote2").value)
'
'    If Not IsNull(RsSavRec.Fields("MaklatMen2").value) Then
'   If RsSavRec.Fields("MaklatMen2").value = True Then Check5(52).value = vbChecked Else Check5(52).value = vbUnchecked
'   Else
'   Check5(52).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("MaklatWomen2").value) Then
'   If RsSavRec.Fields("MaklatWomen2").value = True Then Check5(53).value = vbChecked Else Check5(53).value = vbUnchecked
'   Else
'   Check5(53).value = vbUnchecked
'   End If
'    Txt_MaklatCount2.Text = IIf(IsNull(RsSavRec.Fields("MaklatCount2").value), "0", RsSavRec.Fields("MaklatCount2").value)
'    Txt_MaklatNote2.Text = IIf(IsNull(RsSavRec.Fields("MaklatNote2").value), "", RsSavRec.Fields("MaklatNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("EntranceMen2").value) Then
'   If RsSavRec.Fields("EntranceMen2").value = True Then Check5(54).value = vbChecked Else Check5(54).value = vbUnchecked
'   Else
'   Check5(54).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("EntranceWomen2").value) Then
'   If RsSavRec.Fields("EntranceWomen2").value = True Then Check5(55).value = vbChecked Else Check5(55).value = vbUnchecked
'   Else
'   Check5(55).value = vbUnchecked
'   End If
'    Txt_EntranceCount2.Text = IIf(IsNull(RsSavRec.Fields("EntranceCount2").value), "0", RsSavRec.Fields("EntranceCount2").value)
'    Txt_EntranceNote2.Text = IIf(IsNull(RsSavRec.Fields("EntranceNote2").value), "", RsSavRec.Fields("EntranceNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("Dorginside2").value) Then
'   If RsSavRec.Fields("Dorginside2").value = True Then Check5(56).value = vbChecked Else Check5(56).value = vbUnchecked
'   Else
'   Check5(56).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("DorgOutside2").value) Then
'   If RsSavRec.Fields("DorgOutside2").value = True Then Check5(57).value = vbChecked Else Check5(57).value = vbUnchecked
'   Else
'   Check5(57).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("DorgTwacheh2").value) Then
'   If RsSavRec.Fields("DorgTwacheh2").value = True Then Check5(58).value = vbChecked Else Check5(58).value = vbUnchecked
'   Else
'   Check5(58).value = vbUnchecked
'   End If
'    Txt_DorgCount2.Text = IIf(IsNull(RsSavRec.Fields("DorgCount2").value), "0", RsSavRec.Fields("DorgCount2").value)
'    Txt_DorgNote2.Text = IIf(IsNull(RsSavRec.Fields("DorgNote2").value), "", RsSavRec.Fields("DorgNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("ElevatorInside2").value) Then
'   If RsSavRec.Fields("ElevatorInside2").value = True Then Check5(59).value = vbChecked Else Check5(59).value = vbUnchecked
'   Else
'   Check5(59).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("ElevatorOutSide2").value) Then
'   If RsSavRec.Fields("ElevatorOutSide2").value = True Then Check5(60).value = vbChecked Else Check5(60).value = vbUnchecked
'   Else
'   Check5(60).value = vbUnchecked
'   End If
'    Txt_ElevatorCount2.Text = IIf(IsNull(RsSavRec.Fields("ElevatorCount2").value), "0", RsSavRec.Fields("ElevatorCount2").value)
'    Txt_ElevatorNote2.Text = IIf(IsNull(RsSavRec.Fields("ElevatorNote2").value), "", RsSavRec.Fields("ElevatorNote2").value)
'
'    If Not IsNull(RsSavRec.Fields("HoshInside2").value) Then
'   If RsSavRec.Fields("HoshInside2").value = True Then Check5(61).value = vbChecked Else Check5(61).value = vbUnchecked
'   Else
'   Check5(61).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("HoshOutSide2").value) Then
'   If RsSavRec.Fields("HoshOutSide2").value = True Then Check5(62).value = vbChecked Else Check5(62).value = vbUnchecked
'   Else
'   Check5(62).value = vbUnchecked
'   End If
'    Txt_HoshCount2.Text = IIf(IsNull(RsSavRec.Fields("HoshCount2").value), "0", RsSavRec.Fields("HoshCount2").value)
'    Txt_HoshNote2.Text = IIf(IsNull(RsSavRec.Fields("HoshNote2").value), "", RsSavRec.Fields("HoshNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("MainRoom2").value) Then
'   If RsSavRec.Fields("MainRoom2").value = True Then Check5(63).value = vbChecked Else Check5(63).value = vbUnchecked
'   Else
'   Check5(63).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("MRoom2").value) Then
'   If RsSavRec.Fields("MRoom2").value = True Then Check5(64).value = vbChecked Else Check5(64).value = vbUnchecked
'   Else
'   Check5(64).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("LRoom2").value) Then
'   If RsSavRec.Fields("LRoom2").value = True Then Check5(65).value = vbChecked Else Check5(65).value = vbUnchecked
'   Else
'   Check5(65).value = vbUnchecked
'   End If
'    Txt_MainRoomCount2.Text = IIf(IsNull(RsSavRec.Fields("MainRoomCount2").value), "0", RsSavRec.Fields("MainRoomCount2").value)
'    Txt_MainRoomNote2.Text = IIf(IsNull(RsSavRec.Fields("MainRoomNote2").value), "", RsSavRec.Fields("MainRoomNote2").value)
'
'    If Not IsNull(RsSavRec.Fields("Na3laNormal2").value) Then
'   If RsSavRec.Fields("Na3laNormal2").value = True Then Check5(66).value = vbChecked Else Check5(66).value = vbUnchecked
'   Else
'   Check5(66).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("Na3laDorg2").value) Then
'   If RsSavRec.Fields("Na3laDorg2").value = True Then Check5(67).value = vbChecked Else Check5(67).value = vbUnchecked
'   Else
'   Check5(67).value = vbUnchecked
'   End If
'    Txt_Na3laCount2.Text = IIf(IsNull(RsSavRec.Fields("Na3laCount2").value), "0", RsSavRec.Fields("Na3laCount2").value)
'    Txt_Na3laNote2.Text = IIf(IsNull(RsSavRec.Fields("Na3laNote2").value), "", RsSavRec.Fields("Na3laNote2").value)
'
'    If Not IsNull(RsSavRec.Fields("ClothesInside2").value) Then
'   If RsSavRec.Fields("ClothesInside2").value = True Then Check5(68).value = vbChecked Else Check5(68).value = vbUnchecked
'   Else
'   Check5(68).value = vbUnchecked
'   End If
'
'    If Not IsNull(RsSavRec.Fields("ClothesOutInside2").value) Then
'   If RsSavRec.Fields("ClothesOutInside2").value = True Then Check5(69).value = vbChecked Else Check5(69).value = vbUnchecked
'   Else
'   Check5(69).value = vbUnchecked
'   End If
'    Txt_ClothesCount2.Text = IIf(IsNull(RsSavRec.Fields("ClothesCount2").value), "0", RsSavRec.Fields("ClothesCount2").value)
'    Txt_ClothesNote2.Text = IIf(IsNull(RsSavRec.Fields("ClothesNote2").value), "", RsSavRec.Fields("ClothesNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("ParkingGround2").value) Then
'   If RsSavRec.Fields("ParkingGround2").value = True Then Check5(70).value = vbChecked Else Check5(70).value = vbUnchecked
'   Else
'   Check5(70).value = vbUnchecked
'   End If
'
'   If Not IsNull(RsSavRec.Fields("ParkingGdar2").value) Then
'   If RsSavRec.Fields("ParkingGdar2").value = True Then Check5(71).value = vbChecked Else Check5(71).value = vbUnchecked
'   Else
'   Check5(71).value = vbUnchecked
'   End If
'    Txt_ParkingCount2.Text = IIf(IsNull(RsSavRec.Fields("ParkingCount2").value), "0", RsSavRec.Fields("ParkingCount2").value)
'    Txt_ParkingNote2.Text = IIf(IsNull(RsSavRec.Fields("ParkingNote2").value), "", RsSavRec.Fields("ParkingNote2").value)
'
'
'    If Not IsNull(RsSavRec.Fields("Office2").value) Then
'   If RsSavRec.Fields("Office2").value = True Then Check5(72).value = vbChecked Else Check5(72).value = vbUnchecked
'   Else
'   Check5(72).value = vbUnchecked
'   End If
'    Txt_OfficeCount2.Text = IIf(IsNull(RsSavRec.Fields("OfficeCount2").value), "0", RsSavRec.Fields("OfficeCount2").value)
'    Txt_OfficeNote2.Text = IIf(IsNull(RsSavRec.Fields("OfficeNote2").value), "", RsSavRec.Fields("OfficeNote2").value)
'    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    GridInstallments.Rows = 2
    GridInstallments.Rows = 3
    Dim s As String
    Dim Rs1 As New ADODB.Recordset
    s = " Select *,chkAll=0 from TBL_measureMent2 Where BDet_BD_ID = " & val(TxtSerial1) & " Order By [Level]"
    Rs1.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       
    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
    End If
    GridInstallments.Rows = GridInstallments.Rows + Rs1.RecordCount
    Dim i As Integer, mCol As Long
    With Me.GridInstallments
        For i = .FixedRows To Rs1.RecordCount + 1
            .TextMatrix(i, .ColIndex("Ser")) = i - 1
            For mCol = 1 To .Cols - 1
                If UCase(.ColKey(mCol)) = UCase(Rs1(.ColKey(mCol)).Name) Then
                    .TextMatrix(i, mCol) = Rs1(.ColKey(mCol)) & ""
                End If
            Next
                

            Rs1.MoveNext
        Next i
        End With
'FullGridData
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    '---------------------- check if data Vaclete -----------------------
      If txtCustomerName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ· «·«”„", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Insert Customer Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            txtCustomerName.SetFocus
            Exit Sub
     End If

If Txt_Mobile.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«·  —ﬁ„ «·Â« › "
Else
MsgBox "Please Enter Mobile No..."
End If
Txt_Mobile.SetFocus
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
    StrRecID = new_id("TBL_measureMent", "ID", "")
    Me.TxtSerial1.text = StrRecID

    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)

    FiLLRec
    
ErrTrap:
End Sub
  Sub FullGridData()
     End Sub


Function CountInBooking() As Double
 End Function
Sub AddUnit(Optional Row As Long)
 End Sub
Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 End Sub



Function GetNoBooking(Optional PartID As Double) As Double
 End Function
Sub FillGrid()
 
End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
 End Sub
Function NoBookingUnitInGrid(Optional PartID As Double) As Double
 End Function

Function ChekGrid(Optional PartID As Double) As Boolean
 End Function
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 End Sub

Private Sub ISButton3_Click()
 End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If Row = GridInstallments.Rows - 1 Then
             GridInstallments.Rows = GridInstallments.Rows + 1
        End If
        
        If GridInstallments.ColIndex("chkAll") = Col Then
         Dim mCol  As Long, mChkAll As Boolean
          mChkAll = GridInstallments.TextMatrix(Row, Col)
          For mCol = 1 To GridInstallments.Cols - 1
                
                If GridInstallments.ColDataType(mCol) = flexDTBoolean Then
                    GridInstallments.TextMatrix(Row, mCol) = mChkAll
                End If
            Next
        ElseIf GridInstallments.ColIndex("LevelName") = Col Then
            GridInstallments.TextMatrix(Row, GridInstallments.ColIndex("Level")) = Row - 1

        End If
ReLineGrid
End Sub

Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
ReLineGrid
End Sub

Private Sub ISButton5_Click()
    print_report
End Sub

  
Private Sub ISButton9_Click()

            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.text, "2204201801"
ErrTrap:


End Sub

Private Sub ISButton8_Click()

FrmProjectSearch.C1Tab1.CurrTab = 1
FrmProjectSearch.Caption = "»ÕÀ —›⁄ «·ﬁÌ«”"

FrmProjectSearch.show vbModal
End Sub

Private Sub txtRowNo_Validate(Cancel As Boolean)
GridInstallments.Rows = GridInstallments.FixedRows
GridInstallments.Rows = GridInstallments.FixedRows + val(txtRowNo)
Dim i As Long

For i = GridInstallments.FixedRows To GridInstallments.Rows - 1
    GridInstallments.TextMatrix(i, GridInstallments.ColIndex("LevelName")) = "«·œÊ— " & Write_Qast(CInt(i - 1))
    GridInstallments.TextMatrix(i, GridInstallments.ColIndex("Level")) = i - 1
Next
End Sub

Private Sub TxtName_Change(Index As Integer)
If TxtName(0).text = "" Then Exit Sub
FindRec val(TxtName(0).text)
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
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
      Dim StrSQL As String
                RsSavRec.find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                                          RsSavRec.delete

     
                 StrSQL = "Delete From TBL_measureMent2 Where BDet_BD_ID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          
                                                        
                    GridInstallments.Clear flexClearScrollable, flexClearEverything
                    GridInstallments.Rows = 2
                    
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
               LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
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
           'Cn.Errors.Clear
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
    If TxtModFlg.text = "N" Then
 
        'Command2.Enabled = True
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
   ' Command2.Enabled = False
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
  ' Command2.Enabled = True
  '  XPDtbTrans.Enabled = True
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
    
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount
    
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
'    On Error GoTo ErrTrap
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
        LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount
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
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
           ' Grid.Rows = Grid.Rows + 1
             'VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        Me.DCboUserName.BoundText = user_id
      '  Me.dcBranch.SetFocus
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
    GridInstallments.Tag = 1
    clear_all Me
   TXT_Time.text = Format(Time, "hh : mm AM/PM")
   TxtModFlg.text = "N"
    Me.DCboUserName.BoundText = user_id
    GridInstallments.Rows = 2
    GridInstallments.Rows = 3
    txtRowNo = 1
    txtRowNo_Validate True
    
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
        LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount
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
        LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount
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
 
    Dim XPic As IPictureDisp
  
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Me.Caption = "Request Measurement "
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
         
Label2(3).Caption = "Requerst No"
Label2(2) = "Customer Name"

Label4.Caption = "Mobile"
Label6.Caption = "City"
Label7.Caption = "Time"
Label8.Caption = "Dist"
Label9.Caption = "Order Date"
Label5.Caption = "Measure Date"


With GridInstallments
.TextMatrix(0, .ColIndex("chkAll")) = "All"
.TextMatrix(1, .ColIndex("chkAll")) = "All"

.TextMatrix(0, .ColIndex("Level")) = "Level"
.TextMatrix(1, .ColIndex("Level")) = "Level"

.TextMatrix(0, .ColIndex("WCMen")) = "WC"
.TextMatrix(1, .ColIndex("WCMen")) = "Men"

.TextMatrix(0, .ColIndex("WCWomen")) = "WC"
.TextMatrix(1, .ColIndex("WCWomen")) = "Women"

.TextMatrix(0, .ColIndex("WCChildren")) = "WC"
.TextMatrix(1, .ColIndex("WCChildren")) = "Children"

.TextMatrix(0, .ColIndex("WCGirls")) = "WC"
.TextMatrix(1, .ColIndex("WCGirls")) = "Girls"

.TextMatrix(0, .ColIndex("WCCount")) = "Count"
.TextMatrix(1, .ColIndex("WCCount")) = "Count"

.TextMatrix(0, .ColIndex("WCNote")) = "Note"
.TextMatrix(1, .ColIndex("WCNote")) = "Note"

.TextMatrix(0, .ColIndex("laundryMen")) = "laundry"
.TextMatrix(1, .ColIndex("laundryMen")) = "Men"

.TextMatrix(0, .ColIndex("laundryWomen")) = "laundry"
.TextMatrix(1, .ColIndex("laundryWomen")) = "Women"

.TextMatrix(0, .ColIndex("laundryChildren")) = "laundry"
.TextMatrix(1, .ColIndex("laundryChildren")) = "Children"

.TextMatrix(0, .ColIndex("laundryGirls")) = "laundry"
.TextMatrix(1, .ColIndex("laundryGirls")) = "Girls"

.TextMatrix(0, .ColIndex("laundryCount")) = "Count"
.TextMatrix(1, .ColIndex("laundryCount")) = "Count"

.TextMatrix(0, .ColIndex("laundryNote")) = "Note"
.TextMatrix(1, .ColIndex("laundryNote")) = "Note"

.TextMatrix(0, .ColIndex("laundryareaMen")) = "laundry area"
.TextMatrix(1, .ColIndex("laundryareaMen")) = "Men"

.TextMatrix(0, .ColIndex("laundryareaWomen")) = "laundry area"
.TextMatrix(1, .ColIndex("laundryareaWomen")) = "Women"

.TextMatrix(0, .ColIndex("laundryareaChildren")) = "laundry area"
.TextMatrix(1, .ColIndex("laundryareaChildren")) = "Children"

.TextMatrix(0, .ColIndex("laundryareaGirls")) = "laundry area"
.TextMatrix(1, .ColIndex("laundryareaGirls")) = "Girls"

.TextMatrix(0, .ColIndex("laundryareaCount")) = "Coun"
.TextMatrix(1, .ColIndex("laundryareaCount")) = "Coun"

.TextMatrix(0, .ColIndex("laundryareaNote")) = "Note"
.TextMatrix(1, .ColIndex("laundryareaNote")) = "Note"

.TextMatrix(0, .ColIndex("MainHall")) = "MainHall"
.TextMatrix(1, .ColIndex("MainHall")) = "MainHall"

.TextMatrix(0, .ColIndex("MainHallCount")) = "Count"
.TextMatrix(1, .ColIndex("MainHallCount")) = "Count"

.TextMatrix(0, .ColIndex("MainHallNote")) = "Note"
.TextMatrix(1, .ColIndex("MainHallNote")) = "Note"

.TextMatrix(0, .ColIndex("kitchen")) = "kitchen"
.TextMatrix(1, .ColIndex("kitchen")) = "kitchen"

.TextMatrix(0, .ColIndex("kitchenCount")) = "Count"
.TextMatrix(1, .ColIndex("kitchenCount")) = "Count"

.TextMatrix(0, .ColIndex("kitchenNote")) = "Note"
.TextMatrix(1, .ColIndex("kitchenNote")) = "Note"

.TextMatrix(0, .ColIndex("BoardMen")) = "Board"
.TextMatrix(1, .ColIndex("BoardMen")) = "Men"

.TextMatrix(0, .ColIndex("BoardWomen")) = "Board"
.TextMatrix(1, .ColIndex("BoardWomen")) = "Women"


.TextMatrix(0, .ColIndex("BoardCount")) = "Count"
.TextMatrix(1, .ColIndex("BoardCount")) = "Count"


.TextMatrix(0, .ColIndex("BoardNote")) = "Note"
.TextMatrix(1, .ColIndex("BoardNote")) = "Note"


.TextMatrix(0, .ColIndex("MaklatMen")) = "Maklat"
.TextMatrix(1, .ColIndex("MaklatMen")) = "Men"


.TextMatrix(0, .ColIndex("MaklatWomen")) = "Maklat"
.TextMatrix(1, .ColIndex("MaklatWomen")) = "Women"


.TextMatrix(0, .ColIndex("MaklatCount")) = "Count"
.TextMatrix(1, .ColIndex("MaklatCount")) = "Count"


.TextMatrix(0, .ColIndex("MaklatNote")) = "Note"
.TextMatrix(1, .ColIndex("MaklatNote")) = "Note"


.TextMatrix(0, .ColIndex("EntranceMen")) = "Entrance"
.TextMatrix(1, .ColIndex("EntranceMen")) = "Men"


.TextMatrix(0, .ColIndex("EntranceWomen")) = "Entrance"
.TextMatrix(1, .ColIndex("EntranceWomen")) = "Women"


.TextMatrix(0, .ColIndex("EntranceCount")) = "Count"
.TextMatrix(1, .ColIndex("EntranceCount")) = "Count"

.TextMatrix(0, .ColIndex("EntranceNote")) = "Note"
.TextMatrix(1, .ColIndex("EntranceNote")) = "Note"

.TextMatrix(0, .ColIndex("Dorginside")) = "Dorg"
.TextMatrix(1, .ColIndex("Dorginside")) = "inside"

.TextMatrix(0, .ColIndex("DorgOutside")) = "Dorg"
.TextMatrix(1, .ColIndex("DorgOutside")) = "Outside"

.TextMatrix(0, .ColIndex("DorgTwacheh")) = "Dorg"
.TextMatrix(1, .ColIndex("DorgTwacheh")) = "Twacheh"

.TextMatrix(0, .ColIndex("DorgCount")) = "Count"
.TextMatrix(1, .ColIndex("DorgCount")) = "Count"

.TextMatrix(0, .ColIndex("DorgNote")) = "Note"
.TextMatrix(1, .ColIndex("DorgNote")) = "Note"

.TextMatrix(0, .ColIndex("ElevatorInside")) = "Elevator"
.TextMatrix(1, .ColIndex("ElevatorInside")) = "Inside"

.TextMatrix(0, .ColIndex("ElevatorOutSide")) = "Elevator"
.TextMatrix(1, .ColIndex("ElevatorOutSide")) = "Outside"

.TextMatrix(0, .ColIndex("ElevatorCount")) = "Count"
.TextMatrix(1, .ColIndex("ElevatorCount")) = "Count"

.TextMatrix(0, .ColIndex("ElevatorNote")) = "Note"
.TextMatrix(1, .ColIndex("ElevatorNote")) = "Note"

.TextMatrix(0, .ColIndex("HoshInside")) = "Hosh"
.TextMatrix(1, .ColIndex("HoshInside")) = "Inside"

.TextMatrix(0, .ColIndex("HoshOutSide")) = "Hosh"
.TextMatrix(1, .ColIndex("HoshOutSide")) = "Outside"

.TextMatrix(0, .ColIndex("HoshCount")) = "Count"
.TextMatrix(1, .ColIndex("HoshCount")) = "Count"

.TextMatrix(0, .ColIndex("HoshNote")) = "Note"
.TextMatrix(1, .ColIndex("HoshNote")) = "Note"


.TextMatrix(0, .ColIndex("MainRoom")) = "BedRoom"
.TextMatrix(1, .ColIndex("MainRoom")) = "Main Room"


.TextMatrix(0, .ColIndex("MRoom")) = "BedRoom"
.TextMatrix(1, .ColIndex("MRoom")) = "Children"


.TextMatrix(0, .ColIndex("LRoom")) = "BedRoom"
.TextMatrix(1, .ColIndex("LRoom")) = "Girls"


.TextMatrix(0, .ColIndex("MainRoomCount")) = "Count"
.TextMatrix(1, .ColIndex("MainRoomCount")) = "Count"


.TextMatrix(0, .ColIndex("MainRoomNote")) = "Note"
.TextMatrix(1, .ColIndex("MainRoomNote")) = "Note"


.TextMatrix(0, .ColIndex("Na3laNormal")) = "Na3la"
.TextMatrix(1, .ColIndex("Na3laNormal")) = "Na3laNormal"


.TextMatrix(0, .ColIndex("Na3laDorg")) = "Na3la"
.TextMatrix(1, .ColIndex("Na3laDorg")) = "Na3laDorg"


.TextMatrix(0, .ColIndex("Na3laCount")) = "Count"
.TextMatrix(1, .ColIndex("Na3laCount")) = "Count"

.TextMatrix(0, .ColIndex("Na3laNote")) = "Note"
.TextMatrix(1, .ColIndex("Na3laNote")) = "Note"

.TextMatrix(0, .ColIndex("ClothesInside")) = "Clothes"
.TextMatrix(1, .ColIndex("ClothesInside")) = "Inside"

.TextMatrix(0, .ColIndex("ClothesOutInside")) = "Clothes"
.TextMatrix(1, .ColIndex("ClothesOutInside")) = "Outside"

.TextMatrix(0, .ColIndex("ClothesCount")) = "Count"
.TextMatrix(1, .ColIndex("ClothesCount")) = "Count"

.TextMatrix(0, .ColIndex("ClothesNote")) = "Note"
.TextMatrix(1, .ColIndex("ClothesNote")) = "Note"

.TextMatrix(0, .ColIndex("ParkingGround")) = "Parking Ground"
.TextMatrix(1, .ColIndex("ParkingGround")) = "Inside"

.TextMatrix(0, .ColIndex("ParkingGdar")) = "Parking Ground"
.TextMatrix(1, .ColIndex("ParkingGdar")) = "Outside"

.TextMatrix(0, .ColIndex("ParkingCount")) = "Count"
.TextMatrix(1, .ColIndex("ParkingCount")) = "Count"


.TextMatrix(0, .ColIndex("ParkingNote")) = "Note"
.TextMatrix(1, .ColIndex("ParkingNote")) = "Note"


.TextMatrix(0, .ColIndex("Office")) = "Office"
.TextMatrix(1, .ColIndex("Office")) = "Office"

.TextMatrix(0, .ColIndex("OfficeCount")) = "Count"
.TextMatrix(1, .ColIndex("OfficeCount")) = "Count"

.TextMatrix(0, .ColIndex("OfficeNote")) = "Note"
.TextMatrix(1, .ColIndex("OfficeNote")) = "Note"

End With

   
    Me.Caption = ScreenNameEnglish

End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblBookingBondsInvs"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end
Private Sub TxtSquareCode_Change()

End Sub

Private Sub TxtSquareCode_KeyPress(KeyAscii As Integer)

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
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
 

     
     
        MySQL = " SELECT TBL_measureMent.ID,IsNull(T2.Level,0) as level1,T2.LevelName ,"
       MySQL = MySQL & "  TBL_measureMent.CustomerName as CusName,TBL_measureMent.Cust_Mobile,TBL_measureMent.Cust_City,"
       MySQL = MySQL & " TBL_measureMent.Cust_Time,TBL_measureMent.Cust_District,TBL_measureMent.Date_Order,TBL_measureMent.Date_measureMent,T2.WCMen WCMen1,"
       MySQL = MySQL & " T2.WCWomen WCWomen1,T2.WCChildren WCChildren1,T2.WCGirls WCGirls1,"
       MySQL = MySQL & " T2.WCCount WCCount1,T2.WCNote WCNote1,T2.laundryMen laundryMen1,T2.laundryWomen laundryWomen1,"
       MySQL = MySQL & " T2.laundryChildren laundryChildren1,T2.laundryGirls laundryGirls1,T2.laundryCount laundryCount1,"
       MySQL = MySQL & " T2.laundryNote laundryNote1,T2.laundryareaMen laundryareaMen1,T2.laundryareaWomen laundryareaWomen1,"
       MySQL = MySQL & " T2.laundryareaChildren laundryareaChildren1,T2.laundryareaGirls laundryareaGirls1,"
       MySQL = MySQL & " T2.laundryareaCount laundryareaCount1,T2.laundryareaNote laundryareaNote1,T2.MainHall MainHall1,"
       MySQL = MySQL & " T2.MainHallCount MainHallCount1,T2.MainHallNote MainHallNote1,T2.kitchen kitchen1,T2.kitchenCount kitchenCount1,"
       MySQL = MySQL & " T2.kitchenNote kitchenNote1,T2.BoardMen BoardMen1,T2.BoardWomen BoardWomen1,T2.BoardCount BoardCount1,"
       MySQL = MySQL & " T2.BoardNote BoardNote1,T2.MaklatMen MaklatMen1,T2.MaklatWomen MaklatWomen1,T2.MaklatCount MaklatCount1,"
       MySQL = MySQL & " T2.MaklatNote MaklatNote1,T2.EntranceMen EntranceMen1,T2.EntranceWomen EntranceWomen1,"
       MySQL = MySQL & " T2.EntranceCount EntranceCount1,T2.EntranceNote EntranceNote1,T2.Dorginside Dorginside1,"
       MySQL = MySQL & " T2.DorgOutside DorgOutside1,T2.DorgTwacheh DorgTwacheh1,T2.DorgCount DorgCount1,"
       MySQL = MySQL & " T2.DorgNote DorgNote1,T2.ElevatorInside ElevatorInside1,T2.ElevatorOutSide ElevatorOutSide1,"
       MySQL = MySQL & " T2.ElevatorCount ElevatorCount1,T2.ElevatorNote ElevatorNote1,T2.HoshInside HoshInside1,"
       MySQL = MySQL & " T2.HoshOutSide HoshOutSide1,T2.HoshCount HoshCount1,T2.HoshNote HoshNote1,"
       MySQL = MySQL & " T2.MainRoom MainRoom1,T2.MRoom MRoom1,T2.LRoom LRoom1,T2.MainRoomCount MainRoomCount1,"
       MySQL = MySQL & " T2.MainRoomNote MainRoomNote1,T2.Na3laNormal Na3laNormal1,T2.Na3laDorg Na3laDorg1,"
       MySQL = MySQL & " T2.Na3laCount Na3laCount1,T2.Na3laNote Na3laNote1,T2.ClothesInside ClothesInside1,"
       MySQL = MySQL & " T2.ClothesOutInside ClothesOutInside1,T2.ClothesCount ClothesCount1,"
       MySQL = MySQL & " T2.ClothesNote ClothesNote1,T2.ParkingGround ParkingGround1,T2.ParkingGdar ParkingGdar1,"
       MySQL = MySQL & " T2.ParkingCount ParkingCount1,T2.ParkingNote ParkingNote1,T2.Office Office1,"
       MySQL = MySQL & " T2.OfficeCount OfficeCount1,T2.OfficeNote OfficeNote1"
       MySQL = MySQL & " From TBL_measureMent"
       MySQL = MySQL & "        LEFT Outer JOIN TBL_measureMent2  T2"
       MySQL = MySQL & "       ON  TBL_measureMent.ID= T2.BDet_BD_ID"
       
       'db_createOrUpdateviewSQL "View_measureMent", MySQL
       
       MySQL = MySQL & "  Where (TBL_measureMent.ID = " & val(TxtSerial1.text) & ")"
     
     
  '   db_createOrUpdateviewSQL "View_measureMent", MySQL
     
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_measureMent.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_measureMent.rpt"
        End If
     
    If Dir(StrFileName) = "" Then
     'MsgBox StrFileName
        'GetMsgs 139, vbExclamation
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
      Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
     hide_logo = False
 End Function

Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
    With GridInstallments

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Level")) <> "" Then
                IntCounter = IntCounter + 1

              '  .TextMatrix(i, .ColIndex("NoteNo")) = Me.TxtSerial1.Text & "-" & IntCounter

            End If

        Next i

    End With
       
    IntCounter = 0

End Sub

