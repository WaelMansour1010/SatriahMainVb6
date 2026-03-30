VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_BusinessDialy 
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14265
   Icon            =   "Frm_BusinessDialy.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   14265
   WindowState     =   2  'Maximized
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
      Height          =   315
      ItemData        =   "Frm_BusinessDialy.frx":6852
      Left            =   15480
      List            =   "Frm_BusinessDialy.frx":6862
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
      Top             =   9285
      Width           =   14265
      _cx             =   25162
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
      AutoSizeChildren=   7
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3900
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
         Height          =   288
         Left            =   10152
         TabIndex        =   12
         Top             =   120
         Width           =   2964
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   576
         Left            =   180
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   -840
         Width           =   13908
         _cx             =   24527
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   756
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   720
         Width           =   14052
         _cx             =   24791
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
            Height          =   456
            Left            =   12552
            TabIndex        =   49
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1116
            _ExtentX        =   1958
            _ExtentY        =   794
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
            ButtonImage     =   "Frm_BusinessDialy.frx":687B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   576
            Left            =   8592
            TabIndex        =   50
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
            Width           =   1248
            _ExtentX        =   2196
            _ExtentY        =   1005
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
            ButtonImage     =   "Frm_BusinessDialy.frx":D0DD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   576
            Left            =   10668
            TabIndex        =   51
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   120
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   1005
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
            ButtonImage     =   "Frm_BusinessDialy.frx":D477
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   576
            Left            =   6540
            TabIndex        =   52
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   1005
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
            ButtonImage     =   "Frm_BusinessDialy.frx":13CD9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   576
            Left            =   4824
            TabIndex        =   53
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   120
            Width           =   1272
            _ExtentX        =   2249
            _ExtentY        =   1005
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
            ButtonImage     =   "Frm_BusinessDialy.frx":14073
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   576
            Left            =   120
            TabIndex        =   54
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   120
            Width           =   1128
            _ExtentX        =   1984
            _ExtentY        =   1005
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
            ButtonImage     =   "Frm_BusinessDialy.frx":1460D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   528
            Left            =   3204
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Width           =   1104
            _ExtentX        =   1958
            _ExtentY        =   926
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
            ButtonImage     =   "Frm_BusinessDialy.frx":149A7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   576
            Left            =   1560
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   1005
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
            ButtonImage     =   "Frm_BusinessDialy.frx":1B209
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   276
         Index           =   8
         Left            =   13296
         TabIndex        =   13
         Top             =   120
         Width           =   696
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
            Picture         =   "Frm_BusinessDialy.frx":1B5A3
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1B93D
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1BCD7
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1C071
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1C40B
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1C7A5
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1CB3F
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_BusinessDialy.frx":1D0D9
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
      ButtonImage     =   "Frm_BusinessDialy.frx":1D473
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
      ButtonImage     =   "Frm_BusinessDialy.frx":23CD5
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
      ButtonImage     =   "Frm_BusinessDialy.frx":2A537
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic100 
      Height          =   9285
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   14265
      _cx             =   25162
      _cy             =   16378
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   264
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   13668
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   23
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
            ButtonImage     =   "Frm_BusinessDialy.frx":2A8D1
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   24
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
            ButtonImage     =   "Frm_BusinessDialy.frx":2AC6B
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   25
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
            ButtonImage     =   "Frm_BusinessDialy.frx":2B005
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   26
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
            ButtonImage     =   "Frm_BusinessDialy.frx":2B39F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   612
            Left            =   12960
            Picture         =   "Frm_BusinessDialy.frx":2B739
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ”ÃÌ· «·«⁄„«· «·ÌÊ„Ì…"
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
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   3720
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4560
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4035
         Width           =   13740
         _cx             =   24236
         _cy             =   8043
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
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   4260
            Left            =   60
            TabIndex        =   29
            Top             =   225
            Width           =   13545
            _cx             =   23892
            _cy             =   7514
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Frm_BusinessDialy.frx":2CB3E
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1365
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2460
         Width           =   13740
         _cx             =   24236
         _cy             =   2408
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
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   510
            Width           =   1005
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4260
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   510
            Width           =   1005
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   7170
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   510
            Width           =   1005
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   510
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DCombo1 
            Bindings        =   "Frm_BusinessDialy.frx":2CD3F
            Height          =   315
            Left            =   5580
            TabIndex        =   32
            Top             =   510
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
         Begin MSDataListLib.DataCombo DCombo2 
            Bindings        =   "Frm_BusinessDialy.frx":2CD54
            Height          =   315
            Left            =   2595
            TabIndex        =   33
            Top             =   510
            Width           =   1665
            _ExtentX        =   2937
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
         Begin MSDataListLib.DataCombo DCombo3 
            Bindings        =   "Frm_BusinessDialy.frx":2CD69
            Height          =   315
            Left            =   0
            TabIndex        =   34
            Top             =   510
            Width           =   1440
            _ExtentX        =   2540
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
         Begin MSDataListLib.DataCombo DcCustmer 
            Bindings        =   "Frm_BusinessDialy.frx":2CD7E
            Height          =   315
            Left            =   8580
            TabIndex        =   62
            Top             =   510
            Width           =   2670
            _ExtentX        =   4710
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
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄·„ «·«› —«÷Ï"
            Height          =   330
            Index           =   5
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›Ê—„«‰ «·«› —«÷Ï"
            Height          =   330
            Index           =   12
            Left            =   2940
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄«„· «·«› —«÷Ï"
            Height          =   330
            Index           =   11
            Left            =   6210
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«∆« ⁄·Ï « ›«ﬁÌ…"
            Height          =   450
            Index           =   15
            Left            =   12360
            TabIndex        =   35
            Top             =   510
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   810
         Left            =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1560
         Width           =   13740
         _cx             =   24236
         _cy             =   1429
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
            Height          =   255
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   255
            Width           =   720
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   105
            Width           =   4512
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   270
            Left            =   9990
            TabIndex        =   42
            Top             =   255
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   476
            _Version        =   393216
            Format          =   106496001
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "Frm_BusinessDialy.frx":2CD93
            Height          =   315
            Left            =   5475
            TabIndex        =   43
            Top             =   315
            Width           =   3330
            _ExtentX        =   5874
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„"
            Height          =   210
            Index           =   4
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   225
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ"
            Height          =   330
            Index           =   2
            Left            =   11280
            TabIndex        =   46
            Top             =   315
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   210
            Index           =   7
            Left            =   8910
            TabIndex        =   45
            Top             =   315
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   15
            Index           =   14
            Left            =   4755
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   615
            Width           =   855
         End
      End
      Begin ImpulseButton.ISButton Cmd_DeleteRow 
         Height          =   240
         Left            =   12465
         TabIndex        =   57
         Top             =   8775
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   423
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
         ButtonImage     =   "Frm_BusinessDialy.frx":2CDA8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd_DeleteAll 
         Height          =   240
         Left            =   10800
         TabIndex        =   58
         Top             =   8775
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   423
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
         ButtonImage     =   "Frm_BusinessDialy.frx":2D342
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
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
Attribute VB_Name = "Frm_BusinessDialy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
                Dim RootAccount1 As String
                        Dim RootAccount2 As String
                        Dim RootAccount3 As String
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double
  


Sub FillGrid()

 'On Error GoTo ErrTrap
Dim k As Integer
Dim i As Integer

Dim IntCounter As Integer
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

'Sql = "SELECT dbo.projects.id, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.projects.project_code, dbo.projects.branche_ID, dbo.projects_des.oprid, "
'Sql = Sql & "   dbo.projects_des.fullcode, dbo.projects_des.project_no, dbo.projects_des.project_name AS Expr1, dbo.projects_des.[index], dbo.projects_des.des,"
'Sql = Sql & "   dbo.projects_des.qty, dbo.projects_des.cost, dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id,"
'Sql = Sql & "   dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id, dbo.projects_des.fullcode AS Expr2, dbo.projects_des.Remark, dbo.projects_des.esQty,"
'Sql = Sql & "   dbo.projects_des.PandUnitID, dbo.projects_des.CodeBand, dbo.projects_des.QtyNo, dbo.projects_des.PanID, dbo.projects_des.QtyExe,"
'Sql = Sql & "   dbo.projects_des.PriceExe , dbo.projects_des.TotalExe, dbo.projects_des.PrMainDesID, dbo.projects_des.SortID"
'Sql = Sql & "  FROM  dbo.projects INNER JOIN"
'Sql = Sql & "   dbo.projects_des ON dbo.projects.id = dbo.projects_des.project_id"
 '-----------
 
 'Sql = "SELECT dbo.Tbl_TradingContractDet.TContractDet_TContractID, dbo.Tbl_TradingContractDet.TContractDet_specification,"
  'Sql = Sql & "   dbo.Tbl_TradingContractDet.TContractDet_specificationAr, dbo.Tbl_BusinessDialyDet.BDet_BD_ID, dbo.Tbl_BusinessDialyDet.BDet_BandNo,"
   ' Sql = Sql & "             dbo.Tbl_BusinessDialyDet.BDet_Qun, dbo.Tbl_BusinessDialyDet.BDet_Name, dbo.Tbl_BusinessDialyDet.BDet_NameE, dbo.Tbl_BusinessDialyDet.BDet_EmpID,"
  'Sql = Sql & "               dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID, TblEmployee_1.Emp_Name AS FormanName,"
  'Sql = Sql & "               TblEmployee_1.Emp_Namee AS FormanNameE, dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID, TblEmployee_2.Emp_Name AS TecherName,"
  '       Sql = Sql & "        TblEmployee_2.Emp_Namee AS TecherNameE,dbo.Tbl_TradingContractDet.ID AS fullcode"
 'Sql = Sql & "   FROM  dbo.TblEmployee INNER JOIN"
    '   Sql = Sql & "          dbo.Tbl_TradingContract INNER JOIN"
   'Sql = Sql & "             dbo.Tbl_TradingContractDet ON dbo.Tbl_TradingContract.ID = dbo.Tbl_TradingContractDet.TContractDet_TContractID INNER JOIN"
  'Sql = Sql & "              dbo.Tbl_BusinessDialy INNER JOIN"
   'sql = Sql & "              dbo.Tbl_BusinessDialyDet ON dbo.Tbl_BusinessDialy.ID = dbo.Tbl_BusinessDialyDet.BDet_BD_ID ON"
   ' Sql = Sql & "             dbo.Tbl_TradingContract.ID = dbo.Tbl_BusinessDialy.TradingContractID ON dbo.TblEmployee.Emp_ID = dbo.Tbl_BusinessDialyDet.BDet_EmpID INNER JOIN"
  ' Sql = Sql & "              dbo.TblEmployee AS TblEmployee_1 ON dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID = TblEmployee_1.Emp_ID INNER JOIN"
 ' Sql = Sql & "             dbo.TblEmployee AS T4blEmployee_2 ON dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = TblEmployee_2.Emp_ID"
 
 '----------
                              
'Sql = Sql & " WHERE (dbo.Tbl_TradingContractDet.TContractDet_TContractID = " & val(TxtSearchCode.Text) & ")"

Set Rs2 = New ADODB.Recordset
 Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With Grid
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
i = 0
For k = 1 To .Rows - 2
i = i + 1

.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("NoteNo")) = IIf(IsNull(Rs2("FullCode").value), "", Rs2("FullCode").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("TContractDet_specificationAr").value), "", Rs2("TContractDet_specificationAr").value)
.TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(Rs2("TContractDet_specification").value), "", Rs2("TContractDet_specification").value)
.TextMatrix(i, .ColIndex("EmpName")) = Me.DCombo1
.TextMatrix(i, .ColIndex("EmpID")) = Me.DCombo1.BoundText
.TextMatrix(i, .ColIndex("EmpName")) = Me.DCombo1
.TextMatrix(i, .ColIndex("FromanID")) = Me.DCombo2.BoundText
.TextMatrix(i, .ColIndex("Forman")) = Me.DCombo2
.TextMatrix(i, .ColIndex("TeacherID")) = Me.DCombo3.BoundText
.TextMatrix(i, .ColIndex("TechMan")) = Me.DCombo3


Rs2.MoveNext
Next k
End With
End If
ErrTrap:
End Sub

 

 

Private Sub Cmd_DeleteAll_Click()
If Me.TxtModFlg.Text <> "R" Then


 Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2

End If
End Sub

Private Sub Cmd_DeleteRow_Click()
If Me.TxtModFlg.Text <> "R" Then

RemoveGridRow

End If
End Sub

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub RemoveGridRow()

    With Me.Grid
'MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub


Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
'TxtSerial1.Text = ""
End If
End Sub

 

  
Private Sub DcCustmer_Click(Area As Integer)
    Dim My_SQL As String

    My_SQL = "  select CusID,CusName,TT.ID from TblCustemers  "
    
    My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
    My_SQL = My_SQL & " Where TT.TContract_CustID =  " & val(DcCustmer.BoundText)
    My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
    My_SQL = My_SQL & " order by CusName "
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        TxtSearchCode = rsDummy!ID & ""
    End If

End Sub

Private Sub DCombo1_Click(Area As Integer)
'fillgrid
    'Dim mEmpName As String
    'getemployeeCode val(txtCode(Index)), mEmpName
        
        txtCode(0) = getemployeeCode(val(DCombo1.BoundText))

End Sub

Private Sub DCombo2_Click(Area As Integer)
    
    txtCode(1) = getemployeeCode(val(DCombo2.BoundText))
End Sub

Private Sub DCombo3_Click(Area As Integer)
'fillgrid

txtCode(2) = getemployeeCode(val(DCombo3.BoundText))
End Sub

 Private Sub Form_Load()
   ' On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from Tbl_BusinessDialy order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me

    Dim Dcombos As New ClsDataCombos
    
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DCombo1
    Dcombos.GetEmployees Me.DCombo2
    Dcombos.GetEmployees Me.DCombo3
    
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  "
        My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
        My_SQL = My_SQL & " Where IsNull(Id,0) <>0"
        My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
        My_SQL = My_SQL & " order by CusName "
        
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  "
        My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
        My_SQL = My_SQL & " Where IsNull(Id,0) <>0"
        My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
        My_SQL = My_SQL & " order by CusName "
       
    End If
    fill_combo DcCustmer, My_SQL
    

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
Sub ChangeLang()
On Error GoTo ErrTrap
'Label5.Caption = "Report Calling"
  Dim XPic As IPictureDisp
  
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Me.Caption = "Daily Works "
         
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

lbl(2) = "date"
lbl(15) = "Based on agreement"
'lbl(0) = "Responsible compound"
lbl(7) = "Branch"
lbl(11) = "Default Worker"
lbl(12) = "Default Forman"
Label1(5) = "Default Teacher"

With Grid
.TextMatrix(0, .ColIndex("Qun")) = "Quantity"
.TextMatrix(0, .ColIndex("TConID")) = "Item No"
.TextMatrix(0, .ColIndex("Name")) = "Description"
.TextMatrix(0, .ColIndex("Qun")) = "Quntity"
.TextMatrix(0, .ColIndex("TechMan")) = "Techer"
.TextMatrix(0, .ColIndex("EmpName")) = "Worker Name"
.TextMatrix(0, .ColIndex("Forman")) = "Forman"
.TextMatrix(0, .ColIndex("DayMeter")) = "Day Meter"
.TextMatrix(0, .ColIndex("NameE")) = "Notes"


End With
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
                   StrSQL = "Delete From Tbl_BusinessDialyDet Where BDet_BD_ID=" & val(Me.TxtSerial1.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
               
              End If
    RsSavRec.Fields("BD_Date").value = XPDtbTrans.value
    RsSavRec.Fields("BD_BranchID").value = Dcbranch.BoundText
    RsSavRec.Fields("BD_Notes").value = (Me.txtRemarks.Text)
    RsSavRec.Fields("TradingContractID").value = val((Me.TxtSearchCode.Text))
    RsSavRec.Fields("UserID").value = (Me.DCboUserName.BoundText)
       
    RsSavRec.update
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from Tbl_BusinessDialyDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
    msg = msg & " ‰«“·"
    Else
    msg = msg & "Waiver/Sale of Shares"
    End If
    Dim str2 As String
    With Me.Grid
       For i = .FixedRows To .Rows - 1
       If (.TextMatrix(i, .ColIndex("TConID"))) <> "" Then
       RsDevsub.AddNew
                RsDevsub("BDet_BD_ID").value = val(Me.TxtSerial1.Text)
              

                RsDevsub("BDet_BandNo").value = IIf((.TextMatrix(i, .ColIndex("NoteNo"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteNo"))))
                 RsDevsub("BDet_Qun").value = IIf((.TextMatrix(i, .ColIndex("Qun"))) = "", Null, (.TextMatrix(i, .ColIndex("Qun"))))
                RsDevsub("BDet_Name").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, (.TextMatrix(i, .ColIndex("Name"))))
                RsDevsub("BDet_NameE").value = IIf((.TextMatrix(i, .ColIndex("NameE"))) = "", Null, (.TextMatrix(i, .ColIndex("NameE"))))
                RsDevsub("BDet_EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, (.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("BDet_EmpFormanID").value = IIf((.TextMatrix(i, .ColIndex("FromanID"))) = "", Null, (.TextMatrix(i, .ColIndex("FromanID"))))
                RsDevsub("BDet_EmpTecherID").value = IIf((.TextMatrix(i, .ColIndex("TeacherID"))) = "", Null, (.TextMatrix(i, .ColIndex("TeacherID"))))
                RsDevsub("BDet_DayMeter").value = IIf((.TextMatrix(i, .ColIndex("DayMeter"))) = "", Null, (.TextMatrix(i, .ColIndex("DayMeter"))))
                RsDevsub("TConID").value = IIf((.TextMatrix(i, .ColIndex("TConID"))) = "", Null, (.TextMatrix(i, .ColIndex("TConID"))))
                
        
       RsDevsub.update
       
      End If
     Next i
    End With

    
'''///////////////
  
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                msg = msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               msg = " This record alredy saved... " & CHR(13)
                msg = msg + " You want to enter another record?"
           End If
                If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
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
 Dim TContractCustID As Double

    Dim i As Integer
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)

    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("BD_Date").value), "", RsSavRec.Fields("BD_Date").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BD_BranchID").value), "", RsSavRec.Fields("BD_BranchID").value)
    txtRemarks.Text = IIf(IsNull(RsSavRec.Fields("BD_Notes").value), "", RsSavRec.Fields("BD_Notes").value)
    TxtSearchCode.Text = IIf(IsNull(RsSavRec.Fields("TradingContractID").value), "", RsSavRec.Fields("TradingContractID").value)

    Get_TradingContractinfo TxtSearchCode.Text, TContractCustID, 0

    DcCustmer.BoundText = TContractCustID

    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
     
FullGridData
'RelinGrid
ErrTrap:
End Sub

  
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode As String
    Dim msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
    'Dim Rs2 As ADODB.Recordset
On Error GoTo ErrTrap
    With Grid
               
 If Trim(.TextMatrix(Row, .ColIndex("EmpName"))) = "" Then
 
  .TextMatrix(Row, .ColIndex("EmpId")) = val(DCombo1.BoundText)

  .TextMatrix(Row, .ColIndex("EmpName")) = (DCombo1.Text)
 
 End If
  
 If Trim(.TextMatrix(Row, .ColIndex("Forman"))) = "" Then
 
  .TextMatrix(Row, .ColIndex("FromanID")) = val(DCombo2.BoundText)

  .TextMatrix(Row, .ColIndex("Forman")) = (DCombo2.Text)
 
 End If
  
 If (.TextMatrix(Row, .ColIndex("TechMan"))) = "" Then
 
  .TextMatrix(Row, .ColIndex("TeacherID")) = val(DCombo3.BoundText)

  .TextMatrix(Row, .ColIndex("TechMan")) = (DCombo3.Text)
 
 End If


'If Trim(.TextMatrix(Row, .ColIndex("Name"))) = "" Then
'
'  .TextMatrix(Row, .ColIndex("TConID")) = 0
'
'  .TextMatrix(Row, .ColIndex("Name")) = ""
'
' End If
  
  
Select Case .ColKey(Col)
  
Case "EmpName"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpId"), False, True)
                .TextMatrix(Row, .ColIndex("EmpId")) = StrAccountCode
            

  Case "Forman"
  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("FromanID"), False, True)
                .TextMatrix(Row, .ColIndex("FromanID")) = StrAccountCode


          Case "TechMan"
  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TeacherID"), False, True)
                .TextMatrix(Row, .ColIndex("TeacherID")) = StrAccountCode

 
        Case "Name"
        
         If Trim(.ComboData = "") Then
            .TextMatrix(Row, .ColIndex("Name")) = ""
            .TextMatrix(Row, .ColIndex("TConID")) = ""
            Exit Sub
         End If
          StrAccountCode = .ComboData
                

                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TConID"), False, True)
                
                    .TextMatrix(Row, .ColIndex("TConID")) = StrAccountCode
    
                
         
               
               
                   StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID, dbo.TblProcessDEF.ProcessName,TblProcessDEF.TblProcessDEFID,"
                   StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
                   StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
                   StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
                   StrSQL = StrSQL & " where TblProcessDEF.TblProcessDEFID = '" & val(.TextMatrix(Row, .ColIndex("TConID"))) & "'"
                   StrSQL = StrSQL & " And Tbl_TradingContractDet.TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
                   

                    
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     .TextMatrix(Row, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("ProcessNameE").value), "", rs.Fields("ProcessNameE").value)
                     .TextMatrix(Row, .ColIndex("Name")) = IIf(IsNull(rs.Fields("ProcessName").value), "", rs.Fields("ProcessName").value)
                     .TextMatrix(Row, .ColIndex("NoteNo")) = (rs!ID & "")
                  
                     .TextMatrix(Row, .ColIndex("DayMeter")) = IIf(IsNull(rs.Fields("Interval").value), "", rs.Fields("Interval").value)
                     
                   
                  
'                Else
'                   ' .TextMatrix(Row, .ColIndex("TConID")) = ""
'                    .TextMatrix(Row, .ColIndex("NameE")) = ""
'                    .TextMatrix(Row, .ColIndex("Name")) = ""
'                    .TextMatrix(Row, .ColIndex("DayMeter")) = ""
'                End If

      Case "NameE"
'        StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TConID"), False, True)
'                .TextMatrix(Row, .ColIndex("TConID")) = StrAccountCode
'
'
'               StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID, dbo.TblProcessDEF.ProcessName,"
'               StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
'               StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
'               StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
'               StrSQL = StrSQL & " where ID = '" & val(.TextMatrix(Row, .ColIndex("TConID"))) & "'"
'               StrSQL = StrSQL & " And Tbl_TradingContractDet.TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
'
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                 .TextMatrix(Row, .ColIndex("Name")) = IIf(IsNull(rs.Fields("ProcessName").value), "", rs.Fields("ProcessName").value)
'                 .TextMatrix(Row, .ColIndex("NoteNo")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
'                 .TextMatrix(Row, .ColIndex("DayMeter")) = IIf(IsNull(rs.Fields("Interval").value), "", rs.Fields("Interval").value)
                    

           Case "NoteNo"
              
             StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID, dbo.TblProcessDEF.ProcessName,Tbl_TradingContractDet.ProcessDEFID,"
               StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
               StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
               StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & " where ID = '" & .Cell(flexcpText, Row, Col) & "'"
               StrSQL = StrSQL & " And Tbl_TradingContractDet.TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
               
                 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 
                 .TextMatrix(Row, .ColIndex("Name")) = IIf(IsNull(rs.Fields("ProcessName").value), "", rs.Fields("ProcessName").value)
                 .TextMatrix(Row, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("ProcessNameE").value), "", rs.Fields("ProcessNameE").value)
                 .TextMatrix(Row, .ColIndex("DayMeter")) = IIf(IsNull(rs.Fields("Interval").value), "", rs.Fields("Interval").value)
                 .TextMatrix(Row, .ColIndex("TConID")) = IIf(IsNull(rs.Fields("ProcessDEFID").value), "", rs.Fields("ProcessDEFID").value)
               
              Case "TConID"
                           StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID, dbo.TblProcessDEF.ProcessName,Tbl_TradingContractDet.ProcessDEFID,"
                    StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
                    StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
                    StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
                    StrSQL = StrSQL & " where ProcessDEFID = '" & .Cell(flexcpText, Row, Col) & "'"
                    StrSQL = StrSQL & " And Tbl_TradingContractDet.TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
                    If .TextMatrix(Row, .ColIndex("NoteNo")) <> "" Then
                    '    StrSQL = StrSQL & " and ID = '" & .TextMatrix(Row, .ColIndex("NoteNo")) & "'"
                    End If
               
                 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 If Not rs.EOF Then
                    .TextMatrix(Row, .ColIndex("Name")) = IIf(IsNull(rs.Fields("ProcessName").value), "", rs.Fields("ProcessName").value)
                    .TextMatrix(Row, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("ProcessNameE").value), "", rs.Fields("ProcessNameE").value)
                    .TextMatrix(Row, .ColIndex("DayMeter")) = IIf(IsNull(rs.Fields("Interval").value), "", rs.Fields("Interval").value)
                    .TextMatrix(Row, .ColIndex("NoteNo")) = rs!ID & ""
                End If
                 '.TextMatrix(Row, .ColIndex("TConID")) = IIf(IsNull(rs.Fields("ProcessDEFID").value), "", rs.Fields("ProcessDEFID").value)
               
  End Select



        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

End With


ReLineGrid

ErrTrap:
End Sub


Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
    With Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("NoteNo")) <> "" Then
                IntCounter = IntCounter + 1

              '  .TextMatrix(i, .ColIndex("NoteNo")) = Me.TxtSerial1.Text & "-" & IntCounter

            End If

        Next i

    End With
       
    IntCounter = 0

End Sub



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Grid

   Select Case .ColKey(Col)
        Case "Qun"
        .ComboList = ""
           Case "NoteNo"
        .ComboList = ""
        Case "DayMeter"
        .ComboList = ""
        Case "Name"
       ' Cancel = True
        End Select
        
    End With
 
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
 
            Case "EmpName"
             .TextMatrix(Row, .ColIndex("EmpName")) = ""
                StrSQL = "select * from TblEmployee"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = Grid.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
   
             Case "Forman"
                StrSQL = "select * from TblEmployee"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = Grid.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                   StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
               Case "TechMan"
                StrSQL = "select * from TblEmployee"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = Grid.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
                
                If StrComboList <> "" Then
                   StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                
         
                Case "Name"
                
               StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID,dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
               StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
               StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
               StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & " Where TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 StrComboList = Grid.BuildComboList(rs, "ProcessName", "TblProcessDEFID")
  
                
                If StrComboList <> "" Then
                   StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                
               Case "NameE"
            
               StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID,dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
               StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
               StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
               StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & " Where TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                StrComboList = Grid.BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
   
    
          Case "NoteNo"
            
               StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID,dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
               StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
               StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
               StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & " Where TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                StrComboList = Grid.BuildComboList(rs, "TblProcessDEFID", "ID")
                'StrComboList = Grid.BuildComboList(rs, "ID", "ID")
                
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ColComboList(Col) = StrComboList
                
            Case "TConID"
            
               StrSQL = "Select dbo.Tbl_TradingContractDet.ID,dbo.Tbl_TradingContractDet.TContractDet_TContractID,dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
               StrSQL = StrSQL & " dbo.TblProcessDEF.ProcessNameE , dbo.TblProcessDEF.interval, dbo.TblProcessDEF.IntervalID"
               StrSQL = StrSQL & " FROM  dbo.Tbl_TradingContractDet INNER JOIN"
               StrSQL = StrSQL & " dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & " Where TContractDet_TContractID = '" & TxtSearchCode.Text & "'"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                StrComboList = Grid.BuildComboList(rs, "TblProcessDEFID", "TblProcessDEFID")
                'StrComboList = Grid.BuildComboList(rs, "ID", "ID")
                
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ColComboList(Col) = StrComboList


        End Select

    End With
    'ReLineGrid
    

End Sub

 
Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double

    '---------------------- check if data Vaclete -----------------------
      
           
     If val(Dcbranch.BoundText) = 0 And Dcbranch.Text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·›—⁄  "
     Else
     MsgBox "Please Select Branch"
     End If
     Dcbranch.SetFocus
     Exit Sub
     End If
     
    If val(DcCustmer.BoundText) = 0 And DcCustmer.Text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·« ›«ﬁÌ…  "
     Else
     MsgBox "Please Select Project"
     End If
     DcCustmer.SetFocus
     Exit Sub
     End If
     Dim i As Long
'     With Me.Grid
'        For i = .FixedRows To .Rows - 1
'           If (.TextMatrix(i, .ColIndex("NoteNo"))) <> "" Then
'                If Trim(.TextMatrix(i, .ColIndex("TechMan"))) = "" Then
'                     If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·„⁄·„ «·«› —«÷Ï ”ÿ— —ﬁ„ " & i
'                     Else
'                        MsgBox "Please Select Teacher"
'                     End If
'                        .SetFocus
'                     Exit Sub
'                End If
'                If Trim(.TextMatrix(i, .ColIndex("EmpName"))) = "" Then
'                     If SystemOptions.UserInterface = ArabicInterface Then
'                     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·⁄«„· «·«› —«÷Ï ”ÿ— —ﬁ„ " & i
'                     Else
'                        MsgBox "Please Select Worker"
'                     End If
'                        .SetFocus
'                     Exit Sub
'                End If
'                If Trim(.TextMatrix(i, .ColIndex("Forman"))) = "" Then
'                     If SystemOptions.UserInterface = ArabicInterface Then
'                     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·›Ê—„«‰ «·«› —«÷Ï ”ÿ— —ﬁ„ " & i
'                     Else
'                        MsgBox "Please Select forman"
'                     End If
'                        .SetFocus
'                     Exit Sub
'                End If
'
'               If Trim(.TextMatrix(i, .ColIndex("TConID"))) = "" Then
'                     If SystemOptions.UserInterface = ArabicInterface Then
'                     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·»Ì«‰ ”ÿ— —ﬁ„ " & i
'                     Else
'                        MsgBox "Please Select Note"
'                     End If
'                        .SetFocus
'                     Exit Sub
'                End If
'
'
'           End If
'        Next
  '   End With
     
'    If val(DCombo1.BoundText) = 0 And DCombo1.Text = "" Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·⁄«„· «·«› —«÷Ï  "
'     Else
'     MsgBox "Please Select Branch"
'     End If
'     DCombo1.SetFocus
'     Exit Sub
'     End If
'
'     If val(DCombo2.BoundText) = 0 And DCombo2.Text = "" Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·›Ê—„«‰  "
'     Else
'     MsgBox "Please Select Branch"
'     End If
'     DCombo2.SetFocus
'     Exit Sub
'     End If
'
'
'     If val(DCombo3.BoundText) = 0 And DCombo3.Text = "" Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·„⁄·„ «·«› —«÷Ì  "
'     Else
'     MsgBox "Please Select Branch"
'     End If
'     DCombo3.SetFocus
'     Exit Sub
'     End If
'
     

    With Me.Grid
          If .Rows >= 2 Then
          If val(.TextMatrix(1, .ColIndex("EmpId"))) = 0 Or val(.TextMatrix(1, .ColIndex("TConID"))) = 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «·»Ì«‰«  "
             Else
                 MsgBox "Please Enter Data"
              End If
              Exit Sub
            End If
         End If
        If .Rows < 2 Then
           If SystemOptions.UserInterface = ArabicInterface Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·»Ì«‰«  "
           Else
           MsgBox "Please Enter Data"
           End If
           Exit Sub
           End If
    End With
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
    StrRecID = new_id("Tbl_BusinessDialy", "ID", "")
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
    Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
sql = "Select  dbo.Tbl_BusinessDialyDet.ID,dbo.Tbl_BusinessDialyDet.TConID ,dbo.Tbl_BusinessDialyDet.BDet_BD_ID, dbo.Tbl_BusinessDialyDet.BDet_BandNo, dbo.Tbl_BusinessDialyDet.BDet_Qun,dbo.Tbl_BusinessDialyDet.BDet_DayMeter, "

sql = sql & "   dbo.Tbl_BusinessDialyDet.BDet_Name, dbo.Tbl_BusinessDialyDet.BDet_NameE, dbo.Tbl_BusinessDialyDet.BDet_EmpID, TblEmployee_1.Emp_ID, "

sql = sql & "    TblEmployee_1.Emp_Name AS EmpName, TblEmployee_1.Emp_Namee AS EmpNameE, dbo.TblEmployee.Emp_Name AS Froman, "

sql = sql & "   dbo.TblEmployee.Emp_Namee AS formanE, TblEmployee_2.Emp_Name AS Techer, TblEmployee_2.Emp_Namee AS TecherE, "

sql = sql & "   dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID, dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID"

sql = sql & "  FROM    dbo.Tbl_BusinessDialyDet Left OUTER JOIN"

sql = sql & "  dbo.TblEmployee AS TblEmployee_2 ON dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = TblEmployee_2.Emp_ID Left OUTER JOIN"

sql = sql & " dbo.TblEmployee ON dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID = dbo.TblEmployee.Emp_ID Left OUTER JOIN"

sql = sql & "   dbo.TblEmployee AS TblEmployee_1 ON dbo.Tbl_BusinessDialyDet.BDet_EmpID = TblEmployee_1.Emp_ID"

sql = sql & "  Where (dbo.Tbl_BusinessDialyDet.BDet_BD_ID =" & val(TxtSerial1.Text) & ")"

Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Grid.Rows = Grid.Rows + Rs1.RecordCount
     Dim i As Integer
     With Me.Grid
                For i = .FixedRows To Rs1.RecordCount
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("NoteNo")) = IIf(IsNull(Rs1("BDet_BandNo").value), 0, Rs1("BDet_BandNo").value)
                .TextMatrix(i, .ColIndex("Qun")) = IIf(IsNull(Rs1("BDet_Qun").value), 0, Rs1("BDet_Qun").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("BDet_Name").value), 0, Rs1("BDet_Name").value)
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(Rs1("BDet_NameE").value), 0, Rs1("BDet_NameE").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("BDet_EmpID").value), 0, Rs1("BDet_EmpID").value)
                .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("EmpName").value), 0, Rs1("EmpName").value)
                 .TextMatrix(i, .ColIndex("FromanID")) = IIf(IsNull(Rs1("BDet_EmpFormanID").value), 0, Rs1("BDet_EmpFormanID").value)
                 .TextMatrix(i, .ColIndex("Forman")) = IIf(IsNull(Rs1("Froman").value), 0, Rs1("Froman").value)
                 .TextMatrix(i, .ColIndex("TeacherID")) = IIf(IsNull(Rs1("BDet_EmpTecherID").value), 0, Rs1("BDet_EmpTecherID").value)
                 .TextMatrix(i, .ColIndex("TechMan")) = IIf(IsNull(Rs1("Techer").value), 0, Rs1("Techer").value)
                 .TextMatrix(i, .ColIndex("DayMeter")) = IIf(IsNull(Rs1("BDet_DayMeter").value), 0, Rs1("BDet_DayMeter").value)
                 .TextMatrix(i, .ColIndex("TConID")) = IIf(IsNull(Rs1("TConID").value), 0, Rs1("TConID").value)
                 
                 

                 Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub ISButton8_Click()


FrmProjectSearch.C1Tab1 = 3
FrmProjectSearch.Caption = "»ÕÀ «·«⁄„«· «·ÌÊ„Ì…"
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
 
Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
  'Dim mEmpName As String
    'getemployeeCode val(txtCode(Index)), mEmpName
    If txtCode(Index).Text = "" Then Exit Sub
Dim mEmpId As Integer
If KeyAscii = vbKeyReturn Then
    Select Case Index
    Case 0
        GetEmployeeIDFromCode txtCode(Index), mEmpId
        DCombo1.BoundText = mEmpId

    Case 1
        GetEmployeeIDFromCode txtCode(Index), mEmpId
        DCombo2.BoundText = mEmpId

    Case 2
        GetEmployeeIDFromCode txtCode(Index), mEmpId
        DCombo3.BoundText = mEmpId
    End Select
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim TContractCustID As Double

If TxtSearchCode.Text = "" Then Exit Sub

If KeyAscii = vbKeyReturn Then
Get_TradingContractinfo TxtSearchCode.Text, TContractCustID, 0

If TContractCustID = 0 Then
    TxtSearchCode.Text = ""
End If
DcCustmer.BoundText = TContractCustID
End If
End Sub

Private Sub TxtSearchCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Dim TContractCustID As Double
 FrmProjectSearch.C1Tab1 = 4
 FrmProjectSearch.Label11.Caption = 1
 FrmProjectSearch.Caption = "»ÕÀ «·« ›«ﬁÌ«  "
 FrmProjectSearch.show vbModal
 
' Get_TradingContractinfo TxtSearchCode.Text, TContractCustID, 0


'DcCustmer.BoundText = TContractCustID
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
    Dim msg As String
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
               With Me.Grid
      
    End With
          
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.Delete
                 
                 StrSQL = "Delete From Tbl_BusinessDialyDet Where BDet_BD_ID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          
                                                        
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
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
       ' Grid.Enabled = False
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
    Dim msg As String
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
            msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            msg = msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            msg = msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            msg = "Sorry I've been to delete this record" & CHR(13)
            msg = msg & "By another user on the network " & CHR(13)
            msg = msg & "Data will be updated"
            End If
            MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim msg As String
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
            msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            msg = msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            msg = msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               msg = "Sorry I've been to delete this record" & CHR(13)
            msg = msg & "By another user on the network " & CHR(13)
            msg = msg & "Data will be updated"
            End If
            MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
            Grid.Rows = Grid.Rows + 1
        Me.DCboUserName.BoundText = user_id
Grid.Enabled = True
'ISButton6.Enabled = True
'ISButton4.Enabled = True
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            msg = "⁄›Ê«" & CHR(13)
            msg = msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            msg = msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            msg = "Sorry.." & CHR(13)
            msg = msg & " You can not edit this the record now" & CHR(13)
            msg = msg & "It was being edited by another user on the network"
           
            End If
            MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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
    
    
    Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
            
 
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = branch_id
  
  AddNewRecored

ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim msg As String
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
            msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            msg = msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            msg = msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               msg = "Sorry I've been to delete this record" & CHR(13)
            msg = msg & "By another user on the network " & CHR(13)
            msg = msg & "Data will be updated"
            End If
            MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim msg As String
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
            msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            msg = msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            msg = msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            msg = "Sorry I've been to delete this record" & CHR(13)
            msg = msg & "By another user on the network " & CHR(13)
            msg = msg & "Data will be updated"
            End If
            MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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
    Dim msg As String
 
'--------------------------------------------------------------------------------------------

   MySQL = MySQL & " SELECT dbo.Tbl_BusinessDialy.ID, dbo.Tbl_BusinessDialy.BD_Date, dbo.Tbl_BusinessDialy.BD_BranchID, dbo.TblBranchesData.branch_name,"
   MySQL = MySQL & "         dbo.TblBranchesData.branch_namee, dbo.Tbl_BusinessDialy.BD_Notes, dbo.Tbl_BusinessDialyDet.BDet_BD_ID, dbo.Tbl_BusinessDialyDet.BDet_BandNo,"
    MySQL = MySQL & "        dbo.Tbl_BusinessDialyDet.BDet_Qun, dbo.Tbl_BusinessDialyDet.BDet_Name, dbo.Tbl_BusinessDialyDet.BDet_NameE,dbo.Tbl_BusinessDialyDet.BDet_DayMeter, TblEmployee_1.Emp_ID AS EmpID,"
    MySQL = MySQL & "         TblEmployee_1.Emp_Name AS EmpName, TblEmployee_1.Emp_Namee AS EmpNameE, dbo.TblEmployee.Emp_ID AS ForemanID,"
    MySQL = MySQL & "          dbo.TblEmployee.Emp_Name AS ForemanName, dbo.TblEmployee.Emp_Namee AS ForemanNameE, TblEmployee_2.Emp_ID AS TeacherID,"
    MySQL = MySQL & "          TblEmployee_2.Emp_Name AS TeacherName, TblEmployee_2.Emp_Namee AS TeacherNameE, dbo.Tbl_TradingContract.TContract_CustID,"
    MySQL = MySQL & "          dbo.TblCustemers.CusName , dbo.TblCustemers.Fullcode, dbo.Tbl_BusinessDialy.TradingContractID"
   MySQL = MySQL & "  FROM  dbo.Tbl_BusinessDialy INNER JOIN"
   MySQL = MySQL & "              dbo.Tbl_BusinessDialyDet ON dbo.Tbl_BusinessDialy.ID = dbo.Tbl_BusinessDialyDet.BDet_BD_ID INNER JOIN"
   MySQL = MySQL & "               dbo.TblBranchesData ON dbo.Tbl_BusinessDialy.BD_BranchID = dbo.TblBranchesData.branch_id Left Outer JOIN"
   MySQL = MySQL & "             dbo.TblEmployee AS TblEmployee_1 ON dbo.Tbl_BusinessDialyDet.BDet_EmpID = TblEmployee_1.Emp_ID INNER JOIN"
   MySQL = MySQL & "            dbo.Tbl_TradingContract ON dbo.Tbl_BusinessDialy.TradingContractID = dbo.Tbl_TradingContract.ID INNER JOIN"
   MySQL = MySQL & "             dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID Left OUTER JOIN"
   MySQL = MySQL & "           dbo.TblEmployee ON dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID = dbo.TblEmployee.Emp_ID Left OUTER JOIN"
   MySQL = MySQL & "          dbo.TblEmployee AS TblEmployee_2 ON dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = TblEmployee_2.Emp_ID"
   MySQL = MySQL & "  Where (dbo.Tbl_BusinessDialy.ID =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_BusinessDialy.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_BusinessDialy.rpt"
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
        msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        msg = "No Data"
        End If
        MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
    Dim msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, msg, True
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


 Sub filgrid1()
 End Sub
Sub RelinGrid()
 End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
   Set rs = New ADODB.Recordset
   My_SQL = "Tbl_BusinessDialy"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

  
