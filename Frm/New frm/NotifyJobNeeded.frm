VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form NotifyJobNeeded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Œÿ«— «Õ Ì«Ã«  „ÊŸ›"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14205
   Icon            =   "NotifyJobNeeded.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   14205
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   5040
      Width           =   14295
      Begin ImpulseButton.ISButton ISButton5 
         Height          =   330
         Left            =   11760
         TabIndex        =   53
         ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
         ButtonImage     =   "NotifyJobNeeded.frx":6852
         ButtonImageDisabled=   "NotifyJobNeeded.frx":D0B4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton6 
         Height          =   330
         Left            =   9840
         TabIndex        =   54
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         ButtonImage     =   "NotifyJobNeeded.frx":2C29E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Height          =   2295
      Left            =   0
      TabIndex        =   39
      Top             =   840
      Width           =   14175
      Begin VB.TextBox TxtResons 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   720
         Width           =   6435
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtCuntJob 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   10680
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1665
      End
      Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
         Height          =   255
         Left            =   5760
         TabIndex        =   41
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "NotifyJobNeeded.frx":32B00
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
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
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   8760
         TabIndex        =   42
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   94109697
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   570
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
         Top             =   1560
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   1005
         ButtonPositionImage=   3
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
         ButtonImage     =   "NotifyJobNeeded.frx":32B15
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         LowerToggledContent=   0   'False
      End
      Begin MSDataListLib.DataCombo DcbJob 
         Bindings        =   "NotifyJobNeeded.frx":39377
         Height          =   315
         Left            =   8250
         TabIndex        =   2
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
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
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   2
         Left            =   10290
         TabIndex        =   49
         Top             =   255
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”»«» «·≈Õ Ì«Ã"
         Height          =   285
         Index           =   0
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›… «·„ÿ·Ê»…"
         Height          =   195
         Index           =   15
         Left            =   12480
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”·”·"
         Height          =   285
         Index           =   4
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄œœ «·„ÿ·Ê»"
         Height          =   195
         Index           =   3
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1080
         Width           =   1230
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "NotifyJobNeeded.frx":3938C
      Left            =   15480
      List            =   "NotifyJobNeeded.frx":3939C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   14145
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   9
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
         ButtonImage     =   "NotifyJobNeeded.frx":393B5
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   10
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
         ButtonImage     =   "NotifyJobNeeded.frx":3974F
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   11
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
         ButtonImage     =   "NotifyJobNeeded.frx":39AE9
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
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
         ButtonImage     =   "NotifyJobNeeded.frx":39E83
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈Œÿ«— «Õ Ì«Ã«  „ÊŸ›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   4200
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "NotifyJobNeeded.frx":3A21D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   16
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
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
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1545
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7920
      Width           =   14235
      _cx             =   25109
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
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   19
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12240
            TabIndex        =   20
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "NotifyJobNeeded.frx":3B7CB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   7560
            TabIndex        =   21
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            ButtonImage     =   "NotifyJobNeeded.frx":4202D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   10080
            TabIndex        =   22
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "NotifyJobNeeded.frx":423C7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   5280
            TabIndex        =   23
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            ButtonImage     =   "NotifyJobNeeded.frx":48C29
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   3000
            TabIndex        =   24
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            ButtonImage     =   "NotifyJobNeeded.frx":48FC3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   600
            TabIndex        =   25
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "NotifyJobNeeded.frx":4955D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9720
         TabIndex        =   31
         Top             =   120
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   330
         Left            =   7680
         TabIndex        =   32
         ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
         ButtonImage     =   "NotifyJobNeeded.frx":498F7
         ButtonImageDisabled=   "NotifyJobNeeded.frx":50159
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton4 
         Height          =   330
         Left            =   6000
         TabIndex        =   33
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         ButtonImage     =   "NotifyJobNeeded.frx":6F343
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   12960
         TabIndex        =   34
         Top             =   120
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   1995
      Left            =   0
      TabIndex        =   35
      Top             =   3120
      Width           =   14235
      _cx             =   25109
      _cy             =   3519
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"NotifyJobNeeded.frx":75BA5
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   1200
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
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
            Picture         =   "NotifyJobNeeded.frx":75C5E
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":75FF8
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":76392
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":7672C
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":76AC6
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":76E60
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":771FA
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NotifyJobNeeded.frx":77794
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   37
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
      ButtonImage     =   "NotifyJobNeeded.frx":77B2E
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid Fg 
      Height          =   1995
      Left            =   0
      TabIndex        =   50
      Top             =   5880
      Width           =   14235
      _cx             =   25109
      _cy             =   3519
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"NotifyJobNeeded.frx":7E390
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
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   615
         Left            =   1200
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   17520
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   6960
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
      ButtonImage     =   "NotifyJobNeeded.frx":7E5D7
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   15600
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   6960
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
      ButtonImage     =   "NotifyJobNeeded.frx":84E39
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
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
      TabIndex        =   38
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "NotifyJobNeeded"
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
Private Sub btnQuery_Click()
'Load FrmExpensespaidAdvancedSearch
'FrmExpensespaidAdvancedSearch.show vbModal
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblEmploymentNeed order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
  
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmpJobsTypes Me.DcbJob
    Dcombos.GetUsers Me.DCboUserName
   
 
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
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    
    StrSQL = "Delete From TblEmploymentNeedDet Where   EmpNedDet='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
   
    End If
    RsSavRec.Fields("RecordM").value = XPDtbTrans.value
    RsSavRec.Fields("RecordH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("JobID").value = val(Me.DcbJob.BoundText)
    RsSavRec.Fields("CuntJob").value = val(Me.TxtCuntJob.Text)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("Resons").value = TxtResons.Text

    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblEmploymentNeedDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Grid
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("jobName")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("EmpNedDet").value = Me.TxtSerial1.Text
                RsDevsub("JobID").value = IIf((.TextMatrix(i, .ColIndex("JobID"))) = "", Null, .TextMatrix(i, .ColIndex("JobID")))
                RsDevsub("CuntJob").value = IIf((.TextMatrix(i, .ColIndex("CuntJob"))) = "", Null, .TextMatrix(i, .ColIndex("CuntJob")))
                RsDevsub("Resons").value = IIf((.TextMatrix(i, .ColIndex("Resons"))) = "", Null, .TextMatrix(i, .ColIndex("Resons")))
                RsDevsub("Typ").value = 0
                
       RsDevsub.update
      End If
     Next i
    End With
    '''''''''''''''''''
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblEmploymentNeedDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    With FG
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("JobID")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("EmpNedDet").value = Me.TxtSerial1.Text
                RsDevsub("JobID").value = IIf((.TextMatrix(i, .ColIndex("JobID"))) = "", Null, .TextMatrix(i, .ColIndex("JobID")))
                RsDevsub("Name").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, .TextMatrix(i, .ColIndex("Name")))
                RsDevsub("NationalityID").value = IIf((.TextMatrix(i, .ColIndex("NationalityID"))) = "", Null, .TextMatrix(i, .ColIndex("NationalityID")))
                RsDevsub("Tel").value = IIf((.TextMatrix(i, .ColIndex("Tel"))) = "", Null, .TextMatrix(i, .ColIndex("Tel")))
                RsDevsub("Mobile").value = IIf((.TextMatrix(i, .ColIndex("Mobile"))) = "", Null, .TextMatrix(i, .ColIndex("Mobile")))
                RsDevsub("AdminMobile").value = IIf((.TextMatrix(i, .ColIndex("AdminMobile"))) = "", Null, .TextMatrix(i, .ColIndex("AdminMobile")))
                RsDevsub("Email").value = IIf((.TextMatrix(i, .ColIndex("Email"))) = "", Null, .TextMatrix(i, .ColIndex("Email")))
                RsDevsub("Qualifications").value = IIf((.TextMatrix(i, .ColIndex("Qualifications"))) = "", Null, .TextMatrix(i, .ColIndex("Qualifications")))
                RsDevsub("Experiences").value = IIf((.TextMatrix(i, .ColIndex("Experiences"))) = "", Null, .TextMatrix(i, .ColIndex("Experiences")))
                RsDevsub("LastSalary").value = IIf((.TextMatrix(i, .ColIndex("LastSalary"))) = "", Null, .TextMatrix(i, .ColIndex("LastSalary")))
                RsDevsub("ExpSalary").value = IIf((.TextMatrix(i, .ColIndex("ExpSalary"))) = "", Null, .TextMatrix(i, .ColIndex("ExpSalary")))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, .TextMatrix(i, .ColIndex("Total")))
                RsDevsub("Typ").value = 1
                
       RsDevsub.update
      End If
     Next i
    End With
 '''''''''''''///////////////////
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
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
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordM").value), Date, RsSavRec.Fields("RecordM").value): ProgressBar1.value = 20
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("RecordH").value), "", RsSavRec.Fields("RecordH").value): ProgressBar1.value = 30
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 40
    Me.TxtCuntJob.Text = IIf(IsNull(RsSavRec.Fields("CuntJob").value), "", RsSavRec.Fields("CuntJob").value): ProgressBar1.value = 50
    Me.DcbJob.BoundText = IIf(IsNull(RsSavRec.Fields("JobID").value), "", RsSavRec.Fields("JobID").value): ProgressBar1.value = 60
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 70
    Me.TxtResons.Text = IIf(IsNull(RsSavRec.Fields("Resons").value), "", RsSavRec.Fields("Resons").value): ProgressBar1.value = 80
    
         
    
    
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 80
     ' grid
     FillTextGridData
 ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
 

  Sub FillTextGridData()
'  If check3.value = True Or check5.value = True Then
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
  sql = "SELECT     dbo.TblEmploymentNeedDet.ID, dbo.TblEmploymentNeedDet.EmpNedDet, dbo.TblEmploymentNeedDet.Resons, dbo.TblEmploymentNeedDet.Typ, "
  sql = sql + "                    dbo.TblEmploymentNeedDet.CuntJob , dbo.TblEmploymentNeedDet.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee"
  sql = sql + " FROM         dbo.TblEmploymentNeedDet LEFT OUTER JOIN"
  sql = sql + "                     dbo.TblEmpJobsTypes ON dbo.TblEmploymentNeedDet.JobID = dbo.TblEmpJobsTypes.JobTypeID"
  sql = sql + " Where (dbo.TblEmploymentNeedDet.typ = 0) And (dbo.TblEmploymentNeedDet.EmpNedDet = " & val(TxtSerial1.Text) & ")"
  
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(Rs1("JobID").value), "", Rs1("JobID").value)
                   .TextMatrix(i, .ColIndex("CuntJob")) = IIf(IsNull(Rs1("CuntJob").value), 0, Rs1("CuntJob").value)
                    .TextMatrix(i, .ColIndex("Resons")) = IIf(IsNull(Rs1("Resons").value), "", Rs1("Resons").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("jobName")) = IIf(IsNull(Rs1("JobTypeName").value), "", Rs1("JobTypeName").value)
                     Else
                    .TextMatrix(i, .ColIndex("jobName")) = IIf(IsNull(Rs1("JobTypeNamee").value), "", Rs1("JobTypeNamee").value)
                    End If
       
                    Rs1.MoveNext
             Next i
        End With
'''''''''''''''
  Set Rs1 = New ADODB.Recordset
  
  sql = "SELECT     dbo.TblEmploymentNeedDet.ID, dbo.TblEmploymentNeedDet.EmpNedDet, dbo.TblEmploymentNeedDet.CuntJob, dbo.TblEmploymentNeedDet.Typ, "
  sql = sql + "                    dbo.TblEmploymentNeedDet.Resons, dbo.TblEmploymentNeedDet.Name, dbo.TblEmploymentNeedDet.Tel, dbo.TblEmploymentNeedDet.Mobile,"
  sql = sql + "                    dbo.TblEmploymentNeedDet.AdminMobile, dbo.TblEmploymentNeedDet.Email, dbo.TblEmploymentNeedDet.Qualifications, dbo.TblEmploymentNeedDet.Experiences,"
  sql = sql + "                    dbo.TblEmploymentNeedDet.LastSalary, dbo.TblEmploymentNeedDet.ExpSalary, dbo.TblEmploymentNeedDet.Total, dbo.TblEmpJobsTypes.JobTypeName,"
  sql = sql + "                    dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmploymentNeedDet.JobID, dbo.TblEmploymentNeedDet.NationalityID, dbo.Nationality.name AS Expr1,"
  sql = sql + "                    dbo.nationality.NameE"
  sql = sql + " FROM         dbo.TblEmploymentNeedDet LEFT OUTER JOIN"
  sql = sql + "                    dbo.Nationality ON dbo.TblEmploymentNeedDet.NationalityID = dbo.Nationality.id LEFT OUTER JOIN"
  sql = sql + "                    dbo.TblEmpJobsTypes ON dbo.TblEmploymentNeedDet.JobID = dbo.TblEmpJobsTypes.JobTypeID"
  sql = sql + " Where (dbo.TblEmploymentNeedDet.typ = 1) And (dbo.TblEmploymentNeedDet.EmpNedDet = " & val(TxtSerial1.Text) & ")"
  
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     
     With Me.FG
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(Rs1("Tel").value), "", Rs1("Tel").value)
                   .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(Rs1("Mobile").value), "", Rs1("Mobile").value)
                   .TextMatrix(i, .ColIndex("AdminMobile")) = IIf(IsNull(Rs1("AdminMobile").value), "", Rs1("AdminMobile").value)
                   .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(Rs1("Email").value), "", Rs1("Email").value)
                   .TextMatrix(i, .ColIndex("Qualifications")) = IIf(IsNull(Rs1("Qualifications").value), "", Rs1("Qualifications").value)
                    .TextMatrix(i, .ColIndex("Experiences")) = IIf(IsNull(Rs1("Experiences").value), "", Rs1("Experiences").value)
                      .TextMatrix(i, .ColIndex("LastSalary")) = IIf(IsNull(Rs1("LastSalary").value), "", Rs1("LastSalary").value)
                   .TextMatrix(i, .ColIndex("ExpSalary")) = IIf(IsNull(Rs1("ExpSalary").value), "", Rs1("ExpSalary").value)
                    .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), "", Rs1("Total").value)
                      .TextMatrix(i, .ColIndex("NationalityID")) = IIf(IsNull(Rs1("NationalityID").value), "", Rs1("NationalityID").value)
                   
                    
                      .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(Rs1("JobID").value), "", Rs1("JobID").value)
                  .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
                '    .TextMatrix(i, .ColIndex("Resons")) = IIf(IsNull(Rs1("Resons").value), "", Rs1("Resons").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("NationalityI")) = IIf(IsNull(Rs1("Expr1").value), "", Rs1("Expr1").value)
                   .TextMatrix(i, .ColIndex("jobName")) = IIf(IsNull(Rs1("JobTypeName").value), "", Rs1("JobTypeName").value)
                     Else
                    .TextMatrix(i, .ColIndex("NationalityI")) = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
                    .TextMatrix(i, .ColIndex("jobName")) = IIf(IsNull(Rs1("JobTypeName").value), "", Rs1("JobTypeName").value)
                    End If
       
                    Rs1.MoveNext
             Next i
        End With
    '    End If
    Exit Sub
 End Sub


Sub filgrid()
Dim k As Integer
Dim i As Integer
With Grid
k = .Rows
.Rows = .Rows + 1
Do While k < (.Rows)
.TextMatrix(k, .ColIndex("jobName")) = DcbJob.Text
.TextMatrix(k, .ColIndex("JobID")) = val(DcbJob.BoundText)
.TextMatrix(k, .ColIndex("CuntJob")) = val(TxtCuntJob.Text)
.TextMatrix(k, .ColIndex("Resons")) = TxtResons.Text
k = k + 1
Loop
DcbJob.BoundText = 0
TxtCuntJob.Text = 0
TxtResons.Text = ""
End With
End Sub

Sub RetriveDatt()
Dim last As Integer
Dim i As Integer
Dim Rs1 As ADODB.Recordset
Me.FG.Clear flexClearScrollable, flexClearEverything
FG.Rows = 1
Dim sql As String
For i = 1 To Grid.Rows - 1
If val(Grid.TextMatrix(i, Grid.ColIndex("JobID"))) <> 0 Then
sql = "SELECT     dbo.TblEmploymentModel.ID, dbo.TblEmploymentModel.RecordM, dbo.TblEmploymentModel.RecordH, dbo.TblEmploymentModel.BranchID, "
sql = sql & "                      dbo.TblEmploymentModel.Name1, dbo.TblEmploymentModel.Name2, dbo.TblEmploymentModel.Name3, dbo.TblEmploymentModel.Name4,"
sql = sql & "                    dbo.TblEmploymentModel.Name1 , dbo.TblEmploymentModel.Name2, dbo.TblEmploymentModel.Name3, dbo.TblEmploymentModel.Name4, "
sql = sql & "                      dbo.TblEmploymentModel.Name, dbo.TblEmploymentModel.Tel, dbo.TblEmploymentModel.Mobil, dbo.TblEmploymentModel.Email,"
sql = sql & "                      dbo.TblEmploymentModel.TelAdmin, dbo.TblEmploymentModel.Qualifications, dbo.TblEmploymentModel.Experiences, dbo.TblEmploymentModel.LastSalary,"
sql = sql & "                      dbo.TblEmploymentModel.ExpSalary, dbo.TblEmploymentModel.Total, dbo.Nationality.name AS Expr1, dbo.Nationality.namee, dbo.TblEmpJobsTypes.JobTypeName,"
sql = sql & "                      dbo.TblEmpJobsTypes.JobTypeNamee , dbo.TblEmploymentModel.NationalityID, dbo.TblEmploymentModel.SpecID,       dbo.TblEmploymentModel.JobID"
sql = sql & " FROM         dbo.TblEmploymentModel LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpJobsTypes ON dbo.TblEmploymentModel.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
sql = sql & "                      dbo.Nationality ON dbo.TblEmploymentModel.NationalityID = dbo.Nationality.id"
sql = sql & " where dbo.TblEmploymentModel.JobID =" & val(Grid.TextMatrix(i, Grid.ColIndex("JobID"))) & " "
Set Rs1 = New ADODB.Recordset
Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs1.RecordCount > 0 Then
With FG
last = .Rows
.Rows = .Rows + Rs1.RecordCount
Rs1.MoveFirst
Do While last < (.Rows)
.TextMatrix(last, .ColIndex("NationalityID")) = IIf(IsNull(Rs1("NationalityID").value), "", Rs1("NationalityID").value)
.TextMatrix(last, .ColIndex("Mobile")) = IIf(IsNull(Rs1("Mobil").value), "", Rs1("Mobil").value)
.TextMatrix(last, .ColIndex("Email")) = IIf(IsNull(Rs1("Email").value), "", Rs1("Email").value)
.TextMatrix(last, .ColIndex("AdminMobile")) = IIf(IsNull(Rs1("TelAdmin").value), "", Rs1("TelAdmin").value)
.TextMatrix(last, .ColIndex("Qualifications")) = IIf(IsNull(Rs1("Qualifications").value), "", Rs1("Qualifications").value)
.TextMatrix(last, .ColIndex("Experiences")) = IIf(IsNull(Rs1("Experiences").value), "", Rs1("Experiences").value)
.TextMatrix(last, .ColIndex("LastSalary")) = IIf(IsNull(Rs1("LastSalary").value), "", Rs1("LastSalary").value)
.TextMatrix(last, .ColIndex("ExpSalary")) = IIf(IsNull(Rs1("ExpSalary").value), "", Rs1("ExpSalary").value)
.TextMatrix(last, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), "", Rs1("Total").value)

.TextMatrix(last, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
.TextMatrix(last, .ColIndex("Tel")) = IIf(IsNull(Rs1("Tel").value), "", Rs1("Tel").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(last, .ColIndex("NationalityI")) = IIf(IsNull(Rs1("Expr1").value), "", Rs1("Expr1").value)
.TextMatrix(last, .ColIndex("jobName")) = IIf(IsNull(Rs1("JobTypeName").value), "", Rs1("JobTypeName").value)
Else
.TextMatrix(last, .ColIndex("NationalityI")) = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
.TextMatrix(last, .ColIndex("jobName")) = IIf(IsNull(Rs1("JobTypeNamee").value), "", Rs1("JobTypeNamee").value)
End If
.TextMatrix(last, .ColIndex("JobID")) = IIf(IsNull(Rs1("JobID").value), "", Rs1("JobID").value)

last = last + 1
Rs1.MoveNext
Loop
End With
End If
End If
Next i
End Sub
Private Sub ISButton2_Click()
 If DcbJob.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·ÊŸÌ›… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
                   
                 Else
            MsgBox "Please Select Job ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
            End If
            Exit Sub
            DcbJob.SetFocus
     End If
   '+++++++++++++++++++++++++++++++++++++++++++++++
      If val(TxtCuntJob.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ≈œŒ«· ⁄œœ «·ÊŸ«∆› ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            
                 Else
            MsgBox "Please Enter Count ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                   
            End If
         'slah   TxtCuntJob.SetFocus
            Exit Sub
     End If
     Dim j As Integer
     
    For j = 1 To Grid.Rows - 1
    If val(DcbJob.BoundText) = val(Grid.TextMatrix(j, Grid.ColIndex("JobID"))) Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «· ﬂ—«— Â–Â «·ÊŸÌ›… „ÊÃÊœ… „‰ ﬁ»·"
    Else
    MsgBox "This is Job Already Found"
    End If
    DcbJob.BoundText = 0
    Exit Sub
    End If
    Next j
     filgrid
     RetriveDatt
        '+++++++++++++++++++++++++++++++++++++++++++++++
End Sub
Private Sub ISButton3_Click()
On Error Resume Next
    With Me.FG
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub ISButton4_Click()
On Error Resume Next
Me.Grid.Clear flexClearScrollable, flexClearEverything
cleargriid
End Sub

Private Sub ISButton5_Click()
On Error Resume Next
    With Me.Grid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub

Private Sub ISButton6_Click()
On Error Resume Next
Me.Grid.Clear flexClearScrollable, flexClearEverything
cleargriid
End Sub
Private Sub Txt_DateHigri_LostFocus()
  VBA.Calendar = vbCalGreg
            XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.Text <> "R" Then
              Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
   End If
   End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Arabic Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Dcbranch.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++



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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblEmploymentNeed", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
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
' refrsh sub
Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
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
    Dim x As Integer
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
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TblEmploymentNeedDet Where EmpNedDet='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               cleargriid
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
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
'Private Sub Grid_EnterCell()
 '   On Error GoTo ErrTrap
  '  FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("Ser")))
'ErrTrap:
'End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        
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
        Grid.Enabled = True
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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
    cleargriid
    TxtModFlg.Text = "N"
  
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = branch_id
  Me.FG.Clear flexClearScrollable, flexClearEverything
 
    Me.Grid.Clear flexClearScrollable, flexClearEverything
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
' print Events
'++++++++++++++++++++++++++++++++++++++++++

Private Sub ISButton1_Click()
On Error GoTo ErrTrap
   If val(Me.TxtSerial1.Text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
   sql = "SELECT    dbo.TbExpensespaidAdvanced.IDEXP, dbo.TbExpensespaidAdvanced.DateM, dbo.TbExpensespaidAdvanced.DateH, dbo.TbExpensespaidAdvanced.BranchID, TblBranchesData_2.branch_name,"
   sql = sql & "      TblBranchesData_2.branch_namee, dbo.TbExpensespaidAdvanced.PayWay, dbo.TbExpensespaidAdvanced.Explan, dbo.TbExpensespaidAdvanced.ExpIDD,"
   sql = sql & "     dbo.TbExpensespaidAdvanced.ExpName, dbo.TbExpensesprovided.name, dbo.TbExpensesprovided.namee, dbo.TbExpensespaidAdvanced.ExpAcount, ACCOUNTS_1.Account_Name,"
   sql = sql & "     ACCOUNTS_1.Account_NameEng, dbo.TbExpensespaidAdvanced.ExpAcount1, ACCOUNTS_1.Account_Name AS Account_Name1, ACCOUNTS_1.Account_NameEng AS Account_Name1E,"
   sql = sql & "      dbo.TbExpensespaidAdvanced.ExpSingle, mofrdat_2.mofrad_name, mofrdat_2.mofrad_namee, dbo.TbExpensespaidAdvanced.EXPCheck, dbo.TbExpensespaidAdvanced.ExpValue,"
   sql = sql & "     dbo.TbExpensespaidAdvanced.ExpMonth, dbo.TbExpensespaidAdvanced.ExpYear, dbo.TbExpensespaidAdvanced.ExpNumber, dbo.TbExpensespaidAdvanced.ExpEmpCheck,"
   sql = sql & "      dbo.TbExpensespaidAdvanced.ExpEmpSelect, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
   sql = sql & "       dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1,"
   sql = sql & "     dbo.TblEmployee.Emp_Namee, dbo.TbExpensespaidAdvanced.ExpBourchSelect, TblBranchesData_1.branch_name AS branch_nameSelect,"
   sql = sql & "      TblBranchesData_1.branch_namee AS branch_nameSelectE, dbo.TbExpensespaidAdvanced.ExpMangemtSelect, TblEmpDepartments_1.DepartmentName,"
   sql = sql & "     TblEmpDepartments_1.DepartmentNamee, dbo.TbExpensespaidAdvanced.ExpSingleSelect, mofrdat_1.mofrad_name AS mofrad_nameSelct, mofrdat_1.mofrad_namee AS mofrad_nameSelctE,"
   sql = sql & "     dbo.TbExpensespaidJoin.EmpID, TblEmployee_1.Emp_Name AS Emp_NameDet, TblEmployee_1.Emp_Name1 AS Emp_NameDet1, TblEmployee_1.Emp_Name2 AS Emp_NameDet2,"
   sql = sql & "     TblEmployee_1.Emp_Name3 AS Emp_NameDet3, TblEmployee_1.Emp_Name4 AS Emp_NameDet4, TblEmployee_1.Fullcode AS FullcodeDet, TblEmployee_1.Emp_Namee4 AS Emp_NameeDet4,"
   sql = sql & "     TblEmployee_1.Emp_Namee3 AS Emp_NameeDet3, TblEmployee_1.Emp_Namee2 AS Emp_NameeDet2, TblEmployee_1.Emp_Namee1 AS Emp_NameeDet1,"
   sql = sql & "      TblEmployee_1.Emp_Namee AS Emp_NameeDet, dbo.TbExpensespaidJoin.BranchID AS BranchIDDet, TblBranchesData_2.branch_name AS branch_nameDet,"
   sql = sql & "      TblBranchesData_2.branch_namee AS branch_nameDetE, dbo.TbExpensespaidJoin.MangmentID, TblEmpDepartments_1.DepartmentName AS DepartmentNameDet,"
   sql = sql & "     TblEmpDepartments_1.DepartmentNamee AS DepartmentNameeDet, dbo.TbExpensespaidJoin.Single, mofrdat_2.mofrad_name AS mofrad_nameDet, mofrdat_2.mofrad_namee AS mofrad_nameDetE,"
   sql = sql & "      dbo.TbExpensespaidJoin.SingleValue, dbo.TbExpensespaidJoin.PayType, dbo.TbExpensespaidJoin.Monthe, dbo.TbExpensespaidJoin.SubYear, dbo.TbExpensespaidJoin.PayValue,"
   sql = sql & "      dbo.TbExpensespaidJoin.id , dbo.TbExpensespaidAdvanced.MofrdCheck, dbo.TbExpensespaidAdvanced.TxtSearchCode"
   sql = sql & "        FROM         dbo.TblEmployee RIGHT OUTER JOIN"
   sql = sql & "     dbo.TbExpensespaidAdvanced LEFT OUTER JOIN"
   sql = sql & "      dbo.mofrdat mofrdat_2 RIGHT OUTER JOIN"
   sql = sql & "      dbo.TbExpensespaidJoin ON mofrdat_2.mofrad_code = dbo.TbExpensespaidJoin.Single LEFT OUTER JOIN"
   sql = sql & "      dbo.TblEmpDepartments TblEmpDepartments_1 ON dbo.TbExpensespaidJoin.MangmentID = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
   sql = sql & "      dbo.TblBranchesData TblBranchesData_2 ON dbo.TbExpensespaidJoin.BranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
   sql = sql & "     dbo.TblEmployee TblEmployee_1 ON dbo.TbExpensespaidJoin.EmpID = TblEmployee_1.Emp_ID ON dbo.TbExpensespaidAdvanced.IDEXP = dbo.TbExpensespaidJoin.IDEXP ON"
   sql = sql & "      dbo.TblEmployee.Emp_ID = dbo.TbExpensespaidAdvanced.ExpEmpSelect LEFT OUTER JOIN"
   sql = sql & "      dbo.mofrdat mofrdat_1 ON dbo.TbExpensespaidAdvanced.ExpSingleSelect = mofrdat_1.mofrad_code LEFT OUTER JOIN"
   sql = sql & "      dbo.TblEmpDepartments TblEmpDepartments_2 ON dbo.TbExpensespaidAdvanced.ExpMangemtSelect = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
   sql = sql & "     dbo.TblBranchesData TblBranchesData_1 ON dbo.TbExpensespaidAdvanced.ExpBourchSelect = TblBranchesData_1.branch_id LEFT OUTER JOIN"
   sql = sql & "     dbo.mofrdat mofrdat_3 ON dbo.TbExpensespaidAdvanced.ExpSingle = mofrdat_3.mofrad_code LEFT OUTER JOIN"
   sql = sql & "     dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TbExpensespaidAdvanced.ExpAcount1 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
   sql = sql & "      dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TbExpensespaidAdvanced.ExpAcount = ACCOUNTS_2.Account_Code LEFT OUTER JOIN"
   sql = sql & "      dbo.TbExpensesprovided ON dbo.TbExpensespaidAdvanced.ExpName = dbo.TbExpensesprovided.ID LEFT OUTER JOIN"
   sql = sql & "     dbo.TblBranchesData TblBranchesData_3 ON dbo.TbExpensespaidAdvanced.BranchID = TblBranchesData_3.branch_id"
   sql = sql & " Where (dbo.TbExpensespaidAdvanced.IDEXP = " & val(TxtSerial1.Text) & ")"
                    

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ExpensespaidAdvancedRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ExpensespaidAdvancedRPTEE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
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
ErrTrap:
  End Function
' chang langeg Event
'++++++++++++++++++++++++++++++++++++
'Private Sub TxtVacName_GotFocus()
 '   SwitchKeyboardLang LANG_ARABIC
'End Sub
'Private Sub TxtVacNamee_GotFocus()
'SwitchKeyboardLang LANG_ENGLISH
'End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
       Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    
    Me.Caption = "Jobs Requirements"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Operation ID"
    Me.lbl(2).Caption = "Date"
    Me.lbl(1).Caption = "HJ Date"
    Me.Label3.Caption = "Branch"
    Me.lbl(15).Caption = "Wanted Job"
    Me.Label1(3).Caption = "Wanted NO."
    Me.lbl(0).Caption = "Wanted Cos"
    Me.ISButton2.Caption = "OK"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton3.Caption = "Delet Select"
    ISButton5.Caption = "Delet Select"
    ISButton4.Caption = "Delet All"
    ISButton6.Caption = "Delet All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("jobName")) = "job Name"
        .TextMatrix(0, .ColIndex("CuntJob")) = "job No."
        .TextMatrix(0, .ColIndex("Resons")) = "Wanted Cos"
    End With
    
    With Me.FG
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("NationalityI")) = "Nationality"
        .TextMatrix(0, .ColIndex("jobName")) = "job Name"
        .TextMatrix(0, .ColIndex("Tel")) = "Tel"
        .TextMatrix(0, .ColIndex("Mobile")) = "Mobile"
        .TextMatrix(0, .ColIndex("AdminMobile")) = "Admin Mobile"
        .TextMatrix(0, .ColIndex("Email")) = "Email"
        .TextMatrix(0, .ColIndex("LastSalary")) = "Last Salary"
        .TextMatrix(0, .ColIndex("ExpSalary")) = "Expected Salary"
         .TextMatrix(0, .ColIndex("Qualifications")) = "Qualifications"
        .TextMatrix(0, .ColIndex("Experiences")) = "Experiences"
        .TextMatrix(0, .ColIndex("Total")) = "Reputation"
    End With
    
ErrTrap:
End Sub
Private Sub cleargriid()
Me.Grid.Rows = 1
Me.FG.Rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblEmploymentNeed"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

'+++++++++++++++++++++++++++++++++ end






