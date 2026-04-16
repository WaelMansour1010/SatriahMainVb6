VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmSallingPlan 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ŒÿÂ  ”⁄Ì— «·«’‰«›"
   ClientHeight    =   9030
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   16905
   HelpContextID   =   580
   Icon            =   "FrmSallingPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   16905
   Visible         =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16965
      _cx             =   29924
      _cy             =   15849
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
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      GridRows        =   10
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmSallingPlan.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1770
         Left            =   30
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   16890
         _cx             =   29792
         _cy             =   3122
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
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   1080
            Visible         =   0   'False
            Width           =   4575
            Begin MSComCtl2.DTPicker dbFromDate 
               Height          =   270
               Left            =   2145
               TabIndex        =   69
               Top             =   240
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   476
               _Version        =   393216
               Format          =   173473793
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker dbTodate 
               Height          =   270
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   476
               _Version        =   393216
               Format          =   173473793
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ï"
               Height          =   270
               Index           =   2
               Left            =   1545
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   240
               Width           =   360
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„œ Â« „‰"
               Height          =   270
               Index           =   5
               Left            =   3615
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   240
               Width           =   705
            End
         End
         Begin VB.OptionButton Optfixedintrval 
            Alignment       =   1  'Right Justify
            Caption         =   "„ÕœœÂ"
            Height          =   195
            Index           =   1
            Left            =   15120
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Optfixedintrval 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— „ÕœœÂ «·„œÂ"
            Height          =   195
            Index           =   0
            Left            =   15000
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox TxtPlanID 
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
            Height          =   270
            Left            =   15060
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   840
            Width           =   1200
         End
         Begin VB.CheckBox ChkLocked 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ìﬁ«› «· ⁄«„·"
            Height          =   210
            Left            =   17700
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1395
            Width           =   2295
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
            Height          =   750
            Left            =   0
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   795
            Width           =   9090
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   765
            Index           =   5
            Left            =   0
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   16875
            _cx             =   29766
            _cy             =   1349
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   24
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
            ForeColor       =   4210688
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Picture         =   "FrmSallingPlan.frx":044F
            Caption         =   "ŒÿÂ  ”⁄Ì—«·«’‰«› "
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   0
            ChildSpacing    =   0
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   0
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
            Begin VB.TextBox TxtModFlg 
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
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Text            =   "Text6"
               Top             =   120
               Visible         =   0   'False
               Width           =   615
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   375
               Index           =   0
               Left            =   1695
               TabIndex        =   19
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmSallingPlan.frx":1129
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   375
               Index           =   2
               Left            =   630
               TabIndex        =   20
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmSallingPlan.frx":14C3
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   375
               Index           =   1
               Left            =   2220
               TabIndex        =   21
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmSallingPlan.frx":185D
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   375
               Index           =   3
               Left            =   1155
               TabIndex        =   22
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmSallingPlan.frx":1BF7
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   17955
            TabIndex        =   26
            Top             =   -45
            Width           =   3255
            _ExtentX        =   5741
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
         Begin MSComCtl2.DTPicker DPRecorddate 
            Height          =   270
            Left            =   12600
            TabIndex        =   29
            Top             =   840
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   476
            _Version        =   393216
            Format          =   148504577
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·⁄„·ÌÂ"
            Height          =   225
            Index           =   9
            Left            =   13920
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„"
            Height          =   225
            Index           =   7
            Left            =   16320
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   840
            Width           =   210
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   315
            Index           =   3
            Left            =   9090
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   915
            Width           =   705
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6240
         Left            =   30
         TabIndex        =   1
         Top             =   1815
         Width           =   16890
         _cx             =   29792
         _cy             =   11007
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
         Caption         =   "ŒÿÂ «·«”⁄«—|⁄—Ê÷ Œ«’Â"
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
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   5820
            Index           =   0
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   16800
            _cx             =   29633
            _cy             =   10266
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
            Begin VB.Frame Frame5 
               Height          =   1155
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   360
               Width           =   4545
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«”⁄«— ÌœÊÏ"
                  Height          =   405
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«”⁄«— ¬·Ì"
                  Height          =   405
                  Left            =   2970
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   1125
               End
               Begin VB.CommandButton Command2 
                  Caption         =   " Õ„Ì· «·„·›..."
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   1695
               End
               Begin VB.CommandButton cmdSelectFile 
                  Caption         =   " ÕœÌœ «·„·›..."
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   1545
               End
            End
            Begin VB.TextBox txtFile 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   96
               Top             =   0
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtItemsIDes 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1710
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   -60
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.CheckBox chkIsNewPric 
               Alignment       =   1  'Right Justify
               Caption         =   " ”⁄Ì— ÃœÌœ"
               Height          =   375
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   30
               Width           =   1305
            End
            Begin VB.TextBox TXTOrderNo 
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
               Height          =   315
               Left            =   11880
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   1635
               Width           =   2400
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Õ‰Â „⁄Ì‰Â"
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   15240
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   1710
               Width           =   1320
            End
            Begin VB.ComboBox DCTransactionType 
               Height          =   315
               ItemData        =   "FrmSallingPlan.frx":1F91
               Left            =   13440
               List            =   "FrmSallingPlan.frx":1FA5
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   0
               Width           =   1815
            End
            Begin VB.TextBox txtvalueOrPercentageValue 
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
               Height          =   270
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   4680
               Width           =   1200
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·ﬂ· «·«’‰«›"
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   15360
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1080
               Width           =   1200
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ã„Ê⁄Â „Õœœ"
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   14985
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1320
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.Frame Frame1 
               Caption         =   "Õœœ «·›—⁄"
               Height          =   615
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   360
               Width           =   6015
               Begin VB.OptionButton OptBranch 
                  Alignment       =   1  'Right Justify
                  Caption         =   "·ﬂ· «·›—Ê⁄"
                  Height          =   210
                  Index           =   0
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton OptBranch 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›—⁄ „Õœœ"
                  Height          =   210
                  Index           =   1
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1215
               End
               Begin MSDataListLib.DataCombo dcBranch 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   42
                  Top             =   240
                  Width           =   2760
                  _ExtentX        =   4868
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
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„” ‰œ „⁄Ì‰"
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   15480
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   0
               Width           =   1080
            End
            Begin VB.TextBox TxtInvSerial 
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
               Height          =   315
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   75
               Width           =   1800
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«’‰«› „ÕœœÂ"
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   10425
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   1260
               Width           =   1320
            End
            Begin VB.Frame Frame2 
               Caption         =   "Õœœ «·ÊÕœ« "
               Height          =   615
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   360
               Width           =   5775
               Begin VB.OptionButton optUnits 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÊÕœÂ „ÕœœÂ"
                  Height          =   210
                  Index           =   1
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.OptionButton optUnits 
                  Alignment       =   1  'Right Justify
                  Caption         =   "·ﬂ· «·ÊÕœ« "
                  Height          =   210
                  Index           =   0
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1695
               End
               Begin MSDataListLib.DataCombo DcboUnits 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1920
                  _ExtentX        =   3387
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
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   7
               Left            =   5265
               TabIndex        =   45
               Top             =   1530
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
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
               ButtonImage     =   "FrmSallingPlan.frx":1FDC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   8
               Left            =   4080
               TabIndex        =   46
               Top             =   1530
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmSallingPlan.frx":2376
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcGroup 
               Height          =   315
               Left            =   11865
               TabIndex        =   47
               Top             =   1320
               Width           =   3000
               _ExtentX        =   5292
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
            Begin MSDataListLib.DataCombo dcitems 
               Height          =   315
               Left            =   7080
               TabIndex        =   48
               Top             =   1260
               Width           =   3000
               _ExtentX        =   5292
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   1110
               Left            =   0
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   4530
               Width           =   16785
               _cx             =   29607
               _cy             =   1958
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
               Begin VB.CommandButton CMDDO 
                  Caption         =   "‰›– «·ŒÿÂ"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.CommandButton cmdOperator 
                  Caption         =   "/"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   15.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   3
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   120
                  Width           =   615
               End
               Begin VB.CommandButton cmdOperator 
                  Caption         =   "*"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   15.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   2
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   120
                  Width           =   615
               End
               Begin VB.CommandButton cmdOperator 
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   15.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   1
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   120
                  Width           =   615
               End
               Begin VB.Frame Frame3 
                  Caption         =   "Õœœ ”⁄— «·–Ì Ì „ ⁄·ÌÂ «· €ÌÌ—Â "
                  Height          =   615
                  Left            =   12360
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   0
                  Width           =   4455
                  Begin VB.OptionButton optFixedPrice 
                     Alignment       =   1  'Right Justify
                     Caption         =   "”⁄— „Õœœ"
                     Height          =   210
                     Index           =   1
                     Left            =   2040
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.OptionButton optFixedPrice 
                     Alignment       =   1  'Right Justify
                     Caption         =   "·ﬂ· «·«”⁄«—"
                     Height          =   210
                     Index           =   0
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   240
                     Width           =   1095
                  End
                  Begin MSDataListLib.DataCombo dcSalePriceNames 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   65
                     Top             =   240
                     Width           =   1920
                     _ExtentX        =   3387
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
               End
               Begin VB.ComboBox cbovalueOrPercentage 
                  Height          =   315
                  ItemData        =   "FrmSallingPlan.frx":2910
                  Left            =   2160
                  List            =   "FrmSallingPlan.frx":291A
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.CommandButton cmdOperator 
                  Caption         =   "+"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   15.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   0
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   120
                  Width           =   615
               End
               Begin VB.TextBox txtAnotherPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   270
                  Left            =   9480
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   480
                  Width           =   1200
               End
               Begin VB.ComboBox cbopriceChangeId 
                  Height          =   315
                  ItemData        =   "FrmSallingPlan.frx":2929
                  Left            =   7920
                  List            =   "FrmSallingPlan.frx":2939
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   120
                  Width           =   2775
               End
               Begin VB.Label lblOperator 
                  Alignment       =   2  'Center
                  Caption         =   "+"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   375
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   120
                  Width           =   495
               End
               Begin VB.Label LblPercentage 
                  Alignment       =   1  'Right Justify
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„Â / ‰”»Â"
                  Height          =   225
                  Index           =   8
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   120
                  Width           =   930
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— „Œ ·›"
                  Height          =   225
                  Index           =   1
                  Left            =   10740
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   480
                  Width           =   1050
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„Õœœ«   €ÌÌ— «·”⁄—"
                  Height          =   375
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   120
                  Width           =   1455
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   1800
               TabIndex        =   81
               Top             =   1530
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ﬂ· «·”ÿÊ—"
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
               ButtonImage     =   "FrmSallingPlan.frx":297E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin C1SizerLibCtl.C1Tab TabMain 
               Height          =   2535
               Left            =   0
               TabIndex        =   87
               Top             =   1980
               Width           =   16770
               _cx             =   29580
               _cy             =   4471
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
               Appearance      =   2
               MousePointer    =   0
               Version         =   801
               BackColor       =   12648447
               ForeColor       =   -2147483630
               FrontTabColor   =   14871017
               BackTabColor    =   12648447
               TabOutlineColor =   -2147483632
               FrontTabForeColor=   16711680
               Caption         =   " ”⁄Ì— 1| ”⁄Ì— 2"
               Align           =   0
               CurrTab         =   1
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
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   2160
                  Index           =   2
                  Left            =   -17325
                  TabIndex        =   88
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   16680
                  _cx             =   29422
                  _cy             =   3810
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
                  Begin VSFlex8UCtl.VSFlexGrid FgItems 
                     Height          =   2145
                     Index           =   0
                     Left            =   23250
                     TabIndex        =   89
                     Top             =   135
                     Width           =   16470
                     _cx             =   29051
                     _cy             =   3784
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
                     BackColorFixed  =   -2147483633
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
                     AllowUserResizing=   0
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   50
                     Cols            =   5
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmSallingPlan.frx":2F18
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
                  Begin VSFlex8Ctl.VSFlexGrid Grid1 
                     Height          =   2010
                     Left            =   -46020
                     TabIndex        =   92
                     Top             =   90
                     Width           =   62910
                     _cx             =   110966
                     _cy             =   3545
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
                     SelectionMode   =   1
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   1
                     Cols            =   39
                     FixedRows       =   1
                     FixedCols       =   2
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmSallingPlan.frx":2FD8
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
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   2160
                  Index           =   4
                  Left            =   45
                  TabIndex        =   90
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   16680
                  _cx             =   29422
                  _cy             =   3810
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
                  Begin VSFlex8UCtl.VSFlexGrid FgItems 
                     Height          =   2145
                     Index           =   1
                     Left            =   23220
                     TabIndex        =   91
                     Top             =   180
                     Width           =   16440
                     _cx             =   28998
                     _cy             =   3784
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
                     BackColorFixed  =   -2147483633
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
                     AllowUserResizing=   0
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   50
                     Cols            =   5
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmSallingPlan.frx":35D2
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
                  Begin VSFlex8Ctl.VSFlexGrid Grid2 
                     Height          =   2010
                     Left            =   60
                     TabIndex        =   93
                     Top             =   90
                     Width           =   16530
                     _cx             =   29157
                     _cy             =   3545
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
                     SelectionMode   =   1
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   17
                     FixedRows       =   1
                     FixedCols       =   2
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmSallingPlan.frx":3692
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
                  End
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·‘Õ‰Â"
               Height          =   225
               Index           =   10
               Left            =   14220
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   1755
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·„” ‰œ"
               Height          =   225
               Index           =   0
               Left            =   12540
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   75
               Width           =   810
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   5820
            Index           =   1
            Left            =   17535
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   45
            Width           =   16800
            _cx             =   29633
            _cy             =   10266
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
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   1275
               Left            =   7800
               TabIndex        =   83
               Top             =   720
               Width           =   8865
               _cx             =   15637
               _cy             =   2249
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
               Rows            =   50
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSallingPlan.frx":398D
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
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   1515
               Left            =   7800
               TabIndex        =   85
               Top             =   3000
               Width           =   8865
               _cx             =   15637
               _cy             =   2672
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
               Rows            =   50
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSallingPlan.frx":3A46
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·’‰› «·„÷«›/«·„Ã«‰Ì"
               Height          =   315
               Index           =   12
               Left            =   14160
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   2400
               Width           =   2145
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’‰› «·«”«”"
               Height          =   315
               Index           =   11
               Left            =   14880
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   360
               Width           =   1305
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   885
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   8070
         Width           =   16905
         _cx             =   29819
         _cy             =   1561
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   11880
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   90
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
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
            ButtonImage     =   "FrmSallingPlan.frx":3AFF
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
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
            ButtonImage     =   "FrmSallingPlan.frx":3E99
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSallingPlan.frx":4233
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   11700
            TabIndex        =   9
            Top             =   510
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   10800
            TabIndex        =   10
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   9960
            TabIndex        =   11
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   3
            Left            =   8955
            TabIndex        =   12
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   7920
            TabIndex        =   13
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   6
            Left            =   480
            TabIndex        =   14
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   5
            Left            =   6990
            TabIndex        =   15
            Top             =   510
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   9120
            TabIndex        =   16
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmSallingPlan.frx":45CD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Status 
            Alignment       =   1  'Right Justify
            Height          =   435
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   30
            Width           =   7155
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
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
      ButtonImage     =   "FrmSallingPlan.frx":45E9
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmSallingPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim isFromExcel As Boolean
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Dim mIndexTab As Integer
Dim mGrd As Object
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long




Private Sub cmdSelectFile_Click()

CD1.ShowOpen
txtFile.Text = CD1.filename
End Sub

Private Sub Command2_Click()

  FillItem

End Sub

Sub FillItem()
Dim error_string  As String
  error_string = ""
If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    Dim currentvalue As String, mDesc As String
    Dim Name As String
    Dim NameE As String
    Dim itemcode As String
    Dim ITEMPRICE As Double
    Dim itemDisc As Double
    Dim UnitName As String
    Dim GroupName As String
    Dim mEqu As String
    Dim des As String
    Dim DebitValue As String
    Dim CreditValue As String
   Grid2.Rows = 1
    Set ExcelObj = CreateObject("Excel.Application")
'        Set ExcelSheet = Nothing
'    Set ExcelBook = Nothing
'    Set ExcelObj = Nothing
'
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
    isFromExcel = True
    With ExcelSheet
    i = 2
    
    Do Until .cells(i, 1) & "" = ""
        itemcode = .cells(i, 1)
        GroupName = .cells(i, 2)
        NameE = .cells(i, 3)
        Name = .cells(i, 4)
        ITEMPRICE = val(.cells(i, 5) & "")
        UnitName = .cells(i, 6)
        
        itemDisc = 0
        
    'mDesc = .cells(i, 5)
 addrow2 itemcode, Name, UnitName, ITEMPRICE, itemDisc
       
       Status.Caption = "«·„”·”· :" & i & "«·’‰› : " & Name & " ”⁄—Â :" & ITEMPRICE
       i = i + 1
     '  NewGrid.CountItems
    Loop
        End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

        If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
       isFromExcel = False
       Me.Grid2.Rows = Me.Grid2.Rows + 1
       MsgBox " „ «·«œ—«Ã"
'GetNotinGard
'Coloring
End Sub


Function addrow2(Fullcode As String, Name As String, UnitName As String, ITEMPRICE As Double, Optional itemDisc As Double, Optional des As String)

    Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim UnitID As Double
    Dim LngItemID As Long
    Dim LngUnitID As Long
    Dim ColorID As Integer
    Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
 Dim error_string As String
    Dim Price As Double
    Dim i As Long
  '  UnitID = GetUnitID(Name)
    
                       
                'lllllllllllllll
                
                
                
  
  Dim s As String
  Dim rsDummy As New ADODB.Recordset
  Dim RsUnit As New ADODB.Recordset
   If Name <> "" Then
       s = "Select * from tblItems Where Fullcode Like '" & Trim(Fullcode) & "' Or barCodeNO Like '" & Trim(Fullcode) & "' Or ItemName Like '" & Trim(Name) & "' "
       rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If rsDummy.EOF Then
            Exit Function
        Else
            LngItemID = val(rsDummy!ItemID & "")
        End If
        
    If LngItemID <> 0 Then
    Dim mRow As Long
    
    With Me.Grid2
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("Item_id")) = LngItemID
        .Row = .Rows - 1
            
        .TextMatrix(.Rows - 1, .ColIndex("Item_code")) = rsDummy!itemcode & ""
        .TextMatrix(.Rows - 1, .ColIndex("Item_name")) = IIf(IsNull(rsDummy.Fields("ItemName").value), "", rsDummy.Fields("ItemName").value)
        
        
            s = "Select UnitID,UnitName From TblUnites Where UnitName Like '" & Trim(UnitName) & "'"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open s, Cn, adOpenStatic
            If Not RsUnit.EOF Then
                LngUnitID = val(RsUnit!UnitID & "")
            Else
                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitWholeSalePrice "
                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                Set RsUnit = New ADODB.Recordset
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            End If
               

            If Not RsUnit.EOF Then
                
                .TextMatrix(.Rows - 1, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                .TextMatrix(.Rows - 1, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
                .TextMatrix(.Rows - 1, .ColIndex("SalePrice")) = IIf(IsNull(RsUnit.Fields("UnitWholeSalePrice").value), "", RsUnit.Fields("UnitWholeSalePrice").value)
            End If

            RsUnit.Close
        
        
        .TextMatrix(.Rows - 1, .ColIndex("SalePriceNew")) = ITEMPRICE
    '    .TextMatrix(.Rows - 1, .ColIndex("Discount")) = itemDisc
        
        
        .Row = .Rows - 1
  
         .TextMatrix(.Rows - 1, .ColIndex("Ser")) = .Rows - 1
        
        

      

     End With
    '      Me.TxtItemCodeB.Text = ""
     
    '\      Unload FrmItemSearch2
     ' Me.TxtItemCodeB.SetFocus
         
    Else
         
    End If
    
    Else
           error_string = error_string & Trim(Fullcode) & "," & ITEMPRICE & "," & Name & vbCrLf

End If
'End If

End Function

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub chkIsNewPric_Click()
    Dim s As String
    If chkIsNewPric.value = vbChecked Then
        Frame5.Visible = True
        cmdSelectFile.Visible = True
        Command2.Visible = True
        TabMain.CurrTab = 1
        mIndexTab = 1
        Set mGrd = Grid2
        s = " SELECT  id,PriceName From TblSalePriceNames"
        s = s & " WHERE PriceName LIKE '%”⁄—  Ã“∆…%' OR PriceName LIKE '%”⁄— Ã„·Â%'"
        
        fill_combo dcSalePriceNames, s
        Grid1.Rows = 1
        optUnits(0).value = True
    Else
        Frame5.Visible = False
       cmdSelectFile.Visible = False
        Command2.Visible = False
        TabMain.CurrTab = 0
        mIndexTab = 0
        Set mGrd = Grid1
        Dim Dcombos As ClsDataCombos
    
        Set Dcombos = New ClsDataCombos
        Dcombos.GetSalePriceNames dcSalePriceNames
        Grid2.Rows = 1
    End If
    
    
End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

End Function

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
 
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
        If OptBranch(1).value = True Then
 
            If Dcbranch.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ    «·›—⁄ «Ê·«  "
                Else
                    Msg = "Specify   Branch Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
 
        End If
 
        If optUnits(1).value = True Then
 
            If DcboUnits.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ    «·ÊÕœ… «Ê·«  "
                Else
                    Msg = "Specify   Unit Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboUnits.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
 
        End If
        If chkIsNewPric.value = vbUnchecked Then
            If cbopriceChangeId.ListIndex = -1 Then
          
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ    „ÕÕœœ«  «·”⁄—  «Ê·«  "
                Else
                    Msg = "Specify   Price To Change  Firstly"
                End If
    
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                cbopriceChangeId.SetFocus
                SendKeys "{F4}"
                Exit Sub
     
            End If
     
            If cbovalueOrPercentage.ListIndex = -1 Then
          
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ       «· €ÌÌ— »ﬁÌ„Â «„ ‰”»Â  «Ê·«  "
                Else
                    Msg = "Specify    value Or Percentage      Firstly"
                End If
    
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                cbovalueOrPercentage.SetFocus
                SendKeys "{F4}"
                Exit Sub
     
            End If
     
            If cbovalueOrPercentage.ListIndex = 0 And val(txtvalueOrPercentageValue.Text) = 0 Then
          
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ    ﬁÌ„Â «· €ÌÌ—   «Ê·«  "
                Else
                    Msg = "Specify    value      Firstly"
                End If
    
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                txtvalueOrPercentageValue.SetFocus
                 
                Exit Sub
     
            End If
     
            If cbovalueOrPercentage.ListIndex = 1 And val(txtvalueOrPercentageValue.Text) = 0 Then
          
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ    ‰”»… «· €ÌÌ—   «Ê·«  "
                Else
                    Msg = "Specify    PECENTAGE      Firstly"
                End If
    
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                txtvalueOrPercentageValue.SetFocus
                 
                Exit Sub
     
            End If
     
            If cbopriceChangeId.ListIndex = 3 And val(txtAnotherPrice.Text) = 0 Then
          
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ       «·”⁄— «·„Œ ·›    «Ê·«  "
                Else
                    Msg = "Specify    PECENTAGE      Firstly"
                End If
    
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                txtAnotherPrice.SetFocus
                 
                Exit Sub
     
            End If
            
        End If
    End If
    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblSalesPricesPlanDetails where PlanID=" & val(Me.TxtPlanID.Text)
        Cn.Execute "delete TblSalesPricesPlanDetails2 where PlanID=" & val(Me.TxtPlanID.Text)
   
    End If
    
    rs("PlanID").value = val(TxtPlanID.Text)
    '    rs("CustomerId").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
    rs("RecordDate").value = DPRecorddate.value

    If Optfixedintrval(1).value = True Then
        rs("FixedInterval").value = 1
    Else
        rs("FixedInterval").value = 0
    End If
    

    If chkIsNewPric.value = vbChecked Then
        rs("IsNewPric").value = 1
    Else
        rs("IsNewPric").value = 0
    End If
    

    rs("IntervalFrom").value = dbFromDate.value
    rs("intervalto").value = dbTodate.value
    rs("Remarks").value = IIf(Me.TxtRemarks.Text = "", "", Me.TxtRemarks.Text)
    rs("OrderNo").value = IIf(Me.txtOrderNo.Text = "", "", Me.txtOrderNo.Text)
    rs("InvSerial").value = IIf(Me.TxtInvSerial.Text = "", "", Me.TxtInvSerial.Text)
    rs("TransactionType").value = DCTransactionType.ListIndex
    rs("GroupID").value = IIf(Me.DCGroup.BoundText = "", Null, Me.DCGroup.BoundText)
     
    If Opt(0).value = True Then
        rs("Plantype").value = 0
    ElseIf Opt(1).value = True Then
        rs("Plantype").value = 1
    ElseIf Opt(2).value = True Then
        rs("Plantype").value = 2
    ElseIf Opt(3).value = True Then
        rs("Plantype").value = 3
    ElseIf Opt(4).value = True Then
        rs("Plantype").value = 4
    End If
     
    If OptBranch(1).value = True Then
        rs("FixedBranch").value = 1
    Else
        rs("FixedBranch").value = 0
    End If

    rs("BranchId").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
         
    If optUnits(1).value = True Then
        rs("FixedUnit").value = 1
    Else
        rs("FixedUnit").value = 0
    End If

    rs("Unitid").value = IIf(Me.DcboUnits.BoundText = "", Null, Me.DcboUnits.BoundText)
         
    If optFixedPrice(1).value = True Then
        rs("FixedPrice").value = 1
    Else
        rs("FixedPrice").value = 0
    End If

    rs("priceID").value = IIf(Me.dcSalePriceNames.BoundText = "", Null, Me.dcSalePriceNames.BoundText)
    rs("priceChangeId").value = cbopriceChangeId.ListIndex
    rs("Operator").value = lblOperator.Caption
 
    If optFixedPrice(1).value = True Then
        rs("FixedPrice").value = 1
    Else
        rs("FixedPrice").value = 0
    End If
         
    rs("valueOrPercentage").value = cbovalueOrPercentage.ListIndex
    rs("valueOrPercentageValue").value = val(Me.txtvalueOrPercentageValue.Text)
    rs("AnotherPrice").value = val(Me.txtAnotherPrice.Text)
         
    rs.update
   
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblSalesPricesPlanDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    With Me.Grid1

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Item_id")) <> "" Then
         
                RsDev.AddNew
                RsDev("PlanId").value = Me.TxtPlanID.Text
            
                RsDev("branch_id").value = val(.TextMatrix(i, .ColIndex("BranchId")))
                RsDev("ItemID").value = val(.TextMatrix(i, .ColIndex("Item_id")))
           
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
             
                RsDev("PurchasePrice").value = val(.TextMatrix(i, .ColIndex("PurchasePrice")))
                RsDev("CostPrice").value = val(.TextMatrix(i, .ColIndex("CostPrice")))
                RsDev("SalePrice").value = val(.TextMatrix(i, .ColIndex("SalePrice")))
                RsDev("UnitWholeSalePrice").value = val(.TextMatrix(i, .ColIndex("UnitWholeSalePrice")))
                RsDev("Price1").value = val(.TextMatrix(i, .ColIndex("Price1")))
                RsDev("Price2").value = val(.TextMatrix(i, .ColIndex("Price2")))
                RsDev("Price3").value = val(.TextMatrix(i, .ColIndex("Price3")))
                RsDev("Price4").value = val(.TextMatrix(i, .ColIndex("Price4")))
                RsDev("Price5").value = val(.TextMatrix(i, .ColIndex("Price5")))
                RsDev("Price6").value = val(.TextMatrix(i, .ColIndex("Price6")))
            
                RsDev("NewPrice1").value = val(.TextMatrix(i, .ColIndex("NewPrice1")))
                RsDev("NewPrice2").value = val(.TextMatrix(i, .ColIndex("NewPrice2")))
                RsDev("NewPrice3").value = val(.TextMatrix(i, .ColIndex("NewPrice3")))
                RsDev("NewPrice4").value = val(.TextMatrix(i, .ColIndex("NewPrice4")))
                RsDev("NewPrice5").value = val(.TextMatrix(i, .ColIndex("NewPrice5")))
                RsDev("NewPrice6").value = val(.TextMatrix(i, .ColIndex("NewPrice6")))
             
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
    
    
    RsDev.Close
    
    RsDev.Open "TblSalesPricesPlanDetails2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    i = 1

    With Me.Grid2

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Item_id")) <> "" Then
         
                RsDev.AddNew
                RsDev("PlanId").value = Me.TxtPlanID.Text
            
                RsDev("branch_id").value = val(.TextMatrix(i, .ColIndex("BranchId")))
                RsDev("ItemID").value = val(.TextMatrix(i, .ColIndex("Item_id")))
           
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
             
                RsDev("PurchasePrice").value = val(.TextMatrix(i, .ColIndex("PurchasePrice")))
                RsDev("CostPrice").value = val(.TextMatrix(i, .ColIndex("CostPrice")))
                RsDev("SalePrice").value = val(.TextMatrix(i, .ColIndex("SalePrice")))
                RsDev("UnitWholeSalePrice").value = val(.TextMatrix(i, .ColIndex("UnitWholeSalePrice")))
                RsDev("SalePriceNew").value = val(.TextMatrix(i, .ColIndex("SalePriceNew")))
                RsDev("UnitWholeSalePriceNew").value = val(.TextMatrix(i, .ColIndex("UnitWholeSalePriceNew")))
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
        
    Cn.CommitTrans
    BeginTrans = False
    UpdatePrices
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.Text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub cbopriceChangeId_Change()

    If cbovalueOrPercentage.ListIndex = 3 Then
        txtAnotherPrice.locked = False
    Else
        txtAnotherPrice.locked = False
    End If

End Sub

Private Sub cbopriceChangeId_Click()
    cbopriceChangeId_Change
End Sub

Private Sub cbovalueOrPercentage_Change()

    If cbovalueOrPercentage.ListIndex = 1 Then
        LblPercentage.Visible = True
    Else
        LblPercentage.Visible = False
    End If

End Sub

Private Sub cbovalueOrPercentage_Click()
    cbovalueOrPercentage_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 9
            If mIndexTab = 0 Then
                Set mGrd = Grid1
            ElseIf mIndexTab = 1 Then
                Set mGrd = Grid2
            End If
    
            mGrd.Clear flexClearScrollable, flexClearEverything
            mGrd.Rows = 1

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Me.TxtPlanID.Text = CStr(new_id("TblSalesPricesPlan", "PlanId", "", True))
       
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
            Me.Optfixedintrval(0).value = True
            OptBranch(0).value = True
            optUnits(1).value = True
            Opt(3).value = True
            optFixedPrice(0).value = True

            'XPDtbTrans.SetFocus
            Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 1
            Grid1.Enabled = True
            
            Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.Rows = 1
            Grid2.Enabled = True
        Case 1
 
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            '         Grid.Rows = Grid.Rows + 1

            Grid1.Enabled = True
         
            CuurentLogdata

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            AddToGrid mIndexTab
            '   ViewDataList
            'addrowGroups
    
        Case 8
     
            RemoveGridRow mIndexTab

        Case 20
     
        Case 21
     
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap
    
    If TxtPlanID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtPlanID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            Cn.Execute "delete TblSalesPricesPlanDetails where PlanID=" & val(Me.TxtPlanID.Text)
            Cn.Execute "delete TblSalesPricesPlanDetails2 where PlanID=" & val(Me.TxtPlanID.Text)

            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                
                    Grid1.Clear flexClearScrollable, flexClearEverything
                    Grid1.Rows = 2
                    Grid1.Enabled = False
                
                     Grid2.Clear flexClearScrollable, flexClearEverything
                    Grid2.Rows = 2
                    Grid2.Enabled = False
                    TxtModFlg_Change
                    '     XPTxtCurrent.Caption = 0
                    '     XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub RemoveGridRow(Optional mIndex As Integer)
    If mIndex = 0 Then
        Set mGrd = Grid1
    ElseIf mIndex = 1 Then
        Set mGrd = Grid2
    End If
    
    With mGrd

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
 
End Sub

Function Retrive_Sales_invoice_data(Transaction_ID As Long, Transaction_Type As Integer)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, "
    StrSQL = StrSQL + "  dbo.TblItems.ItemName, dbo.Transaction_Details.Price, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transaction_Details.CostPrice,"
    StrSQL = StrSQL + " dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee,  dbo.Transaction_Details.BranchId"
    StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblBranchesData ON dbo.Transaction_Details.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transactions.Transaction_ID=" & Transaction_ID & " and  Transactions.Transaction_Type=" & Transaction_Type

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    Grid1.Rows = 2
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Refresh

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = Grid1.Rows
    
        If Grid1.TextMatrix(row_count - 1, Grid1.ColIndex("Item_code")) = "" Then
            row_count = row_count - 1
        End If
     
        Grid1.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To Grid1.Rows - 1 'RsDetails.RecordCount
 
            Grid1.TextMatrix(Num, Grid1.ColIndex("BranchId")) = IIf(IsNull(RsDetails("BranchId")), "", (RsDetails("BranchId").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                Grid1.TextMatrix(Num, Grid1.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_name")), "", (RsDetails("branch_name").value))
            Else
                Grid1.TextMatrix(Num, Grid1.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_namee")), "", (RsDetails("branch_namee").value))
            End If

            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_id")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
        
            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))

            If Transaction_Type = 22 Then '„‘ —Ì« 
                Grid1.TextMatrix(Num, Grid1.ColIndex("PurchasePrice")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            ElseIf Transaction_Type = 20 Then '«–‰ «÷«›Â
                Grid1.TextMatrix(Num, Grid1.ColIndex("CostPrice")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            ElseIf Transaction_Type = 21 Then '„»Ì⁄« 
                Grid1.TextMatrix(Num, Grid1.ColIndex("SalePrice")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
             
            ElseIf Transaction_Type = 19 Then '«–‰ ’—›
                Grid1.TextMatrix(Num, Grid1.ColIndex("CostPrice")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            End If
         
            Grid1.TextMatrix(Num, Grid1.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Function

Function RetriveAllItems(Optional BranchID As Integer = 0, Optional UnitID As Integer = 0, Optional GroupID As Integer = 0, Optional ItemID As Integer = 0, Optional orderNo As String = "")
 
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    Dim Begin  As Boolean
    Begin = False
    StrSQL = " SELECT     dbo.TblSalesPrices.ItemID, dbo.TblSalesPrices.Price1, dbo.TblSalesPrices.Price2, dbo.TblSalesPrices.Price3, dbo.TblSalesPrices.Price5, dbo.TblSalesPrices.Price4,"
    StrSQL = StrSQL & "  dbo.TblSalesPrices.Price6, dbo.TblSalesPrices.Discount1, dbo.TblSalesPrices.Discount2, dbo.TblSalesPrices.Discount3, dbo.TblSalesPrices.Discount4,"
    StrSQL = StrSQL & " dbo.TblSalesPrices.Discount5, dbo.TblSalesPrices.Discount6, dbo.TblUnites.UnitName, dbo.TblSalesPrices.UnitID, dbo.TblSalesPrices.BranchId,"
    StrSQL = StrSQL & " dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblItems.GroupID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    StrSQL = StrSQL & " dbo.Groups.GroupName"
    StrSQL = StrSQL & " FROM         dbo.TblSalesPrices INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.TblSalesPrices.ItemID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & " dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblSalesPrices.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblSalesPrices.UnitID = dbo.TblUnites.UnitID"

    If orderNo <> "" Then
        StrSQL = "SELECT DISTINCT "
        StrSQL = StrSQL & " dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.order_no, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID,"
        StrSQL = StrSQL & " dbo.TblItems.PurchasePrice, dbo.TblItems.SallingPrice, dbo.TblSalesPrices.Price1, dbo.TblSalesPrices.Price2, dbo.TblSalesPrices.Price3, dbo.TblSalesPrices.Price4,"
        StrSQL = StrSQL & " dbo.TblSalesPrices.Price6, dbo.TblSalesPrices.Price5, dbo.TblSalesPrices.UnitID, dbo.TblUnites.UnitName, dbo.TblSalesPrices.BranchId,"
        StrSQL = StrSQL & " dbo.TblBranchesData.branch_namee , dbo.TblBranchesData.branch_name"
        StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblSalesPrices ON dbo.TblItems.ItemID = dbo.TblSalesPrices.ItemID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblSalesPrices.UnitID = dbo.TblUnites.UnitID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblSalesPrices.BranchId = dbo.TblBranchesData.branch_id"
 
        StrSQL = StrSQL & " WHERE     (dbo.Transaction_Details.order_no = '" & orderNo & "')"

        GoTo ll
    End If

    If BranchID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and  BranchId=" & BranchID
        Else
            StrSQL = StrSQL + " where  BranchId=" & BranchID
            Begin = True
        End If
    End If

    If UnitID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and  TblSalesPrices.UnitID=" & UnitID
        Else
            StrSQL = StrSQL + " where  TblSalesPrices.UnitID =" & UnitID
            Begin = True
        End If
    End If

    If GroupID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and  TblItems.GroupID=" & GroupID
        Else
            StrSQL = StrSQL + " where  TblItems.GroupID =" & GroupID
            Begin = True
        End If
    End If

    If ItemID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and TblSalesPrices.itemid=" & ItemID
        Else
            StrSQL = StrSQL + " where  TblSalesPrices.itemid=" & ItemID
            Begin = True
        End If
    End If
  
ll:
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    'Grid1.Rows = 2
    'Grid1.Clear flexClearScrollable, flexClearEverything
    'Grid1.Refresh
    Dim costPrice As Double
    Dim LngItemID As Long

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = Grid1.Rows
    
        If Grid1.TextMatrix(row_count - 1, Grid1.ColIndex("Item_code")) = "" Then
            row_count = row_count - 1
        End If
     
        Grid1.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To Grid1.Rows - 1 'RsDetails.RecordCount
            Grid1.TextMatrix(Num, Grid1.ColIndex("SalePrice")) = IIf(IsNull(RsDetails("Price1")), "", (RsDetails("Price1").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price1")) = IIf(IsNull(RsDetails("Price1")), "", (RsDetails("Price1").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price2")) = IIf(IsNull(RsDetails("Price2")), "", (RsDetails("Price2").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price3")) = IIf(IsNull(RsDetails("Price3")), "", (RsDetails("Price3").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price4")) = IIf(IsNull(RsDetails("Price4")), "", (RsDetails("Price4").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price5")) = IIf(IsNull(RsDetails("Price5")), "", (RsDetails("Price5").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price6")) = IIf(IsNull(RsDetails("Price6")), "", (RsDetails("Price6").value))
      
            Grid1.TextMatrix(Num, Grid1.ColIndex("BranchId")) = IIf(IsNull(RsDetails("BranchId")), "", (RsDetails("BranchId").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                Grid1.TextMatrix(Num, Grid1.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_name")), "", (RsDetails("branch_name").value))
            Else
                Grid1.TextMatrix(Num, Grid1.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_namee")), "", (RsDetails("branch_namee").value))
            End If
                 
            If orderNo <> "" Then
                LngItemID = IIf(IsNull(RsDetails("Item_ID")), 0, (RsDetails("Item_ID").value))
        
            Else
                LngItemID = IIf(IsNull(RsDetails("ItemID")), 0, (RsDetails("ItemID").value))
            End If

            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_id")) = LngItemID
            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
        
            Dim DblTemp As Double
            '›Ï «·ŒÿÊ… «·√Ê·Ï ‰Õ«Ê· «‰ ‰« Ï »«Œ— ”⁄— ‘—«¡
            DblTemp = GetPrice(LngItemID, 1, False)

            If DblTemp = 0 Then '·«ÌÊÃœ «Œ— ”⁄— ‘—«¡
                DblTemp = GetPrice(LngItemID, 3, False)   '‰Õ«Ê· «·Õ’Ê· ⁄·Ï ”⁄— «·—’Ìœ «·√›  «ÕÏ

                If DblTemp = 0 Then
                    DblTemp = GetPrice(LngItemID, 9)  '«Œ— ‘Ï¡ ÂÊ «·Õ’Ê· ⁄·Ï ”⁄— «Œ— „— Ã⁄ „»Ì⁄« 
                End If
            End If
         
            Grid1.TextMatrix(Num, Grid1.ColIndex("PurchasePrice")) = DblTemp
            costPrice = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod)
            Grid1.TextMatrix(Num, Grid1.ColIndex("CostPrice")) = costPrice
         
            Grid1.TextMatrix(Num, Grid1.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Function




Function RetriveAllItems2(Optional BranchID As Integer = 0, Optional UnitID As Integer = 0, Optional GroupID As Integer = 0, Optional ItemID As Integer = 0, Optional orderNo As String = "")
 
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    Dim Begin  As Boolean
    Begin = False
    StrSQL = " SELECT     dbo.TblSalesPrices.ItemID, dbo.TblSalesPrices.Price1, dbo.TblSalesPrices.Price2, dbo.TblSalesPrices.Price3, dbo.TblSalesPrices.Price5, dbo.TblSalesPrices.Price4,"
    StrSQL = StrSQL & "  dbo.TblSalesPrices.Price6, dbo.TblSalesPrices.Discount1, dbo.TblSalesPrices.Discount2, dbo.TblSalesPrices.Discount3, dbo.TblSalesPrices.Discount4,"
    StrSQL = StrSQL & " dbo.TblSalesPrices.Discount5, dbo.TblSalesPrices.Discount6, dbo.TblUnites.UnitName, dbo.TblSalesPrices.UnitID, dbo.TblSalesPrices.BranchId,"
    StrSQL = StrSQL & " dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblItems.GroupID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    StrSQL = StrSQL & " dbo.Groups.GroupName"
    StrSQL = StrSQL & " FROM         dbo.TblSalesPrices INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.TblSalesPrices.ItemID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & " dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblSalesPrices.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblSalesPrices.UnitID = dbo.TblUnites.UnitID"

    If orderNo <> "" Then
        StrSQL = "SELECT DISTINCT "
        StrSQL = StrSQL & " dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.order_no, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID,"
        StrSQL = StrSQL & " dbo.TblItems.PurchasePrice, dbo.TblItems.SallingPrice, dbo.TblSalesPrices.Price1, dbo.TblSalesPrices.Price2, dbo.TblSalesPrices.Price3, dbo.TblSalesPrices.Price4,"
        StrSQL = StrSQL & " dbo.TblSalesPrices.Price6, dbo.TblSalesPrices.Price5, dbo.TblSalesPrices.UnitID, dbo.TblUnites.UnitName, dbo.TblSalesPrices.BranchId,"
        StrSQL = StrSQL & " dbo.TblBranchesData.branch_namee , dbo.TblBranchesData.branch_name"
        StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblSalesPrices ON dbo.TblItems.ItemID = dbo.TblSalesPrices.ItemID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblSalesPrices.UnitID = dbo.TblUnites.UnitID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblSalesPrices.BranchId = dbo.TblBranchesData.branch_id"
 
        StrSQL = StrSQL & " WHERE     (dbo.Transaction_Details.order_no = '" & orderNo & "')"

        GoTo ll
    End If

    If BranchID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and  BranchId=" & BranchID
        Else
            StrSQL = StrSQL + " where  BranchId=" & BranchID
            Begin = True
        End If
    End If

    If UnitID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and  TblSalesPrices.UnitID=" & UnitID
        Else
            StrSQL = StrSQL + " where  TblSalesPrices.UnitID =" & UnitID
            Begin = True
        End If
    End If

    If GroupID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and  TblItems.GroupID=" & GroupID
        Else
            StrSQL = StrSQL + " where  TblItems.GroupID =" & GroupID
            Begin = True
        End If
    End If

    If ItemID <> 0 Then

        If Begin = True Then
            StrSQL = StrSQL + " and TblSalesPrices.itemid=" & ItemID
        Else
            StrSQL = StrSQL + " where  TblSalesPrices.itemid=" & ItemID
            Begin = True
        End If
    End If
  
ll:
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    'Grid1.Rows = 2
    'Grid1.Clear flexClearScrollable, flexClearEverything
    'Grid1.Refresh
    Dim costPrice As Double
    Dim LngItemID As Long

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = Grid1.Rows
    
        If Grid1.TextMatrix(row_count - 1, Grid1.ColIndex("Item_code")) = "" Then
            row_count = row_count - 1
        End If
     
        Grid1.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To Grid1.Rows - 1 'RsDetails.RecordCount
            Grid1.TextMatrix(Num, Grid1.ColIndex("SalePrice")) = IIf(IsNull(RsDetails("Price1")), "", (RsDetails("Price1").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price1")) = IIf(IsNull(RsDetails("Price1")), "", (RsDetails("Price1").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price2")) = IIf(IsNull(RsDetails("Price2")), "", (RsDetails("Price2").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price3")) = IIf(IsNull(RsDetails("Price3")), "", (RsDetails("Price3").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price4")) = IIf(IsNull(RsDetails("Price4")), "", (RsDetails("Price4").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price5")) = IIf(IsNull(RsDetails("Price5")), "", (RsDetails("Price5").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Price6")) = IIf(IsNull(RsDetails("Price6")), "", (RsDetails("Price6").value))
      
            Grid1.TextMatrix(Num, Grid1.ColIndex("BranchId")) = IIf(IsNull(RsDetails("BranchId")), "", (RsDetails("BranchId").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                Grid1.TextMatrix(Num, Grid1.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_name")), "", (RsDetails("branch_name").value))
            Else
                Grid1.TextMatrix(Num, Grid1.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_namee")), "", (RsDetails("branch_namee").value))
            End If
                 
            If orderNo <> "" Then
                LngItemID = IIf(IsNull(RsDetails("Item_ID")), 0, (RsDetails("Item_ID").value))
        
            Else
                LngItemID = IIf(IsNull(RsDetails("ItemID")), 0, (RsDetails("ItemID").value))
            End If

            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_id")) = LngItemID
            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("Item_name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
        
            Dim DblTemp As Double
            '›Ï «·ŒÿÊ… «·√Ê·Ï ‰Õ«Ê· «‰ ‰« Ï »«Œ— ”⁄— ‘—«¡
            DblTemp = GetPrice(LngItemID, 1, False)

            If DblTemp = 0 Then '·«ÌÊÃœ «Œ— ”⁄— ‘—«¡
                DblTemp = GetPrice(LngItemID, 3, False)   '‰Õ«Ê· «·Õ’Ê· ⁄·Ï ”⁄— «·—’Ìœ «·√›  «ÕÏ

                If DblTemp = 0 Then
                    DblTemp = GetPrice(LngItemID, 9)  '«Œ— ‘Ï¡ ÂÊ «·Õ’Ê· ⁄·Ï ”⁄— «Œ— „— Ã⁄ „»Ì⁄« 
                End If
            End If
         
            Grid1.TextMatrix(Num, Grid1.ColIndex("PurchasePrice")) = DblTemp
            costPrice = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod)
            Grid1.TextMatrix(Num, Grid1.ColIndex("CostPrice")) = costPrice
         
            Grid1.TextMatrix(Num, Grid1.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            Grid1.TextMatrix(Num, Grid1.ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Function

Function AddToGrid(Optional ByVal mIndex As Integer)
    Dim Transaction_ID As Long
    Dim BranchID As Integer
    Dim UnitID As Integer
    Dim GroupID As Integer
    Dim ItemID  As Integer
    Dim Msg   As String

    If OptBranch(1).value = True Then
 
        If Dcbranch.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ    «·›—⁄ «Ê·«  "
            Else
                Msg = "Specify   Branch Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Dcbranch.SetFocus
            SendKeys "{F4}"
            Exit Function
        End If
 
    End If
 
    If optUnits(1).value = True And Opt(0).value <> True Then
 
        If DcboUnits.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ    «·ÊÕœ… «Ê·«  "
            Else
                Msg = "Specify   Unit Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboUnits.SetFocus
            SendKeys "{F4}"
            Exit Function
        End If
 
    End If
 
    If Opt(0).value = True Then  '„” ‰œ „⁄Ì‰
 
        If DCTransactionType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ ‰Ê⁄ «·„” ‰œ "
            Else
                Msg = "Specify Voucher type Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCTransactionType.SetFocus
            SendKeys "{F4}"
            Exit Function
        End If
        
        If Trim(TxtInvSerial.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ —ﬁ„ «·„” ‰œ "
            Else
                Msg = "Specify Voucher No Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtInvSerial.SetFocus
             
            Exit Function
        End If

    ElseIf Opt(1).value = True Then ' ﬂ· «·«’‰«›
 
    ElseIf Opt(2).value = True Then '„Ã„Ê⁄Â „ÕœœÂ

        If DCGroup.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ    «·„Ã„Ê⁄Â  «Ê·«  "
            Else
                Msg = "Specify   group  Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCGroup.SetFocus
            SendKeys "{F4}"
            Exit Function
        End If
        
    ElseIf Opt(3).value = True Then '«’‰«› „ÕœœÂ

        If dcitems.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ    «·’‰›  «Ê·«  "
            Else
                Msg = "Specify   Items  Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcitems.SetFocus
            SendKeys "{F4}"
            Exit Function
        End If

    ElseIf Opt(4).value = True Then '‘Õ‰Â „⁄Ì‰Â

        If Trim(txtOrderNo.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ —ﬁ„ «·‘Õ‰Â "
            Else
                Msg = "Specify order No Firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            txtOrderNo.SetFocus
             
            Exit Function
        End If

    End If
     
    BranchID = val(Dcbranch.BoundText)
    UnitID = val(DcboUnits.BoundText)
 
    If Opt(0).value = True Then '„” ‰œ
        GetTransIDFromNoteSerial1 Me.TxtInvSerial.Text, Transaction_ID, , DCTransactionType.ItemData(DCTransactionType.ListIndex)
 
        Retrive_Sales_invoice_data Transaction_ID, DCTransactionType.ItemData(DCTransactionType.ListIndex)

    ElseIf Opt(1).value = True Then 'ﬂ· «·«’‰«›
        TxtItemsIDes.Text = ""
        RetriveAllItems BranchID, UnitID

    ElseIf Opt(2).value = True Then '„Ã„Ê⁄Â „ÕœœÂ
                  
        GroupID = val(DCGroup.BoundText)
        If chkIsNewPric.value = vbChecked Then
            RetriveAllItems2 BranchID, UnitID, GroupID
        Else
            RetriveAllItems BranchID, UnitID, GroupID
        End If

    ElseIf Opt(3).value = True Then '«’‰«› „ÕœœÂ
 
        ItemID = val(dcitems.BoundText)
        TxtItemsIDes.Text = ItemID
        If chkIsNewPric.value = vbChecked Then
            
            Retrive_Items_data1
            'RetriveAllItems2 BranchID, UnitID, GroupID, ItemID
        Else
            RetriveAllItems BranchID, UnitID, GroupID, ItemID
        End If
    ElseIf Opt(4).value = True Then '‘Õ‰Â „ÕœœÂ
        If chkIsNewPric.value = vbChecked Then
        
            RetriveAllItems2 BranchID, UnitID, , , Me.txtOrderNo.Text
        Else
            RetriveAllItems BranchID, UnitID, , , Me.txtOrderNo.Text
        End If

    End If

End Function
 
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)

End Sub

Private Sub Command1_Click()
 
        Load FrmItemSearch
        FrmItemSearch.mRow = Grid2.Row
        FrmItemSearch.RetrunType = 9888
        FrmItemSearch.show vbModal
    
End Sub

Function UpdatePrices()
    Dim UnitID As Long
    Dim ItemID As Long
    Dim BranchID As Long
    Dim StrSQL As String
    Dim Price1 As Double
    Dim Price2 As Double
    Dim Price3 As Double
    Dim Price4 As Double
    Dim Price5 As Double
    Dim Price6 As Double
    Dim i As Integer
    If mIndexTab = 1 And chkIsNewPric.value = vbChecked Then
        With Me.Grid2

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Item_id")) <> "" Then
                Price1 = val(.TextMatrix(i, .ColIndex("SalePriceNew")))
                Price2 = val(.TextMatrix(i, .ColIndex("UnitWholeSalePriceNew")))
                UnitID = val(.TextMatrix(i, .ColIndex("UnitId")))
                ItemID = val(.TextMatrix(i, .ColIndex("Item_id")))
                BranchID = val(.TextMatrix(i, .ColIndex("BranchId")))
                 



                If val(dcSalePriceNames.BoundText) = 2 Or val(dcSalePriceNames.BoundText) = 1 Then
                    StrSQL = "update TblItemsUnits set OldUnitSalesPrice= UnitSalesPrice,UnitSalesPrice=" & Price1 & " where ItemID=" & ItemID & " and UnitID=" & UnitID
                ElseIf val(dcSalePriceNames.BoundText) = 4 Then
                    StrSQL = "update TblItemsUnits set OldUnitWholeSalePrice = UnitWholeSalePrice,UnitWholeSalePrice=" & Price2 & " where ItemID=" & ItemID & " and UnitID=" & UnitID
                Else
                    StrSQL = "update TblItemsUnits set OldUnitSalesPrice= UnitSalesPrice,UnitSalesPrice=" & Price1 & " where ItemID=" & ItemID & " and UnitID=" & UnitID
                End If
                
                Cn.Execute StrSQL
            
                'StrSQL = "update TblItems set SallingPrice=" & Price1 & ",CustomerPrice=" & Price2 & ",DealerPrice=" & Price3 & "where ItemID=" & ItemID
             
                'Cn.Execute StrSQL
            End If

        Next i

    End With

    Else
        With Me.Grid1
    
            For i = 1 To .Rows - 1
    
                If .TextMatrix(i, .ColIndex("Item_id")) <> "" Then
                    Price1 = val(.TextMatrix(i, .ColIndex("NewPrice1")))
                    Price2 = val(.TextMatrix(i, .ColIndex("NewPrice2")))
                    Price3 = val(.TextMatrix(i, .ColIndex("NewPrice3")))
                    Price4 = val(.TextMatrix(i, .ColIndex("NewPrice4")))
                    Price5 = val(.TextMatrix(i, .ColIndex("NewPrice5")))
                    Price6 = val(.TextMatrix(i, .ColIndex("NewPrice6")))
                    UnitID = val(.TextMatrix(i, .ColIndex("UnitId")))
                    ItemID = val(.TextMatrix(i, .ColIndex("Item_id")))
                    BranchID = val(.TextMatrix(i, .ColIndex("BranchId")))
                    StrSQL = "update TblSalesPrices set Price1=" & Price1 & ",Price2=" & Price2 & ",Price3=" & Price3 & ",Price4=" & Price4 & ",Price5=" & Price5 & ",Price6=" & Price6 & "where ItemID=" & ItemID & " and UnitID=" & UnitID & "and BranchId=" & BranchID
                    Cn.Execute StrSQL
                
                    StrSQL = "update TblItems set SallingPrice=" & Price1 & ",CustomerPrice=" & Price2 & ",DealerPrice=" & Price3 & "where ItemID=" & ItemID
                 
                    Cn.Execute StrSQL
                End If
    
            Next i
    
        End With
    End If
End Function

Private Sub CmdDo_Click()
    Dim oldprice As Double

    Dim Newprice   As Double
    Dim OperationValue As Double
    Dim objScript As Object
    Dim i As Long
    Dim new_str As String
    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
    If mIndexTab = 1 And chkIsNewPric.value = vbChecked Then
           With Me.Grid2
    
            For i = 1 To .Rows - 1
    
                If .TextMatrix(i, .ColIndex("Item_id")) <> "" Then
                    If cbopriceChangeId.ListIndex = 0 Then
                        oldprice = val(.TextMatrix(i, .ColIndex("PurchasePrice")))
                    ElseIf cbopriceChangeId.ListIndex = 1 Then
                        oldprice = val(.TextMatrix(i, .ColIndex("CostPrice")))
                    ElseIf cbopriceChangeId.ListIndex = 2 Then
                        oldprice = val(.TextMatrix(i, .ColIndex("SalePrice")))
                    ElseIf cbopriceChangeId.ListIndex = 3 Then
                        oldprice = val(txtAnotherPrice)
                    End If
    
                    If cbovalueOrPercentage.ListIndex = -1 Then
                        OperationValue = 0
                    ElseIf cbovalueOrPercentage.ListIndex = 0 Then
                        OperationValue = val(txtvalueOrPercentageValue.Text)
                    ElseIf cbovalueOrPercentage.ListIndex = 1 Then
                        OperationValue = oldprice * val(txtvalueOrPercentageValue.Text) / 100
                    End If
    
                    new_str = oldprice & lblOperator.Caption & OperationValue
                    Newprice = objScript.Eval(new_str)
                    Newprice = Round(Newprice, 2)
    
                    If optFixedPrice(0).value = True Then
                 
                        .TextMatrix(i, .ColIndex("UnitWholeSalePriceNew")) = Newprice
                        
                    Else
    

                        If val(dcSalePriceNames.BoundText) = 2 Then
                            .TextMatrix(i, .ColIndex("SalePriceNew")) = Newprice
                        ElseIf val(dcSalePriceNames.BoundText) = 1 Then
                            .TextMatrix(i, .ColIndex("SalePriceNew")) = Newprice


                        ElseIf val(dcSalePriceNames.BoundText) = 4 Then
                            .TextMatrix(i, .ColIndex("UnitWholeSalePriceNew")) = Newprice

                      
                        End If
                    End If
                        
                End If
                
                '
            Next i
    
        End With
    
    Else
        With Me.Grid1
    
            For i = 1 To .Rows - 1
    
                If .TextMatrix(i, .ColIndex("Item_id")) <> "" Then
                    If cbopriceChangeId.ListIndex = 0 Then
                        oldprice = val(.TextMatrix(i, .ColIndex("PurchasePrice")))
                    ElseIf cbopriceChangeId.ListIndex = 1 Then
                        oldprice = val(.TextMatrix(i, .ColIndex("CostPrice")))
                    ElseIf cbopriceChangeId.ListIndex = 2 Then
                        oldprice = val(.TextMatrix(i, .ColIndex("SalePrice")))
                    ElseIf cbopriceChangeId.ListIndex = 3 Then
                        oldprice = val(txtAnotherPrice)
                    End If
    
                    If cbovalueOrPercentage.ListIndex = -1 Then
                        OperationValue = 0
                    ElseIf cbovalueOrPercentage.ListIndex = 0 Then
                        OperationValue = val(txtvalueOrPercentageValue.Text)
                    ElseIf cbovalueOrPercentage.ListIndex = 1 Then
                        OperationValue = oldprice * val(txtvalueOrPercentageValue.Text) / 100
                    End If
    
                    new_str = oldprice & lblOperator.Caption & OperationValue
                    Newprice = objScript.Eval(new_str)
                    Newprice = Round(Newprice, 0)
    
                    If optFixedPrice(0).value = True Then
                 
                        .TextMatrix(i, .ColIndex("NewPrice1")) = Newprice
                        .TextMatrix(i, .ColIndex("NewPrice2")) = Newprice
                        .TextMatrix(i, .ColIndex("NewPrice3")) = Newprice
                        .TextMatrix(i, .ColIndex("NewPrice4")) = Newprice
                        .TextMatrix(i, .ColIndex("NewPrice5")) = Newprice
                        .TextMatrix(i, .ColIndex("NewPrice6")) = Newprice
                    Else
    
                        If val(dcSalePriceNames.BoundText) = 1 Then
                            .TextMatrix(i, .ColIndex("NewPrice1")) = Newprice
                        ElseIf val(dcSalePriceNames.BoundText) = 2 Then
                            .TextMatrix(i, .ColIndex("NewPrice2")) = Newprice
                        ElseIf val(dcSalePriceNames.BoundText) = 3 Then
                            .TextMatrix(i, .ColIndex("NewPrice3")) = Newprice
                        ElseIf val(dcSalePriceNames.BoundText) = 4 Then
                            .TextMatrix(i, .ColIndex("NewPrice4")) = Newprice
                        ElseIf val(dcSalePriceNames.BoundText) = 5 Then
                            .TextMatrix(i, .ColIndex("NewPrice5")) = Newprice
                        ElseIf val(dcSalePriceNames.BoundText) = 6 Then
                            .TextMatrix(i, .ColIndex("NewPrice6")) = Newprice
                      
                        End If
                    End If
                        
                End If
                
                '
            Next i
    
        End With
    End If
End Sub

Private Sub cmdOperator_Click(Index As Integer)
    Me.lblOperator.Caption = cmdOperator(Index).Caption
End Sub

Private Sub dcitems_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        If mIndexTab = 1 And chkIsNewPric.value = vbChecked Then
            FrmItemSearch.mRow = Grid2.Row
            FrmItemSearch.RetrunType = 9888
        Else
            FrmItemSearch.RetrunType = 11
        End If
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub Form_Load()

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = "  ŒÿÂ «”⁄«— «·«’‰«›   "
    ScreenNameEnglish = " Items Price  Plan "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    'With Me.Grid
    '   Set .WallPaper = GrdBack.Picture
     
    'End With

    With Me.Grid1
        Set .WallPaper = GrdBack.Picture
     
    End With

    My_SQL = " select branch_id,branch_name from TblBranchesData"
    fill_combo Dcbranch, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetItemsNames dcitems
    Dcombos.GetItemSGroups DCGroup
    Dcombos.GetItemsUnits Me.DcboUnits
    Dcombos.GetSalePriceNames dcSalePriceNames

    'With Me.Grid
    '    .Rows = 1
    '     .ExplorerBar = flexExSortShowAndMove
    '   .RowHeightMin = 300
    '     .ExtendLastCol = True
    'End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblSalesPricesPlan  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Items Prices Plan "
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(9).Caption = "Date"
 
    lbl(3).Caption = "Remarks"
 
    Optfixedintrval(0).Caption = " Unknown"
    Optfixedintrval(1).Caption = " Specified"
    lbl(5).Caption = "From"
    lbl(2).Caption = "To"
    Frame1.Caption = "Select Branch"
    OptBranch(0).Caption = "All Branches"
    OptBranch(1).Caption = "Specific Branch"
    Frame2.Caption = "Select Units"
    optUnits(0).Caption = "All Units"
    optUnits(1).Caption = "Specific Unit"
    Opt(1).Caption = "All Items"
    Opt(2).Caption = "Specific Unit"
    Opt(4).Caption = "Shipment"
 
    lbl(10).Caption = "No."
    Opt(3).Caption = "Specific Item"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Remove"
    Cmd(9).Caption = "Remove All Line"
    Frame3.Caption = "Select Price name to change it"
    optFixedPrice(0).Caption = "All Prices"
    optFixedPrice(1).Caption = "Specific Price"

    Label2.Caption = "Determinants "

    With cbopriceChangeId
        .Clear
        .AddItem "Last Purchase Price"
        .AddItem "Average Cost"
        .AddItem "Current Sale Price"
        .AddItem "Other"
    End With

    lbl(8).Caption = "Value/Percentage"
    CMDDO.Caption = "Start Plan"
    CmdRemove.Caption = "Remove Line"

    lbl(1).Caption = "Other Price"

    Opt(0).Caption = "Specific Voucher"
    lbl(0).Caption = "NO"

    With Me.Grid1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("branchname")) = " Branch  Name"
        .TextMatrix(0, .ColIndex("item_code")) = "Item code"
        .TextMatrix(0, .ColIndex("Item_name")) = "Item Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("purchaseprice")) = "Purchase Price"
        .TextMatrix(0, .ColIndex("costPrice")) = "Cost Price"
        .TextMatrix(0, .ColIndex("SalePrice")) = "Sale Price"
        .TextMatrix(0, .ColIndex("Price1")) = "Price1"
        .TextMatrix(0, .ColIndex("NewPrice1")) = "NewPrice1"
        .TextMatrix(0, .ColIndex("Price2")) = "Price2"
        .TextMatrix(0, .ColIndex("NewPrice2")) = "NewPrice2"
        .TextMatrix(0, .ColIndex("Price3")) = "Price3"
        .TextMatrix(0, .ColIndex("NewPrice3")) = "NewPrice3"
        .TextMatrix(0, .ColIndex("Price4")) = "Price4"
        .TextMatrix(0, .ColIndex("NewPrice4")) = "NewPrice4"
        .TextMatrix(0, .ColIndex("Price5")) = "Price5"
        .TextMatrix(0, .ColIndex("NewPrice5")) = "NewPrice5"
        .TextMatrix(0, .ColIndex("Price6")) = "Price6"
        .TextMatrix(0, .ColIndex("NewPrice6")) = "NewPrice6"
    End With
    
    C1Tab1.TabCaption(0) = "Prices Plan"
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    'Grid.Clear flexClearScrollable, flexClearEverything
    'Grid.Rows = 1
          
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Rows = 1

    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtPlanID.Text = IIf(IsNull(rs("PlanID").value), "", rs("PlanID").value)
    Me.DPRecorddate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    DCGroup.BoundText = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)
    DCTransactionType.ListIndex = IIf(IsNull(rs("TransactionType").value), -1, rs("TransactionType").value)
    TxtInvSerial.Text = IIf(IsNull(rs("InvSerial").value), "", rs("InvSerial").value)

    If (rs("FixedInterval").value) = True Then
        Optfixedintrval(1).value = True
    Else
        Optfixedintrval(0).value = True
    End If



    If (rs("IsNewPric").value) = 1 Then
        chkIsNewPric.value = vbChecked
    Else
        chkIsNewPric.value = vbUnchecked
    End If
chkIsNewPric_Click

    dbFromDate.value = IIf(IsNull(rs("IntervalFrom").value), Date, rs("IntervalFrom").value)
    dbTodate.value = IIf(IsNull(rs("intervalto").value), Date, rs("intervalto").value)
    TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    Me.txtOrderNo.Text = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)

    Opt(val((rs("Plantype").value))).value = True

    If (rs("FixedBranch").value) = True Then
        OptBranch(1).value = True
    Else
        OptBranch(0).value = True
    End If

    Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    If (rs("FixedUnit").value) = True Then
        optUnits(1).value = True
    Else
        optUnits(0).value = True
    End If

    DcboUnits.BoundText = IIf(IsNull(rs("Unitid").value), "", rs("Unitid").value)

    If (rs("FixedPrice").value) = True Then
        optFixedPrice(1).value = True
    Else
        optFixedPrice(0).value = True
    End If

    dcSalePriceNames.BoundText = IIf(IsNull(rs("priceID").value), "", rs("priceID").value)

    cbopriceChangeId.ListIndex = IIf(IsNull(rs("priceChangeId").value), -1, val(rs("priceChangeId").value))
    lblOperator.Caption = IIf(IsNull(rs("Operator").value), "", (rs("Operator").value))
    cbovalueOrPercentage.ListIndex = IIf(IsNull(rs("valueOrPercentage").value), -1, val(rs("valueOrPercentage").value))
    txtvalueOrPercentageValue.Text = IIf(IsNull(rs("valueOrPercentageValue").value), 0, val(rs("valueOrPercentageValue").value))
    txtAnotherPrice.Text = IIf(IsNull(rs("AnotherPrice").value), 0, val(rs("AnotherPrice").value))
    
    
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblSalesPricesPlanDetails.Id,dbo.TblSalesPricesPlanDetails.UnitWholeSalePrice, dbo.TblSalesPricesPlanDetails.PlanId, dbo.TblSalesPricesPlanDetails.branch_id, "
    StrSQL = StrSQL & "    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblSalesPricesPlanDetails.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    StrSQL = StrSQL & "    dbo.TblUnites.UnitName, dbo.TblSalesPricesPlanDetails.UnitID, dbo.TblSalesPricesPlanDetails.PurchasePrice, dbo.TblSalesPricesPlanDetails.CostPrice,"
    StrSQL = StrSQL & "    dbo.TblSalesPricesPlanDetails.SalePrice, dbo.TblSalesPricesPlanDetails.Price1, dbo.TblSalesPricesPlanDetails.Price2, dbo.TblSalesPricesPlanDetails.Price3,"
    StrSQL = StrSQL & "    dbo.TblSalesPricesPlanDetails.Price4, dbo.TblSalesPricesPlanDetails.Price5, dbo.TblSalesPricesPlanDetails.Price6, dbo.TblSalesPricesPlanDetails.newprice1,"
    StrSQL = StrSQL & "    dbo.TblSalesPricesPlanDetails.newprice2, dbo.TblSalesPricesPlanDetails.newprice3, dbo.TblSalesPricesPlanDetails.newprice4,"
    StrSQL = StrSQL & "    dbo.TblSalesPricesPlanDetails.newprice5 , dbo.TblSalesPricesPlanDetails.newprice6"
    StrSQL = StrSQL & "    FROM         dbo.TblSalesPricesPlanDetails INNER JOIN"
    StrSQL = StrSQL & "    dbo.TblBranchesData ON dbo.TblSalesPricesPlanDetails.branch_id = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & "    dbo.TblItems ON dbo.TblSalesPricesPlanDetails.ItemID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "    dbo.TblUnites ON dbo.TblSalesPricesPlanDetails.UnitID = dbo.TblUnites.UnitID"
      
    StrSQL = StrSQL & "  WHERE     (dbo.TblSalesPricesPlanDetails.PlanId = " & val(TxtPlanID.Text) & ")"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
            
                Else
                    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
            
                End If
            
                .TextMatrix(i, .ColIndex("Item_id")) = IIf(IsNull(RsDev("ItemID").value), "", RsDev("ItemID").value)
            
                .TextMatrix(i, .ColIndex("Item_code")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
            
                .TextMatrix(i, .ColIndex("Item_name")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
            
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(RsDev("UnitID").value), "", RsDev("UnitID").value)
            
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("PurchasePrice")) = IIf(IsNull(RsDev("PurchasePrice").value), "", val(RsDev("PurchasePrice").value))
                .TextMatrix(i, .ColIndex("UnitWholeSalePrice")) = IIf(IsNull(RsDev("UnitWholeSalePrice").value), "", val(RsDev("UnitWholeSalePrice").value))
            
                .TextMatrix(i, .ColIndex("CostPrice")) = IIf(IsNull(RsDev("CostPrice").value), "", val(RsDev("CostPrice").value))
            
                .TextMatrix(i, .ColIndex("SalePrice")) = IIf(IsNull(RsDev("SalePrice").value), "", val(RsDev("SalePrice").value))
            
                .TextMatrix(i, .ColIndex("Price1")) = IIf(IsNull(RsDev("Price1").value), "", val(RsDev("Price1").value))
            
                .TextMatrix(i, .ColIndex("Price2")) = IIf(IsNull(RsDev("Price2").value), "", val(RsDev("Price2").value))
            
                .TextMatrix(i, .ColIndex("Price3")) = IIf(IsNull(RsDev("Price3").value), "", val(RsDev("Price3").value))
            
                .TextMatrix(i, .ColIndex("Price4")) = IIf(IsNull(RsDev("Price4").value), "", val(RsDev("Price4").value))
            
                .TextMatrix(i, .ColIndex("Price5")) = IIf(IsNull(RsDev("Price5").value), "", val(RsDev("Price5").value))
            
                .TextMatrix(i, .ColIndex("Price6")) = IIf(IsNull(RsDev("Price6").value), "", val(RsDev("Price6").value))
            
                .TextMatrix(i, .ColIndex("NewPrice1")) = IIf(IsNull(RsDev("NewPrice1").value), "", val(RsDev("NewPrice1").value))
            
                .TextMatrix(i, .ColIndex("NewPrice2")) = IIf(IsNull(RsDev("NewPrice2").value), "", val(RsDev("NewPrice2").value))
            
                .TextMatrix(i, .ColIndex("NewPrice3")) = IIf(IsNull(RsDev("NewPrice3").value), "", val(RsDev("NewPrice3").value))
            
                .TextMatrix(i, .ColIndex("NewPrice4")) = IIf(IsNull(RsDev("NewPrice4").value), "", val(RsDev("NewPrice4").value))
            
                .TextMatrix(i, .ColIndex("NewPrice5")) = IIf(IsNull(RsDev("NewPrice5").value), "", val(RsDev("NewPrice5").value))
            
                .TextMatrix(i, .ColIndex("NewPrice6")) = IIf(IsNull(RsDev("NewPrice6").value), "", val(RsDev("NewPrice6").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    
    
Grid2.Rows = 1
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblSalesPricesPlanDetails2.Id,TblSalesPricesPlanDetails2.SalePrice,"
    StrSQL = StrSQL & "    dbo.TblSalesPricesPlanDetails2.UnitWholeSalePrice,dbo.TblSalesPricesPlanDetails2.UnitWholeSalePriceNew, dbo.TblSalesPricesPlanDetails2.PlanId, dbo.TblSalesPricesPlanDetails2.branch_id, "
    StrSQL = StrSQL & "    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblSalesPricesPlanDetails2.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    StrSQL = StrSQL & "    dbo.TblUnites.UnitName, dbo.TblSalesPricesPlanDetails2.UnitID, dbo.TblSalesPricesPlanDetails2.PurchasePrice, dbo.TblSalesPricesPlanDetails2.CostPrice,TblSalesPricesPlanDetails2.SalePriceNew "
    
    StrSQL = StrSQL & "    FROM         dbo.TblSalesPricesPlanDetails2 Left Outer JOIN"
    StrSQL = StrSQL & "    dbo.TblBranchesData ON dbo.TblSalesPricesPlanDetails2.branch_id = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & "    dbo.TblItems ON dbo.TblSalesPricesPlanDetails2.ItemID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "    dbo.TblUnites ON dbo.TblSalesPricesPlanDetails2.UnitID = dbo.TblUnites.UnitID"
      
    StrSQL = StrSQL & "  WHERE     (dbo.TblSalesPricesPlanDetails2.PlanId = " & val(TxtPlanID.Text) & ")"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid2
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
                .TextMatrix(i, .ColIndex("Ser")) = i
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
            
                Else
                    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
            
                End If
            
                .TextMatrix(i, .ColIndex("Item_id")) = IIf(IsNull(RsDev("ItemID").value), "", RsDev("ItemID").value)
            
                .TextMatrix(i, .ColIndex("Item_code")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
            
                .TextMatrix(i, .ColIndex("Item_name")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
            
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(RsDev("UnitID").value), "", RsDev("UnitID").value)
            
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("PurchasePrice")) = IIf(IsNull(RsDev("PurchasePrice").value), "", val(RsDev("PurchasePrice").value))
                .TextMatrix(i, .ColIndex("UnitWholeSalePrice")) = IIf(IsNull(RsDev("UnitWholeSalePrice").value), "", val(RsDev("UnitWholeSalePrice").value))
                .TextMatrix(i, .ColIndex("UnitWholeSalePriceNew")) = IIf(IsNull(RsDev("UnitWholeSalePriceNew").value), "", val(RsDev("UnitWholeSalePriceNew").value))
            
                .TextMatrix(i, .ColIndex("CostPrice")) = IIf(IsNull(RsDev("CostPrice").value), "", val(RsDev("CostPrice").value))
            
                .TextMatrix(i, .ColIndex("SalePrice")) = IIf(IsNull(RsDev("SalePrice").value), "", val(RsDev("SalePrice").value))
                .TextMatrix(i, .ColIndex("SalePriceNew")) = IIf(IsNull(RsDev("SalePriceNew").value), "", val(RsDev("SalePriceNew").value))
                
            
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
    
    
    Exit Sub
ErrTrap:
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " „ " & TxtPlanID.Text & CHR(13) & "    «—ÌŒ «·Œÿ… " & DPRecorddate
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " No " & TxtPlanID.Text & CHR(13) & " Date" & DPRecorddate
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function
 
Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid1

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
         
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If
     
    End With

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With Grid1

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub Option3_Click()

End Sub

Private Sub Option4_Click()

End Sub

Private Sub GRID2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Grid2

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If
        If .ColKey(Col) = "UnitWholeSalePrice" Or .ColKey(Col) = "SalePrice" Or .ColKey(Col) = "PurchasePrice" Then
            Cancel = True
        End If
    End With
End Sub

Private Sub Optfixedintrval_Click(Index As Integer)

    Select Case Index

        Case 0
            Frame4.Visible = False

        Case 1
            Frame4.Visible = True

    End Select

End Sub

Private Sub TabMain_Click()
mIndexTab = TabMain.CurrTab
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        'Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        CMDDO.Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
 
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False
        CMDDO.Enabled = True
    Else
        'Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True
        CMDDO.Enabled = False

    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub



Function Retrive_Items_data1()
    Dim StrSQL  As String
    Dim row_count As Long
    Dim Num As Long
    Dim i As Long
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    StrSQL = "select TblItems.ItemName,TblItems.ItemID,TblItemsUnits.UnitId,TblItemsUnits.UnitPurPrice PurchasePrice, TblUnites.UnitName,TblItemsUnits.UnitWholeSalePrice,TblItemsUnits.UnitSalesPrice SalePrice,TblItems.ItemCode from TblItems"
    StrSQL = StrSQL & " Inner join TblItemsUnits On TblItemsUnits.ItemId = TblItems.ItemId "
    StrSQL = StrSQL & " Inner join TblUnites On TblItemsUnits.UnitId = TblUnites.UnitId "
    
    If TxtItemsIDes.Text <> "" Then
        StrSQL = StrSQL & " where TblItemsUnits.ItemID in(" & TxtItemsIDes.Text & ")"
    End If
    StrSQL = StrSQL & " Order By TblItems.ItemId"
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   If rs2.RecordCount > 0 Then
        
        If Grid2.TextMatrix(Grid2.Rows - 1, Grid2.ColIndex("Item_id")) = "" Then
            Grid2.Rows = Grid2.Rows - 1
        End If
     With Grid2
     row_count = Grid2.Rows
       rs2.MoveFirst
       .Rows = rs2.RecordCount + .Rows
        For Num = row_count To .Rows - 1 'RsDetails.RecordCount
            .TextMatrix(Num, .ColIndex("Item_id")) = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            .TextMatrix(Num, .ColIndex("Item_code")) = IIf(IsNull(rs2("ItemCode").value), 0, rs2("ItemCode").value)
            .TextMatrix(Num, .ColIndex("Item_name")) = IIf(IsNull(rs2("ItemName").value), 0, rs2("ItemName").value)
            .TextMatrix(Num, .ColIndex("UnitId")) = IIf(IsNull(rs2("UnitId").value), 0, rs2("UnitId").value)
            .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitName").value), 0, rs2("UnitName").value)
            .TextMatrix(Num, .ColIndex("SalePrice")) = IIf(IsNull(rs2("SalePrice").value), 0, rs2("SalePrice").value)
            .TextMatrix(Num, .ColIndex("UnitWholeSalePrice")) = IIf(IsNull(rs2("UnitWholeSalePrice").value), 0, rs2("UnitWholeSalePrice").value)
            .TextMatrix(Num, .ColIndex("PurchasePrice")) = IIf(IsNull(rs2("PurchasePrice").value), 0, rs2("PurchasePrice").value)
            
    
       'TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(Rs2("ItemName").value), "", Rs2("ItemName").value)
        rs2.MoveNext
        Next Num
       
    End With
    End If
End Function

Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\Gard.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡  «·«’‰«› €Ì— «·„ÊÃÊœ… "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub
 

