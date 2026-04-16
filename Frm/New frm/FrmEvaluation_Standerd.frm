VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmEvaluation_Standerd 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„⁄«ÌÌ— «· ﬁÌÌ„"
   ClientHeight    =   9675
   ClientLeft      =   6705
   ClientTop       =   1620
   ClientWidth     =   10155
   Icon            =   "FrmEvaluation_Standerd.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   10155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9672
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10152
      _cx             =   17912
      _cy             =   17066
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   510
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   8235
         Width           =   9690
         _cx             =   17092
         _cy             =   900
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
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   330
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   330
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   330
            Index           =   2
            Left            =   6390
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   120
            Width           =   1965
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   330
            Index           =   4
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   1455
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   732
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   10140
         _cx             =   17886
         _cy             =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "   „⁄«ÌÌ— «· ﬁÌÌ„  "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   1
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
            Height          =   345
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   5
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation_Standerd.frx":038A
            ColorButton     =   -2147483634
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
            Height          =   345
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation_Standerd.frx":0724
            ColorButton     =   -2147483634
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
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   7
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation_Standerd.frx":0ABE
            ColorButton     =   -2147483634
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
            Height          =   345
            Index           =   3
            Left            =   615
            TabIndex        =   8
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmEvaluation_Standerd.frx":0E58
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
      End
      Begin C1SizerLibCtl.C1Elastic pnlHeader 
         Height          =   3228
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   9840
         _cx             =   17357
         _cy             =   5689
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   615
            Left            =   120
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3975
            _cx             =   7011
            _cy             =   1085
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
            Begin VB.TextBox TxtAvgAbscen 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   120
               Width           =   855
            End
            Begin VB.TextBox TxtNoDayAbcen 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì”«ÊÌ"
               Height          =   255
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «Ì«„ «·€Ì«»"
               Height          =   255
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   120
               Width           =   1095
            End
         End
         Begin XtremeSuiteControls.RadioButton Emp_Stude 
            Height          =   375
            Index           =   0
            Left            =   8880
            TabIndex        =   62
            Top             =   840
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "„ÊŸ›"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox MaxDgree 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   600
            Width           =   1764
         End
         Begin VB.TextBox ExcelTo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   2760
            Width           =   1092
         End
         Begin VB.TextBox ExcelFrom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   2760
            Width           =   1212
         End
         Begin VB.TextBox VeryGTo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   2400
            Width           =   1092
         End
         Begin VB.TextBox VeryGFrom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   2400
            Width           =   1212
         End
         Begin VB.TextBox GoodTo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2040
            Width           =   1092
         End
         Begin VB.TextBox GoodFrom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2040
            Width           =   1212
         End
         Begin VB.TextBox InterTo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1680
            Width           =   1092
         End
         Begin VB.TextBox InterFrom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1680
            Width           =   1212
         End
         Begin VB.TextBox WeakTo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1320
            Width           =   1092
         End
         Begin VB.TextBox WeakFrom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1320
            Width           =   1212
         End
         Begin VB.TextBox ENameE 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   3156
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   600
            Width           =   2004
         End
         Begin VB.TextBox EName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   600
            Width           =   1980
         End
         Begin VB.TextBox ID 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6600
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   1980
         End
         Begin MSComCtl2.DTPicker SDate 
            Height          =   312
            Left            =   3156
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   2004
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95551491
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo BranchID 
            Height          =   288
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Visible         =   0   'False
            Width           =   1764
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Emp_Stude 
            Height          =   375
            Index           =   1
            Left            =   7440
            TabIndex        =   63
            Top             =   840
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "„ œ—»"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·œ—Ã… «·ﬁ’ÊÏ"
            Height          =   372
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   372
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   2760
            Width           =   492
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   372
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   2760
            Width           =   612
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„„ «“"
            Height          =   372
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   2760
            Width           =   852
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   372
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   2400
            Width           =   492
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   372
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   2400
            Width           =   612
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÌœ Ãœ«"
            Height          =   372
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   2400
            Width           =   852
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   372
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   2040
            Width           =   492
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   372
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   2040
            Width           =   612
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÌœ"
            Height          =   372
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2040
            Width           =   852
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   372
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1680
            Width           =   492
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   372
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ Ê”ÿ"
            Height          =   372
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1680
            Width           =   852
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   372
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1320
            Width           =   492
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   372
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "÷⁄Ì› "
            Height          =   372
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1320
            Width           =   852
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ï"
            Height          =   375
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ï"
            Height          =   372
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   852
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   24
            Left            =   2196
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   120
            Visible         =   0   'False
            Width           =   768
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·ÌÊ„"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   5424
            TabIndex        =   26
            Top             =   120
            Width           =   744
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   336
            Index           =   8
            Left            =   8460
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   1188
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   750
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8925
         Width           =   10155
         _cx             =   17912
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   0
            Left            =   8940
            TabIndex        =   16
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":11F2
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   1
            Left            =   7830
            TabIndex        =   17
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":7A54
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   2
            Left            =   6600
            TabIndex        =   18
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":E2B6
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   3
            Left            =   5580
            TabIndex        =   19
            Top             =   120
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":14B18
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   4
            Left            =   4230
            TabIndex        =   20
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":1B37A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   6
            Left            =   1155
            TabIndex        =   21
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":21BDC
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   510
            Left            =   105
            TabIndex        =   22
            Top             =   120
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":4B7FE
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   7
            Left            =   3345
            TabIndex        =   23
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   900
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":52060
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   9
            Left            =   2055
            TabIndex        =   24
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   900
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
            ButtonImage     =   "FrmEvaluation_Standerd.frx":588C2
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
      End
      Begin C1SizerLibCtl.C1Elastic pnlGrid 
         Height          =   4020
         Left            =   120
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   4080
         Width           =   9855
         _cx             =   17383
         _cy             =   7091
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
         Begin VSFlex8Ctl.VSFlexGrid fg_Details 
            Height          =   3960
            Left            =   120
            TabIndex        =   61
            Top             =   0
            Width           =   9705
            _cx             =   17124
            _cy             =   6985
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
            FormatString    =   $"FrmEvaluation_Standerd.frx":5F124
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
      End
   End
End
Attribute VB_Name = "FrmEvaluation_Standerd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    
    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            ID.Text = CStr(new_id("TblEvaluationStandered", "ID", "", True))
            fg_Details.Rows = fg_Details.FixedRows
            fg_Details.Rows = fg_Details.FixedRows + 10
            Emp_Stude(0).value = True
            Emp_Stude_Click (0)
        Case 1
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "E"
            fg_Details.Rows = fg_Details.Rows + 1
        Case 2
            SaveData
        Case 3
            Undo
        Case 4
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_Action
        Case 5
        Case 6
            Unload Me
        Case 7
            print_report2
        Case 9
            Unload FrmInsurancesSearch
            FrmInsurancesSearch.SendForm = 3
            FrmInsurancesSearch.show
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Emp_Stude_Click(Index As Integer)
    pnlGrid.Visible = False
    C1Elastic2.Visible = False
    If Me.Emp_Stude(1).value = True Then
        C1Elastic2.Visible = True
    ElseIf Me.Emp_Stude(0).value = True Then
        pnlGrid.Visible = True
    End If
End Sub
Private Sub ExcelFrom_KeyPress(KeyAscii As Integer)
     KeyAscii = KeyAscii_Num(KeyAscii, Me.ExcelFrom.Text, 1)
End Sub
Private Sub ExcelTo_Change()
    If TxtModFlg <> "R" Then
        If val(ExcelTo.Text) > val(MaxDgree.Text) Then
            ExcelTo.Text = MaxDgree.Text
        End If
    End If
End Sub
Private Sub ExcelTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.ExcelTo.Text, 1)
End Sub
Private Sub fg_Details_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With fg_Details
        Select Case .ColKey(Col)
            Case "Allowance"
                .TextMatrix(Row, .ColIndex("AllowanceID")) = .ComboData
            Case Is = "InfluenceType"
                If .TextMatrix(Row, .ColIndex("InfluenceType")) = "+" Then
                    .TextMatrix(Row, .ColIndex("InfluenceTypeID")) = 1
                ElseIf .TextMatrix(Row, .ColIndex("InfluenceType")) = "-" Then
                    .TextMatrix(Row, .ColIndex("InfluenceTypeID")) = 2
                End If
        End Select
    End With
End Sub
Private Sub fg_Details_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With fg_Details
        Select Case .ColKey(Col)
            Case "Allowance"
                'Full Path Display
                'strSQL = " select * from mofrdat "
                StrSQL = " select * from mofrad  "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = fg_Details.BuildComboList(rs, "eq_sys, *name", "id")
                Else
                    StrComboList = fg_Details.BuildComboList(rs, "eq_sys, *namee", "id")
                End If
                Debug.Print StrSQL
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
            Case "InfluenceType"
                StrComboList = "+|-"
                .ComboList = StrComboList
            Case "AllowanceName"
                .ComboList = ""
            Case "Points"
                .ComboList = ""
        End Select
    End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Fill_Combos()
    Dim Dcombos As ClsDataCombos
    Dim str As String
  
    Set Dcombos = New ClsDataCombos
   
    Dcombos.GetBranches BranchID
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  „·› «·„œ«—”  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 '  Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    
    Dim StrSQL As String
    StrSQL = ""
    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = "SELECT  *  From TblEvaluationStandered    "
    Else
        StrSQL = "SELECT  *  From TblEvaluationStandered"
    End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub

ErrTrap:
End Sub
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
                SaveData
            Case vbCancel
                Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
    EleHeader.Caption = "Evaluation Standerds"
    
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(9).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"
    
    Label21.Caption = "Equal"
    Label20.Caption = "No Day"
    
    lbl(8).Caption = "Ser"
    Label3.Caption = "Todat Date"
    lbl(24).Caption = "Branch"
    Label1.Caption = "Arabic name"
    Label2.Caption = "English name"
    Label19.Caption = "Max Mark"
    
    Emp_Stude(0).Caption = "Employee"
    Emp_Stude(1).Caption = "Intern"
    
    Label4.Caption = "Poor"
    Label7.Caption = "Average"
    Label10.Caption = "Good"
    Label13.Caption = "Very Good"
    Label16.Caption = "Excellent"
    
    Label6.Caption = "From"
    Label9.Caption = "From"
    Label12.Caption = "From"
    Label15.Caption = "From"
    Label18.Caption = "From"
    
    Label5.Caption = "To"
    Label8.Caption = "To"
    Label11.Caption = "To"
    Label14.Caption = "To"
    Label17.Caption = "To"
    
    With fg_Details
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Allowance")) = "Bonus Name"
        .TextMatrix(0, .ColIndex("InfluenceType")) = "Influence Type"
        .TextMatrix(0, .ColIndex("Points")) = "Points"
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo ErrTrap
    
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If
        rs.Close
    End If
    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With fg_Details
        Select Case .ColKey(Col)
        End Select
    End With
End Sub
Private Sub GoodFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.GoodFrom.Text, 1)
End Sub
Private Sub GoodTo_Change()
    If TxtModFlg <> "R" Then
        If val(GoodTo.Text) > val(MaxDgree.Text) Then
            GoodTo.Text = MaxDgree.Text
        End If
    End If
End Sub
Private Sub GoodTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.GoodTo.Text, 1)
End Sub
Private Sub InterFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, InterFrom.Text, 1)
End Sub
Private Sub InterTo_Change()
    If TxtModFlg <> "R" Then
       If val(InterTo.Text) > val(MaxDgree.Text) Then
            InterTo.Text = MaxDgree.Text
        End If
    End If
End Sub
Private Sub InterTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.InterTo.Text, 1)
End Sub
Private Sub TxtModFlg_Change()
    
    On Error GoTo ErrTrap
    
    Select Case Me.TxtModFlg.Text
        Case "R"
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰"
            Else
                Me.Caption = "School  Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(9).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            ID.locked = True
      

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            pnlHeader.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰ ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«    ﬁÌÌ„ «·„ÊŸ›Ì‰ ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
            pnlHeader.Enabled = True
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«    ﬁÌÌ„ «·„ÊŸ›Ì‰ «·ÕÃ“ (  ⁄œÌ· )"
            Else
                Me.Caption = "Booking Request Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            ID.locked = True
           pnlHeader.Enabled = True
    End Select
    Exit Sub
ErrTrap:
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
    
    MySQL = MySQL & "SELECT dbo.TblEvaluationStandered.ID, dbo.TblEvaluationStandered.ProgrammID, dbo.TblEvaluationStandered.AirLineID, dbo.TblEvaluationStandered.AirPortID, dbo.TblEvaluationStandered.CompanyID,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.MekkaHotelID, TblHotels_2.Name AS MekkaHotelName, dbo.TblEvaluationStandered.MadinaHotelID, TblHotels_1.Name AS MadinaHotelName,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.JeddahHotelID, dbo.TblHotels.Name AS JeddahHotelName, dbo.TblEvaluationStandered.InClientID, TblCustemers_1.CusName AS InClientName,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.OutClientID, dbo.TblCustemers.CusName AS OutClientName, dbo.TblAirport.Name AS AirPortName, dbo.TblAirlines.Name AS AirLineName,"
    MySQL = MySQL & " dbo.TblTourismCompanies.Name AS CompanyName, dbo.TblBranchesData.branch_name AS BranchName, dbo.TblCompaniesGroup.Name AS GroupName,"
    MySQL = MySQL & " dbo.TblProgrammTypes.Name AS ProgrammName, dbo.TblEvaluationStandered.SDate, dbo.TblEvaluationStandered.BranchID, dbo.TblEvaluationStandered.FlightNo,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.emp, dbo.TblEvaluationStandered.GroupID, dbo.TblEvaluationStandered.other, dbo.TblEvaluationStandered.EmpID, dbo.TblEvaluationStandered.EmpName,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.EmpCode, dbo.TblEvaluationStandered.EmpMbile, convert ( char(10) , dbo.TblEvaluationStandered.ArriveTime ,108)  ArriveTime, dbo.TblEvaluationStandered.ArriveDate,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.VehicleNo, dbo.TblEvaluationStandered.Model, dbo.TblEvaluationStandered.VehicleType, dbo.TblEvaluationStandered.CreationUserID,"
    MySQL = MySQL & " dbo.TblEvaluationStandered.CreationDate, dbo.TblEvaluationStandered_Details.FromCity, dbo.TblEvaluationStandered_Details.TOCity, dbo.TblEvaluationStandered_Details.Date, dbo.TblEvaluationStandered_Details.HID,"
    MySQL = MySQL & " CONVERT(char(10), dbo.TblEvaluationStandered_Details.Time, 108) [Time] , dbo.TblEvaluationStandered_Details.CreationUserID AS Expr1, dbo.TblEvaluationStandered_Details.CreationDate AS Expr2, dbo.TblEvaluationStandered_Details.Remarks,"
    MySQL = MySQL & " dbo.TblCountriesGovernments.GovernmentName AS FromCityName, TblCountriesGovernments_1.GovernmentName AS ToCityName"
    MySQL = MySQL & " FROM dbo.TblCustemers INNER JOIN"
    MySQL = MySQL & " dbo.TblEvaluationStandered INNER JOIN"
    MySQL = MySQL & " dbo.TblProgrammTypes ON dbo.TblEvaluationStandered.ProgrammID = dbo.TblProgrammTypes.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblEvaluationStandered.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
    MySQL = MySQL & " dbo.TblCompaniesGroup ON dbo.TblEvaluationStandered.GroupID = dbo.TblCompaniesGroup.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblAirlines ON dbo.TblEvaluationStandered.AirLineID = dbo.TblAirlines.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblAirport ON dbo.TblEvaluationStandered.AirPortID = dbo.TblAirport.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblTourismCompanies ON dbo.TblEvaluationStandered.CompanyID = dbo.TblTourismCompanies.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblHotels AS TblHotels_2 ON dbo.TblEvaluationStandered.MekkaHotelID = TblHotels_2.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblHotels AS TblHotels_1 ON dbo.TblEvaluationStandered.MadinaHotelID = TblHotels_1.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblHotels ON dbo.TblEvaluationStandered.JeddahHotelID = dbo.TblHotels.ID INNER JOIN"
    MySQL = MySQL & " dbo.TblCustemers AS TblCustemers_1 ON dbo.TblEvaluationStandered.InClientID = TblCustemers_1.CusID ON"
    MySQL = MySQL & " dbo.TblCustemers.CusID = dbo.TblEvaluationStandered.OutClientID INNER JOIN"
    MySQL = MySQL & " dbo.TblEvaluationStandered_Details ON dbo.TblEvaluationStandered.ID = dbo.TblEvaluationStandered_Details.HID INNER JOIN"
    MySQL = MySQL & " dbo.TblCountriesGovernments ON dbo.TblEvaluationStandered_Details.FromCity = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
    MySQL = MySQL & " dbo.TblCountriesGovernments AS TblCountriesGovernments_1 ON dbo.TblEvaluationStandered_Details.TOCity = TblCountriesGovernments_1.GovernmentID"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequest.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequest.rpt"
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
            Msg = "There's no data to show"
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
Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
     If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
   
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    BranchID.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    EName.Text = IIf(IsNull(rs("EName").value), "", Trim(rs("EName").value))
    ENameE.Text = IIf(IsNull(rs("ENameE").value), "", Trim(rs("ENameE").value))
    
    WeakFrom.Text = IIf(IsNull(rs("WeakFrom").value), "", rs("WeakFrom").value)
    WeakTo.Text = IIf(IsNull(rs("WeakTo").value), "", rs("WeakTo").value)
    
    InterFrom.Text = IIf(IsNull(rs("InterFrom").value), "", rs("InterFrom").value)
    InterTo.Text = IIf(IsNull(rs("InterTo").value), "", rs("InterTo").value)
    
    GoodFrom.Text = IIf(IsNull(rs("GoodFrom").value), "", rs("GoodFrom").value)
    GoodTo.Text = IIf(IsNull(rs("GoodTo").value), "", rs("GoodTo").value)
    
    VeryGFrom.Text = IIf(IsNull(rs("VeryGFrom").value), "", rs("VeryGFrom").value)
    VeryGTo.Text = IIf(IsNull(rs("VeryGTo").value), "", rs("VeryGTo").value)
    
    ExcelFrom.Text = IIf(IsNull(rs("ExcelFrom").value), "", rs("ExcelFrom").value)
    ExcelTo.Text = IIf(IsNull(rs("ExcelTo").value), "", rs("ExcelTo").value)
    MaxDgree.Text = IIf(IsNull(rs("MaxDgree").value), "", rs("MaxDgree").value)

    TxtNoDayAbcen.Text = IIf(IsNull(rs("NoDayAbcen").value), 0, rs("NoDayAbcen").value)
    TxtAvgAbscen.Text = IIf(IsNull(rs("AvgAbscen").value), "", rs("AvgAbscen").value)
    If Not IsNull(rs("Emp_Stude").value) Then
    If rs("Emp_Stude").value = 1 Then
    Emp_Stude(1).value = True
    Else
    Emp_Stude(0).value = True
    End If
    Else
    Emp_Stude(0).value = True
    End If
    
    Set Rs_Temp = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = " SELECT H.* , name  from  TblEvaluationStandered_Details H , mofrad D "
    StrSQL = StrSQL & " Where AllowanceID = D.ID and  H.HID = " & val(ID.Text)
    
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst
        With fg_Details
            .Rows = Rs_Temp.RecordCount + 1
            Dim j As Integer
            For j = 1 To .Rows - 1
                .TextMatrix(j, .ColIndex("id")) = IIf(IsNull(Rs_Temp("id").value), "", Rs_Temp("id").value)
                .TextMatrix(j, .ColIndex("hid")) = IIf(IsNull(Rs_Temp("hid").value), 0, Rs_Temp("hid").value)
                
                .TextMatrix(j, .ColIndex("AllowanceID")) = IIf(IsNull(Rs_Temp("AllowanceID").value), "", Rs_Temp("AllowanceID").value)
                .TextMatrix(j, .ColIndex("Allowance")) = IIf(IsNull(Rs_Temp("name").value), "", Rs_Temp("name").value)
                
                .TextMatrix(j, .ColIndex("AllowanceName")) = IIf(IsNull(Rs_Temp("AllowanceName").value), "", Rs_Temp("AllowanceName").value)
                
                .TextMatrix(j, .ColIndex("InfluenceTypeID")) = IIf(IsNull(Rs_Temp("InfluenceType").value), "", Rs_Temp("InfluenceType").value)
            
                If val(.TextMatrix(j, .ColIndex("InfluenceTypeID"))) = 1 Then
                    .TextMatrix(j, .ColIndex("InfluenceType")) = "+"
                ElseIf val(.TextMatrix(j, .ColIndex("InfluenceTypeID"))) = 2 Then
                    .TextMatrix(j, .ColIndex("InfluenceType")) = "-"
                End If
                .TextMatrix(j, .ColIndex("Points")) = IIf(IsNull(Rs_Temp("Points").value), "", Rs_Temp("Points").value)
                Rs_Temp.MoveNext
            Next
        End With
    End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub EName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub ENameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
Private Sub VeryGFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.VeryGFrom.Text, 1)
End Sub
Private Sub VeryGTo_Change()
    If TxtModFlg <> "R" Then
        If val(VeryGTo.Text) > val(MaxDgree.Text) Then
            VeryGTo.Text = MaxDgree.Text
        End If
    End If
End Sub
Private Sub VeryGTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.VeryGTo.Text, 1)
End Sub
Private Sub WeakFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.WeakFrom.Text, 1)
End Sub
Private Sub WeakTo_Change()
    If TxtModFlg <> "R" Then
        If val(WeakTo.Text) > val(MaxDgree.Text) Then
            WeakTo.Text = MaxDgree.Text
        End If
End If
End Sub
Private Sub WeakTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.WeakTo.Text, 1)
End Sub
Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        fg_Details.Rows = fg_Details.FixedRows
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
        fg_Details.Rows = fg_Details.FixedRows
    End If

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
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Trim(EName.Text) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Standered Name "
            Else
                Msg = "«œŒ· «”„ «·„⁄Ì«— «Ê·« "
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            EName.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text
           Case "N"
                rs.AddNew
                ID.Text = CStr(new_id("TblEvaluationStandered", "ID", "", True))
           Case "E"
                StrSQL = "delete From TblEvaluationStandered_Details where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
        End Select
        
        rs("ID").value = val(ID.Text)
        rs("SDate").value = SDate.value
        rs("BranchID").value = IIf(BranchID.BoundText = "", Null, BranchID.BoundText)
        rs("EName").value = IIf(EName.Text = "", Null, EName.Text)
        rs("ENameE").value = IIf(ENameE.Text = "", Null, ENameE.Text)
        
        rs("WeakFrom").value = val(WeakFrom.Text)
        rs("WeakTo").value = val(WeakTo.Text)
        rs("InterFrom").value = val(InterFrom.Text)
        rs("InterTo").value = val(InterTo.Text)
        rs("GoodFrom").value = val(GoodFrom.Text)
        rs("GoodTo").value = val(GoodTo.Text)
        rs("VeryGFrom").value = val(VeryGFrom.Text)
        rs("VeryGTo").value = val(VeryGTo.Text)
        rs("ExcelFrom").value = val(ExcelFrom.Text)
        rs("ExcelTo").value = val(ExcelTo.Text)
        rs("MaxDgree").value = val(MaxDgree.Text)
        ''/
        rs("NoDayAbcen").value = val(TxtNoDayAbcen.Text)
        rs("AvgAbscen").value = val(TxtAvgAbscen.Text)
        If Emp_Stude(1).value = True Then
            rs("Emp_Stude").value = 1
        Else
            rs("Emp_Stude").value = 0
        End If
        rs("creationdate").value = Date
        rs("creationuserID").value = user_id
      
        rs.update
        
        
        Dim Rs_Temp As ADODB.Recordset
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = " select * from TblEvaluationStandered_Details  where 1 = -1 "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With fg_Details
            Dim j As Integer
            For j = 1 To fg_Details.Rows - 1
                If .TextMatrix(j, .ColIndex("AllowanceID")) <> "" Then
                    Rs_Temp.AddNew
                    Rs_Temp("ID") = CStr(new_id("TblEvaluationStandered_Details", "ID", "", True))
                    Rs_Temp("HID") = val(ID.Text)
                    Rs_Temp("AllowanceID") = val(.TextMatrix(j, .ColIndex("AllowanceID")))
                    Rs_Temp("AllowanceName") = .TextMatrix(j, .ColIndex("AllowanceName"))
                    Rs_Temp("InfluenceType") = val(.TextMatrix(j, .ColIndex("InfluenceTypeID")))
                    Rs_Temp("Points") = val(.TextMatrix(j, .ColIndex("Points")))
                
                    Rs_Temp.update
                End If
            Next
        End With
        
        Dim StrDes As String

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        Select Case Me.TxtModFlg.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰ " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text
        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)
            fg_Details.Rows = fg_Details.FixedRows
        Case "E"
            rs.find " ID='" & val(ID.Text) & "'", , adSearchForward, adBookmarkFirst
            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Action()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
 
    On Error GoTo ErrTrap
            
    If ID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰  —ﬁ„ " & CHR(13)
            Msg = Msg + (ID.Text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Delete Booking Request File ? " & CHR(13)
            Msg = Msg + (ID.Text) & CHR(13)
            Msg = Msg + "  Are you sure you want to delete ?"
        End If
         
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblEvaluationStandered_Details where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "delete From TblEvaluationStandered where  ID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                    
                StrSQL = "SELECT  *  From TblEvaluationStandered "
                rs.Close
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If
    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "this process Not Aailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If
    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ…  ﬁÌÌ„ «·„ÊŸ›Ì‰ "
    Else
    Msg = "Sorry can't delete data" & CHR(13) & "for its integration with employee evaluation"
    End If
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub
Private Sub AddTip()
    Dim Wrap As String
    
    On Error GoTo ErrTrap
    
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–«  ﬁÌÌ„ «·„ÊŸ›Ì‰" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With
    Exit Sub
ErrTrap:
End Sub
