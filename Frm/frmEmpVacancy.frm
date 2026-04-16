VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpContract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄ﬁÊœ «·„ÊŸ›Ì‰"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13155
   Icon            =   "frmEmpVacancy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   13155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   705
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13155
      _cx             =   23204
      _cy             =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   178
         Weight          =   700
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
      Caption         =   "  ⁄ﬁÊœ «·„ÊŸ›Ì‰    "
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
         Height          =   135
         Left            =   0
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   600
         Width           =   12375
         _cx             =   21828
         _cy             =   238
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
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmEmpVacancy.frx":6852
         ColorButton     =   16777215
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
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmEmpVacancy.frx":6BEC
         ColorButton     =   16777215
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
         Left            =   1800
         TabIndex        =   12
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmEmpVacancy.frx":6F86
         ColorButton     =   16777215
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
         Left            =   645
         TabIndex        =   13
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmEmpVacancy.frx":7320
         ColorButton     =   16777215
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblnfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   120
         Width           =   4695
      End
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   8535
      Left            =   0
      TabIndex        =   22
      Top             =   480
      Width           =   13185
      _cx             =   23257
      _cy             =   15055
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
      Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«  «·«Ã«“«  «·„” Õﬁ…|«· ÃœÌœ «· ·ﬁ«∆Ì ··⁄ﬁœ"
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
      Picture(0)      =   "frmEmpVacancy.frx":76BA
      Picture(1)      =   "frmEmpVacancy.frx":DF1C
      Picture(2)      =   "frmEmpVacancy.frx":1477E
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   8070
         Left            =   14130
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   45
         Width           =   13095
         _cx             =   23098
         _cy             =   14235
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
         Begin VB.Frame Frame11 
            Height          =   8070
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   0
            Width           =   13095
            Begin VB.Frame Frame12 
               Height          =   735
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   7200
               Width           =   6375
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   330
                  Left            =   4200
                  TabIndex        =   140
                  ToolTipText     =   " ÕœÌœ «·ﬂ·"
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmEmpVacancy.frx":1AFE0
                  ButtonImageDisabled=   "frmEmpVacancy.frx":21842
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton5 
                  Height          =   330
                  Left            =   840
                  TabIndex        =   141
                  ToolTipText     =   " ÕœÌœ «·ﬂ·"
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "«·€«¡ «· ÕœÌœ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmEmpVacancy.frx":40A2C
                  ButtonImageDisabled=   "frmEmpVacancy.frx":4728E
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VB.Frame Frame13 
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   7200
               Width           =   6375
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   330
                  Left            =   4080
                  TabIndex        =   137
                  ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·’› «·Õ«·Ì"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmEmpVacancy.frx":66478
                  ButtonImageDisabled=   "frmEmpVacancy.frx":6CCDA
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   330
                  Left            =   1080
                  TabIndex        =   138
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   240
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ· "
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmEmpVacancy.frx":8BEC4
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   6915
               Left            =   120
               TabIndex        =   142
               Top             =   240
               Width           =   12915
               _cx             =   22781
               _cy             =   12197
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
               Rows            =   2
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEmpVacancy.frx":92726
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
                  Left            =   600
                  TabIndex        =   143
                  Top             =   4320
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8070
         Index           =   2
         Left            =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   13095
         _cx             =   23098
         _cy             =   14235
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   8415
            Left            =   0
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   13215
            _cx             =   23310
            _cy             =   14843
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
               Height          =   8415
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   0
               Width           =   13095
               Begin VB.ComboBox DcbStatus 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":9280C
                  Left            =   3600
                  List            =   "frmEmpVacancy.frx":92816
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ÃœÌœ  ·ﬁ«∆Ì ··⁄ﬁœ"
                  Height          =   255
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   2160
                  Width           =   2055
               End
               Begin VB.ComboBox DataCombo5 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":92824
                  Left            =   0
                  List            =   "frmEmpVacancy.frx":92826
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   98
                  Top             =   960
                  Width           =   1575
               End
               Begin VB.Frame Frame8 
                  Caption         =   " –«ﬂ— «·”›— «·”‰ÊÌ…"
                  Height          =   1695
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   6240
                  Width           =   6855
                  Begin VB.TextBox TxtTicketValue 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   600
                     Width           =   615
                  End
                  Begin VB.TextBox no_of_Child_ticket 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   1320
                     Width           =   615
                  End
                  Begin VB.CheckBox Child_ticket 
                     Alignment       =   1  'Right Justify
                     Caption         =   " –«ﬂ— ·· «»⁄Ì‰ "
                     Height          =   255
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   1320
                     Width           =   2175
                  End
                  Begin VB.CheckBox wife_ticket 
                     Alignment       =   1  'Right Justify
                     Caption         =   " –ﬂ—… ··“ÊÃ…"
                     Height          =   255
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   960
                     Width           =   2175
                  End
                  Begin VB.CheckBox have_ticket 
                     Alignment       =   1  'Right Justify
                     Caption         =   " –ﬂ—… ·Â"
                     Height          =   255
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   600
                     Width           =   2055
                  End
                  Begin MSDataListLib.DataCombo DataCombo1 
                     Height          =   315
                     Left            =   3960
                     TabIndex        =   89
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo2 
                     Height          =   315
                     Left            =   3960
                     TabIndex        =   90
                     Top             =   960
                     Visible         =   0   'False
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo3 
                     Height          =   315
                     Left            =   3960
                     TabIndex        =   91
                     Top             =   1320
                     Visible         =   0   'False
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label27 
                     Alignment       =   1  'Right Justify
                     Caption         =   "»ﬁÌ„…"
                     Height          =   255
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   600
                     Width           =   495
                  End
                  Begin VB.Shape Shape3 
                     BorderWidth     =   2
                     Height          =   1215
                     Left            =   120
                     Top             =   360
                     Width           =   1455
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0FFFF&
                     Caption         =   " –«ﬂ— «·”›— €Ì— „Õœœ… «·ﬁÌ„… Ê–·ﬂ ·«Œ ·«› ﬁÌ„ Â« „‰ › —… «·Ï «Œ—Ï"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000080&
                     Height          =   1020
                     Index           =   5
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   450
                     Width           =   1215
                  End
                  Begin VB.Label Label26 
                     Alignment       =   1  'Right Justify
                     Caption         =   "»⁄œœ"
                     Height          =   255
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   1320
                     Width           =   495
                  End
                  Begin VB.Label Label25 
                     Alignment       =   2  'Center
                     Caption         =   "«·œ—Ã…"
                     Height          =   255
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   855
                  End
                  Begin VB.Label Label24 
                     Alignment       =   1  'Right Justify
                     Height          =   375
                     Left            =   4800
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   360
                     Width           =   1815
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„·«ÕŸ… Â«„…:-"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000080&
                     Height          =   255
                     Index           =   6
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   120
                     Width           =   1275
                  End
               End
               Begin VB.Frame Frame6 
                  Caption         =   "„Œ’’ ‰Â«Ì… «·Œœ„…"
                  Height          =   2535
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   4200
                  Width           =   5415
                  Begin VB.Frame Frame7 
                     Caption         =   "‰”»…"
                     Height          =   2295
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   240
                     Width           =   5295
                     Begin VSFlex8Ctl.VSFlexGrid Grid2 
                        Height          =   1380
                        Left            =   120
                        TabIndex        =   79
                        Top             =   480
                        Width           =   5040
                        _cx             =   8890
                        _cy             =   2434
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
                        Rows            =   1
                        Cols            =   5
                        FixedRows       =   1
                        FixedCols       =   0
                        RowHeightMin    =   300
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"frmEmpVacancy.frx":92828
                        ScrollTrack     =   0   'False
                        ScrollBars      =   2
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
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   390
                        Index           =   8
                        Left            =   3600
                        TabIndex        =   80
                        Top             =   1800
                        Width           =   1410
                        _ExtentX        =   2487
                        _ExtentY        =   688
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "Õ–› ”ÿ—"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ButtonImage     =   "frmEmpVacancy.frx":928DD
                        DrawFocusRectangle=   0   'False
                     End
                     Begin VB.Label Label21 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Õœœ «·„›—œ«   «· Ì   ” œŒ· ›Ì Õ”«» ‰Â«Ì… «·Œœ„…"
                        Height          =   375
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   81
                        Top             =   240
                        Width           =   4935
                     End
                  End
                  Begin VB.Label Label22 
                     Alignment       =   1  'Right Justify
                     Caption         =   "%"
                     Height          =   255
                     Left            =   600
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   840
                     Width           =   735
                  End
               End
               Begin VB.TextBox Due_period_no 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   9840
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   2160
                  Width           =   855
               End
               Begin VB.TextBox Holiday_period_no 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10560
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   5880
                  Width           =   735
               End
               Begin VB.TextBox Contract_ID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10200
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.TextBox emp_code 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.ComboBox Contract_period 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":92E77
                  Left            =   3600
                  List            =   "frmEmpVacancy.frx":92E81
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   240
                  Width           =   975
               End
               Begin VB.TextBox Contract_period_no 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   240
                  Width           =   495
               End
               Begin VB.ComboBox test_period 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":92E8F
                  Left            =   0
                  List            =   "frmEmpVacancy.frx":92E99
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.TextBox test_period_no 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   600
                  Width           =   855
               End
               Begin VB.ComboBox due_period 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":92EA7
                  Left            =   8760
                  List            =   "frmEmpVacancy.frx":92EB4
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.ComboBox Holiday_period 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":92EC7
                  Left            =   9600
                  List            =   "frmEmpVacancy.frx":92ED1
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   5880
                  Width           =   975
               End
               Begin VB.TextBox salary_period_no 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   9840
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   2760
                  Width           =   855
               End
               Begin VB.ComboBox salary_period 
                  Height          =   315
                  ItemData        =   "frmEmpVacancy.frx":92EDF
                  Left            =   8760
                  List            =   "frmEmpVacancy.frx":92EE9
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   2760
                  Width           =   975
               End
               Begin VB.Frame Frame1 
                  Caption         =   "—« »"
                  Height          =   2655
                  Left            =   7200
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   3120
                  Width           =   5295
                  Begin VSFlex8Ctl.VSFlexGrid Grid 
                     Height          =   1620
                     Left            =   120
                     TabIndex        =   62
                     Top             =   480
                     Width           =   5040
                     _cx             =   8890
                     _cy             =   2857
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
                     Rows            =   1
                     Cols            =   5
                     FixedRows       =   1
                     FixedCols       =   0
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"frmEmpVacancy.frx":92EF7
                     ScrollTrack     =   0   'False
                     ScrollBars      =   2
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
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   9
                     Left            =   3720
                     TabIndex        =   63
                     Top             =   2040
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ–› ”ÿ—"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "frmEmpVacancy.frx":92FAC
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Õœœ „›—œ«  „” Õﬁ«  «·«Ã«“…"
                     Height          =   375
                     Left            =   840
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   240
                     Width           =   4335
                  End
               End
               Begin VB.OptionButton salary_or_fixed_value 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ﬁÌ„… À«» …"
                  Height          =   255
                  Index           =   1
                  Left            =   7200
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   2760
                  Width           =   1335
               End
               Begin VB.TextBox Fixed_value 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   2760
                  Width           =   1275
               End
               Begin VB.OptionButton salary_or_fixed_value 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—« »"
                  Height          =   255
                  Index           =   0
                  Left            =   11040
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   2760
                  Width           =   1455
               End
               Begin VB.Frame Frame2 
                  Caption         =   "«·“Ì«œ… «·”‰ÊÌ…"
                  Height          =   2775
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   1320
                  Width           =   5415
                  Begin VB.OptionButton yearly_increase_fixed_value_or_percentage 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬁÌ„… À«» …"
                     Height          =   255
                     Index           =   0
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   360
                     Width           =   1455
                  End
                  Begin VB.TextBox yearly_increase 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   0
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   360
                     Width           =   975
                  End
                  Begin VB.TextBox yearly_increase 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   1
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   360
                     Width           =   855
                  End
                  Begin VB.OptionButton yearly_increase_fixed_value_or_percentage 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰”»…"
                     Height          =   255
                     Index           =   1
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   360
                     Width           =   1095
                  End
                  Begin VB.Frame Frame3 
                     Caption         =   "‰”»…"
                     Height          =   2175
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   600
                     Width           =   5295
                     Begin VSFlex8Ctl.VSFlexGrid Grid1 
                        Height          =   1260
                        Left            =   120
                        TabIndex        =   50
                        Top             =   480
                        Width           =   5040
                        _cx             =   8890
                        _cy             =   2222
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
                        Rows            =   1
                        Cols            =   5
                        FixedRows       =   1
                        FixedCols       =   0
                        RowHeightMin    =   300
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"frmEmpVacancy.frx":93546
                        ScrollTrack     =   0   'False
                        ScrollBars      =   2
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
                     Begin ImpulseButton.ISButton Cmd 
                        Height          =   390
                        Index           =   7
                        Left            =   3600
                        TabIndex        =   51
                        Top             =   1680
                        Width           =   1410
                        _ExtentX        =   2487
                        _ExtentY        =   688
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "Õ–› ”ÿ—"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ButtonImage     =   "frmEmpVacancy.frx":935FB
                        DrawFocusRectangle=   0   'False
                     End
                     Begin VB.Label Label14 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Õœœ «·„›—œ«   «· Ì ”Ì „ «·“Ì«œ… ⁄·ÌÂ«"
                        Height          =   255
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   52
                        Top             =   240
                        Width           =   4695
                     End
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "%"
                     Height          =   255
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   360
                     Width           =   135
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "«· √„Ì‰ «·ÿ»Ì"
                  Height          =   1455
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   6720
                  Width           =   5415
                  Begin VB.TextBox TxtInsuranceNO 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2160
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   120
                     Width           =   1575
                  End
                  Begin VB.CheckBox have_insurance 
                     Alignment       =   1  'Right Justify
                     Caption         =   " √„Ì‰ ·Â"
                     Height          =   255
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   720
                     Width           =   1215
                  End
                  Begin VB.CheckBox wife_insurance 
                     Alignment       =   1  'Right Justify
                     Caption         =   " √„Ì‰ ··“ÊÃ…"
                     Height          =   255
                     Left            =   1080
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   720
                     Width           =   1335
                  End
                  Begin VB.CheckBox Child_insurance 
                     Alignment       =   1  'Right Justify
                     Caption         =   " √„Ì‰ ·· «»⁄Ì‰  "
                     Height          =   255
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   1080
                     Width           =   1455
                  End
                  Begin VB.TextBox no_of_Child_insurance 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   1080
                     Width           =   735
                  End
                  Begin MSDataListLib.DataCombo have_insurance_class 
                     Height          =   315
                     Left            =   3000
                     TabIndex        =   41
                     Top             =   720
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo wife_insurance_class 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   42
                     Top             =   720
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo Child_insurance_class 
                     Height          =   315
                     Left            =   3000
                     TabIndex        =   43
                     Top             =   1080
                     Width           =   735
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label28 
                     Alignment       =   2  'Center
                     Caption         =   "«·—ﬁ„ «· √„Ì‰Ì"
                     Height          =   255
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.Label Label23 
                     Alignment       =   2  'Center
                     Caption         =   "«·›∆…"
                     Height          =   255
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   480
                     Width           =   855
                  End
                  Begin VB.Label Label15 
                     Alignment       =   2  'Center
                     Caption         =   "«·›∆…"
                     Height          =   255
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   480
                     Width           =   855
                  End
                  Begin VB.Label Label16 
                     Alignment       =   1  'Right Justify
                     Caption         =   "»⁄œœ"
                     Height          =   255
                     Left            =   1080
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   1080
                     Width           =   1215
                  End
               End
               Begin VB.TextBox Emp_id 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.TextBox emp_name 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   0
                  Left            =   7440
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.TextBox emp_name 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   1
                  Left            =   6240
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.TextBox emp_name 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   2
                  Left            =   4920
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.TextBox emp_name 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   3
                  Left            =   3600
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.TextBox XPTxtEmpName 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Text            =   "Text1"
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Text            =   "TxtModFlg"
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Command1"
                  Height          =   195
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   1680
                  Width           =   375
               End
               Begin MSComCtl2.DTPicker Contract_date 
                  Height          =   315
                  Left            =   7560
                  TabIndex        =   99
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   96600065
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker Issue_date 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   100
                  Top             =   1680
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   96600065
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker Holiday_date 
                  Height          =   315
                  Left            =   6480
                  TabIndex        =   101
                  Top             =   5880
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   96600065
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo Departement 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   102
                  Top             =   1320
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo job 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   103
                  Top             =   960
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo Contract_type 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   104
                  Top             =   960
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   96600065
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
                  Height          =   255
                  Left            =   6240
                  TabIndex        =   106
                  Top             =   240
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateHigri1 
                  Height          =   255
                  Left            =   0
                  TabIndex        =   107
                  Top             =   240
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
               End
               Begin ImpulseButton.ISButton ShowTab 
                  Height          =   255
                  Left            =   5760
                  TabIndex        =   131
                  Top             =   2160
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  Caption         =   "<<<<<"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmEmpVacancy.frx":93B95
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin ImpulseButton.ISButton ISButton1 
                  Height          =   255
                  Left            =   5640
                  TabIndex        =   132
                  Top             =   5880
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  Caption         =   "<<<<<"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmEmpVacancy.frx":9A3F7
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  Caption         =   "Õ«·… «·⁄ﬁœ"
                  Height          =   375
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   960
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   2
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   240
                  Width           =   915
               End
               Begin VB.Label Label30 
                  Alignment       =   2  'Center
                  Caption         =   "‰Ê⁄ «·⁄ﬁœ"
                  Height          =   255
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label29 
                  Alignment       =   2  'Center
                  Caption         =   "‰Â«Ì… «·⁄ﬁœ"
                  Height          =   255
                  Left            =   2520
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "ﬂ· «·„›—œ«  «·„Œ «—… ” œŒ· ÷„‰ «Õ ”«» „Œ’’ «·«Ã«“…"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   1140
                  Index           =   25
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   3810
                  Width           =   1215
               End
               Begin VB.Shape Shape2 
                  BorderWidth     =   2
                  Height          =   1335
                  Left            =   5640
                  Top             =   3720
                  Width           =   1455
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H0080C0FF&
                  Height          =   6015
                  Left            =   5520
                  Top             =   2040
                  Width           =   7215
               End
               Begin VB.Label lbl 
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„” Õﬁ«  «·„«œÌ… ⁄‰ ﬂ· "
                  Height          =   285
                  Index           =   1
                  Left            =   10800
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   2160
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «Ì«„ «·«Ã«“…"
                  Height          =   285
                  Index           =   0
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   5880
                  Width           =   1125
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
                  Caption         =   "—ﬁ„ «·⁄ﬁœ"
                  Height          =   375
                  Left            =   11760
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  Caption         =   "»œ«Ì… «·⁄ﬁœ"
                  Height          =   255
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„›—œ«  «· Ì ÌÕ’· ⁄·ÌÂ« «·„ÊŸ›"
                  Height          =   375
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   1680
                  Width           =   3735
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  Caption         =   "ﬂÊœ «·„ÊŸ›"
                  Height          =   375
                  Left            =   11760
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Caption         =   "«”„ «·„ÊŸ›"
                  Height          =   255
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
                  Height          =   375
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œÂ «· ⁄«ﬁœ"
                  Height          =   375
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label10 
                  Alignment       =   2  'Center
                  Caption         =   "„œ… «·«Œ »«—"
                  Height          =   255
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   375
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  Caption         =   "«·ﬁ”„"
                  Height          =   255
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   " «—ÌŒ „»«‘—… «·⁄„·"
                  Height          =   375
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   1680
                  Width           =   1815
               End
               Begin VB.Line Line1 
                  X1              =   5760
                  X2              =   12600
                  Y1              =   2640
                  Y2              =   2640
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Caption         =   " «—ÌŒ «” Õﬁ«ﬁ  «·«Ã«“…"
                  Height          =   255
                  Left            =   7920
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   5880
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   375
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„·«ÕŸ… Â«„…:-"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Index           =   24
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   3360
                  Width           =   1275
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   8070
         Left            =   13830
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   13095
         _cx             =   23098
         _cy             =   14235
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
         Begin VSFlex8Ctl.VSFlexGrid gridHolidayDue 
            Height          =   7635
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   12945
            _cx             =   22834
            _cy             =   13467
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmEmpVacancy.frx":A0C59
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
      End
   End
   Begin VB.TextBox Basic_salary 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13920
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame10 
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9000
      Width           =   13095
      Begin VB.Label XPTxtCount 
         Alignment       =   2  'Center
         Caption         =   "Label19"
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·⁄ﬁÊœ"
         Height          =   255
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   2  'Center
         Caption         =   "XPTxtCurrent"
         Height          =   255
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄ﬁœ «·Õ«·Ì"
         Height          =   255
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   9360
      Width           =   13095
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   11880
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÃœÌœ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Height          =   375
         Index           =   1
         Left            =   10560
         TabIndex        =   2
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ⁄œÌ·"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   9120
         TabIndex        =   3
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ›Ÿ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Height          =   375
         Index           =   3
         Left            =   7560
         TabIndex        =   4
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " —«Ã⁄"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   6120
         TabIndex        =   5
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
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
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton CmdHelp 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "„”«⁄œ…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   6
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   420
         Index           =   10
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   741
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
   End
End
Attribute VB_Name = "frmEmpContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim EmpReport As ClsEmployeeReport
Dim xReport As New CRAXDRT.Report
Dim NO As Double
Private objScript As Object
Dim case_id As Integer
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
    With Me.Grid
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Mofradtype")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If
        Next i
    End With
    IntCounter = 0
    With Me.GRID1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Mofradtype")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If
        Next i
    End With
    IntCounter = 0
    With Me.GRID2
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Mofradtype")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Function GetDefaultComponents()
    Dim LBLWhereSTR1 As String
    Dim LBLWhereSTR2 As String
    Dim SQLStr As String
    Dim rscomponent As New ADODB.Recordset
    LBLWhereSTR1 = GetComponentIncalculations(0)
'    SQLSTR = "SELECT     id, name, nameE"
'    SQLSTR = SQLSTR & " from dbo.MOFRAD"
'    SQLSTR = SQLSTR & " WHERE     (id IN (" & LBLWhereSTR1 & "))"
 
 
           SQLStr = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
                SQLStr = SQLStr & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
                SQLStr = SQLStr & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
                SQLStr = SQLStr & " Where (  (dbo.mofrad.Aloc1 = 1) and dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.Text) & ")"
                SQLStr = SQLStr & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
 
    rscomponent.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.Grid
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Mofradtype")) = IIf(IsNull(rscomponent("mofrad_type").value), "", rscomponent("mofrad_type").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("name").value), "", rscomponent("name").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("namee").value), "", rscomponent("namee").value)
                End If
            
                rscomponent.MoveNext
            Next
   
        End With

    End If

    rscomponent.Close

    LBLWhereSTR2 = GetComponentIncalculations(1)
'    SQLSTR = "SELECT     id, name, nameE"
'    SQLSTR = SQLSTR & " from dbo.MOFRAD"
'    SQLSTR = SQLSTR & " WHERE     (id IN (" & LBLWhereSTR2 & "))"
    
    
           SQLStr = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
                SQLStr = SQLStr & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
                SQLStr = SQLStr & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
                SQLStr = SQLStr & " Where (  (dbo.mofrad.Aloc2 = 1) and dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.Text) & ")"
                SQLStr = SQLStr & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
     
     
    rscomponent.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.GRID2
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Mofradtype")) = IIf(IsNull(rscomponent("mofrad_type").value), "", rscomponent("mofrad_type").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("name").value), "", rscomponent("name").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("namee").value), "", rscomponent("namee").value)
                End If
                        
                rscomponent.MoveNext
            Next i
               
        End With

    End If



    rscomponent.Close
 
    
           SQLStr = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
                SQLStr = SQLStr & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
                SQLStr = SQLStr & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
                SQLStr = SQLStr & " Where (  (dbo.mofrad.InCrease = 1) and dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.Text) & ")"
                SQLStr = SQLStr & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
     
     
    rscomponent.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.GRID1
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Mofradtype")) = IIf(IsNull(rscomponent("mofrad_type").value), "", rscomponent("mofrad_type").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("name").value), "", rscomponent("name").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("namee").value), "", rscomponent("namee").value)
                End If
                        
                rscomponent.MoveNext
            Next i
               
        End With

    End If


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


    MySQL = " SELECT     TOP 1 TblEmployee_1.Emp_ID, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee, TblEmployee_1.Nationality, "
MySQL = MySQL & "  TblEmployee_1.dean, TblEmployee_1.NumPasp, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, TblEmployee_1.DateEndPasp,"
MySQL = MySQL & "  TblEmployee_1.DateExpPasp, TblEmployee_1.pasplace, dbo.Contract.Contract_period_no, dbo.Contract.Contract_period, dbo.Contract.test_period_no,"
MySQL = MySQL & "  dbo.Contract.test_period, dbo.EmpSalaryComponent.[Value], dbo.Contract.Due_period_no, dbo.Contract.due_period, dbo.Contract.Holiday_period_no,"
MySQL = MySQL & "  dbo.Contract.Holiday_period , dbo.Contract.Contract_ID,  TblEmployee_1.IssueDateH, TblEmployee_1.BignDateWork"
MySQL = MySQL & " FROM         dbo.TblEmployee TblEmployee_1 INNER JOIN"
MySQL = MySQL & "  dbo.TblEmpJobsTypes ON TblEmployee_1.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID INNER JOIN"
MySQL = MySQL & "  dbo.Contract ON TblEmployee_1.Emp_ID = dbo.Contract.Emp_id INNER JOIN"
MySQL = MySQL & "   dbo.EmpSalaryComponent ON TblEmployee_1.Emp_ID = dbo.EmpSalaryComponent.emp_ID"
MySQL = MySQL & " Where (dbo.Contract.Contract_ID = " & val(Contract_ID) & ")"
MySQL = MySQL & " ORDER BY dbo.EmpSalaryComponent.id"
 
 MySQL = " SELECT   TblEmployee_1.Emp_ID, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee, TblEmployee_1.Nationality, TblEmployee_1.dean, TblEmployee_1.NumPasp,"
 MySQL = MySQL & "    dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, TblEmployee_1.DateEndPasp, TblEmployee_1.DateExpPasp, TblEmployee_1.pasplace,"
 MySQL = MySQL & "    dbo.Contract.Contract_period_no, dbo.Contract.Contract_period, dbo.Contract.test_period_no, dbo.Contract.test_period, dbo.EmpSalaryComponent.[Value], dbo.Contract.Due_period_no,"
 MySQL = MySQL & "   dbo.Contract.due_period, dbo.Contract.Holiday_period_no, dbo.Contract.Holiday_period, dbo.Contract.Contract_ID, TblEmployee_1.IssueDateH, TblEmployee_1.BignDateWork, dbo.Contract.DateH,"
 MySQL = MySQL & "    dbo.Contract.Contract_date, dbo.EmpSalaryComponent.mofrad_type, dbo.EmpSalaryComponent.AccountName, dbo.EmpSalaryComponent.AccountCode, dbo.EmpSalaryComponent.des,"
 MySQL = MySQL & "    dbo.EmpSalaryComponent.specific_value , dbo.Contract.DateH1"
 MySQL = MySQL & "  , TblEmployee_1.NumEkama, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name3, "
  MySQL = MySQL & "                      TblEmployee_1.Emp_Name4, TblEmployee_1.Emp_Name2, TblEmployee_1.DateEndekamah, TblEmployee_1.DateExpoekamaH, TblEmployee_1.KafelName,"
  MySQL = MySQL & "                      TblEmployee_1.hdoddate, TblEmployee_1.hdodno, TblEmployee_1.hdomnfaz, TblEmployee_1.jopstatusid, TblEmployee_1.workstate, TblEmployee_1.Emp_Mail,"
    MySQL = MySQL & "                    TblEmployee_1.Emp_mobile, TblEmployee_1.Emp_Remark , TblEmployee_1.placeEkama      FROM         dbo.TblEmployee TblEmployee_1 LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblEmpJobsTypes ON TblEmployee_1.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.EmpSalaryComponent ON TblEmployee_1.Emp_ID = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.Contract ON TblEmployee_1.Emp_ID = dbo.Contract.Emp_id"
 MySQL = MySQL & "  Where (dbo.Contract.Contract_ID =" & val(Contract_ID) & " )"
 MySQL = MySQL & "   ORDER BY dbo.EmpSalaryComponent.id"




 




     If DataCombo5.ListIndex = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'              StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpContracts.rpt"
'        Else
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpContracts.rpt"
'        End If
        
        
        
        
                If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\EmpContracts.rpt"
        Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\EmpContracts.rpt"
        End If
        
        
        ElseIf DataCombo5.ListIndex = 1 Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpContractsfimaly.rpt"
'        Else
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpContractsfimaly.rpt"
'        End If
'
        
                        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\EmpContractsfimaly.rpt"
        Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\EmpContractsfimaly.rpt"
        End If
        
        
        Else
        Msg = "ÌÃ» «œŒ«· ‰Ê⁄ «·⁄ﬁœ "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DataCombo5.SetFocus
     '   Exit Sub
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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
 
    End If
     If Contract_period_no.Text = "" Then
     Contract_period_no.Text = 0
     End If
     If Contract_period_no.Text = "" Then
     Contract_period.Text = 0
     End If
     xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(IIf(IsNull(RsData("Value").value), 0, RsData("Value").value), "0.00"), 0, True, ".")
     xReport.ParameterFields(5).AddCurrentValue Contract_period_no.Text
     xReport.ParameterFields(6).AddCurrentValue Contract_period.Text
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Public Sub FillGridWithData(EmpID As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblVacationData where EmpID=" & EmpID & "  order by ExpectedacationDate"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.gridHolidayDue
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("ExpectedacationDate")) = IIf(IsNull(rs.Fields("ExpectedacationDate").value), "", rs.Fields("ExpectedacationDate").value)
               
                .TextMatrix(i, .ColIndex("ExpectedacationDateH")) = IIf(IsNull(rs.Fields("ExpectedacationDateH").value), "", rs.Fields("ExpectedacationDateH").value)
            .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(rs.Fields("Remark").value), "", rs.Fields("Remark").value)
             
            .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs.Fields("Value").value), "", rs.Fields("Value").value)
           .TextMatrix(i, .ColIndex("Status1")) = IIf(IsNull(rs.Fields("Status1").value), "", rs.Fields("Status1").value)
           
         
                rs.MoveNext
            Next

            rs.Close
        End If
  .AutoSize 0, .Cols - 1, False
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
   Private Sub Check1_Click()
   Dim STS As Boolean
    If Me.TxtModFlg.Text <> "R" Then
        If val(Me.Contract_period_no.Text) = 0 Or (Me.Contract_period_no.Text) = "" Then
           If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈œŒ«· „œ… «·⁄ﬁœ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Contract_period_no.SetFocus
            STS = True
            Exit Sub
         Else
            MsgBox "Write contract period ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Contract_period_no.SetFocus
            STS = True
            Exit Sub
            End If
        End If
        
       If Me.Contract_period.Text = "" Then
           If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ „œ… «·⁄ﬁœ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Contract_period.SetFocus
            STS = True
            Exit Sub
         Else
            MsgBox "Write Select contract period ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Contract_period.SetFocus
            STS = True
            Exit Sub
            End If
        End If
       If STS = True Then GoTo 10
   
        If Check1.value = vbChecked Then
        cleargriid
        ADDVSFlexGrid1Data
        Else
        cleargriid
        End If
10:
    End If
   End Sub
 Private Sub ADDVSFlexGrid1Data()
   Dim i As Integer
   Dim AddYR As Date
   Dim Addmonth As Date
   With Me.VSFlexGrid1
              For i = .FixedRows To 50
                   .Rows = .FixedRows + i
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("id")) = Me.Contract_ID.Text
                    If i = 1 Then
                    .TextMatrix(i, .ColIndex("startd")) = Me.Contract_date.value
                    .TextMatrix(i, .ColIndex("endd")) = Me.DTPicker1.value
                    Addmonth = Me.Contract_date.value
                    AddYR = Me.DTPicker1.value
                    Else
                    Select Case Contract_period.ListIndex
                    Case 0
                    Addmonth = DateAdd("m", val(Contract_period_no.Text), Addmonth)
                    Addmonth = DateAdd("d", 1, Addmonth)
                    .TextMatrix(i, .ColIndex("startd")) = Addmonth
                    AddYR = DateAdd("m", val(Contract_period_no.Text), AddYR)
                    AddYR = DateAdd("d", 1, AddYR)
                    .TextMatrix(i, .ColIndex("endd")) = AddYR
                    Case 1
                    Addmonth = DateAdd("yyyy", val(Contract_period_no.Text), Addmonth)
                    Addmonth = DateAdd("d", 1, Addmonth)
                   .TextMatrix(i, .ColIndex("startd")) = Addmonth
                    AddYR = DateAdd("yyyy", val(Contract_period_no.Text), AddYR)
                    AddYR = DateAdd("d", 1, AddYR)
                   .TextMatrix(i, .ColIndex("endd")) = AddYR
                    End Select
                    End If
              Next i
        End With
   End Sub
   Private Sub cleargriid()
      Me.VSFlexGrid1.Rows = 1
    End Sub
 Private Sub Cmd_Click(Index As Integer)
  On Error GoTo ErrTrap
    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
               lblnfo.Caption = ""
            GetDefaultComponents
        salary_or_fixed_value(0).value = True
        due_period.ListIndex = 0
        salary_period.ListIndex = 1
        C1Tab1.CurrTab = 0
          Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Grid.Rows = Grid.Rows + 1
            GRID1.Rows = GRID1.Rows + 1
            GRID2.Rows = GRID2.Rows + 1
            getemployeeIformatio Emp_Code

        Case 2
        If CheckRepeatContofEmployee(val(Contract_ID.Text)) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ ⁄„· «ﬂÀ— „‰ ⁄ﬁœ ··„ÊŸ›.ÌÊÃœ ⁄ﬁœ ·Â–« «·„ÊŸ›"
        Else
        MsgBox "The contract can not be repeated to the employee"
        End If
        Exit Sub
        End If
            SaveData

        Case 3
            Undo

        Case 4
            '  If DoPremis(Do_Delete, Me.name, True) = False Then
            '      Exit Sub
            '  End If
            Del_ProfData

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
        
                    If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
Set FrmEmployeeSearch1.RetrunFrm = Me
         FrmEmployeeSearch1.show
         
        Case 6
            Unload Me
 
        Case 9
            DeleteRow

        Case 7
            DeleteRow1

        Case 8
            DeleteRow2

Case 10
        If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.Contract_ID.Text) <> 0 Then
             hide_logo = True
                print_report val(Me.Contract_ID.Text)
        
        
            End If
            
    End Select

    Exit Sub
ErrTrap:

End Sub

Sub DeleteRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub
 
Sub DeleteRow1()

    With Me.GRID1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub

Sub DeleteRow2()

    With Me.GRID2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

End Sub
 
Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
   Private Sub Command1_Click()
    If val(Emp_id.Text) <> 0 Then
        frmEmpSalaryComponent.show
        frmEmpSalaryComponent.Contract_ID = Me.Contract_ID
        frmEmpSalaryComponent.Retrive val(Emp_id.Text)
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ﬂÊœ „ÊŸ› Œÿ√"
        Else
            MsgBox "Invalid Employee Code"
        End If
    End If
End Sub
Private Sub getemployeeIformatio2(Emp_id1 As Double)
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "select * from  TblEmployee where Emp_id=" & Emp_id1
    'sql = "select * from  TblEmployee where Emp_Code='" & emp_codee & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    If rs.RecordCount > 0 Then
        Emp_Code.Text = IIf(IsNull(rs("fullcode").value), "", rs("fullcode"))

        Emp_id.Text = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id"))

        If SystemOptions.UserInterface = ArabicInterface Then
            emp_Name(0).Text = IIf(IsNull(rs("Emp_Name1").value), "", rs("Emp_Name1"))
            emp_Name(1).Text = IIf(IsNull(rs("Emp_Name2").value), "", rs("Emp_Name2"))
            emp_Name(2).Text = IIf(IsNull(rs("Emp_Name3").value), "", rs("Emp_Name3"))
            emp_Name(3).Text = IIf(IsNull(rs("Emp_Name4").value), "", rs("Emp_Name4"))
        Else

            emp_Name(0).Text = IIf(IsNull(rs("Emp_Namee1").value), "", rs("Emp_Namee1"))
            emp_Name(1).Text = IIf(IsNull(rs("Emp_Namee2").value), "", rs("Emp_Namee2"))
            emp_Name(2).Text = IIf(IsNull(rs("Emp_Namee3").value), "", rs("Emp_Namee3"))
            emp_Name(3).Text = IIf(IsNull(rs("Emp_Namee4").value), "", rs("Emp_Namee4"))

        End If

        job.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID"))
        Departement.BoundText = IIf(IsNull(rs("DepartmentID").value), "", rs("DepartmentID"))
        Issue_date.value = IIf(IsNull(rs("BignDateWork").value), Date, rs("BignDateWork"))

        Basic_salary.Text = IIf(IsNull(rs("Emp_Salary").value), 0, rs("Emp_Salary"))
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ﬂÊœ „ÊŸ› Œÿ√"
        Else
            MsgBox "Invalid Employee Code"
        End If

        Emp_Code.Text = ""
    End If

End Sub
 
Private Sub getemployeeIformatio(emp_codee As String)
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "select * from  TblEmployee where fullcode='" & emp_codee & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    If rs.RecordCount > 0 Then
        Emp_id.Text = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id"))

        If SystemOptions.UserInterface = ArabicInterface Then
            emp_Name(0).Text = IIf(IsNull(rs("Emp_Name1").value), "", rs("Emp_Name1"))
            emp_Name(1).Text = IIf(IsNull(rs("Emp_Name2").value), "", rs("Emp_Name2"))
            emp_Name(2).Text = IIf(IsNull(rs("Emp_Name3").value), "", rs("Emp_Name3"))
            emp_Name(3).Text = IIf(IsNull(rs("Emp_Name4").value), "", rs("Emp_Name4"))
        Else

            emp_Name(0).Text = IIf(IsNull(rs("Emp_Namee1").value), "", rs("Emp_Namee1"))
            emp_Name(1).Text = IIf(IsNull(rs("Emp_Namee2").value), "", rs("Emp_Namee2"))
            emp_Name(2).Text = IIf(IsNull(rs("Emp_Namee3").value), "", rs("Emp_Namee3"))
            emp_Name(3).Text = IIf(IsNull(rs("Emp_Namee4").value), "", rs("Emp_Namee4"))

        End If

        job.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID"))
        Departement.BoundText = IIf(IsNull(rs("DepartmentID").value), "", rs("DepartmentID"))
        Issue_date.value = IIf(IsNull(rs("BignDateWork").value), Date, rs("BignDateWork"))

        Basic_salary.Text = IIf(IsNull(rs("Emp_Salary").value), 0, rs("Emp_Salary"))
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ﬂÊœ „ÊŸ› Œÿ√"
        Else
            MsgBox "Invalid Employee Code"
        End If

        Emp_Code.Text = ""
    End If

End Sub

Private Sub Contract_date_Change()
DTPicker1.value = calcenaddate(Contract_date.value, val(Contract_period_no.Text), val(Contract_period.ListIndex))
Txt_DateHigri.value = ToHijriDate(Contract_date.value)
End Sub

Private Sub Contract_period_Click()
DTPicker1.value = calcenaddate(Contract_date.value, val(Contract_period_no.Text), val(Contract_period.ListIndex))

Txt_DateHigri1.value = ToHijriDate(DTPicker1.value)
End Sub

Private Sub Contract_period_no_Change()
DTPicker1.value = calcenaddate(Contract_date.value, val(Contract_period_no.Text), val(Contract_period.ListIndex))
Txt_DateHigri1.value = ToHijriDate(DTPicker1.value)
End Sub
Private Sub Contract_period_no_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Contract_period_no.Text, 0)
End Sub

Private Sub DTPicker1_Change()
Txt_DateHigri1.value = ToHijriDate(DTPicker1.value)
End Sub
Private Sub emp_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        If Me.TxtModFlg.Text <> "R" Then
            getemployeeIformatio Emp_Code
        End If
    End If

End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub
Private Sub Form_Load()
    system_path = App.path
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    Dim Msg As String

    'On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
 
    End If
C1Tab1.CurrTab = 0

    '
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    FullComb
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmpDepartments Me.Departement
    Dcombos.GetEmpJobsTypes Me.job

    Dcombos.GetEmpContractTypes Me.Contract_type
    Dcombos.GetInsuranceClass Me.have_insurance_class
    Dcombos.GetInsuranceClass Me.wife_insurance_class
    Dcombos.GetInsuranceClass Me.Child_insurance_class
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbStatus
    .Clear
    .AddItem "»·«"
    .AddItem " „ «· ⁄ÌÌ‰"
    .AddItem " „ «·«‰Â«¡"
    End With
    Else
     With DcbStatus
    .Clear
    .AddItem "With out"
    .AddItem "Hired"
    .AddItem "No Hired"
    End With
    End If
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    rs.Open "[Contract]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:

End Sub
Function CheckRepeatContofEmployee(Optional Contract_ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     Contract_ID, Emp_id"
sql = sql & " From dbo.Contract"
sql = sql & " Where (Emp_id = " & val(Emp_id.Text) & ") And (Contract_ID <> " & Contract_ID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckRepeatContofEmployee = True
Else
CheckRepeatContofEmployee = False
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    Set EmpReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With Grid

        Select Case .ColKey(Col)

            Case "name"
                'Full Path Display
                 
             '   StrSQL = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
             '   StrSQL = StrSQL & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
             '   StrSQL = StrSQL & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
             '   StrSQL = StrSQL & " Where (dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.text) & ")"
             '   StrSQL = StrSQL & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
 '
     SQLStr = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
                SQLStr = SQLStr & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
                SQLStr = SQLStr & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
                SQLStr = SQLStr & " Where (  (dbo.mofrad.Aloc1 = 1) and dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.Text) & ")"
                SQLStr = SQLStr & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
     
     
                rs.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "name", "mofrad_type")
                Else
                    StrComboList = Grid.BuildComboList(rs, "namee", "mofrad_type")
                End If
                
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
 
            Case "name"
                 
                StrAccountCode = .ComboData
                '    LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("Mofradtype")) = StrAccountCode
 
        End Select
 
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
ErrTrap:
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
 
    With GRID1

        Select Case .ColKey(Col)

            Case "name"
                'Full Path Display
                 
               ' StrSQL = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
               ' StrSQL = StrSQL & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
               ' StrSQL = StrSQL & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
               ' StrSQL = StrSQL & " Where (dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.text) & ")"
               ' StrSQL = StrSQL & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
 
     SQLStr = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
                SQLStr = SQLStr & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
                SQLStr = SQLStr & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
                SQLStr = SQLStr & " Where (  (dbo.mofrad.InCrease = 1) and dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.Text) & ")"
                SQLStr = SQLStr & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
     
     
                rs.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GRID1.BuildComboList(rs, "name", "mofrad_type")
                Else
                    StrComboList = GRID1.BuildComboList(rs, "namee", "mofrad_type")
                End If
                
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With GRID1

        Select Case .ColKey(Col)
 
            Case "name"
                 
                StrAccountCode = .ComboData
                '    LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("Mofradtype")) = StrAccountCode
 
        End Select
 
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLinegrid1
    End With

    ReLineGrid
ErrTrap:
End Sub

Private Sub grid2_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With GRID2

        Select Case .ColKey(Col)

            Case "name"
                'Full Path Display
                 
'                StrSQL = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
'                StrSQL = StrSQL & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
'                StrSQL = StrSQL & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
'                StrSQL = StrSQL & " Where (dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.text) & ")"
'                StrSQL = StrSQL & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
'
    SQLStr = " SELECT     TOP 100 PERCENT dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.emp_ID, dbo.EmpSalaryComponent.mofrad_type"
                SQLStr = SQLStr & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
                SQLStr = SQLStr & " dbo.mofrad ON dbo.EmpSalaryComponent.mofrad_type = dbo.mofrad.id"
                SQLStr = SQLStr & " Where (  (dbo.mofrad.Aloc2= 1) and dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.Text) & ")"
                SQLStr = SQLStr & " ORDER BY dbo.EmpSalaryComponent.emp_ID"
     
     
                rs.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GRID2.BuildComboList(rs, "name", "mofrad_type")
                Else
                    StrComboList = GRID2.BuildComboList(rs, "namee", "mofrad_type")
                End If
                
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub grid2_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With GRID2

        Select Case .ColKey(Col)
 
            Case "name"
                 
                StrAccountCode = .ComboData
                '    LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("Mofradtype")) = StrAccountCode
 
        End Select
 
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLinegrid2
    End With

    ReLineGrid
ErrTrap:
End Sub

Private Sub ISButton1_Click()
C1Tab1.CurrTab = 1
End Sub

Private Sub salary_or_fixed_value_Click(Index As Integer)
    If salary_or_fixed_value(0).value = True Then
        salary_period_no.Enabled = True
        salary_period.Enabled = True
        Fixed_value.Text = ""
    Else
        Fixed_value.Enabled = True
        salary_period_no.Enabled = False
        salary_period.Enabled = False
        salary_period_no.Text = ""
        salary_period.Text = ""
    End If
End Sub
Private Sub salary_period_Click()
If Me.TxtModFlg <> "R" Then
Holiday_period.ListIndex = salary_period.ListIndex
End If
End Sub

Private Sub salary_period_no_Change()
If Me.TxtModFlg <> "R" Then
Holiday_period_no.Text = salary_period_no.Text
End If
End Sub

Private Sub ShowTab_Click()
C1Tab1.CurrTab = 2
End Sub

Private Sub Txt_DateHigri_Validate(Cancel As Boolean)
 On Error GoTo ErrTrap
 VBA.Calendar = vbCalGreg
            Contract_date.value = ToGregorianDate(Txt_DateHigri.value)
                Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "⁄ﬁÊœ «·„ÊŸ›Ì‰"
            Else
                Me.Caption = "Contract Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            '   Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                '      Me.Cmd(7).Enabled = False
            
            End If

            Frame5.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·⁄ﬁÊœ ( ”ÃÌ· ”Ã· ÃœÌœ)"
            Else
                Me.Caption = "Contract  Data(Enter New Record)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '  Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            Frame5.Enabled = True

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ««·⁄ﬁÊœ(  ⁄œÌ· )"
            Else
                Me.Caption = "Contarct Data(Edit Current Record)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '  Me.Cmd(7).Enabled = False
       
            Frame5.Enabled = True

    End Select

    Exit Sub
ErrTrap:
End Sub
 Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If
    Select Case Index
        Case 0
            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious
                cleargriid
                If Check1.value = True Then
                 cleargriid
                 RetriveGrid
                 End If
                If rs.BOF Then rs.MoveFirst
            End If
        Case 1
            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                cleargriid
                 If Check1.value = True Then
                 cleargriid
                 RetriveGrid
                 End If
            End If
        Case 2
            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
                 cleargriid
                 If Check1.value = True Then
                 cleargriid
                 RetriveGrid
                 End If
            End If
        Case 3
            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext
                 cleargriid
                 If Check1.value = True Then
                 cleargriid
                 ADDVSFlexGrid1Data
                 End If
                If rs.EOF Then rs.MoveLast
            End If
    End Select
    Retrive
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, _
                   Optional EmpID As Double = 0, _
                   Optional called As Boolean = False)
    On Error GoTo ErrTrap
    Grid.Rows = 1
    GRID1.Rows = 1
    GRID2.Rows = 1
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
    End If

     If called = False Then
           ' Exit Sub
                   If Lngid <> 0 Then
            rs.find "Contract_ID=" & Lngid, , adSearchForward, adBookmarkFirst

                        If rs.EOF Or rs.BOF Then
                            Exit Sub
                        End If
        End If
           GoTo ViewData
        Else
            GoTo CreateNew
        End If
        
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "Contract_ID=" & Lngid, , adSearchForward, adBookmarkFirst

                        If rs.EOF Or rs.BOF Then
                            Exit Sub
                        End If
        End If
    
CreateNew:
        Dim X As Integer
    
        If EmpID <> 0 Then
                rs.find "Emp_id=" & EmpID, , adSearchForward, adBookmarkFirst
  
                If rs.EOF Or rs.BOF Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox("€Ì— „”Ã· ⁄ﬁœ ·Â–« «·„ÊŸ›  Â·  —Ìœ  ”ÃÌ·  «·⁄ﬁœ «·Œ«’  »Â ‰⁄„/·«", vbInformation + vbYesNo)
                    Else
                        X = MsgBox("No Contract For This Employee Create Contarct y / n", vbInformation + vbYesNo)
                    End If
     
                    If X = vbYes Then
                        Cmd_Click (0)
                        getemployeeIformatio2 EmpID
                        GetDefaultComponents
                    Else
                        Unload Me
                    End If
    
                    '
                    Exit Sub
                
                End If
            Else
               Exit Sub
            
            End If
        
              
    End If
ViewData:
    Contract_ID.Text = IIf(IsNull(rs("Contract_ID").value), "", val(rs("Contract_ID").value))
    ' aladein add
    Me.Contract_date.value = IIf(IsNull(rs("Contract_date").value), Date, rs("Contract_date").value)
    Me.DataCombo5.ListIndex = IIf(IsNull(rs("ADDtype_Contract").value), -1, rs("ADDtype_Contract").value)
    'DTPicker1.value = IIf(IsNull(rs("Contract_EndDate").value), Date, Format(rs("Contract_EndDate").value, "DD/MM/YYYY"))
'        DTPicker1.value = Format(rs("Contract_EndDate").value, "DD\MM\YYYY")
        
    Me.Txt_DateHigri.value = IIf(IsNull(rs("DateH").value), "", rs("DateH").value)
    Me.Txt_DateHigri1.value = IIf(IsNull(rs("DateH1").value), "", rs("DateH").value)
    Me.Contract_type.BoundText = IIf(IsNull(rs("Contract_type").value), "", rs("Contract_type").value)
    Contract_period_no.Text = IIf(IsNull(rs("Contract_period_no").value), 0, rs("Contract_period_no").value)
    DcbStatus.ListIndex = IIf(IsNull(rs("StutsID").value), -1, rs("StutsID").value)
    If IsNull(rs("Contract_period").value) Then
        Me.Contract_period.ListIndex = 0
    Else
        Me.Contract_period.ListIndex = rs("Contract_period").value
    End If

    test_period_no.Text = IIf(IsNull(rs("test_period_no").value), 0, rs("test_period_no").value)

    If IsNull(rs("test_period").value) Then
        Me.test_period.ListIndex = 0
    Else
        Me.test_period.ListIndex = rs("test_period").value
    End If
    Me.job.BoundText = IIf(IsNull(rs("job").value), "", rs("job").value)

   
    Me.Departement.BoundText = IIf(IsNull(rs("Departement").value), "", rs("Departement").value)
    Issue_date.value = IIf(IsNull(rs("Issue_date").value), Date, rs("Issue_date").value)
    Basic_salary.Text = IIf(IsNull(rs("Basic_salary").value), "", rs("Basic_salary").value)
    Emp_id.Text = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
       getemployeeIformatio2 Emp_id
       FillGridWithData Emp_id
       
   ' emp_code.text = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
    XPTxtEmpName.Text = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
  '  emp_name(0).text = IIf(IsNull(rs("Emp_Name1").value), "", Trim(rs("Emp_Name1").value))
  '  emp_name(1).text = IIf(IsNull(rs("Emp_Name2").value), "", Trim(rs("Emp_Name2").value))
  '  emp_name(2).text = IIf(IsNull(rs("Emp_Name3").value), "", Trim(rs("Emp_Name3").value))
  '  emp_name(3).text = IIf(IsNull(rs("Emp_Name4").value), "", Trim(rs("Emp_Name4").value))

    Due_period_no.Text = IIf(IsNull(rs("Due_period_no").value), 0, rs("Due_period_no").value)

    If IsNull(rs("due_period").value) Then
        Me.due_period.ListIndex = 0
    Else
        Me.due_period.ListIndex = rs("due_period").value
    End If

    Holiday_date.value = IIf(IsNull(rs("Holiday_date").value), Date, rs("Holiday_date").value)

    Holiday_period_no.Text = IIf(IsNull(rs("Holiday_period_no").value), 0, rs("Holiday_period_no").value)

    If IsNull(rs("Holiday_period").value) Then
        Me.Holiday_period.ListIndex = 0
    Else
        Me.Holiday_period.ListIndex = rs("Holiday_period").value
    End If

    If IsNull(rs("salary_or_fixed_value").value) Or rs("salary_or_fixed_value").value = 0 Then
        salary_or_fixed_value(0).value = True
    Else
        salary_or_fixed_value(1).value = True
    End If

    salary_period_no.Text = IIf(IsNull(rs("salary_period_no").value), 0, rs("salary_period_no").value)

    If IsNull(rs("salary_period").value) Then
        Me.salary_period.ListIndex = 0
    Else
        Me.salary_period.ListIndex = rs("salary_period").value
    End If

    If rs("have_ticket").value = True Then
        have_ticket.value = Checked
    Else
        have_ticket.value = Unchecked
    End If

    If rs("wife_ticket").value = True Then
        wife_ticket.value = Checked
    Else
        wife_ticket.value = Unchecked
    End If

    If rs("Child_ticket").value = True Then
        Child_ticket.value = Checked
    Else
        Child_ticket.value = Unchecked
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''
     If rs("AutRenContract").value = True Then
        Check1.value = Checked
    Else
        Check1.value = Unchecked
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
    no_of_Child_ticket.Text = IIf(IsNull(rs("no_of_Child_ticket").value), 0, rs("no_of_Child_ticket").value)
    txtTicketValue.Text = IIf(IsNull(rs("TicketValue").value), 0, rs("TicketValue").value)
    TxtInsuranceNo.Text = IIf(IsNull(rs("InsuranceNO").value), "", rs("InsuranceNO").value)

    If IsNull(rs("yearly_increase_fixed_value_or_percentage").value) Or rs("yearly_increase_fixed_value_or_percentage").value = 0 Then
        yearly_increase_fixed_value_or_percentage(0).value = True
        yearly_increase(0).Text = IIf(IsNull(rs("yearly_increase").value), 0, rs("yearly_increase").value)
    Else
    
        yearly_increase_fixed_value_or_percentage(1).value = True
        yearly_increase(1).Text = IIf(IsNull(rs("yearly_increase").value), 0, rs("yearly_increase").value)
    End If

    If rs("have_insurance").value = True Then
        have_insurance.value = Checked
    Else
        have_insurance.value = Unchecked

    End If

    Me.have_insurance_class.BoundText = IIf(IsNull(rs("have_insurance_class").value), "", rs("have_insurance_class").value)
    
    If rs("wife_insurance").value = True Then
        wife_insurance.value = Checked
    Else
        wife_insurance.value = Unchecked

    End If

    Me.wife_insurance_class.BoundText = IIf(IsNull(rs("wife_insurance_class").value), "", rs("wife_insurance_class").value)
        
    If rs("Child_insurance").value = True Then
        Child_insurance.value = Checked
    Else
        Child_insurance.value = Unchecked

    End If

    Me.Child_insurance_class.BoundText = IIf(IsNull(rs("Child_insurance_class").value), "", rs("Child_insurance_class").value)
    no_of_Child_insurance.Text = IIf(IsNull(rs("no_of_Child_insurance").value), 0, rs("no_of_Child_insurance").value)
    
    'ret details
    Dim rscomponent As ADODB.Recordset
    Dim sql As String

    'sql = " select * from EmpSalaryComponent where emp_ID=" & Val(Emp_id.text)
    sql = " SELECT     dbo.mofrad.name, dbo.mofrad.nameE, dbo.TblContractDetails.Mofradtype"
    sql = sql & " FROM         dbo.TblContractDetails INNER JOIN"
    sql = sql & " dbo.mofrad ON dbo.TblContractDetails.Mofradtype = dbo.mofrad.id"
    sql = sql & " WHERE     (dbo.TblContractDetails.DefDataType = 0) AND (dbo.TblContractDetails.Contract_ID = " & val(Contract_ID.Text) & ")"

    Set rscomponent = New ADODB.Recordset
    rscomponent.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.Grid
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Mofradtype")) = IIf(IsNull(rscomponent("Mofradtype").value), "", rscomponent("Mofradtype").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("name").value), "", rscomponent("name").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("namee").value), "", rscomponent("namee").value)
                End If
            
                rscomponent.MoveNext
            Next
   
        End With

    End If

    Set rscomponent = Nothing
    
    sql = " SELECT     dbo.mofrad.name, dbo.mofrad.nameE, dbo.TblContractDetails.Mofradtype"
    sql = sql & " FROM         dbo.TblContractDetails INNER JOIN"
    sql = sql & " dbo.mofrad ON dbo.TblContractDetails.Mofradtype = dbo.mofrad.id"
    sql = sql & " WHERE     (dbo.TblContractDetails.DefDataType = 1) AND (dbo.TblContractDetails.Contract_ID = " & val(Contract_ID.Text) & ")"

    Set rscomponent = New ADODB.Recordset
    rscomponent.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.GRID1
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Mofradtype")) = IIf(IsNull(rscomponent("Mofradtype").value), "", rscomponent("Mofradtype").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("name").value), "", rscomponent("name").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("namee").value), "", rscomponent("namee").value)
                End If
            
                rscomponent.MoveNext
            Next
   
        End With

    End If

    Set rscomponent = Nothing

    sql = " SELECT     dbo.mofrad.name, dbo.mofrad.nameE, dbo.TblContractDetails.Mofradtype"
    sql = sql & " FROM         dbo.TblContractDetails INNER JOIN"
    sql = sql & " dbo.mofrad ON dbo.TblContractDetails.Mofradtype = dbo.mofrad.id"
    sql = sql & " WHERE     (dbo.TblContractDetails.DefDataType = 2) AND (dbo.TblContractDetails.Contract_ID = " & val(Contract_ID.Text) & ")"

    Set rscomponent = New ADODB.Recordset
    rscomponent.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.GRID2
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Mofradtype")) = IIf(IsNull(rscomponent("Mofradtype").value), "", rscomponent("Mofradtype").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("name").value), "", rscomponent("name").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rscomponent("namee").value), "", rscomponent("namee").value)
                End If
            
                rscomponent.MoveNext
            Next
           End With
          RetriveGrid
    End If
    lblnfo.Caption = Emp_Code & "    " & emp_Name(0).Text & " " & emp_Name(1).Text & " " & emp_Name(2).Text & " " & emp_Name(3).Text & " "
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub RetriveGrid()
      Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
  sql = "SELECT     ID, Contract_ID, StartDate, EndDate, status"
  sql = sql + "       From dbo.TBLContractAuto"
  sql = sql + "  Where (Contract_ID = " & val(Contract_ID.Text) & ") "
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.VSFlexGrid1
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("Contract_ID").value), "", Rs1("Contract_ID").value)
                   .TextMatrix(i, .ColIndex("startd")) = IIf(IsNull(Rs1("StartDate").value), "", Rs1("StartDate").value)
                   .TextMatrix(i, .ColIndex("endd")) = IIf(IsNull(Rs1("EndDate").value), "", Rs1("EndDate").value)
                   .TextMatrix(i, .ColIndex("stats")) = IIf(IsNull(Rs1("status").value), "", Rs1("status").value)
                    Rs1.MoveNext
             Next i
        End With
        Exit Sub
End Sub

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
  ' On Error GoTo ErrTrap
    XPTxtEmpName = Trim(emp_Name(0).Text) & " " & Trim(emp_Name(1).Text) & " " & Trim(emp_Name(2).Text) & " " & Trim(emp_Name(3).Text)
     If Me.TxtModFlg.Text <> "R" Then
        If emp_Name(0).Text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„ «·„ÊŸ› "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            emp_Name(0).SetFocus
            SelectText emp_Name(0)
            Exit Sub
        End If
  '   If SystemOptions.UserInterface = EnglishInterface Then
        '   If DataCombo5.text = "" Then
       '    Msg = "Select Contract Type "
      ''      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
          '  DataCombo5.SetFocus
       '     SelectText DataCombo5
       '     Exit Sub
        '   Else
        If DataCombo5.Text = "" Then
            Msg = "ÌÃ» «œŒ«· ‰Ê⁄ «·⁄ﬁœ "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DataCombo5.SetFocus
            SelectText DataCombo5
            Exit Sub
       End If
    '   End If
    
        '        If emp_name(2).text = "" Then
        '        Msg = "ÌÃ» «œŒ«· «”„ «·Ãœ "
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        emp_name(2).SetFocus
        '        SelectText emp_name(2)
        '        Exit Sub
        '       End If
    
        '      If emp_name(3).text = "" Then
        '        Msg = "ÌÃ» «œŒ«· «”„ «·⁄«∆·… "
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        emp_name(3).SetFocus
        '        SelectText emp_name(3)
        '        Exit Sub
        '       End If
    
        If Not IsNumeric(Basic_salary.Text) Then
            Msg = "ÌÃ» «œŒ«· «·—« » «·«”«”Ì ··„ÊŸ›  "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Basic_salary.SetFocus
            SelectText Basic_salary
            Exit Sub
        End If
    
        If val(Due_period_no.Text) = 0 Then
            Msg = "«·„” Õﬁ«  «·„«œÌ… ⁄‰ ﬂ·  › —… ·« Ì„ﬂ‰« «‰  ﬂÊ‰ ’›—"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
            Due_period_no.SetFocus
            SendKeys "{F4}"
        
            Exit Sub
        End If
    
  '      If job.BoundText = "" Then
    
  '          Msg = "»—Ã«¡  ÕœÌœ «·ÊŸÌ›…"
  '          MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  '
  '          job.SetFocus
  '          SendKeys "{F4}"
  '
  '          Exit Sub
  '      End If
    
'        If Departement.BoundText = "" Then
    
'            Msg = "»—Ã«¡  ÕœÌœ «·ﬁ”„ «·–Ì Ì »⁄Â «·„ÊŸ›"
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'
'            Departement.SetFocus
'            SendKeys "{F4}"
'
'            Exit Sub
'        End If
    
        Select Case TxtModFlg.Text

            Case "N"
                StrVacCode = IsRecExist("contract", "Emp_id", Trim(Emp_id.Text), " Emp_Name")

                 If StrVacCode <> "" Then
                    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« ⁄ﬁœ  ·Â–« «·„ÊŸ›  "
                    MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
                    Emp_Code.SetFocus
                    SelectText Emp_Code
                    Exit Sub
                End If

                StrSQL = "select * From contract where Emp_Name='" & Trim(XPTxtEmpName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    Msg = "ÌÊÃœ „ÊŸ› „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

            Case "E"
                '     StrSQL = "select * From contract where Emp_Name='" & Trim(XPTxtEmpName.text) & "'"
                '     RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '     If RsTemp.RecordCount > 0 Then
                '         If RsTemp("Emp_ID").value <> Val(XPTxtEmpID) Then
                '             Msg = "ÌÊÃœ „ÊŸ› „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                '             Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
                '             Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                '             MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '             Exit Sub
                '         End If
                '     End If
        End Select

        Cn.BeginTrans
        BeginTrans = True
    
        If TxtModFlg.Text = "N" Then
      Contract_ID.Text = CStr(new_id("contract", "Contract_ID", "", True))

            rs.AddNew
        Else
     
            sql = "delete    dbo.TblContractDetails  where Contract_ID=" & val(Me.Contract_ID.Text)
            Cn.Execute sql
        End If
       
        rs("Contract_ID").value = val(Contract_ID.Text)
        rs("Contract_date").value = Contract_date.value
        rs("DateH").value = Me.Txt_DateHigri.value
        rs("DateH1").value = Me.Txt_DateHigri1.value
        rs("Contract_Enddate").value = DTPicker1.value
        rs("StutsID").value = val(DcbStatus.ListIndex)
        ' aladein add
         If val(Me.DataCombo5.ListIndex) = -1 Then
            rs("ADDtype_Contract").value = Null
        Else
            rs("ADDtype_Contract").value = val(Me.DataCombo5.ListIndex)
        End If
        '''''''''
     
        If val(Me.Contract_type.BoundText) = 0 Then
            rs("Contract_type").value = Null
        Else
            rs("Contract_type").value = val(Me.Contract_type.BoundText)
        End If

        rs("Contract_period_no").value = val(Contract_period_no.Text)
        rs("Contract_period").value = Contract_period.ListIndex
        rs("test_period_no").value = val(test_period_no.Text)
        rs("test_period").value = test_period.ListIndex
        rs("Emp_ID").value = val(Emp_id.Text)
        rs("Emp_Code").value = IIf(Emp_Code.Text = "", Null, Trim(Emp_Code.Text))
        rs("Emp_Name1").value = Trim(emp_Name(0).Text)
        rs("Emp_Name2").value = Trim(emp_Name(1).Text)
        rs("Emp_Name3").value = Trim(emp_Name(2).Text)
        rs("Emp_Name4").value = Trim(emp_Name(3).Text)
        rs("Emp_Name").value = Trim(emp_Name(0).Text) & " " & Trim(emp_Name(1).Text) & " " & Trim(emp_Name(2).Text) & " " & Trim(emp_Name(3).Text)
        rs("Basic_salary").value = val(Basic_salary.Text)
     
        If val(Me.job.BoundText) = 0 Then
            rs("job").value = Null
        Else
            rs("job").value = val(Me.job.BoundText)
        End If
    
        If val(Me.Departement.BoundText) = 0 Then
            rs("Departement").value = Null
        Else
            rs("Departement").value = val(Me.Departement.BoundText)
        End If
    
        rs("Issue_date").value = Issue_date.value
    
        rs("Due_period_no").value = val(Due_period_no.Text)
        rs("due_period").value = due_period.ListIndex
    
        rs("Holiday_date").value = Holiday_date.value
        
        rs("Holiday_period_no").value = val(Holiday_period_no.Text)
        rs("Holiday_period").value = Holiday_period.ListIndex
    
        If salary_or_fixed_value(0).value = True Then
            rs("salary_or_fixed_value").value = 0
        Else
            rs("salary_or_fixed_value").value = 1
        End If
    
        rs("salary_period_no").value = val(salary_period_no.Text)
        rs("salary_period").value = salary_period.ListIndex
        rs("Fixed_value").value = val(Fixed_value.Text)
      
        If have_ticket.value = Checked Then
            rs("have_ticket").value = 1
        Else
            rs("have_ticket").value = 0
        End If
    
        If wife_ticket.value = Checked Then
            rs("wife_ticket").value = 1
        Else
            rs("wife_ticket").value = 0
        End If
    
        If Child_ticket.value = Checked Then
            rs("Child_ticket").value = 1
        Else
            rs("Child_ticket").value = 0
        End If
       '''''''''''''
       If Check1.value = Checked Then
            rs("AutRenContract").value = 1
        Else
            rs("AutRenContract").value = 0
        End If
       '''''''''''''''''''''''''''''
        rs("no_of_Child_ticket").value = val(no_of_Child_ticket.Text)
        rs("TicketValue").value = val(txtTicketValue.Text)
        rs("InsuranceNO").value = val(TxtInsuranceNo.Text)
   
        Dim noofticket As Integer
        noofticket = noofticket

        If have_ticket.value = vbChecked Then
            noofticket = noofticket + 1
        End If
   
        If wife_ticket.value = vbChecked Then
            noofticket = noofticket + 1
        End If
   
        If Child_ticket.value = vbChecked Then
            noofticket = noofticket + val(no_of_Child_ticket)
        End If
   
        rs("TicketValueTotal").value = val(txtTicketValue.Text) * noofticket
   
        If yearly_increase_fixed_value_or_percentage(0).value = True Then
            rs("yearly_increase_fixed_value_or_percentage").value = 0
            rs("yearly_increase").value = val(yearly_increase(0).Text)
        Else
            rs("yearly_increase_fixed_value_or_percentage").value = 1
            rs("yearly_increase").value = val(yearly_increase(1).Text)
        End If
    
        If have_insurance.value = Checked Then
            rs("have_insurance").value = 1
        Else
            rs("have_insurance").value = 0
        End If
    
        If val(Me.have_insurance_class.BoundText) = 0 Then
            rs("have_insurance_class").value = Null
        Else
            rs("have_insurance_class").value = val(Me.have_insurance_class.BoundText)
        End If
  
        If wife_insurance.value = Checked Then
            rs("wife_insurance").value = 1
        Else
            rs("wife_insurance").value = 0
        End If
    
        If val(Me.wife_insurance_class.BoundText) = 0 Then
            rs("wife_insurance_class").value = Null
        Else
            rs("wife_insurance_class").value = val(Me.wife_insurance_class.BoundText)
        End If
  
        If Child_insurance.value = Checked Then
            rs("Child_insurance").value = 1
        Else
            rs("Child_insurance").value = 0
        End If
    
        If val(Me.Child_insurance_class.BoundText) = 0 Then
            rs("Child_insurance_class").value = Null
        Else
            rs("Child_insurance_class").value = val(Me.Child_insurance_class.BoundText)
        End If
    
        rs("no_of_Child_insurance").value = val(no_of_Child_insurance.Text)
          
        rs.update
    
        'save SubDetails
        Dim rscomponent As ADODB.Recordset
     '   sql = "TblContractDetails"
        Set rscomponent = New ADODB.Recordset
       ' rscomponent.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdTable
        sql = "SELECT     * from dbo.TblContractDetails Where (1 = -1)"
        rscomponent.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        With Me.Grid
            For i = .FixedRows To .Rows - 1
                If val(.TextMatrix(i, .ColIndex("Mofradtype"))) <> 0 Then
                    rscomponent.AddNew
                    rscomponent("Contract_ID").value = val(Contract_ID.Text)
                    rscomponent("emp_ID").value = val(Emp_id.Text) '„—»Êÿ »—ﬁ„ «·„ÊŸ›
                    rscomponent("Mofradtype").value = IIf(.TextMatrix(i, .ColIndex("Mofradtype")) = "", 0, .TextMatrix(i, .ColIndex("Mofradtype")))
                    rscomponent("DefDataType").value = 0
                    rscomponent.update
                End If
              Next i
        End With
        
        With Me.GRID1
            For i = .FixedRows To .Rows - 1
                If val(.TextMatrix(i, .ColIndex("Mofradtype"))) <> 0 Then
                    rscomponent.AddNew
                    rscomponent("Contract_ID").value = val(Contract_ID.Text)
                    rscomponent("emp_ID").value = val(Emp_id.Text) '„—»Êÿ »—ﬁ„ «·„ÊŸ›
                    rscomponent("Mofradtype").value = IIf(.TextMatrix(i, .ColIndex("Mofradtype")) = "", 0, .TextMatrix(i, .ColIndex("Mofradtype")))
                    rscomponent("DefDataType").value = 1
                    rscomponent.update
                End If
  
            Next i
        End With
        With Me.GRID2
            For i = .FixedRows To .Rows - 1
                If val(.TextMatrix(i, .ColIndex("Mofradtype"))) <> 0 Then
                    rscomponent.AddNew
                    rscomponent("Contract_ID").value = val(Contract_ID.Text)
                    rscomponent("emp_ID").value = val(Emp_id.Text) '„—»Êÿ »—ﬁ„ «·„ÊŸ›
                    rscomponent("Mofradtype").value = IIf(.TextMatrix(i, .ColIndex("Mofradtype")) = "", 0, .TextMatrix(i, .ColIndex("Mofradtype")))
                    rscomponent("DefDataType").value = 2
                    rscomponent.update
                End If
              Next i
        End With
        Cn.CommitTrans
        FullGridRece
       
        RetriveGrid
        BeginTrans = False
         
         
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
            Select Case Me.TxtModFlg.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ÊŸ› " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
 
                    Msg = "saved Success " & CHR(13)
                    Msg = Msg + "Do you want another entry?"
       
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
        CreateVacationData (val(Emp_id.Text))
        FillGridWithData val(Emp_id.Text)
        
        TxtModFlg.Text = "R"
     End If
    Exit Sub
ErrTrap:
    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
 Private Sub FullGridRece()
  '  On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    If Check1.value = vbChecked Then
    StrSQL = "Delete From TBLContractAuto Where Contract_ID='" & val(Contract_ID.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    End If
    Dim RsDevsub As ADODB.Recordset
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLContractAuto Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With VSFlexGrid1
       For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("Contract_ID").value = Me.Contract_ID.Text
                RsDevsub("StartDate").value = IIf((.TextMatrix(i, .ColIndex("startd"))) = "", Null, .TextMatrix(i, .ColIndex("startd")))
                RsDevsub("EndDate").value = IIf((.TextMatrix(i, .ColIndex("endd"))) = "", Null, .TextMatrix(i, .ColIndex("endd")))
                If val(.TextMatrix(i, .ColIndex("stats"))) = vbChecked Then
                RsDevsub("status").value = 1
                Else
                RsDevsub("status").value = 0
                End If
                RsDevsub.update
      End If
      Next i
    End With
ErrTrap:
    End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "Contract_ID='" & val(Contract_ID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_ProfData()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Contract_ID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„ÊŸ› —ﬁ„ " & CHR(13)
        Msg = Msg + (Contract_ID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst
            
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·⁄ﬁœ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            KeyCode = 0
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If Shift = VBRUN.ShiftConstants.vbShiftMask Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄ﬁœ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄ﬁÊœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄ﬁÊœ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› «·⁄ﬁÊœ „ÊŸ›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄ﬁœ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New Record ..." & Wrap & "Click here to add a new Contract" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the current Contract data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new record or " & Wrap & "save the edit in the " & Wrap & "current record", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo" & Wrap & "Undo in the adding new record" & Wrap & "Or undo in the current editing" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current Contract data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an Contract" & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
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

Private Sub ChangeLang()
Label31.Caption = "Status"
    ISButton2.Caption = "Select All"
    ISButton5.Caption = "Undo Select"
    Me.Caption = "Employee Contract"
    EleHeader.Caption = Me.Caption
    Label29.Caption = "End Date"
    Label27.Caption = "Value"
    Cmd(10).Caption = "Print"
    Label30.Caption = "Type of Contract"
    Label3.Caption = "Contract #"
    Label4.Caption = "Date"
    Label8.Caption = "Type"
    Label9.Caption = "Period"
    Label10.Caption = "Exam period"
    Label28.Caption = "Insurance Number"
    Label6.Caption = "Emp Code"
    Label7.Caption = "Emp Name"
    Label11.Caption = "Job"
    Label12.Caption = "Departement"
    Label13.Caption = "Start date"
    lbl(1).Caption = "Dues for each"
    C1Tab1.CurrTab = 0
    C1Tab1.TabCaption(0) = "Basic data"
    C1Tab1.TabCaption(1) = "Due Holidya Data"
    C1Tab1.TabCaption(2) = "Automatic Renewal Contract"
    Check1.Caption = "Automatic Renewal Contract"
    Label17.Caption = "The due date of leave"
    lbl(0).Caption = "vacation days"
    Label18.Caption = "Basic Salary"
    ISButton3.Caption = "Remove row"
    ISButton4.Caption = "Remove All"

    yearly_increase_fixed_value_or_percentage(0).Caption = "Salary"
    yearly_increase_fixed_value_or_percentage(1).Caption = "Fixed Value"

    Label5.Caption = "The salary earned by the employee"

    With due_period
        .Clear
 
        .AddItem "Month"
        .AddItem "Year"
        .AddItem "Day"
    End With

    With salary_period
        .Clear
        .AddItem "Day"
        .AddItem "Month"

    End With
    
    With test_period
        .Clear
        .AddItem "Month"
        .AddItem "Year"

    End With
With Contract_period
        .Clear
        .AddItem "Month"
        .AddItem "Year"

    End With
    

    With Holiday_period
        .Clear
        .AddItem "Day"
        .AddItem "month"

    End With

    'With Me.VSFlexGrid2
    '.TextMatrix(0, .ColIndex("LineNo")) = "LineNo"
    '.TextMatrix(0, .ColIndex("AccountName")) = "Component Name "
    '.TextMatrix(0, .ColIndex("value")) = "value"
    '.TextMatrix(0, .ColIndex("des")) = "des"

    'End With

    'With Me.xx
    '.TextMatrix(0, .ColIndex("LineNo")) = "LineNo"
    '.TextMatrix(0, .ColIndex("AccountName")) = "Component Name "
    '.TextMatrix(0, .ColIndex("value")) = "value"
    '.TextMatrix(0, .ColIndex("des")) = "des"

    'End With
 
    Frame1.Caption = "Salary"

    Label2.Caption = "Leave entitlements"
    have_ticket.Caption = "His ticket"
    wife_ticket.Caption = "Ticket for Wife"
    Child_ticket.Caption = "Ticket for child ,count"

    Frame2.Caption = "Annual increase"

    salary_or_fixed_value(0).Caption = "Fixed Value"
    salary_or_fixed_value(1).Caption = "Percentage"
    Frame3.Caption = "Percentage"

    Label14.Caption = "Select Salary"
    Frame4.Caption = "Medical Insurance"
    Label15.Caption = "Class"
    have_insurance.Caption = "His"
    wife_insurance.Caption = "Wife"
    Child_insurance.Caption = "Child"
    Label19.Caption = "Current rec"
    Label20.Caption = "Total rec"

    Label16.Caption = "Count"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Label2.Caption = "Vacation Components"

    With Grid
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("name")) = "name"
    End With

    Cmd(9).Caption = "Remove "

    Label14.Caption = "Annual increase Components"

    With GRID1
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("name")) = "name"
    End With
    Cmd(7).Caption = "Remove "
    Frame6.Caption = "End Of Service Components"
    With GRID2
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("name")) = "name"
    End With
    Cmd(8).Caption = "Remove "
    lbl(24).Caption = "Notes"
    lbl(25).Caption = "Select Components"
    lbl(6).Caption = "Notes"
    lbl(5).Caption = "Select Components"
    Frame3.Caption = "percentage"
    Label25.Caption = "Class"
    Label26.Caption = "Count"
    Label15.Caption = "Class"
    Label23.Caption = "Class"
    Frame8.Caption = "Tickets"
    Frame7.Caption = "Percentage"
    Label21.Caption = "End Of Service Components"
       
    With gridHolidayDue
        .TextMatrix(0, .ColIndex("ExpectedacationDate")) = "Expec.Date"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("ExpectedacationDateH")) = "Expec. DateH"
        .TextMatrix(0, .ColIndex("Status1")) = "Done"
        .TextMatrix(0, .ColIndex("Remark")) = "Remark"
    End With
    
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Chkk")) = "select"
        .TextMatrix(0, .ColIndex("startd")) = "Start Date"
        .TextMatrix(0, .ColIndex("endd")) = "End Date"
        .TextMatrix(0, .ColIndex("stats")) = "Status"
    End With
    
End Sub
 
Private Sub xx_Click()

End Sub

Private Sub yearly_increase_fixed_value_or_percentage_Click(Index As Integer)

    If yearly_increase_fixed_value_or_percentage(0).value = True Then
        yearly_increase(0).Enabled = True
        yearly_increase(1).Enabled = False
 
        yearly_increase(1).Text = ""

    Else
        yearly_increase(1).Enabled = True
        yearly_increase(0).Enabled = False
        yearly_increase(0).Text = ""
    End If

End Sub
'''''''''''''''''''''''''''''''''''''''
 Private Sub FullComb()
    If SystemOptions.UserInterface = EnglishInterface Then
    With Me.DataCombo5
        .Clear
        .AddItem "Single"
        .AddItem "Family"
        End With
    Else
    With Me.DataCombo5
        .Clear
        .AddItem "›—œÌ"
        .AddItem "⁄«∆·Ì"
        End With
    End If
 End Sub
 Private Sub ISButton2_Click()
   On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim DelRow As Integer
    With Me.VSFlexGrid1
                  Selrow = True
                  For DelRow = .FixedRows To .Rows - 1
                  If val(.TextMatrix(DelRow, .ColIndex("Chkk"))) = vbUnchecked Then
                 .TextMatrix(DelRow, .ColIndex("Chkk")) = Selrow
                  Else
                  End If
           Next DelRow
    End With
 Exit Sub
ErrTrap:
End Sub
Private Sub ISButton5_Click()
  On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim DelRow As Integer
    With Me.VSFlexGrid1
                  Selrow = False
                  For DelRow = .FixedRows To .Rows - 1
                  If val(.TextMatrix(DelRow, .ColIndex("Chkk"))) = True Then
                 .TextMatrix(DelRow, .ColIndex("Chkk")) = Selrow
                  Else
                  End If
           Next DelRow
    End With
 Exit Sub
ErrTrap:
End Sub
Private Sub ISButton3_Click()
    On Error GoTo ErrTrap
    Dim i As Integer
    With Me.VSFlexGrid1
                  For i = .FixedRows To .Rows - 1
                  If val(.TextMatrix(i, .ColIndex("Chkk"))) = True Then
                  VSFlexGrid1.RemoveItem i
                  Else
                  GoTo 10
                  End If
                  i = i - 1
10:        Next i
    End With
 Exit Sub
ErrTrap:
End Sub
Private Sub ISButton4_Click()
      cleargriid
End Sub
''''''''''''''''''''''''''' end


