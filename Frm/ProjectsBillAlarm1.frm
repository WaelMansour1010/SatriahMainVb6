VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form ProjectsBillAlarm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17310
   Icon            =   "ProjectsBillAlarm1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10860
   ScaleWidth      =   17310
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
   Begin VB.Timer tmrScrolling 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   10920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   10920
   End
   Begin C1SizerLibCtl.C1Elastic frm_Dash 
      Height          =   10875
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   17430
      _cx             =   30745
      _cy             =   19182
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   -45
         RightToLeft     =   -1  'True
         TabIndex        =   204
         Top             =   -45
         Width           =   17520
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ƒ‘—«  «·ÕÌ…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   120
            Width           =   4920
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   10440
         Left            =   -165
         TabIndex        =   1
         Top             =   360
         Width           =   17400
         _cx             =   30692
         _cy             =   18415
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
         Caption         =   "«· Õ·Ì· «·„«·Ï|«·«‰ «Ã|«·ÕÃ“|«·«Ã„«·Ì« |«⁄œ«œ« |«·„»Ì⁄« |«·„ÊŸ›Ì‰+«·„Œ«“‰+«·⁄„·«¡+«·„Ê—œÌ‰|„” Œ·’«  «·„‘«—Ì⁄|«·„»Ì⁄«  Ê «·„’—Ê›« "
         Align           =   0
         CurrTab         =   8
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic20 
            Height          =   10020
            Left            =   -17955
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic21 
               Height          =   10080
               Left            =   0
               TabIndex        =   195
               TabStop         =   0   'False
               Top             =   0
               Width           =   17310
               _cx             =   30533
               _cy             =   17780
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic22 
                  Height          =   1185
                  Left            =   -510
                  TabIndex        =   196
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   16470
                  _cx             =   29051
                  _cy             =   2090
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
                  Begin MSComCtl2.DTPicker FrmDate4 
                     Height          =   390
                     Left            =   13260
                     TabIndex        =   197
                     Top             =   555
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   96337921
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker ToDate4 
                     Height          =   390
                     Left            =   10320
                     TabIndex        =   198
                     Top             =   555
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   96337921
                     CurrentDate     =   41640
                  End
                  Begin ALLButtonS.ALLButton ALLButton5 
                     Height          =   390
                     Left            =   2040
                     TabIndex        =   199
                     Top             =   555
                     Width           =   7230
                     _ExtentX        =   12753
                     _ExtentY        =   688
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
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
                     BCOL            =   16711680
                     BCOLO           =   16711680
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "ProjectsBillAlarm1.frx":058A
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„” Œ·’«  «·„‘«—Ì⁄"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   14.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   375
                     Index           =   6
                     Left            =   5040
                     RightToLeft     =   -1  'True
                     TabIndex        =   202
                     Top             =   120
                     Width           =   5160
                  End
                  Begin VB.Label Label52 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰  «—ÌŒ"
                     Height          =   315
                     Left            =   15300
                     RightToLeft     =   -1  'True
                     TabIndex        =   201
                     Top             =   570
                     Width           =   690
                  End
                  Begin VB.Label Label51 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï  «—ÌŒ"
                     Height          =   270
                     Left            =   12180
                     RightToLeft     =   -1  'True
                     TabIndex        =   200
                     Top             =   570
                     Width           =   810
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid6 
                  Height          =   8790
                  Left            =   0
                  TabIndex        =   203
                  Top             =   1170
                  Width           =   15975
                  _cx             =   28178
                  _cy             =   15505
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
                  Rows            =   2
                  Cols            =   14
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"ProjectsBillAlarm1.frx":05A6
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic11 
            Height          =   10020
            Left            =   -18555
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin VB.TextBox TxtData 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   435
               Index           =   5
               Left            =   6885
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   7050
               Width           =   4035
            End
            Begin VB.TextBox TxtData 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   435
               Index           =   4
               Left            =   6885
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   6495
               Width           =   4035
            End
            Begin VB.TextBox TxtData 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Index           =   3
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   7050
               Width           =   2370
            End
            Begin VB.TextBox TxtData 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Index           =   2
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   6495
               Width           =   2370
            End
            Begin VB.TextBox TxtData 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Index           =   1
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   5940
               Width           =   2370
            End
            Begin VB.TextBox TxtData 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   420
               Index           =   0
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   146
               Top             =   5400
               Width           =   2370
            End
            Begin VSFlex8UCtl.VSFlexGrid FGSales 
               Height          =   4605
               Left            =   6720
               TabIndex        =   139
               Top             =   555
               Width           =   10440
               _cx             =   18415
               _cy             =   8123
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
               ForeColorFixed  =   16711680
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   0
               HighLight       =   0
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   500
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"ProjectsBillAlarm1.frx":07E1
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
               ExplorerBar     =   1
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
            Begin MSChart20Lib.MSChart ChartSales 
               Height          =   5565
               Left            =   -495
               OleObjectBlob   =   "ProjectsBillAlarm1.frx":0A12
               TabIndex        =   140
               Top             =   -135
               Width           =   7035
            End
            Begin MSChart20Lib.MSChart ChartSales1 
               Height          =   5010
               Left            =   0
               OleObjectBlob   =   "ProjectsBillAlarm1.frx":2EDB
               TabIndex        =   141
               Top             =   5250
               Width           =   7215
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»Ì«‰«  «·„⁄—Ê÷… Õ Ì  «—ÌŒ"
               Height          =   570
               Left            =   13605
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   135
               Width           =   3555
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Left            =   7365
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   135
               Width           =   6390
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               Caption         =   "«ﬁ· „»Ì⁄« "
               Height          =   570
               Left            =   15135
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   7050
               Width           =   1350
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "«ﬂ»— „»Ì⁄« "
               Height          =   285
               Left            =   15135
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   6645
               Width           =   1350
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "„ Ê”ÿ «·„»Ì⁄«  ··‰ﬁ«ÿ"
               Height          =   570
               Left            =   14055
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   5940
               Width           =   2430
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
               Height          =   570
               Left            =   15135
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   5400
               Width           =   1350
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   10020
            Left            =   -18855
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin VB.Timer tmrLoadingAll 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   0
               Top             =   0
            End
            Begin VB.Frame Frame8 
               Height          =   3012
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   2040
               Width           =   3852
               Begin VB.OptionButton chkT 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«Ã„«·Ì« "
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   2040
                  Width           =   1692
               End
               Begin VB.OptionButton chkR 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ÕÃ“"
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   1560
                  Width           =   1692
               End
               Begin VB.OptionButton chkP 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«‰ «Ã"
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   1080
                  Width           =   1692
               End
               Begin VB.OptionButton chkF 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· Õ·Ì· «·„«·Ï"
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   600
                  Width           =   1692
               End
               Begin VB.CommandButton Command8 
                  Caption         =   "Save"
                  Height          =   492
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   2400
                  Width           =   1932
               End
               Begin VB.Label Label32 
                  Alignment       =   2  'Center
                  Caption         =   " »œ√ «·‘«‘… »"
                  Height          =   372
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   240
                  Width           =   3012
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "„‰  «—ÌŒ"
               Height          =   2052
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   3000
               Width           =   4575
               Begin VB.CommandButton Command4 
                  Caption         =   "Save"
                  Height          =   492
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   1320
                  Width           =   1932
               End
               Begin MSComCtl2.DTPicker dtpFromDate 
                  Height          =   312
                  Left            =   960
                  TabIndex        =   64
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   2004
                  _ExtentX        =   3545
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   96337923
                  CurrentDate     =   37140
               End
               Begin MSComCtl2.DTPicker dtpToDate 
                  Height          =   312
                  Left            =   960
                  TabIndex        =   65
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   2004
                  _ExtentX        =   3545
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   96337923
                  CurrentDate     =   37140
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï  «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   252
                  Left            =   3228
                  TabIndex        =   67
                  Top             =   720
                  Width           =   744
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰  «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   252
                  Left            =   3240
                  TabIndex        =   66
                  Top             =   240
                  Width           =   744
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   " €ÌÌ— «⁄œ«œ«  «·„ƒﬁ « "
               Height          =   2535
               Left            =   10065
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   3000
               Width           =   4596
               Begin VB.CommandButton Command9 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   1920
                  Width           =   1092
               End
               Begin VB.CommandButton Command3 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   1320
                  Width           =   1092
               End
               Begin VB.CommandButton Command2 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   840
                  Width           =   1092
               End
               Begin VB.CommandButton Command1 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   360
                  Width           =   1092
               End
               Begin MSComCtl2.DTPicker IntervalAll 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "h:mm:ss AMPM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   4
                  EndProperty
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   56
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   65929219
                  UpDown          =   -1  'True
                  CurrentDate     =   37140.875
               End
               Begin MSComCtl2.DTPicker IntervalTab 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "h:mm:ss AMPM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   4
                  EndProperty
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   57
                  TabStop         =   0   'False
                  Top             =   840
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   65929219
                  UpDown          =   -1  'True
                  CurrentDate     =   37140.875
               End
               Begin MSComCtl2.DTPicker IntervalData 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "h:mm:ss AMPM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   4
                  EndProperty
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   58
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   65929219
                  UpDown          =   -1  'True
                  CurrentDate     =   37140.875
               End
               Begin MSComCtl2.DTPicker DataUpdate 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "h:mm:ss AMPM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   4
                  EndProperty
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   114
                  TabStop         =   0   'False
                  Top             =   1920
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   65929219
                  UpDown          =   -1  'True
                  CurrentDate     =   37140.875
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ÕœÌÀ «·»Ì«‰« "
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   1920
                  Width           =   1095
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ‰ﬁ· »Ì‰ «·«”ÿ—"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1320
                  Width           =   1092
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ‰ﬁ· »Ì‰ «·‘«‘« "
                  Height          =   372
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ÕœÌÀ «·ﬂ·"
                  Height          =   372
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   360
                  Width           =   852
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "«· Õ·Ì· «·„«·Ï"
               Height          =   2880
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   0
               Width           =   4596
               Begin VB.ComboBox cbSection1 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":5393
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":53A3
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   600
                  Width           =   2052
               End
               Begin VB.ComboBox cbSection2 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":53D3
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":53E3
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1080
                  Width           =   2052
               End
               Begin VB.ComboBox cbSection3 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":5413
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":5423
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   1560
                  Width           =   2052
               End
               Begin VB.ComboBox cbSection4 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":5453
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":5463
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   2040
                  Width           =   2052
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ Ê«Õœ"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   600
                  Width           =   852
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «À‰Ì‰"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   1080
                  Width           =   852
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·À«·À"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   1560
                  Width           =   852
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·—«»⁄"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   2040
                  Width           =   852
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "«·«‰ «Ã"
               Height          =   2880
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   0
               Width           =   4575
               Begin VB.ComboBox cbPSection1 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":5493
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":54A3
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   600
                  Width           =   2052
               End
               Begin VB.ComboBox cbPSection4 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":54D2
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":54E2
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   2040
                  Width           =   2052
               End
               Begin VB.ComboBox cbPSection3 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":5511
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":5521
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   1560
                  Width           =   2052
               End
               Begin VB.ComboBox cbPSection2 
                  Height          =   315
                  ItemData        =   "ProjectsBillAlarm1.frx":5550
                  Left            =   960
                  List            =   "ProjectsBillAlarm1.frx":5560
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   2052
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·—«»⁄"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   2040
                  Width           =   852
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·À«·À"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   1560
                  Width           =   852
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «À‰Ì‰"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   852
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ Ê«Õœ"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   600
                  Width           =   852
               End
            End
            Begin VB.Frame Frame7 
               Height          =   2052
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   0
               Width           =   3852
               Begin VB.OptionButton opt_TabChangeNot 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œ„  ›⁄Ì· «· ‰ﬁ·"
                  Height          =   312
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   840
                  Width           =   1572
               End
               Begin VB.OptionButton opt_TabChange 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ›⁄Ì· «· ‰ﬁ·"
                  Height          =   312
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   480
                  Width           =   2052
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Save"
                  Height          =   492
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   1320
                  Width           =   1572
               End
               Begin VB.CheckBox chkInvisble 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Œ›«¡ «·—”„ «·»Ì«‰Ï"
                  Height          =   492
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   3012
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   10020
            Left            =   -19155
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic10 
               Height          =   3030
               Left            =   5730
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   4545
               Width           =   5340
               _cx             =   9419
               _cy             =   5345
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
               Caption         =   "«·„»Ì⁄«   "
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   3
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
               Begin VB.TextBox txtAvg 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   1365
                  Width           =   2745
               End
               Begin VB.TextBox txtSalestotal 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   810
                  Width           =   2745
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„ Ê”ÿ «·”⁄— ··› —…"
                  Height          =   300
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   1365
                  Width           =   1590
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ﬁÌ„… «·„»Ì⁄«  «·‘Â—Ì…"
                  Height          =   300
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   810
                  Width           =   1875
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   6900
               Left            =   11235
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   3030
               Width           =   5745
               _cx             =   10134
               _cy             =   12171
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
               Begin VSFlex8UCtl.VSFlexGrid fg_Charge_Totals 
                  Height          =   6240
                  Left            =   0
                  TabIndex        =   91
                  Top             =   480
                  Width           =   5325
                  _cx             =   9393
                  _cy             =   11007
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial (Arabic)"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
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
                  BackColorAlternate=   16776960
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
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":558F
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
                  ExplorerBar     =   1
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
               Begin VB.Label Label28 
                  Alignment       =   2  'Center
                  Caption         =   "«·ﬂ„Ì«  «·„‘ÕÊ‰…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1395
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   135
                  Width           =   2295
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   2640
               Left            =   11235
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   270
               Width           =   5745
               _cx             =   10134
               _cy             =   4657
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
               Begin VB.TextBox txtCharge 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   1530
                  Width           =   2880
               End
               Begin VB.TextBox TxttotalRecivedShippedQty 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1950
                  Width           =   2880
               End
               Begin VB.TextBox txtProduct_Total 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   285
                  Width           =   2880
               End
               Begin VB.TextBox txtYes 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   690
                  Width           =   2880
               End
               Begin VB.TextBox txtReserve 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   1080
                  Width           =   2880
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„Ê—œ «·ÌÊ„-;ﬂ„Ì… „—”·Â"
                  Height          =   285
                  Left            =   3015
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1530
                  Width           =   2160
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„Ê—œ «·ÌÊ„-;ﬂ„Ì… „” ·„Â"
                  Height          =   285
                  Left            =   3015
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   1950
                  Width           =   2160
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Caption         =   "√Ã„«·Ï √„ «— «·‘Â—"
                  Height          =   285
                  Left            =   3450
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   285
                  Width           =   1440
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«‰ «Ã «„”"
                  Height          =   300
                  Left            =   3450
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   690
                  Width           =   1440
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬂ„Ì… «·„ÕÃÊ“…"
                  Height          =   330
                  Left            =   3450
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   1080
                  Width           =   1440
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   3330
               Left            =   5730
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   270
               Width           =   5340
               _cx             =   9419
               _cy             =   5874
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
               Begin VB.TextBox TxtPayed 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   156
                  Top             =   2505
                  Width           =   1740
               End
               Begin VB.CommandButton Command10 
                  Caption         =   " ›«’Ì· «·„œ›Ê⁄« "
                  Height          =   420
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   155
                  Top             =   2505
                  Width           =   1440
               End
               Begin VB.CommandButton Command7 
                  Caption         =   " ›«’Ì· «·«Ì—«œ« "
                  Height          =   465
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   1350
                  Width           =   1440
               End
               Begin VB.CommandButton Command6 
                  Caption         =   " ›«’Ì· «·„’—Ê›« "
                  Height          =   435
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   1920
                  Width           =   1440
               End
               Begin VB.TextBox txtBank 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   285
                  Width           =   3030
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   840
                  Width           =   3030
               End
               Begin VB.TextBox txtRevenue 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   1350
                  Width           =   1740
               End
               Begin VB.TextBox txtExpenses 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   1920
                  Width           =   1740
               End
               Begin VB.Label Label42 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œ›Ê⁄«  «·ÌÊ„"
                  Height          =   285
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   157
                  Top             =   2505
                  Width           =   2025
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—’Ìœ «·»‰Êﬂ"
                  Height          =   285
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   285
                  Width           =   2025
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—’Ìœ «·Œ“‰…"
                  Height          =   285
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   840
                  Width           =   2025
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ì—«œ«  «·ÌÊ„"
                  Height          =   330
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   1350
                  Width           =   2025
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„’—Ê›«  «·ÌÊ„"
                  Height          =   300
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   1920
                  Width           =   2025
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   9525
               Left            =   0
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   270
               Width           =   5565
               _cx             =   9816
               _cy             =   16801
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
               Begin VB.TextBox txtInv 
                  Alignment       =   1  'Right Justify
                  Height          =   525
                  Left            =   435
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   225
                  Width           =   2865
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_MaterialTotal 
                  Height          =   8190
                  Left            =   0
                  TabIndex        =   71
                  Top             =   1140
                  Width           =   5160
                  _cx             =   9102
                  _cy             =   14446
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
                  Rows            =   12
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":5633
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
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ã„«·Ï ﬁÌ„… «·„Œ“Ê‰"
                  Height          =   480
                  Left            =   3300
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   225
                  Width           =   1575
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   10020
            Left            =   -19455
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin C1SizerLibCtl.C1Elastic Frm_Reserve 
               Height          =   4710
               Left            =   0
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   0
               Width           =   17160
               _cx             =   30268
               _cy             =   8308
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
               Caption         =   "«·ÕÃ“"
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_Reserve 
                  Height          =   4650
                  Left            =   -435
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":56A2
                  TabIndex        =   5
                  Top             =   255
                  Width           =   6780
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Reserve 
                  Height          =   4185
                  Left            =   6345
                  TabIndex        =   6
                  Top             =   585
                  Width           =   9225
                  _cx             =   16272
                  _cy             =   7382
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
                  Cols            =   13
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":7B5A
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
                  AutoSizeMouse   =   0   'False
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
            Begin C1SizerLibCtl.C1Elastic frm_Material 
               Height          =   4680
               Left            =   0
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   4950
               Width           =   17160
               _cx             =   30268
               _cy             =   8255
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
               Caption         =   "«·„Ê«œ «·Œ«„"
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_Material 
                  Height          =   4650
                  Left            =   -285
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":7D4A
                  TabIndex        =   8
                  Top             =   255
                  Width           =   6480
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Material 
                  Height          =   4170
                  Left            =   6345
                  TabIndex        =   9
                  Top             =   450
                  Width           =   9225
                  _cx             =   16272
                  _cy             =   7355
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
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":A202
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   10020
            Left            =   -19755
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin VB.Timer tmrStoping 
               Enabled         =   0   'False
               Interval        =   60000
               Left            =   6600
               Top             =   8640
            End
            Begin VB.Timer tmrLoading 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   6000
               Top             =   8640
            End
            Begin C1SizerLibCtl.C1Elastic Frm_Charge 
               Height          =   4680
               Left            =   150
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   4950
               Width           =   17010
               _cx             =   30004
               _cy             =   8255
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
               Caption         =   "”‰œ ‘Õ‰"
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_Charge 
                  Height          =   4650
                  Left            =   -435
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":A271
                  TabIndex        =   12
                  Top             =   255
                  Width           =   7635
               End
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   4155
                  Left            =   7350
                  TabIndex        =   13
                  Top             =   420
                  Width           =   8220
                  _cx             =   14499
                  _cy             =   7329
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
                  Cols            =   15
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":C729
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
            Begin C1SizerLibCtl.C1Elastic Frm_Production 
               Height          =   4710
               Left            =   150
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   135
               Width           =   17010
               _cx             =   30004
               _cy             =   8308
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
               Caption         =   "«„— «·«‰ «Ã"
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_product 
                  Height          =   4650
                  Left            =   -150
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":C97A
                  TabIndex        =   15
                  Top             =   255
                  Width           =   7350
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Product 
                  Height          =   4185
                  Left            =   7200
                  TabIndex        =   16
                  Top             =   585
                  Width           =   8235
                  _cx             =   14526
                  _cy             =   7382
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
                  Cols            =   7
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":EE32
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   10020
            Left            =   -20055
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin C1SizerLibCtl.C1Elastic frm_Bank 
               Height          =   4710
               Left            =   150
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   135
               Width           =   8565
               _cx             =   15108
               _cy             =   8308
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
               Caption         =   "«—’œ… «·»‰Êﬂ"
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_Bankes 
                  Height          =   4185
                  Left            =   0
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":EF49
                  TabIndex        =   19
                  Top             =   525
                  Width           =   2730
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Bank 
                  Height          =   4065
                  Left            =   2880
                  TabIndex        =   20
                  Top             =   465
                  Width           =   4890
                  _cx             =   8625
                  _cy             =   7170
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":11401
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
            End
            Begin C1SizerLibCtl.C1Elastic frm_Boxes 
               Height          =   4710
               Left            =   8715
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   135
               Width           =   8445
               _cx             =   14896
               _cy             =   8308
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
               Caption         =   "«—’œ… «·Œ“‰ Ê«·⁄Âœ"
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin VSFlex8UCtl.VSFlexGrid Fg_Boxes 
                  Height          =   4050
                  Left            =   2745
                  TabIndex        =   22
                  Top             =   480
                  Width           =   4905
                  _cx             =   8652
                  _cy             =   7144
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":11566
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
                  AutoSizeMouse   =   0   'False
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
               Begin MSChart20Lib.MSChart chrt_Boxes 
                  Height          =   4260
                  Left            =   -285
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":11694
                  TabIndex        =   23
                  Top             =   525
                  Width           =   3315
               End
            End
            Begin C1SizerLibCtl.C1Elastic frm_Receipt 
               Height          =   4695
               Left            =   8715
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   4905
               Width           =   8445
               _cx             =   14896
               _cy             =   8281
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
               Caption         =   "„ﬁ»Ê÷« "
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_Receipts 
                  Height          =   4380
                  Left            =   0
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":13B4C
                  TabIndex        =   25
                  Top             =   420
                  Width           =   2745
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Receipts 
                  Height          =   4185
                  Left            =   2745
                  TabIndex        =   26
                  Top             =   465
                  Width           =   5055
                  _cx             =   8916
                  _cy             =   7382
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":16004
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
            End
            Begin C1SizerLibCtl.C1Elastic frm_Expenses 
               Height          =   4695
               Left            =   150
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   4905
               Width           =   8565
               _cx             =   15108
               _cy             =   8281
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
               Caption         =   "„’—Ê›« "
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               Begin MSChart20Lib.MSChart chrt_Expenses 
                  Height          =   4110
                  Left            =   -435
                  OleObjectBlob   =   "ProjectsBillAlarm1.frx":1609D
                  TabIndex        =   28
                  Top             =   615
                  Width           =   3315
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Expenses 
                  Height          =   4035
                  Left            =   2880
                  TabIndex        =   29
                  Top             =   465
                  Width           =   4755
                  _cx             =   8387
                  _cy             =   7117
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":18555
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
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic12 
            Height          =   10020
            Left            =   -18255
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic13 
               Height          =   4980
               Left            =   0
               TabIndex        =   159
               TabStop         =   0   'False
               Top             =   0
               Width           =   17310
               _cx             =   30533
               _cy             =   8784
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   3735
                  Left            =   8070
                  TabIndex        =   160
                  Top             =   1110
                  Width           =   8070
                  _cx             =   14235
                  _cy             =   6588
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
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":185EF
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
                  Height          =   3735
                  Left            =   0
                  TabIndex        =   161
                  Top             =   1140
                  Width           =   8070
                  _cx             =   14235
                  _cy             =   6588
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
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"ProjectsBillAlarm1.frx":18686
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic15 
                  Height          =   1140
                  Left            =   8070
                  TabIndex        =   162
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   8070
                  _cx             =   14235
                  _cy             =   2011
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
                  Begin MSComCtl2.DTPicker FrmDate 
                     Height          =   390
                     Left            =   4260
                     TabIndex        =   163
                     Top             =   435
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   65929217
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker ToDate 
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   164
                     Top             =   435
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   65929217
                     CurrentDate     =   41640
                  End
                  Begin ALLButtonS.ALLButton ALLButton1 
                     Height          =   390
                     Left            =   120
                     TabIndex        =   165
                     Top             =   435
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
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
                     BCOL            =   16711680
                     BCOLO           =   16711680
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "ProjectsBillAlarm1.frx":18719
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„ÊŸ›Ì‰"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   14.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   375
                     Index           =   0
                     Left            =   -3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   168
                     Top             =   0
                     Width           =   7920
                  End
                  Begin VB.Label Label45 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰  «—ÌŒ"
                     Height          =   315
                     Left            =   5820
                     RightToLeft     =   -1  'True
                     TabIndex        =   167
                     Top             =   450
                     Width           =   690
                  End
                  Begin VB.Label Label44 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï  «—ÌŒ"
                     Height          =   270
                     Left            =   3420
                     RightToLeft     =   -1  'True
                     TabIndex        =   166
                     Top             =   450
                     Width           =   810
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic16 
                  Height          =   1140
                  Left            =   120
                  TabIndex        =   169
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   7920
                  _cx             =   13970
                  _cy             =   2011
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
                  Begin MSComCtl2.DTPicker FrmDate1 
                     Height          =   390
                     Left            =   4260
                     TabIndex        =   170
                     Top             =   435
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   65929217
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker ToDate1 
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   171
                     Top             =   435
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   65929217
                     CurrentDate     =   41640
                  End
                  Begin ALLButtonS.ALLButton ALLButton2 
                     Height          =   390
                     Left            =   120
                     TabIndex        =   172
                     Top             =   435
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
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
                     BCOL            =   16711680
                     BCOLO           =   16711680
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "ProjectsBillAlarm1.frx":18735
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.Label Label46 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï  «—ÌŒ"
                     Height          =   270
                     Left            =   3420
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   450
                     Width           =   810
                  End
                  Begin VB.Label Label43 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰  «—ÌŒ"
                     Height          =   315
                     Left            =   5820
                     RightToLeft     =   -1  'True
                     TabIndex        =   174
                     Top             =   450
                     Width           =   690
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„Œ«“‰"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   14.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   375
                     Index           =   3
                     Left            =   -4080
                     RightToLeft     =   -1  'True
                     TabIndex        =   173
                     Top             =   0
                     Width           =   7920
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic14 
               Height          =   5100
               Left            =   0
               TabIndex        =   176
               TabStop         =   0   'False
               Top             =   4980
               Width           =   17160
               _cx             =   30268
               _cy             =   8996
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic17 
                  Height          =   4965
                  Left            =   0
                  TabIndex        =   177
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   15885
                  _cx             =   28019
                  _cy             =   8758
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
                  Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
                     Height          =   3720
                     Left            =   8085
                     TabIndex        =   178
                     Top             =   1110
                     Width           =   8070
                     _cx             =   14235
                     _cy             =   6562
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
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"ProjectsBillAlarm1.frx":18751
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
                  Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
                     Height          =   3720
                     Left            =   0
                     TabIndex        =   179
                     Top             =   1140
                     Width           =   8085
                     _cx             =   14261
                     _cy             =   6562
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
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"ProjectsBillAlarm1.frx":187E8
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
                  Begin C1SizerLibCtl.C1Elastic C1Elastic18 
                     Height          =   1140
                     Left            =   8085
                     TabIndex        =   180
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   8070
                     _cx             =   14235
                     _cy             =   2011
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
                     Begin MSComCtl2.DTPicker FrmDate2 
                        Height          =   390
                        Left            =   4260
                        TabIndex        =   181
                        Top             =   435
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   688
                        _Version        =   393216
                        Format          =   65929217
                        CurrentDate     =   41640
                     End
                     Begin MSComCtl2.DTPicker ToDate2 
                        Height          =   390
                        Left            =   1800
                        TabIndex        =   182
                        Top             =   435
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   688
                        _Version        =   393216
                        Format          =   65929217
                        CurrentDate     =   41640
                     End
                     Begin ALLButtonS.ALLButton ALLButton3 
                        Height          =   390
                        Left            =   120
                        TabIndex        =   183
                        Top             =   435
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   688
                        BTYPE           =   3
                        TX              =   "⁄—÷"
                        ENAB            =   -1  'True
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
                        BCOL            =   16711680
                        BCOLO           =   16711680
                        FCOL            =   16777215
                        FCOLO           =   16777215
                        MCOL            =   12632256
                        MPTR            =   1
                        MICON           =   "ProjectsBillAlarm1.frx":1887F
                        UMCOL           =   -1  'True
                        SOFT            =   0   'False
                        PICPOS          =   0
                        NGREY           =   0   'False
                        FX              =   0
                        HAND            =   0   'False
                        CHECK           =   0   'False
                        VALUE           =   0   'False
                     End
                     Begin VB.Label Label48 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·Ï  «—ÌŒ"
                        Height          =   270
                        Left            =   3420
                        RightToLeft     =   -1  'True
                        TabIndex        =   186
                        Top             =   450
                        Width           =   810
                     End
                     Begin VB.Label Label47 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "„‰  «—ÌŒ"
                        Height          =   315
                        Left            =   5820
                        RightToLeft     =   -1  'True
                        TabIndex        =   185
                        Top             =   450
                        Width           =   690
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·⁄„·«¡"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00C00000&
                        Height          =   375
                        Index           =   4
                        Left            =   -3240
                        RightToLeft     =   -1  'True
                        TabIndex        =   184
                        Top             =   0
                        Width           =   7920
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic19 
                     Height          =   1140
                     Left            =   120
                     TabIndex        =   187
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   7920
                     _cx             =   13970
                     _cy             =   2011
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
                     Begin MSComCtl2.DTPicker FrmDate3 
                        Height          =   390
                        Left            =   4260
                        TabIndex        =   188
                        Top             =   435
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   688
                        _Version        =   393216
                        Format          =   65929217
                        CurrentDate     =   41640
                     End
                     Begin MSComCtl2.DTPicker ToDate3 
                        Height          =   390
                        Left            =   1800
                        TabIndex        =   189
                        Top             =   435
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   688
                        _Version        =   393216
                        Format          =   65929217
                        CurrentDate     =   41640
                     End
                     Begin ALLButtonS.ALLButton ALLButton4 
                        Height          =   390
                        Left            =   120
                        TabIndex        =   190
                        Top             =   435
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   688
                        BTYPE           =   3
                        TX              =   "⁄—÷"
                        ENAB            =   -1  'True
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
                        BCOL            =   16711680
                        BCOLO           =   16711680
                        FCOL            =   16777215
                        FCOLO           =   16777215
                        MCOL            =   12632256
                        MPTR            =   1
                        MICON           =   "ProjectsBillAlarm1.frx":1889B
                        UMCOL           =   -1  'True
                        SOFT            =   0   'False
                        PICPOS          =   0
                        NGREY           =   0   'False
                        FX              =   0
                        HAND            =   0   'False
                        CHECK           =   0   'False
                        VALUE           =   0   'False
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·„Ê—œÌ‰"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00C00000&
                        Height          =   375
                        Index           =   5
                        Left            =   -4080
                        RightToLeft     =   -1  'True
                        TabIndex        =   193
                        Top             =   0
                        Width           =   7920
                     End
                     Begin VB.Label Label50 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "„‰  «—ÌŒ"
                        Height          =   315
                        Left            =   5820
                        RightToLeft     =   -1  'True
                        TabIndex        =   192
                        Top             =   450
                        Width           =   690
                     End
                     Begin VB.Label Label49 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·Ï  «—ÌŒ"
                        Height          =   270
                        Left            =   3420
                        RightToLeft     =   -1  'True
                        TabIndex        =   191
                        Top             =   450
                        Width           =   810
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic100 
            Height          =   10020
            Left            =   45
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   45
            Width           =   17310
            _cx             =   30533
            _cy             =   17674
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic23 
               Height          =   10020
               Left            =   0
               TabIndex        =   207
               TabStop         =   0   'False
               Top             =   0
               Width           =   17310
               _cx             =   30533
               _cy             =   17674
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic24 
                  Height          =   690
                  Left            =   0
                  TabIndex        =   208
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   17310
                  _cx             =   30533
                  _cy             =   1217
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
                  Begin MSComCtl2.DTPicker KFromDate 
                     Height          =   390
                     Left            =   13380
                     TabIndex        =   209
                     Top             =   195
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   65929217
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker KToDate 
                     Height          =   390
                     Left            =   10680
                     TabIndex        =   210
                     Top             =   195
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   65929217
                     CurrentDate     =   41640
                  End
                  Begin ALLButtonS.ALLButton ALLButton6 
                     Height          =   390
                     Left            =   7440
                     TabIndex        =   211
                     Top             =   195
                     Width           =   2670
                     _ExtentX        =   4710
                     _ExtentY        =   688
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
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
                     BCOL            =   16711680
                     BCOLO           =   16711680
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "ProjectsBillAlarm1.frx":188B7
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton ALLButton7 
                     Height          =   390
                     Left            =   3960
                     TabIndex        =   236
                     Top             =   195
                     Width           =   2670
                     _ExtentX        =   4710
                     _ExtentY        =   688
                     BTYPE           =   3
                     TX              =   "ÿ»«⁄… «·›Ê« Ì—"
                     ENAB            =   -1  'True
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
                     BCOL            =   14871017
                     BCOLO           =   14871017
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "ProjectsBillAlarm1.frx":188D3
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton ALLButton8 
                     Height          =   390
                     Left            =   360
                     TabIndex        =   237
                     Top             =   195
                     Width           =   2910
                     _ExtentX        =   5133
                     _ExtentY        =   688
                     BTYPE           =   3
                     TX              =   "ÿ»«⁄… «·„’—Ê›« "
                     ENAB            =   -1  'True
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
                     BCOL            =   14871017
                     BCOLO           =   14871017
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "ProjectsBillAlarm1.frx":188EF
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.Label Label54 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï  «—ÌŒ"
                     Height          =   270
                     Left            =   12420
                     RightToLeft     =   -1  'True
                     TabIndex        =   213
                     Top             =   210
                     Width           =   810
                  End
                  Begin VB.Label Label53 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰  «—ÌŒ"
                     Height          =   315
                     Left            =   15180
                     RightToLeft     =   -1  'True
                     TabIndex        =   212
                     Top             =   210
                     Width           =   690
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid7 
                  Height          =   8070
                  Left            =   120
                  TabIndex        =   214
                  Top             =   1410
                  Width           =   10530
                  _cx             =   18574
                  _cy             =   14235
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
                  Rows            =   2
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"ProjectsBillAlarm1.frx":1890B
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
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid8 
                  Height          =   8670
                  Left            =   10680
                  TabIndex        =   215
                  Top             =   840
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   15293
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
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"ProjectsBillAlarm1.frx":18AB2
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic25 
                  Height          =   570
                  Left            =   0
                  TabIndex        =   216
                  TabStop         =   0   'False
                  Top             =   9480
                  Width           =   17310
                  _cx             =   30533
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
                  GridRows        =   5
                  GridCols        =   29
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"ProjectsBillAlarm1.frx":18B7A
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.TextBox Text6 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   240
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.TextBox Text7 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   14835
                     RightToLeft     =   -1  'True
                     TabIndex        =   234
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.TextBox Text5 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   12345
                     RightToLeft     =   -1  'True
                     TabIndex        =   221
                     Top             =   120
                     Width           =   1320
                  End
                  Begin VB.TextBox Text4 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   9930
                     RightToLeft     =   -1  'True
                     TabIndex        =   220
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.TextBox Text3 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   7665
                     RightToLeft     =   -1  'True
                     TabIndex        =   219
                     Top             =   120
                     Width           =   1320
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   2550
                     RightToLeft     =   -1  'True
                     TabIndex        =   218
                     Top             =   120
                     Width           =   1305
                  End
                  Begin VB.TextBox text2 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   4890
                     RightToLeft     =   -1  'True
                     TabIndex        =   217
                     Top             =   120
                     Width           =   1350
                  End
                  Begin VB.Label Label69 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "„ Ê”ÿ «·Â«„‘"
                     Height          =   210
                     Left            =   1485
                     RightToLeft     =   -1  'True
                     TabIndex        =   241
                     Top             =   180
                     Width           =   1050
                  End
                  Begin VB.Label Label67 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·ﬂ„Ì…"
                     Height          =   210
                     Left            =   16215
                     RightToLeft     =   -1  'True
                     TabIndex        =   235
                     Top             =   180
                     Width           =   960
                  End
                  Begin VB.Label Label60 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·’«›Ì «·—»Õ"
                     Height          =   210
                     Left            =   3870
                     RightToLeft     =   -1  'True
                     TabIndex        =   226
                     Top             =   180
                     Width           =   930
                  End
                  Begin VB.Label Label59 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·„’—Ê›« "
                     Height          =   210
                     Left            =   6255
                     RightToLeft     =   -1  'True
                     TabIndex        =   225
                     Top             =   180
                     Width           =   1320
                  End
                  Begin VB.Label Label57 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·—»Õ"
                     Height          =   210
                     Left            =   9000
                     RightToLeft     =   -1  'True
                     TabIndex        =   224
                     Top             =   180
                     Width           =   915
                  End
                  Begin VB.Label Label56 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «· ﬂ·›…"
                     Height          =   210
                     Left            =   11280
                     RightToLeft     =   -1  'True
                     TabIndex        =   223
                     Top             =   180
                     Width           =   1035
                  End
                  Begin VB.Label Label55 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
                     Height          =   210
                     Left            =   13695
                     RightToLeft     =   -1  'True
                     TabIndex        =   222
                     Top             =   180
                     Width           =   1065
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic26 
                  Height          =   570
                  Left            =   120
                  TabIndex        =   227
                  TabStop         =   0   'False
                  Top             =   810
                  Width           =   10575
                  _cx             =   18653
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
                  Begin VB.Label Label68 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄Ì‰« "
                     Height          =   240
                     Left            =   6765
                     RightToLeft     =   -1  'True
                     TabIndex        =   239
                     Top             =   120
                     Width           =   1635
                  End
                  Begin VB.Label Label58 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     Height          =   240
                     Left            =   6720
                     RightToLeft     =   -1  'True
                     TabIndex        =   238
                     Top             =   330
                     Width           =   1635
                  End
                  Begin VB.Label Label66 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     Height          =   240
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   233
                     Top             =   315
                     Width           =   1635
                  End
                  Begin VB.Label Label61 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·›Ê« Ì— «·„ ⁄œœ…"
                     Height          =   240
                     Left            =   4605
                     RightToLeft     =   -1  'True
                     TabIndex        =   232
                     Top             =   105
                     Width           =   1635
                  End
                  Begin VB.Label Label65 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     Height          =   240
                     Left            =   2655
                     RightToLeft     =   -1  'True
                     TabIndex        =   231
                     Top             =   315
                     Width           =   1515
                  End
                  Begin VB.Label Label64 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·›Ê« Ì— «·‰ﬁœÌ…"
                     Height          =   240
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   230
                     Top             =   120
                     Width           =   1515
                  End
                  Begin VB.Label Label63 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     Height          =   240
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   229
                     Top             =   315
                     Width           =   1515
                  End
                  Begin VB.Label Label62 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·›Ê« Ì— «·√Ã·…"
                     Height          =   240
                     Left            =   2655
                     RightToLeft     =   -1  'True
                     TabIndex        =   228
                     Top             =   105
                     Width           =   1515
                  End
               End
            End
         End
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Label39ddd"
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   120
         Width           =   5700
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Height          =   2676
      Left            =   0
      TabIndex        =   75
      Top             =   480
      Visible         =   0   'False
      Width           =   3804
      _cx             =   6710
      _cy             =   4720
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ProjectsBillAlarm1.frx":18D0B
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
   Begin C1SizerLibCtl.C1Elastic frm_alarm 
      Height          =   10860
      Left            =   0
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   0
      Width           =   17310
      _cx             =   30533
      _cy             =   19156
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
         Height          =   600
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   0
         Width           =   17400
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   450
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   128
               Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
               Top             =   15
               Width           =   2340
               _ExtentX        =   4128
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
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   45
               Width           =   855
            End
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Text            =   "modflag"
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   510
            Visible         =   0   'False
            Width           =   945
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   480
            Top             =   0
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
                  Picture         =   "ProjectsBillAlarm1.frx":18D7A
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":19114
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":194AE
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":19848
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":19BE2
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":19F7C
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":1A316
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1.frx":1A8B0
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ‰»Ì… «·„” Œ·’«  «·Œ«’… »«·„‘«—Ì⁄ Ê «· Ì ·„  ”œœ »«·ﬂ«„·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   2
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   120
            Width           =   7920
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "œ·«·«  «·«·Ê«‰"
         Height          =   555
         Left            =   10395
         RightToLeft     =   -1  'True
         TabIndex        =   119
         Top             =   10230
         Width           =   6030
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— „”œœ »«·ﬂ«„·"
            Height          =   255
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "„”œœ Ã“∆Ì«"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   510
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   2175
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton btnProductionSave 
         Caption         =   "Save"
         Height          =   540
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Top             =   2685
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSChart20Lib.MSChart chrt_alarm 
         Height          =   4755
         Left            =   645
         OleObjectBlob   =   "ProjectsBillAlarm1.frx":1AC4A
         TabIndex        =   131
         Top             =   0
         Width           =   18225
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   4950
         Left            =   0
         TabIndex        =   132
         Top             =   5235
         Width           =   17355
         _cx             =   30612
         _cy             =   8731
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
         Rows            =   2
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ProjectsBillAlarm1.frx":1D102
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
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   390
         Left            =   285
         TabIndex        =   133
         Top             =   10350
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
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
         ButtonImage     =   "ProjectsBillAlarm1.frx":1D33D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin Cfx62ClientServerCtl.Chart Chart1 
         Height          =   1125
         Left            =   0
         TabIndex        =   134
         Top             =   810
         Visible         =   0   'False
         Width           =   1515
         _Data_          =   "ProjectsBillAlarm1.frx":1D6D7
      End
      Begin Cfx62ClientServerCtl.Chart Chart2 
         Height          =   1260
         Left            =   1695
         TabIndex        =   135
         Top             =   780
         Visible         =   0   'False
         Width           =   1320
         _Data_          =   "ProjectsBillAlarm1.frx":1DB78
      End
      Begin Cfx62ClientServerCtl.Chart Chart3 
         Height          =   555
         Left            =   3000
         TabIndex        =   136
         Top             =   780
         Visible         =   0   'False
         Width           =   1725
         _Data_          =   "ProjectsBillAlarm1.frx":1E019
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   450
         Left            =   1500
         TabIndex        =   137
         Top             =   10350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   794
         ButtonStyle     =   1
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
         ButtonImage     =   "ProjectsBillAlarm1.frx":1E4BA
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ï ﬁÌ„… «·„Œ“Ê‰"
      Height          =   252
      Left            =   2412
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   0
      Width           =   1440
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4800
   End
End
Attribute VB_Name = "ProjectsBillAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Public SendForm As String
Dim tmr As Integer
Dim Curr_Row_Box As Integer
Dim Curr_Row_Bank As Integer
Dim Curr_Row_Expenses As Integer
Dim Curr_Row_Receipt As Integer
Dim Curr_Row_Product As Integer
Dim Curr_Row_Charge As Integer
Dim Curr_Row_Reserve As Integer
Dim Curr_Row_Material As Integer
Dim Curr_Row_MaterialTotal As Integer
Dim Curr_Charge_Totals As Integer

Dim loadingcount As Integer
Dim Move_Row_Fixed As Long
Dim Move_Row As Long
Dim Move_Tab_Fixed As Long
Dim Move_Tab As Long
 Dim DataUpdateSecond As Long
Dim DataUpdateAll As Long
Dim Move_y As Integer, Move_N As Integer
'meeeeeeeeeeeee
Dim total1, total2, total3, total4, total5 As Double
Dim Sum1, Sum2, Sum3 As Double
Public Sub FillGridWithData()

    'On Error GoTo ErrTrap
 
    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim sql As String
    sql = "SELECT  * FROM     project_billl  "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If
 
    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable
 
        rs.MoveFirst
        
        For X = 1 To rs.RecordCount
       
            ActualTotal = getBillPayedToproject(val(rs.Fields("id").value))
            Result = val(rs.Fields("total").value) - ActualTotal
            Dim Total As Double
            Total = IIf(val(rs.Fields("total").value) = 0, 1, val(rs.Fields("total").value))
            resultpercentage = Round(ActualTotal / Total * 100, 2)
 
            If val(rs.Fields("total").value) > ActualTotal Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("bill_date").value), "", rs.Fields("bill_date").value)
                .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), "", rs.Fields("project_no").value)
                .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), "", rs.Fields("project_name").value)
            
                .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs.Fields("End_user_name").value), "", rs.Fields("End_user_name").value)
            
                .TextMatrix(i, .ColIndex("Sub_user_name")) = IIf(IsNull(rs.Fields("Sub_user_name").value), "", rs.Fields("Sub_user_name").value)
            
                .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
 
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
            
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = Result

                If Result = val(.TextMatrix(i, .ColIndex("total"))) Then
                    .Cell(flexcpBackColor, i, 12, i, 12) = vbRed
                Else
                    .Cell(flexcpBackColor, i, 12, i, 12) = vbYellow
                End If
                
                chrt_alarm.ShowLegend = True
                chrt_alarm.ColumnCount = rs.RecordCount
                chrt_alarm.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_alarm.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_alarm.RowLabel = "Chart"
                End If
                
                chrt_alarm.Column = i
                chrt_alarm.Row = 1
                chrt_alarm.Data = val(.TextMatrix(i, .ColIndex("ResultPercentage")))
                chrt_alarm.ColumnLabel = .TextMatrix(i, .ColIndex("Project_name"))
 
                
            End If

            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Public Sub FillGridWithData2()

    'On Error GoTo ErrTrap
 
    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim sql As String
        With Me.VSFlexGrid6
        .Rows = 1
        .Clear flexClearScrollable
       End With
    sql = "SELECT  * FROM     project_billl  "
    sql = sql & " where bill_date>=" & SQLDate(FrmDate4, True) & " and bill_date<=" & SQLDate(ToDate4, True) & ""
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If
 
    With Me.VSFlexGrid6
        .Rows = 1
        .Clear flexClearScrollable
 
        rs.MoveFirst
        
        For X = 1 To rs.RecordCount
       
            ActualTotal = getBillPayedToproject(val(rs.Fields("id").value))
            Result = val(rs.Fields("total").value) - ActualTotal
            Dim Total As Double
            Total = IIf(val(rs.Fields("total").value) = 0, 1, val(rs.Fields("total").value))
            resultpercentage = Round(ActualTotal / Total * 100, 2)
 
            If val(rs.Fields("total").value) > ActualTotal Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("bill_date").value), "", rs.Fields("bill_date").value)
                .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), "", rs.Fields("project_no").value)
                .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), "", rs.Fields("project_name").value)
            
                .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs.Fields("End_user_name").value), "", rs.Fields("End_user_name").value)
            
                .TextMatrix(i, .ColIndex("Sub_user_name")) = IIf(IsNull(rs.Fields("Sub_user_name").value), "", rs.Fields("Sub_user_name").value)
            
                .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
 
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
            
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = Result

              '  If result = val(.TextMatrix(i, .ColIndex("total"))) Then
              '      .Cell(flexcpBackColor, i, 12, i, 12) = vbRed
              '  Else
              '      .Cell(flexcpBackColor, i, 12, i, 12) = vbYellow
              '  End If
              '
              '  chrt_alarm.ShowLegend = True
              '  chrt_alarm.ColumnCount = rs.RecordCount
              '  chrt_alarm.RowCount = 1
              '  If SystemOptions.UserInterface = ArabicInterface Then
              '          chrt_alarm.RowLabel = "«·—”„ «·»Ì«‰Ì"
              '  Else
              '          chrt_alarm.RowLabel = "Chart"
              '  End If
              '
              '  chrt_alarm.Column = i
              '  chrt_alarm.Row = 1
              '  chrt_alarm.Data = val(.TextMatrix(i, .ColIndex("ResultPercentage")))
              '  chrt_alarm.ColumnLabel = .TextMatrix(i, .ColIndex("Project_name"))
 '
                
            End If

            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub ALLButton1_Click()
FillEmployee
End Sub

Private Sub ALLButton2_Click()
FillStore
End Sub

Private Sub ALLButton3_Click()
FillCustomer
End Sub

Private Sub ALLButton4_Click()
FillSuppler
End Sub

Private Sub ALLButton5_Click()
FillGridWithData2
End Sub



Private Sub BtnCancel_Click()
    Me.Hide
End Sub
Sub FillStore()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
          VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 1
sql = " SELECT     dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, "
sql = sql & "                      dbo.Accounts.Account_NameEng ,"
sql = sql & " Bala=(select (SUM(DEV_Value1) + SUM(DEV_Value2)) FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
sql = sql & "             DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
sql = sql & "     FROM  DOUBLE_ENTREY_VOUCHERS as do WHERE (do.RecordDate >= " & SQLDate(FrmDate1, True) & " AND do.RecordDate <= " & SQLDate(Me.ToDate1.value, True) & ") and    do.Account_Code = dbo.ACCOUNTS.Account_Code and(do.Posted IS NULL)) x)"
sql = sql & " FROM         dbo.TblStore INNER JOIN"
sql = sql & "                      dbo.ACCOUNTS ON dbo.TblStore.Account_Code = dbo.ACCOUNTS.Account_Code"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With VSFlexGrid3
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Bala")) = IIf(IsNull(rs2("Bala").value), 0, rs2("Bala").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Sub FillSuppler()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
          VSFlexGrid5.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid5.Rows = 1
sql = " SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Name, "
sql = sql & "                      dbo.Accounts.account_serial , dbo.Accounts.Account_NameEng, "
sql = sql & " Bala=(select (SUM(DEV_Value1) + SUM(DEV_Value2)) FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
sql = sql & "             DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
sql = sql & "     FROM  DOUBLE_ENTREY_VOUCHERS as do WHERE (do.RecordDate >= " & SQLDate(FrmDate3, True) & " AND do.RecordDate <= " & SQLDate(Me.ToDate3.value, True) & ") and    do.Account_Code = dbo.ACCOUNTS.Account_Code and(do.Posted IS NULL)) x)"
sql = sql & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
sql = sql & "                      dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
sql = sql & " Where (dbo.TblCustemers.Type = 2) And (dbo.TblCustemers.CusID <> 1)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With VSFlexGrid5
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Bala")) = IIf(IsNull(rs2("Bala").value), 0, rs2("Bala").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Sub FillCustomer()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
          VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid4.Rows = 1
sql = " SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Name, "
sql = sql & "                      dbo.Accounts.account_serial , dbo.Accounts.Account_NameEng, "
sql = sql & " Bala=(select (SUM(DEV_Value1) + SUM(DEV_Value2)) FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
sql = sql & "             DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
sql = sql & "     FROM  DOUBLE_ENTREY_VOUCHERS as do WHERE (do.RecordDate >= " & SQLDate(FrmDate2, True) & " AND do.RecordDate <= " & SQLDate(Me.ToDate2.value, True) & ") and    do.Account_Code = dbo.ACCOUNTS.Account_Code and(do.Posted IS NULL)) x)"
sql = sql & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
sql = sql & "                      dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
sql = sql & " Where (dbo.TblCustemers.Type = 1) And (dbo.TblCustemers.CusID <> 2)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With VSFlexGrid4
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Bala")) = IIf(IsNull(rs2("Bala").value), 0, rs2("Bala").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Sub FillEmployee()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
          VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
sql = "SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Account_code, dbo.ACCOUNTS.Account_Name, "
sql = sql & "  dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial ,"
sql = sql & " Bala=(select (SUM(DEV_Value1) + SUM(DEV_Value2)) FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
sql = sql & "             DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
sql = sql & "     FROM  DOUBLE_ENTREY_VOUCHERS as do WHERE (do.RecordDate >= " & SQLDate(FrmDate, True) & " AND do.RecordDate <= " & SQLDate(Me.ToDate.value, True) & ") and    do.Account_Code = dbo.ACCOUNTS.Account_Code and(do.Posted IS NULL)) x)"
sql = sql & " FROM         dbo.TblEmployee INNER JOIN"
sql = sql & "   dbo.ACCOUNTS ON dbo.TblEmployee.Account_code = dbo.ACCOUNTS.Account_Code        "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With VSFlexGrid1
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Bala")) = IIf(IsNull(rs2("Bala").value), 0, rs2("Bala").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Private Sub btnProductionSave_Click()

'If cbPSection1.ListIndex = -1 Then
'    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 1 ")
'    Exit Sub
'End If
'
'If cbPSection2.ListIndex = -1 Then
'    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 2 ")
'    Exit Sub
'End If

'If cbPSection3.ListIndex = -1 Then
'    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 3 ")
'    Exit Sub
'End If
'
'If cbPSection4.ListIndex = -1 Then
'    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 4 ")
'    Exit Sub
'End If


SaveSetting "Win_Sys_EX_B", "Setting", "PSection_1", cbPSection1.Text
SaveSetting "Win_Sys_EX_B", "Setting", "PSection_2", cbPSection2.Text
SaveSetting "Win_Sys_EX_B", "Setting", "PSection_3", cbPSection3.Text
SaveSetting "Win_Sys_EX_B", "Setting", "PSection_4", cbPSection4.Text

SaveSetting "Win_Sys_EX_B", "Setting", "PSec1", cbPSection1.ListIndex
SaveSetting "Win_Sys_EX_B", "Setting", "PSec2", cbPSection2.ListIndex
SaveSetting "Win_Sys_EX_B", "Setting", "PSec3", cbPSection3.ListIndex
SaveSetting "Win_Sys_EX_B", "Setting", "PSec4", cbPSection4.ListIndex

 PositionAllocation_Production

End Sub

Private Sub btnSave_Click()

If cbSection1.ListIndex = -1 Then
    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 1 ")
    Exit Sub
End If

If cbSection2.ListIndex = -1 Then
    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 2 ")
    Exit Sub
End If

If cbSection3.ListIndex = -1 Then
    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 3 ")
    Exit Sub
End If

If cbSection4.ListIndex = -1 Then
    MsgBox ("„‰ ›÷·ﬂ «Œ — „—»⁄ 4 ")
    Exit Sub
End If


SaveSetting "Win_Sys_EX_B", "Setting", "Section_1", cbSection1.Text
SaveSetting "Win_Sys_EX_B", "Setting", "Section_2", cbSection2.Text
SaveSetting "Win_Sys_EX_B", "Setting", "Section_3", cbSection3.Text
SaveSetting "Win_Sys_EX_B", "Setting", "Section_4", cbSection4.Text

SaveSetting "Win_Sys_EX_B", "Setting", "Sec1", cbSection1.ListIndex
SaveSetting "Win_Sys_EX_B", "Setting", "Sec2", cbSection2.ListIndex
SaveSetting "Win_Sys_EX_B", "Setting", "Sec3", cbSection3.ListIndex
SaveSetting "Win_Sys_EX_B", "Setting", "Sec4", cbSection4.ListIndex
PositionAllocation

End Sub




Private Sub C1Tab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmrStoping.Enabled = True
    tmrScrolling.Enabled = False
    Timer1.Enabled = False

End Sub

Private Sub cbPSection1_Click()

If cbPSection2.ListIndex = cbPSection1.ListIndex Then
        cbPSection2.ListIndex = -1
End If

If cbPSection3.ListIndex = cbPSection1.ListIndex Then
        cbPSection3.ListIndex = -1
End If


If cbPSection4.ListIndex = cbPSection1.ListIndex Then
        cbPSection4.ListIndex = -1
End If

End Sub


Private Sub cbPSection2_Click()
If cbPSection1.ListIndex = cbPSection2.ListIndex Then
        cbPSection1.ListIndex = -1
End If

If cbPSection3.ListIndex = cbPSection2.ListIndex Then
        cbPSection3.ListIndex = -1
End If

If cbPSection4.ListIndex = cbPSection2.ListIndex Then
        cbPSection4.ListIndex = -1
End If

End Sub


Private Sub cbPSection3_Click()

If cbPSection1.ListIndex = cbPSection3.ListIndex Then
        cbPSection1.ListIndex = -1
End If

If cbPSection2.ListIndex = cbPSection3.ListIndex Then
        cbPSection2.ListIndex = -1
End If


If cbPSection4.ListIndex = cbPSection3.ListIndex Then
        cbPSection4.ListIndex = -1
End If


End Sub

Private Sub cbPSection4_Click()

If cbPSection1.ListIndex = cbPSection4.ListIndex Then
        cbPSection1.ListIndex = -1
End If

If cbPSection2.ListIndex = cbPSection4.ListIndex Then
        cbPSection2.ListIndex = -1
End If


If cbPSection3.ListIndex = cbPSection4.ListIndex Then
        cbPSection3.ListIndex = -1
End If

End Sub

Private Sub cbSection1_Click()

If cbSection2.ListIndex = cbSection1.ListIndex Then
        cbSection2.ListIndex = -1
End If

If cbSection3.ListIndex = cbSection1.ListIndex Then
        cbSection3.ListIndex = -1
End If


If cbSection4.ListIndex = cbSection1.ListIndex Then
        cbSection4.ListIndex = -1
End If

End Sub

Private Sub cbSection2_Click()
If cbSection1.ListIndex = cbSection2.ListIndex Then
        cbSection1.ListIndex = -1
End If

If cbSection3.ListIndex = cbSection2.ListIndex Then
        cbSection3.ListIndex = -1
End If


If cbSection4.ListIndex = cbSection2.ListIndex Then
        cbSection4.ListIndex = -1
End If
End Sub


Private Sub cbSection3_Click()

If cbSection1.ListIndex = cbSection3.ListIndex Then
        cbSection1.ListIndex = -1
End If

If cbSection2.ListIndex = cbSection3.ListIndex Then
        cbSection2.ListIndex = -1
End If


If cbSection4.ListIndex = cbSection3.ListIndex Then
        cbSection4.ListIndex = -1
End If

End Sub

Private Sub cbSection4_Click()
If cbSection1.ListIndex = cbSection4.ListIndex Then
        cbSection1.ListIndex = -1
End If

If cbSection2.ListIndex = cbSection4.ListIndex Then
        cbSection2.ListIndex = -1
End If


If cbSection3.ListIndex = cbSection4.ListIndex Then
        cbSection3.ListIndex = -1
End If
End Sub

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ‰»Ì…    „” Œ·’«  ·„  ”œœ »«·ﬂ«„·", True, 2, 1, 1500
End Sub



Private Sub Command1_Click()

Dim ss As Long, mm As Long, hh As Long, allsecond As Long

ss = Second(IntervalAll.value)
mm = Minute(IntervalAll.value)
hh = Hour(IntervalAll.value)
mm = mm * 60
hh = hh * 60 * 60
allsecond = ss + mm + hh


SaveSetting "Win_Sys_EX_B", "Setting", "All_Time", IntervalAll.value

SaveSetting "Win_Sys_EX_B", "Setting", "Move_Row", allsecond
SaveSetting "Win_Sys_EX_B", "Setting", "Move_Row_Time", IntervalAll.value

SaveSetting "Win_Sys_EX_B", "Setting", "Move_Tab", allsecond
SaveSetting "Win_Sys_EX_B", "Setting", "Move_Tab_Time", IntervalAll.value

TimerSetting

End Sub

Private Sub Command10_Click()
  TxtPayed.Text = PayedTotals
PrintPayedVchr
End Sub

Private Sub Command2_Click()

Dim ss As Long, mm As Long, hh As Long, allsecond As Long

ss = Second(IntervalTab.value)
mm = Minute(IntervalTab.value)
hh = Hour(IntervalTab.value)
mm = mm * 60
hh = hh * 60 * 60
allsecond = ss + mm + hh


SaveSetting "Win_Sys_EX_B", "Setting", "Move_Tab", allsecond
SaveSetting "Win_Sys_EX_B", "Setting", "Move_Tab_Time", IntervalTab.value

TimerSetting

End Sub

Private Sub Command3_Click()
Dim ss As Long, mm As Long, hh As Long, allsecond As Long

ss = Second(IntervalData.value)
mm = Minute(IntervalData.value)
hh = Hour(IntervalData.value)
mm = mm * 60
hh = hh * 60 * 60
allsecond = ss + mm + hh


SaveSetting "Win_Sys_EX_B", "Setting", "Move_Row", allsecond
SaveSetting "Win_Sys_EX_B", "Setting", "Move_Row_Time", IntervalData.value

TimerSetting

End Sub

Private Sub Command4_Click()

SaveSetting "Win_Sys_EX_B", "Setting", "FromDate", dtpFromDate.value
SaveSetting "Win_Sys_EX_B", "Setting", "ToDate", dtpToDate.value

End Sub

Private Sub Command5_Click()
        SaveSetting "Win_Sys_EX_B", "Setting", "chart_visible", chkInvisble.value
        
        SaveSetting "Win_Sys_EX_B", "Setting", "TabChange", CInt(opt_TabChange.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "TabChangeNot", CInt(opt_TabChangeNot.value)
End Sub

Private Sub Command6_Click()
  txtExpenses.Text = ExpensesTotals
Print_Expenses

End Sub

Private Sub Command7_Click()
txtRevenue.Text = RevenueTotals

Print_Revenue
End Sub

Private Sub Command8_Click()
        SaveSetting "Win_Sys_EX_B", "Setting", "chkF", CInt(chkF.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "chkP", CInt(chkP.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "chkR", CInt(chkR.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "chkT", CInt(chkT.value)
End Sub

Private Sub Command9_Click()
Dim ss As Long, mm As Long, hh As Long, allsecond As Long

ss = Second(DataUpdate.value)
mm = Minute(DataUpdate.value)
hh = Hour(DataUpdate.value)
mm = mm * 60
hh = hh * 60 * 60
allsecond = ss + mm + hh


SaveSetting "Win_Sys_EX_B", "Setting", "LoadingAll", allsecond
'
SaveSetting "Win_Sys_EX_B", "Setting", "DataUpdate", DataUpdate.value

TimerSetting
End Sub


Private Sub tmrLoadingAll_Timer()
loadingcount = loadingcount + 1

If loadingcount >= DataUpdateAll Then
All_Total
 loadingcount = 0
 End If

End Sub

Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FillGridWithData
    
    Chart1.Gallery = Gallery_Pie
    Chart2.Gallery = Gallery_Curve
    Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
    FrmDate = FirstPeriodDateInthisYear
    FrmDate1 = FirstPeriodDateInthisYear
    FrmDate2 = FirstPeriodDateInthisYear
    FrmDate3 = FirstPeriodDateInthisYear
    FrmDate4 = FirstPeriodDateInthisYear
   KFromDate = FirstPeriodDateInthisYear

    ToDate.value = Date
    ToDate1.value = Date
    ToDate2.value = Date
    ToDate3.value = Date
    ToDate4.value = Date
    KToDate = Date
    
       'FrmDate4 = KFromDate
       '    ToDate4.value = KToDate
       
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        cahngelang
    End If
    
   If SendForm = "Dash" Then
            frm_Dash.Visible = True
            frm_alarm.Visible = False
            tmrLoading.Enabled = True
            tmrScrolling.Enabled = True
            Timer1.Enabled = True
  
          All_Total
  
    Else
            frm_alarm.Visible = True
            frm_Dash.Visible = False
 
           tmrLoading.Enabled = False
            tmrScrolling.Enabled = False
            Timer1.Enabled = False
   End If
    
    
'ShowBoxesAccouns

tmrLoading.Enabled = True
StartTab
PositionAllocation
Retrive_SectionData
PositionAllocation_Production
Retrive_SectionData_Production
TimerSetting
tmrLoadingAll.Enabled = True
'tmrLoadingAll.interval = DataUpdateAll
HideCharts
StartTab

End Sub

Private Sub HideCharts()
Dim i As Integer
i = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chart_visible") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chart_visible"))

Dim tt
tt = GetSetting("Win_Sys_EX_B", "Setting", "TabChange")
Move_y = IIf(GetSetting("Win_Sys_EX_B", "Setting", "TabChange") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "TabChange"))
Move_N = IIf(GetSetting("Win_Sys_EX_B", "Setting", "TabChangeNot") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "TabChangeNot"))

If Move_N <> 0 Then
        Timer1.Enabled = False
        opt_TabChangeNot = True
Else
        Timer1.Enabled = True
        opt_TabChange = True
End If

If i = 1 Then
'// hide all charts

chrt_Bankes.Visible = False
fg_Bank.Width = frm_Bank.Width
fg_Bank.Refresh



chrt_Boxes.Visible = False
Fg_Boxes.Width = frm_Boxes.Width
Fg_Boxes.Refresh

chrt_Expenses.Visible = False
fg_Expenses.Width = frm_Expenses.Width


chrt_Receipts.Visible = False
fg_Receipts.Width = frm_Receipt.Width


chrt_product.Visible = False
fg_Product.Width = Frm_Production.Width


chrt_Charge.Visible = False
GridInstallments.Width = Frm_Charge.Width

chrt_Reserve.Visible = False
fg_Reserve.Width = Frm_Reserve.Width

chrt_Material.Visible = False
fg_Material.Width = frm_Material.Width

Else


End If

End Sub


Private Sub TimerSetting()

Dim i1 As Long, i2 As Long, i3 As Long

i1 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Move_Row") = "", 1, GetSetting("Win_Sys_EX_B", "Setting", "Move_Row"))
Move_Row_Fixed = i1
Move_Row = i1


i2 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Move_Tab") = "", 1, GetSetting("Win_Sys_EX_B", "Setting", "Move_Tab"))
Move_Tab_Fixed = i2
Move_Tab = i2


'***********************************************************************
'i3 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "DataUpdateSecond") = "", 5, GetSetting("Win_Sys_EX_B", "Setting", "DataUpdateSecond"))
'DataUpdateSecond = i3
'DataUpdateAll = i3
'***********************************************************************


IntervalData.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Move_Row_Time") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "Move_Row_Time"))
IntervalTab.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Move_Tab_Time") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "Move_Tab_Time"))
IntervalAll.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "All_Time") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "All_Time"))

DataUpdateAll = IIf(GetSetting("Win_Sys_EX_B", "Setting", "LoadingAll") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "LoadingAll"))

DataUpdate.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "DataUpdate") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "DataUpdate"))


End Sub

Private Sub StartTab()

Dim chkF As Integer, chkP As Integer, chkR As Integer, chkT As Integer

chkF = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkF") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkF"))
chkP = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkP") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkP"))
chkR = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkR") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkR"))
chkT = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkT") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkT"))
C1Tab1.CurrTab = 0

If chkF <> 0 Then
C1Tab1.CurrTab = 0
Me.chkF.value = True
End If
   
   
If chkP <> 0 Then
C1Tab1.CurrTab = 1
Me.chkP.value = True
End If


If chkR <> 0 Then
C1Tab1.CurrTab = 2
Me.chkR.value = True
End If


If chkT <> 0 Then
C1Tab1.CurrTab = 3
Me.chkT.value = True
End If


End Sub






Private Sub PositionAllocation()
Dim Section_1 As String, Section_2 As String, Section_3 As String, Section_4 As String

Section_1 = GetSetting("Win_Sys_EX_B", "Setting", "Section_1")
Section_2 = GetSetting("Win_Sys_EX_B", "Setting", "Section_2")
Section_3 = GetSetting("Win_Sys_EX_B", "Setting", "Section_3")
Section_4 = GetSetting("Win_Sys_EX_B", "Setting", "Section_4")

If Section_1 = "«·»‰Êﬂ" Then
        frm_Bank.left = frm_Bank.Width + 250
        frm_Bank.top = 100
ElseIf Section_2 = "«·»‰Êﬂ" Then
        frm_Bank.left = 100
        frm_Bank.top = 100
ElseIf Section_3 = "«·»‰Êﬂ" Then
        frm_Bank.left = frm_Bank.Width + 250
        frm_Bank.top = frm_Bank.Height + 200
ElseIf Section_4 = "«·»‰Êﬂ" Then
        frm_Bank.left = 100
        frm_Bank.top = frm_Bank.Height + 200
End If


If Section_1 = "«·„ﬁ»Ê÷« " Then
        frm_Receipt.left = frm_Receipt.Width + 250
        frm_Receipt.top = 100
ElseIf Section_2 = "«·„ﬁ»Ê÷« " Then
        frm_Receipt.left = 100
        frm_Receipt.top = 100
ElseIf Section_3 = "«·„ﬁ»Ê÷« " Then
        frm_Receipt.left = frm_Receipt.Width + 250
        frm_Receipt.top = frm_Receipt.Height + 200
ElseIf Section_4 = "«·„ﬁ»Ê÷« " Then
        frm_Receipt.left = 100
        frm_Receipt.top = frm_Receipt.Height + 200
End If


If Section_1 = "«·„’—Ê›« " Then
        frm_Expenses.left = frm_Expenses.Width + 250
        frm_Expenses.top = 100
ElseIf Section_2 = "«·„’—Ê›« " Then
        frm_Expenses.left = 100
        frm_Expenses.top = 100
ElseIf Section_3 = "«·„’—Ê›« " Then
        frm_Expenses.left = frm_Expenses.Width + 250
        frm_Expenses.top = frm_Expenses.Height + 200
ElseIf Section_4 = "«·„’—Ê›« " Then
        frm_Expenses.left = 100
        frm_Expenses.top = frm_Expenses.Height + 200
End If


If Section_1 = "«·Œ“‰ Ê«·⁄Âœ" Then
        frm_Boxes.left = frm_Boxes.Width + 250
        frm_Boxes.top = 100
ElseIf Section_2 = "«·Œ“‰ Ê«·⁄Âœ" Then
        frm_Boxes.left = 100
        frm_Boxes.top = 100
ElseIf Section_3 = "«·Œ“‰ Ê«·⁄Âœ" Then
        frm_Boxes.left = frm_Boxes.Width + 250
        frm_Boxes.top = frm_Boxes.Height + 200
ElseIf Section_4 = "«·Œ“‰ Ê«·⁄Âœ" Then
        frm_Boxes.left = 100
        frm_Boxes.top = frm_Boxes.Height + 200
End If


End Sub

'////////////////////////////////////////
Private Sub PositionAllocation_Production()
Dim PSection_1 As String, PSection_2 As String, PSection_3 As String, PSection_4 As String

PSection_1 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_1")
PSection_2 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_2")
PSection_3 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_3")
PSection_4 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_4")

'If PSection_1 = "«„— «·«‰ «Ã" Then
'        Frm_Production.left = Frm_Production.Width + 250
'        Frm_Production.top = 100
'ElseIf PSection_2 = "«„— «·«‰ «Ã" Then
'        Frm_Production.left = 100
'        Frm_Production.top = 100
'ElseIf PSection_3 = "«„— «·«‰ «Ã" Then
'        Frm_Production.left = Frm_Production.Width + 250
'        Frm_Production.top = Frm_Production.Height + 200
'ElseIf PSection_4 = "«„— «·«‰ «Ã" Then
'        Frm_Production.left = 100
'        Frm_Production.top = Frm_Production.Height + 200
'End If
    

'If PSection_1 = "”‰œ ‘Õ‰" Then
'        Frm_Charge.left = Frm_Charge.Width + 250
'        Frm_Charge.top = 100
'ElseIf PSection_2 = "”‰œ ‘Õ‰" Then
'        Frm_Charge.left = 100
'        Frm_Charge.top = 100
'ElseIf PSection_3 = "”‰œ ‘Õ‰" Then
'        Frm_Charge.left = Frm_Charge.Width + 250
'        Frm_Charge.top = Frm_Charge.Height + 200
'ElseIf PSection_4 = "”‰œ ‘Õ‰" Then
'        Frm_Charge.left = 100
'        Frm_Charge.top = Frm_Charge.Height + 200
'End If
'
'
'If PSection_1 = "«·ÕÃ“" Then
'        Frm_Reserve.left = Frm_Reserve.Width + 250
'        Frm_Reserve.top = 100
'ElseIf PSection_2 = "«·ÕÃ“" Then
'        Frm_Reserve.left = 100
'        Frm_Reserve.top = 100
'ElseIf PSection_3 = "«·ÕÃ“" Then
'        Frm_Reserve.left = Frm_Reserve.Width + 250
'        Frm_Reserve.top = Frm_Reserve.Height + 200
'ElseIf PSection_4 = "«·ÕÃ“" Then
'        Frm_Reserve.left = 100
'        Frm_Reserve.top = Frm_Reserve.Height + 200
'End If
'
'If PSection_1 = "«·„Ê«œ «·Œ«„" Then
'        frm_Material.left = frm_Material.Width + 250
'        frm_Material.top = 100
'ElseIf PSection_2 = "«·„Ê«œ «·Œ«„" Then
'        frm_Material.left = 100
'        frm_Material.top = 100
'ElseIf PSection_3 = "«·„Ê«œ «·Œ«„" Then
'        frm_Material.left = frm_Material.Width + 250
'        frm_Material.top = frm_Material.Height + 200
'ElseIf PSection_4 = "«·„Ê«œ «·Œ«„" Then
'        frm_Material.left = 100
'        frm_Material.top = frm_Material.Height + 200
'End If
'
Dim Fromdate, ToDate
Fromdate = GetSetting("Win_Sys_EX_B", "Setting", "FromDate")
ToDate = GetSetting("Win_Sys_EX_B", "Setting", "ToDate")

dtpFromDate.value = IIf(Fromdate = "", Now, Fromdate)
dtpToDate.value = IIf(ToDate = "", Now, ToDate)


End Sub



Private Sub Retrive_SectionData()

Dim Section_1 As String, Section_2 As String, Section_3 As String, Section_4 As String
Dim Sec1 As Integer, Sec2 As Integer, Sec3 As Integer, Sec4 As Integer

Section_1 = GetSetting("Win_Sys_EX_B", "Setting", "Section_1")
Section_2 = GetSetting("Win_Sys_EX_B", "Setting", "Section_2")
Section_3 = GetSetting("Win_Sys_EX_B", "Setting", "Section_3")
Section_4 = GetSetting("Win_Sys_EX_B", "Setting", "Section_4")

Sec1 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Sec1") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "Sec1")))
Sec2 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Sec2") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "Sec2")))
Sec3 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Sec3") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "Sec3")))
Sec4 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Sec4") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "Sec4")))

cbSection1.ListIndex = Sec1
cbSection2.ListIndex = Sec2
cbSection3.ListIndex = Sec3
cbSection4.ListIndex = Sec4



End Sub
'//////////////////////
Private Sub Retrive_SectionData_Production()

Dim PSection_1 As String, PSection_2 As String, PSection_3 As String, PSection_4 As String
Dim PSec1 As Integer, PSec2 As Integer, PSec3 As Integer, PSec4 As Integer

PSection_1 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_1")
PSection_2 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_2")
PSection_3 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_3")
PSection_4 = GetSetting("Win_Sys_EX_B", "Setting", "PSection_4")

PSec1 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "PSec1") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "PSec1")))
PSec2 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "PSec2") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "PSec2")))
PSec3 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "PSec3") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "PSec3")))
PSec4 = IIf(GetSetting("Win_Sys_EX_B", "Setting", "PSec4") = "", -1, val(GetSetting("Win_Sys_EX_B", "Setting", "PSec4")))

cbPSection1.ListIndex = PSec1
cbPSection2.ListIndex = PSec2
cbPSection3.ListIndex = PSec3
cbPSection4.ListIndex = PSec4



End Sub


Function cahngelang()
    Label53.Caption = "From"
    Label54.Caption = "To"
    ALLButton6.Caption = "Process"
    Label67.Caption = "Volume"
    Label64.Caption = "Total Cash Inv"
    Label62.Caption = "Total Credit Inv"
    Label61.Caption = "Total ATM Inv"
    Label55.Caption = "Total Sales"
    Label56.Caption = "RMC "
    Label57.Caption = "Gross Profit"
'    Label58.Caption = "Sales Net"
    Label59.Caption = "Total Expenses"
    Label60.Caption = "Net Profit"
    ALLButton7.Caption = "Print Invoices"
    ALLButton8.Caption = "Print Expenses"
    Label68.Caption = "Sample"
    Label69.Caption = "Avg Margin"
    

    With VSFlexGrid7
        .TextMatrix(0, .ColIndex("Serial")) = "Ser"
        .TextMatrix(0, .ColIndex("InvDate")) = "Date"
        .TextMatrix(0, .ColIndex("InvNo")) = "Invoice No"
        .TextMatrix(0, .ColIndex("itemName")) = "Product Name"
        .TextMatrix(0, .ColIndex("cont")) = "Qty"
      '  .TextMatrix(0, .ColIndex("Price")) = "Sales Price"
        .TextMatrix(0, .ColIndex("total")) = "Sales Price"
        .TextMatrix(0, .ColIndex("RowValue")) = "Rmc Price"
        .TextMatrix(0, .ColIndex("profit")) = "Profit"
    End With
    
    With VSFlexGrid8
        .TextMatrix(0, .ColIndex("Serial")) = "Ser"
        .TextMatrix(0, .ColIndex("voucherNo")) = "Ref No"
        .TextMatrix(0, .ColIndex("voucherDate")) = "Date"
        .TextMatrix(0, .ColIndex("Des")) = "Description"
        .TextMatrix(0, .ColIndex("voucherValue")) = "Amount"
    End With
        
    '3333333333333333333
    
    
    Label1(2).Caption = " Dash Board"
    Me.Caption = Label1(2).Caption
    Frame1.Caption = "Color Map"
    Label3.Caption = "Fully"
    Label5.Caption = "Partial"

    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"
opt_TabChange.Caption = "Enable Tabs"
opt_TabChangeNot.Caption = "Disable Tabs"
Label32.Caption = "Start With"
chkF.Caption = "Fin Analysis"
chkP.Caption = "Productios"
chkR.Caption = "Reserve"
chkT.Caption = "Totals"
Command9.Caption = "Save"

Label34.Caption = "Update within"
Label33.Caption = "Total Delivered"
Label42.Caption = "Today A/P"
Command10.Caption = "A/P Details"
C1Elastic10.Caption = "Sales"

'//////////////////////////////////////////
frm_Boxes.Caption = " Boxes Balance "
    With Me.Fg_Boxes
        .TextMatrix(0, .ColIndex("Serial")) = "s"
        .TextMatrix(0, .ColIndex("BoxName")) = " Box Name"
        .TextMatrix(0, .ColIndex("BoxCredit")) = "Box Balance "
        .TextMatrix(0, .ColIndex("credit")) = "credit"
        .TextMatrix(0, .ColIndex("debit")) = "debit"

    End With
    
'//////////////////////////////

frm_Bank.Caption = " Banks  Balance "
    With Me.fg_Bank
        .TextMatrix(0, .ColIndex("Serial")) = "s"
        .TextMatrix(0, .ColIndex("BankName")) = " Bank Name "
        .TextMatrix(0, .ColIndex("BankCredit")) = "Bank  Balance "
        .TextMatrix(0, .ColIndex("credit")) = "credit"
        .TextMatrix(0, .ColIndex("debit")) = "debit"
    End With


'//////////////////////////////

frm_Receipt.Caption = " Receipts  "
    With Me.fg_Receipts
        .TextMatrix(0, .ColIndex("Serial")) = "s"
        .TextMatrix(0, .ColIndex("Branch")) = " Branch "
        .TextMatrix(0, .ColIndex("Value")) = "Analytical value "

    End With

'//////////////////////////////

frm_Expenses.Caption = " Expenses  "
With Me.fg_Expenses
    .TextMatrix(0, .ColIndex("Serial")) = "s"
    .TextMatrix(0, .ColIndex("Branch")) = " Branch "
    .TextMatrix(0, .ColIndex("Value")) = "Analytical value "

End With


'//////////////////////////////

Frm_Production.Caption = " Production Order  "
With Me.fg_Product
    .TextMatrix(0, .ColIndex("Ser")) = "s"
    .TextMatrix(0, .ColIndex("NoteSerial1")) = " NoteSerial1 "
    .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date "
    
    .TextMatrix(0, .ColIndex("CusName")) = " Customer Name "
    .TextMatrix(0, .ColIndex("FullCode")) = " Code "
       .TextMatrix(0, .ColIndex("ItemName")) = " Item Name "
    .TextMatrix(0, .ColIndex("showqty")) = "Quantity "
    
End With

'//////////////////////////////

Frm_Charge.Caption = " Production Order  "
With Me.GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "s"
    .TextMatrix(0, .ColIndex("Notes")) = " Pump "
    .TextMatrix(0, .ColIndex("CusName")) = "Customer Name "
    .TextMatrix(0, .ColIndex("HeyName")) = " Location "
     .TextMatrix(0, .ColIndex("productiontypename")) = " Type "
    
    .TextMatrix(0, .ColIndex("Contacttime")) = " Time "
    .TextMatrix(0, .ColIndex("ShowQty")) = " Required Quantity "
    .TextMatrix(0, .ColIndex("ShipedQty")) = " Charged Quantity "
    .TextMatrix(0, .ColIndex("diff")) = "Difference "
    
    .TextMatrix(0, .ColIndex("Emp_Name")) = " sale rep "
    .TextMatrix(0, .ColIndex("cus_mobile")) = " Mobile  "
    .TextMatrix(0, .ColIndex("transactioncomment")) = " Remark  "
End With


Frm_Reserve.Caption = " Reservation  "
With Me.fg_Reserve
    .TextMatrix(0, .ColIndex("Ser")) = "s"
    .TextMatrix(0, .ColIndex("Notes")) = " Pump "
    .TextMatrix(0, .ColIndex("CusName")) = "Customer Name "
    .TextMatrix(0, .ColIndex("HeyName")) = " Location "
     .TextMatrix(0, .ColIndex("productiontypename")) = " Type "
    
    .TextMatrix(0, .ColIndex("Contacttime")) = " Time "
    .TextMatrix(0, .ColIndex("ShowQty")) = " Required Quantity "

    .TextMatrix(0, .ColIndex("Emp_Name")) = " sale rep "
    .TextMatrix(0, .ColIndex("cus_mobile")) = " Mobile  "
    .TextMatrix(0, .ColIndex("transactioncomment")) = " Remark  "
End With


frm_Material.Caption = " Material  "
With Me.fg_Material
    .TextMatrix(0, .ColIndex("Ser")) = "s"
    .TextMatrix(0, .ColIndex("ItemName")) = " Item Name "
   .TextMatrix(0, .ColIndex("qty")) = "Quantity "
   
End With

'///////////////////////////////////////////////

Label22.Caption = "Month Total"
Label23.Caption = "Box Balance"
Label24.Caption = " Reserved Quantity "
Label25.Caption = "Shipped quantity"

Label19.Caption = "Bank Balance"
Label20.Caption = "Yesterdaya Production"
Label21.Caption = " Revenue Today "
Label27.Caption = "Expenses"

Label28.Caption = "Shipping"

With Me.fg_Charge_Totals
'    .TextMatrix(0, .ColIndex("CusNamee")) = " Customer Name "
    .TextMatrix(0, .ColIndex("total")) = " Quantity "
    .TextMatrix(0, .ColIndex("productiontypename")) = " Type "
End With

With Me.fg_MaterialTotal
    .TextMatrix(0, .ColIndex("ser")) = " S "
    .TextMatrix(0, .ColIndex("ItemName")) = " Item "
    .TextMatrix(0, .ColIndex("qty")) = " Quantity "
End With


'/////////////////////////////////////////
Frame3.Caption = "financial"
Label6.Caption = "Section 1"
Label7.Caption = "Section 2"
Label8.Caption = "Section 3 "
Label9.Caption = "Section 4"

Frame4.Caption = "Production"
Label13.Caption = "Section 1"
Label12.Caption = "Section 2"
Label11.Caption = "Section 3 "
Label10.Caption = "Section 4"

Frame5.Caption = "Timers Settings"
Label14.Caption = " Refersh All "
Label15.Caption = "Screens Change "
Label16.Caption = "Row Change"
Command1.Caption = "Activate"
Command2.Caption = "Activate"
Command3.Caption = "Activate"

Frame6.Caption = " Date Setting"
Label17.Caption = " From Date "
Label18.Caption = " To Date "

Label1(1).Caption = "Dash Board"
C1Tab1.Caption = " Financial|Production|Reservation|Totals|Settings|Sales|Emp&Clients&Suppliers|Projects|Waiting "
C1Tab1.TabCaption(8) = "Profit And loss"

Label26.Caption = "Inventory Total"

Label30.Caption = "Inventory Total"
Label29.Caption = "Average Price"

''////////////////////////////////
Command7.Caption = "Revenue Det."
Command6.Caption = "Expenses Det."


End Function


Private Sub GetBoxesData()

Show_BoxesAccouns

End Sub


Private Sub GetExpensesData()
        Dim StrSQL As String
        StrSQL = "select  b.branch_id , b.branch_name ,  b.branch_namee ,  sum (n.Note_Value ) Note_Value  from notes N , TblBranchesData b "
        StrSQL = StrSQL & " WHERE  n.branch_no = b.branch_id and  n.notetype = 5 and  "
        StrSQL = StrSQL & " N.NoteDate >= '" & SQLDate(dtpFromDate.value) & "'"
        StrSQL = StrSQL & " and N.NoteDate <= '" & SQLDate(dtpToDate.value) & "'"
        
        StrSQL = StrSQL & " group by b.branch_id , b.branch_name , b.branch_namee   "
        
        
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
         fg_Expenses.Rows = fg_Expenses.FixedRows
         
        If Rs_Temp.RecordCount > 0 Then
              With chrt_Expenses
                .ShowLegend = True
                .ColumnCount = Rs_Temp.RecordCount
                .RowCount = 1
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .RowLabel = "«·—”„ «·»Ì«‰Ì"
                    Else
                   .RowLabel = "Chart"
                   End If
             End With
             
             
          '  Chart4.
             
             
              Dim i As Integer
              With fg_Expenses
                  fg_Expenses.Rows = fg_Expenses.FixedRows + Rs_Temp.RecordCount
                  For i = 1 To fg_Expenses.Rows - 1
                  
                            chrt_Expenses.Column = i
                            chrt_Expenses.Row = 1
                            chrt_Expenses.Data = IIf(IsNull(Rs_Temp.Fields("Note_Value").value), "", Rs_Temp.Fields("Note_Value").value)
                        
                  
                  
                            .TextMatrix(i, .ColIndex("Serial")) = i
                            
                            If SystemOptions.UserInterface = ArabicInterface Then
                                      .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs_Temp("branch_name").value), "", Rs_Temp("branch_name").value)
                                          chrt_Expenses.ColumnLabel = IIf(IsNull(Rs_Temp.Fields("branch_name").value), "", Rs_Temp.Fields("branch_name").value)
                           Else
                                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs_Temp("branch_namee").value), "", Rs_Temp("branch_namee").value)
                                        chrt_Expenses.ColumnLabel = IIf(IsNull(Rs_Temp.Fields("branch_namee").value), "", Rs_Temp.Fields("branch_namee").value)
                           End If
                            
                            .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(Rs_Temp("Note_Value").value), "", Rs_Temp("Note_Value").value)
                           ' .TextMatrix(i, .ColIndex("Percent")) = IIf(IsNull(Rs_Temp("").value), "", Rs_Temp("").value)
                            Rs_Temp.MoveNext
                  Next
              End With
             
        End If
        Expenses_Percent
        
End Sub

Private Sub Expenses_Percent()
Dim i As Integer, value As Double

For i = 1 To fg_Expenses.Rows - 1
        value = value + val(fg_Expenses.TextMatrix(i, fg_Expenses.ColIndex("Value")))
Next

If value > 0 Then
For i = 1 To fg_Expenses.Rows - 1
        fg_Expenses.TextMatrix(i, fg_Expenses.ColIndex("Percent")) = Math.Round((val(fg_Expenses.TextMatrix(i, fg_Expenses.ColIndex("Value"))) / value) * 100, 2)
Next
End If

End Sub


Private Sub Receipts_Percent()
Dim i As Integer, value As Double

For i = 1 To fg_Receipts.Rows - 1
        value = value + val(fg_Receipts.TextMatrix(i, fg_Receipts.ColIndex("Value")))
Next

If value > 0 Then
For i = 1 To fg_Receipts.Rows - 1
        fg_Receipts.TextMatrix(i, fg_Receipts.ColIndex("Percent")) = Math.Round((val(fg_Receipts.TextMatrix(i, fg_Receipts.ColIndex("Value"))) / value) * 100, 2)
Next
End If

End Sub
Private Sub GetReceiptsData()
        Dim StrSQL As String
        StrSQL = " select  b.branch_id , b.branch_name , b.branch_namee , sum (n.Note_Value ) Note_Value  from notes N , TblBranchesData b "
        StrSQL = StrSQL & " WHERE  n.branch_no = b.branch_id and  n.notetype = 4 "
        
        StrSQL = StrSQL & " and  N.NoteDate >= '" & SQLDate(dtpFromDate.value) & "'"
        StrSQL = StrSQL & " and  N.NoteDate <= '" & SQLDate(dtpToDate.value) & "'"
        
        StrSQL = StrSQL & " group by b.branch_id , b.branch_name  , b.branch_namee  "
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        
        If Rs_Temp.RecordCount > 0 Then
              With chrt_Receipts
                .ShowLegend = True
                .ColumnCount = Rs_Temp.RecordCount
                .RowCount = 1
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .RowLabel = "«·—”„ «·»Ì«‰Ì"
                    Else
                   .RowLabel = "Chart"
                   End If
             End With
             
             Dim i As Integer
              With fg_Receipts
                  fg_Receipts.Rows = fg_Receipts.FixedRows + Rs_Temp.RecordCount
                  For i = 1 To fg_Receipts.Rows - 1
                  
                            chrt_Receipts.Column = i
                            chrt_Receipts.Row = 1
                            chrt_Receipts.Data = IIf(IsNull(Rs_Temp.Fields("Note_Value").value), "", Rs_Temp.Fields("Note_Value").value)
                           
                            .TextMatrix(i, .ColIndex("Serial")) = i
                            
                            If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs_Temp("branch_name").value), "", Rs_Temp("branch_name").value)
                             chrt_Receipts.ColumnLabel = IIf(IsNull(Rs_Temp.Fields("branch_name").value), "", Rs_Temp.Fields("branch_name").value)
                            Else
                             .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(Rs_Temp("branch_namee").value), "", Rs_Temp("branch_namee").value)
                              chrt_Receipts.ColumnLabel = IIf(IsNull(Rs_Temp.Fields("branch_namee").value), "", Rs_Temp.Fields("branch_namee").value)
                            End If
                            
                            .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(Rs_Temp("Note_Value").value), "", Rs_Temp("Note_Value").value)
                           ' .TextMatrix(i, .ColIndex("Percent")) = IIf(IsNull(Rs_Temp("").value), "", Rs_Temp("").value)
                           Rs_Temp.MoveNext
                           
                  Next
              End With
             
        End If
        Receipts_Percent
End Sub

Private Sub GetBanksData()
        
        Dim StrSQL As String
        StrSQL = " select   bankID , BankName  from  banksdata "
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Show_BankesData

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
tmrScrolling.Enabled = False
End Sub

Private Sub Form_Resize()
PositionAllocation
End Sub

 

Private Sub Timer1_Timer()

    Move_Tab = Move_Tab - 1
    
    If Move_Tab <= 0 Then
           C1Tab1.CurrTab = tmr
            tmr = tmr + 1
            If tmr = 4 Then
                    tmr = 0
            End If
            Move_Tab = Move_Tab_Fixed
    End If
    
End Sub



Private Sub tmrLoading_Timer()
  If SendForm = "Dash" Then
           ' C1Tab1.CurrTab = 0
            
            frm_Dash.Visible = True
            
            
            
            GetExpensesData
            DoEvents

           Sleep 50
            
            GetReceiptsData
            DoEvents
            
            
           Sleep 50
           
            GetBoxesData
            DoEvents
            
            Sleep 50
           
            GetBanksData
            DoEvents
            
            
            Sleep 50
            Show_POSData
            DoEvents
           '////////////////////////////////////
            
            Sleep 50
       '    C1Tab1.CurrTab = 1
              
           GetCharge
           DoEvents
           
           Sleep 50
           GetProduct2
           DoEvents
           
           '////////////////////////////////////
     '      C1Tab1.CurrTab = 2
           
           Sleep 50
           GetReserve2
           DoEvents
           
          Sleep 50
          GetMaterial
          DoEvents
    End If
    
    tmrLoading.Enabled = False
End Sub

Private Sub tmrScrolling_Timer()
       
Move_Row = Move_Row - 1
If Move_Row <= 0 Then
         Scroll_Bankes
         Scroll_Boxes
         Scroll_Expenses
         Scroll_Receipt
         Scroll_Product
         Scroll_Charge
         Scroll_Material
         Scroll_Reserve
         Scroll_MateriaTotall
         Scroll_Charge_Totals
         Move_Row = Move_Row_Fixed
End If

End Sub

Private Sub Scroll_Charge_Totals()
         fg_Charge_Totals.TopRow = Curr_Charge_Totals * 10
         Curr_Charge_Totals = Curr_Charge_Totals + 1
         If Curr_Charge_Totals * 10 > fg_Charge_Totals.Rows Then
                    Curr_Charge_Totals = 0
         End If
End Sub

Private Sub Scroll_MateriaTotall()
         fg_MaterialTotal.TopRow = Curr_Row_MaterialTotal * 10
         Curr_Row_MaterialTotal = Curr_Row_MaterialTotal + 1
         If Curr_Row_MaterialTotal * 10 > fg_MaterialTotal.Rows Then
                    Curr_Row_MaterialTotal = 0
         End If
End Sub


Private Sub Scroll_Material()
         fg_Material.TopRow = Curr_Row_Material * 10
         Curr_Row_Material = Curr_Row_Material + 1
         If Curr_Row_Material * 10 > fg_Material.Rows Then
                    Curr_Row_Material = 0
         End If
End Sub

Private Sub Scroll_Reserve()
         fg_Reserve.TopRow = Curr_Row_Reserve * 10
         Curr_Row_Reserve = Curr_Row_Reserve + 1
         If Curr_Row_Reserve * 10 > fg_Reserve.Rows Then
                    Curr_Row_Reserve = 0
         End If
End Sub


Private Sub Scroll_Charge()
         GridInstallments.TopRow = Curr_Row_Charge * 10
         Curr_Row_Charge = Curr_Row_Charge + 1
         If Curr_Row_Charge * 10 > GridInstallments.Rows Then
                    Curr_Row_Charge = 0
         End If
End Sub

Private Sub Scroll_Product()
         fg_Product.TopRow = Curr_Row_Product * 10
         Curr_Row_Product = Curr_Row_Product + 1
         If Curr_Row_Product * 10 > fg_Product.Rows Then
                    Curr_Row_Product = 0
         End If
End Sub

Private Sub Scroll_Boxes()
         Fg_Boxes.TopRow = Curr_Row_Box * 10
         Curr_Row_Box = Curr_Row_Box + 1
         If Curr_Row_Box * 10 > Fg_Boxes.Rows Then
                    Curr_Row_Box = 0
         End If
End Sub


Private Sub Scroll_Bankes()
          fg_Bank.TopRow = Curr_Row_Bank * 10
          Curr_Row_Bank = Curr_Row_Bank + 1
          If Curr_Row_Bank * 10 > fg_Bank.Rows Then
                    Curr_Row_Bank = 0
          End If
End Sub

Private Sub Scroll_Expenses()
         fg_Expenses.TopRow = Curr_Row_Expenses * 10
         Curr_Row_Expenses = Curr_Row_Expenses + 1
         If Curr_Row_Expenses * 10 > fg_Expenses.Rows Then
                    Curr_Row_Expenses = 0
         End If
End Sub

Private Sub Scroll_Receipt()
         fg_Receipts.TopRow = Curr_Row_Receipt * 10
         Curr_Row_Receipt = Curr_Row_Receipt + 1
         If Curr_Row_Receipt * 10 > fg_Receipts.Rows Then
                    Curr_Row_Receipt = 0
         End If
End Sub

Public Sub Show_BoxesAccouns()

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim FirstPeriod As Date
    Dim Balance As Double

 
    StrSQL = "SELECT * from TblBoxesData  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
       ' Load FrmBoxesAccounts

        With Fg_Boxes
        
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst


            
            
            
            For i = .FixedRows To .Rows - 1
                
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
             Else
             .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxNamee").value), "", rs("BoxNamee").value)
            End If
             
             
             
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
      
                getFirstPeriodDateInthisYear FirstPeriod
 
                Balance = GetActualAccountBalance(rs("Account_Code").value, 0, FirstPeriod, Date)

       '         .TextMatrix(i, .ColIndex("BoxCredit")) = Abs(Balance) 'GetActualAccountBalance(rs("Account_Code").value, branch_id, FirstPeriod, Date)


                
                .TextMatrix(i, .ColIndex("credit")) = Math.Abs(Credit_Debit(.TextMatrix(i, .ColIndex("AccountCode")), 0))
               .TextMatrix(i, .ColIndex("debit")) = Math.Abs(Credit_Debit(.TextMatrix(i, .ColIndex("AccountCode")), 1))
               

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                         .TextMatrix(i, .ColIndex("BoxCredit")) = "" & Abs(Balance) & ""
                        .TextMatrix(i, .ColIndex("Type")) = "„œÌ‰"
                     
                    ElseIf Balance < 0 Then
                    
                        .TextMatrix(i, .ColIndex("Type")) = "œ«∆‰"
                        .TextMatrix(i, .ColIndex("BoxCredit")) = " ( " & CStr(Abs(Balance)) & " ) "
                        
                    Else
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                Else

                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Debit"
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Credit"
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                End If
                
                
                
                With chrt_Boxes
                chrt_Boxes.ShowLegend = True
                chrt_Boxes.ColumnCount = rs.RecordCount
                chrt_Boxes.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Boxes.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Boxes.RowLabel = "Chart"
                End If
                End With
                chrt_Boxes.Column = i
                chrt_Boxes.Row = 1
                chrt_Boxes.Data = val(.TextMatrix(i, .ColIndex("BoxCredit")))
                
               If SystemOptions.UserInterface = ArabicInterface Then
                             chrt_Boxes.ColumnLabel = IIf(IsNull(rs.Fields("BoxName").value), "", rs.Fields("BoxName").value)
                Else
                                  chrt_Boxes.ColumnLabel = IIf(IsNull(rs.Fields("BoxNamee").value), "", rs.Fields("BoxNamee").value)
                End If
                
                
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    Exit Sub

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where TblBoxesData.BoxID <>1"
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblBoxesData.BoxID,dbo.TblBoxesData.BoxName, QryBoxesCredit.BoxCredit" & " FROM dbo.TblBoxesData INNER JOIN " & "dbo.QryBoxesCredit() QryBoxesCredit ON dbo.TblBoxesData.BoxID = QryBoxesCredit.BoxID"

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <>1"
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
      '  Load FrmBoxesAccounts

        With frm_Boxes
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)

                If Not IsNull(rs("BoxCredit").value) Then
                    .TextMatrix(i, .ColIndex("BoxCredit")) = Format(rs("BoxCredit").value, SystemOptions.SysDefCurrencyForamt)
                Else
                    .TextMatrix(i, .ColIndex("BoxCredit")) = 0
                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Msg = "·«ÌÊÃœ «Ï Œ“‰ „”Ã·… ›Ï «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«·«Ì„ﬂ‰ ⁄—÷ «·Œ“‰ «·Õ«·Ì… ›Ï «·»—‰«„Ã...!!!"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

End Sub


Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer

    With FGSales

        For i = .FixedRows To .Rows - 1
        
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 10) = &HFFFFC0
            Else
                .Cell(flexcpBackColor, i, 1, i, 10) = vbWhite
            End If

        Next i

    End With

    'line_no1 = IntCounter

End Sub
Public Sub Show_POSData()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim FirstPeriod As Date
    Dim Balance As Double
          Dim Maxsalename As String
       Dim minsalename As String
    
          Dim Fromdate As Date
       Dim ToDate As Date
     Fromdate = "01/" & Month(Date) & "/" & year(Date)
ToDate = MonthLastDay(Date)
   Dim totaldouble As Double
   
            Dim Maxsale     As Double
Dim Minsale As Double
   
  StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.Transaction_Details.Price) AS Total, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, "
StrSQL = StrSQL & "  dbo.TblStore.code"
StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID where 1=1"
StrSQL = StrSQL & " AND  (dbo.Transactions.Transaction_Date >=   " & SQLDate(Fromdate, True) & "    AND   dbo.Transactions.Transaction_Date<=" & SQLDate(ToDate, True) & " )"
StrSQL = StrSQL & "   AND  (dbo.Transactions.Transaction_Type=21)  "
   
StrSQL = StrSQL & "   GROUP BY dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code"
StrSQL = StrSQL & "   Order BY Total  Desc"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
       ' Load FrmBoxesAccounts

        With FGSales
        
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i

                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
             Else
             .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
             End If
                .TextMatrix(i, .ColIndex("StoreCode")) = IIf(IsNull(rs("Code").value), "", rs("Code").value)
      .TextMatrix(i, .ColIndex("Results")) = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
            totaldouble = totaldouble + val(.TextMatrix(i, .ColIndex("Results")))

       
                                If i = 1 Then
                          .TextMatrix(i, .ColIndex("Remarks")) = "«ﬂÀ—  „»Ì⁄« "
                          Maxsale = val(.TextMatrix(i, .ColIndex("Results")))
                             Maxsalename = (.TextMatrix(i, .ColIndex("StoreName")))
                             
                 ElseIf i = rs.RecordCount Then
                 .TextMatrix(i, .ColIndex("Remarks")) = "«ﬁ·  „»Ì⁄« "
                 Minsale = val(.TextMatrix(i, .ColIndex("Results")))
                 minsalename = (.TextMatrix(i, .ColIndex("StoreName")))
                End If
                
                
              
                ChartSales.ShowLegend = True
                ChartSales.ColumnCount = rs.RecordCount
                ChartSales.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Bankes.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Bankes.RowLabel = "Chart"
                End If
      
                ChartSales.Column = i
                ChartSales.Row = 1
                ChartSales.Data = val(.TextMatrix(i, .ColIndex("Results")))
                ChartSales.ColumnLabel = (.TextMatrix(i, .ColIndex("StoreName")))  'IIf(IsNull(rs.Fields("Results").value), 0, rs.Fields("Results").value)
                
            
            
            
            
            
                ChartSales1.ShowLegend = True
                ChartSales1.ColumnCount = rs.RecordCount
                ChartSales1.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Bankes.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Bankes.RowLabel = "Chart"
                End If
      
                ChartSales1.Column = i
                ChartSales1.Row = 1
                ChartSales1.Data = val(.TextMatrix(i, .ColIndex("Results")))
                ChartSales1.ColumnLabel = (.TextMatrix(i, .ColIndex("StoreName")))  'IIf(IsNull(rs.Fields("Results").value), 0, rs.Fields("Results").value)
                            
                            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
      TxtData(0).Text = totaldouble
      TxtData(1).Text = Round(totaldouble / rs.RecordCount, 2)
      TxtData(2).Text = Maxsale
      TxtData(3).Text = Minsale
       TxtData(4).Text = Maxsalename
        TxtData(5).Text = minsalename
        
      Label40.Caption = Now
      
        End With

    End If
Coloring

    Exit Sub

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where TblBoxesData.BoxID <>1"
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblBoxesData.BoxID,dbo.TblBoxesData.BoxName, QryBoxesCredit.BoxCredit" & " FROM dbo.TblBoxesData INNER JOIN " & "dbo.QryBoxesCredit() QryBoxesCredit ON dbo.TblBoxesData.BoxID = QryBoxesCredit.BoxID"

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <>1"
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
      '  Load FrmBoxesAccounts

        With frm_Boxes
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)

                If Not IsNull(rs("BoxCredit").value) Then
                    .TextMatrix(i, .ColIndex("BoxCredit")) = Format(rs("BoxCredit").value, SystemOptions.SysDefCurrencyForamt)
                Else
                    .TextMatrix(i, .ColIndex("BoxCredit")) = 0
                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Msg = "·«ÌÊÃœ «Ï Œ“‰ „”Ã·… ›Ï «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«·«Ì„ﬂ‰ ⁄—÷ «·Œ“‰ «·Õ«·Ì… ›Ï «·»—‰«„Ã...!!!"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

End Sub

Public Sub Show_BankesData()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim FirstPeriod As Date
    Dim Balance As Double
    
    StrSQL = "SELECT * from BanksData  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
       ' Load FrmBoxesAccounts

        With fg_Bank
        
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(rs("BankID").value), "", rs("BankID").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
             Else
             .TextMatrix(i, .ColIndex("BankNameE")) = IIf(IsNull(rs("BankNamee").value), "", rs("BankNamee").value)
             End If
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
      
                getFirstPeriodDateInthisYear FirstPeriod
 
                Balance = GetActualAccountBalance(rs("Account_Code").value, 0, FirstPeriod, Date)
                          
               .TextMatrix(i, .ColIndex("BankCredit")) = Abs(Balance) 'GetActualAccountBalance(rs("Account_Code").value, branch_id, FirstPeriod, Date)


               .TextMatrix(i, .ColIndex("credit")) = Math.Abs(Credit_Debit(.TextMatrix(i, .ColIndex("AccountCode")), 0))
               .TextMatrix(i, .ColIndex("debit")) = Math.Abs(Credit_Debit(.TextMatrix(i, .ColIndex("AccountCode")), 1))
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "„œÌ‰"
                        .TextMatrix(i, .ColIndex("BankCredit")) = Abs(Balance)
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "œ«∆‰"
                        .TextMatrix(i, .ColIndex("BankCredit")) = "( " & Abs(Balance) & " ) "
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                Else

                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Debit"
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Credit"
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                End If
                
                
                
              
                chrt_Bankes.ShowLegend = True
                chrt_Bankes.ColumnCount = rs.RecordCount
                chrt_Bankes.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Bankes.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Bankes.RowLabel = "Chart"
                End If
      
                chrt_Bankes.Column = i
                chrt_Bankes.Row = 1
                chrt_Bankes.Data = val(.TextMatrix(i, .ColIndex("BankCredit")))
                chrt_Bankes.ColumnLabel = IIf(IsNull(rs.Fields("BankName").value), "", rs.Fields("BankName").value)
                
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    Exit Sub

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where TblBoxesData.BoxID <>1"
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblBoxesData.BoxID,dbo.TblBoxesData.BoxName, QryBoxesCredit.BoxCredit" & " FROM dbo.TblBoxesData INNER JOIN " & "dbo.QryBoxesCredit() QryBoxesCredit ON dbo.TblBoxesData.BoxID = QryBoxesCredit.BoxID"

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <>1"
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
      '  Load FrmBoxesAccounts

        With frm_Boxes
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)

                If Not IsNull(rs("BoxCredit").value) Then
                    .TextMatrix(i, .ColIndex("BoxCredit")) = Format(rs("BoxCredit").value, SystemOptions.SysDefCurrencyForamt)
                Else
                    .TextMatrix(i, .ColIndex("BoxCredit")) = 0
                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Msg = "·«ÌÊÃœ «Ï Œ“‰ „”Ã·… ›Ï «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«·«Ì„ﬂ‰ ⁄—÷ «·Œ“‰ «·Õ«·Ì… ›Ï «·»—‰«„Ã...!!!"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

End Sub

Private Sub GetCharge()
        FillGrid
End Sub


Public Sub FillGrid(Optional str As String)

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 
 
My_SQL = My_SQL & " select * , (totalShippedQty - showqty)  _diff from ( "
My_SQL = My_SQL & "  SELECT TOP (100) PERCENT dbo.GetNoOfShipments(dbo.Transactions.Transaction_ID) AS noofShipments, dbo.Transactions.Transaction_ID,"
My_SQL = My_SQL & "  dbo.gettotalShippedQty1(dbo.Transactions.Transaction_ID) AS totalShippedQty, dbo.GetminTimeForShipments(dbo.Transactions.Transaction_ID) AS MinTime,"
My_SQL = My_SQL & "  dbo.GetmaxTimeForShipments(dbo.Transactions.Transaction_ID) AS MaxTime, dbo.Transactions.Without, dbo.Transactions.Wait, dbo.Transactions.FixesAssetsID,"
My_SQL = My_SQL & "  dbo.TblEquipments.Code, dbo.TblEquipments.name, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblProductionType.name AS productiontypename,"
My_SQL = My_SQL & "  dbo.TblProductionType.namee AS productiontypenamee, dbo.Transactions.ContactTime, dbo.Transaction_Details.ShowQty, dbo.TblUnites.UnitName,"
My_SQL = My_SQL & "  dbo.TblUnites.UnitNamee, dbo.Transactions.TransactionComment, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
My_SQL = My_SQL & "  dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Remark, dbo.Transactions.Transaction_Date,"
My_SQL = My_SQL & "  dbo.TblCountriesGovernmentsCities.CityName AS HeyName, dbo.TblCustemers.Account_Code, ISNULL(dbo.ACCOUNTS.DepitBalance, 0) AS DepitBalance,"
My_SQL = My_SQL & "  ISNULL(dbo.ACCOUNTS.CreditBalance, 0) AS CreditBalance, ISNULL(dbo.ACCOUNTS.opening_balance, 0) AS opening_balance, dbo.TblEquipments.Notes,"
My_SQL = My_SQL & "  dbo.TblItems.Wight, dbo.TblItems.[Content], dbo.TblItems.Dippre, dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.ItemName,"
My_SQL = My_SQL & "  TblEmployee_1.Emp_Name AS Workername, TblEmployee_2.Emp_Name AS Helpername, TblEmployee_1.Emp_Namee AS Workernamee,"
My_SQL = My_SQL & "  TblEmployee_2.Emp_Namee AS Helpernamee"
My_SQL = My_SQL & "  FROM     dbo.Transactions INNER JOIN"
My_SQL = My_SQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
My_SQL = My_SQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblEmployee AS TblEmployee_2 ON dbo.Transactions.empID2 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblEmployee AS TblEmployee_1 ON dbo.Transactions.empID1 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
My_SQL = My_SQL & "  dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid"
My_SQL = My_SQL & "  Where (dbo.Transactions.Transaction_Type = 61)"

 My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date >= '" & SQLDate(dtpFromDate.value) & "'"
 My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date <= '" & SQLDate(dtpToDate.value) & "'"

    
 
 My_SQL = My_SQL + "   ) as tb1    order by   _diff  desc "

   
Dim ActualTotal As Double
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i

              .TextMatrix(i, .ColIndex("ShipedQty")) = (IIf(IsNull(rs.Fields("totalShippedQty").value), 0, rs.Fields("totalShippedQty").value))
              .TextMatrix(i, .ColIndex("showqty")) = IIf(IsNull(rs.Fields("showqty").value), 0, rs.Fields("showqty").value)
              .TextMatrix(i, .ColIndex("diff")) = val(.TextMatrix(i, .ColIndex("ShipedQty"))) - val(.TextMatrix(i, .ColIndex("showqty")))
              
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("Notes")) = (IIf(IsNull(rs.Fields("Notes").value), "", rs.Fields("Notes").value))
              
              If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              Else
                        .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
              End If
              
              
              .TextMatrix(i, .ColIndex("ShowQty")) = (IIf(IsNull(rs.Fields("ShowQty").value), "", rs.Fields("ShowQty").value))
              
              .TextMatrix(i, .ColIndex("HeyName")) = (IIf(IsNull(rs.Fields("HeyName").value), "", rs.Fields("HeyName").value))
              .TextMatrix(i, .ColIndex("productiontypename")) = (IIf(IsNull(rs.Fields("productiontypename").value), "", rs.Fields("productiontypename").value))
              .TextMatrix(i, .ColIndex("Contacttime")) = (IIf(IsNull(rs.Fields("Contacttime").value), "", rs.Fields("Contacttime").value))
              
              .TextMatrix(i, .ColIndex("ShowQty")) = (IIf(IsNull(rs.Fields("ShowQty").value), "", rs.Fields("ShowQty").value))
              .TextMatrix(i, .ColIndex("typenew")) = (IIf(IsNull(rs.Fields("typenew").value), "", rs.Fields("typenew").value))
              .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value))
                      
               .TextMatrix(i, .ColIndex("cus_mobile")) = (IIf(IsNull(rs.Fields("cus_mobile").value), "", rs.Fields("cus_mobile").value))
               .TextMatrix(i, .ColIndex("transactioncomment")) = (IIf(IsNull(rs.Fields("transactioncomment").value), "", rs.Fields("transactioncomment").value))
             
                chrt_Charge.ShowLegend = True
                chrt_Charge.ColumnCount = rs.RecordCount
                chrt_Charge.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Charge.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Charge.RowLabel = "Chart"
                End If
      
                chrt_Charge.Column = i
                chrt_Charge.Row = 1
                chrt_Charge.Data = val(.TextMatrix(i, .ColIndex("diff")))
                chrt_Charge.ColumnLabel = .TextMatrix(i, .ColIndex("CusName"))
              
             
             rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub


Private Sub GetProduct()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
    My_SQL = My_SQL & "   SELECT            dbo.GetNoOfShipments(dbo.Transactions.Transaction_ID) AS noofShipments, dbo.Transactions.Transaction_ID,"
    My_SQL = My_SQL & "                   dbo.gettotalShippedQty1(dbo.Transactions.Transaction_ID) AS totalShippedQty, dbo.GetminTimeForShipments(dbo.Transactions.Transaction_ID) AS MinTime,"
    My_SQL = My_SQL & "                   dbo.GetmaxTimeForShipments(dbo.Transactions.Transaction_ID) AS MaxTime, dbo.Transactions.Without, dbo.Transactions.Wait, dbo.Transactions.FixesAssetsID,"
    My_SQL = My_SQL & "                  dbo.TblEquipments.Code, dbo.TblEquipments.name, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblProductionType.name AS productiontypename,"
    My_SQL = My_SQL & "                  dbo.TblProductionType.namee AS productiontypenamee, dbo.Transactions.ContactTime, dbo.Transaction_Details.ShowQty, dbo.TblUnites.UnitName,"
    My_SQL = My_SQL & "                  dbo.TblUnites.UnitNamee, dbo.Transactions.TransactionComment, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
    My_SQL = My_SQL & "                  dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Remark, dbo.Transactions.Transaction_Date,"
    My_SQL = My_SQL & "                  dbo.TblCountriesGovernmentsCities.CityName AS HeyName, dbo.TblCustemers.Account_Code, ISNULL(dbo.ACCOUNTS.DepitBalance, 0) AS DepitBalance,"
    My_SQL = My_SQL & "                   ISNULL(dbo.ACCOUNTS.CreditBalance, 0) AS CreditBalance, ISNULL(dbo.ACCOUNTS.opening_balance, 0) AS opening_balance, dbo.TblEquipments.Notes,"
    My_SQL = My_SQL & "                 dbo.TblItems.Wight, dbo.TblItems.[Content], dbo.TblItems.Dippre, dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.ItemName,"
    My_SQL = My_SQL & "                  TblEmployee_1.Emp_Name AS Workername, TblEmployee_2.Emp_Name AS Helpername, TblEmployee_1.Emp_Namee AS Workernamee,"
    My_SQL = My_SQL & "                 TblEmployee_2.Emp_Namee AS Helpernamee"
    My_SQL = My_SQL & "  FROM     dbo.Transactions INNER JOIN"
    My_SQL = My_SQL & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    My_SQL = My_SQL & "                 dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    My_SQL = My_SQL & "                   dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.TblEmployee AS TblEmployee_2 ON dbo.Transactions.empID2 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.TblEmployee AS TblEmployee_1 ON dbo.Transactions.empID1 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid"
    My_SQL = My_SQL & "  WHERE  (dbo.Transactions.Transaction_Type = 61)"
    My_SQL = My_SQL & "  AND  (dbo.Transactions.Transaction_Date >= '" & SQLDate(dtpFromDate.value) & "')    AND  (dbo.Transactions.Transaction_Date <=  '" & SQLDate(dtpToDate.value) & "') "
    My_SQL = My_SQL & "  ORDER BY dbo.Transactions.Wait, dbo.Transactions.Without"
    


   
Dim ActualTotal As Double

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.fg_Product
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              .TextMatrix(i, .ColIndex("totalShippedQty")) = (IIf(IsNull(rs.Fields("totalShippedQty").value), "", rs.Fields("totalShippedQty").value))
         
                chrt_product.ShowLegend = True
                chrt_product.ColumnCount = rs.RecordCount
                chrt_product.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_product.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_product.RowLabel = "Chart"
                End If
      
                chrt_product.Column = i
                chrt_product.Row = 1
                chrt_product.Data = val(.TextMatrix(i, .ColIndex("totalShippedQty")))
                chrt_product.ColumnLabel = .TextMatrix(i, .ColIndex("CusName"))
              
                          
             rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With


End Sub

Private Sub GetProduct2()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
My_SQL = My_SQL & "  SELECT     dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
My_SQL = My_SQL & "  dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName,"
My_SQL = My_SQL & "  dbo.TblItems.ItemNamee , dbo.TblItems.fullcode           "
My_SQL = My_SQL & "  FROM         dbo.Transactions INNER JOIN                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
My_SQL = My_SQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
My_SQL = My_SQL & "  Where (dbo.Transactions.Transaction_Type = 26)"

 
    My_SQL = My_SQL & "  AND  (dbo.Transactions.Transaction_Date >= '" & SQLDate(dtpFromDate.value) & "')    AND  (dbo.Transactions.Transaction_Date <=  '" & SQLDate(dtpToDate.value) & "') "
    My_SQL = My_SQL & "   order by dbo.Transactions.Transaction_Date  "
    


   
Dim ActualTotal As Double

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.fg_Product
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              
              If SystemOptions.UserInterface Then
                        .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              Else
                            .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
              End If
              
           .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
           .TextMatrix(i, .ColIndex("Transaction_Date")) = (IIf(IsNull(rs.Fields("Transaction_Date").value), "", rs.Fields("Transaction_Date").value))
           .TextMatrix(i, .ColIndex("FullCode")) = (IIf(IsNull(rs.Fields("FullCode").value), "", rs.Fields("FullCode").value))
           .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
              
              .TextMatrix(i, .ColIndex("showqty")) = (IIf(IsNull(rs.Fields("showqty").value), "", rs.Fields("showqty").value))
         
                chrt_product.ShowLegend = True
                chrt_product.ColumnCount = rs.RecordCount
                chrt_product.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_product.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_product.RowLabel = "Chart"
                End If
      
                chrt_product.Column = i
                chrt_product.Row = 1
                chrt_product.Data = val(.TextMatrix(i, .ColIndex("totalShippedQty")))
                chrt_product.ColumnLabel = .TextMatrix(i, .ColIndex("CusName"))
              
                          
             rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With


End Sub



Private Function Credit_Debit(AccountCode As String, Credit_Or_Debit As Integer) As Double

Dim str As String
Dim Result As Double

str = "  select   dbo.GetBalanceCreditORdepit ( '" & SQLDate(dtpFromDate.value) & "' , '" & SQLDate(dtpToDate.value) & "' , '" & AccountCode & "' , " & Credit_Or_Debit & ",1)   as result  "
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
        Result = IIf(IsNull(Rs_Temp("result").value), 0, Rs_Temp("result").value)
End If
Credit_Debit = Result
End Function

Private Sub GetReserve()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
    My_SQL = My_SQL & "  SELECT dbo.TblItems.HaveSerial, dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.ShowQty, dbo.Transactions.Transaction_Date"
    My_SQL = My_SQL & "  FROM     dbo.TblItems INNER JOIN"
    My_SQL = My_SQL & "  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    My_SQL = My_SQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID INNER JOIN"
    My_SQL = My_SQL & "  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"

    My_SQL = My_SQL & "  where dbo.Transactions.Transaction_Date >= '" & SQLDate(dtpFromDate.value) & "'"
    My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date <= '" & SQLDate(dtpToDate.value) & "'"
   
Dim ActualTotal As Double

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.fg_Reserve
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst
i = 1
Dim nn As Long
nn = val(rs.RecordCount + 1 - 1)

            For i = 1 To nn
            
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("ItemCode")) = (IIf(IsNull(rs.Fields("ItemCode").value), "", rs.Fields("ItemCode").value))
              .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
              .TextMatrix(i, .ColIndex("ShowQty")) = (IIf(IsNull(rs.Fields("ShowQty").value), "", rs.Fields("ShowQty").value))
              
                chrt_Reserve.ShowLegend = True
                chrt_Reserve.ColumnCount = rs.RecordCount
                chrt_Reserve.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Reserve.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Reserve.RowLabel = "Chart"
                End If
      
                chrt_Reserve.Column = i
                chrt_Reserve.Row = 1
                chrt_Reserve.Data = val(.TextMatrix(i, .ColIndex("ShowQty")))
                chrt_Reserve.ColumnLabel = .TextMatrix(i, .ColIndex("ItemName"))
              
                          
                rs.MoveNext
           Next
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With


End Sub


Private Sub GetReserve2()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    

 My_SQL = My_SQL & "  SELECT dbo.Transactions.Without, dbo.Transactions.Wait, dbo.Transactions.FixesAssetsID, dbo.TblEquipments.Code, dbo.TblEquipments.name,"
 My_SQL = My_SQL & "  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblProductionType.name AS productiontypename, dbo.TblProductionType.namee AS productiontypenamee,"
 My_SQL = My_SQL & "  dbo.Transactions.ContactTime, dbo.Transaction_Details.ShowQty, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.TransactionComment,"
 My_SQL = My_SQL & "  dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_mobile,"
 My_SQL = My_SQL & "  dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Remark, dbo.Transactions.Transaction_Date, dbo.TblCountriesGovernmentsCities.CityName AS HeyName,"
 My_SQL = My_SQL & "  dbo.TblCustemers.Account_Code, ISNULL(dbo.ACCOUNTS.DepitBalance, 0) AS DepitBalance, ISNULL(dbo.ACCOUNTS.CreditBalance, 0) AS CreditBalance,"
 My_SQL = My_SQL & "  ISNULL(dbo.ACCOUNTS.opening_balance, 0) AS opening_balance, dbo.TblEquipments.Notes, dbo.TblItems.Wight, dbo.TblItems.[Content], dbo.TblItems.Dippre,"
 My_SQL = My_SQL & "  dbo.TblItems.Source , dbo.TblItems.Typenew, dbo.TblItems.itemname"
 My_SQL = My_SQL & "  FROM     dbo.Transactions INNER JOIN"
 My_SQL = My_SQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
 My_SQL = My_SQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
 My_SQL = My_SQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
 My_SQL = My_SQL & "  dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
 My_SQL = My_SQL & "  dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
 My_SQL = My_SQL & "  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 My_SQL = My_SQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 My_SQL = My_SQL & "  dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
 My_SQL = My_SQL & "  dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid"
 My_SQL = My_SQL & "  WHERE  (dbo.Transactions.Transaction_Type = 61) "
 My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date >= '" & SQLDate(dtpFromDate.value) & "'"
 My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date <= '" & SQLDate(dtpToDate.value) & "'"
 My_SQL = My_SQL & "  ORDER BY dbo.Transactions.Wait, dbo.Transactions.Without"

   
Dim ActualTotal As Double

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.fg_Reserve
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst
i = 1
Dim nn As Long, DepitBalance As Double, CreditBalance As Double, finaldebit As Double
Dim finalcredit   As String, opening_balance As String

nn = val(rs.RecordCount + 1 - 1)

            For i = 1 To nn
            
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("Notes")) = (IIf(IsNull(rs.Fields("Notes").value), "", rs.Fields("Notes").value))
              
            If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              Else
               .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
              End If
              
              
              .TextMatrix(i, .ColIndex("ShowQty")) = (IIf(IsNull(rs.Fields("ShowQty").value), "", rs.Fields("ShowQty").value))
              
              .TextMatrix(i, .ColIndex("HeyName")) = (IIf(IsNull(rs.Fields("HeyName").value), "", rs.Fields("HeyName").value))
              .TextMatrix(i, .ColIndex("productiontypename")) = (IIf(IsNull(rs.Fields("productiontypename").value), "", rs.Fields("productiontypename").value))
              .TextMatrix(i, .ColIndex("Contacttime")) = (IIf(IsNull(rs.Fields("Contacttime").value), "", rs.Fields("Contacttime").value))
              
              .TextMatrix(i, .ColIndex("ShowQty")) = (IIf(IsNull(rs.Fields("ShowQty").value), "", rs.Fields("ShowQty").value))
              .TextMatrix(i, .ColIndex("typenew")) = (IIf(IsNull(rs.Fields("typenew").value), "", rs.Fields("typenew").value))
              .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value))
              
              
              DepitBalance = IIf(IsNull(rs.Fields("DepitBalance").value), 0, rs.Fields("DepitBalance").value)
              CreditBalance = IIf(IsNull(rs.Fields("CreditBalance").value), 0, rs.Fields("CreditBalance").value)
              opening_balance = IIf(IsNull(rs.Fields("opening_balance").value), 0, rs.Fields("opening_balance").value)
              
                      
               .TextMatrix(i, .ColIndex("cus_mobile")) = (IIf(IsNull(rs.Fields("cus_mobile").value), "", rs.Fields("cus_mobile").value))
               .TextMatrix(i, .ColIndex("transactioncomment")) = (IIf(IsNull(rs.Fields("transactioncomment").value), "", rs.Fields("transactioncomment").value))
              
                chrt_Reserve.ShowLegend = True
                chrt_Reserve.ColumnCount = rs.RecordCount
                chrt_Reserve.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Reserve.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Reserve.RowLabel = "Chart"
                End If
      
                chrt_Reserve.Column = i
                chrt_Reserve.Row = 1
                chrt_Reserve.Data = val(.TextMatrix(i, .ColIndex("ShowQty")))
                                
                If SystemOptions.UserInterface = ArabicInterface Then
                         chrt_Reserve.ColumnLabel = .TextMatrix(i, .ColIndex("CusName"))
              Else
                         chrt_Reserve.ColumnLabel = .TextMatrix(i, .ColIndex("CusNamee"))
              End If
                          
                rs.MoveNext
           Next
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With


End Sub



Private Sub GetMaterial()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    

My_SQL = My_SQL & "  select distinct   dbo.TblItems.ItemName , dbo.TblItems.ItemNamee ,   ( SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) as SumQty"
My_SQL = My_SQL & "  FROM         dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
My_SQL = My_SQL & "  INNER JOIN  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
My_SQL = My_SQL & "  WHERE     (dbo.TransactionTypes.StockEffect <> 0) AND (dbo.Transaction_Details.Item_ID = TblItems.ItemID  )     ) qty"   ' AND (dbo.Transactions.Transaction_Date <= '" & SQLDate(dtpToDate.value) & "'
My_SQL = My_SQL & "  From dbo.TblItems, Transaction_Details, Transactions, Groups"
My_SQL = My_SQL & "  where  dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID  and"
My_SQL = My_SQL & "  Transaction_Details.Transaction_ID = Transactions.Transaction_ID And dbo.TblItems.GroupID = dbo.Groups.GroupID And Groups.ISMaterial = 1"
   
    Dim ActualTotal As Double

    rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      
      With Me.fg_Material
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           fg_MaterialTotal.Rows = rs.RecordCount + 1
            rs.MoveFirst
            i = 1
            Dim nn As Long
            nn = val(rs.RecordCount + 1 - 1)

            For i = 1 To nn
            
              .TextMatrix(i, .ColIndex("Ser")) = i
                If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
                   fg_MaterialTotal.TextMatrix(i, fg_MaterialTotal.ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
                Else
                         .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value))
                        fg_MaterialTotal.TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value))
                End If
               .TextMatrix(i, .ColIndex("qty")) = (IIf(IsNull(rs.Fields("qty").value), "", rs.Fields("qty").value))
              fg_MaterialTotal.TextMatrix(i, .ColIndex("qty")) = (IIf(IsNull(rs.Fields("qty").value), "", rs.Fields("qty").value))
              
              
                chrt_Material.ShowLegend = True
                chrt_Material.ColumnCount = rs.RecordCount
                chrt_Material.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Material.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Material.RowLabel = "Chart"
                End If
      
                chrt_Material.Column = i
                chrt_Material.Row = 1
                chrt_Material.Data = val(.TextMatrix(i, .ColIndex("qty")))
                
                 If SystemOptions.UserInterface = ArabicInterface Then
                chrt_Material.ColumnLabel = .TextMatrix(i, .ColIndex("ItemName"))
              Else
                     chrt_Material.ColumnLabel = .TextMatrix(i, .ColIndex("ItemName"))
              End If
                                        
                rs.MoveNext
           Next
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub

Private Sub GetMaterialxxxx()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    

My_SQL = My_SQL & "  select distinct   dbo.TblItems.ItemName , dbo.TblItems.ItemNamee ,   ( SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) as SumQty"
My_SQL = My_SQL & "  FROM         dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
My_SQL = My_SQL & "  INNER JOIN  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
My_SQL = My_SQL & "  WHERE     (dbo.TransactionTypes.StockEffect <> 0) AND (dbo.Transaction_Details.Item_ID = TblItems.ItemID  )   )  ) qty"   ' AND (dbo.Transactions.Transaction_Date <= '" & SQLDate(dtpToDate.value) & "'
My_SQL = My_SQL & "  From dbo.TblItems, Transaction_Details, Transactions, Groups"
My_SQL = My_SQL & "  where  dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID  and"
My_SQL = My_SQL & "  Transaction_Details.Transaction_ID = Transactions.Transaction_ID And dbo.TblItems.GroupID = dbo.Groups.GroupID And Groups.ISMaterial = 1"
   
    Dim ActualTotal As Double

    rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      
      With Me.fg_Material
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           fg_MaterialTotal.Rows = rs.RecordCount + 1
            rs.MoveFirst
            i = 1
            Dim nn As Long
            nn = val(rs.RecordCount + 1 - 1)

            For i = 1 To nn
            
              .TextMatrix(i, .ColIndex("Ser")) = i
                If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
                   fg_MaterialTotal.TextMatrix(i, fg_MaterialTotal.ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
                Else
                         .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value))
                        fg_MaterialTotal.TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value))
                End If
               .TextMatrix(i, .ColIndex("qty")) = (IIf(IsNull(rs.Fields("qty").value), "", rs.Fields("qty").value))
              fg_MaterialTotal.TextMatrix(i, .ColIndex("qty")) = (IIf(IsNull(rs.Fields("qty").value), "", rs.Fields("qty").value))
              
              
                chrt_Material.ShowLegend = True
                chrt_Material.ColumnCount = rs.RecordCount
                chrt_Material.RowCount = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                        chrt_Material.RowLabel = "«·—”„ «·»Ì«‰Ì"
                Else
                        chrt_Material.RowLabel = "Chart"
                End If
      
                chrt_Material.Column = i
                chrt_Material.Row = 1
                chrt_Material.Data = val(.TextMatrix(i, .ColIndex("qty")))
                
                 If SystemOptions.UserInterface = ArabicInterface Then
                chrt_Material.ColumnLabel = .TextMatrix(i, .ColIndex("ItemName"))
              Else
                     chrt_Material.ColumnLabel = .TextMatrix(i, .ColIndex("ItemName"))
              End If
                                        
                rs.MoveNext
           Next
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub

Private Sub tmrStoping_Timer()

If Move_y <> 0 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If

tmrScrolling.Enabled = True
tmrStoping.Enabled = False

End Sub


Private Sub All_Total()

        Dim Account_Code_dynamic As String, balanceString As String, Balance As String
        
        Account_Code_dynamic = get_account_code_branch(20, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtBank.Text = balanceString
        
         Account_Code_dynamic = get_account_code_branch(6, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtBox.Text = balanceString
        
       ' Account_Code_dynamic = get_account_code_branch(34, my_branch)
       ' WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtRevenue.Text = RevenueTotals
        
        Account_Code_dynamic = get_account_code_branch(20, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtBank.Text = balanceString
        
          Account_Code_dynamic = get_account_code_branch(0, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtInv.Text = balanceString
        
      '  Account_Code_dynamic = get_account_code_branch(33, my_branch)
      '  WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtExpenses.Text = ExpensesTotals
        
       TxtPayed.Text = PayedTotals
        
        
        Get_Product_Total
        Get_Reserved_Total
       Get_Charge_Total
       Get_Product_Yester_Total
       Get_Sales_Total
       Get_Charge_grid
        
End Sub
Private Function BuildSql3(NoteType As Integer) As String
    Dim StrSQL As String
    Dim Begin As Boolean
    'On Error GoTo ErrTrap

    StrSQL = "SELECT    dbo.Notes.BankName as BankName2 , dbo.Notes.Remark,  dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.BankID, dbo.Notes.ChqueNum, "
    StrSQL = StrSQL & " dbo.Notes.DueDate, dbo.Notes.ExpensesID, dbo.Notes.Note_Value, dbo.BanksData.BankName, dbo.BanksData.BankNamee, dbo.TblBoxesData.BoxName,"
    StrSQL = StrSQL & " dbo.TblBoxesData.BoxNameE, dbo.Notes.CusID, dbo.Notes.BoxID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.RevenuesID,"
    StrSQL = StrSQL & " dbo.TblRevenuesTypes.RevenuesName, dbo.TblRevenuesTypes.RevenuesNamee, dbo.Notes.project_id, dbo.projects.Project_name, dbo.Notes.EmpAccountCode,"
    StrSQL = StrSQL & "  dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.Notes.AccountsCode, dbo.Notes.person, dbo.TblNotesTypes.NotesTypeName,"
    StrSQL = StrSQL & " dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    StrSQL = StrSQL & "  dbo.TblBranchesData.ActivityTypeId, dbo.tblActivitesType.Name AS ACtivityName, dbo.tblActivitesType.namee AS ACtivityNameE, dbo.Notes.NoteCashingType,"
    StrSQL = StrSQL & "  dbo.Notes.Emp_id , dbo.Notes.Doctype"
    StrSQL = StrSQL & " FROM         dbo.tblActivitesType RIGHT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblActivitesType.id = dbo.TblBranchesData.branch_id RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.Notes ON dbo.TblBranchesData.branch_id = dbo.Notes.branch_no LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.ACCOUNTS ON dbo.Notes.AccountsCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.projects ON dbo.Notes.project_id = dbo.projects.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblRevenuesTypes ON dbo.Notes.RevenuesID = dbo.TblRevenuesTypes.RevenuesID LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID"
    StrSQL = StrSQL & "  WHERE     (dbo.Notes.ExpensesID IS NULL)"
 
  StrSQL = StrSQL + " and dbo.Notes.NoteType  =" & NoteType & " "
 StrSQL = StrSQL + " and dbo.Notes.NoteDate =" & SQLDate(Date, True) & ""

 
 
     BuildSql3 = StrSQL
    Exit Function
 
End Function

Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhFirstDayInMonth = DateSerial(year(dtmDate), _
     Month(dtmDate), 1)
End Function

Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInMonth = DateSerial(year(dtmDate), _
     Month(dtmDate) + 1, 0)
End Function

Private Sub Get_Product_Total()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
    On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
   Dim EndFirstPer As String, BeginLastPer As String
    
    My_SQL = My_SQL & "  SELECT     Sum (ShowQty) total"
    My_SQL = My_SQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    My_SQL = My_SQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    My_SQL = My_SQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    My_SQL = My_SQL & "  Where (dbo.Transactions.Transaction_Type = 26)"
        
    My_SQL = My_SQL & "  AND  (dbo.Transactions.Transaction_Date >= '" & SQLDate(dhFirstDayInMonth(DateTime.Now)) & "')    AND  (dbo.Transactions.Transaction_Date <=  '" & SQLDate(dhLastDayInMonth(DateTime.Now)) & "') "
  '  My_SQL = My_SQL & "   order by dbo.Transactions.Transaction_Date  "

EndFirstPer = dhLastDayInMonth(DateTime.Now)

BeginLastPer = dhFirstDayInMonth(DateTime.Now)



    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtProduct_Total.Text = IIf(IsNull(rs("total").value), 0, rs("total").value)
    End If
End Sub

Private Sub Get_Product_Yester_Total()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
    On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
    My_SQL = My_SQL & "  SELECT     Sum (ShowQty) total"
    My_SQL = My_SQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    My_SQL = My_SQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    My_SQL = My_SQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    My_SQL = My_SQL & "  Where (dbo.Transactions.Transaction_Type = 26)"
        
       Dim ss As String
    VBA.Calendar = vbCalGreg
          
          
          
       ss = SQLDate(DateAdd("D", -1, DateTime.Now))
        
        
    My_SQL = My_SQL & "  AND  (dbo.Transactions.Transaction_Date >= '" & SQLDate(DateAdd("D", -1, DateTime.Now)) & "'   AND  (dbo.Transactions.Transaction_Date <=  '" & SQLDate(DateAdd("D", -1, DateTime.Now)) & "') )"


    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtYes.Text = IIf(IsNull(rs("total").value), 0, rs("total").value)
    End If
End Sub



Private Sub Get_Reserved_Total()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
    On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
       

My_SQL = My_SQL & "  SELECT sum( showqty) total"
My_SQL = My_SQL & "    FROM     dbo.Transactions INNER JOIN"
My_SQL = My_SQL & "    dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
My_SQL = My_SQL & "    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
My_SQL = My_SQL & "    dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
My_SQL = My_SQL & "    dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
My_SQL = My_SQL & "    dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
My_SQL = My_SQL & "    dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
My_SQL = My_SQL & "    dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid"
My_SQL = My_SQL & "    Where (dbo.Transactions.Transaction_Type = 61)"

        
My_SQL = My_SQL & "  AND  (dbo.Transactions.Transaction_Date >= '" & SQLDate(DateTime.Now) & "')    AND  (dbo.Transactions.Transaction_Date <=  '" & SQLDate(DateTime.Now) & "') "

   ' My_SQL = My_SQL & "   order by dbo.Transactions.Transaction_Date  "

    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtReserve.Text = IIf(IsNull(rs("total").value), 0, rs("total").value)
    End If
End Sub


Private Sub Get_Charge_Total()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
    On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
       
My_SQL = "SELECT     SUM(dbo.Transaction_Details.ShowQty) AS TotalsSend ,sum(RecivedShippedQty) as totalRecivedShippedQty"
My_SQL = My_SQL & " FROM         dbo.Transactions INNER JOIN"
My_SQL = My_SQL & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
My_SQL = My_SQL & "  GROUP BY dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Date"
My_SQL = My_SQL & "  HAVING      (dbo.Transactions.Transaction_Type = 55) AND (dbo.Transactions.Transaction_Date = " & SQLDate(Date, True) & " )"
 
    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtCharge.Text = IIf(IsNull(rs("TotalsSend").value), 0, rs("TotalsSend").value)
              TxttotalRecivedShippedQty.Text = IIf(IsNull(rs("totalRecivedShippedQty").value), 0, rs("totalRecivedShippedQty").value)
              
    End If
    
End Sub



Public Sub Get_Charge_grid(Optional str As String)

Dim My_SQL As String

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 
 
 
My_SQL = "SELECT       dbo.gettotalShippedQty1(Transactions.ProkerId) AS totaRQty, Transactions.Transaction_ID, dbo.TblCustemers.CusName, "
My_SQL = My_SQL & "                       dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, Transactions_1.ProkerId, Transactions_1.Transaction_Type, dbo.Transaction_Details.ShipedQty,"
My_SQL = My_SQL & "                       dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.RecivedShippedQty, dbo.Transaction_Details.TotallRecivedShippedQty, Transactions_1.ProductionTypeid,"
My_SQL = My_SQL & "                       dbo.TblProductionType.name AS productionname, dbo.TblProductionType.namee AS productionnamee, Transactions_1.Transaction_Date,"
My_SQL = My_SQL & "                       Transactions_1.FixesAssetsID, Transactions_1.TransactionComment AS additions, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
My_SQL = My_SQL & "                       dbo.TblItems.Fullcode AS ItemCode, dbo.TblItems.Dippre, dbo.TblItems.Content, dbo.TblItems.Wight, Transactions_1.Neighborhoodid,"
My_SQL = My_SQL & "                       dbo.TblCountriesGovernmentsCities.CityName AS location, Transactions_1.CarId, dbo.TblCarsData.BoardNO, Transactions_1.empID1, TblEmployee_2.Emp_Name,"
My_SQL = My_SQL & "                       TblEmployee_2.Fullcode AS empfullcode, TblEmployee_2.Emp_Namee, Transactions_1.empID2, TblEmployee_2.Emp_Name AS supervisorname,"
My_SQL = My_SQL & "                       TblEmployee_2.Fullcode AS supervisorCode, TblEmployee_1.Emp_Code AS driverCode, TblEmployee_1.Emp_Name AS drivername, dbo.TblItems.Source,"
My_SQL = My_SQL & "                       dbo.TblItems.Typenew, TblEmployee_3.Emp_Name AS Workername, TblEmployee_3.Emp_Namee AS WorkernameEng, dbo.TblCarsData.Notes AS carecode2,"
My_SQL = My_SQL & "                       Transactions.NoteSerial1, Transactions.ContactTime, TblEmployee_4.Emp_Name AS sabasupervisor, dbo.FixedAssets.id, dbo.TblEquipments.Notes AS pubpcodenew,"
My_SQL = My_SQL & "                       dbo.Transaction_Details.lotNo , dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.ProductionDate"
My_SQL = My_SQL & "  FROM         dbo.TblProductionType RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblEquipments RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.FixedAssets ON dbo.TblEquipments.fixedAssetid = dbo.FixedAssets.id RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.Transactions Transactions_1 ON TblEmployee_2.Emp_ID = Transactions_1.empID2 ON"
My_SQL = My_SQL & "                       dbo.TblCountriesGovernmentsCities.CityID = Transactions_1.Neighborhoodid ON TblEmployee_3.Emp_ID = Transactions_1.empID1 ON"
My_SQL = My_SQL & "                       dbo.FixedAssets.id = Transactions_1.FixesAssetsID ON dbo.TblProductionType.Id = Transactions_1.ProductionTypeid RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblCarsData RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblEmployee TblEmployee_4 RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.Transactions Transactions INNER JOIN"
My_SQL = My_SQL & "                       dbo.Transaction_Details ON Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON TblEmployee_4.Emp_ID = Transactions.empID1 ON"
My_SQL = My_SQL & "                       dbo.TblCarsData.id = Transactions.CarId ON TblEmployee_1.Emp_ID = Transactions.DriverId ON"
My_SQL = My_SQL & "                       Transactions_1.Transaction_ID = Transactions.ProkerId LEFT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblCustemers ON Transactions.CusID = dbo.TblCustemers.CusID"
My_SQL = My_SQL & " WHERE     (Transactions_1.Transaction_Date =" & SQLDate(DateTime.Now, True) & ""
My_SQL = My_SQL & "  )ORDER BY dbo.Transaction_Details.ID "



'new
 
 
 My_SQL = "   select sum(ShowQty) as ShowQty ,CusName,CusNamee,productionname,productionnamee"
My_SQL = My_SQL & "   From"
My_SQL = My_SQL & "    ("
My_SQL = My_SQL & "    SELECT       dbo.gettotalShippedQty1(Transactions.ProkerId) AS totaRQty, Transactions.Transaction_ID, dbo.TblCustemers.CusName,                        dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
My_SQL = My_SQL & "  Transactions_1.ProkerId, Transactions_1.Transaction_Type, dbo.Transaction_Details.ShipedQty,"
My_SQL = My_SQL & "           dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.RecivedShippedQty, dbo.Transaction_Details.TotallRecivedShippedQty,"
My_SQL = My_SQL & "   Transactions_1.ProductionTypeid,                       dbo.TblProductionType.name AS productionname,"
My_SQL = My_SQL & "   dbo.TblProductionType.namee AS productionnamee, Transactions_1.Transaction_Date,                       Transactions_1.FixesAssetsID,"
 My_SQL = My_SQL & "   Transactions_1.TransactionComment AS additions, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,                       dbo.TblItems.Fullcode AS ItemCode,"
My_SQL = My_SQL & "  dbo.TblItems.Dippre, dbo.TblItems.Content, dbo.TblItems.Wight, Transactions_1.Neighborhoodid,                       dbo.TblCountriesGovernmentsCities.CityName AS location, Transactions_1.CarId"
My_SQL = My_SQL & "    , dbo.TblCarsData.BoardNO, Transactions_1.empID1, TblEmployee_2.Emp_Name,                       TblEmployee_2.Fullcode AS empfullcode, TblEmployee_2.Emp_Namee, Transactions_1.empID2, TblEmployee_2.Emp_Name AS supervisorname,"
My_SQL = My_SQL & "          TblEmployee_2.Fullcode AS supervisorCode, TblEmployee_1.Emp_Code AS driverCode, TblEmployee_1.Emp_Name AS drivername,"
My_SQL = My_SQL & "  dbo.TblItems.Source,                       dbo.TblItems.Typenew, TblEmployee_3.Emp_Name AS Workername, TblEmployee_3.Emp_Namee AS WorkernameEng,"
My_SQL = My_SQL & "   dbo.TblCarsData.Notes AS carecode2,                       Transactions.NoteSerial1, Transactions.ContactTime, TblEmployee_4.Emp_Name AS sabasupervisor, dbo.FixedAssets.id, dbo.TblEquipments.Notes AS pubpcodenew,                       dbo.Transaction_Details.lotNo , dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.ProductionDate"
My_SQL = My_SQL & "   , dbo.TblCustemers.CusID"
My_SQL = My_SQL & "     FROM         dbo.TblProductionType RIGHT OUTER JOIN                       dbo.TblEquipments RIGHT OUTER JOIN                       dbo.FixedAssets ON dbo.TblEquipments.fixedAssetid = dbo.FixedAssets.id RIGHT OUTER JOIN                       dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN                       dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
My_SQL = My_SQL & "     dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN                       dbo.Transactions Transactions_1 ON TblEmployee_2.Emp_ID = Transactions_1.empID2 ON                       dbo.TblCountriesGovernmentsCities.CityID = Transactions_1.Neighborhoodid ON TblEmployee_3.Emp_ID = Transactions_1.empID1 ON                       dbo.FixedAssets.id = Transactions_1.FixesAssetsID ON dbo.TblProductionType.Id = Transactions_1.ProductionTypeid RIGHT OUTER JOIN"
My_SQL = My_SQL & "      dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN                       dbo.TblCarsData RIGHT OUTER JOIN                       dbo.TblEmployee TblEmployee_4 RIGHT OUTER JOIN                       dbo.Transactions Transactions INNER JOIN                       dbo.Transaction_Details ON Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON TblEmployee_4.Emp_ID = Transactions.empID1 ON                       dbo.TblCarsData.id = Transactions.CarId ON TblEmployee_1.Emp_ID = Transactions.DriverId ON                       Transactions_1.Transaction_ID = Transactions.ProkerId LEFT OUTER JOIN                       dbo.TblCustemers"
My_SQL = My_SQL & "   ON Transactions.CusID = dbo.TblCustemers.CusID"
My_SQL = My_SQL & " WHERE     (  (Transactions.Transaction_Type = 55) AND Transactions.Transaction_Date =" & SQLDate(DateTime.Now, True) & ")"
My_SQL = My_SQL & "    )x"
My_SQL = My_SQL & "    group by CusID,CusName,CusNamee,productionname,productionnamee"
My_SQL = My_SQL & "   order by CusName"

'  order by CusName
 
    
 'SQLDate(DateTime.Now)

   
Dim ActualTotal As Double
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.fg_Charge_Totals
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
  
              .TextMatrix(i, .ColIndex("total")) = (IIf(IsNull(rs.Fields("ShowQty").value), "", rs.Fields("ShowQty").value))
              If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              .TextMatrix(i, .ColIndex("productiontypename")) = (IIf(IsNull(rs.Fields("productionname").value), "", rs.Fields("productionname").value))
              Else
                 .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
                 .TextMatrix(i, .ColIndex("productiontypenamee")) = (IIf(IsNull(rs.Fields("productionname").value), "", rs.Fields("productionname").value))
            End If
                 
                 
             rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub



Private Sub Get_Sales_Total()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
    On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    
       Dim Fromdate As Date
       Dim ToDate As Date
     Fromdate = "01/" & Month(Date) & "/" & year(Date)
ToDate = MonthLastDay(Date)

My_SQL = My_SQL & "  SELECT  sum ( dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice ) Tot ,( sum (showprice) / count(*) ) Avg_Tot"

My_SQL = My_SQL & "  FROM     dbo.Transaction_Details INNER JOIN"
My_SQL = My_SQL & "                    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"


        
   My_SQL = My_SQL & " AND  (dbo.Transactions.Transaction_Date >=   " & SQLDate(Fromdate, True) & "    AND   dbo.Transactions.Transaction_Date<=" & SQLDate(ToDate, True) & " )"
   My_SQL = My_SQL & "   AND  (dbo.Transactions.Transaction_Type=21)  "

    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
         '     txtAvg.text = Math.Round(IIf(IsNull(rs("Avg_Tot").value), 0, rs("Avg_Tot").value), 2)
              txtSalestotal.Text = Math.Round(IIf(IsNull(rs("Tot").value), 0, rs("Tot").value), 2)
    End If
rs.Close


My_SQL = My_SQL & "  SELECT  sum ( dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice ) Tot ,( sum (showprice) / count(*) ) Avg_Tot"

My_SQL = My_SQL & "  FROM     dbo.Transaction_Details INNER JOIN"
My_SQL = My_SQL & "                    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"


        
   My_SQL = My_SQL & " AND  (dbo.Transactions.Transaction_Date >= (SELECT datesatrt FROM  dbo.TblyearsData WHERE  CurrentYear = 1))    AND  (dbo.Transactions.Transaction_Date <= ((SELECT DateEnd FROM  dbo.TblyearsData WHERE  CurrentYear = 1)))"
   ' My_SQL = My_SQL & "   order by dbo.Transactions.Transaction_Date  "

     
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtAvg.Text = Math.Round(IIf(IsNull(rs("Avg_Tot").value), 0, rs("Avg_Tot").value), 2)
   '           txtSalestotal.text = Math.Round(IIf(IsNull(rs("Tot").value), 0, rs("Tot").value), 2)
    End If
    
End Sub

Private Sub Get_DatePer()

Dim My_SQL As String

  '  On Error GoTo ErrTrap
    On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    
       

My_SQL = My_SQL & "  SELECT  sum ( dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice ) Tot ,( sum (showprice) / count(*) ) Avg_Tot"

My_SQL = My_SQL & "  FROM     dbo.Transaction_Details INNER JOIN"
My_SQL = My_SQL & "                    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"


    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtAvg.Text = Math.Round(IIf(IsNull(rs("Avg_Tot").value), 0, rs("Avg_Tot").value), 2)
              txtSalestotal.Text = Math.Round(IIf(IsNull(rs("Tot").value), 0, rs("Tot").value), 2)
    End If
    
End Sub
Function PrintPayedVchr()
    Dim Msg As String
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim BolShowNotes As Boolean
    On Error GoTo ErrTrap

 
 
 Dim Reports As ClsRepoerts

    Set Reports = New ClsRepoerts
 
    Reports.ShowSallingTime BuildSql3(5), Date, Date, , 9, , , ""
 
     
ErrTrap:


End Function

Function Print_Expenses(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""
    
MySQL = MySQL & "  SELECT     dbo.Notes.NoteDate, dbo.Notes.NoteSerial, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.ACCOUNTS.Account_Name,"
MySQL = MySQL & "      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description"
MySQL = MySQL & "      FROM         dbo.ACCOUNTS INNER JOIN"
MySQL = MySQL & "      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
MySQL = MySQL & "      dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
MySQL = MySQL & "      WHERE     (dbo.ACCOUNTS.Account_Code IN"
MySQL = MySQL & "      (SELECT     Account_Code"
MySQL = MySQL & "      From dbo.ExpensesType"
MySQL = MySQL & "                )"
MySQL = MySQL & "      and   dbo.Notes.NoteDate=" & SQLDate(Date, True) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ExpesesTotal.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ExpesesTotal.rpt"
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
        StrReportTitle = ""
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
 


Private Function RevenueTotals() As Double
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = "SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) As totalRevenue"
MySQL = MySQL & "  FROM         dbo.Notes INNER JOIN"
MySQL = MySQL & "  dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType INNER JOIN"
MySQL = MySQL & "                        dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
MySQL = MySQL & "  WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)  AND (dbo.Notes.NoteDate = " & SQLDate(Date, True) & ") AND"
MySQL = MySQL & "                        (dbo.Notes.NoteType = 4) OR"
MySQL = MySQL & "                        (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1) AND (dbo.Notes.NoteID IN"
MySQL = MySQL & "                            (SELECT     NoteId"
MySQL = MySQL & "                               From dbo.Transactions"
MySQL = MySQL & "                               WHERE     (Transactions.PaymentType = 0 AND Transaction_Type = 21 AND Transaction_Date = " & SQLDate(Date, True) & " ))) AND (dbo.Notes.NoteType = 170)"
                             
 
 
MySQL = " SELECT     SUM(dbo.Notes.Note_Value) AS RevenueTotals"
MySQL = MySQL & "                  FROM         dbo.tblActivitesType RIGHT OUTER JOIN"
MySQL = MySQL & "                                        dbo.TblBranchesData ON dbo.tblActivitesType.id = dbo.TblBranchesData.branch_id RIGHT OUTER JOIN"
MySQL = MySQL & "                                        dbo.Notes ON dbo.TblBranchesData.branch_id = dbo.Notes.branch_no LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.ACCOUNTS ON dbo.Notes.AccountsCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.projects ON dbo.Notes.project_id = dbo.projects.id LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.TblRevenuesTypes ON dbo.Notes.RevenuesID = dbo.TblRevenuesTypes.RevenuesID LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                                        dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID"
MySQL = MySQL & "                  WHERE     (dbo.Notes.NoteType = 4) "
 MySQL = MySQL & "      and   ( NoteDate =" & SQLDate(Date, True) & " )"


'AND (dbo.Notes.NoteDate = CONVERT(DATETIME, '2017-01-08 00:00:00', 102))


    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
 RevenueTotals = 0
        Exit Function
    End If
RevenueTotals = IIf(IsNull(RsData("RevenueTotals").value), 0, RsData("RevenueTotals").value)

  End Function


Private Function Print_Revenue()
     
    Dim Msg As String
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim BolShowNotes As Boolean
    On Error GoTo ErrTrap

 
 
 Dim Reports As ClsRepoerts

    Set Reports = New ClsRepoerts
 
    Reports.ShowSallingTime BuildSql3(4), Date, Date, , 9, , , ""
 
     
ErrTrap:
     
     
Exit Function
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    'Dim Msg As String

MySQL = ""
    
MySQL = MySQL & " SELECT    dbo.Notes.NoteSerial1, dbo.TblNotesTypes.NotesTypeName, dbo.Notes.Remark, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
MySQL = MySQL & " dbo.Notes.NoteDate"
MySQL = MySQL & " FROM         dbo.Notes INNER JOIN"
MySQL = MySQL & " dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType INNER JOIN"
MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
 
MySQL = MySQL & "    WHERE    (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1  AND      dbo.Notes.NoteDate =  " & SQLDate(Date, True) & " and  dbo.Notes.NoteType = 4   AND NoteDate = " & SQLDate(Date, True) & ") "
MySQL = MySQL & "   OR"
  MySQL = MySQL & "  ("
MySQL = MySQL & "    dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1"
MySQL = MySQL & "    and   dbo.Notes.NoteType = 170"
MySQL = MySQL & "     AND  dbo.Notes.NoteID IN"
MySQL = MySQL & "        ("
MySQL = MySQL & "          SELECT     NoteId        From dbo.Transactions"
MySQL = MySQL & "          WHERE   (Transactions.PaymentType = 0 and   Transaction_Type = 21 AND Transaction_Date =   " & SQLDate(Date, True) & ")"
MySQL = MySQL & "  )"
MySQL = MySQL & "  )"
 MySQL = MySQL & "   ORDER BY dbo.Notes.NoteType DESC"


 
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Revenue.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Revenue.rpt"
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
        StrReportTitle = ""
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

Private Function PayedTotals() As Double
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT     SUM(dbo.Notes.Note_Value) AS TotalPayed"
MySQL = MySQL & "     FROM         dbo.tblActivitesType RIGHT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblBranchesData ON dbo.tblActivitesType.id = dbo.TblBranchesData.branch_id RIGHT OUTER JOIN"
MySQL = MySQL & "                          dbo.Notes ON dbo.TblBranchesData.branch_id = dbo.Notes.branch_no LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.ACCOUNTS ON dbo.Notes.AccountsCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.projects ON dbo.Notes.project_id = dbo.projects.id LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblRevenuesTypes ON dbo.Notes.RevenuesID = dbo.TblRevenuesTypes.RevenuesID LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID"
MySQL = MySQL & "    WHERE     (dbo.Notes.NoteType = 5)"
 MySQL = MySQL & "      and   ( NoteDate =" & SQLDate(Date, True) & " )"
                             
     Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
 PayedTotals = 0
        Exit Function
    End If
PayedTotals = IIf(IsNull(RsData("TotalPayed").value), 0, RsData("TotalPayed").value)

  End Function
Private Function ExpensesTotals() As Double
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = " SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS TotalExpenses"
    MySQL = MySQL & "    FROM         dbo.ACCOUNTS INNER JOIN"
    MySQL = MySQL & "                          dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
    MySQL = MySQL & "                          dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
    MySQL = MySQL & "    WHERE     (dbo.ACCOUNTS.Account_Code IN"
    MySQL = MySQL & "                              (SELECT     Account_Code"
    MySQL = MySQL & "                                 FROM         dbo.ExpensesType)) AND (dbo.Notes.NoteDate =" & SQLDate(Date, True) & " )"
                             
 
                             
 

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        ExpensesTotals = 0
        Exit Function
    End If
    ExpensesTotals = IIf(IsNull(RsData("TotalExpenses").value), 0, RsData("TotalExpenses").value)
    End Function
'_____________________________________________________________________________________________________________________________________________________________________________________________________________________________________
'Khaled Part
'#####################################################################################################################################################################################################################################
Private Sub fillgrid1()

    Dim KSQLSrt As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    KSQLSrt = "SELECT Transaction_Details.CostPrice, TblCustemers.Fullcode, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, CONVERT(varchar, MONTH(Transactions.Transaction_Date)) + '-' + CONVERT(varchar,"
    KSQLSrt = KSQLSrt & " YEAR(Transactions.Transaction_Date)) AS Monthdate1, Transaction_Details.ShowQty, Transaction_Details.showPrice, Transactions.Transaction_ID, Transactions.Transaction_Serial, Transactions.Transaction_Date,"
    KSQLSrt = KSQLSrt & " TransactionTypes.TransactionTypeName, dbo.GetItemUnitFactor(Transaction_Details.Item_ID, 0) AS UnitFactor, TransactionTypes.TransactionEnglishName, TransactionTypes.StockEffect, Transactions.Transaction_Type,"
    KSQLSrt = KSQLSrt & " Transactions.PaymentType, Transactions.CusID, Transactions.StoreID, Transactions.UserID, Transactions.Emp_ID, Transaction_Details.Item_ID, Transaction_Details.Quantity, Transaction_Details.Price,"
    KSQLSrt = KSQLSrt & " Transaction_Details.UnitId, Transaction_Details.OpeningBurcahseValue, Transaction_Details.OpeningBurcahseQty, Transaction_Details.OpeningSalesQty, Transaction_Details.OpeningSalesValue, TblUnites.UnitName,"
    KSQLSrt = KSQLSrt & " TblStore.StoreName, TblCustemers.CusName, TblCustemers.CusNamee, Transactions.PayedValue, Transactions.BranchId, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBranchesData.Tel,"
    KSQLSrt = KSQLSrt & " TblItems.ItemCode, TblItems.ItemName, TblItems.GroupID, Groups.GroupName, Groups.GroupNamee, Transaction_Details.OpeningReSalesQty, Transaction_Details.OpeningReSalesValue,"
    KSQLSrt = KSQLSrt & " ISNULL(Transaction_Details.TotalDiscountPerLine, 0) * ISNULL(Transactions.Currency_rate, 1) AS TotalDiscountPerLine, ISNULL(Transactions.Currency_rate, 1) AS Currency_rate,"
    KSQLSrt = KSQLSrt & " CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END AS 'ItemDiscountValue',"
    KSQLSrt = KSQLSrt & " Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) AS localprice, Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) * Transaction_Details.ShowQty AS LineNet,"
    KSQLSrt = KSQLSrt & " Transactions.NoteSerial1, TblItems.ItemNamee, Transactions.CashCustomerName, Transactions.CashCustomerPhone, Transactions.CashCustomerMobile, Transactions.CashCustomerAddress,"
    KSQLSrt = KSQLSrt & " Transactions.CashCustomerComment, Transactions.VAT, Transactions.VATNO, Transactions.Transaction_NetValue, Transaction_Details.Vat AS VatDet, Transaction_Details.Vatyo, Transactions.BasedOn"
    KSQLSrt = KSQLSrt & " FROM Transactions INNER JOIN"
    KSQLSrt = KSQLSrt & " Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblStore ON Transactions.StoreID = TblStore.StoreID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblCustemers ON Transactions.CusID = TblCustemers.CusID INNER JOIN"
    KSQLSrt = KSQLSrt & " TransactionTypes ON Transactions.Transaction_Type = TransactionTypes.Transaction_Type INNER JOIN"
    KSQLSrt = KSQLSrt & " TblItems ON Transaction_Details.Item_ID = TblItems.ItemID INNER JOIN"
    KSQLSrt = KSQLSrt & " Groups ON TblItems.GroupID = Groups.GroupID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblUnites ON Transaction_Details.UnitId = TblUnites.UnitID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblEmployee ON Transactions.Emp_ID = TblEmployee.Emp_ID"
    KSQLSrt = KSQLSrt & " Where (transactions.Transaction_Type = 21)"

    
    If Not IsNull(KFromDate.value) Then
        KSQLSrt = KSQLSrt & " and Transactions.Transaction_Date >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        KSQLSrt = KSQLSrt & "and Transactions.Transaction_Date <=" & SQLDate(KToDate.value, True) & ""
    End If
         KSQLSrt = KSQLSrt & "  order by Transactions.Transaction_Date,NoteSerial1"
    
    Set rs = New ADODB.Recordset
    rs.Open KSQLSrt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    VSFlexGrid7.Rows = VSFlexGrid5.FixedRows
    If rs.BOF Or rs.EOF Then
    Else
        With Me.VSFlexGrid7
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .Rows - 1
'            MsgBox IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value)
'            MsgBox IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)
            
            '
                .TextMatrix(i, .ColIndex("Serial")) = i
              .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                .TextMatrix(i, .ColIndex("InvDate")) = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
                .TextMatrix(i, .ColIndex("InvNo")) = IIf(IsNull(rs("NOTESERIAL1").value), "", rs("NOTESERIAL1").value)
                .TextMatrix(i, .ColIndex("itemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("cont")) = IIf(IsNull(rs("SHOWQTY").value), "", rs("SHOWQTY").value)
                
                
                
                total5 = total5 + IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)
                
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs("SHOWpRICE").value), "", rs("SHOWpRICE").value)
                .TextMatrix(i, .ColIndex("total")) = (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value)) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)
                
                total1 = total1 + (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value)) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)
                
                .TextMatrix(i, .ColIndex("RowValue")) = (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * IIf(IsNull(rs("CostPrice").value), "", rs("CostPrice").value)
                
                total2 = total2 + (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * IIf(IsNull(rs("CostPrice").value), 0, rs("CostPrice").value)
                

                .TextMatrix(i, .ColIndex("profit")) = (((IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value))) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)) - (IIf(IsNull(rs("CostPrice").value), 0, rs("CostPrice").value) * IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value))
                total3 = total3 + (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value)) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value) - (IIf(IsNull(rs("CostPrice").value), 0, rs("CostPrice").value) * IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value))
                If (((IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value))) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)) - (IIf(IsNull(rs("CostPrice").value), 0, rs("CostPrice").value) * IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) < 0 Then
                    .Cell(flexcpBackColor, i, 1, i, 8) = &H8080FF
                End If
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub

Private Sub FillGrid2()
    Dim KSQLSrt As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    KSQLSrt = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.Note_Value, Notes.ExpensesID, ExpensesType.Name,ExpensesType.Namee, Notes.NoteSerial, Notes.Remark, TblUsers.UserName, Notes.BoxID, Notes.UserID, TblBoxesData.BoxName,"
    KSQLSrt = KSQLSrt & " BanksData.BankName , Notes.BankID, Notes.ChqueNum, Notes.dueDate, Notes.NoteSerial1, Notes.branch_no"
    KSQLSrt = KSQLSrt & " FROM TblUsers INNER JOIN"
    KSQLSrt = KSQLSrt & " ExpensesType INNER JOIN"
    KSQLSrt = KSQLSrt & " Notes ON ExpensesType.ID = Notes.ExpensesID ON TblUsers.UserID = Notes.UserID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblBoxesData ON Notes.BoxID = TblBoxesData.BoxID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " BanksData ON Notes.BankID = BanksData.BankID"
    KSQLSrt = KSQLSrt & " Where (Notes.NoteType = 3)"
    
    If Not IsNull(KFromDate.value) Then
        KSQLSrt = KSQLSrt & " and Notes.NoteDate >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        KSQLSrt = KSQLSrt & " and Notes.NoteDate <=" & SQLDate(KToDate.value, True) & ""
    End If
    KSQLSrt = KSQLSrt & "  order by NoteDate,NoteSerial1"
    
    VSFlexGrid8.Rows = VSFlexGrid5.FixedRows
    
    Set rs = New ADODB.Recordset
    rs.Open KSQLSrt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
    Else
        With Me.VSFlexGrid8
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("voucherNo")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("voucherDate")) = IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
                .TextMatrix(i, .ColIndex("voucherValue")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                total4 = total4 + IIf(IsNull(rs("Note_Value").value), 0, rs("Note_Value").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
                End If
                
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Private Sub FillOthers()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim KSQLSrt As String
        
    KSQLSrt = "SELECT Transaction_Details.CostPrice, TblCustemers.Fullcode, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, CONVERT(varchar, MONTH(Transactions.Transaction_Date)) + '-' + CONVERT(varchar,"
    KSQLSrt = KSQLSrt & " YEAR(Transactions.Transaction_Date)) AS Monthdate1, Transaction_Details.ShowQty, Transaction_Details.showPrice, Transactions.Transaction_ID, Transactions.Transaction_Serial, Transactions.Transaction_Date,"
    KSQLSrt = KSQLSrt & " TransactionTypes.TransactionTypeName, dbo.GetItemUnitFactor(Transaction_Details.Item_ID, 0) AS UnitFactor, TransactionTypes.TransactionEnglishName, TransactionTypes.StockEffect, Transactions.Transaction_Type,"
    KSQLSrt = KSQLSrt & " Transactions.PaymentType, Transactions.CusID, Transactions.StoreID, Transactions.UserID, Transactions.Emp_ID, Transaction_Details.Item_ID, Transaction_Details.Quantity, Transaction_Details.Price,"
    KSQLSrt = KSQLSrt & " Transaction_Details.UnitId, Transaction_Details.OpeningBurcahseValue, Transaction_Details.OpeningBurcahseQty, Transaction_Details.OpeningSalesQty, Transaction_Details.OpeningSalesValue, TblUnites.UnitName,"
    KSQLSrt = KSQLSrt & " TblStore.StoreName, TblCustemers.CusName, TblCustemers.CusNamee, Transactions.PayedValue, Transactions.BranchId, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBranchesData.Tel,"
    KSQLSrt = KSQLSrt & " TblItems.ItemCode, TblItems.ItemName, TblItems.GroupID, Groups.GroupName, Groups.GroupNamee, Transaction_Details.OpeningReSalesQty, Transaction_Details.OpeningReSalesValue,"
    KSQLSrt = KSQLSrt & " ISNULL(Transaction_Details.TotalDiscountPerLine, 0) * ISNULL(Transactions.Currency_rate, 1) AS TotalDiscountPerLine, ISNULL(Transactions.Currency_rate, 1) AS Currency_rate,"
    KSQLSrt = KSQLSrt & " CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END AS 'ItemDiscountValue',"
    KSQLSrt = KSQLSrt & " Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) AS localprice, Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) * Transaction_Details.ShowQty AS LineNet,"
    KSQLSrt = KSQLSrt & " Transactions.NoteSerial1, TblItems.ItemNamee, Transactions.CashCustomerName, Transactions.CashCustomerPhone, Transactions.CashCustomerMobile, Transactions.CashCustomerAddress,"
    KSQLSrt = KSQLSrt & " Transactions.CashCustomerComment, Transactions.VAT, Transactions.VATNO, Transactions.Transaction_NetValue, Transaction_Details.Vat AS VatDet, Transaction_Details.Vatyo, Transactions.BasedOn"
    KSQLSrt = KSQLSrt & " FROM Transactions INNER JOIN"
    KSQLSrt = KSQLSrt & " Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblStore ON Transactions.StoreID = TblStore.StoreID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblCustemers ON Transactions.CusID = TblCustemers.CusID INNER JOIN"
    KSQLSrt = KSQLSrt & " TransactionTypes ON Transactions.Transaction_Type = TransactionTypes.Transaction_Type INNER JOIN"
    KSQLSrt = KSQLSrt & " TblItems ON Transaction_Details.Item_ID = TblItems.ItemID INNER JOIN"
    KSQLSrt = KSQLSrt & " Groups ON TblItems.GroupID = Groups.GroupID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblUnites ON Transaction_Details.UnitId = TblUnites.UnitID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblEmployee ON Transactions.Emp_ID = TblEmployee.Emp_ID"
    KSQLSrt = KSQLSrt & " Where (transactions.Transaction_Type = 21)"
    KSQLSrt = KSQLSrt & " and Transactions.PaymentType = 0"
    
    If Not IsNull(KFromDate.value) Then
        KSQLSrt = KSQLSrt & " and Transactions.Transaction_Date >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        KSQLSrt = KSQLSrt & "and Transactions.Transaction_Date <=" & SQLDate(KToDate.value, True) & ""
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open KSQLSrt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    For i = 1 To rs.RecordCount
        Sum1 = Sum1 + (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value)) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)
        rs.MoveNext
    Next i
    
    KSQLSrt = "SELECT Transaction_Details.CostPrice, TblCustemers.Fullcode, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, CONVERT(varchar, MONTH(Transactions.Transaction_Date)) + '-' + CONVERT(varchar,"
    KSQLSrt = KSQLSrt & " YEAR(Transactions.Transaction_Date)) AS Monthdate1, Transaction_Details.ShowQty, Transaction_Details.showPrice, Transactions.Transaction_ID, Transactions.Transaction_Serial, Transactions.Transaction_Date,"
    KSQLSrt = KSQLSrt & " TransactionTypes.TransactionTypeName, dbo.GetItemUnitFactor(Transaction_Details.Item_ID, 0) AS UnitFactor, TransactionTypes.TransactionEnglishName, TransactionTypes.StockEffect, Transactions.Transaction_Type,"
    KSQLSrt = KSQLSrt & " Transactions.PaymentType, Transactions.CusID, Transactions.StoreID, Transactions.UserID, Transactions.Emp_ID, Transaction_Details.Item_ID, Transaction_Details.Quantity, Transaction_Details.Price,"
    KSQLSrt = KSQLSrt & " Transaction_Details.UnitId, Transaction_Details.OpeningBurcahseValue, Transaction_Details.OpeningBurcahseQty, Transaction_Details.OpeningSalesQty, Transaction_Details.OpeningSalesValue, TblUnites.UnitName,"
    KSQLSrt = KSQLSrt & " TblStore.StoreName, TblCustemers.CusName, TblCustemers.CusNamee, Transactions.PayedValue, Transactions.BranchId, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBranchesData.Tel,"
    KSQLSrt = KSQLSrt & " TblItems.ItemCode, TblItems.ItemName, TblItems.GroupID, Groups.GroupName, Groups.GroupNamee, Transaction_Details.OpeningReSalesQty, Transaction_Details.OpeningReSalesValue,"
    KSQLSrt = KSQLSrt & " ISNULL(Transaction_Details.TotalDiscountPerLine, 0) * ISNULL(Transactions.Currency_rate, 1) AS TotalDiscountPerLine, ISNULL(Transactions.Currency_rate, 1) AS Currency_rate,"
    KSQLSrt = KSQLSrt & " CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END AS 'ItemDiscountValue',"
    KSQLSrt = KSQLSrt & " Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) AS localprice, Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) * Transaction_Details.ShowQty AS LineNet,"
    KSQLSrt = KSQLSrt & " Transactions.NoteSerial1, TblItems.ItemNamee, Transactions.CashCustomerName, Transactions.CashCustomerPhone, Transactions.CashCustomerMobile, Transactions.CashCustomerAddress,"
    KSQLSrt = KSQLSrt & " Transactions.CashCustomerComment, Transactions.VAT, Transactions.VATNO, Transactions.Transaction_NetValue, Transaction_Details.Vat AS VatDet, Transaction_Details.Vatyo, Transactions.BasedOn"
    KSQLSrt = KSQLSrt & " FROM Transactions INNER JOIN"
    KSQLSrt = KSQLSrt & " Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblStore ON Transactions.StoreID = TblStore.StoreID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblCustemers ON Transactions.CusID = TblCustemers.CusID INNER JOIN"
    KSQLSrt = KSQLSrt & " TransactionTypes ON Transactions.Transaction_Type = TransactionTypes.Transaction_Type INNER JOIN"
    KSQLSrt = KSQLSrt & " TblItems ON Transaction_Details.Item_ID = TblItems.ItemID INNER JOIN"
    KSQLSrt = KSQLSrt & " Groups ON TblItems.GroupID = Groups.GroupID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblUnites ON Transaction_Details.UnitId = TblUnites.UnitID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblEmployee ON Transactions.Emp_ID = TblEmployee.Emp_ID"
    KSQLSrt = KSQLSrt & " Where (transactions.Transaction_Type = 21)"
    KSQLSrt = KSQLSrt & " and Transactions.PaymentType = 1"
    
    If Not IsNull(KFromDate.value) Then
        KSQLSrt = KSQLSrt & " and Transactions.Transaction_Date >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        KSQLSrt = KSQLSrt & "and Transactions.Transaction_Date <=" & SQLDate(KToDate.value, True) & ""
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open KSQLSrt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    For i = 1 To rs.RecordCount
        Sum2 = Sum2 + (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value)) - IIf(IsNull(rs("TotalDiscountPerLine").value), 0, rs("TotalDiscountPerLine").value) - IIf(IsNull(rs("ItemDiscountValue").value), 0, rs("ItemDiscountValue").value)
        rs.MoveNext
    Next i
    
    KSQLSrt = "SELECT Transaction_Details.CostPrice, TblCustemers.Fullcode, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, CONVERT(varchar, MONTH(Transactions.Transaction_Date)) + '-' + CONVERT(varchar,"
    KSQLSrt = KSQLSrt & " YEAR(Transactions.Transaction_Date)) AS Monthdate1, Transaction_Details.ShowQty, Transaction_Details.showPrice, Transactions.Transaction_ID, Transactions.Transaction_Serial, Transactions.Transaction_Date,"
    KSQLSrt = KSQLSrt & " TransactionTypes.TransactionTypeName, dbo.GetItemUnitFactor(Transaction_Details.Item_ID, 0) AS UnitFactor, TransactionTypes.TransactionEnglishName, TransactionTypes.StockEffect, Transactions.Transaction_Type,"
    KSQLSrt = KSQLSrt & " Transactions.PaymentType, Transactions.CusID, Transactions.StoreID, Transactions.UserID, Transactions.Emp_ID, Transaction_Details.Item_ID, Transaction_Details.Quantity, Transaction_Details.Price,"
    KSQLSrt = KSQLSrt & " Transaction_Details.UnitId, Transaction_Details.OpeningBurcahseValue, Transaction_Details.OpeningBurcahseQty, Transaction_Details.OpeningSalesQty, Transaction_Details.OpeningSalesValue, TblUnites.UnitName,"
    KSQLSrt = KSQLSrt & " TblStore.StoreName, TblCustemers.CusName, TblCustemers.CusNamee, Transactions.PayedValue, Transactions.BranchId, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBranchesData.Tel,"
    KSQLSrt = KSQLSrt & " TblItems.ItemCode, TblItems.ItemName, TblItems.GroupID, Groups.GroupName, Groups.GroupNamee, Transaction_Details.OpeningReSalesQty, Transaction_Details.OpeningReSalesValue,"
    KSQLSrt = KSQLSrt & " ISNULL(Transaction_Details.TotalDiscountPerLine, 0) * ISNULL(Transactions.Currency_rate, 1) AS TotalDiscountPerLine, ISNULL(Transactions.Currency_rate, 1) AS Currency_rate,"
    KSQLSrt = KSQLSrt & " CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
    KSQLSrt = KSQLSrt & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END AS 'ItemDiscountValue',"
    KSQLSrt = KSQLSrt & " Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) AS localprice, Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) * Transaction_Details.ShowQty AS LineNet,"
    KSQLSrt = KSQLSrt & " Transactions.NoteSerial1, TblItems.ItemNamee, Transactions.CashCustomerName, Transactions.CashCustomerPhone, Transactions.CashCustomerMobile, Transactions.CashCustomerAddress,"
    KSQLSrt = KSQLSrt & " Transactions.CashCustomerComment, Transactions.VAT, Transactions.VATNO, Transactions.Transaction_NetValue, Transaction_Details.Vat AS VatDet, Transaction_Details.Vatyo, Transactions.BasedOn"
    KSQLSrt = KSQLSrt & " FROM Transactions INNER JOIN"
    KSQLSrt = KSQLSrt & " Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblStore ON Transactions.StoreID = TblStore.StoreID INNER JOIN"
    KSQLSrt = KSQLSrt & " TblCustemers ON Transactions.CusID = TblCustemers.CusID INNER JOIN"
    KSQLSrt = KSQLSrt & " TransactionTypes ON Transactions.Transaction_Type = TransactionTypes.Transaction_Type INNER JOIN"
    KSQLSrt = KSQLSrt & " TblItems ON Transaction_Details.Item_ID = TblItems.ItemID INNER JOIN"
    KSQLSrt = KSQLSrt & " Groups ON TblItems.GroupID = Groups.GroupID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblUnites ON Transaction_Details.UnitId = TblUnites.UnitID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblEmployee ON Transactions.Emp_ID = TblEmployee.Emp_ID"
    KSQLSrt = KSQLSrt & " Where (transactions.Transaction_Type = 21)"
    KSQLSrt = KSQLSrt & " and Transactions.PaymentType = 2"
    
    If Not IsNull(KFromDate.value) Then
        KSQLSrt = KSQLSrt & " and Transactions.Transaction_Date >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        KSQLSrt = KSQLSrt & "and Transactions.Transaction_Date <=" & SQLDate(KToDate.value, True) & ""
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open KSQLSrt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    For i = 1 To rs.RecordCount
        Sum3 = Sum3 + (IIf(IsNull(rs("SHOWQTY").value), 0, rs("SHOWQTY").value)) * (IIf(IsNull(rs("SHOWpRICE").value), 0, rs("SHOWpRICE").value))
        rs.MoveNext
    Next i
    
    Text7.Text = total5
    Text5.Text = total1
    Text4.Text = total2
    Text3.Text = total3
    Text2.Text = total4
    Text1.Text = total3 - total4
    Label63 = Sum1
    Label65 = Sum2
    Label66 = Sum3
    If total5 <> 0 Then
    Text6.Text = Round((total3 / total5), 2)
    Else
    Text6.Text = 0
    End If
End Sub
Private Sub ALLButton6_Click()
    total1 = 0
    total2 = 0
    total3 = 0
    total4 = 0
    total5 = 0
    Sum1 = 0
    Sum2 = 0
    Sum3 = 0
    
    fillgrid1
    FillGrid2
    FillOthers
End Sub
Private Sub ALLButton7_Click()
    Dim MySQL As String
    Dim StrFileName As String
    
    MySQL = "SELECT Transaction_Details.CostPrice, TblCustemers.Fullcode, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, CONVERT(varchar, MONTH(Transactions.Transaction_Date)) + '-' + CONVERT(varchar,"
    MySQL = MySQL & " YEAR(Transactions.Transaction_Date)) AS Monthdate1, Transaction_Details.ShowQty, Transaction_Details.showPrice, Transactions.Transaction_ID, Transactions.Transaction_Serial, Transactions.Transaction_Date,"
    MySQL = MySQL & " TransactionTypes.TransactionTypeName, dbo.GetItemUnitFactor(Transaction_Details.Item_ID, 0) AS UnitFactor, TransactionTypes.TransactionEnglishName, TransactionTypes.StockEffect, Transactions.Transaction_Type,"
    MySQL = MySQL & " Transactions.PaymentType, Transactions.CusID, Transactions.StoreID, Transactions.UserID, Transactions.Emp_ID, Transaction_Details.Item_ID, Transaction_Details.Quantity, Transaction_Details.Price,"
    MySQL = MySQL & " Transaction_Details.UnitId, Transaction_Details.OpeningBurcahseValue, Transaction_Details.OpeningBurcahseQty, Transaction_Details.OpeningSalesQty, Transaction_Details.OpeningSalesValue, TblUnites.UnitName,"
    MySQL = MySQL & " TblStore.StoreName, TblCustemers.CusName, TblCustemers.CusNamee, Transactions.PayedValue, Transactions.BranchId, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBranchesData.Tel,"
    MySQL = MySQL & " TblItems.ItemCode, TblItems.ItemName, TblItems.GroupID, Groups.GroupName, Groups.GroupNamee, Transaction_Details.OpeningReSalesQty, Transaction_Details.OpeningReSalesValue,"
    MySQL = MySQL & " ISNULL(Transaction_Details.TotalDiscountPerLine, 0) * ISNULL(Transactions.Currency_rate, 1) AS TotalDiscountPerLine, ISNULL(Transactions.Currency_rate, 1) AS Currency_rate,"
    MySQL = MySQL & " CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))"
    MySQL = MySQL & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
    MySQL = MySQL & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END AS 'ItemDiscountValue',"
    MySQL = MySQL & " Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) AS localprice, Transaction_Details.showPrice * ISNULL(Transactions.Currency_rate, 1) * Transaction_Details.ShowQty AS LineNet,"
    MySQL = MySQL & " Transactions.NoteSerial1, TblItems.ItemNamee, Transactions.CashCustomerName, Transactions.CashCustomerPhone, Transactions.CashCustomerMobile, Transactions.CashCustomerAddress,"
    MySQL = MySQL & " Transactions.CashCustomerComment, Transactions.VAT, Transactions.VATNO, Transactions.Transaction_NetValue, Transaction_Details.Vat AS VatDet, Transaction_Details.Vatyo, Transactions.BasedOn"
    MySQL = MySQL & " FROM Transactions INNER JOIN"
    MySQL = MySQL & " Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID INNER JOIN"
    MySQL = MySQL & " TblStore ON Transactions.StoreID = TblStore.StoreID INNER JOIN"
    MySQL = MySQL & " TblCustemers ON Transactions.CusID = TblCustemers.CusID INNER JOIN"
    MySQL = MySQL & " TransactionTypes ON Transactions.Transaction_Type = TransactionTypes.Transaction_Type INNER JOIN"
    MySQL = MySQL & " TblItems ON Transaction_Details.Item_ID = TblItems.ItemID INNER JOIN"
    MySQL = MySQL & " Groups ON TblItems.GroupID = Groups.GroupID LEFT OUTER JOIN"
    MySQL = MySQL & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & " TblUnites ON Transaction_Details.UnitId = TblUnites.UnitID LEFT OUTER JOIN"
    MySQL = MySQL & " TblEmployee ON Transactions.Emp_ID = TblEmployee.Emp_ID"
    MySQL = MySQL & " Where (transactions.Transaction_Type = 21)"
    
    If Not IsNull(KFromDate.value) Then
        MySQL = MySQL & " and Transactions.Transaction_Date >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        MySQL = MySQL & "and Transactions.Transaction_Date <=" & SQLDate(KToDate.value, True) & ""
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepInvAlert.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepInvAlertE.rpt"
    End If
    
    print_report StrFileName, MySQL
End Sub
Private Sub ALLButton8_Click()
    
    Dim KSQLSrt  As String
    Dim StrFileName As String
    
    KSQLSrt = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.Note_Value, Notes.ExpensesID, ExpensesType.Name,ExpensesType.Namee, Notes.NoteSerial, Notes.Remark, TblUsers.UserName, Notes.BoxID, Notes.UserID, TblBoxesData.BoxName,"
    KSQLSrt = KSQLSrt & " BanksData.BankName , Notes.BankID, Notes.ChqueNum, Notes.dueDate, Notes.NoteSerial1, Notes.branch_no"
    KSQLSrt = KSQLSrt & " FROM TblUsers INNER JOIN"
    KSQLSrt = KSQLSrt & " ExpensesType INNER JOIN"
    KSQLSrt = KSQLSrt & " Notes ON ExpensesType.ID = Notes.ExpensesID ON TblUsers.UserID = Notes.UserID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " TblBoxesData ON Notes.BoxID = TblBoxesData.BoxID LEFT OUTER JOIN"
    KSQLSrt = KSQLSrt & " BanksData ON Notes.BankID = BanksData.BankID"
    KSQLSrt = KSQLSrt & " Where (Notes.NoteType = 3)"
    
    If Not IsNull(KFromDate.value) Then
        KSQLSrt = KSQLSrt & " and Notes.NoteDate >=" & SQLDate(KFromDate.value, True) & ""
    End If
    If Not IsNull(KToDate.value) Then
        KSQLSrt = KSQLSrt & " and Notes.NoteDate <=" & SQLDate(KToDate.value, True) & ""
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExpAlert.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExpAlertE.rpt"
    End If
    
    print_report StrFileName, KSQLSrt
End Sub
Function print_report(StrFileName As String, MySQL As String)
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim Msg As String
    

    
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If


    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabComment
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.EngComment
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue KFromDate.value
    xReport.ParameterFields(5).AddCurrentValue KToDate.value
    xReport.ParameterFields(6).AddCurrentValue Sum1
    xReport.ParameterFields(7).AddCurrentValue Sum2
    xReport.ParameterFields(8).AddCurrentValue Sum3
    xReport.ParameterFields(9).AddCurrentValue total1
    Dim temp1, temp2, temp3, temp4, temp5 As String             '}
    temp1 = Format(total2, "#.00")                              '}
    temp2 = Format(total3, "#.00")                              '}
    xReport.ParameterFields(10).AddCurrentValue temp1           '}
    xReport.ParameterFields(11).AddCurrentValue temp2           '}
    xReport.ParameterFields(12).AddCurrentValue total4          '}      <========     3ayez el 7ar2 bgaz
    xReport.ParameterFields(13).AddCurrentValue total5          '}
    temp3 = Format((total1 - total2), "#.00")                   '}
    xReport.ParameterFields(14).AddCurrentValue temp3           '}
    temp4 = Format((total3 - total4), "#.00")                   '}
    xReport.ParameterFields(15).AddCurrentValue temp4           '}
    temp5 = Format((total3 / total5), "#.00")
    xReport.ParameterFields(16).AddCurrentValue temp5
    
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

Private Sub VSFlexGrid7_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim StrSQL As String
Dim allTransaction_ID As Double
    With Me.VSFlexGrid7
 allTransaction_ID = val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
        
If allTransaction_ID = 0 Then Exit Sub
'allTransaction_ID = Mid(allTransaction_ID, 2, Len(allTransaction_ID))
StrSQL = "select * from Transactions where Transaction_ID in (" & allTransaction_ID & ")"
frmsalebill.show
  frmsalebill.generalSearch (StrSQL)
            
frmsalebill.invoiceSerach = True

  
End With
End Sub

Private Sub VSFlexGrid7_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid7
Select Case .ColKey(Col)
           Case "Show"
           .ColComboList(.ColIndex("Show")) = "..."
 End Select
 End With
 
End Sub
