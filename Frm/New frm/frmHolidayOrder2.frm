VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmHolidayorder2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· „” Õﬁ«  «·«Ã«“…"
   ClientHeight    =   8760
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14700
   HelpContextID   =   580
   Icon            =   "frmHolidayOrder2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14700
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8760
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14700
      _cx             =   25929
      _cy             =   15452
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmHolidayOrder2.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7725
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14640
         _cx             =   25823
         _cy             =   13626
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
         Caption         =   "»Ì«‰«  «·ÿ·»|„Êﬁ› «·⁄Âœ"
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
         Begin VB.Frame Frame8 
            Caption         =   "„Êﬁ› «·⁄Âœ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7305
            Left            =   15285
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   45
            Width           =   14550
            Begin VB.TextBox Text6 
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
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Text            =   "0"
               Top             =   1320
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox Text5 
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
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   6240
               Left            =   5880
               TabIndex        =   73
               Top             =   360
               Width           =   8535
               _cx             =   15055
               _cy             =   11007
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
               Rows            =   50
               Cols            =   25
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmHolidayOrder2.frx":040F
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«’Ê· «· Ì  ﬂ«‰  »⁄Âœ …"
               Height          =   195
               Index           =   37
               Left            =   11760
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   1680
               Width           =   2460
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·”·› ⁄·Ï «·„ÊŸ›"
               Height          =   195
               Index           =   36
               Left            =   11760
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   1320
               Visible         =   0   'False
               Width           =   2460
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «·⁄Âœ «·‰ﬁœÌ… ⁄·Ï «·„ÊŸ›"
               Height          =   195
               Index           =   34
               Left            =   11760
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   960
               Visible         =   0   'False
               Width           =   2460
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Height          =   435
               Index           =   35
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   360
               Width           =   1860
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " ﬁÊ„ Â–… «·Ã“∆Ì… » €ÌÌ— „Êﬁ› «·„ÊŸ› „‰ «·”·› Ê «·⁄Âœ Ê«·«’Ê· «· Ì »⁄Âœ …"
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
               Height          =   615
               Index           =   32
               Left            =   120
               TabIndex        =   66
               Top             =   120
               Visible         =   0   'False
               Width           =   4725
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7305
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   14550
            _cx             =   25665
            _cy             =   12885
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7275
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   14745
               _cx             =   26009
               _cy             =   12832
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
               Begin VB.Frame Frame9 
                  Caption         =   "«Ã„«·Ï «·„” Õﬁ« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   1185
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   6600
                  Width           =   14550
                  Begin VB.TextBox Text15 
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
                     Left            =   9360
                     RightToLeft     =   -1  'True
                     TabIndex        =   106
                     Text            =   "0"
                     Top             =   720
                     Width           =   1335
                  End
                  Begin VB.TextBox Text3 
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
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   105
                     Text            =   "0"
                     Top             =   720
                     Width           =   1095
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·«Ã„«·Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Index           =   27
                     Left            =   1560
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
                     Top             =   120
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   26
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   120
                     Width           =   1860
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "»ﬁÌ„…"
                     Height          =   195
                     Index           =   42
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   720
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " –ﬂ—…"
                     Height          =   195
                     Index           =   40
                     Left            =   11880
                     RightToLeft     =   -1  'True
                     TabIndex        =   109
                     Top             =   720
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·«Ã„«·Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Index           =   39
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   720
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   25
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   840
                     Width           =   1500
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "«Ã„«·Ì «·«” ﬁÿ«⁄« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   3405
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   3240
                  Width           =   14415
                  Begin VB.CheckBox Check4 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Ì „ Œ’„"
                     Height          =   255
                     Left            =   12240
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   1080
                     Width           =   1815
                  End
                  Begin VB.CheckBox Check3 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Ì „ Œ’„"
                     Height          =   255
                     Left            =   12240
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   720
                     Width           =   1815
                  End
                  Begin VB.CheckBox Check2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Ì „ Œ’„"
                     Height          =   255
                     Left            =   12360
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.TextBox Text13 
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
                     Left            =   9360
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Text            =   "0"
                     Top             =   720
                     Width           =   1335
                  End
                  Begin VB.TextBox Text11 
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
                     Left            =   9360
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Text            =   "0"
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
                     Height          =   1860
                     Left            =   4560
                     TabIndex        =   99
                     Top             =   1320
                     Width           =   8160
                     _cx             =   14393
                     _cy             =   3281
                     Appearance      =   1
                     BorderStyle     =   1
                     Enabled         =   0   'False
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
                     GridLines       =   3
                     GridLinesFixed  =   2
                     GridLineWidth   =   5
                     Rows            =   1
                     Cols            =   12
                     FixedRows       =   1
                     FixedCols       =   0
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"frmHolidayOrder2.frx":07F2
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
                     Begin VB.PictureBox PicDes 
                        BorderStyle     =   0  'None
                        Height          =   1635
                        Left            =   240
                        RightToLeft     =   -1  'True
                        ScaleHeight     =   1635
                        ScaleWidth      =   2925
                        TabIndex        =   100
                        Top             =   960
                        Visible         =   0   'False
                        Width           =   2925
                        Begin VB.TextBox TxtDes 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H80000018&
                           BorderStyle     =   0  'None
                           Height          =   1125
                           Left            =   30
                           MultiLine       =   -1  'True
                           RightToLeft     =   -1  'True
                           ScrollBars      =   3  'Both
                           TabIndex        =   101
                           Top             =   360
                           Visible         =   0   'False
                           Width           =   2115
                        End
                        Begin VB.Label LblDes 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H8000000C&
                           Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                           ForeColor       =   &H0000C8FF&
                           Height          =   315
                           Left            =   0
                           RightToLeft     =   -1  'True
                           TabIndex        =   102
                           Top             =   0
                           Width           =   2445
                        End
                     End
                     Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                        Height          =   315
                        Left            =   240
                        TabIndex        =   103
                        ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                        Top             =   600
                        Visible         =   0   'False
                        Width           =   2955
                        _cx             =   1973752924
                        _cy             =   1973748268
                        Alignment       =   0
                        Appearance      =   3
                        AutoSearch      =   0   'False
                        BackColor       =   -2147483624
                        BackgroundColor =   -2147483633
                        BorderColor     =   0
                        BorderVisible   =   -1  'True
                        Caption         =   "SmartCombo1"
                        CaptionAlignment=   4
                        CaptionBackColor=   -2147483633
                        BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        CaptionForeColor=   -2147483630
                        CaptionHeight   =   15
                        CaptionOnTop    =   0   'False
                        CaptionMultiLine=   0
                        Checkbox3D      =   0   'False
                        CheckboxAlignment=   5
                        CheckboxBackColor=   16777215
                        CheckboxSize    =   13
                        CheckboxValue   =   0
                        BrowsePictureAlignment=   5
                        BrowsePictureStretchH=   0
                        BrowsePictureStretchV=   0
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
                        ForeColor       =   0
                        Gap             =   0
                        HideSelection   =   -1  'True
                        Locked          =   0   'False
                        MaxLength       =   0
                        MultiLine       =   0
                        OnFocus         =   3
                        PasswordChar    =   ""
                        Picture         =   "frmHolidayOrder2.frx":09CF
                        PictureAlignment=   5
                        PictureBackColor=   -2147483624
                        PictureStretchH =   0
                        PictureStretchV =   0
                        Redraw          =   -1  'True
                        ScrollBar       =   0
                        Style           =   0
                        Text            =   ""
                        UnderLine       =   0   'False
                        Enabled0        =   -1  'True
                        Position0       =   0
                        Tip0            =   "Caption"
                        Visible0        =   0   'False
                        Width0          =   90
                        Enabled1        =   -1  'True
                        Position1       =   1
                        Tip1            =   ""
                        Visible1        =   -1  'True
                        Width1          =   32
                        Enabled2        =   -1  'True
                        Position2       =   2
                        Tip2            =   "Check Box (Space, Ctrl + Space)"
                        Visible2        =   0   'False
                        Width2          =   16
                        Enabled3        =   -1  'True
                        Position3       =   3
                        Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                        Visible3        =   -1  'True
                        Width3          =   145
                        Enabled4        =   -1  'True
                        Position4       =   4
                        Tip4            =   "Left Spinner (Alt + Left)"
                        Visible4        =   0   'False
                        Width4          =   16
                        Enabled5        =   -1  'True
                        Position5       =   5
                        Tip5            =   "Right Spinner (Alt + Right)"
                        Visible5        =   0   'False
                        Width5          =   16
                        Enabled6        =   -1  'True
                        Position6       =   6
                        Tip6            =   "Up Spinner (Ctrl + Up)"
                        Visible6        =   0   'False
                        Width6          =   16
                        Enabled7        =   -1  'True
                        Position7       =   7
                        Tip7            =   "Down Spinner (Ctrl + Down)"
                        Visible7        =   0   'False
                        Width7          =   16
                        Enabled8        =   -1  'True
                        Position8       =   8
                        Tip8            =   "Browse (Alt + Enter)"
                        Visible8        =   0   'False
                        Width8          =   16
                        Enabled9        =   -1  'True
                        Position9       =   9
                        Tip9            =   " (Alt + Down, F4)"
                        Visible9        =   -1  'True
                        Width9          =   16
                        Enabled10       =   -1  'True
                        Position10      =   10
                        Tip10           =   "Right Arrow (Alt + >)"
                        Visible10       =   0   'False
                        Width10         =   16
                     End
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„’—Ê›«  «Œ—Ï"
                     Height          =   195
                     Index           =   24
                     Left            =   10320
                     RightToLeft     =   -1  'True
                     TabIndex        =   98
                     Top             =   1080
                     Width           =   1860
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·”·› »ﬁÌ„…"
                     Height          =   195
                     Index           =   23
                     Left            =   10680
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   240
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄Âœ »ﬁÌ„…"
                     Height          =   195
                     Index           =   22
                     Left            =   10800
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   720
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·«Ã„«·Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Index           =   21
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   2880
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   20
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   91
                     Top             =   2880
                     Width           =   1860
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "«·„” Õﬁ«  «·„«œÌ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   1245
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   2085
                  Width           =   14415
                  Begin VB.TextBox Text10 
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
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Text            =   "0"
                     Top             =   720
                     Width           =   1095
                  End
                  Begin VB.TextBox Text9 
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
                     Left            =   9360
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Text            =   "0"
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.TextBox Text1 
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
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Text            =   "0"
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.TextBox Text2 
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
                     Left            =   9360
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Text            =   "0"
                     Top             =   720
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   18
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   840
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·«Ã„«·Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Index           =   16
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   720
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " –ﬂ—…"
                     Height          =   195
                     Index           =   17
                     Left            =   11880
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   720
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ÌÊ„"
                     Height          =   195
                     Index           =   14
                     Left            =   11880
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   240
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "»ﬁÌ„…"
                     Height          =   195
                     Index           =   19
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   720
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "»ﬁÌ„…"
                     Height          =   195
                     Index           =   13
                     Left            =   10200
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   240
                     Width           =   1500
                  End
               End
               Begin VB.Frame Frame2 
                  Caption         =   "ÊÕœ… «·„›—œ"
                  Enabled         =   0   'False
                  Height          =   510
                  Left            =   1125
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   8175
                  Visible         =   0   'False
                  Width           =   5910
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬁÌ„…"
                     Height          =   195
                     Index           =   0
                     Left            =   3840
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«Ì«„"
                     Height          =   195
                     Index           =   1
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "”«⁄« "
                     Height          =   195
                     Index           =   2
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.TextBox TxtRowNumber 
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
                  Left            =   600
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Text            =   "0"
                  Top             =   1215
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«— ﬂ· «·„ÊŸ›Ì‰"
                  Height          =   255
                  Left            =   -4170
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.Frame Frame1 
                  Caption         =   "»Ì«‰«  «·ÿ·»"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   960
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   1440
                  Width           =   14475
                  Begin VB.TextBox TxtSearchCode 
                     Alignment       =   2  'Center
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
                     Left            =   5640
                     TabIndex        =   78
                     Top             =   240
                     Width           =   1050
                  End
                  Begin VB.TextBox Text8 
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
                     Height          =   330
                     Left            =   11040
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   240
                     Width           =   1470
                  End
                  Begin VB.OptionButton Option4 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·›—⁄"
                     Height          =   210
                     Left            =   16680
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   500
                     Width           =   1575
                  End
                  Begin VB.OptionButton Option3 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬁ”„ „⁄Ì‰"
                     Height          =   210
                     Left            =   14640
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   840
                     Width           =   1575
                  End
                  Begin VB.TextBox TxtValue1 
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
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Text            =   "0"
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   495
                  End
                  Begin VB.TextBox TxtValue 
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
                     Left            =   4200
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Text            =   "0"
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   855
                  End
                  Begin VB.TextBox TxtRemarks 
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
                     Height          =   615
                     Left            =   960
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   53
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1455
                  End
                  Begin VB.OptionButton Option1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
                     Height          =   210
                     Left            =   15720
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.OptionButton Option2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«Œ Ì«— «·„ÊŸ›Ì‰"
                     Height          =   210
                     Left            =   14280
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   1200
                     Width           =   1455
                  End
                  Begin MSComCtl2.DTPicker DTPicker10 
                     Height          =   360
                     Left            =   7920
                     TabIndex        =   77
                     Top             =   240
                     Width           =   1830
                     _ExtentX        =   3228
                     _ExtentY        =   635
                     _Version        =   393216
                     Format          =   96796673
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DCEmployee 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   79
                     Top             =   240
                     Width           =   2790
                     _ExtentX        =   4921
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "  «·„ÊŸ›"
                     Height          =   195
                     Index           =   3
                     Left            =   6315
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   240
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «·„Ê«›ﬁ…"
                     Height          =   435
                     Index           =   12
                     Left            =   9600
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   240
                     Width           =   1185
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·ÿ·»"
                     Height          =   195
                     Index           =   11
                     Left            =   13455
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   240
                     Width           =   750
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " œﬁÌﬁ…"
                     Height          =   225
                     Index           =   10
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   1155
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     Height          =   225
                     Index           =   9
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «·ﬁÌ„…"
                     Height          =   465
                     Index           =   4
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   795
                  End
               End
               Begin VB.TextBox txtid 
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
                  Height          =   330
                  Left            =   11100
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   1125
                  Width           =   1470
               End
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
                  Height          =   480
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   2895
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   930
                  Index           =   5
                  Left            =   -510
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   15045
                  _cx             =   26538
                  _cy             =   1640
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
                  Picture         =   "frmHolidayOrder2.frx":0F69
                  Caption         =   " ”ÃÌ· „” Õﬁ«  «·«Ã«“…"
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
                  Begin ImpulseButton.ISButton XPBtnMove 
                     Height          =   375
                     Index           =   0
                     Left            =   1695
                     TabIndex        =   8
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
                     ButtonImage     =   "frmHolidayOrder2.frx":1C43
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
                     TabIndex        =   9
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
                     ButtonImage     =   "frmHolidayOrder2.frx":1FDD
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
                     TabIndex        =   10
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
                     ButtonImage     =   "frmHolidayOrder2.frx":2377
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
                     TabIndex        =   11
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
                     ButtonImage     =   "frmHolidayOrder2.frx":2711
                     ColorHighlight  =   4194304
                     ColorHoverText  =   16777215
                     ColorShadow     =   -2147483631
                     ColorOutline    =   -2147483631
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                     ColorToggledHoverText=   16777215
                     ColorTextShadow =   16777215
                  End
                  Begin VB.Image ImgFavorites 
                     Height          =   390
                     Left            =   5280
                     Picture         =   "frmHolidayOrder2.frx":2AAB
                     Stretch         =   -1  'True
                     Top             =   0
                     Width           =   525
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   600
                  Index           =   3
                  Left            =   -4095
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   1005
                  Visible         =   0   'False
                  Width           =   4950
                  _cx             =   8731
                  _cy             =   1058
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
                  Caption         =   " Õœœ «·”‰… «·„«·Ì…"
                  Align           =   0
                  AutoSizeChildren=   0
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
                  Style           =   1
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
                  Begin VB.ComboBox CboYear 
                     Height          =   315
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   14
                     Top             =   165
                     Width           =   1005
                  End
                  Begin VB.ComboBox CmbMonth 
                     Height          =   315
                     Left            =   75
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   13
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   1485
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰…"
                     Height          =   240
                     Index           =   2
                     Left            =   3795
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   180
                     Width           =   1020
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     Height          =   195
                     Index           =   0
                     Left            =   1425
                     RightToLeft     =   -1  'True
                     TabIndex        =   15
                     Top             =   270
                     Visible         =   0   'False
                     Width           =   645
                  End
               End
               Begin MSComCtl2.DTPicker XPDtbTrans 
                  Height          =   360
                  Left            =   8055
                  TabIndex        =   21
                  Top             =   1095
                  Width           =   1830
                  _ExtentX        =   3228
                  _ExtentY        =   635
                  _Version        =   393216
                  Format          =   96796673
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   12360
                  TabIndex        =   54
                  Top             =   6840
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmHolidayOrder2.frx":6713
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   7560
                  TabIndex        =   59
                  Top             =   1590
                  Visible         =   0   'False
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmHolidayOrder2.frx":6CAD
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„œ… «·«Ã«“… «·„ÿ·Ê»…"
                  Height          =   195
                  Index           =   15
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   0
                  Width           =   1620
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   6960
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.Label LblSum 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   6840
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.Label LBLWhereSTR 
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
                  Height          =   255
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   1710
                  Visible         =   0   'False
                  Width           =   1740
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„›—œ"
                  Height          =   195
                  Index           =   5
                  Left            =   14475
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   1770
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·”‰œ"
                  Height          =   435
                  Index           =   8
                  Left            =   10110
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   1125
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”‰œ"
                  Height          =   435
                  Index           =   7
                  Left            =   13635
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   1125
                  Width           =   750
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   540
                  Left            =   13305
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   1125
                  Width           =   975
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   7770
         Width           =   14640
         _cx             =   25823
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
            TabIndex        =   26
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
            ButtonImage     =   "frmHolidayOrder2.frx":7047
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   27
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
            ButtonImage     =   "frmHolidayOrder2.frx":73E1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   28
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
            ButtonImage     =   "frmHolidayOrder2.frx":777B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   11100
            TabIndex        =   34
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
            Left            =   10200
            TabIndex        =   35
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
            Left            =   9390
            TabIndex        =   36
            Top             =   480
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
            Left            =   8235
            TabIndex        =   37
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
            Left            =   7080
            TabIndex        =   38
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
            Left            =   5160
            TabIndex        =   39
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
            Left            =   5910
            TabIndex        =   40
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
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   32
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
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   2
            Left            =   3570
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   225
            Width           =   4695
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   10185
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   225
            Width           =   1455
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
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
      ButtonImage     =   "frmHolidayOrder2.frx":7B15
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmHolidayorder2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
  lbl(16).Caption = "Total"
   lbl(21).Caption = "Total"
    lbl(27).Caption = "Total"
        lbl(13).Caption = "Value"
            lbl(19).Caption = "Value"
            lbl(14).Caption = "Day'"
            lbl(17).Caption = "Ticket"
       lbl(7).Caption = "No Reg"
       lbl(8).Caption = "Date Reg"
       Frame1.Caption = "Data of Order"
       lbl(11).Caption = "OPR#"
       lbl(3).Caption = "Employee'"
     Check3.Caption = "Be Deducted"
     Check4.Caption = "Be Deducted"
       Check2.Caption = "Be Deducted"
       lbl(23).Caption = "Advances value"
       Frame3.Caption = "Financial Receivables"
       Frame4.Caption = "Financial Receivables"
       lbl(12).Caption = "Accept Date"
       lbl(22).Caption = "Testament Values"
       lbl(24).Caption = "Other Expenses"
    FrmHolidayorder2.Caption = "Registration Dues Holiday"
 Ele(5).Caption = "Registration Dues Holiday"
 Frame9.Caption = "Total Dues"
 lbl(4).Caption = "Value"
 lbl(10).Caption = "Minutes"
 Ele(3).Caption = "Select  Fiscal Year"
 lbl(9).Caption = "Remarks"
 lbl(36).Caption = "Total Advances to Employee"
 lbl(34).Caption = "Testament total cash employee"
 LblDes.Caption = "You can write a comment here"
 lbl(37).Caption = "Assets that were in the custody of"
 With Fg_Journal
  .TextMatrix(0, .ColIndex("LineNo")) = "Serial"
  .TextMatrix(0, .ColIndex("AccountName")) = "Expense"
  .TextMatrix(0, .ColIndex("value")) = "Value"
    End With
    C1Tab1.Caption = "Data Accounting  | Data Order"
    lbl(32).Caption = "These partial change the position of the employee advances and Testament and assets under his custody"
    Frame8.Caption = "Important Data"
   With Grid
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("Emp_Code")) = "Assets Code"
  .TextMatrix(0, .ColIndex("Emp_Name")) = "Assets Name"
  .TextMatrix(0, .ColIndex("remarks")) = "Recipient of the new"
    End With
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

 
 
   

    Label2(0).Caption = "Current Record"
    Label2(2).Caption = "Total Record"

    

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Emp_code")) = "Emp_code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp_Name"
        .TextMatrix(0, .ColIndex("Hdate")) = "Holday Date"
        .TextMatrix(0, .ColIndex("NoofDays")) = "Long Holday"
        .TextMatrix(0, .ColIndex("LastArrDate")) = "Last Back"
        .TextMatrix(0, .ColIndex("ActualDate")) = "Actual Date vacation"
        .TextMatrix(0, .ColIndex("remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("NoOfHour")) = "Hours"

    End With

End Sub

Private Sub Form_Load()
Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption

End Sub

