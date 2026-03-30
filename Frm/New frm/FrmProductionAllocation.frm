VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmProductionAllocation 
   BackColor       =   &H00E2E9E9&
   Caption         =   "  Œ’Ì’ ŒÿÊÿ «·√‰ «Ã ·√Ê«„— «·‘€· Ê«·ﬂ„Ì«  «·„‰ ÃÂ"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18870
   HelpContextID   =   580
   Icon            =   "FrmProductionAllocation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   18870
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   10500
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18870
      _cx             =   33285
      _cy             =   18521
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   9720
         Left            =   30
         TabIndex        =   1
         Top             =   0
         Width           =   18945
         _cx             =   33417
         _cy             =   17145
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
         Caption         =   " Œ’Ì’ «·«‰ «Ã|«·„Ê«œ «·Œ«„ «·„ﬁœ—…"
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
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   9300
            Left            =   19590
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   45
            Width           =   18855
            _cx             =   33258
            _cy             =   16404
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
            Begin VSFlex8UCtl.VSFlexGrid FG1 
               Height          =   7575
               Left            =   480
               TabIndex        =   95
               Top             =   600
               Width           =   12300
               _cx             =   21696
               _cy             =   13361
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
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   17
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProductionAllocation.frx":038A
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   9300
            Left            =   45
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   18855
            _cx             =   33258
            _cy             =   16404
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
               Left            =   -720
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   8640
               Width           =   18975
               _cx             =   33470
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
               Begin VB.TextBox TxtNoteSerial1V 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   13320
                  TabIndex        =   135
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.TextBox txtTransaction_ID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3600
                  TabIndex        =   134
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10800
                  TabIndex        =   130
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   6000
                  TabIndex        =   129
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.TextBox TxtresiveVoucher 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   11160
                  Locked          =   -1  'True
                  TabIndex        =   127
                  Top             =   840
                  Width           =   1935
               End
               Begin VB.CommandButton CmdResiveVoucher 
                  Caption         =   "«‰‘«¡ «–‰ «÷«›… «·Ì"
                  Height          =   315
                  Left            =   14040
                  TabIndex        =   126
                  Top             =   840
                  Width           =   2880
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "⁄—÷ «·«–‰"
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   125
                  Top             =   120
                  Width           =   1560
               End
               Begin VB.CommandButton CnsSHowGl 
                  Caption         =   "⁄—÷ «·ﬁÌœ"
                  Height          =   315
                  Left            =   7680
                  TabIndex        =   124
                  Top             =   120
                  Width           =   1200
               End
               Begin MSComCtl2.DTPicker ReciveDate 
                  Height          =   315
                  Left            =   16320
                  TabIndex        =   128
                  Top             =   120
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   94502913
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   315
                  Index           =   41
                  Left            =   12000
                  TabIndex        =   138
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ ”‰œ «·«” ·«„"
                  Height          =   315
                  Index           =   40
                  Left            =   15000
                  TabIndex        =   137
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «·«” ·«„"
                  Height          =   315
                  Index           =   39
                  Left            =   17640
                  TabIndex        =   136
                  Top             =   120
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "»Ì«‰«  «·ﬂ„Ì… «·„‰ Ã…"
               Height          =   1815
               Left            =   -6120
               TabIndex        =   47
               Top             =   1200
               Visible         =   0   'False
               Width           =   5295
               Begin VB.TextBox TxttotalMaterials 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   600
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   50
                  Top             =   600
                  Width           =   2145
               End
               Begin VB.TextBox TxttotalMaterialsForItems 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   600
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   49
                  Top             =   960
                  Width           =   2145
               End
               Begin VB.TextBox TxtTotalProductionQty 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   600
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   48
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·„Ê«œ"
                  Height          =   315
                  Index           =   23
                  Left            =   3120
                  TabIndex        =   53
                  Top             =   720
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„…«·ÊÕœ… ÿ»ﬁ« ··‰”»"
                  Height          =   315
                  Index           =   22
                  Left            =   3120
                  TabIndex        =   52
                  Top             =   1080
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "√Ã„«·Ì «·ﬂ„Ì… «·„‰ Ã…"
                  Height          =   315
                  Index           =   21
                  Left            =   3120
                  TabIndex        =   51
                  Top             =   360
                  Width           =   1935
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "»Ì«‰«  √„— «·‘€·"
               Height          =   2055
               Left            =   9120
               TabIndex        =   40
               Top             =   -1440
               Visible         =   0   'False
               Width           =   4575
               Begin VB.TextBox TxttotalOrderQty 
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
                  Left            =   1800
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   42
                  Top             =   960
                  Width           =   1425
               End
               Begin VB.TextBox TxtWorkOrderNO 
                  Alignment       =   2  'Center
                  BackColor       =   &H80000002&
                  Height          =   315
                  Left            =   1080
                  MaxLength       =   50
                  TabIndex        =   41
                  Top             =   240
                  Width           =   2145
               End
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   43
                  Top             =   600
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
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
               Begin MSDataListLib.DataCombo DDcunits 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   96
                  Top             =   960
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
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
               Begin MSDataListLib.DataCombo DCItemID 
                  Height          =   312
                  Left            =   0
                  TabIndex        =   98
                  Top             =   1680
                  Width           =   3132
                  _ExtentX        =   5530
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
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
                  Caption         =   "«·’‰› «·„‰ Ã"
                  Height          =   312
                  Index           =   28
                  Left            =   2520
                  TabIndex        =   99
                  Top             =   1680
                  Width           =   1932
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÊÕœ…"
                  Height          =   555
                  Index           =   6
                  Left            =   960
                  TabIndex        =   97
                  Top             =   960
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬂ„Ì… «·„ÿ·Ê»Â ··«‰ «Ã"
                  Height          =   435
                  Index           =   12
                  Left            =   3120
                  TabIndex        =   46
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„Ì·"
                  Height          =   555
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   45
                  Top             =   600
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·√„—"
                  Height          =   315
                  Index           =   19
                  Left            =   3480
                  TabIndex        =   44
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "»Ì«‰«  «·Œÿ"
               Height          =   1575
               Left            =   7800
               TabIndex        =   28
               Top             =   1200
               Width           =   11055
               Begin VB.TextBox TxtStoreID 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   133
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.TextBox TXTUsedPowerPriceH 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   7680
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   114
                  Top             =   720
                  Width           =   1188
               End
               Begin VB.TextBox TXTTotalUsedPowerPriceH 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   6120
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   113
                  Top             =   720
                  Width           =   1425
               End
               Begin VB.TextBox TxtTotalSalariesaLL 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1440
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   34
                  Top             =   240
                  Width           =   1425
               End
               Begin VB.TextBox TxtTotalElectricalsaLL 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1440
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   33
                  Top             =   720
                  Width           =   1425
               End
               Begin VB.TextBox TxttotalLineExpensesaLL 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   6120
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   32
                  Top             =   1080
                  Width           =   1425
               End
               Begin VB.TextBox TxttotalLineExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   7680
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   31
                  Top             =   1080
                  Width           =   1188
               End
               Begin VB.TextBox TxtTotalElectricals 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2880
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   30
                  Top             =   720
                  Width           =   1425
               End
               Begin VB.TextBox TxtTotalSalaries 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2880
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1425
               End
               Begin MSDataListLib.DataCombo dcLineID 
                  Height          =   288
                  Left            =   6120
                  TabIndex        =   35
                  Top             =   240
                  Width           =   2772
                  _ExtentX        =   4895
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   -2147483646
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
               Begin MSDataListLib.DataCombo DCStores 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   131
                  Top             =   1200
                  Width           =   4320
                  _ExtentX        =   7620
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   -2147483646
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
                  Caption         =   "«·„Œ“‰"
                  Height          =   312
                  Index           =   26
                  Left            =   5280
                  TabIndex        =   132
                  Top             =   1200
                  Width           =   732
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì „’«—Ì› «·ÊﬁÊœ"
                  Height          =   315
                  Index           =   36
                  Left            =   8880
                  TabIndex        =   115
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œÿ «·«‰ «Ã"
                  Height          =   228
                  Index           =   11
                  Left            =   9120
                  TabIndex        =   39
                  Top             =   360
                  Width           =   1572
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„’«—Ì›"
                  Height          =   315
                  Index           =   18
                  Left            =   8880
                  TabIndex        =   38
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì „’«—Ì› «·ﬂÂ—»«¡"
                  Height          =   315
                  Index           =   17
                  Left            =   3840
                  TabIndex        =   37
                  Top             =   720
                  Width           =   2175
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·«ÃÊ—"
                  Height          =   312
                  Index           =   16
                  Left            =   4560
                  TabIndex        =   36
                  Top             =   240
                  Width           =   1452
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "««·› —…"
               Height          =   615
               Left            =   6360
               TabIndex        =   21
               Top             =   2880
               Width           =   12375
               Begin VB.TextBox Text1 
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
                  Left            =   120
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   117
                  Top             =   1440
                  Width           =   1425
               End
               Begin VB.TextBox TxtNoOfHours 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
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
                  Left            =   1560
                  MaxLength       =   50
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1425
               End
               Begin MSComCtl2.DTPicker DBfromTime 
                  Height          =   285
                  Left            =   8400
                  TabIndex        =   23
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CustomFormat    =   "'Time: 'hh:mm tt"
                  Format          =   94502915
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker DBtoTime 
                  Height          =   285
                  Left            =   4080
                  TabIndex        =   24
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   12648447
                  CustomFormat    =   "'Time: 'hh:mm tt"
                  Format          =   94502915
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker fromdate 
                  Height          =   270
                  Left            =   10080
                  TabIndex        =   119
                  Top             =   240
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   476
                  _Version        =   393216
                  CalendarBackColor=   -2147483646
                  Format          =   94502913
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker todate 
                  Height          =   270
                  Left            =   5760
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   476
                  _Version        =   393216
                  CalendarBackColor=   -2147483646
                  Format          =   94502913
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„‘—›"
                  Height          =   315
                  Index           =   37
                  Left            =   1680
                  TabIndex        =   116
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·”«⁄« "
                  Height          =   315
                  Index           =   15
                  Left            =   2760
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   195
                  Index           =   14
                  Left            =   6960
                  TabIndex        =   26
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   13
                  Left            =   11760
                  TabIndex        =   25
                  Top             =   240
                  Width           =   375
               End
            End
            Begin VB.TextBox Txtid 
               Alignment       =   1  'Right Justify
               Height          =   285
               HideSelection   =   0   'False
               Left            =   15900
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   840
               Width           =   1320
            End
            Begin VB.CheckBox ChkLocked 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ìﬁ«› «· ⁄«„·"
               Height          =   210
               Left            =   19260
               TabIndex        =   19
               Top             =   3315
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
               Height          =   1155
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   1515
               Width           =   7530
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   5
               Left            =   0
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   0
               Width           =   18795
               _cx             =   33152
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
               Picture         =   "FrmProductionAllocation.frx":063B
               Caption         =   "  Œ’Ì’ ŒÿÊÿ «·√‰ «Ã ·√Ê«„— «·‘€· Ê«·ﬂ„Ì«  «·„‰ ÃÂ  "
               Align           =   0
               AutoSizeChildren=   7
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
                  TabIndex        =   55
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
                  ButtonImage     =   "FrmProductionAllocation.frx":1315
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
                  TabIndex        =   56
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
                  ButtonImage     =   "FrmProductionAllocation.frx":16AF
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
                  TabIndex        =   57
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
                  ButtonImage     =   "FrmProductionAllocation.frx":1A49
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
                  TabIndex        =   58
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
                  ButtonImage     =   "FrmProductionAllocation.frx":1DE3
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
            Begin MSComCtl2.DTPicker dbFromDate 
               Height          =   270
               Left            =   2985
               TabIndex        =   59
               Top             =   480
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   476
               _Version        =   393216
               Format          =   94502913
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker dbTodate 
               Height          =   270
               Left            =   240
               TabIndex        =   60
               Top             =   480
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   476
               _Version        =   393216
               Format          =   94502913
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker dbOPrDate 
               Height          =   270
               Left            =   13320
               TabIndex        =   61
               Top             =   840
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   476
               _Version        =   393216
               CalendarBackColor=   -2147483646
               Format          =   94502913
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo DCranch 
               Height          =   315
               Left            =   8160
               TabIndex        =   62
               Top             =   840
               Width           =   4215
               _ExtentX        =   7435
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6180
               Index           =   2
               Left            =   0
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   2760
               Width           =   18855
               _cx             =   33258
               _cy             =   10901
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
                  Left            =   1920
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   122
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   12930
               End
               Begin VB.TextBox TxtAlarm 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   8160
                  TabIndex        =   112
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.TextBox TxtWorkOrderNOSub 
                  Alignment       =   2  'Center
                  BackColor       =   &H80000002&
                  Height          =   315
                  Left            =   14640
                  MaxLength       =   50
                  TabIndex        =   109
                  Top             =   1080
                  Width           =   1905
               End
               Begin VB.TextBox TxtDiscount 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   6480
                  Locked          =   -1  'True
                  TabIndex        =   107
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ Ì«— ’‰›"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   20700
                  TabIndex        =   67
                  Top             =   2040
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ ﬂ«›Â «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   19140
                  TabIndex        =   66
                  Top             =   2040
                  Width           =   1680
               End
               Begin VB.TextBox TxtQty 
                  Alignment       =   2  'Center
                  BackColor       =   &H80000002&
                  Height          =   315
                  Left            =   13320
                  MaxLength       =   50
                  TabIndex        =   65
                  Top             =   1080
                  Width           =   705
               End
               Begin VB.TextBox TXTREMARKSDeails 
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
                  Height          =   675
                  Left            =   1920
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   64
                  Top             =   1440
                  Width           =   12930
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   3555
                  Index           =   1
                  Left            =   0
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   2160
                  Width           =   19065
                  _cx             =   33629
                  _cy             =   6271
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
                     Height          =   195
                     Left            =   19680
                     TabIndex        =   73
                     Top             =   1515
                     Visible         =   0   'False
                     Width           =   2295
                  End
                  Begin VB.TextBox Txtidxx 
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
                     Height          =   165
                     Index           =   0
                     Left            =   -4020
                     TabIndex        =   72
                     Top             =   3915
                     Width           =   2220
                  End
                  Begin VB.CheckBox Check1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ "
                     Height          =   180
                     Left            =   18810
                     TabIndex        =   71
                     Top             =   1485
                     Width           =   2325
                  End
                  Begin VB.TextBox txtType 
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
                     Height          =   105
                     Left            =   20085
                     TabIndex        =   70
                     Text            =   "0"
                     Top             =   1620
                     Visible         =   0   'False
                     Width           =   495
                  End
                  Begin VB.CheckBox ChKauto 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ì"
                     Enabled         =   0   'False
                     Height          =   150
                     Left            =   19380
                     TabIndex        =   69
                     Top             =   1440
                     Width           =   1530
                  End
                  Begin MSDataListLib.DataCombo dcopr 
                     Height          =   315
                     Left            =   19545
                     TabIndex        =   74
                     Top             =   825
                     Width           =   4140
                     _ExtentX        =   7303
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
                  Begin MSDataListLib.DataCombo dcproject 
                     Height          =   315
                     Left            =   19425
                     TabIndex        =   75
                     Top             =   675
                     Width           =   1635
                     _ExtentX        =   2884
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
                  Begin MSDataListLib.DataCombo Dcterm 
                     Height          =   315
                     Left            =   19890
                     TabIndex        =   76
                     Top             =   360
                     Width           =   3150
                     _ExtentX        =   5556
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
                  Begin VSFlex8Ctl.VSFlexGrid Grid 
                     Height          =   3135
                     Left            =   120
                     TabIndex        =   77
                     Top             =   240
                     Width           =   18690
                     _cx             =   32967
                     _cy             =   5530
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
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmProductionAllocation.frx":217D
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
                     Caption         =   "0"
                     Height          =   315
                     Index           =   31
                     Left            =   6480
                     TabIndex        =   103
                     Top             =   3480
                     Visible         =   0   'False
                     Width           =   1575
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·ﬂ„Ì… "
                     Height          =   312
                     Index           =   30
                     Left            =   8160
                     TabIndex        =   102
                     Top             =   3480
                     Visible         =   0   'False
                     Width           =   1608
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
                     Height          =   90
                     Left            =   14115
                     TabIndex        =   79
                     Top             =   450
                     Width           =   900
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "»œ«Ì… «· Œ’Ì’"
                     Height          =   135
                     Index           =   8
                     Left            =   18195
                     TabIndex        =   78
                     Top             =   1095
                     Width           =   1830
                  End
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   1065
                  TabIndex        =   80
                  Top             =   1680
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
                  ButtonImage     =   "FrmProductionAllocation.frx":2763
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   396
                  Index           =   21
                  Left            =   120
                  TabIndex        =   81
                  Top             =   1680
                  Width           =   912
                  _ExtentX        =   1614
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmProductionAllocation.frx":2AFD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo dcitemclass 
                  Height          =   315
                  Left            =   9240
                  TabIndex        =   82
                  Top             =   1080
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   -2147483646
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
               Begin MSDataListLib.DataCombo DDcunits1 
                  Height          =   315
                  Left            =   11160
                  TabIndex        =   100
                  Top             =   1080
                  Width           =   1320
                  _ExtentX        =   2328
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
               Begin MSDataListLib.DataCombo DcAccount1 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   3240
                  _ExtentX        =   5715
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  BackColor       =   -2147483646
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
                  Caption         =   "«÷€ÿ · Œ’Ì’ ”‰œ«  «·’—›"
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Index           =   38
                  Left            =   14760
                  TabIndex        =   121
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·–—«⁄"
                  Height          =   315
                  Index           =   35
                  Left            =   8760
                  TabIndex        =   111
                  Top             =   1080
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·√„—"
                  Height          =   315
                  Index           =   34
                  Left            =   16440
                  TabIndex        =   110
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰”»… «·Œ’„"
                  Height          =   315
                  Index           =   33
                  Left            =   7200
                  TabIndex        =   108
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ”«» «·Œ’„"
                  Height          =   315
                  Index           =   32
                  Left            =   5040
                  TabIndex        =   106
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÊÕœ…"
                  Height          =   315
                  Index           =   29
                  Left            =   12615
                  TabIndex        =   101
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„ÊŸ›"
                  Height          =   315
                  Index           =   1
                  Left            =   8400
                  TabIndex        =   86
                  Top             =   3570
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   315
                  Index           =   24
                  Left            =   13920
                  TabIndex        =   85
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›—“"
                  Height          =   315
                  Index           =   25
                  Left            =   10560
                  TabIndex        =   84
                  Top             =   1080
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   27
                  Left            =   15960
                  TabIndex        =   83
                  Top             =   1440
                  Width           =   975
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   315
               Index           =   4
               Left            =   6600
               TabIndex        =   118
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »  Œ’Ì’ ŒÿÊÿ «·√‰ «Ã ·√Ê«„— «·‘€· Ê«·ﬂ„Ì«  «·„‰ ÃÂ  "
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
               Height          =   540
               Index           =   44
               Left            =   360
               TabIndex        =   104
               Top             =   840
               Width           =   3975
            End
            Begin VB.Shape Shape2 
               BorderWidth     =   2
               Height          =   615
               Left            =   360
               Top             =   840
               Width           =   4005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   315
               Index           =   20
               Left            =   6960
               TabIndex        =   93
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·›—⁄"
               Height          =   315
               Index           =   10
               Left            =   12240
               TabIndex        =   92
               Top             =   840
               Width           =   720
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«· «—ÌŒ"
               Height          =   225
               Index           =   9
               Left            =   15000
               TabIndex        =   91
               Top             =   840
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               Height          =   225
               Index           =   7
               Left            =   17400
               TabIndex        =   90
               Top             =   840
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„œ Â« „‰"
               Height          =   270
               Index           =   5
               Left            =   4575
               TabIndex        =   89
               Top             =   480
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ï"
               Height          =   270
               Index           =   2
               Left            =   1785
               TabIndex        =   88
               Top             =   480
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   195
               Index           =   3
               Left            =   10410
               TabIndex        =   87
               Top             =   1395
               Visible         =   0   'False
               Width           =   945
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   570
         Left            =   1800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   9840
         Width           =   12825
         _cx             =   22622
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
            Left            =   10320
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
            ButtonImage     =   "FrmProductionAllocation.frx":3097
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   11205
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
            ButtonImage     =   "FrmProductionAllocation.frx":3431
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
            ButtonImage     =   "FrmProductionAllocation.frx":37CB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   492
            Index           =   0
            Left            =   8940
            TabIndex        =   9
            Top             =   36
            Width           =   732
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
            Height          =   492
            Index           =   1
            Left            =   8040
            TabIndex        =   10
            Top             =   36
            Width           =   852
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
            Height          =   492
            Index           =   2
            Left            =   7200
            TabIndex        =   11
            Top             =   36
            Width           =   768
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
            Height          =   492
            Index           =   3
            Left            =   6192
            TabIndex        =   12
            Top             =   36
            Width           =   768
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
            Height          =   492
            Index           =   4
            Left            =   5160
            TabIndex        =   13
            Top             =   0
            Width           =   768
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
            Height          =   492
            Index           =   6
            Left            =   2400
            TabIndex        =   14
            Top             =   36
            Width           =   768
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
            Height          =   492
            Index           =   5
            Left            =   4236
            TabIndex        =   15
            Top             =   36
            Width           =   768
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
            Left            =   12120
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
            MICON           =   "FrmProductionAllocation.frx":3B65
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   492
            Index           =   7
            Left            =   3360
            TabIndex        =   141
            Top             =   0
            Width           =   768
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Left            =   1560
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
            TabIndex        =   7
            Top             =   240
            Width           =   1515
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   15240
         TabIndex        =   139
         Top             =   9840
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   270
         Index           =   42
         Left            =   17985
         TabIndex        =   140
         Top             =   9915
         Width           =   900
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
      ButtonImage     =   "FrmProductionAllocation.frx":3B81
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmProductionAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Public LngRow As Double
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal x As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Function CuurentLogdata(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·« ›«ﬁÌ…    " & txtid.Text & CHR(13) & " «·⁄„»· " & DBCboClientName.Text & CHR(13) & "  „œ Â« „‰  " & dbFromDate & CHR(13) & "  «·Ï " & dbTodate & CHR(13) & "  „·«ÕŸ«  " & TxtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "   „ «Ìﬁ«› «· ⁄«„· "
    End If
                    
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Contract No    " & txtid.Text & CHR(13) & " Customer " & DBCboClientName.Text & CHR(13) & " From   " & dbFromDate & CHR(13) & "  To  " & dbTodate & CHR(13) & "  Remarks " & TxtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & CHR(13) & " Locked "
    End If
                    
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
    
End Function

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
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



MySQL = "SELECT     dbo.tblProductionAlloc.ID, dbo.tblProductionAlloc.OPrDate, dbo.tblProductionAlloc.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "                        dbo.tblProductionAlloc.StoreId, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.tblProductionAlloc.LineID, dbo.TblProductLine.name,"
MySQL = MySQL & "                       dbo.tblProductionAlloc.fromTime, dbo.tblProductionAlloc.toTime, dbo.tblProductionAlloc.NoOfHours, dbo.tblProductionAlloc.TotalSalaries,"
MySQL = MySQL & "                       dbo.tblProductionAlloc.TotalElectricals, dbo.tblProductionAlloc.WorkOrderNO, dbo.tblProductionAlloc.totalLineExpenses, dbo.tblProductionAlloc.totalOrderQty,"
MySQL = MySQL & "                       dbo.tblProductionAlloc.TotalProductionQty, dbo.tblProductionAlloc.totalMaterialsForItems, dbo.tblProductionAlloc.totalMaterials, dbo.tblProductionAlloc.REMARKS,"
MySQL = MySQL & "                       dbo.tblProductionAlloc.UsedPowerPriceH, dbo.tblProductionAlloc.Transaction_ID, dbo.tblProductionAlloc.NoteSerial, dbo.tblProductionAlloc.NoteSerial1,"
MySQL = MySQL & "                       dbo.tblProductionAlloc.ReciveDate, dbo.tblProductionAlloc.NoteSerial1V, dbo.tblProductionAllocDetails.Qty, dbo.tblProductionAllocDetails.NoteSerial AS NoteSerialDet,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.NoteSerial1 AS NoteSerial1Det, dbo.tblProductionAllocDetails.REMARKS AS REMARKSDet, dbo.tblProductionAllocDetails.Price,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.MaterialsValue, dbo.tblProductionAllocDetails.SalariesValue, dbo.tblProductionAllocDetails.LineExpensesValue,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.DiscountPercentage, dbo.tblProductionAllocDetails.DiscountValue, dbo.tblProductionAllocDetails.Cost,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.NProductionOrderNO, dbo.tblProductionAllocDetails.ClassId, dbo.TblItemsclasses.SizeName,  dbo.TblItemsclasses.SizeNameE,   "
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.StoreID AS StoreIDDet, TblStore_1.StoreName AS StoreNameDet, TblStore_1.StoreNamee AS StoreNameDetE,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.tblProductionAllocDetails.itemid, dbo.TblItems.ItemName,"
MySQL = MySQL & "                       dbo.TblItems.ItemNamee, dbo.tblProductionAllocDetails.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.tblProductionAllocDetails.Account_Code1,"
MySQL = MySQL & "                       ACCOUNTS_1.Account_Name AS Account_Name2, dbo.tblProductionAllocDetails.Alarm, dbo.tblProductionAllocDetails.fromdate, dbo.tblProductionAllocDetails.todate,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.fromTime AS fromTimeDet, dbo.tblProductionAllocDetails.toTime AS toTimeDet, dbo.tblProductionAllocDetails.OverHead,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.StrSand, dbo.tblProductionAllocDetails.totalss, dbo.tblProductionAllocDetails.StrSelectSands, dbo.tblProductionAllocDetails.hours,"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails.ElectricExpenses , dbo.tblProductionAllocDetails.gasExpenses , dbo.TblItems.Fullcode "
MySQL = MySQL & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ACCOUNTS_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.tblProductionAllocDetails ON ACCOUNTS_1.Account_Code = dbo.tblProductionAllocDetails.Account_Code1 LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.tblProductionAllocDetails.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblItems ON dbo.tblProductionAllocDetails.itemid = dbo.TblItems.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblUnites ON dbo.tblProductionAllocDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblStore TblStore_1 ON dbo.tblProductionAllocDetails.StoreID = TblStore_1.StoreID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblItemsclasses ON dbo.tblProductionAllocDetails.ClassId = dbo.TblItemsclasses.SizeId RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.tblProductionAlloc ON dbo.tblProductionAllocDetails.ID = dbo.tblProductionAlloc.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblProductLine ON dbo.tblProductionAlloc.LineID = dbo.TblProductLine.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblStore ON dbo.tblProductionAlloc.StoreId = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.tblProductionAlloc.BranchID"
MySQL = MySQL & "  Where (dbo.tblProductionAlloc.id =" & val(txtid.Text) & ")"
 


  If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\REPORTS\REPORTS NEW\RepProductAllocation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\RepProductAllocationE.rpt"
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
        Msg = "No Data "
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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

End Function
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

Function Create_dev()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
          MsgBox "Branch wasn't created", vbCritical
        End If
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
             MsgBox "Salaries Account for employees wasn't created for this branch ", vbCritical
            End If
            
            GoTo ErrTrap
         
        End If
    End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "
    Else
    Msg = "Salaries Entry for employees  for Month  " & "   Year "
    End If

    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = Null
    rs("Remark").value = Msg

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = user_id
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With
 
 If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    Else
    MsgBox "ıEntry wasn't created", vbInformation
    End If
    
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
    Else
    MsgBox "An error occur while saving data", vbExclamation
    End If
  
End Function
Function create_report_data()

End Function
Function Create_dev1()
  
    
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    dcproject.Text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub

Private Sub CmdOk_Click()
    CmdOk_Click
End Sub



Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .Rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
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

    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub



Private Sub Del_Trans()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If txtid.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
  Msg = "Process No Will be Deleted " & CHR(13)
        Msg = Msg + (txtid.Text) & CHR(13)
        Msg = Msg + " are oyu sure to delete this process"

End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
DeleteTransactiomsVoucher val(txtTransaction_ID.Text)

                Cn.Execute "delete tblProductionAllocDetails where id=" & val(Me.txtid.Text)
                 Cn.Execute "delete TblProductionAllocDetails1 where ProID=" & val(Me.txtid.Text)
                
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '   XPTxtCurrent.Caption = 0
                    '   XPTxtCount.Caption = 0
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
         Msg = "This process Not Allowed Where there is no data"
        End If
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then '
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        Else
         Msg = "you can't delete this record as data integrity " & CHR(13) & "there is data connect with supplier "
        End If
        
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub
Function checkVouvher() As Boolean
Dim Vchr_result As String
       Vchr_result = Voucher_coding(val(my_branch), ReciveDate.value, 19, 250, , 28, , val(Me.DCStores.BoundText))

                If Vchr_result = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «” ·«„ „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
                    Else
                    MsgBox "can't add new bill receipt as you exceed receipts limit ": Exit Function
                    End If
                    
               checkVouvher = False
                Else
                       
                    If Vchr_result = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
                    Else
                    MsgBox "You must Write Receipt No Manual  ": Exit Function
                    End If
                        
                        checkVouvher = False
                    Else
                        TxtNoteSerial1V = Vchr_result
                        checkVouvher = True
                    End If
                End If
End Function

Function createReciveVoucher(NoteID As Double, NoteSerial As String)
Dim Transaction_ID As String
Dim Vchr_result As String
DeleteTransactiomsVoucher val(txtTransaction_ID.Text)

'If txtTransaction_ID.text = "" Then
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
txtTransaction_ID.Text = Transaction_ID
'Else

'End If

        If TxtNoteSerial1V = "" Then
      Vchr_result = Voucher_coding(val(DCranch.BoundText), ReciveDate.value, 19, 250, , 28, , val(DCStores.BoundText))
      TxtNoteSerial1V = Vchr_result
           
            End If
   
  '*****************************************************************************************************************************************************************Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,StoreID,UserID,NoteSerial1,BranchId
          Cn.Execute "INSERT INTO  Transactions ( Transaction_ID,Transaction_Serial ,Transaction_Date,Transaction_Type ,StoreID,UserID,NoteSerial1,BranchId,NoteSerial)SELECT " & Transaction_ID & ",0," & SQLDate(ReciveDate.value, True) & ", 28,StoreID,UserID" & ",NoteSerial1='" & TxtNoteSerial1V & "',BranchID ," & TxtNoteSerial.Text & "  From tblProductionAlloc Where ID =" & val(txtid.Text)
        '
        '****************************************************************************************************************************************************************************** C         (showPrice,Transaction_ID,Item_ID,ItemCase,Quantity,Price,ColorID,itemsize,UnitId,ShowQty,NProductionOrderNO,classid)
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,Transaction_ID,Item_ID,ItemCase,Quantity,Price,ColorID,itemsize,UnitId,ShowQty,NProductionOrderNO,classid)SELECT   (cost/qty) ," & Transaction_ID & ",itemid       ,1 , Qty,     (cost/qty ), 1,1, UnitID, qty,NProductionOrderNO,classid From dbo.tblProductionAllocDetails Where id = " & val(txtid.Text)
       
       rs!Transaction_ID = Transaction_ID
        rs!NoteSerial1V = TxtNoteSerial1V
      '  rs!NoteSerial = NoteSerial
        
        rs.update
    Dim StrSQL As String
    StrSQL = "update notes set noteseial1=" & val(TxtNoteSerial1V) & ",Remark='" & TxtNoteSerial1V & "',Transaction_ID=" & val(txtTransaction_ID.Text) & " where NoteID=" & NoteID
     
End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String

If SystemOptions.UserInterface = ArabicInterface Then
des = "”‰œ  Œ’Ì’ «‰ «Ã —ﬁ„  " & txtid & " » «—ÌŒ " & dbOPrDate.value
Else
des = "Prodaction Allocation Receipt No.  " & txtid & " With Date  " & dbOPrDate.value
End If

Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "tblProductionAlloc"
Filedname = "ID"
ContNo = txtid
Notevalue = 0


                     If Me.TxtModFlg = "N" Then
                                 CreateNotes NoteID, (ReciveDate.value), val(DCranch.BoundText), 250, Notevalue, NoteSerial, "", tablename, Filedname, ContNo, des, ToHijriDate((ReciveDate.value))
                                     TxtNoteID.Text = NoteID
                                    TxtNoteSerial.Text = NoteSerial
                    Else
                                      If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                    CreateNotes NoteID, (ReciveDate.value), val(DCranch.BoundText), 250, Notevalue, NoteSerial, "", tablename, Filedname, ContNo, des, ToHijriDate((ReciveDate.value))
                                                       TxtNoteID.Text = NoteID
                                                  TxtNoteSerial.Text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",NoteSerial1='" & "" & "'"
                                                    sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
ReLineGrid
createReciveVoucher val(TxtNoteID.Text), TxtNoteSerial.Text
CREATE_VOUCHER_GE val(TxtNoteID.Text), val(DCranch.BoundText), user_id, ReciveDate.value
rs.Resync adAffectCurrent


End Function



Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Double
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

 
   
  
            StrTempDes = "”‰œ «” ·«„ «‰ «Ã    " & "" & "  »‰«¡ ⁄·Ï ”‰œ  Œ’Ì’ »—ﬁ„   " & txtid.Text
            LngDevNO = LngDevNO + 1
 
Notevalue = 0
 
 Dim Account_Code_dynamic37 As String '„Ê«œ
  Dim Account_Code_dynamic38 As String '⁄„«·Â
   Dim Account_Code_dynamic39 As String ',ÊﬁÊœ
    Dim Account_Code_dynamic79 As String ' ﬂÂ—»«¡
        
   Account_Code_dynamic37 = get_account_code_branch(37, my_branch)
   Account_Code_dynamic38 = get_account_code_branch(38, my_branch)
   Account_Code_dynamic39 = get_account_code_branch(39, my_branch)
   Account_Code_dynamic79 = get_account_code_branch(79, my_branch)
           
            
'll:
   LngDevNO = 0
  
'****************************************************************************************
Dim sql As String

    With Me.Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemID")) <> "" Then
                         If val(.TextMatrix(i, .ColIndex("Price"))) > 0 Then
                                
                               Notevalue = val(.TextMatrix(i, .ColIndex("Price")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = get_store_Account(DCStores.BoundText, "Account_Code")
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                            
                        
                                    
                                    
                           End If
  
                           If val(.TextMatrix(i, .ColIndex("MaterialsValue"))) > 0 Then
                                
                                Notevalue = val(.TextMatrix(i, .ColIndex("MaterialsValue")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = Account_Code_dynamic37
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                            
                        
                                    
                                    
                           End If
                           
                     
                     
                     
                         If val(.TextMatrix(i, .ColIndex("SalariesValue"))) > 0 Then
                                
                               Notevalue = val(.TextMatrix(i, .ColIndex("SalariesValue")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = Account_Code_dynamic38
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                            
                        
                                    
                                    
                           End If
                           
                     
                                    If val(.TextMatrix(i, .ColIndex("gasExpenses"))) > 0 Then
                                
                               Notevalue = val(.TextMatrix(i, .ColIndex("gasExpenses")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = Account_Code_dynamic39
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                            
                        
                                    
                                    
                           End If
                           
                           
                                   If val(.TextMatrix(i, .ColIndex("ElectricExpenses"))) > 0 Then
                                
                               Notevalue = val(.TextMatrix(i, .ColIndex("ElectricExpenses")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = Account_Code_dynamic79
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                            
                        
                                    
                                    
                           End If
                           
                           
                     
                     
           If val(.TextMatrix(i, .ColIndex("DiscountValue"))) > 0 And (.TextMatrix(i, .ColIndex("Account2"))) <> "" Then
                                
                                
                               Notevalue = val(.TextMatrix(i, .ColIndex("DiscountValue")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = (.TextMatrix(i, .ColIndex("Account2")))
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                            
                        
                      Notevalue = val(.TextMatrix(i, .ColIndex("DiscountValue")))
                           LngDevNO = LngDevNO + 1
                           StrTempAccountCode = get_store_Account(DCStores.BoundText, "Account_Code")
            
                          
                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                        GoTo ErrTrap
                                    
                                    
                                    End If
                                    
                                LngDevNO = LngDevNO + 1
                                
       
          
          End If
                    
             
             
             
            '*/**********************************************
            StrSQL = "SELECT     TOP 100 PERCENT dbo.TblDistriExpensItemDet3.Vlue AS TOTAL, dbo.ACCOUNTS.Account_Code"
 StrSQL = StrSQL & "  FROM         dbo.TblDistriExpensItemDet2 LEFT OUTER JOIN"
 StrSQL = StrSQL & "   dbo.TblDistriExpensItemDet3 ON dbo.TblDistriExpensItemDet2.ID = dbo.TblDistriExpensItemDet3.IDDet LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.ACCOUNTS ON REPLACE(REPLACE(dbo.TblDistriExpensItemDet3.Account_Code, CHAR(10), ''), CHAR(13), '') = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL & "   Where (dbo.TblDistriExpensItemDet2.ItemID = " & val(.TextMatrix(i, .ColIndex("ItemID"))) & ")"
StrSQL = StrSQL & "   ORDER BY dbo.TblDistriExpensItemDet2.ItemID"
               
   
    Dim j As Integer
    Dim RsDev As ADODB.Recordset
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
           
   
            For j = 1 To RsDev.RecordCount
            
                                 
                               Notevalue = IIf(IsNull(RsDev("TOTAL").value), 0, RsDev("TOTAL").value) * val(.TextMatrix(i, .ColIndex("Qty")))
                                 StrTempAccountCode = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                                 
                               If Notevalue > 0 And ModAccounts.check_account_exist(StrTempAccountCode) = True Then
                                       LngDevNO = LngDevNO + 1
                                       'StrTempAccountCode = (.TextMatrix(i, .ColIndex("Account2")))
                        
                                      
                                                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "   ﬁÌ„… «·„Œ“Ê‰    ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                                    GoTo ErrTrap
                                                
                                                
                                                End If
                                                
                                            LngDevNO = LngDevNO + 1
                                
                                End If
                                
           
           
RsDev.MoveNext
            Next j
        End With
        
       End If
        
 RsDev.Close
 Set RsDev = Nothing
             
             
             End If
            
            '
        Next i

    End With

ErrTrap:
  
 End Function
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
 Dim nElements As Integer
Dim RsDetails2 As ADODB.Recordset
  Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
  Dim st As String
  '  On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
If checkVouvher = False Then Exit Sub
                If DCStores.Text = "" Then
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
                            Else
                            Msg = "Specify Store"
                            End If
                
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DCStores.SetFocus
                    SendKeys "{F4}"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
    
 
    End If


 
    If DCranch.BoundText = "" Then
                             If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "ÌÃ»  ÕœÌœ «·›—⁄"
                            Else
                            Msg = "Specify Branch"
                            End If
                
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DCranch.SetFocus
                    SendKeys "{F4}"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
    
 
    
    
    
    
    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
                Me.txtid.Text = CStr(new_id("tblProductionAlloc", "id", "", True))

        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete tblProductionAllocDetails where id=" & val(Me.txtid.Text)

          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords


    End If
    
       rs("id").value = txtid.Text
    
       rs("OPrDate").value = dbOPrDate.value
       rs("BranchID").value = IIf(Me.DCranch.BoundText = "", Null, Me.DCranch.BoundText)
       rs("LineID").value = IIf(Me.dcLineID.BoundText = "", Null, Me.dcLineID.BoundText)
       rs("StoreID").value = IIf(Me.DCStores.BoundText = "", Null, Me.DCStores.BoundText)
       
     rs("UserID").value = user_id
    
     rs("TotalSalaries").value = IIf(Not IsNumeric(Me.TxtTotalSalaries.Text), Null, Me.TxtTotalSalaries.Text)
     rs("TotalElectricals").value = IIf(Not IsNumeric(Me.TxtTotalElectricals.Text), Null, Me.TxtTotalElectricals.Text)
       rs("UsedPowerPriceH").value = IIf(Not IsNumeric(Me.TXTUsedPowerPriceH.Text), Null, Me.TXTUsedPowerPriceH.Text)
   
     rs("totalLineExpenses").value = IIf(Not IsNumeric(Me.TxttotalLineExpenses.Text), Null, Me.TxttotalLineExpenses.Text)
     
     rs("TotalProductionQty").value = IIf(Not IsNumeric(Me.TxtTotalProductionQty.Text), Null, Me.TxtTotalProductionQty.Text)
     rs("totalMaterialsForItems").value = IIf(Not IsNumeric(Me.TxttotalMaterialsForItems.Text), Null, Me.TxttotalMaterialsForItems.Text)
     rs("totalMaterials").value = IIf(Not IsNumeric(Me.TxtTotalMaterials.Text), Null, Me.TxtTotalMaterials.Text)
     
             rs("fromTime").value = FormatDateTime(Me.DBfromTime.value, vbShortTime)
        rs("toTime").value = FormatDateTime(Me.DBtoTime.value, vbShortTime)
        
         rs("NoOfHours").value = IIf((Me.TxtNoOfHours.Text) = "", Null, Me.TxtNoOfHours.Text)
         rs("WorkOrderNO").value = IIf((Me.TxtWorkOrderNO.Text) = "", Null, Me.TxtWorkOrderNO.Text)
       rs("ReciveDate").value = ReciveDate.value
         
         
   
    rs("customerid").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)

rs("totalOrderQty").value = IIf(Not IsNumeric(Me.TxttotalOrderQty.Text), Null, Me.TxttotalOrderQty.Text)

      
    rs("Remarks").value = IIf(Me.TxtRemarks.Text = "", "", Me.TxtRemarks.Text)
 
     rs("ItemID").value = IIf(Me.DCItemID.BoundText = "", Null, Me.DCItemID.BoundText)

   rs("UnitID").value = IIf(Me.DDcunits.BoundText = "", Null, Me.DDcunits.BoundText)
 
 

    rs.update
    Set RsDetails2 = New ADODB.Recordset
    RsDetails2.Open "TblProductionAllocDetails1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "tblProductionAllocDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer
 


Dim sql As String

    With Me.Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemID")) <> "" Then
            If Me.TxtModFlg.Text = "E" Then
        '«Œ Ì«— ”‰œ«  «·„Ê«œ «·Œ«„
            sql = " update Transaction_Details set flag=0 where ProIDdet= " & val(.TextMatrix(i, .ColIndex("id"))) & ""
            
        Cn.Execute sql

    End If
                RsDev.AddNew
                RsDev("id").value = Me.txtid.Text
            
                RsDev("ItemId").value = val(.TextMatrix(i, .ColIndex("ItemID")))
                 RsDev("totalss").value = val(.TextMatrix(i, .ColIndex("totalss")))
                  RsDev("StrSelectSands").value = .TextMatrix(i, .ColIndex("sandat"))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                RsDev("Alarm").value = IIf((.TextMatrix(i, .ColIndex("Alarm"))) = "", Null, (.TextMatrix(i, .ColIndex("Alarm"))))
                
                'TxtAlarm
                
                  RsDev("StrSand").value = .TextMatrix(i, .ColIndex("sand"))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("Qty").value = val(.TextMatrix(i, .ColIndex("Qty")))
                RsDev("ClassId").value = val(.TextMatrix(i, .ColIndex("ClassId")))
                RsDev("StoreID").value = val(.TextMatrix(i, .ColIndex("StoreID")))
                RsDev("REMARKS").value = val(.TextMatrix(i, .ColIndex("REMARKS")))
                
                
 
                
                  RsDev("MaterialsValue").value = val(.TextMatrix(i, .ColIndex("MaterialsValue")))
                    RsDev("SalariesValue").value = val(.TextMatrix(i, .ColIndex("SalariesValue")))
                      RsDev("LineExpensesValue").value = val(.TextMatrix(i, .ColIndex("LineExpensesValue")))
                      
                        RsDev("gasExpenses").value = val(.TextMatrix(i, .ColIndex("gasExpenses")))
                          RsDev("ElectricExpenses").value = val(.TextMatrix(i, .ColIndex("ElectricExpenses")))
                             
                            
                        RsDev("DiscountPercentage").value = val(.TextMatrix(i, .ColIndex("DiscountPercentage")))
                          RsDev("discountvalue").value = val(.TextMatrix(i, .ColIndex("discountvalue")))
                            RsDev("cost").value = val(.TextMatrix(i, .ColIndex("cost")))
                            RsDev("OverHead").value = val(.TextMatrix(i, .ColIndex("OverHead")))
                            
                            RsDev("Account_Code").value = (.TextMatrix(i, .ColIndex("Account1")))
                            RsDev("Account_Code1").value = (.TextMatrix(i, .ColIndex("Account2")))
                
                RsDev("NProductionOrderNO").value = (.TextMatrix(i, .ColIndex("NProductionOrderNO")))
                  






                 RsDev("hours").value = (.TextMatrix(i, .ColIndex("hours")))
                 RsDev("fromdate").value = (.TextMatrix(i, .ColIndex("fromdate")))
                 RsDev("todate").value = (.TextMatrix(i, .ColIndex("todate")))
                 RsDev("fromTime").value = (.TextMatrix(i, .ColIndex("fromTime")))
                 RsDev("toTime").value = (.TextMatrix(i, .ColIndex("toTime")))
                
                
                
                
                RsDev.update
           If .TextMatrix(i, .ColIndex("sandat")) <> "" Then
          st = .TextMatrix(i, .ColIndex("sandat"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails2.AddNew
        
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails2("ProID").value = val(txtid.Text)
         RsDetails2("ProIDdet").value = IIf(IsNull(RsDev("id2").value), Null, RsDev("id2").value)
         RsDetails2("NProductionOrderNO").value = astrSplit2tems2(0)
         RsDetails2("Transaction_ID").value = val(astrSplit2tems2(1))
         RsDetails2("Selcted").value = val(astrSplit2tems2(2))
         RsDetails2("NoteSerial1").value = astrSplit2tems2(3)
         RsDetails2("Transaction_Date").value = astrSplit2tems2(4)
         RsDetails2("total").value = val(astrSplit2tems2(5))
         RsDetails2("idd").value = val(astrSplit2tems2(6))
         If val(astrSplit2tems2(2)) = 1 Then
        StrSQL = " update Transaction_Details set flag =1 where id=" & val(astrSplit2tems2(6)) & ""
         Cn.Execute StrSQL
         StrSQL = " update Transaction_Details set ProIDdet =" & val(IIf(IsNull(RsDev("id2").value), Null, RsDev("id2").value)) & " where id=" & val(astrSplit2tems2(6)) & ""
         Cn.Execute StrSQL
        End If
         RsDetails2.update
         Next j
                  
          
          End If
                    
            End If
            
            '
        Next i

    End With
 
    RsDev.Close
    'save Groups
 
    
    Cn.CommitTrans
    BeginTrans = False
    createVoucher
      updateNotesValueAndNobytext (val(TxtNoteID.Text))
 
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
Else
        Msg = " Process Saved Successfully " & CHR(13)
            Msg = Msg + "do you want to add another data"
End If
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If
Retrive val(Me.txtid.Text)
        Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
             MsgBox "Updates Saved Successfully ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            
            '  Fg_Journal.Enabled = False
            Retrive val(Me.txtid.Text)
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
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
        Msg = "Can't save data " & CHR(13)
        Msg = Msg + "Invalid data entered " & CHR(13)
        Msg = Msg + "Pleae try again"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Cmd_Click(Index As Integer)
' On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
       
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Grid.Enabled = True
            Option2.value = True
 DCranch.BoundText = Current_branch
 DCboUserName.BoundText = user_id
ReciveDate.value = Date
        Case 1
                    If ChekClodePeriod(dbOPrDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
            DCboUserName.BoundText = user_id
            CuurentLogdata

        Case 2
                    If ChekClodePeriod(dbOPrDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
    ReLineGrid
            SaveData
           
        Case 3
            Undo

        Case 4
                    If ChekClodePeriod(dbOPrDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            Load FrmSearchProdAllocation
             FrmSearchProdAllocation.show vbModal
          

        Case 6
            Unload Me

        Case 7
         
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

           'sa If val(Me.TxtTblVendorContractD.text) <> 0 Then
                print_report
        
        
           ' End If
    
        Case 8
          '  RemoveGridRowGroup
    
        Case 20
            addrow
ReLineGrid
        Case 21
            RemoveGridRow
            ReLineGrid
    End Select

    Exit Sub
ErrTrap:

End Sub



Private Sub RemoveGridRow()
Dim Msg As String
    With Me.Grid
If Me.TxtModFlg.Text <> "R" Then
        If .Row <= 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("ItemId")) <> "" Then
             
       If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› ”‰œ«  «· Œ’Ì’  —ﬁ„ " & CHR(13)
        Msg = Msg + .TextMatrix(.Row, .ColIndex("sand")) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Daa " & CHR(13)
        Msg = Msg + .TextMatrix(.Row, .ColIndex("sand")) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
         If .TextMatrix(.Row, .ColIndex("sandat")) <> "" Then
        StrSQL = " update Transaction_Details set flag =0 where ProIDdet=" & val(.TextMatrix(.Row, .ColIndex("id"))) & ""
        Cn.Execute StrSQL
         End If
         
         .RemoveItem .Row
         
         Else
         Exit Sub
         End If
         End If
      End If
       
    End With

    ReLineGrid
End Sub

Function addrow()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer

 
        If DCItemID.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— «·’‰›  «·„‰ Ã ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If


        If DDcunits1.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«—  ÊÕœ… «·’‰›  «·„‰ Ã ...!!!"
            Else
                Msg = "must Specify item unit Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If



        If dcitemclass.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«—  ›—“ «·’‰›  «·„‰ Ã ...!!!"
            Else
                Msg = "must Specify item class Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If




        If DcAccount1.BoundText = "" And val(TxtDiscount.Text) > 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«—  Õ”«» «·Œ’„ ...!!!"
            Else
                Msg = "must Specify item class Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If
        
        


        If DCStores.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«—  „Œ“‰  «·’‰›  «·„‰ Ã ...!!!"
            Else
                Msg = "must Specify item Store Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If
        
        
        If val(TxtQty.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»   «œŒ«· «·ﬂ„Ì… «·„‰ Ã…  «·„‰ Ã ...!!!"
            Else
                Msg = "must Specify item price   ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If
                
                
        wherestr = "  where ItemID= " & val(DCItemID.BoundText)
 

    sql = "Select * from TblItems "

    If wherestr <> "" Then
        sql = sql & wherestr
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
 
    With Grid
 
        lastrow = .Rows
    
        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + lastrow
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                LngItemID = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                       
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs3.Fields("ItemCode").value), "", Rs3.Fields("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemName").value), "", Rs3.Fields("ItemName").value)
                       
      
                    .TextMatrix(i, .ColIndex("UnitId")) = val(DDcunits1.BoundText)
                    .TextMatrix(i, .ColIndex("UnitName")) = DDcunits1.Text
               
               .TextMatrix(i, .ColIndex("Alarm")) = TxtAlarm.Text
               
               'TxtAlarm
               
          
             .TextMatrix(i, .ColIndex("StoreID")) = val(DCStores.BoundText)
                    .TextMatrix(i, .ColIndex("StoreName")) = DCStores.Text
                    
                       
            .TextMatrix(i, .ColIndex("ClassId")) = val(dcitemclass.BoundText)
                    .TextMatrix(i, .ColIndex("classname")) = dcitemclass.Text
                                          
                                     .TextMatrix(i, .ColIndex("Qty")) = val(TxtQty.Text)
                                     '.TextMatrix(i, .ColIndex("Price")) = val(TxtPrice.text)
                                     
         .TextMatrix(i, .ColIndex("REMARKS")) = (TXTREMARKSDeails.Text)
             .TextMatrix(i, .ColIndex("NProductionOrderNO")) = (TxtWorkOrderNOSub.Text)
             
           
             
                  '.TextMatrix(i, .ColIndex("MaterialsValue")) = val(TxttotalMaterialsForItems.Text)
            '   .TextMatrix(i, .ColIndex("MaterialsValue")) = Round(GetProductionTotalIssue(TxtWorkOrderNOSub.text), 2)
                  .TextMatrix(i, .ColIndex("SalariesValue")) = val(TxtTotalSalariesaLL.Text)
             '     .TextMatrix(i, .ColIndex("LineExpensesValue")) = val(TxtTotalElectricalsaLL.text)
                 .TextMatrix(i, .ColIndex("DiscountPercentage")) = val(TxtDiscount.Text)
                 
         '        .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("MaterialsValue"))) + val(TxtTotalSalariesaLL.Text) + val(TxtTotalElectricalsaLL.Text)
                 
         '        .TextMatrix(i, .ColIndex("DiscountValue")) = val(TxtDiscount.Text) * .TextMatrix(i, .ColIndex("Price")) / 100
                 
         '        .TextMatrix(i, .ColIndex("Cost")) = val(.TextMatrix(i, .ColIndex("Price"))) - val(.TextMatrix(i, .ColIndex("DiscountValue")))
                 If DcAccount1.BoundText <> "" And val(TxtDiscount.Text) > 0 Then
                                        .TextMatrix(i, .ColIndex("Account2")) = (DcAccount1.BoundText)
                                          .TextMatrix(i, .ColIndex("Account2name")) = (DcAccount1.Text)
              End If
 
             .TextMatrix(i, .ColIndex("hours")) = (TxtNoOfHours.Text)
            .TextMatrix(i, .ColIndex("Fromdate")) = Fromdate.value
            .TextMatrix(i, .ColIndex("todate")) = todate.value
            .TextMatrix(i, .ColIndex("fromTime")) = DBfromTime.value
            .TextMatrix(i, .ColIndex("toTime")) = DBtoTime.value
              .TextMatrix(i, .ColIndex("gasExpenses")) = val(TXTTotalUsedPowerPriceH.Text)
               .TextMatrix(i, .ColIndex("ElectricExpenses")) = val(TxtTotalElectricalsaLL.Text)
               'ma   .TextMatrix(i, .ColIndex("OverHead")) = GetOverHeadForItems(val(.TextMatrix(i, .ColIndex("ItemId"))))
                 
               'GetOverHeadForItems
             '
             '              .TextMatrix(i, .ColIndex("gasExpenses")) = Round(val(TXTTotalUsedPowerPriceH.text) / val(TxttotalOrderQty.text), 2)
             '           .TextMatrix(i, .ColIndex("ElectricExpenses")) = Round(val(TxtTotalElectricalsaLL.text) / val(TxttotalOrderQty.text), 2)
                        
                        '.TextMatrix(i, .ColIndex("LineExpensesValue")) = Round(val(TxtTotalElectricalsaLL.text) / val(TxtTotalProductionQty.text), 2)
                        
             '               .TextMatrix(i, .ColIndex("SalariesValue")) = Round(val(TxtTotalSalariesaLL.text) / val(TxttotalOrderQty.text), 2)
                            .TextMatrix(i, .ColIndex("LineExpensesValue")) = Round(val(.TextMatrix(i, .ColIndex("gasExpenses"))) + val(.TextMatrix(i, .ColIndex("ElectricExpenses"))) + val(.TextMatrix(i, .ColIndex("SalariesValue"))), 2) 'Round(val(TxtTotalElectricalsaLL.text) / val(TxttotalOrderQty.text), 2)
                      'DiscountPercentage
                            
                            .TextMatrix(i, .ColIndex("Price")) = Round(val(.TextMatrix(i, .ColIndex("MaterialsValue"))) + val(.TextMatrix(i, .ColIndex("LineExpensesValue")) + val(.TextMatrix(i, .ColIndex("OverHead")))), 2)
                            
                          .TextMatrix(i, .ColIndex("DiscountValue")) = Round(val(.TextMatrix(i, .ColIndex("DiscountPercentage"))) * .TextMatrix(i, .ColIndex("Price")) / 100, 2)
                            
                       .TextMatrix(i, .ColIndex("Cost")) = Round(val(.TextMatrix(i, .ColIndex("Price"))) - val(.TextMatrix(i, .ColIndex("DiscountValue"))), 2)


                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close
TxtQty.Text = ""
dcitemclass.BoundText = ""
'DCStoreS.BoundText = ""
TXTREMARKSDeails.Text = ""
DcAccount1.BoundText = ""

    ReLineGrid


   
 
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
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If Rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If Rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If Rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If Rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If Rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If Rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If Rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If Rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If Rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If Rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If Rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If Rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If Rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If Rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If Rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If Rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If Rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If Rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

Private Sub CmdRemove_Click()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If x = vbNo Then Exit Sub
    
    If Grid.Rows > 1 Then
        If Grid.Rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.Rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub



 

Private Sub CnsSHowGl_Click()
If SystemOptions.UserInterface = ArabicInterface Then
       
  Else
        If val(TxtNoteSerial) = 0 Then MsgBox " VChr Not Exist ": Exit Sub


End If
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (TxtNoteSerial)
End Sub

Private Sub Command4_Click()
    Dim Transaction_ID As Integer
    Transaction_ID = Me.txtTransaction_ID.Text
    
If SystemOptions.UserInterface = ArabicInterface Then
    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    Else
    If Transaction_ID = 0 Then MsgBox "Receipt not recorded ": Exit Sub
End If
    
    FrmInpoutWorkOrder.show
    FrmInpoutWorkOrder.Retrive (Transaction_ID)

End Sub

Private Sub dbOPrDate_Change()
If Me.TxtModFlg <> "R" Then
ReciveDate.value = dbOPrDate.value
End If

End Sub

Private Sub DcAccount1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 195
            
    End If

End Sub

Private Sub DBfromTime_Change()

    If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
        TxtNoOfHours.Text = CalculateTimes(Me.DBfromTime.value, Me.DBtoTime.value)
    End If


End Sub

 

Private Sub DBtoTime_Change()
    If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
        TxtNoOfHours.Text = CalculateTimes(Me.DBfromTime.value, Me.DBtoTime.value)
    End If
End Sub

Private Sub dcitemclass_Change()
Dim Account_Code As String
TxtDiscount.Text = getClassInformations(val(dcitemclass.BoundText), , , , Account_Code)
DcAccount1.BoundText = Account_Code
End Sub
 
Private Sub dcitemclass_Click(Area As Integer)
dcitemclass_Change
End Sub

Private Sub dcLineID_Change()
If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then

 add_line (val(Me.dcLineID.BoundText))
 
End If
End Sub

Private Sub dcLineID_Click(Area As Integer)
dcLineID_Change
End Sub

Private Sub dcproject_Click(Area As Integer)
    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub DCStores_Change()
 TxtStoreID.Text = getStoreCoding(val(DCStores.BoundText))
End Sub

Private Sub Dcterm_Click(Area As Integer)
    If Dcterm.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
End Sub

Function add_line(ID As Integer)
    On Error Resume Next
    Dim LngRow As Long
    
    Dim sql As String
 
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sql = "select * from TblProductLine where id=" & ID

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then Exit Function
    
    If Me.dcLineID.BoundText = "" Then Exit Function


 
       TxtTotalElectricals.Text = IIf(Not IsNumeric(rs("UsedElectricPriceH").value), 0, rs("UsedElectricPriceH").value)
         TXTUsedPowerPriceH.Text = IIf(Not IsNumeric(rs("UsedPowerPriceH").value), 0, rs("UsedPowerPriceH").value)
         
       TxtTotalSalaries.Text = IIf(Not IsNumeric(rs("WorkerPriceH").value), 0, rs("WorkerPriceH").value)
    
    
    TxttotalLineExpenses.Text = IIf(Not IsNumeric(rs("LinePriceH").value), 0, rs("LinePriceH").value)
 
    
        Dim Hour As Integer
        Dim Minute As Double
        Dim totalhour As Double
        Hour = val(Mid(Me.TxtNoOfHours.Text, 1, 2))
        Minute = val(Mid(Me.TxtNoOfHours.Text, 4, 2)) / 60
        totalhour = Round(Hour + Minute, 2)
  totalhour = val(Me.TxtNoOfHours.Text)
        TxtTotalElectricalsaLL.Text = TxtTotalElectricals.Text * totalhour
  TxtTotalSalariesaLL.Text = TxtTotalSalaries.Text * totalhour
    TxttotalLineExpensesaLL.Text = TxttotalLineExpenses.Text * totalhour
   TXTTotalUsedPowerPriceH = TXTUsedPowerPriceH * totalhour
     

     
End Function

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With Fg1

  
               Cancel = True
  

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
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

 
    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL
    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
 
 
Dcombos.GetBranches Me.DCranch
 Dcombos.GetLine Me.dcLineID
 Dcombos.GetStores Me.DCStores
 
     Dcombos.GetItemsUnits Me.DDcunits
     Dcombos.GetItemsUnits Me.DDcunits1
     
    Dcombos.GetItemsNames DCItemID, , , , True
    
        Dcombos.GetItemsClasses Me.dcitemclass
    
  If SystemOptions.UserInterface = ArabicInterface Then
        Dcombos.GetAccountingCodes DcAccount1, True
      
    Else
 
        Dcombos.GetAccountingCodesENg DcAccount1, True
       
    End If




    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From tblProductionAlloc  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
  '  Cmd(2).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
Cmd(7).Caption = "Print"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Production Allocation"
    lbl(39).Caption = "Recive Date"
     lbl(40).Caption = "R.V No."
 lbl(41).Caption = "GE No."
 Command4.Caption = "Show No"
 CnsSHowGl.Caption = "Show Ge"
 lbl(42).Caption = "User Name"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "OPR#"
    lbl(9).Caption = "Date"
    lbl(10).Caption = "Branch"
    Frame2.Caption = "Line Data  "
    lbl(11).Caption = "Production line"
    lbl(16).Caption = "Total Salaries"
  lbl(36).Caption = "Total expenses, fuel "
   lbl(17).Caption = "Total expenses, electricity"
  lbl(18).Caption = "Total Expenses"
  Frame1.Caption = "Period"
  lbl(15).Caption = " NO. hours"
  Frame4.Caption = "Data Qantity Produced"
  '  ChkLocked.Caption = "Locked"
  lbl(21).Caption = "Total qantity produced"
  lbl(22).Caption = "According unit value ratios"
  lbl(23).Caption = "Valueof materials"
  lbl(3).Caption = "Remarks"
  lbl(13).Caption = "From"
  lbl(14).Caption = "To"
 Frame3.Caption = "Work Order Data"
 lbl(19).Caption = "Order No"
 lbl(0).Caption = "Customer"
 lbl(12).Caption = "Req QTY"
 lbl(28).Caption = "Product Category"
  lbl(4).Caption = "Remarks"
  lbl(27).Caption = "Remarks"
   lbl(34).Caption = "Order No"
    lbl(24).Caption = "QTY"
    lbl(29).Caption = "Unit"
    lbl(25).Caption = "Sort"
    lbl(35).Caption = "Arm"
    lbl(33).Caption = "Discount%"
      lbl(26).Caption = "Store"
        lbl(32).Caption = "Debit ACC"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Remove"

   ' Option1.Caption = "All Groups"
   ' Option2.Caption = "Select Group"
   ' Cmd(20).Caption = "Add"
   ' Cmd(21).Caption = "Remove"

   ' CmdRemove.Caption = "Remove Line"
 lbl(30).Caption = "Total QTY"
 C1Tab1.Caption = "Production Allocation"
lbl(44).Caption = "This screen you customize the production lines for the job orders and quantities produced"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "No"
        .TextMatrix(0, .ColIndex("NProductionOrderNO")) = "Order No"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit"
        .TextMatrix(0, .ColIndex("ClassName")) = "Sort"
        .TextMatrix(0, .ColIndex("MaterialsValue")) = "materials  "
        .TextMatrix(0, .ColIndex("SalariesValue")) = "Salaries  "
        .TextMatrix(0, .ColIndex("LineExpensesValue")) = "Line Expenses "
        .TextMatrix(0, .ColIndex("Price")) = "Initial cost"
        .TextMatrix(0, .ColIndex("DiscountPercentage")) = "Discount %"
         .TextMatrix(0, .ColIndex("DiscountValue")) = "Discount Value "
        .TextMatrix(0, .ColIndex("Cost")) = "Cost"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store"
        .TextMatrix(0, .ColIndex("Account2name")) = "Debit account"
         .TextMatrix(0, .ColIndex("NoteSerial1")) = "VCHR NO."
        .TextMatrix(0, .ColIndex("NoteSerial")) = "GL No."
        .TextMatrix(0, .ColIndex("gasExpenses")) = "gas Expenses"
        .TextMatrix(0, .ColIndex("ElectricExpenses")) = "Electric Expenses"
        .TextMatrix(0, .ColIndex("OverHead")) = "OverHead"
                .TextMatrix(0, .ColIndex("Alarm")) = "Arm"
                        .TextMatrix(0, .ColIndex("sand")) = "Vouchers"
                        
        
        .TextMatrix(0, .ColIndex("remarks")) = "ÒRemarks"
        .TextMatrix(0, .ColIndex("hours")) = "No Hours"
        .TextMatrix(0, .ColIndex("todate")) = "To Date"
        .TextMatrix(0, .ColIndex("fromdate")) = "From Date"
        
    End With

   ' Me.C1Tab1.TabCaption(1) = "Groups"
   ' Me.C1Tab1.TabCaption(0) = "Items"
End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
     Grid.Clear flexClearScrollable, flexClearEverything
              Grid.Rows = 1
              
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .Rows = .Rows + 1
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "Total"
        End If
        
        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

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

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
 ReLineGrid

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid


   If .Rows > 1 Then
 lbl(31).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Qty"), .Rows - 1, .ColIndex("Qty"))
 TxtTotalProductionQty.Text = lbl(31).Caption
 Else
  lbl(31).Caption = 0
 TxtTotalProductionQty.Text = lbl(31).Caption
 End If
 
 
        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              'ma .TextMatrix(i, .ColIndex("OverHead")) = GetOverHeadForItems(val(.TextMatrix(i, .ColIndex("ItemId"))))
                                .TextMatrix(i, .ColIndex("Price")) = Round(val(.TextMatrix(i, .ColIndex("MaterialsValue"))) + val(.TextMatrix(i, .ColIndex("LineExpensesValue"))) + (val(.TextMatrix(i, .ColIndex("OverHead"))) * val(.TextMatrix(i, .ColIndex("Qty")))), 2)
                                '+  val(.TextMatrix(i, .ColIndex("OverHead")))  * val(.TextMatrix(i, .ColIndex("Qty")))
                            
                          .TextMatrix(i, .ColIndex("DiscountValue")) = Round(val(.TextMatrix(i, .ColIndex("DiscountPercentage"))) * .TextMatrix(i, .ColIndex("Price")) / 100, 2)
                            
                       .TextMatrix(i, .ColIndex("Cost")) = Round(val(.TextMatrix(i, .ColIndex("Price"))) - val(.TextMatrix(i, .ColIndex("DiscountValue"))), 2)


               'X = val(.TextMatrix(i, .ColIndex("gasExpenses"))) + val(.TextMatrix(i, .ColIndex("ElectricExpenses"))) + val(.TextMatrix(i, .ColIndex("SalariesValue")))
                        If val(TxtTotalProductionQty.Text) <> 0 Then
                      '  .TextMatrix(i, .ColIndex("gasExpenses")) = Round(val(TXTTotalUsedPowerPriceH.text) / val(TxttotalOrderQty.text), 2)
                      '  .TextMatrix(i, .ColIndex("ElectricExpenses")) = Round(val(TxtTotalElectricalsaLL.text) / val(TxttotalOrderQty.text), 2)
                      '  .TextMatrix(i, .ColIndex("SalariesValue")) = Round(val(TxtTotalSalariesaLL.text) / val(TxttotalOrderQty.text), 2)
                      '      .TextMatrix(i, .ColIndex("LineExpensesValue")) = val(.TextMatrix(i, .ColIndex("gasExpenses"))) + val(.TextMatrix(i, .ColIndex("ElectricExpenses"))) + val(.TextMatrix(i, .ColIndex("SalariesValue"))) 'Round(val(TxtTotalElectricalsaLL.text) / val(TxttotalOrderQty.text), 2)
                      '   .TextMatrix(i, .ColIndex("Price")) = Round(val(.TextMatrix(i, .ColIndex("MaterialsValue"))) + val(.TextMatrix(i, .ColIndex("LineExpensesValue"))), 2)
                      '   .TextMatrix(i, .ColIndex("DiscountValue")) = Round(val(.TextMatrix(i, .ColIndex("DiscountPercentage"))) * .TextMatrix(i, .ColIndex("Price")) / 100, 2)
                      '  .TextMatrix(i, .ColIndex("Cost")) = Round(val(.TextMatrix(i, .ColIndex("Price"))) - val(.TextMatrix(i, .ColIndex("DiscountValue"))), 2)
                    
                    End If
         
  
            End If

        Next i

 
    End With

 

End Sub
 
 
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

     On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
 

    If rs.RecordCount < 1 Then
        Exit Sub
    End If

        If Lngid <> 0 Then
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    
 
    Me.txtid.Text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
    dbOPrDate.value = IIf(IsNull(rs("OPrDate").value), Date, rs("OPrDate").value)
    ReciveDate.value = IIf(IsNull(rs("ReciveDate").value), dbOPrDate.value, rs("ReciveDate").value)
 'ReciveDate
 DCranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)
 dcLineID.BoundText = IIf(IsNull(rs("LineID").value), "", rs("LineID").value)
 
 TxtNoteID.Text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
 TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
 '****
 TxtNoteSerial1V.Text = IIf(IsNull(rs("NoteSerial1V").value), "", rs("NoteSerial1V").value)
 txtTransaction_ID.Text = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
 DCStores.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
 '***
 DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
 TxtTotalSalaries.Text = IIf(IsNull(rs("TotalSalaries").value), 0, rs("TotalSalaries").value)
 TxtTotalElectricals.Text = IIf(IsNull(rs("TotalElectricals").value), 0, rs("TotalElectricals").value)
 TXTUsedPowerPriceH.Text = IIf(IsNull(rs("UsedPowerPriceH").value), 0, rs("UsedPowerPriceH").value)
 
 TxttotalLineExpenses.Text = IIf(IsNull(rs("totalLineExpenses").value), 0, rs("totalLineExpenses").value)
 DCItemID.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
 DDcunits.BoundText = IIf(IsNull(rs("UNITID").value), "", rs("UNITID").value)
DDcunits1.BoundText = IIf(IsNull(rs("UNITID").value), "", rs("UNITID").value)

 TxtTotalProductionQty.Text = IIf(IsNull(rs("TotalProductionQty").value), 0, rs("TotalProductionQty").value)
 TxttotalMaterialsForItems.Text = IIf(IsNull(rs("totalMaterialsForItems").value), 0, rs("totalMaterialsForItems").value)
 
 TxtTotalMaterials.Text = IIf(IsNull(rs("totalMaterials").value), 0, rs("totalMaterials").value)
 
     Dim fomshift1 As Date
    Dim todate1 As Date

    If Not IsNull(rs("fromTime").value) Then
        fomshift1 = FormatDateTime(rs("fromTime").value, vbShortTime)
        Me.DBfromTime.value = fomshift1
   
    End If

    If Not IsNull(rs("totime").value) Then
        todate1 = FormatDateTime(rs("totime").value, vbShortTime)
        Me.DBtoTime.value = todate1
   
   
    End If
    
 TxtNoOfHours.Text = IIf(IsNull(rs("NoOfHours").value), 0, rs("NoOfHours").value)
 
 TxtWorkOrderNO.Text = IIf(IsNull(rs("WorkOrderNO").value), "", rs("WorkOrderNO").value)
 
   DBCboClientName.BoundText = IIf(IsNull(rs("CustomerId").value), "", rs("CustomerId").value)

TxttotalOrderQty.Text = IIf(IsNull(rs("totalOrderQty").value), 0, rs("totalOrderQty").value)

    TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
 

    StrSQL = " SELECT   dbo.tblProductionAllocDetails.OverHead,   dbo.tblProductionAllocDetails.fromdate, dbo.tblProductionAllocDetails.todate, dbo.tblProductionAllocDetails.fromTime, dbo.tblProductionAllocDetails.toTime, "
   StrSQL = StrSQL & "                    dbo.tblProductionAllocDetails.hours, dbo.tblProductionAllocDetails.gasExpenses, dbo.tblProductionAllocDetails.ElectricExpenses, dbo.tblProductionAlloc.OPrDate,"
  StrSQL = StrSQL & "                     dbo.tblProductionAlloc.BranchID, dbo.tblProductionAlloc.LineID, dbo.tblProductionAlloc.UserID, dbo.tblProductionAlloc.fromTime AS Expr1,"
  StrSQL = StrSQL & "                     dbo.tblProductionAlloc.toTime AS Expr2, dbo.tblProductionAlloc.NoOfHours, dbo.tblProductionAlloc.TotalSalaries, dbo.tblProductionAlloc.TotalElectricals,"
  StrSQL = StrSQL & "                     dbo.tblProductionAlloc.totalLineExpenses, dbo.tblProductionAlloc.WorkOrderNO, dbo.tblProductionAlloc.customerid, dbo.tblProductionAlloc.totalOrderQty,"
  StrSQL = StrSQL & "                     dbo.tblProductionAlloc.TotalProductionQty, dbo.tblProductionAlloc.totalMaterialsForItems, dbo.tblProductionAlloc.totalMaterials, dbo.tblProductionAlloc.REMARKS,"
  StrSQL = StrSQL & "                     dbo.tblProductionAlloc.ItemID, dbo.tblProductionAlloc.UnitID, dbo.tblProductionAllocDetails.ClassId, dbo.tblProductionAllocDetails.Qty,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.NoteSerial, dbo.tblProductionAllocDetails.NoteSerial1, dbo.tblProductionAllocDetails.REMARKS AS RemarksDetails,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.UnitID AS UnitDetails, dbo.tblProductionAllocDetails.Price, dbo.tblProductionAllocDetails.itemid AS itemDetails,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.StoreID, dbo.tblProductionAllocDetails.NProductionOrderNO, dbo.TblItemsclasses.SizeName , dbo.TblItemsclasses.SizeNamee, dbo.TblStore.StoreName,"
  StrSQL = StrSQL & "                     dbo.TblStore.StoreNamee, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.MaterialsValue, dbo.tblProductionAllocDetails.SalariesValue, dbo.tblProductionAllocDetails.LineExpensesValue,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.DiscountPercentage, dbo.tblProductionAllocDetails.Cost, dbo.tblProductionAllocDetails.Account_Code,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.Account_Code1, dbo.tblProductionAllocDetails.DiscountValue, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_NameEng,"
  StrSQL = StrSQL & "                     ACCOUNTS_1.Account_Name AS DiscountAccount_Name, ACCOUNTS_1.Account_NameEng AS DiscountAccount_NameEng, dbo.tblProductionAllocDetails.Alarm,"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails.id2 , dbo.tblProductionAllocDetails.StrSand, dbo.tblProductionAllocDetails.totalss, dbo.tblProductionAllocDetails.StrSelectSands"
  StrSQL = StrSQL & "  FROM         dbo.tblProductionAlloc INNER JOIN"
  StrSQL = StrSQL & "                     dbo.tblProductionAllocDetails ON dbo.tblProductionAlloc.ID = dbo.tblProductionAllocDetails.ID INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblItemsclasses ON dbo.tblProductionAllocDetails.ClassId = dbo.TblItemsclasses.SizeId INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblStore ON dbo.tblProductionAllocDetails.StoreID = dbo.TblStore.StoreID INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblUnites ON dbo.tblProductionAllocDetails.UnitID = dbo.TblUnites.UnitID INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblItems ON dbo.tblProductionAllocDetails.itemid = dbo.TblItems.ItemID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.ACCOUNTS ACCOUNTS_1 ON dbo.tblProductionAllocDetails.Account_Code1 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.ACCOUNTS ACCOUNTS_2 ON dbo.tblProductionAllocDetails.Account_Code = ACCOUNTS_2.Account_Code"
  StrSQL = StrSQL & "   Where (dbo.tblProductionAlloc.id = " & val(Me.txtid.Text) & ")"
   
    
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

   
            For i = .FixedRows To .Rows - 1
            
             .TextMatrix(i, .ColIndex("sand")) = IIf(IsNull(RsDev("StrSand").value), "", RsDev("StrSand").value)
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("id2").value), "", RsDev("id2").value)
            .TextMatrix(i, .ColIndex("sandat")) = IIf(IsNull(RsDev("StrSelectSands").value), "", RsDev("StrSelectSands").value)
  .TextMatrix(i, .ColIndex("totalss")) = IIf(IsNull(RsDev("totalss").value), "", RsDev("totalss").value)
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDev("itemDetails").value), "", RsDev("itemDetails").value)
            
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDev("StoreName").value), "", RsDev("StoreName").value)
                .TextMatrix(i, .ColIndex("itemname")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
          Else
           .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDev("StoreNamee").value), "", RsDev("StoreNamee").value)
          .TextMatrix(i, .ColIndex("itemname")) = IIf(IsNull(RsDev("ItemNamee").value), "", RsDev("ItemNamee").value)
          End If
          
                .TextMatrix(i, .ColIndex("unitid")) = IIf(IsNull(RsDev("UnitDetails").value), "", RsDev("UnitDetails").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            .TextMatrix(i, .ColIndex("Alarm")) = IIf(IsNull(RsDev("Alarm").value), "", RsDev("Alarm").value)
            
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
                .TextMatrix(i, .ColIndex("OverHead")) = IIf(IsNull(RsDev("OverHead").value), 0, (RsDev("OverHead").value))
            
            '
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(RsDev("Qty").value), 0, val(RsDev("Qty").value))
            
                 .TextMatrix(i, .ColIndex("ClassId")) = IIf(IsNull(RsDev("ClassId").value), "", RsDev("ClassId").value)
                 
                 If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("classname")) = IIf(IsNull(RsDev("SizeName").value), "", RsDev("SizeName").value)
            Else
            .TextMatrix(i, .ColIndex("classname")) = IIf(IsNull(RsDev("SizeNamee").value), "", RsDev("SizeNamee").value)
            End If
            
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(RsDev("StoreID").value), "", RsDev("StoreID").value)
               
            .TextMatrix(i, .ColIndex("REMARKS")) = IIf(IsNull(RsDev("REMARKS").value), "", RsDev("REMARKS").value)
            .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsDev("NoteSerial").value), "", RsDev("NoteSerial").value)
            .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), "", RsDev("NoteSerial1").value)
            

          .TextMatrix(i, .ColIndex("Account1")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
        .TextMatrix(i, .ColIndex("Account2")) = IIf(IsNull(RsDev("Account_Code1").value), "", RsDev("Account_Code1").value)
               If SystemOptions.UserInterface = ArabicInterface Then
      .TextMatrix(i, .ColIndex("Account1name")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
        .TextMatrix(i, .ColIndex("Account2name")) = IIf(IsNull(RsDev("DiscountAccount_Name").value), "", RsDev("DiscountAccount_Name").value)
 Else
       .TextMatrix(i, .ColIndex("Account1name")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
        .TextMatrix(i, .ColIndex("Account2name")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)

 
 End If
 



            
                 .TextMatrix(i, .ColIndex("MaterialsValue")) = IIf(IsNull(RsDev("MaterialsValue").value), "", RsDev("MaterialsValue").value)
                     .TextMatrix(i, .ColIndex("SalariesValue")) = IIf(IsNull(RsDev("SalariesValue").value), "", RsDev("SalariesValue").value)
                          .TextMatrix(i, .ColIndex("LineExpensesValue")) = IIf(IsNull(RsDev("LineExpensesValue").value), "", RsDev("LineExpensesValue").value)
                               .TextMatrix(i, .ColIndex("DiscountPercentage")) = IIf(IsNull(RsDev("DiscountPercentage").value), "", RsDev("DiscountPercentage").value)
                                    .TextMatrix(i, .ColIndex("discountvalue")) = IIf(IsNull(RsDev("discountvalue").value), "", RsDev("discountvalue").value)
                                         .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
                                         .TextMatrix(i, .ColIndex("NProductionOrderNO")) = IIf(IsNull(RsDev("NProductionOrderNO").value), "", RsDev("NProductionOrderNO").value)
                                         
  .TextMatrix(i, .ColIndex("fromdate")) = IIf(IsNull(RsDev("fromdate").value), "", RsDev("fromdate").value)
  .TextMatrix(i, .ColIndex("todate")) = IIf(IsNull(RsDev("todate").value), "", RsDev("todate").value)
  .TextMatrix(i, .ColIndex("fromTime")) = IIf(IsNull(RsDev("fromTime").value), "", RsDev("fromTime").value)
  .TextMatrix(i, .ColIndex("toTime")) = IIf(IsNull(RsDev("toTime").value), "", RsDev("toTime").value)
  .TextMatrix(i, .ColIndex("hours")) = IIf(IsNull(RsDev("hours").value), "", RsDev("hours").value)
  
  .TextMatrix(i, .ColIndex("gasExpenses")) = IIf(IsNull(RsDev("gasExpenses").value), "", RsDev("gasExpenses").value)
  .TextMatrix(i, .ColIndex("ElectricExpenses")) = IIf(IsNull(RsDev("ElectricExpenses").value), "", RsDev("ElectricExpenses").value)
  
  
                                        
                 
            RsDev.MoveNext
            Next i
     .AutoSize 0, .Cols - 1, False
        End With

    End If

    RsDev.Close
 
 
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
 

 


Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
  With Me.Grid

        Select Case .ColKey(Col)
    
                 Case "sand"
                  LngRow = Row

 'LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmSandSelected
                FrmSandSelected.show

                    
                End Select
                End With
End Sub

Private Sub Grid_Click()
' With Me.Grid
  
'            If .TextMatrix(i, .ColIndex("MaterialsValue")) <> "" Then
'
'            End If
' End With
            
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid

Select Case .ColKey(Col)
Case "sand"
            .ColComboList(.ColIndex("sand")) = "..."
            End Select
            End With
End Sub

Private Sub lbl_Click(Index As Integer)
Select Case Index

Case 38
FrmItemsDetails2.show

End Select
End Sub

Private Sub ReciveDate_Change()
If Me.TxtModFlg = "E" Then
        If Month(rs("ReciveDate").value) = Month(ReciveDate.value) Then Exit Sub
    End If
If Me.TxtModFlg <> "R" Then
TxtNoteSerial1V.Text = ""
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True
   Ele(1).Enabled = True
        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TxtNoOfHours_Change()
If Me.TxtModFlg <> "R" Then

 add_line (val(Me.dcLineID.BoundText))
 
End If
End Sub

Public Sub RetriveOrder(Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    'On Error GoTo ErrTrap
 

    Dim RsMainData  As New ADODB.Recordset
    Dim StrSQLMain As String
    Dim i As Integer
    Dim LngItemID As Long
    Dim LngItemID2 As Long
    Dim lngShowQty As Double
    Dim currentrow As Integer
    Dim UnitID As Integer
    currentrow = 0
    StrSQLMain = " SELECT    Transactions.CusID,  dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.UnitId"
    StrSQLMain = StrSQLMain & " FROM         dbo.Transactions INNER JOIN"
    StrSQLMain = StrSQLMain & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQLMain = StrSQLMain & "  WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & order_no & "')"
    RsMainData.Open StrSQLMain, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsMainData.RecordCount < 1 Then
 
 
 
            DBCboClientName.BoundText = ""
 
 
 DCItemID.BoundText = ""
 TxttotalOrderQty.Text = ""
 DDcunits.BoundText = ""
  DDcunits1.BoundText = ""
 
        Exit Sub
        
        
        Else
        
           DBCboClientName.BoundText = IIf(IsNull(RsMainData("CusID").value), "", RsMainData("CusID").value)
    LngItemID = IIf(IsNull(RsMainData("Item_ID")), 0, (RsMainData("Item_ID").value))
         lngShowQty = IIf(IsNull(RsMainData("ShowQty")), 0, (RsMainData("ShowQty").value))
 UnitID = IIf(IsNull(RsMainData("unitid")), 0, (RsMainData("unitid").value))
 
 DCItemID.BoundText = LngItemID
 TxttotalOrderQty.Text = lngShowQty
 DDcunits.BoundText = UnitID
 DDcunits1.BoundText = UnitID
    End If

 
     

 
 
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub TxtQty_Change()
  If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
       
        show_parts
    End If
    
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TxtQty.Text, 0)

End Sub

Private Sub TxtStoreID_KeyPress(KeyAscii As Integer)
 Dim StoreID As Integer

    If KeyAscii = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCStores.BoundText = StoreID
    End If
End Sub

Private Sub TxtTotalProductionQty_Change()
  If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
       
        show_parts
    End If
    
End Sub

Private Sub TxtTotalProductionQty_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TxtTotalProductionQty.Text, 0)
End Sub

Public Function show_parts()
 
    Dim RowNum As Integer
    Fg1.Clear flexClearScrollable, flexClearEverything
    Fg1.Rows = 2
          
 

        If DCItemID.BoundText <> "" Then
            If add_part_item(val(DCItemID.BoundText), val(TxtTotalProductionQty.Text)) Then
        
            End If
        End If

 

End Function


Public Function add_part_item(LngItemID As Long, _
                              Optional Qty As Long) As Boolean
    '131315
    Dim StrSQL As String
    Dim RsParts As ADODB.Recordset
    Dim i As Integer
  
    StrSQL = "SELECT  dbo.TblItemsParts.Unitid,  dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.TableID   ,dbo.TblItems.ItemName, dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.ItemID, dbo.TblItems.ItemCode"
    StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN "
    StrSQL = StrSQL + " dbo.TblItemsParts ON dbo.TblItems.ItemID = dbo.TblItemsParts.PartItemID"
    StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & LngItemID
    StrSQL = StrSQL + " Order By TableID"
    Dim item_cost As Long
    Set RsParts = New ADODB.Recordset
    RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsParts.EOF Or RsParts.BOF) Then

        For i = 0 To RsParts.RecordCount - 1
               
            item_cost = ModItemCostPrice.GetCostItemPrice(RsParts("PartItemID").value, 0, , , SystemOptions.SysMainStockCostMethod, , , , , RsParts("Unitid").value)

            If add_item_to_parts_grid(val(RsParts("PartItemID").value), RsParts("ItemCode").value, RsParts("ItemName").value, item_cost, val(RsParts("PartItemQty").value), Qty, val(RsParts("Unitid").value)) = True Then
            End If
                  
            RsParts.MoveNext
        Next i

    End If

End Function


Public Function add_item_to_parts_grid(ItemID As Long, _
                                       itemcode As String, _
                                       itemname As String, _
                                       cost As Long, _
                                       Qty As Long, _
                                       productQty As Long, Optional UnitID As Integer)
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim StrSQL As String
    LngNewRow = ModFgLib.SetFgForNewRow(Fg1, Fg1.ColIndex("Code"))

    StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & "TblUnites.UnitName, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & "TblItemsUnits.UnitID "
    StrSQL = StrSQL + " Where  TblUnites.UnitID=" & val(UnitID)
    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    Dim UnitName As String

    If Not (rs.BOF Or rs.EOF) Then
        UnitID = IIf(IsNull(rs("UnitID").value), 0, rs("UnitID").value)
        UnitName = IIf(IsNull(rs("UnitName").value), 0, rs("UnitName").value)
    End If

    With Me.Fg1
        .TextMatrix(LngNewRow, .ColIndex("id")) = ItemID
        .TextMatrix(LngNewRow, .ColIndex("code")) = itemcode
        .TextMatrix(LngNewRow, .ColIndex("Name")) = itemname
        .TextMatrix(LngNewRow, .ColIndex("count")) = Qty
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = UnitID
        .TextMatrix(LngNewRow, .ColIndex("Unitname")) = UnitName
        .TextMatrix(LngNewRow, .ColIndex("Cost")) = cost
        .TextMatrix(LngNewRow, .ColIndex("Valu")) = cost * Qty
        .TextMatrix(LngNewRow, .ColIndex("TotalQty")) = productQty * Qty
        .TextMatrix(LngNewRow, .ColIndex("Total")) = productQty * cost * Qty
    
        .AutoSize 0, .Cols - 1, False
   
        If .Rows > 1 Then
            Me.TxtTotalMaterials.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))
            Me.TxttotalMaterialsForItems.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
            
      '      If val(TxtTotalProductionQty.text) > 0 Then
      '              TxttotalMaterialsForItems = TxttotalMaterials / TxtTotalProductionQty
      '      Else
      '           TxttotalMaterialsForItems = 0
      '      End If
            
        Else
            Me.TxtTotalMaterials.Text = 0
            TxttotalMaterialsForItems = 0
        End If

    End With

End Function
Private Sub TxtWorkOrderNO_Change()
    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder (Me.TxtWorkOrderNO.Text)
        
    End If
End Sub

Private Sub TxtWorkOrderNO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
         Order_no_search2.show
            Order_no_search2.RetrunType = 7


    End If
    
End Sub

Private Sub TxtWorkOrderNOSub_Change()
    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder (Me.TxtWorkOrderNOSub.Text)
        
    End If
    
End Sub

Private Sub TxtWorkOrderNOSub_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
         Order_no_search2.show
            Order_no_search2.RetrunType = 9


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
