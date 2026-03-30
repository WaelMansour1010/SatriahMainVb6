VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form ProjectsBillAlarmX 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13470
   Icon            =   "ProjectsBillAlarm1X.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   13470
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic frm_alarm 
      Height          =   9684
      Left            =   13680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11544
      _cx             =   20373
      _cy             =   17092
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
      Begin VB.CommandButton btnProductionSave 
         Caption         =   "Save"
         Height          =   492
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2400
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   492
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Frame Frame1 
         Caption         =   "œ·«·«  «·«·Ê«‰"
         Height          =   495
         Left            =   6888
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   9120
         Width           =   3996
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "„”œœ Ã“∆Ì«"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— „”œœ »«·ﬂ«„·"
            Height          =   255
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   11628
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   510
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Text            =   "modflag"
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   450
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   3
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
               TabIndex        =   4
               Top             =   45
               Width           =   855
            End
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
                  Picture         =   "ProjectsBillAlarm1X.frx":058A
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":0924
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":0CBE
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":1058
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":13F2
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":178C
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":1B26
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectsBillAlarm1X.frx":20C0
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
            Height          =   372
            Index           =   2
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   7920
         End
      End
      Begin MSChart20Lib.MSChart chrt_alarm 
         Height          =   4224
         Left            =   -360
         OleObjectBlob   =   "ProjectsBillAlarm1X.frx":245A
         TabIndex        =   8
         Top             =   360
         Width           =   12108
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   4392
         Left            =   0
         TabIndex        =   9
         Top             =   4680
         Width           =   11592
         _cx             =   20447
         _cy             =   7747
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
         FormatString    =   $"ProjectsBillAlarm1X.frx":4912
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
         Height          =   336
         Left            =   240
         TabIndex        =   15
         Top             =   9240
         Width           =   564
         _ExtentX        =   1005
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
         ButtonImage     =   "ProjectsBillAlarm1X.frx":4B4D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin Cfx62ClientServerCtl.Chart Chart1 
         Height          =   972
         Left            =   0
         TabIndex        =   16
         Top             =   744
         Visible         =   0   'False
         Width           =   1020
         _Data_          =   "ProjectsBillAlarm1X.frx":4EE7
      End
      Begin Cfx62ClientServerCtl.Chart Chart2 
         Height          =   1092
         Left            =   1140
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   1020
         _Data_          =   "ProjectsBillAlarm1X.frx":5388
      End
      Begin Cfx62ClientServerCtl.Chart Chart3 
         Height          =   492
         Left            =   2148
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1020
         _Data_          =   "ProjectsBillAlarm1X.frx":5829
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   396
         Left            =   1008
         TabIndex        =   21
         Top             =   9240
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   688
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
         ButtonImage     =   "ProjectsBillAlarm1X.frx":5CCA
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VB.Timer tmrScrolling 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   9960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   9960
   End
   Begin C1SizerLibCtl.C1Elastic frm_Dash 
      Height          =   9792
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   13464
      _cx             =   23760
      _cy             =   17277
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
         Height          =   360
         Left            =   72
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   72
         Width           =   13320
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   372
            Index           =   1
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   0
            Width           =   7920
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   9120
         Left            =   72
         TabIndex        =   25
         Top             =   480
         Width           =   13320
         _cx             =   23495
         _cy             =   16087
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
         Caption         =   "«· Õ·Ì· «·„«·Ï|«·«‰ «Ã|«·ÕÃ“|«·«Ã„«·Ì« |«⁄œ«œ« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   8700
            Left            =   14565
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   45
            Width           =   13230
            _cx             =   23336
            _cy             =   15346
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
            Begin VB.Frame Frame8 
               Height          =   3012
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   2040
               Width           =   3852
               Begin VB.OptionButton chkT 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«Ã„«·Ì« "
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   2040
                  Width           =   1692
               End
               Begin VB.OptionButton chkR 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ÕÃ“"
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   1560
                  Width           =   1692
               End
               Begin VB.OptionButton chkP 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«‰ «Ã"
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   1080
                  Width           =   1692
               End
               Begin VB.OptionButton chkF 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· Õ·Ì· «·„«·Ï"
                  Height          =   252
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   600
                  Width           =   1692
               End
               Begin VB.CommandButton Command8 
                  Caption         =   "Save"
                  Height          =   492
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   2400
                  Width           =   1932
               End
               Begin VB.Label Label32 
                  Alignment       =   2  'Center
                  Caption         =   " »œ√ «·‘«‘… »"
                  Height          =   372
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   240
                  Width           =   3012
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "„‰  «—ÌŒ"
               Height          =   2052
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   3000
               Width           =   4575
               Begin VB.CommandButton Command4 
                  Caption         =   "Save"
                  Height          =   492
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   1320
                  Width           =   1932
               End
               Begin MSComCtl2.DTPicker dtpFromDate 
                  Height          =   312
                  Left            =   960
                  TabIndex        =   88
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   2004
                  _ExtentX        =   3545
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   94109699
                  CurrentDate     =   37140
               End
               Begin MSComCtl2.DTPicker dtpToDate 
                  Height          =   312
                  Left            =   960
                  TabIndex        =   89
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   2004
                  _ExtentX        =   3545
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   94109699
                  CurrentDate     =   37140
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈·Ï  «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   252
                  Left            =   3228
                  TabIndex        =   91
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
                  TabIndex        =   90
                  Top             =   240
                  Width           =   744
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   " €ÌÌ— «⁄œ«œ«  «·„ƒﬁ « "
               Height          =   2052
               Left            =   8628
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   3000
               Width           =   4596
               Begin VB.CommandButton Command3 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   1320
                  Width           =   1092
               End
               Begin VB.CommandButton Command2 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   840
                  Width           =   1092
               End
               Begin VB.CommandButton Command1 
                  Caption         =   " ›⁄Ì·"
                  Height          =   372
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
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
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   94109699
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
                  TabIndex        =   81
                  TabStop         =   0   'False
                  Top             =   840
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   94109699
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
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   1452
                  _ExtentX        =   2566
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "HH:mm:ss"
                  Format          =   94109699
                  UpDown          =   -1  'True
                  CurrentDate     =   37140.875
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ‰ﬁ· »Ì‰ «·«”ÿ—"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   1320
                  Width           =   1092
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ‰ﬁ· »Ì‰ «·‘«‘« "
                  Height          =   372
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ÕœÌÀ «·ﬂ·"
                  Height          =   372
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   360
                  Width           =   852
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "«· Õ·Ì· «·„«·Ï"
               Height          =   2880
               Left            =   8640
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   0
               Width           =   4596
               Begin VB.ComboBox cbSection1 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":6064
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":6074
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   600
                  Width           =   2052
               End
               Begin VB.ComboBox cbSection2 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":60A4
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":60B4
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   1080
                  Width           =   2052
               End
               Begin VB.ComboBox cbSection3 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":60E4
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":60F4
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1560
                  Width           =   2052
               End
               Begin VB.ComboBox cbSection4 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":6124
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":6134
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   2040
                  Width           =   2052
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ Ê«Õœ"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   600
                  Width           =   852
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «À‰Ì‰"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1080
                  Width           =   852
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·À«·À"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1560
                  Width           =   852
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·—«»⁄"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   2040
                  Width           =   852
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "«·«‰ «Ã"
               Height          =   2880
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   0
               Width           =   4575
               Begin VB.ComboBox cbPSection1 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":6164
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":6174
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   600
                  Width           =   2052
               End
               Begin VB.ComboBox cbPSection4 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":61A3
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":61B3
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   2040
                  Width           =   2052
               End
               Begin VB.ComboBox cbPSection3 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":61E2
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":61F2
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   1560
                  Width           =   2052
               End
               Begin VB.ComboBox cbPSection2 
                  Height          =   288
                  ItemData        =   "ProjectsBillAlarm1X.frx":6221
                  Left            =   960
                  List            =   "ProjectsBillAlarm1X.frx":6231
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   1080
                  Width           =   2052
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·—«»⁄"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   2040
                  Width           =   852
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «·À«·À"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   1560
                  Width           =   852
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ «À‰Ì‰"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   1080
                  Width           =   852
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„—»⁄ Ê«Õœ"
                  Height          =   372
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   600
                  Width           =   852
               End
            End
            Begin VB.Frame Frame7 
               Height          =   2052
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   0
               Width           =   3852
               Begin VB.OptionButton opt_TabChangeNot 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œ„  ›⁄Ì· «· ‰ﬁ·"
                  Height          =   312
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   840
                  Width           =   1572
               End
               Begin VB.OptionButton opt_TabChange 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ›⁄Ì· «· ‰ﬁ·"
                  Height          =   312
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   480
                  Width           =   2052
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Save"
                  Height          =   492
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   1320
                  Width           =   1572
               End
               Begin VB.CheckBox chkInvisble 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Œ›«¡ «·—”„ «·»Ì«‰Ï"
                  Height          =   492
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   3012
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   8700
            Left            =   14265
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   45
            Width           =   13230
            _cx             =   23336
            _cy             =   15346
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
               Height          =   2640
               Left            =   4440
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   2625
               Width           =   4110
               _cx             =   7250
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
               Caption         =   "«·„»Ì⁄« "
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
                  Height          =   288
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   1200
                  Width           =   2412
               End
               Begin VB.TextBox txtSalestotal 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   720
                  Width           =   2412
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„ Ê”ÿ «·”⁄—"
                  Height          =   252
                  Left            =   2412
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   1200
                  Width           =   1452
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ﬁÌ„… «·„»Ì⁄« "
                  Height          =   252
                  Left            =   2412
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   720
                  Width           =   1452
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   2640
               Left            =   8595
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   2745
               Width           =   4440
               _cx             =   7832
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
               Begin VSFlex8UCtl.VSFlexGrid fg_Charge_Totals 
                  Height          =   2124
                  Left            =   0
                  TabIndex        =   117
                  Top             =   480
                  Width           =   4452
                  _cx             =   7853
                  _cy             =   3746
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":6260
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
                  Caption         =   "«·‘Õ‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   372
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   120
                  Width           =   1932
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   2520
               Left            =   8595
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   240
               Width           =   4440
               _cx             =   7832
               _cy             =   4445
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
               Begin VB.TextBox TxttotalRecivedShippedQty 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   2040
                  Width           =   2412
               End
               Begin VB.TextBox txtProduct_Total 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   240
                  Width           =   2412
               End
               Begin VB.TextBox txtYes 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   720
                  Width           =   2412
               End
               Begin VB.TextBox txtReserve 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   1200
                  Width           =   2412
               End
               Begin VB.TextBox txtCharge 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   1680
                  Width           =   2412
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„Ê—œ «·ÌÊ„-;ﬂ„Ì… „” ·„Â"
                  Height          =   255
                  Left            =   2535
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   2040
                  Width           =   1800
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Caption         =   "√Ã„«·Ï √„ «— «·‘Â—"
                  Height          =   255
                  Left            =   3015
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«‰ «Ã «„”"
                  Height          =   255
                  Left            =   3015
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   720
                  Width           =   1200
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬂ„Ì… «·„ÕÃÊ“…"
                  Height          =   255
                  Left            =   3015
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   1200
                  Width           =   1200
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„Ê—œ «·ÌÊ„-;ﬂ„Ì… „—”·Â"
                  Height          =   255
                  Left            =   2535
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   1680
                  Width           =   1800
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   2280
               Left            =   4440
               TabIndex        =   96
               TabStop         =   0   'False
               Top             =   240
               Width           =   4110
               _cx             =   7250
               _cy             =   4022
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
               Begin VB.CommandButton Command7 
                  Caption         =   " ›«’Ì· «·«Ì—«œ« "
                  Height          =   372
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   1200
                  Width           =   1212
               End
               Begin VB.CommandButton Command6 
                  Caption         =   " ›«’Ì· «·„’—Ê›« "
                  Height          =   372
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   1680
                  Width           =   1212
               End
               Begin VB.TextBox txtBank 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   240
                  Width           =   2652
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   720
                  Width           =   2652
               End
               Begin VB.TextBox txtRevenue 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1332
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1200
                  Width           =   1440
               End
               Begin VB.TextBox txtExpenses 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   1332
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   1680
                  Width           =   1440
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—’Ìœ «·»‰Êﬂ"
                  Height          =   252
                  Left            =   2112
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1752
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—’Ìœ «·Œ“‰…"
                  Height          =   252
                  Left            =   2112
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   720
                  Width           =   1752
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ì—«œ«  «·ÌÊ„"
                  Height          =   252
                  Left            =   2112
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   1200
                  Width           =   1752
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„’—Ê›« "
                  Height          =   252
                  Left            =   2112
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1680
                  Width           =   1752
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   5025
               Left            =   0
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   240
               Width           =   4320
               _cx             =   7620
               _cy             =   8864
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
                  Height          =   288
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   120
                  Width           =   2412
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_MaterialTotal 
                  Height          =   4356
                  Left            =   0
                  TabIndex        =   95
                  Top             =   600
                  Width           =   4392
                  _cx             =   7747
                  _cy             =   7683
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":6304
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
                  Height          =   252
                  Left            =   2772
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   120
                  Width           =   1440
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   8700
            Left            =   13965
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   45
            Width           =   13230
            _cx             =   23336
            _cy             =   15346
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
               Height          =   4080
               Left            =   0
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   0
               Width           =   13155
               _cx             =   23204
               _cy             =   7197
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
                  Height          =   4032
                  Left            =   -300
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":6373
                  TabIndex        =   29
                  Top             =   228
                  Width           =   5664
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Reserve 
                  Height          =   3636
                  Left            =   5364
                  TabIndex        =   30
                  Top             =   504
                  Width           =   7824
                  _cx             =   13801
                  _cy             =   6413
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":882B
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
               Height          =   4080
               Left            =   0
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   4290
               Width           =   13155
               _cx             =   23204
               _cy             =   7197
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
                  Height          =   4044
                  Left            =   -180
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":8A1B
                  TabIndex        =   32
                  Top             =   228
                  Width           =   5424
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Material 
                  Height          =   3636
                  Left            =   5424
                  TabIndex        =   33
                  Top             =   396
                  Width           =   7764
                  _cx             =   13695
                  _cy             =   6413
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":AED3
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
            Height          =   8700
            Left            =   45
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   45
            Width           =   13230
            _cx             =   23336
            _cy             =   15346
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
            Begin VB.Timer tmrLoadingAll 
               Left            =   7320
               Top             =   8520
            End
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
               Height          =   4080
               Left            =   120
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   4290
               Width           =   12975
               _cx             =   22886
               _cy             =   7197
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
                  Height          =   4044
                  Left            =   -360
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":AF42
                  TabIndex        =   36
                  Top             =   228
                  Width           =   6384
               End
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   3636
                  Left            =   6144
                  TabIndex        =   37
                  Top             =   360
                  Width           =   6924
                  _cx             =   12213
                  _cy             =   6413
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":D3FA
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
               Height          =   4080
               Left            =   120
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   120
               Width           =   12975
               _cx             =   22886
               _cy             =   7197
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
                  Height          =   4032
                  Left            =   -120
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":D64B
                  TabIndex        =   39
                  Top             =   228
                  Width           =   6144
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Product 
                  Height          =   3636
                  Left            =   6024
                  TabIndex        =   40
                  Top             =   504
                  Width           =   6984
                  _cx             =   12319
                  _cy             =   6413
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":FB03
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
            Height          =   8700
            Left            =   -13875
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   45
            Width           =   13230
            _cx             =   23336
            _cy             =   15346
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
               Height          =   4080
               Left            =   120
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   120
               Width           =   6495
               _cx             =   11456
               _cy             =   7197
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
                  Height          =   3636
                  Left            =   -60
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":FC1A
                  TabIndex        =   43
                  Top             =   456
                  Width           =   2352
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Bank 
                  Height          =   3528
                  Left            =   2412
                  TabIndex        =   44
                  Top             =   408
                  Width           =   4032
                  _cx             =   7112
                  _cy             =   6223
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":120D2
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
               Height          =   4080
               Left            =   6675
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   120
               Width           =   6495
               _cx             =   11456
               _cy             =   7197
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
                  Height          =   3516
                  Left            =   2352
                  TabIndex        =   46
                  Top             =   420
                  Width           =   4092
                  _cx             =   7218
                  _cy             =   6202
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":12237
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
                  Height          =   3696
                  Left            =   -180
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":12365
                  TabIndex        =   47
                  Top             =   456
                  Width           =   2832
               End
            End
            Begin C1SizerLibCtl.C1Elastic frm_Receipt 
               Height          =   4080
               Left            =   6675
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   4260
               Width           =   6495
               _cx             =   11456
               _cy             =   7197
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
                  Height          =   3816
                  Left            =   -60
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":1481D
                  TabIndex        =   49
                  Top             =   360
                  Width           =   2412
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Receipts 
                  Height          =   3636
                  Left            =   2352
                  TabIndex        =   50
                  Top             =   408
                  Width           =   4212
                  _cx             =   7429
                  _cy             =   6413
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":16CD5
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
               Height          =   4080
               Left            =   120
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   4260
               Width           =   6495
               _cx             =   11456
               _cy             =   7197
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
                  Height          =   3576
                  Left            =   -300
                  OleObjectBlob   =   "ProjectsBillAlarm1X.frx":16D6E
                  TabIndex        =   52
                  Top             =   540
                  Width           =   2712
               End
               Begin VSFlex8UCtl.VSFlexGrid fg_Expenses 
                  Height          =   3516
                  Left            =   2412
                  TabIndex        =   53
                  Top             =   408
                  Width           =   3912
                  _cx             =   6900
                  _cy             =   6202
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
                  FormatString    =   $"ProjectsBillAlarm1X.frx":19226
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
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Height          =   2676
      Left            =   0
      TabIndex        =   99
      Top             =   480
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
      FormatString    =   $"ProjectsBillAlarm1X.frx":192C0
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
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ï ﬁÌ„… «·„Œ“Ê‰"
      Height          =   252
      Left            =   2412
      RightToLeft     =   -1  'True
      TabIndex        =   98
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
Attribute VB_Name = "ProjectsBillAlarmX"
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


Dim Move_Row_Fixed As Long
Dim Move_Row As Long
Dim Move_Tab_Fixed As Long
Dim Move_Tab As Long
Dim LoadingAll As Long
Dim Move_y As Integer, Move_N As Integer
Public Sub FillGridWithData()

    'On Error GoTo ErrTrap
 
    Dim i As Integer
    Dim x As Integer
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
        
        For x = 1 To rs.RecordCount
       
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

Private Sub BtnCancel_Click()
    Me.Hide
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




Private Sub C1Tab1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

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

Print_Expenses

End Sub

Private Sub Command7_Click()
Print_Revenue
End Sub

Private Sub Command8_Click()
        SaveSetting "Win_Sys_EX_B", "Setting", "chkF", CInt(chkF.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "chkP", CInt(chkP.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "chkR", CInt(chkR.value)
        SaveSetting "Win_Sys_EX_B", "Setting", "chkT", CInt(chkT.value)
End Sub

Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FillGridWithData
    
    Chart1.Gallery = Gallery_Pie
    Chart2.Gallery = Gallery_Curve

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



IntervalData.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Move_Row_Time") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "Move_Row_Time"))
IntervalTab.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "Move_Tab_Time") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "Move_Tab_Time"))
IntervalAll.value = IIf(GetSetting("Win_Sys_EX_B", "Setting", "All_Time") = "", Date, GetSetting("Win_Sys_EX_B", "Setting", "All_Time"))

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
Dim Fromdate, todate
Fromdate = GetSetting("Win_Sys_EX_B", "Setting", "FromDate")
todate = GetSetting("Win_Sys_EX_B", "Setting", "ToDate")

dtpFromDate.value = IIf(Fromdate = "", Now, Fromdate)
dtpToDate.value = IIf(todate = "", Now, todate)


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
    Label1(2).Caption = "Dash Board"
    Me.Caption = Label1(2).Caption
    Frame1.Caption = "Color Map"
    Label3.Caption = "Fully"
    Label5.Caption = "Partial"

    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"

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
    .TextMatrix(0, .ColIndex("CusNamee")) = " Customer Name "
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

Label1(1).Caption = " Dash Board "
C1Tab1.Caption = " Financial|Production|Reservation|Totals|Settings "


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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Private Sub LoadData()

Dim chkF As Integer, chkP As Integer, chkR As Integer, chkT As Integer



If Move_N <> 0 Then

chkF = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkF") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkF"))
chkP = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkP") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkP"))
chkR = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkR") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkR"))
chkT = IIf(GetSetting("Win_Sys_EX_B", "Setting", "chkT") = "", 0, GetSetting("Win_Sys_EX_B", "Setting", "chkT"))

If chkF <> 0 Then ' Õ·Ì· „«·Ì
GetExpensesData
GetReceiptsData
GetBoxesData
GetBanksData
End If

If chkP <> 0 Then '«‰ «Ã
GetCharge
GetProduct2
End If

If chkR <> 0 Then 'ÕÃ“
GetReserve2
GetMaterial
End If

If chkT <> 0 Then '«Ã„«·Ì« 
All_Total
End If

Else

GetExpensesData
GetReceiptsData
GetBoxesData
GetBanksData
GetCharge
GetProduct2
GetReserve2
GetMaterial
All_Total
End If




End Sub


Private Sub F()
 

LoadingAll = LoadingAll - 1

If Move_Row <= 0 Then
         LoadData
         LoadingAll = 60 * 1000 * 5
End If

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
        
        Account_Code_dynamic = get_account_code_branch(34, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtRevenue.Text = balanceString
        
        Account_Code_dynamic = get_account_code_branch(20, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtBank.Text = balanceString
        
          Account_Code_dynamic = get_account_code_branch(0, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtInv.Text = balanceString
        
        Account_Code_dynamic = get_account_code_branch(33, my_branch)
        WriteCustomerBalPublic Account_Code_dynamic, Balance, balanceString
        txtExpenses.Text = balanceString
        
        
        
        Get_Product_Total
        Get_Reserved_Total
       Get_Charge_Total
       Get_Product_Yester_Total
       Get_Sales_Total
       Get_Charge_grid
        
End Sub

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
 
 
 
My_SQL = My_SQL & "  select qwq.PTID , qwq.cusid  , qwq.CusName , qwq.CusNamee, qwq.productiontypename , qwq.productiontypenamee , sum(totalShippedQty) total"
My_SQL = My_SQL & "  from ("
My_SQL = My_SQL & "  SELECT TblProductionType.id PTID , TblCustemers.CusID , dbo.GetNoOfShipments(dbo.Transactions.Transaction_ID) AS noofShipments, dbo.Transactions.Transaction_ID,"
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
My_SQL = My_SQL & "  Where (dbo.Transactions.Transaction_Type = 61 "
My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date >= '" & SQLDate(DateTime.Now) & "'"
My_SQL = My_SQL & "  and dbo.Transactions.Transaction_Date <= '" & SQLDate(DateTime.Now) & "'"

My_SQL = My_SQL & " )) qwq  "
My_SQL = My_SQL & "  group by PTID , cusid  , CusNamee , CusName,productiontypename , productiontypenamee"



    
 

   
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
  
              .TextMatrix(i, .ColIndex("total")) = (IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value))
              If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              .TextMatrix(i, .ColIndex("productiontypename")) = (IIf(IsNull(rs.Fields("productiontypename").value), "", rs.Fields("productiontypename").value))
              Else
                 .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
                 .TextMatrix(i, .ColIndex("productiontypenamee")) = (IIf(IsNull(rs.Fields("productiontypenamee").value), "", rs.Fields("productiontypenamee").value))
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
    
       

My_SQL = My_SQL & "  SELECT  sum ( dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice ) Tot ,( sum (showprice) / count(*) ) Avg_Tot"

My_SQL = My_SQL & "  FROM     dbo.Transaction_Details INNER JOIN"
My_SQL = My_SQL & "                    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"


        
   My_SQL = My_SQL & " AND  (dbo.Transactions.Transaction_Date >= (SELECT datesatrt FROM  dbo.TblyearsData WHERE  CurrentYear = 1))    AND  (dbo.Transactions.Transaction_Date <= ((SELECT DateEnd FROM  dbo.TblyearsData WHERE  CurrentYear = 1)))"
   ' My_SQL = My_SQL & "   order by dbo.Transactions.Transaction_Date  "

    Dim ActualTotal As Double
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
    rs.MoveFirst
              txtAvg.Text = Math.Round(IIf(IsNull(rs("Avg_Tot").value), 0, rs("Avg_Tot").value), 2)
              txtSalestotal.Text = Math.Round(IIf(IsNull(rs("Tot").value), 0, rs("Tot").value), 2)
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
MySQL = MySQL & "      From dbo.ACCOUNTS"
MySQL = MySQL & "      WHERE     (AccountTab = 3) AND (AccountTypes = 2)))"

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
 


Private Function Print_Revenue()
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""
    
MySQL = MySQL & " SELECT    dbo.Notes.NoteSerial1, dbo.TblNotesTypes.NotesTypeName, dbo.Notes.Remark, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
MySQL = MySQL & " dbo.Notes.NoteDate"
MySQL = MySQL & " FROM         dbo.Notes INNER JOIN"
MySQL = MySQL & " dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType INNER JOIN"
MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
MySQL = MySQL & " WHERE     (dbo.Notes.NoteType = 4) OR"
MySQL = MySQL & " (dbo.Notes.NoteType = 170) AND (dbo.Notes.NoteID IN"
MySQL = MySQL & " (SELECT     NoteId"
MySQL = MySQL & " From dbo.Transactions"
MySQL = MySQL & " WHERE     (Transaction_Type = 21) AND (Transaction_Date = " & SQLDate(Date, True) & ") AND (PaymentType = 0))) AND"
MySQL = MySQL & " (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)"
MySQL = MySQL & " AND (dbo.Notes.NoteDate = " & SQLDate(Date, True) & " )"
MySQL = MySQL & " ORDER BY dbo.Notes.NoteType DESC"

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

