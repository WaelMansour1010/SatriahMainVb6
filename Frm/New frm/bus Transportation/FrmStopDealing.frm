VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStopDealing 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " «Ìﬁ«› «· ⁄«„·"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "FrmStopDealing.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   7905
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9450
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7905
      _cx             =   13944
      _cy             =   16669
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
      Begin C1SizerLibCtl.C1Elastic cCancel 
         Height          =   5775
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   7710
         _cx             =   13600
         _cy             =   10186
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
         Caption         =   "≈·€«¡ «·«Ìﬁ«›"
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
         Begin VB.TextBox txtRecordNo4 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   885
            Width           =   2055
         End
         Begin VB.TextBox txtFullCode4 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   885
            Width           =   1455
         End
         Begin VB.CheckBox chk4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ — «·ﬂ·"
            Height          =   195
            Left            =   4335
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   2640
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcCustomer4 
            Height          =   315
            Left            =   1080
            TabIndex        =   33
            Top             =   1470
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMinistry4 
            Height          =   315
            Left            =   1080
            TabIndex        =   34
            Top             =   360
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker StopDate4 
            Height          =   390
            Left            =   3660
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1995
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal StopDateH4 
            Height          =   390
            Left            =   1110
            TabIndex        =   36
            Top             =   1995
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   688
         End
         Begin VSFlex8Ctl.VSFlexGrid grid4 
            Height          =   2595
            Left            =   240
            TabIndex        =   37
            Top             =   3000
            Width           =   6120
            _cx             =   10795
            _cy             =   4577
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmStopDealing.frx":038A
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄ﬁœ «·«”‰«œ"
            Height          =   360
            Index           =   24
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   360
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   390
            Index           =   23
            Left            =   3090
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   885
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ ⁄Âœ"
            Height          =   390
            Index           =   22
            Left            =   6135
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   885
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   360
            Index           =   21
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1470
            Width           =   705
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   6255
            TabIndex        =   38
            Top             =   1995
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Elastic cStop 
         Height          =   5775
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1920
         Width           =   7710
         _cx             =   13600
         _cy             =   10186
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
         Begin VB.OptionButton opt_canceled 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·€«¡ „‰ «·⁄ﬁœ"
            Height          =   255
            Left            =   3735
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton opt_temp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ìﬁ«› „ƒﬁ "
            Height          =   255
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton opt_final 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ìﬁ«› ‰Â«∆Ï"
            Height          =   195
            Left            =   6375
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   120
            Width           =   1215
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5160
            Index           =   2
            Left            =   120
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   480
            Width           =   7485
            _cx             =   13203
            _cy             =   9102
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
            Caption         =   "«Ìﬁ«› ‰Â«∆Ï"
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
            Begin VB.ComboBox cbStop1 
               Height          =   315
               ItemData        =   "FrmStopDealing.frx":0438
               Left            =   1680
               List            =   "FrmStopDealing.frx":0442
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   390
               Width           =   4785
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   3540
               Left            =   75
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   960
               Width           =   7320
               _cx             =   12912
               _cy             =   6244
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
               Begin VB.TextBox txtFullCode1 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   4065
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   1110
                  Width           =   1455
               End
               Begin VB.TextBox txtRecordNo1 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   690
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1110
                  Width           =   2070
               End
               Begin MSDataListLib.DataCombo dcCustomer1 
                  Height          =   315
                  Left            =   690
                  TabIndex        =   48
                  Top             =   1680
                  Width           =   4815
                  _ExtentX        =   8493
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcMinistry 
                  Height          =   315
                  Left            =   690
                  TabIndex        =   49
                  Top             =   570
                  Width           =   4815
                  _ExtentX        =   8493
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker StopDate1 
                  Height          =   390
                  Left            =   3285
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   2205
                  Width           =   2250
                  _ExtentX        =   3969
                  _ExtentY        =   688
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   98762755
                  CurrentDate     =   37140
               End
               Begin Dynamic_Byte.NourHijriCal StopDateH1 
                  Height          =   390
                  Left            =   720
                  TabIndex        =   51
                  Top             =   2205
                  Width           =   2445
                  _ExtentX        =   4313
                  _ExtentY        =   688
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   390
                  Left            =   5895
                  TabIndex        =   56
                  Top             =   2205
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ ⁄Âœ"
                  Height          =   360
                  Index           =   9
                  Left            =   6255
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   1680
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·„ ⁄Âœ"
                  Height          =   390
                  Index           =   10
                  Left            =   5775
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   1110
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ã·"
                  Height          =   390
                  Index           =   11
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   1110
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄ﬁœ «·«”‰«œ"
                  Height          =   360
                  Index           =   12
                  Left            =   6255
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   570
                  Width           =   705
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3540
               Left            =   75
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   960
               Width           =   7320
               _cx             =   12912
               _cy             =   6244
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
               Begin VB.TextBox txtRecordNo 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   360
                  Width           =   2052
               End
               Begin VB.TextBox txtFullCode 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   4164
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   360
                  Width           =   1452
               End
               Begin VB.ComboBox cbStopType 
                  Height          =   315
                  ItemData        =   "FrmStopDealing.frx":0452
                  Left            =   720
                  List            =   "FrmStopDealing.frx":045C
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   1680
                  Width           =   4896
               End
               Begin MSDataListLib.DataCombo dcar 
                  Height          =   288
                  Left            =   720
                  TabIndex        =   61
                  Top             =   1200
                  Width           =   4896
                  _ExtentX        =   8625
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   336
                  Left            =   3396
                  TabIndex        =   62
                  TabStop         =   0   'False
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   2220
                  _ExtentX        =   3916
                  _ExtentY        =   582
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   98762755
                  CurrentDate     =   37140
               End
               Begin Dynamic_Byte.NourHijriCal ToDateH 
                  Height          =   336
                  Left            =   1080
                  TabIndex        =   63
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   2436
                  _ExtentX        =   4286
                  _ExtentY        =   582
               End
               Begin MSComCtl2.DTPicker FromDate 
                  Height          =   336
                  Left            =   3396
                  TabIndex        =   64
                  TabStop         =   0   'False
                  Top             =   2160
                  Width           =   2220
                  _ExtentX        =   3916
                  _ExtentY        =   582
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   98762755
                  CurrentDate     =   37140
               End
               Begin Dynamic_Byte.NourHijriCal FromDateH 
                  Height          =   336
                  Left            =   720
                  TabIndex        =   65
                  Top             =   2160
                  Width           =   2436
                  _ExtentX        =   4286
                  _ExtentY        =   582
               End
               Begin MSDataListLib.DataCombo dcCustomer 
                  Height          =   288
                  Left            =   720
                  TabIndex        =   66
                  Top             =   720
                  Width           =   4896
                  _ExtentX        =   8625
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ã·"
                  Height          =   336
                  Index           =   5
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   360
                  Width           =   936
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·„ ⁄Âœ"
                  Height          =   336
                  Index           =   1
                  Left            =   5736
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   360
                  Width           =   1176
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ ⁄Âœ"
                  Height          =   312
                  Index           =   6
                  Left            =   6204
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   720
                  Width           =   708
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï  «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   336
                  Index           =   0
                  Left            =   5856
                  TabIndex        =   70
                  Top             =   2640
                  Visible         =   0   'False
                  Width           =   1056
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "  «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   336
                  Left            =   5856
                  TabIndex        =   69
                  Top             =   2160
                  Width           =   1056
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·«Ìﬁ«›"
                  Height          =   348
                  Index           =   7
                  Left            =   5784
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1656
                  Width           =   1128
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„⁄œÂ/«·”Ì«—…"
                  Height          =   312
                  Index           =   3
                  Left            =   5784
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   1188
                  Width           =   1128
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·«Ìﬁ«›"
               Height          =   285
               Index           =   8
               Left            =   6390
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   390
               Width           =   885
            End
         End
         Begin C1SizerLibCtl.C1Elastic Frame1 
            Height          =   5160
            Left            =   120
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   480
            Width           =   7485
            _cx             =   13203
            _cy             =   9102
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
            Caption         =   "«Ìﬁ«› „ƒﬁ "
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
            Begin VB.CheckBox chk2 
               Alignment       =   1  'Right Justify
               Caption         =   "«Œ — «·ﬂ·"
               Height          =   195
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   2580
               Width           =   1590
            End
            Begin VB.TextBox txtFullCode2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   330
               Left            =   4425
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   870
               Width           =   1455
            End
            Begin VB.TextBox txtRecordNo2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   330
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   870
               Width           =   2055
            End
            Begin MSDataListLib.DataCombo dcCustomer2 
               Height          =   315
               Left            =   1080
               TabIndex        =   79
               Top             =   1425
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcMinistry2 
               Height          =   315
               Left            =   1080
               TabIndex        =   80
               Top             =   345
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker StopDate2 
               Height          =   390
               Left            =   3660
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   1950
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   688
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   98762755
               CurrentDate     =   37140
            End
            Begin Dynamic_Byte.NourHijriCal StopDateH2 
               Height          =   390
               Left            =   1110
               TabIndex        =   82
               Top             =   1950
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   688
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid2 
               Height          =   2070
               Left            =   600
               TabIndex        =   83
               Top             =   2925
               Width           =   5520
               _cx             =   9737
               _cy             =   3651
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmStopDealing.frx":046C
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
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ"
               ForeColor       =   &H00000000&
               Height          =   390
               Left            =   6255
               TabIndex        =   88
               Top             =   1950
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ ⁄Âœ"
               Height          =   360
               Index           =   16
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   1425
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„ ⁄Âœ"
               Height          =   390
               Index           =   15
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   870
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·”Ã·"
               Height          =   390
               Index           =   14
               Left            =   3090
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   870
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄ﬁœ «·«”‰«œ"
               Height          =   360
               Index           =   13
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   345
               Width           =   705
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   5160
            Left            =   120
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   480
            Width           =   7485
            _cx             =   13203
            _cy             =   9102
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
            Caption         =   "≈·€«¡ „‰ «·⁄ﬁœ "
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
            Begin VB.CheckBox chk3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — «·ﬂ·"
               Height          =   195
               Left            =   4215
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   2580
               Width           =   1575
            End
            Begin VB.TextBox txtFullCode3 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   330
               Left            =   3570
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   870
               Width           =   1470
            End
            Begin VB.TextBox txtRecordNo3 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   330
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   870
               Width           =   2055
            End
            Begin MSDataListLib.DataCombo dcCustomer3 
               Height          =   315
               Left            =   240
               TabIndex        =   93
               Top             =   1425
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcMinistry3 
               Height          =   315
               Left            =   240
               TabIndex        =   94
               Top             =   345
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker StopDate3 
               Height          =   390
               Left            =   2820
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   1950
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   688
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   98762755
               CurrentDate     =   37140
            End
            Begin Dynamic_Byte.NourHijriCal StopDateH3 
               Height          =   390
               Left            =   270
               TabIndex        =   96
               Top             =   1950
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   688
            End
            Begin VSFlex8Ctl.VSFlexGrid grid3 
               Height          =   2295
               Left            =   240
               TabIndex        =   97
               Top             =   2820
               Width           =   5520
               _cx             =   9737
               _cy             =   4048
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmStopDealing.frx":051A
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
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ"
               ForeColor       =   &H00000000&
               Height          =   390
               Left            =   5415
               TabIndex        =   102
               Top             =   1950
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ ⁄Âœ"
               Height          =   360
               Index           =   20
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1425
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„ ⁄Âœ"
               Height          =   390
               Index           =   19
               Left            =   5295
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   870
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·”Ã·"
               Height          =   390
               Index           =   18
               Left            =   2250
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   870
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄ﬁœ «·«”‰«œ"
               Height          =   360
               Index           =   17
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   345
               Width           =   705
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   630
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   7755
         _cx             =   13679
         _cy             =   1111
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "   «Ìﬁ«› «· ⁄«„·  "
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
            TabIndex        =   2
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   3
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
            ButtonImage     =   "FrmStopDealing.frx":05C8
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
            TabIndex        =   4
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
            ButtonImage     =   "FrmStopDealing.frx":0962
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
            ButtonImage     =   "FrmStopDealing.frx":0CFC
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
            ButtonImage     =   "FrmStopDealing.frx":1096
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   1065
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   7620
         _cx             =   13441
         _cy             =   1879
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
         Begin VB.OptionButton opt_Stop 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ìﬁ«›"
            Height          =   432
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   852
         End
         Begin VB.OptionButton opt_Cancel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·€«¡ «·«Ìﬁ«›"
            Height          =   372
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1092
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   384
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   72
            Width           =   1824
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
            Height          =   372
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   600
            Width           =   852
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ"
            Height          =   264
            Index           =   0
            Left            =   5496
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   120
            Width           =   816
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   765
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   7815
         Width           =   7695
         _cx             =   13573
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   345
            Index           =   4
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   345
            Index           =   2
            Left            =   4770
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   345
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   810
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   345
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   690
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   8655
         Width           =   7815
         _cx             =   13785
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
            Height          =   450
            Index           =   0
            Left            =   6915
            TabIndex        =   16
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":1430
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
            Height          =   450
            Index           =   1
            Left            =   5985
            TabIndex        =   17
            Top             =   120
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":7C92
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
            Height          =   450
            Index           =   2
            Left            =   5160
            TabIndex        =   18
            Top             =   120
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":E4F4
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
            Height          =   450
            Index           =   3
            Left            =   4365
            TabIndex        =   19
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":14D56
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
            Height          =   450
            Index           =   4
            Left            =   3420
            TabIndex        =   20
            Top             =   120
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":1B5B8
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
            Height          =   450
            Index           =   6
            Left            =   990
            TabIndex        =   21
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":21E1A
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
            Height          =   450
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":4BA3C
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
            Height          =   450
            Index           =   7
            Left            =   2670
            TabIndex        =   23
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":5229E
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
            Height          =   450
            Index           =   9
            Left            =   1785
            TabIndex        =   24
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStopDealing.frx":58B00
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
   End
End
Attribute VB_Name = "FrmStopDealing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp1 As ADODB.Recordset
Dim Rs_Temp2 As ADODB.Recordset
Dim Rs_Temp3 As ADODB.Recordset
Dim TTP As clstooltip

Private Sub cbDiscType_Change()



End Sub

Private Sub cbStop1_Change()

If cbStop1.ListIndex = 0 Then
        C1Elastic2.Visible = True
        C1Elastic5.Visible = False
        
ElseIf cbStop1.ListIndex = 1 Then
        C1Elastic2.Visible = False
        C1Elastic5.Visible = True
End If

End Sub

Private Sub cbStop1_Click()
If cbStop1.ListIndex = 0 Then
        C1Elastic2.Visible = True
        C1Elastic5.Visible = False
        
ElseIf cbStop1.ListIndex = 1 Then
        C1Elastic2.Visible = False
        C1Elastic5.Visible = True
End If
End Sub

Private Sub cbStopType_Change()

If cbStopType.ListIndex = 0 Then
    ToDate.Enabled = True
    ToDateH.Enabled = True
ElseIf cbStopType.ListIndex = 1 Then
    ToDate.Enabled = False
    ToDateH.Enabled = False
End If

End Sub

Private Sub cbStopType_Click()
If cbStopType.ListIndex = 0 Then
    ToDate.Enabled = True
    ToDateH.Enabled = True
ElseIf cbStopType.ListIndex = 1 Then
    ToDate.Enabled = False
    ToDateH.Enabled = False
End If

End Sub

Private Sub Check2_Click()
Dim value As Boolean
'value = Ch
End Sub

Private Sub chk2_Click()
Dim value As Boolean, i As Integer

value = chk2.value
With GRID2
For i = 1 To GRID2.Rows - 1
        .TextMatrix(i, .ColIndex("check")) = value
Next
End With
End Sub

Private Sub chk3_Click()
Dim value As Boolean, i As Integer

value = chk3.value
With Grid3
For i = 1 To Grid3.Rows - 1
        .TextMatrix(i, .ColIndex("check")) = value
Next
End With
End Sub

Private Sub chk4_Click()
Dim value As Boolean, i As Integer

value = chk4.value
With Grid4
For i = 1 To .Rows - 1
        .TextMatrix(i, .ColIndex("check")) = value
Next
End With
End Sub

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            txtID.Text = CStr(new_id("TblStopDealing", "ID", "", True))
           ' txtName.SetFocus
           cbStop1.ListIndex = 0
           
           opt_Stop.value = True
           opt_final.value = True
           GRID2.Rows = GRID2.FixedRows
           Grid3.Rows = Grid3.FixedRows
           Grid4.Rows = Grid4.FixedRows
        Case 1
                                             If ChekClodePeriod(Me.FromDate.value) = True Then
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

        Case 2

              If ChekClodePeriod(Me.FromDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            SaveData

        Case 3
            Undo

        Case 4
                                             If ChekClodePeriod(Me.FromDate.value) = True Then
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

            Del_Company

        Case 5

        Case 6
            Unload Me
         Case 7
         print_report
   '      print_report2
   Case 9
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "stopdealing"
            FrmSearch_BasicData.show
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
                        On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtID, "15062020005"


End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub dcDiscAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 291115
    End If
    
End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then

       Unload FrmSearch_MinistryContract
        FrmSearch_MinistryContract.SendForm = "StopDeal"
        FrmSearch_MinistryContract.show
        
End If

End Sub

Private Sub dcCustomer_Change()


Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     TxtRecordNo.Text = recordno
     TxtFullcode.Text = Fullcode


Dim StrSQL As String
 Set Rs_Temp = New ADODB.Recordset
 Set dcar.RowSource = Rs_Temp
    
  StrSQL = "  SELECT  ID , BoardNo from TblVendorCars  where  customerID = " & val(dcCustomer.BoundText) & "   ORDER BY ID "
    fill_combo dcar, StrSQL
dcar.Refresh
End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmCompanySearch
        FrmCompanySearch.lblSearchtype = "2030"
        FrmCompanySearch.show vbModal
End If
End Sub

Private Sub dcCustomer1_Change()
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer1.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer1.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordNo1.Text = recordno
     txtFullCode1.Text = Fullcode

End Sub

Private Sub dcCustomer1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmCompanySearch
        FrmCompanySearch.lblSearchtype = "2025"
        FrmCompanySearch.show vbModal
End If

End Sub

Private Sub dcCustomer2_Change()
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer2.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer2.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordNo2.Text = recordno
     txtFullCode2.Text = Fullcode
End Sub

Private Sub dcCustomer3_Change()
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer3.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer3.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordNo3.Text = recordno
     TxtFullcode3.Text = Fullcode

End Sub


Private Sub dcCustomer4_Change()
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer4.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer4.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordNo4.Text = recordno
     TxtFullcode4.Text = Fullcode
End Sub


Private Sub dcMinistry_Change()
    Dim str As String
    str = " select * From TblAttributionContract where idac = " & val(dcMinistry.BoundText)
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp2.RecordCount > 0 Then
        Rs_Temp2.MoveFirst '
        dcCustomer1.BoundText = IIf(IsNull(Rs_Temp2("VendorID").value), "", Rs_Temp2("VendorID").value)
     End If


End Sub

Private Sub dcMinistry_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmSearch_MinistryContract
        FrmSearch_MinistryContract.SendForm = "StopDeal"
        FrmSearch_MinistryContract.show
End If

End Sub

Private Sub dcMinistry2_Click(Area As Integer)
 Dim str As String
    str = " select * From TblAttributionContract where idac = " & val(dcMinistry2.BoundText)
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp2.RecordCount > 0 Then
        Rs_Temp2.MoveFirst '
        dcCustomer2.BoundText = IIf(IsNull(Rs_Temp2("VendorID").value), "", Rs_Temp2("VendorID").value)
     End If
     
     Fill_With_Cars2
End Sub

Private Sub dcMinistry3_Click(Area As Integer)
 Dim str As String
    str = " select * From TblAttributionContract where idac = " & val(dcMinistry3.BoundText)
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp2.RecordCount > 0 Then
        Rs_Temp2.MoveFirst '
        dcCustomer3.BoundText = IIf(IsNull(Rs_Temp2("VendorID").value), "", Rs_Temp2("VendorID").value)
     End If
     
     Fill_With_Cars3
End Sub

Private Sub dcMinistry4_Click(Area As Integer)
 Dim str As String
    str = " select * From TblAttributionContract where idac = " & val(dcMinistry4.BoundText)
    Set Rs_Temp2 = New ADODB.Recordset
    Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp2.RecordCount > 0 Then
        Rs_Temp2.MoveFirst '
        dcCustomer4.BoundText = IIf(IsNull(Rs_Temp2("VendorID").value), "", Rs_Temp2("VendorID").value)
     End If
     
     Fill_With_Cars4
End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    'Dcombos.GetAccountingCodes Me.dcDiscAccount, True, , 3

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  «Ìﬁ«› «· ⁄«„·  "
    LogTexte = " Open Window " & "  Violation Types "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
       
       
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
   ' Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
   
   
    Resize_Form Me
    
    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblStopDealing "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
     Dim str As String
     str = " select idac , idac  from TblAttributionContract "
     fill_combo dcMinistry, str
     fill_combo dcMinistry2, str
     fill_combo dcMinistry3, str
     fill_combo dcMinistry4, str
    
     Dcombos.GetCustomersSuppliers 2, dcCustomer
     Dcombos.GetCustomersSuppliers 2, dcCustomer1
     Dcombos.GetCustomersSuppliers 2, dcCustomer2
     Dcombos.GetCustomersSuppliers 2, dcCustomer3
     Dcombos.GetCustomersSuppliers 2, dcCustomer4
      
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«   «Ìﬁ«› «· ⁄«„·  "
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






Private Sub FromDate_Change()
       FromDateH.value = ToHijriDate(FromDate.value)
End Sub

Private Sub Fromdateh_LostFocus()
VBA.Calendar = vbCalGreg
            FromDate.value = ToGregorianDate(FromDateH.value)
End Sub

Private Sub opt_Cancel_Click()
cCancel.Visible = True
cStop.Visible = False
End Sub

Private Sub opt_canceled_Click()
Ele(2).Visible = False
Frame1.Visible = False
C1Elastic8.Visible = True
End Sub

Private Sub opt_final_Click()
Ele(2).Visible = True
Frame1.Visible = False
C1Elastic8.Visible = False
End Sub

Private Sub opt_Stop_Click()
cStop.Visible = True
cCancel.Visible = False
End Sub



Private Sub opt_temp_Click()
Ele(2).Visible = False
Frame1.Visible = True
C1Elastic8.Visible = False
End Sub

Private Sub StopDate1_Change()
          StopDateH1.value = ToHijriDate(StopDate1.value)
End Sub


Private Sub StopDateH1_LostFocus()
  VBA.Calendar = vbCalGreg
            StopDate1.value = ToGregorianDate(StopDateH1.value)
End Sub



Private Sub ToDate_Change()
  ToDateH.value = ToHijriDate(ToDate.value)
End Sub

Private Sub ToDateH_LostFocus()
        VBA.Calendar = vbCalGreg
        ToDate.value = ToGregorianDate(ToDateH.value)
End Sub

Private Sub txtBoxNo_Change()

End Sub

Private Sub txtfullcode_Change()
Dim val1, val2
If TxtFullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & TxtFullcode & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
        TxtRecordNo.Text = ""
        dcCustomer.BoundText = ""
    End If
    
    TxtRecordNo.Text = recordno
    dcCustomer.BoundText = CusID
End Sub

Private Sub txtFullCode1_Change()
Dim val1, val2
If TxtFullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & txtFullCode1.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
        txtRecordNo1.Text = ""
        dcCustomer1.BoundText = ""
    End If
    
    txtRecordNo1.Text = recordno
    dcCustomer1.BoundText = CusID
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «Ìﬁ«› «· ⁄«„·"
            Else
                Me.Caption = "Violation Types"
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
        
            Me.txtID.locked = True
           ' Me.txtName.locked = True
          '  Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            C1Elastic2.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «Ìﬁ«› «· ⁄«„· ( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «Ìﬁ«› «· ⁄«„·( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            Me.txtID.locked = True
           ' Me.txtName.locked = False
            C1Elastic2.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «Ìﬁ«› «· ⁄«„· (  ⁄œÌ· )"
            Else
                Me.Caption = "Violation Types(Edit)"
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
        
            Me.txtID.locked = True
          '  Me.txtName.locked = False
       '     Me.XPMTxtRemark.locked = False
            C1Elastic2.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

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

        GRID2.Rows = GRID2.FixedRows
        Grid3.Rows = Grid3.FixedRows
        Grid4.Rows = Grid4.FixedRows
        
        txtID.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
        dcCustomer.BoundText = IIf(IsNull(rs("CustomerID").value), "", rs("CustomerID").value)
        dcar.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
        cbStopType.ListIndex = IIf(IsNull(rs("Stop").value), -1, rs("Stop").value)
        FromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
        ToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
        FromDateH.value = IIf(IsNull(rs("FromDateh").value), Date, rs("FromDateh").value)
        ToDateH.value = IIf(IsNull(rs("ToDateh").value), Date, rs("ToDateh").value)
      
        cbStop1.ListIndex = IIf(IsNull(rs("StopM").value), -1, rs("StopM").value)
        dcMinistry.BoundText = IIf(IsNull(rs("AttribID").value), "", rs("AttribID").value)
        dcCustomer1.BoundText = IIf(IsNull(rs("CustomerID1").value), "", rs("CustomerID1").value)
        txtFullCode1.Text = IIf(IsNull(rs("CFullCode1").value), "", rs("CFullCode1").value)
        txtRecordNo1.Text = IIf(IsNull(rs("CRecordNo1").value), "", rs("CRecordNo1").value)
        StopDate1.value = IIf(IsNull(rs("StopDate1").value), Date, rs("StopDate1").value)
        StopDateH1.value = IIf(IsNull(rs("StopDateH1").value), ToHijriDate(Date), rs("StopDateH1").value)
               
         dcMinistry2.BoundText = IIf(IsNull(rs("AttribID2").value), "", rs("AttribID2").value)
        dcCustomer2.BoundText = IIf(IsNull(rs("CustomerID2").value), "", rs("CustomerID2").value)
        txtFullCode2.Text = IIf(IsNull(rs("CFullCode2").value), "", rs("CFullCode2").value)
        txtRecordNo2.Text = IIf(IsNull(rs("CRecordNo2").value), "", rs("CRecordNo2").value)
        StopDate2.value = IIf(IsNull(rs("StopDate2").value), Date, rs("StopDate2").value)
        StopDateH2.value = IIf(IsNull(rs("StopDateH2").value), ToHijriDate(Date), rs("StopDateH2").value)
        
        
         dcMinistry3.BoundText = IIf(IsNull(rs("AttribID3").value), "", rs("AttribID3").value)
        dcCustomer3.BoundText = IIf(IsNull(rs("CustomerID3").value), "", rs("CustomerID3").value)
        TxtFullcode3.Text = IIf(IsNull(rs("CFullCode3").value), "", rs("CFullCode3").value)
        txtRecordNo3.Text = IIf(IsNull(rs("CRecordNo3").value), "", rs("CRecordNo3").value)
        StopDate3.value = IIf(IsNull(rs("StopDate3").value), Date, rs("StopDate3").value)
        StopDateH3.value = IIf(IsNull(rs("StopDateH3").value), ToHijriDate(Date), rs("StopDateH3").value)
        
         dcMinistry4.BoundText = IIf(IsNull(rs("AttribID4").value), "", rs("AttribID4").value)
        dcCustomer4.BoundText = IIf(IsNull(rs("CustomerID4").value), "", rs("CustomerID4").value)
        TxtFullcode4.Text = IIf(IsNull(rs("CFullCode4").value), "", rs("CFullCode4").value)
        txtRecordNo4.Text = IIf(IsNull(rs("CRecordNo4").value), "", rs("CRecordNo4").value)
        StopDate4.value = IIf(IsNull(rs("StopDate4").value), Date, rs("StopDate4").value)
        StopDateH4.value = IIf(IsNull(rs("StopDateH4").value), ToHijriDate(Date), rs("StopDateH4").value)
        
        
        opt_Stop.value = IIf(IsNull(rs("stp").value), False, rs("stp").value)
        opt_Cancel.value = IIf(IsNull(rs("cancl").value), False, rs("cancl").value)
              
        
        Dim rr As Integer
        
        rr = IIf(IsNull(rs("StopDealingType").value), 1, rs("StopDealingType").value)
        
        If rr = 1 Then
        ElseIf rr = 2 Then
        ElseIf rr = 3 Then
        End If
        If rr = 1 Then
                opt_final.value = True
        ElseIf rr = 2 Then
                opt_temp.value = True
        ElseIf rr = 3 Then
                opt_canceled.value = True
        End If
       
       Dim ss As String, mm As Integer
       ss = "  select h.carid , h.hid , d.boardno , h.IDAC_D from  tblvendorcars d ,  TblStopDealing_Details h where d.id = h.carid and hid =  " & val(txtID.Text)
       
       If opt_Stop.value = True Then
            If rr = 2 Then
                    Set Rs_Temp1 = New ADODB.Recordset
                    Rs_Temp1.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If Rs_Temp1.RecordCount > 0 Then
                             With GRID2
                                    .Rows = .FixedRows + Rs_Temp1.RecordCount
                                    For mm = 1 To .Rows - 1
                                               .TextMatrix(mm, .ColIndex("id")) = IIf(IsNull(Rs_Temp1("carid").value), "", Rs_Temp1("carid").value)
                                               .TextMatrix(mm, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp1("BoardNo").value), "", Rs_Temp1("BoardNo").value)
                                               .TextMatrix(mm, .ColIndex("IDAC_D")) = IIf(IsNull(Rs_Temp1("IDAC_D").value), "", Rs_Temp1("IDAC_D").value)
                                               .TextMatrix(mm, .ColIndex("check")) = 1
                                    Next
                             End With
                    End If
            ElseIf rr = 3 Then
                     Set Rs_Temp1 = New ADODB.Recordset
                    Rs_Temp1.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If Rs_Temp1.RecordCount > 0 Then
                             With Grid3
                                    .Rows = .FixedRows + Rs_Temp1.RecordCount
                                    For mm = 1 To .Rows - 1
                                               .TextMatrix(mm, .ColIndex("id")) = IIf(IsNull(Rs_Temp1("carid").value), "", Rs_Temp1("carid").value)
                                               .TextMatrix(mm, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp1("BoardNo").value), "", Rs_Temp1("BoardNo").value)
                                               .TextMatrix(mm, .ColIndex("IDAC_D")) = IIf(IsNull(Rs_Temp1("IDAC_D").value), "", Rs_Temp1("IDAC_D").value)
                                               .TextMatrix(mm, .ColIndex("check")) = 1
                                    Next
                             End With
                    End If
             End If
      ElseIf opt_Cancel.value = True Then
            Set Rs_Temp1 = New ADODB.Recordset
            Rs_Temp1.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp1.RecordCount > 0 Then
                     With Grid4
                            .Rows = .FixedRows + Rs_Temp1.RecordCount
                            For mm = 1 To .Rows - 1
                                       .TextMatrix(mm, .ColIndex("id")) = IIf(IsNull(Rs_Temp1("carid").value), "", Rs_Temp1("carid").value)
                                       .TextMatrix(mm, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp1("BoardNo").value), "", Rs_Temp1("BoardNo").value)
                                       .TextMatrix(mm, .ColIndex("IDAC_D")) = IIf(IsNull(Rs_Temp1("IDAC_D").value), "", Rs_Temp1("IDAC_D").value)
                                       .TextMatrix(mm, .ColIndex("check")) = 1
                            Next
                     End With
            End If
        End If
       
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub




Private Sub TxtName_GotFocus()
On Error Resume Next
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
 SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtRecordNo_Change()
Dim val1, val2, CusID As String, Fullcode As String
If TxtRecordNo.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & TxtRecordNo.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcCustomer.BoundText = ""
        TxtFullcode.Text = ""
    End If
    
   dcCustomer.BoundText = CusID
   TxtFullcode.Text = Fullcode
End Sub

Private Sub txtRecordNo1_Change()
Dim val1, val2, CusID As String, Fullcode As String
If TxtRecordNo.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & txtRecordNo1.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcCustomer1.BoundText = ""
        txtFullCode1.Text = ""
    End If
    
   dcCustomer1.BoundText = CusID
   txtFullCode1.Text = Fullcode

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
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

Function CuurentLogdata(Optional Currentmode As String)
 
 
End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
     
        
                
        If opt_Stop = True Then
                If opt_temp = True Then
                        If IS_Grid2_SelectedRow = False Then
                                MsgBox ("«Œ — «·„⁄œ« /«·”Ì«—«  «Ê·«")
                                Exit Sub
                        End If
                ElseIf opt_final.value = True Then
                           If cbStop1.ListIndex = 0 Then
                                    If dcMinistry.BoundText = "" Then
                                            MsgBox ("  „‰ ›÷·ﬂ «Œ — ⁄ﬁœ «·«”‰«œ «Ê·« ")
                                            Exit Sub
                                    End If
                            ElseIf cbStop1.ListIndex = 1 Then
                                    If dcCustomer.BoundText = "" Then
                                        MsgBox "„‰ ›÷·ﬂ «Œ — «·„ ⁄Âœ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                        dcCustomer.SetFocus
                                        Exit Sub
                                    End If
                                
                                    If dcar.BoundText = "" Then
                                        MsgBox "„‰ ›÷·ﬂ «Œ — «·„⁄œÂ/«·”Ì«—… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                        dcar.SetFocus
                                        Exit Sub
                                    End If
                              End If
                ElseIf opt_canceled.value = True Then
                        If IS_Grid3_SelectedRow = False Then
                                MsgBox ("«Œ — «·„⁄œ« /«·”Ì«—«  «Ê·«")
                                Exit Sub
                        End If
                End If
        End If
        
        
        
       Select Case Me.TxtModFlg.Text
            Case "N"
                If cbStop1.ListIndex = 0 Then
                            StrSQL = " select * from tblstopdealing where AttribID = " & val(dcMinistry.BoundText)
                            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            If RsTemp.RecordCount > 0 Then
                                    Msg = " „ «Ìﬁ«› Â–« «·⁄ﬁœ „”»ﬁ« " & CHR(13)
                                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·⁄ﬁœ «·’ÕÌÕ " & CHR(13)
                                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                    dcMinistry.SetFocus
                                    Exit Sub
                            End If
                ElseIf cbStop1.ListIndex = 1 Then
                            StrSQL = " select * from tblstopdealing where carid = " & val(dcar.BoundText)
                            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            If RsTemp.RecordCount > 0 Then
                                    Msg = " „ «Ìﬁ«› Â–Â «·„⁄œÂ/«·”Ì«—… „”»ﬁ« " & CHR(13)
                                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·„⁄œÂ/«·”Ì«—… " & CHR(13)
                                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                    dcar.SetFocus
                                    Exit Sub
                            End If
                End If
            Case "E"
                If cbStop1.ListIndex = 1 Then
                            StrSQL = " select * from tblstopdealing where carid = " & val(dcar.BoundText)
                            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            If RsTemp.RecordCount > 0 Then
                                If RsTemp("ID").value <> val(txtID.Text) Then
                                        Msg = " „ «Ìﬁ«› Â–Â «·„⁄œÂ/«·”Ì«—… „”»ﬁ« " & CHR(13)
                                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·„⁄œÂ/«·”Ì«—… " & CHR(13)
                                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                        dcar.SetFocus
                                        Exit Sub
                                 End If
                            End If
                ElseIf cbStop1.ListIndex = 0 Then
                            StrSQL = " select * from tblstopdealing where AttribID = " & val(dcMinistry.BoundText)
                            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            If RsTemp.RecordCount > 0 Then
                                If RsTemp("ID").value <> val(txtID.Text) Then
                                        Msg = " „ «Ìﬁ«› Â–Â «·⁄ﬁœ „”»ﬁ« " & CHR(13)
                                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·⁄ﬁœ " & CHR(13)
                                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                        dcMinistry.SetFocus
                                        Exit Sub
                                 End If
                            End If
                End If
        End Select
        
         
        Select Case Me.TxtModFlg.Text
           Case "N"
            rs.AddNew
            txtID.Text = CStr(new_id("TblStopDealing", "ID", "", True))
            Case "E"
             
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        If TxtModFlg.Text = "" Then
                   Cancel_StopDealing
        End If
          
        rs("ID").value = val(txtID.Text)
       
        rs("CustomerID").value = IIf(dcCustomer.BoundText = "", Null, dcCustomer.BoundText)
        rs("CarID").value = IIf(dcar.BoundText = "", Null, dcar.BoundText)
        rs("Stop").value = IIf(cbStopType.ListIndex = -1, Null, cbStopType.ListIndex)
        rs("FromDate").value = FromDate.value
        rs("ToDate").value = ToDate.value
        rs("FromDateH").value = FromDateH.value
        rs("ToDateH").value = ToDateH.value
        rs("StopM").value = IIf(cbStop1.ListIndex = -1, Null, cbStop1.ListIndex)
        
        rs("AttribID").value = IIf(dcMinistry.BoundText = "", Null, dcMinistry.BoundText)
        rs("CustomerID1").value = IIf(dcCustomer1.BoundText = "", Null, dcCustomer1.BoundText)
        rs("CFullCode1").value = txtFullCode1.Text
        rs("CRecordNo1").value = txtRecordNo1.Text
        rs("StopDate1").value = StopDate1.value
        rs("StopDateH1").value = StopDateH1.value
           
        rs("AttribID2").value = IIf(dcMinistry2.BoundText = "", Null, dcMinistry2.BoundText)
        rs("CustomerID2").value = IIf(dcCustomer2.BoundText = "", Null, dcCustomer2.BoundText)
        rs("CFullCode2").value = txtFullCode2.Text
        rs("CRecordNo2").value = txtRecordNo2.Text
        rs("StopDate2").value = StopDate2.value
        rs("StopDateH2").value = StopDateH2.value
        
        
        rs("AttribID3").value = IIf(dcMinistry3.BoundText = "", Null, dcMinistry3.BoundText)
        rs("CustomerID3").value = IIf(dcCustomer3.BoundText = "", Null, dcCustomer3.BoundText)
        rs("CFullCode3").value = TxtFullcode3.Text
        rs("CRecordNo3").value = txtRecordNo3.Text
        rs("StopDate3").value = StopDate3.value
        rs("StopDateH3").value = StopDateH3.value
        
        rs("AttribID4").value = IIf(dcMinistry4.BoundText = "", Null, dcMinistry4.BoundText)
        rs("CustomerID4").value = IIf(dcCustomer4.BoundText = "", Null, dcCustomer4.BoundText)
        rs("CFullCode4").value = TxtFullcode4.Text
        rs("CRecordNo4").value = txtRecordNo4.Text
        rs("StopDate4").value = StopDate4.value
        rs("StopDateH4").value = StopDateH4.value
        
        rs("stp").value = opt_Stop.value
        rs("cancl").value = opt_Cancel.value
        
        If opt_Stop.value = True Then
                If opt_final.value = True Then
                        rs("StopDealingType").value = 1
                ElseIf opt_temp.value = True Then
                        rs("StopDealingType").value = 2
                ElseIf opt_canceled.value = True Then
                        rs("StopDealingType").value = 3
                End If
        End If
        rs.update
                
        '---------  Temperary Stoping
        Dim d As Integer, l As String
          If opt_Stop = True Then '«·«Ìﬁ«›
                
                
                
                If opt_temp = True Then '„ƒﬁ 
                        
                        Set Rs_Temp = New ADODB.Recordset
                        Rs_Temp.Open "TblStopDealing_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                        For d = 1 To GRID2.Rows - 1
                                If GRID2.Cell(flexcpChecked, d, GRID2.ColIndex("check")) = flexChecked And GRID2.TextMatrix(d, GRID2.ColIndex("ID")) <> "" Then
                                        Rs_Temp.AddNew
                                        Rs_Temp("ID").value = new_id("TblStopDealing_Details", "ID", "", True)
                                        Rs_Temp("HID").value = val(txtID.Text)
                                        Rs_Temp("CarID").value = val(GRID2.TextMatrix(d, GRID2.ColIndex("ID")))
                                        Rs_Temp("IDAC_D").value = val(GRID2.TextMatrix(d, GRID2.ColIndex("IDAC_D")))
                                        Rs_Temp.update
                                             
                                        Set Rs_Temp1 = New ADODB.Recordset ' «·„⁄œ« /«·”Ì«—«  ··«Ìﬁ«› «·„ƒﬁ 
                                        l = " select * from tblvehicleallocation_details where id =  " & val(GRID2.TextMatrix(d, GRID2.ColIndex("IDAC_D")))
                                        Rs_Temp1.Open l, Cn, adOpenStatic, adLockOptimistic, adCmdText
                                        If Rs_Temp1.RecordCount > 0 Then
                                                Rs_Temp1("StartStopDealing").value = StopDate2.value
                                                Rs_Temp1("StartStopDealingH").value = StopDateH2.value
                                                Rs_Temp1("StopDealingID").value = val(txtID.Text)
                                                Rs_Temp1.update
                                        End If
                                                                                
                                End If
                        Next
                       
                ElseIf opt_canceled.value = True Then '«·€«¡  «Ìﬁ«› „ƒﬁ 
                        
                        Set Rs_Temp = New ADODB.Recordset
                        Rs_Temp.Open "TblStopDealing_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                        For d = 1 To Grid3.Rows - 1
                                If Grid3.Cell(flexcpChecked, d, Grid3.ColIndex("check")) = flexChecked And Grid3.TextMatrix(d, Grid3.ColIndex("ID")) <> "" Then
                                        Rs_Temp.AddNew
                                        Rs_Temp("ID").value = new_id("TblStopDealing_Details", "ID", "", True)
                                        Rs_Temp("HID").value = val(txtID.Text)
                                        Rs_Temp("CarID").value = val(Grid3.TextMatrix(d, Grid3.ColIndex("ID")))
                                        Rs_Temp("IDAC_D").value = val(Grid3.TextMatrix(d, Grid3.ColIndex("IDAC_D")))
                                        Rs_Temp.update
                                        
                                       Set Rs_Temp1 = New ADODB.Recordset
                                        l = " select * from tblvehicleallocation_details where id =  " & val(Grid3.TextMatrix(d, Grid3.ColIndex("IDAC_D")))
                                        Rs_Temp1.Open l, Cn, adOpenStatic, adLockOptimistic, adCmdText
                                        If Rs_Temp1.RecordCount > 0 Then
                                                Rs_Temp1("stoped").value = True
                                                Rs_Temp1("StopDealingID").value = val(txtID.Text)
                                                Rs_Temp1.update
                                        End If
                                        
                                End If
                        Next
                ElseIf opt_final.value = True Then '‰Â«∆Ì  ' «Ìﬁ«› ⁄ﬁœ
                        '//////////////////////////////////////  final
                        If cbStop1.ListIndex = 0 Then
                        Dim m As Integer
                        Set Rs_Temp2 = New ADODB.Recordset
                        StrSQL = "SELECT  *  From TblVendorCars where CustomerID  = " & val(dcCustomer1.BoundText)
                        Rs_Temp2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If Rs_Temp2.RecordCount > 0 Then
                                For m = 0 To Rs_Temp2.RecordCount - 1
                                '        Rs_Temp2("StopDeal").value = True
                                '        Rs_Temp2("StopDate").value = StopDate1.value
                                '        Rs_Temp2("StopDateH").value = StopDateH1.value
                                '        Rs_Temp2("stopdealingID").value = val(txtid.Text)
                                '        Rs_Temp2.update
                                '        Rs_Temp2.MoveNext
                                 Next
                        End If
                        Set Rs_Temp = New ADODB.Recordset '
                        StrSQL = "SELECT  *  From TblAttributionContract where idac = " & val(dcMinistry.BoundText)
                        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If Rs_Temp.RecordCount > 0 Then
                                  Rs_Temp("StopDeal").value = True
                                  Rs_Temp("StopDate").value = StopDate1.value
                                  Rs_Temp("StopDateH").value = StopDateH1.value
                                  Rs_Temp("stopdealingID").value = val(txtID.Text)
                                  Rs_Temp.update
                        End If
                    ElseIf cbStop1.ListIndex = 1 Then '‰Â«∆Ì ' «Ìﬁ«› ”Ì«—…
                        Set Rs_Temp = New ADODB.Recordset
                        StrSQL = "SELECT  *  From TblVendorCars where id = " & val(dcar.BoundText)
                        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If Rs_Temp.RecordCount > 0 Then
                                 Rs_Temp("StopDeal").value = True
                                  Rs_Temp("StopDate").value = FromDate.value
                                  Rs_Temp("StopDateH").value = FromDateH.value
                                  Rs_Temp("stopdealingID").value = val(txtID.Text)
                                  Rs_Temp.update
                        End If
                    End If
                End If
          ElseIf opt_Cancel.value = True Then '«·€«¡
          
                        Set Rs_Temp = New ADODB.Recordset
                        Rs_Temp.Open "TblStopDealing_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                        For d = 1 To Grid4.Rows - 1
                                If Grid4.Cell(flexcpChecked, d, Grid4.ColIndex("check")) = flexChecked And Grid4.TextMatrix(d, Grid4.ColIndex("ID")) <> "" Then
                                        Rs_Temp.AddNew
                                        Rs_Temp("ID").value = new_id("TblStopDealing_Details", "ID", "", True)
                                        Rs_Temp("HID").value = val(txtID.Text)
                                        Rs_Temp("CarID").value = val(Grid4.TextMatrix(d, Grid4.ColIndex("ID")))
                                        Rs_Temp("IDAC_D").value = val(Grid4.TextMatrix(d, Grid4.ColIndex("IDAC_D")))
                                        Rs_Temp.update
                                        
                                        Set Rs_Temp1 = New ADODB.Recordset
                                        l = " select * from tblvehicleallocation_details where id =  " & val(Grid4.TextMatrix(d, Grid4.ColIndex("IDAC_D")))
                                        Rs_Temp1.Open l, Cn, adOpenStatic, adLockOptimistic, adCmdText
                                        If Rs_Temp1.RecordCount > 0 Then
                                                Rs_Temp1("EndStopDealing").value = StopDate4.value
                                                Rs_Temp1("EndStopDealingH").value = StopDateH4.value
                                                Rs_Temp1("StopDealingID").value = val(txtID.Text)
                                                Rs_Temp1.update
                                        End If
                                End If
                        Next
           End If
                
        

        
        
           
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  «Ìﬁ«› «· ⁄«„· " & CHR(13)
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(txtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtID.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«   «Ìﬁ«› «· ⁄«„· —ﬁ„ " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                   
                    Cancel_StopDealing
                   
                   StrSQL = "delete From TblStopDealing_details  where  HID =" & val(txtID.Text)
                   Cn.Execute StrSQL, , adExecuteNoRecords: Del
                   
                   StrSQL = "delete From TblStopDealing where  ID =" & val(txtID.Text)
                   Cn.Execute StrSQL, , adExecuteNoRecords: Del
                   
                   
                   StrSQL = "SELECT  *  From TblStopDealing"
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
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
        .Create Me.hwnd, "»Ì«‰«  «‰Ê«⁄ «·„Œ«·›« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „Œ«·›… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  „Œ«·›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  „Œ«·›… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·„Œ«·›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„‰«ÿﬁ «·«œ«—Ì…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ «·„Œ«·›…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   «·„Œ«·›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub


Function print_report(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



'MySQL = MySQL & "  SELECT ID, IDAC, CusID, CusName, CusNamee, Fullcode, RecordNo, StopDate, StopDateH, BoardNo, SM"
'MySQL = MySQL & "  FROM     (SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
''MySQL = MySQL & "  dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate1 AS StopDate, dbo.TblStopDealing.StopDateH1 AS StopDateH,"
'MySQL = MySQL & "  '' AS BoardNo, '⁄ﬁœ' AS SM"
'MySQL = MySQL & "  FROM      dbo.TblStopDealing INNER JOIN"
'MySQL = MySQL & "  dbo.TblCustemers ON dbo.TblStopDealing.CustomerID1 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'MySQL = MySQL & "  dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID = dbo.TblAttributionContract.IDAC"
'MySQL = MySQL & "  Where (dbo.TblStopDealing.StopM = 0)"
'MySQL = MySQL & "  Union"
''MySQL = MySQL & "  SELECT TblStopDealing_1.ID, '' AS IDAC, TblCustemers_1.CusID, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode,"
'MySQL = MySQL & "  TblCustemers_1.RecordNo, TblStopDealing_1.FromDate AS StopDate, TblStopDealing_1.FromDateH AS StopDateH, dbo.TblVendorCars.BoardNo, '”Ì«—…' AS SM"
'MySQL = MySQL & "  FROM     dbo.TblStopDealing AS TblStopDealing_1 LEFT OUTER JOIN"
'MySQL = MySQL & "  dbo.TblCustemers AS TblCustemers_1 ON TblStopDealing_1.CustomerID = TblCustemers_1.CusID LEFT OUTER JOIN"
'MySQL = MySQL & "  dbo.TblVendorCars ON TblStopDealing_1.CarID = dbo.TblVendorCars.ID"
'MySQL = MySQL & "  WHERE  (TblStopDealing_1.StopM = 1)) AS tbl1"
'MySQL = MySQL & "  Where (1 = 1)"

MySQL = MySQL & "    select * from ("
          
MySQL = MySQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate1 AS StopDate, dbo.TblStopDealing.StopDateH1 AS StopDateH,"
MySQL = MySQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
MySQL = MySQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID1 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID = dbo.TblAttributionContract.IDAC"
MySQL = MySQL & "    Where (dbo.TblStopDealing.StopM = 0)"
MySQL = MySQL & "    Union"


MySQL = MySQL & "    SELECT dbo.TblStopDealing.ID, '' AS IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "    dbo.TblCustemers.RecordNo,  dbo.TblStopDealing.FromDate AS StopDate, dbo.TblStopDealing.FromDateH AS StopDateH,         dbo.TblVendorCars.BoardNo, '”Ì«—…' AS SM"
MySQL = MySQL & "    FROM     dbo.TblStopDealing LEFT OUTER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID = dbo.TblCustemers.CusID"
MySQL = MySQL & "    LEFT OUTER JOIN           dbo.TblVendorCars ON dbo.TblStopDealing.CarID = dbo.TblVendorCars.ID"
MySQL = MySQL & "    Where (dbo.TblStopDealing.StopM = 1)"

MySQL = MySQL & "    Union"

MySQL = MySQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate2 AS StopDate, dbo.TblStopDealing.StopDateH2 AS StopDateH,"
MySQL = MySQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
MySQL = MySQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID2 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID2 = dbo.TblAttributionContract.IDAC"
MySQL = MySQL & "    Where (dbo.TblStopDealing.stp = 1 And StopDealingType = 2)"

MySQL = MySQL & "    Union"

MySQL = MySQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate3 AS StopDate, dbo.TblStopDealing.StopDateH3 AS StopDateH,"
MySQL = MySQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
MySQL = MySQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID3 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID3 = dbo.TblAttributionContract.IDAC"
MySQL = MySQL & "    Where (dbo.TblStopDealing.stp = 1 And StopDealingType = 3)"


MySQL = MySQL & "    Union"

MySQL = MySQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate4 AS StopDate, dbo.TblStopDealing.StopDateH4 AS StopDateH,"
MySQL = MySQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
MySQL = MySQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID4 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID4 = dbo.TblAttributionContract.IDAC"
MySQL = MySQL & "    Where (dbo.TblStopDealing.stp = 1 And StopDealingType = 4)"

MySQL = MySQL & "    ) tb1 where 1= 1"
       
       



  MySQL = MySQL & "   and id = " & val(txtID.Text)
     
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_StopDealing.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_StopDealing.rpt"
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
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
    End If
    
    
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

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Del()

   If cbStop1.ListIndex = 0 Then
            Dim m As Integer, i As Integer, StrSQL As String
          
            Set Rs_Temp = New ADODB.Recordset
            StrSQL = "SELECT  *  From TblAttributionContract where idac = " & val(dcMinistry.BoundText)
            Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp.RecordCount > 0 Then
                      Rs_Temp("StopDeal").value = Null
                      Rs_Temp("StopDate").value = Null
                      Rs_Temp("StopDateH").value = Null
                      Rs_Temp.update
            End If
                 
                 
           Set Rs_Temp1 = New ADODB.Recordset
            StrSQL = "SELECT  *  From TblVehicleAllocation_Details where idva = " & val(dcMinistry.BoundText)
            Rs_Temp1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp1.RecordCount > 0 Then
                 For i = 0 To Rs_Temp1.RecordCount - 1
                            Set Rs_Temp2 = New ADODB.Recordset
                            StrSQL = "SELECT  *  From Tblvendorcars where id = " & IIf(IsNull(Rs_Temp1("carid").value), 0, Rs_Temp1("carid").value)
                            Rs_Temp2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            If Rs_Temp2.RecordCount > 0 Then
                                    Rs_Temp2("StopDeal").value = Null
                                    Rs_Temp2("StopDate").value = Null
                                    Rs_Temp2("StopDateH").value = Null
                                    Rs_Temp2.update
                            End If
                 Next
            End If
                 
         
            
            
        
        ElseIf cbStop1.ListIndex = 1 Then
            Set Rs_Temp = New ADODB.Recordset
            StrSQL = "SELECT  *  From TblVendorCars where id = " & val(dcar.BoundText)
            Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp.RecordCount > 0 Then
                    Rs_Temp("StopDeal").value = Null
                     Rs_Temp("StopDate").value = Null
                      Rs_Temp("StopDateH").value = Null
                      Rs_Temp.update
            End If
        
  End If

End Sub

Public Function IS_Grid2_SelectedRow() As Boolean
    Dim slct As Boolean, i As Integer
    slct = False
    For i = 1 To GRID2.Rows - 1
            If GRID2.Cell(flexcpChecked, i, GRID2.ColIndex("check")) = flexChecked Then
                    IS_Grid2_SelectedRow = True
                    Exit Function
            End If
    Next
    IS_Grid2_SelectedRow = False
End Function

Public Function IS_Grid3_SelectedRow() As Boolean
    Dim slct As Boolean, i As Integer
    slct = False
    For i = 1 To Grid3.Rows - 1
            If Grid3.Cell(flexcpChecked, i, Grid3.ColIndex("check")) = flexChecked Then
                    IS_Grid3_SelectedRow = True
                    Exit Function
            End If
    Next
    IS_Grid3_SelectedRow = False
End Function


Public Function IS_Grid4_SelectedRow() As Boolean
    Dim slct As Boolean, i As Integer
    slct = False
    For i = 1 To Grid4.Rows - 1
           If Grid4.Cell(flexcpChecked, i, Grid4.ColIndex("check")) = flexChecked Then
                    IS_Grid4_SelectedRow = True
                    Exit Function
            End If
    Next
    IS_Grid4_SelectedRow = False
End Function

Private Sub Fill_With_Cars2()

Dim str As String, i As Integer
str = "select d.id , d.carid , h.boardNo from  tblvendorCars H ,  tblvehicleallocation_details  D  where  startstopdealing = '' and  h.id = d.carid and d.IDVA = " & val(dcMinistry2.BoundText)
Set Rs_Temp3 = New ADODB.Recordset
Rs_Temp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

With GRID2
If Rs_Temp3.RecordCount > 0 Then
    .Rows = .FixedRows + Rs_Temp3.RecordCount
    For i = 1 To .Rows - 1
           .TextMatrix(i, .ColIndex("IDAC_D")) = IIf(IsNull(Rs_Temp3("ID").value), "", Rs_Temp3("ID").value)
           .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_Temp3("carID").value), "", Rs_Temp3("carID").value)
           .TextMatrix(i, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp3("BoardNo").value), "", Rs_Temp3("BoardNo").value)
           Rs_Temp3.MoveNext
    Next
End If

End With
End Sub


Private Sub Fill_With_Cars3()

Dim str As String, i As Integer
str = "select  d.id , d.carid , h.boardNo from  tblvendorCars H ,  tblvehicleallocation_details  D  where ( stoped is null  or stoped = 0 ) and  h.id = d.carid and d.IDVA = " & val(dcMinistry3.BoundText)
Set Rs_Temp3 = New ADODB.Recordset
Rs_Temp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Grid3
If Rs_Temp3.RecordCount > 0 Then
    .Rows = .FixedRows + Rs_Temp3.RecordCount
    For i = 1 To .Rows - 1
           .TextMatrix(i, .ColIndex("IDAC_D")) = IIf(IsNull(Rs_Temp3("ID").value), "", Rs_Temp3("ID").value)
           .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_Temp3("carID").value), "", Rs_Temp3("carID").value)
           .TextMatrix(i, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp3("BoardNo").value), "", Rs_Temp3("BoardNo").value)
           Rs_Temp3.MoveNext
    Next
End If

End With
End Sub

Private Sub Fill_With_Cars4()

Dim str As String, i As Integer
str = "select  d.id , d.carid , h.boardNo from  tblvendorCars H ,  tblvehicleallocation_details  D  where stoped is null  and  startstopdealing  <> '' and endstopdealing = ''  and  h.id = d.carid and d.IDVA = " & val(dcMinistry4.BoundText)
Set Rs_Temp3 = New ADODB.Recordset
Rs_Temp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

With Grid4
.Rows = .FixedRows
If Rs_Temp3.RecordCount > 0 Then
    .Rows = .FixedRows + Rs_Temp3.RecordCount
    For i = 1 To .Rows - 1
           .TextMatrix(i, .ColIndex("IDAC_D")) = IIf(IsNull(Rs_Temp3("ID").value), "", Rs_Temp3("ID").value)
           .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs_Temp3("carID").value), "", Rs_Temp3("carID").value)
           .TextMatrix(i, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp3("BoardNo").value), "", Rs_Temp3("BoardNo").value)
           Rs_Temp3.MoveNext
    Next
End If

End With
End Sub


Private Sub Cancel_StopDealing()
    Dim StrSQL As String
    
    StrSQL = "update tblvendorcars set StopDate = null , StopDateH = null , StopDeal = null , stopdealingID = null where stopdealingID =" & val(txtID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords: Del
    
    StrSQL = "update TblAttributionContract set StopDate = null , StopDateH = null , StopDeal = null , stopdealingID = null where stopdealingID =" & val(txtID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords: Del
    
    StrSQL = " update tblvehicleallocation_details set StartStopDealing = null ,StartStopDealingh=null , EndStopDealing = null , EndStopDealingH = null ,stoped = null , stopdealingID = null  where stopdealingID = " & val(txtID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords: Del

End Sub






