VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "TASKPA~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "COMMAN~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Begin VB.Form FrmPane 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   0  'None
   Caption         =   "M"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt 
      Height          =   285
      Left            =   30
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   630
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   285
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   300
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox PicLeftPan 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   8160
      Left            =   810
      RightToLeft     =   -1  'True
      ScaleHeight     =   8160
      ScaleWidth      =   4185
      TabIndex        =   0
      Top             =   90
      Width           =   4185
      Begin XtremeTaskPanel.TaskPanel TaskPanel1 
         Height          =   705
         Left            =   1170
         TabIndex        =   1
         Top             =   1980
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   1244
         _StockProps     =   64
         VisualTheme     =   3
         ItemLayout      =   2
         HotTrackStyle   =   1
         BorderStyle     =   2
      End
      Begin C1SizerLibCtl.C1Elastic EleManSearch 
         Height          =   4185
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   4125
         _cx             =   7276
         _cy             =   7382
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
         Appearance      =   5
         MousePointer    =   0
         Version         =   801
         BackColor       =   14737632
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1980
            Index           =   0
            Left            =   30
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   510
            Width           =   4065
            _cx             =   7170
            _cy             =   3493
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   16777215
            ForeColor       =   128
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "⁄Ê«„· «·»ÕÀ"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   4
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   1
            TagSplit        =   2
            PicturePos      =   8
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
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "»ÕÀ „ÿ«»ﬁ"
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
               Height          =   285
               Index           =   1
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   600
               Width           =   1155
            End
            Begin MSDataListLib.DataCombo DcboItemName 
               Height          =   315
               Left            =   60
               TabIndex        =   35
               Top             =   1590
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.TextBox TxtSearch 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Index           =   3
               Left            =   60
               TabIndex        =   30
               Top             =   1260
               Width           =   2925
            End
            Begin MSDataListLib.DataCombo DcboCustomer 
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   29
               Top             =   900
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.TextBox TxtSearch 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   2
               Left            =   1290
               TabIndex        =   27
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox TxtSearch 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   1
               Left            =   30
               TabIndex        =   25
               Top             =   300
               Width           =   1155
            End
            Begin VB.TextBox TxtSearch 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   0
               Left            =   2070
               TabIndex        =   23
               Top             =   300
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·’‰›"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   1620
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   930
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«·”Ì—Ì«·"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   3090
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   600
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «· ﬂ "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1170
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   300
               Width           =   825
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ì’«· «·œŒÊ·"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   3090
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "»ÕÀ ⁄‰ ’Ì«‰… ’‰› «Ê ﬁÿ⁄…"
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
            Height          =   225
            Index           =   0
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   30
            Value           =   -1  'True
            Width           =   3525
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "»ÕÀ ⁄‰  Ã„Ì⁄ ÃÂ«“ ﬂ„»ÌÊ —"
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
            Height          =   225
            Index           =   1
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   270
            Width           =   3525
         End
         Begin ImpulseButton.ISButton CmdSeacrh 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   2505
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            ButtonStyle     =   1
            Caption         =   "»ÕÀ"
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
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8UCtl.VSFlexGrid FgSearch 
            Height          =   1320
            Left            =   30
            TabIndex        =   17
            Top             =   2835
            Width           =   4065
            _cx             =   7170
            _cy             =   2328
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
            WallPaperAlignment=   4
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1980
            Index           =   1
            Left            =   30
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   510
            Width           =   4065
            _cx             =   7170
            _cy             =   3493
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   16777215
            ForeColor       =   128
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "⁄Ê«„· «·»ÕÀ"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   4
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
            Begin VB.TextBox TxtSearch 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   4
               Left            =   1410
               TabIndex        =   32
               Top             =   390
               Width           =   1545
            End
            Begin MSDataListLib.DataCombo DcboCustomer 
               Height          =   315
               Index           =   1
               Left            =   30
               TabIndex        =   34
               Top             =   720
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "«”„ «·⁄„Ì·"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "—ﬁ„ «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   390
               Width           =   975
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "‰ «∆Ã «·»ÕÀ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2595
            Width           =   1335
         End
      End
      Begin VB.Timer TimManAlram 
         Left            =   240
         Top             =   270
      End
      Begin VB.PictureBox pboxForm 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   300
         ScaleHeight     =   2625
         ScaleWidth      =   3765
         TabIndex        =   4
         Top             =   5040
         Visible         =   0   'False
         Width           =   3765
         Begin VB.CheckBox chkOption 
            Alignment       =   1  'Right Justify
            Caption         =   "› Õ √ﬂÀ— „‰ „Ã„Ê⁄… ›Ï ‰›” «·Êﬁ "
            Height          =   255
            Index           =   1
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1050
            Width           =   2805
         End
         Begin VB.CheckBox chkOption 
            Alignment       =   1  'Right Justify
            Caption         =   "≈Œ›«¡ ‰Â«∆Ì«"
            Height          =   255
            Index           =   0
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   780
            Width           =   2805
         End
         Begin VB.ComboBox CboStyle 
            Height          =   315
            Left            =   450
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   3195
         End
         Begin ImpulseButton.ISButton BtnOK 
            Height          =   375
            Left            =   1050
            TabIndex        =   8
            Top             =   2190
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„Ê«›ﬁ"
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
            BackStyle       =   0
            ButtonImage     =   "FrmPane.frx":0000
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton btnCancelPop 
            Cancel          =   -1  'True
            Height          =   375
            Left            =   60
            TabIndex        =   9
            Top             =   2190
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈·€«¡"
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
            BackStyle       =   0
            ButtonImage     =   "FrmPane.frx":039A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘ﬂ· «·⁄«„"
            Height          =   225
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   1125
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   0
            X2              =   3735
            Y1              =   2130
            Y2              =   2145
         End
      End
      Begin VB.Timer TimerData 
         Interval        =   1000
         Left            =   3630
         Top             =   150
      End
      Begin VB.PictureBox PicChart 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Index           =   0
         Left            =   2790
         RightToLeft     =   -1  'True
         ScaleHeight     =   2055
         ScaleWidth      =   1125
         TabIndex        =   3
         Top             =   1590
         Width           =   1125
         Begin Cfx62ClientServerCtl.Chart itemchart 
            Height          =   1695
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   3135
            _Data_          =   "FrmPane.frx":0734
         End
      End
      Begin VB.PictureBox PicFg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1665
         Index           =   0
         Left            =   900
         RightToLeft     =   -1  'True
         ScaleHeight     =   1665
         ScaleWidth      =   1545
         TabIndex        =   2
         Top             =   210
         Width           =   1545
         Begin VSFlex8Ctl.VSFlexGrid Fg 
            Height          =   1485
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   150
            Width           =   1485
            _cx             =   2619
            _cy             =   2619
            Appearance      =   2
            BorderStyle     =   0
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPane.frx":0BD5
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
      Begin MSComctlLib.ImageList imlViewIcons 
         Left            =   2610
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":0CAA
               Key             =   "Payment"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":1984
               Key             =   "ManSearch"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":265E
               Key             =   "SAStock"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":3338
               Key             =   "EgStock"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":4012
               Key             =   "RePur"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":4CEC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":59C6
               Key             =   "Sales"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":66A0
               Key             =   "Pur"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":737A
               Key             =   "Man"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":8054
               Key             =   "Emps"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":8D2E
               Key             =   "Safe"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":9A08
               Key             =   "Expensive"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":A6E2
               Key             =   "Cash"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":B3BC
               Key             =   "Gold"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":C096
               Key             =   "Tip"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":CD70
               Key             =   "SA"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":CEFC
               Key             =   "News"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":DBD6
               Key             =   "CreditCart"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":E8B0
               Key             =   "ReSales"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":F58A
               Key             =   "Notes"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlTaskPanelIcons 
         Left            =   3210
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   65280
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":10264
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":104D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":10772
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":108F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":10A8D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":10C35
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":10DD7
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":10F62
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":110ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":111F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11480
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":1158C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11808
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":118CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11A6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11C0A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlToolbarIcons 
         Left            =   2610
         Top             =   930
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11CB5
               Key             =   "New"
               Object.Tag             =   "100"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11DC7
               Key             =   "Open"
               Object.Tag             =   "101"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11ED9
               Key             =   "Save"
               Object.Tag             =   "103"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":11FEB
               Key             =   "Print"
               Object.Tag             =   "113"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":120FD
               Key             =   "Cut"
               Object.Tag             =   "108"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":1220F
               Key             =   "Copy"
               Object.Tag             =   "106"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":12321
               Key             =   "Paste"
               Object.Tag             =   "107"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":12433
               Key             =   "About"
               Object.Tag             =   "112"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":12545
               Key             =   "Undo"
               Object.Tag             =   "140"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPane.frx":12897
               Key             =   ""
               Object.Tag             =   "145"
            EndProperty
         EndProperty
      End
      Begin XtremeCommandBars.CommandBars CommandBars 
         Left            =   120
         Top             =   7650
         _Version        =   786432
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         RightToLeft     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PopupControl PopUpControl 
      Index           =   0
      Left            =   240
      Top             =   1080
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
End
Attribute VB_Name = "FrmPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_PanelType As Integer
Dim TTP As clstooltip
Dim cSearchDcbo(2) As clsDCboSearch
Const IDOK = 1
Const IDCLOSE = 2
Const IDSITE = 3

Private Sub CboStyle_Change()

    If CboStyle.ListIndex = -1 Then Exit Sub
    Me.TaskPanel1.VisualTheme = val(Me.CboStyle.ItemData(Me.CboStyle.ListIndex))
    pboxForm.SetFocus
End Sub

Private Sub CboStyle_Click()
    CboStyle_Change
End Sub

Private Sub Chk_Click(Index As Integer)

    Select Case Index

        Case 0
            Me.TimManAlram.Enabled = IIf(Chk(Index).value = vbChecked, True, False)
    End Select

End Sub

Private Sub Command1_Click()
 
End Sub

Private Sub CommandBars_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)

    If (CommandBar.Title = "Form") Then
        Dim ControlForm As CommandBarControlCustom
        CommandBar.Controls.DeleteAll
        Set ControlForm = CommandBar.Controls.Add(xtpControlCustom, 0, "Form")
        ControlForm.Handle = pboxForm.hWnd
        pboxForm.BackColor = CommandBars.GetSpecialColor(XPCOLOR_MENUBAR_FACE)
        chkOption(0).BackColor = pboxForm.BackColor
        chkOption(1).BackColor = pboxForm.BackColor
        Exit Sub
    End If

End Sub

Private Sub btnCancelPop_Click()
    CommandBars.ClosePopups
End Sub

Private Sub BtnOK_Click()

    If chkOption(1).value = vbChecked Then
        Me.TaskPanel1.Behaviour = xtpTaskPanelBehaviourList
    Else
        Me.TaskPanel1.Behaviour = xtpTaskPanelBehaviourExplorer
    End If

End Sub

Private Sub Fg_DblClick(Index As Integer)
    Exit Sub
    Dim LngTransID As Long

    If FG(Index).Row <= 0 Then Exit Sub

    With FG(Index)

        Select Case Index

            Case 0
                LngTransID = val(FG(Index).TextMatrix(FG(Index).Row, FG(Index).ColIndex("TransID")))
                OpenScreen InvoiceScreen, LngTransID

            Case 1

            Case 2

            Case 3
        End Select

    End With

End Sub

Private Sub Form_Load()
    LoadCommandBar
    SystemOptions.BolStopUpdateTask = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If Me.TimerData.Enabled = True Then
        SystemOptions.BolStopUpdateTask = True
        Me.TimerData.Enabled = False
    End If

    If Me.TimManAlram.Enabled = True Then
        Me.TimManAlram.Enabled = True
    End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.PicLeftPan.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    UnloadArrayControls
    Set TTP = Nothing

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

End Sub

Private Sub Opt_Click(Index As Integer)
    Me.Ele(0).Visible = Me.Opt(0).value
    Me.Ele(1).Visible = Me.Opt(1).value
End Sub

Private Sub PopUpControl_ItemClick(Index As Integer, _
                                   ByVal Item As XtremeSuiteControls.IPopupControlItem)

    If Item.id = IDCLOSE Or Item.id = IDOK Then
        PopUpControl(Index).Close
    End If

End Sub

Private Sub PopUpControl_StateChanged(Index As Integer)

    If PopUpControl(Index).State = xtpPopupStateClosed Then
        Debug.Print Index & " xtpPopupStateClosed"
    End If

    'If Index > 0 Then
    '    If PopUpControl(Index).State = xtpPopupStateClosed Then
    '        Unload PopUpControl(Index)
    '    End If
    'End If
End Sub

Private Sub TaskPanel1_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
    Dim Reports As ClsRepoerts
    Dim startDate As Date
    Dim EndDate As Date
    Dim StrSQL As String
    Dim IntTemp As Integer

    If Not Item Is Nothing Then

        Select Case Item.Group.id

            Case TaskPanelGroupsIDs.TskGroupSales

                ' ﬁ«—Ì— Œ«’… »«·„»Ì⁄« 
                If Item.id = TaskPnlTransGroupItemsIDs.TskItemDayValTrans Then
                    ' ﬁ«—Ì— «·„»Ì⁄«  «·ÌÊ„
                    Set Reports = New ClsRepoerts
                    startDate = GetWeekStartEND(Date, 0)
                    EndDate = GetWeekStartEND(Date, 1)
                    StrSQL = "Select * From ReportSallingTime"
                    StrSQL = StrSQL + " Where Transaction_Date =" & SQLDate(Date, True)
                    startDate = Date
                    EndDate = Date
                    Reports.ShowSallingTime StrSQL, startDate, EndDate, False
                    Set Reports = Nothing
                ElseIf Item.id = TaskPnlTransGroupItemsIDs.TskItemWeekValTrans Then
                    ' ﬁ«—Ì— «·„»Ì⁄«  ›Ï «·√”»Ê⁄
                    Set Reports = New ClsRepoerts
                    startDate = GetWeekStartEND(Date, 0)
                    EndDate = GetWeekStartEND(Date, 1)
                    StrSQL = "Select * From ReportSallingTime"
                    StrSQL = StrSQL + " Where Transaction_Date >=" & SQLDate(startDate, True)
                    StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(EndDate, True)
                    Reports.ShowSallingTime StrSQL, startDate, EndDate, False
                    Set Reports = Nothing
                ElseIf Item.id = TaskPnlTransGroupItemsIDs.TskItemMonthValTrans Then
                    ' ﬁ«—Ì— «·„»Ì⁄«  ›Ï «·‘Â— «·Õ«·Ï
                    Set Reports = New ClsRepoerts
                    StrSQL = "Select * From ReportSallingTime"
                    StrSQL = StrSQL + " Where Month(Transaction_Date)=" & Month(Date)
                    StrSQL = StrSQL + " AND Year(Transaction_Date)=" & year(Date)
                    startDate = CDate("1/" & Month(Date) & "/" & year(Date))
                    IntTemp = GetMonthDaysCount(Month(Date), year(Date))
                    EndDate = DateAdd("d", IntTemp - 1, startDate)
                    Reports.ShowSallingTime StrSQL, startDate, EndDate, False
                    Set Reports = Nothing
                End If

            Case TaskPanelGroupsIDs.TskGroupPurchase

                ' ﬁ«—Ì— Œ«’… »«·„‘ —Ì« 
                If Item.id = TaskPnlTransGroupItemsIDs.TskItemWeekValTrans Then
                    ' ﬁ«—Ì— «·„‘ —Ì«  ›Ï «·√”»Ê⁄
                    Set Reports = New ClsRepoerts
                    startDate = GetWeekStartEND(Date, 0)
                    EndDate = GetWeekStartEND(Date, 1)
                    StrSQL = "Select * From ReportBuyTime_Client"
                    StrSQL = StrSQL + " Where Transaction_Date >=" & SQLDate(startDate, True)
                    StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(EndDate, True)
                    Reports.ShowBuyTime StrSQL, startDate, EndDate
                    Set Reports = Nothing
                ElseIf Item.id = TaskPnlTransGroupItemsIDs.TskItemMonthValTrans Then
                    ' ﬁ«—Ì— «·„‘ —Ì«  ›Ï «·‘Â— «·Õ«·Ï
                    Set Reports = New ClsRepoerts
                    StrSQL = "Select * From ReportBuyTime_Client"
                    StrSQL = StrSQL + " Where Month(Transaction_Date)=" & Month(Date)
                    StrSQL = StrSQL + " AND Year(Transaction_Date)=" & year(Date)
                    startDate = CDate("1/" & Month(Date) & "/" & year(Date))
                    IntTemp = GetMonthDaysCount(Month(Date), year(Date))
                    EndDate = DateAdd("d", IntTemp - 1, startDate)
                    Reports.ShowBuyTime StrSQL, startDate, EndDate
                    Set Reports = Nothing
                End If

        End Select

    End If

End Sub

Private Sub TimerData_Timer()
    Dim i As Integer
    Static J As Integer
    J = J + 1

    If J < 15 Then
        Exit Sub
    ElseIf SystemOptions.BolUpdateTaskInProgress = True Then
        '·Ê «‰ ⁄„·Ì… «· ÕœÌÀ «·”«»ﬁ… „«“«·   Ã—Ï Ê·„  ‰ ÂÏ
        Exit Sub
    End If

    If SystemOptions.BolStopUpdateTask = True Then
        '·Ê «‰ «·»—‰«„Ã ÌÃ—Ï Õ«·Ì« ≈€·«ﬁÂ
        '·«»œ „‰ ≈·€«¡ ⁄„·Ì… «· ÕœÌÀ
        TimerData.Enabled = False
        Exit Sub
    End If

    SystemOptions.BolUpdateTaskInProgress = True
    J = 0

    For i = 1 To Me.TaskPanel1.Groups.count

        If SystemOptions.BolStopUpdateTask = True Then
            '·Ê «‰ «·»—‰«„Ã ÌÃ—Ï Õ«·Ì« ≈€·«ﬁÂ
            '·«»œ „‰ ≈·€«¡ ⁄„·Ì… «· ÕœÌÀ
            TimerData.Enabled = False
            Exit Sub
        End If

        If Me.TaskPanel1.Groups(i).Tag = "Opened" Then
            WriteTaskPanlData Me.TaskPanel1.Groups(i).id
        End If

    Next i

    SystemOptions.BolUpdateTaskInProgress = False
End Sub

Private Sub WriteTaskPanlData(Optional IntGroupID As Integer = -1)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim DblTemp As Double

    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        'Exit Sub
    End If

    SystemOptions.BolUpdateTaskInProgress = True

    If IntGroupID = TaskPanelGroupsIDs.TskGroupSales Then
        '================================«·„»Ì⁄« ==================================
        GetTransactionGroup 2, Me.TaskPanel1.Groups(1), FG(0), itemchart(0)
        StrSQL = "Select TOP 5 * From ReportSallingTime"
        StrSQL = StrSQL + " Where Transaction_Date =" & SQLDate(Date, True)

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

        Do While rs.State = adStateExecuting

            If SystemOptions.BolStopUpdateTask = True Then
                rs.Cancel
                rs.Close
                SystemOptions.BolUpdateTaskInProgress = False
                Exit Sub
            End If

            DoEvents
        Loop

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
       
            rs.MoveFirst

            With Me.itemchart(0)
                .Visible = True
                .Gallery = Gallery_Bar
                .Chart3D = True
                .ShowTips = True
                .LegendBox = False
                .LegendBoxObj.Alignment = ToolAlignment_Far
                .LegendBoxObj.Font.Size = 8
                .AllowEdit = False
                .MultipleColors = True
                Set .DataSourceAdo = rs
            End With

        End If
        
        With Me.FG(0)
            .Rows = 2
            .Clear flexClearScrollable

            If rs.RecordCount > 0 Then
                .Rows = rs.RecordCount + 1
                rs.MoveFirst

                For i = 1 To .Rows - 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 2) = IIf(IsNull(rs.Fields("Transaction_ID").value), "", rs.Fields("Transaction_ID").value)
         
                    .TextMatrix(i, 3) = IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value)
            
                    .TextMatrix(i, 4) = IIf(IsNull(rs.Fields("TransNet").value), "", rs.Fields("TransNet").value)
                    '
                    rs.MoveNext
                Next

                rs.Close
            End If

            .RowHeight(-1) = 300
        End With

        '=========================================================================
    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupPurchase Then
        '================================«·„‘ —Ì« ==================================
        GetTransactionGroup 1, Me.TaskPanel1.Groups(2), FG(1), itemchart(1)
    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupReSales Then
        GetTransactionGroup 9, Me.TaskPanel1.Groups(3), FG(2), itemchart(2)
    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupRePurchase Then
        GetTransactionGroup 5, Me.TaskPanel1.Groups(4), FG(3), itemchart(3)
    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupBoxes Then
        Dim RsBoxes As ADODB.Recordset
        Dim RsBeg As ADODB.Recordset
        Dim Item As TaskPanelGroupItem
        Dim Group As TaskPanelGroup
        Dim IntGroupIndex As Integer

        For i = 1 To Me.TaskPanel1.Groups.count

            If Me.TaskPanel1.Groups(i).id = IntGroupID Then
                IntGroupIndex = i
                Exit For
            End If

        Next i

        Set Group = Me.TaskPanel1.Groups(IntGroupIndex)
        Group.Items.Clear
        StrSQL = "Select BoxID ,BoxName From TblBoxesData Order By BoxID"
        Set RsBoxes = New ADODB.Recordset
        RsBoxes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsBoxes.BOF Or RsBoxes.EOF) Then
            StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName,"
            StrSQL = StrSQL + " Sum(QryBoxBalance.Note_Value * TransDir) As BoxAccount"
            StrSQL = StrSQL + " From  dbo.QryBoxBalance()QryBoxBalance"
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate < " & SQLDate(Date, True) & ""
            StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName"
            StrSQL = StrSQL + " Order By QryBoxBalance.BoxID"
            Set RsBeg = New ADODB.Recordset
            RsBeg.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

            Do While RsBeg.State = adStateExecuting

                If SystemOptions.BolStopUpdateTask = True Then
                    RsBeg.Cancel
                    RsBeg.Close
                    SystemOptions.BolUpdateTaskInProgress = False
                    Exit Sub
                End If

                DoEvents
            Loop

            RsBoxes.MoveFirst

            For i = 1 To RsBoxes.RecordCount
                Set Item = Group.Items.Add(0, RsBoxes("BoxName").value, xtpTaskItemTypeText)
                Item.Bold = True
                DblTemp = 0

                If Not (RsBeg.BOF Or RsBeg.EOF) Then
                    RsBeg.MoveFirst
                    RsBeg.find "BoxID=" & RsBoxes("BoxID").value, , adSearchForward, 1

                    If Not (RsBeg.BOF Or RsBeg.EOF) Then
                        DblTemp = IIf(IsNull(RsBeg("BoxAccount").value), 0, RsBeg("BoxAccount").value)
                    Else
                        RsBeg.MoveFirst
                    End If
                End If

                If SystemOptions.UserInterface = ArabicInterface Then
                    Group.Items.Add 0, "«·—’Ìœ «·√›  «ÕÏ ··Œ“‰… : " & Format(DblTemp, SystemOptions.SysDefCurrencyForamt), xtpTaskItemTypeLink
                Else
                    Group.Items.Add 0, "Box Totay Openning Balance: " & Format(DblTemp, SystemOptions.SysDefCurrencyForamt), xtpTaskItemTypeLink
                End If

                RsBoxes.MoveNext
            Next i

        End If

        Group.EnsureVisible
    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupExpenses Then
        '«·„’—Ê›« 
        StrSQL = "Select Count(NoteID)CountX,Sum(Notes.Note_Value)as SumX"
        StrSQL = StrSQL + " From NOTES "
        StrSQL = StrSQL + " Where (NOTES.NoteType = 3)"
        StrSQL = StrSQL + " AND Month(NOTES.NOTEDate)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Day(NOTES.NOTEDate)=" & Day(Date) & ""
        StrSQL = StrSQL + " AND Year(NOTES.NOTeDate)=" & year(Date) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

        Do While rs.State = adStateExecuting

            If SystemOptions.BolStopUpdateTask = True Then
                rs.Cancel
                rs.Close
                SystemOptions.BolUpdateTaskInProgress = False
                Exit Sub
            End If

            DoEvents
        Loop

        DblTemp = 0

        If Not (rs.BOF Or rs.EOF) Then
            DblTemp = IIf(IsNull(rs("CountX").value), 0, rs("CountX").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.TaskPanel1.Groups(6).Items(2).Caption = "⁄œœ ⁄„·Ì«  «·’—› : " & DblTemp
            Else
                Me.TaskPanel1.Groups(6).Items(2).Caption = "Count Of Expenses Operation: " & DblTemp
            End If

            DblTemp = 0
            DblTemp = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.TaskPanel1.Groups(6).Items(1).Caption = "ﬁÌ„… „’—Ê›«  «·ÌÊ„ : " & DblTemp
            Else
                Me.TaskPanel1.Groups(6).Items(1).Caption = "Total of Today Expenses:" & DblTemp
            End If
        End If

        '-------------------
        StrSQL = " Select TOP 1 NoteID,NOTESERIAL"
        StrSQL = StrSQL + " From Notes"
        StrSQL = StrSQL + " Where NoteType = 3"
        StrSQL = StrSQL + " AND Month(NOTES.NOTEDate)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Day(NOTES.NOTEDate)=" & Day(Date) & ""
        StrSQL = StrSQL + " AND Year(NOTES.NOTeDate)=" & year(Date) & ""
        StrSQL = StrSQL + " Order By NOTEID DESC "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

        Do While rs.State = adStateExecuting

            If SystemOptions.BolStopUpdateTask = True Then
                rs.Cancel
                rs.Close
                SystemOptions.BolUpdateTaskInProgress = False
                Exit Sub
            End If

            DoEvents
        Loop

        DblTemp = 0

        If Not (rs.BOF Or rs.EOF) Then
            DblTemp = IIf(IsNull(rs("NoteID").value), 0, rs("NoteID").value)
            Me.TaskPanel1.Groups(6).Items(3).Tag = DblTemp
            Me.TaskPanel1.Groups(6).Items(3).ToolTip = "≈÷€ÿ Â‰« Õ Ï Ì „ › Õ ‘«‘… «·„’—Ê›«  ⁄·Ï «Œ— ⁄„·Ì… ’—›"
            DblTemp = 0
            DblTemp = IIf(IsNull(rs("NOTESERIAL").value), 0, rs("NOTESERIAL").value)
            Me.TaskPanel1.Groups(6).Items(3).Caption = "«Œ— ⁄„·Ì… „’—Ê›«  : " & DblTemp
        End If

        '-------------------«·—”„ «·»Ì«‰Ï
        StrSQL = "Select DATENAME(Day ,NoteDate)as DayNumber,DATENAME(WeekDay ," & "NOTEDATE)AS DayName,Sum(Note_Value)AS SumX "
        StrSQL = StrSQL + " From Notes Where NoteType = 3 "
        StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(GetWeekStartEND(Date, 0), True) & ""
        StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(GetWeekStartEND(Date, 1), True) & ""
        StrSQL = StrSQL + " Group By NoteDate"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

        Do While rs.State = adStateExecuting

            If SystemOptions.BolStopUpdateTask = True Then
                rs.Cancel
                rs.Close
                SystemOptions.BolUpdateTaskInProgress = False
                Exit Sub
            End If

            DoEvents
        Loop

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            DblTemp = 0

            For i = 0 To rs.RecordCount - 1
                DblTemp = DblTemp + IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)
                rs.MoveNext
            Next i

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.TaskPanel1.Groups(6).Items(7).Caption = "≈Ã„«·Ï „’—Ê›«  «·√”»Ê⁄ «·Õ«·Ï: " & DblTemp
            Else
                Me.TaskPanel1.Groups(6).Items(7).Caption = "Total Current Week Expenses:" & DblTemp
            End If

            rs.MoveFirst

            With Me.itemchart(4)
                .Visible = True
                .Gallery = Gallery_Bar
                .Chart3D = True
                .ShowTips = True
                .LegendBox = False
                .LegendBoxObj.Alignment = ToolAlignment_Far
                .LegendBoxObj.Font.Size = 8
                .AllowEdit = False
                .MultipleColors = True
                Set .DataSourceAdo = rs
            End With

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.itemchart(4).SetMessageText "NoData", "·« ÊÃœ »Ì«‰«  ··⁄—÷"
            End If

            With Me.itemchart(4)
                .ClearData ClearDataFlag_AllData
            End With

        End If

        StrSQL = "Select " & "Sum(Note_Value)AS SumX "
        StrSQL = StrSQL + " From Notes Where NoteType = 3 "
        StrSQL = StrSQL + " AND Month(NOTES.NOTEDate)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Year(NOTES.NOTeDate)=" & year(Date) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

        Do While rs.State = adStateExecuting

            If SystemOptions.BolStopUpdateTask = True Then
                rs.Cancel
                rs.Close
                SystemOptions.BolUpdateTaskInProgress = False
                Exit Sub
            End If

            DoEvents
        Loop

        DblTemp = 0
        DblTemp = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.TaskPanel1.Groups(6).Items(8).Caption = "≈Ã„«·Ï „’—Ê›«  «·‘Â— «·Ã«—Ï:" & DblTemp
        Else
            Me.TaskPanel1.Groups(6).Items(8).Caption = "Total Current Month Expenses:" & DblTemp
        End If

        '---------------------------------------------
    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupNotesRecivable Then
        '«·„ﬁ»Ê÷« 
        StrSQL = " SELECT SUM(Note_Value) AS SumX"
        StrSQL = StrSQL + " From Notes "
        StrSQL = StrSQL + " Where (NoteType = 4) "
        StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            DblTemp = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.TaskPanel1.Groups(7).Items(1).Caption = "Total Receipts Today : " & DblTemp

            Else
                Me.TaskPanel1.Groups(7).Items(1).Caption = "≈Ã„«·Ï „ﬁ»Ê÷«  «·ÌÊ„ : " & DblTemp

            End If

            Me.TaskPanel1.Groups(7).Items(1).ToolTip = WriteNo(CStr(DblTemp), 0)
        End If

        StrSQL = "SELECT NoteID, NoteSerial, Note_Value From NOTES "
        StrSQL = StrSQL + " WHERE (NoteID =(SELECT MAX(NOTEID)"
        StrSQL = StrSQL + " From NOTES WHERE NoteType = 4)) "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Me.TaskPanel1.Groups(7).Items(2).Caption = "Last Receipt opr : " & rs("NoteSerial").value
                Me.TaskPanel1.Groups(7).Items(2).ToolTip = "Last Receipt Value : " & WriteNo(CStr(rs("Note_Value").value), 0)
       
            Else
                Me.TaskPanel1.Groups(7).Items(2).Caption = "√Œ— ⁄„·Ì… ﬁ»÷ : " & rs("NoteSerial").value
                Me.TaskPanel1.Groups(7).Items(2).ToolTip = "ﬁÌ„… «Œ— ⁄„·Ì… ﬁ»÷ : " & WriteNo(CStr(rs("Note_Value").value), 0)
            End If
        
        End If

    ElseIf IntGroupID = TaskPanelGroupsIDs.TskGroupNotesPayable Then
        '«·„œ›Ê⁄« 
        StrSQL = " SELECT SUM(Note_Value) AS SumX"
        StrSQL = StrSQL + " From Notes "
        StrSQL = StrSQL + " Where (NoteType = 5) "
        StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            DblTemp = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.TaskPanel1.Groups(8).Items(1).Caption = "Total payments today : " & DblTemp
            Else
                Me.TaskPanel1.Groups(8).Items(1).Caption = "≈Ã„«·Ï „œ›Ê⁄«  «·ÌÊ„ : " & DblTemp
            End If
        
            Me.TaskPanel1.Groups(8).Items(1).ToolTip = WriteNo(CStr(DblTemp), 0)
        End If

        StrSQL = "SELECT NoteID, NoteSerial, Note_Value From NOTES "
        StrSQL = StrSQL + " WHERE (NoteID =(SELECT MAX(NOTEID)"
        StrSQL = StrSQL + " From NOTES WHERE NoteType = 5)) "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Me.TaskPanel1.Groups(8).Items(2).Caption = "last Payment  Opr : " & rs("NoteSerial").value
                Me.TaskPanel1.Groups(8).Items(2).ToolTip = "last Payment  Value : " & WriteNo(CStr(rs("Note_Value").value), 0)
        
            Else
                Me.TaskPanel1.Groups(8).Items(2).Caption = "√Œ— ⁄„·Ì… œ›⁄ : " & rs("NoteSerial").value
                Me.TaskPanel1.Groups(8).Items(2).ToolTip = "ﬁÌ„… «Œ— ⁄„·Ì… œ›⁄ : " & WriteNo(CStr(rs("Note_Value").value), 0)
        
            End If
        End If
    End If

    SystemOptions.BolUpdateTaskInProgress = False
    Exit Sub
ErrTrap:
    SystemOptions.BolUpdateTaskInProgress = False
End Sub

Public Sub CreateTaskPanel()
    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim J As Integer
    Dim GrdBack  As New ClsBackGroundPic
    Dim LastFgIndex As Integer
    Dim LastPicIndex As Integer
    Dim StrToolTip As String

    On Local Error GoTo ErrTrap
    ''-----------------------------------------------------
    UnloadArrayControls
    ClearTaskPanel

    If SystemOptions.UserInterface = ArabicInterface Then
        TaskPanel1.RightToLeft = True
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        TaskPanel1.RightToLeft = False
    End If

    TaskPanel1.SetImageList Me.imlViewIcons
    TaskPanel1.VisualTheme = xtpTaskPanelThemeOffice2000
    TaskPanel1.ItemLayout = xtpTaskItemLayoutDefault
    '----------------------------------------------------

    If PanelType = 1 Then

        For i = 1 To 3
            Load PicFg(PicFg.count)
            Load FG(FG.count)
            Load PicChart(PicChart.count)
            Load itemchart(itemchart.count)
        Next i

        '---------------------------------
        '√œÊ«  «·—”„ «·»Ì«‰Ï «·Œ«’ »«·„’—Ê›« 
        Load PicChart(PicChart.count)
        Load itemchart(itemchart.count)
        '---------------------------------
        '---Sales Group
        Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupSales, "")
        CreateTransGroup 2, Group, FG(0), PicFg(0), itemchart(0), PicChart(0)
        '---Purchase Group
        Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupPurchase, "")
        CreateTransGroup 1, Group, FG(1), PicFg(1), itemchart(1), PicChart(1)
        '---Re Sales Group
        Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupReSales, "")
        CreateTransGroup 9, Group, FG(2), PicFg(2), itemchart(2), PicChart(2)
        '---Re Purchase Group
        Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupRePurchase, "")
        CreateTransGroup 5, Group, FG(3), PicFg(3), itemchart(3), PicChart(3)

        '--------------------------------------------------------------------------
        If SystemOptions.UserInterface = ArabicInterface Then
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupBoxes, "«·Œ“Ì‰…")
            Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ „·Œ’ Õ—ﬂ«  ﬂ· «·Œ“‰ «·ÌÊ„"
        Else
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupBoxes, "Boxes")
            Group.ToolTip = "Show Information about Totay Box Balance"
        End If

        Group.IconIndex = Me.imlViewIcons.ListImages("Safe").Index
        Set rs = New ADODB.Recordset
        StrSQL = "SELECT BoxID, BoxName From TblBoxesData Order By BoxID"
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then

            For i = 1 To rs.RecordCount
                Set Item = Group.Items.Add(0, rs("BoxName").value, xtpTaskItemTypeText)
                Item.Bold = True
                Group.Items.Add 0, "«·—’Ìœ «·√›  «ÕÏ ··Œ“‰…:", xtpTaskItemTypeLink
                Group.Items.Add 0, "ﬁÌ„… œŒÊ· «·‰ﬁœÌ…:", xtpTaskItemTypeLink
                Group.Items.Add 0, "ﬁÌ„… Œ—ÊÃ «·‰ﬁœÌ…:", xtpTaskItemTypeLink
                Group.Items.Add 0, "—’Ìœ «·Œ“‰… «·√‰:", xtpTaskItemTypeLink
                rs.MoveNext
            Next i

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                Group.Items.Add 0, "·« ÊÃœ Œ“‰ „⁄—›… ›Ï «·»—‰«„Ã", xtpTaskItemTypeText
            Else
                Group.Items.Add 0, "There is NO Boxes", xtpTaskItemTypeText
            End If
        End If

        '--------------------------------------------------------------------------
        If SystemOptions.UserInterface = ArabicInterface Then
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupExpenses, "„’—Ê›« ")
        Else
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupExpenses, "Expenses")
        End If

        PicChart(4).Visible = True
        itemchart(4).Visible = True
        Set itemchart(4).Container = PicChart(4)
        Group.IconIndex = Me.imlViewIcons.ListImages("Expensive").Index

        If SystemOptions.UserInterface = ArabicInterface Then
            Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ ÕÃ„ «·„’—Ê›«  «·ÌÊ„"
            Group.Items.Add 0, "ﬁÌ„… „’—Ê›«  «·ÌÊ„:", xtpTaskItemTypeLink
            Group.Items.Add 0, "⁄œœ ⁄„·Ì«  «·’—›:", xtpTaskItemTypeLink
            Group.Items.Add 0, "√Œ— ⁄„·Ì… „’—Ê›« :", xtpTaskItemTypeLink
            Set Item = Group.Items.Add(0, "—”„ »Ì«‰Ï ÌÊ÷Õ ÕÃ„ «·„’—Ê›«  Â–« «·√”»Ê⁄", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = vbRed
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            Set Item.Control = PicChart(4)
            Set Item = Group.Items.Add(0, "„⁄·Ê„«  «Œ—Ï „Â„…...", xtpTaskItemTypeText)
            Item.Bold = True
            Group.Items.Add 0, "≈Ã„«·Ï „’—Ê›«  «·√”»Ê⁄ «·Õ«·Ï : ", xtpTaskItemTypeLink
            Group.Items.Add 0, "≈Ã„«·Ï „’—Ê›«  «·‘Â— «·Ã«—Ï : ", xtpTaskItemTypeLink

        Else
            Group.ToolTip = "Show Information about Today Expenses"
            Group.Items.Add 0, "Today Total Expenses:", xtpTaskItemTypeLink
            Group.Items.Add 0, "Today Count Expenses:", xtpTaskItemTypeLink
            Group.Items.Add 0, "Last Operation:", xtpTaskItemTypeLink
            Set Item = Group.Items.Add(0, "Chart show the Volume of Expenses in Current Week", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = vbRed
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            Set Item.Control = PicChart(4)
            Set Item = Group.Items.Add(0, "More Important Information", xtpTaskItemTypeText)
            Item.Bold = True
            Group.Items.Add 0, "Current Week Total Expenses:", xtpTaskItemTypeLink
            Group.Items.Add 0, "Current Month Total Expenses: ", xtpTaskItemTypeLink
        End If

        '--------------------------------------------------------------------------
        If SystemOptions.UserInterface = EnglishInterface Then
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupNotesRecivable, "Receipts")
            Group.ToolTip = "Total Receipts Today"
            Group.IconIndex = Me.imlViewIcons.ListImages("Cash").Index
            Group.Items.Add 0, "Total Receipts Today:", xtpTaskItemTypeLink
            Group.Items.Add 0, "Last Receipt :", xtpTaskItemTypeLink
        
        Else
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupNotesRecivable, "„ﬁ»Ê÷« ")
            Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ ÕÃ„ «·„ﬁ»Ê÷«  «·ÌÊ„"
            Group.IconIndex = Me.imlViewIcons.ListImages("Cash").Index
            Group.Items.Add 0, "ﬁÌ„… „ﬁ»Ê÷«  «·ÌÊ„ :", xtpTaskItemTypeLink
            Group.Items.Add 0, "√Œ— ⁄„·Ì… ﬁ»÷ :", xtpTaskItemTypeLink
    
        End If

        If SystemOptions.UserInterface = EnglishInterface Then
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupNotesPayable, "Payments")
        Else
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupNotesPayable, "„œ›Ê⁄« ")
        End If

        Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ ÕÃ„ «·„œ›Ê⁄«  «·ÌÊ„"
        Group.IconIndex = Me.imlViewIcons.ListImages("Payment").Index
        Group.Items.Add 0, "ﬁÌ„… „œ›Ê⁄«  «·ÌÊ„:", xtpTaskItemTypeLink
        Group.Items.Add 0, "√Œ— ⁄„·Ì… „œ›Ê⁄« :", xtpTaskItemTypeLink

        If SystemOptions.UserInterface = EnglishInterface Then

            Set Group = TaskPanel1.Groups.Add(0, " installments")
        Else
            Set Group = TaskPanel1.Groups.Add(0, "√ﬁ”«ÿ")

        End If

        Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ ÕÃ„ „— Ã⁄ «·„»Ì⁄«  «·ÌÊ„"
        Group.IconIndex = Me.imlViewIcons.ListImages("Notes").Index

        If SystemOptions.UserInterface = EnglishInterface Then
       
            Set Group = TaskPanel1.Groups.Add(0, "Employees")
        Else
            Set Group = TaskPanel1.Groups.Add(0, "„ÊŸ›Ì‰")
        End If

        Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ «·„ÊŸ›Ì‰ «·ÌÊ„"
        Group.IconIndex = Me.imlViewIcons.ListImages("Emps").Index

        If SystemOptions.UserInterface = EnglishInterface Then
       
            Group.Items.Add 0, "The attendance of staff", xtpTaskItemTypeText
            Group.Items(1).Bold = True
            Group.Items.Add 0, "Absence of staff", xtpTaskItemTypeText
            Group.Items(2).Bold = True
        Else
            Group.Items.Add 0, "Õ÷Ê— «·„ÊŸ›Ì‰", xtpTaskItemTypeText
            Group.Items(1).Bold = True
            Group.Items.Add 0, "€Ì«» «·„ÊŸ›Ì‰", xtpTaskItemTypeText
            Group.Items(2).Bold = True
        
        End If
        
    ElseIf PanelType = 2 Then
        CreateInternetNewsBar
    ElseIf PanelType = 3 Then

        If SystemOptions.SysMantainceAllow = True Then
            EleManSearch.Visible = True

            For i = 1 To 3
                Load PicFg(PicFg.count)
                Load FG(FG.count)
                Load PicChart(PicChart.count)
                Load itemchart(itemchart.count)
            Next i

            Set Group = TaskPanel1.Groups.Add(0, " ”·Ì„ ’Ì«‰… ··⁄„·«¡")
            Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ «·√’‰«› √Ê «·ﬁÿ⁄ «· Ï ÌÃ»  ”·Ì„Â« ··⁄„·«¡"
            Group.IconIndex = Me.imlViewIcons.ListImages("Man").Index
            Set Item = Group.Items.Add(0, "’Ì«‰… Õ«‰ „Ì⁄«œ  ”Ì·„Â« ··⁄„·«¡", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = SysMaronColor
            Item.SetMargins 0, 0, 0, 0
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            LastFgIndex = 0
            FG(LastFgIndex).Visible = True
            ModTaskPanel.SetupGrid FG(LastFgIndex), 1
            PicFg(LastFgIndex).Visible = True
            PicFg(LastFgIndex).Height = (FG(LastFgIndex).RowHeight(0) * 6) + 100
            Set FG(LastFgIndex).Container = PicFg(LastFgIndex)
            Set Item.Control = PicFg(LastFgIndex)
            Set Group = TaskPanel1.Groups.Add(0, "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œÌ‰")
            Group.ToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ «·√’‰«› √Ê «·ﬁÿ⁄ «· Ï ÌÃ» ≈— Ã«⁄Â« „‰ «·„Ê—œÌ‰"
            Group.IconIndex = Me.imlViewIcons.ListImages("Man").Index

            Set Item = Group.Items.Add(0, "÷„«‰ ÊÃ» —ÃÊ⁄Â „‰ «·„Ê—œÌ‰", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = SysMaronColor
            Item.SetMargins 0, 0, 0, 0
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            LastFgIndex = 1
            FG(LastFgIndex).Visible = True
            ModTaskPanel.SetupGrid FG(LastFgIndex), 1
            PicFg(LastFgIndex).Visible = True
            PicFg(LastFgIndex).Height = (FG(LastFgIndex).RowHeight(0) * 6) + 100
            Set FG(LastFgIndex).Container = PicFg(LastFgIndex)
            Set Item.Control = PicFg(LastFgIndex)
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupManCompsAsmply, " Ã„Ì⁄ «·√ÃÂ“…")
            Group.ToolTip = "›Ê« Ì— »√ÃÂ“… ﬂ«„·… ÌÃ»  Ã„Ì⁄Â«"
            Group.IconIndex = Me.imlViewIcons.ListImages("Man").Index
            Set Item = Group.Items.Add(0, "√ÃÂ“… „ÿ·Ê»  Ã„Ì⁄Â«", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = vbRed
            Item.SetMargins 0, 0, 0, 0
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            LastFgIndex = 2
            FG(LastFgIndex).Visible = True
            ModTaskPanel.SetupGrid FG(LastFgIndex), 1
            PicFg(LastFgIndex).Visible = True
            PicFg(LastFgIndex).Height = (FG(LastFgIndex).RowHeight(0) * 6) + 100
            Set FG(LastFgIndex).Container = PicFg(LastFgIndex)
            Set Item.Control = PicFg(LastFgIndex)
            '-------------------------
            Set Item = Group.Items.Add(0, "√ÃÂ“…  „  Ã„Ì⁄Â«", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = vbRed
            Item.SetMargins 0, 0, 0, 0
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            LastFgIndex = 3
            FG(LastFgIndex).Visible = True
            ModTaskPanel.SetupGrid FG(LastFgIndex), 1
            PicFg(LastFgIndex).Visible = True
            PicFg(LastFgIndex).Height = (FG(LastFgIndex).RowHeight(0) * 6) + 100
            Set FG(LastFgIndex).Container = PicFg(LastFgIndex)
            Set Item.Control = PicFg(LastFgIndex)
            Set Group = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupManSearch, "»ÕÀ ›Ï «·’Ì«‰…")
            Group.IconIndex = Me.imlViewIcons.ListImages("ManSearch").Index
            Set Item = Group.Items.Add(0, "«Œ — ‰Ê⁄ «·»ÕÀ «·–Ï  —ÌœÂ", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = vbRed
            Item.SetMargins 0, 0, 0, 0
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            Set Item.Control = Me.EleManSearch
            Me.EleManSearch.BackColor = Item.BackColor
            '-----------------------------------------
            Set Me.CmdSeacrh.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Find").Picture
            Me.CmdSeacrh.ButtonPositionImage = impRightOfText
            '-----------------------------------------
            Set Group = TaskPanel1.Groups.Add(0, "ŒÌ«—«  «·’Ì«‰…")
            Group.ToolTip = "⁄Ê«„· ÊŒÌ«—«  Œ«’… » ‰»ÌÂ«  «·’Ì«‰…"
            Group.IconIndex = Me.imlViewIcons.ListImages("Man").Index
            Set Item = Group.Items.Add(0, " ‰»ÌÂ«   Ã„Ì⁄ «·√ÃÂ“…", xtpTaskItemTypeText)
            Item.Bold = True
            Item.TextColor = vbRed
            Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
            Me.Chk(0).Caption = " ‘€Ì·  ‰»ÌÂ«   Ã„Ì⁄ «·√ÃÂ“…"
            Me.Chk(0).Visible = True
            Me.Chk(0).RightToLeft = True
            Me.Chk(0).BackColor = vbWhite
            Set Item.Control = Me.Chk(0)
            '---------------------------------------
            AddManTip
            SetManDefaluts
            '    '---------------------------------------
        End If
    End If

    'TaskPanel1.SetMargins 5, 5, 5, 5, 5
    For i = Me.TaskPanel1.Groups.count To 1 Step -1
        Me.TaskPanel1.Groups(i).Expanded = False
        Me.TaskPanel1.Groups(i).AllowDrag = xtpTaskItemAllowDrag
    Next i

    TaskPanel1.SetIconSize 16, 16
    TaskPanel1.ToolTipContext.Style = xtpToolTipOffice2007
    TaskPanel1.ToolTipContext.ShowShadow = True

    If TaskPanel1.ToolTipContext.IsBalloonStyleSupported = True Then
        TaskPanel1.ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconInfo
    End If

    TaskPanel1.SetGroupInnerMargins 0, 0, 0, 0
    Exit Sub
ErrTrap:
End Sub

Private Sub PicChart_Resize(Index As Integer)
    On Error Resume Next
    Me.itemchart(Index).Move PicChart(Index).ScaleLeft, PicChart(Index).ScaleTop, PicChart(Index).ScaleWidth, PicChart(Index).ScaleHeight

End Sub

Private Sub PicLeftPan_Resize()
    On Error Resume Next
    Me.TaskPanel1.Move Me.PicLeftPan.ScaleLeft, Me.PicLeftPan.ScaleTop, Me.PicLeftPan.ScaleWidth, (Me.PicLeftPan.ScaleHeight - 500)
End Sub

Private Sub LoadCommandBar()
    Dim i As Integer
    Dim ExtendedBar As CommandBar
    Dim ControlForm As CommandBarPopup

    CommandBarsGlobalSettings.App = App
    CommandBars.DeleteAll
    Exit Sub
    Set ExtendedBar = CommandBars.Add("Extended", xtpBarTop)
    ExtendedBar.EnableDocking xtpFlagAlignBottom
    ExtendedBar.ShowExpandButton = False

    With ExtendedBar.Controls
        Set ControlForm = .Add(xtpControlSplitButtonPopup, 12, "ŒÌ«—« ...")
        ControlForm.CommandBar.SetPopupToolBar True
        ControlForm.CommandBar.Title = "Form"
    End With

    CommandBars.AddImageList imlToolbarIcons

    With Me.CboStyle
        .AddItem "√Ê›Ì” 2000 „Ã”„"
        .ItemData(.NewIndex) = xtpTaskPanelThemeOffice2000
        .AddItem "√Ê›Ì” 2000 „”ÿÕ"
        .ItemData(.NewIndex) = xtpTaskPanelThemeOffice2000Plain
        .AddItem "√Ê›Ì” «ﬂ” »Ï „”ÿÕ"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeOfficeXPPlain
        .AddItem "√Ê›Ì” 2003 „Ã”„"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeOffice2003
        .AddItem "√Ê›Ì” 2003 „”ÿÕ"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeOffice2003Plain
        .AddItem "ÊÌ‰œÊ“ √ﬂ” »Ï „Ã”„"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeNativeWinXP
        .AddItem "ÊÌ‰œÊ“ √ﬂ” »Ï „”ÿÕ"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeNativeWinXPPlain
        .AddItem "⁄‰«’— ›Ï ‘ﬂ· ﬁ«∆„…"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeListView
        .AddItem "⁄‰«’— ›Ï ‘ﬂ· ﬁ«∆„…(»‰Ÿ«„ «Ê›Ì” √ﬂ” »Ï)"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeListViewOfficeXP
        .AddItem "⁄‰«’— ›Ï ‘ﬂ· ﬁ«∆„…(»‰Ÿ«„ «Ê›Ì” 2003 )"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeListViewOffice2003
        .AddItem "‘—Ìÿ ≈Œ ’«—«  »‰Ÿ«„ √Ê›Ì” 2003"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeShortcutBarOffice2003
        .AddItem "ﬁ«∆„… √“«— „Ã”„…"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeToolbox
        .AddItem "ﬁ«∆„… √“«— „”ÿÕ…"
        .ItemData(.NewIndex) = XTPTaskPanelVisualTheme.xtpTaskPanelThemeToolboxWhidbey
    
    End With

End Sub

Private Sub TaskPanel1_GroupExpanded(ByVal Group As XtremeTaskPanel.ITaskPanelGroup)

    If Group.Expanded = False Then
        'The User Close the group
        Group.Tag = "Closed"
    ElseIf Group.Expanded = True Then
        'The User Open the group
        Group.Tag = "Opened"

        'Make Code to Load this group Data
        If Me.PanelType = 1 Then
            WriteTaskPanlData Group.id
        ElseIf Me.PanelType = 3 Then
            SetDcboSearch
        End If

        Group.EnsureVisible

        If Group.Items.count > 0 Then
            'Make Sure the Last Item In the Group
            'to Ensure that all the Group are visible
            Group.Items(Group.Items.count).EnsureVisible
        End If
    End If

End Sub

Private Sub TaskPanel1_GroupExpanding(ByVal Group As XtremeTaskPanel.ITaskPanelGroup, _
                                      ByVal Expanding As Boolean, _
                                      Cancel As Boolean)

    If Expanding = True And Group.Tag = "Opened" Then
        Cancel = True
    Else

        If Group.id = TaskPanelGroupsIDs.TskGroupBoxes Then
            Group.Items.Clear
        End If
    End If

End Sub

Public Sub CreateTransGroup(IntTransType As Integer, _
                            xGroup As TaskPanelGroup, _
                            Optional FG As VSFlex8Ctl.vsFlexGrid = Nothing, _
                            Optional PicFg As PictureBox = Nothing, _
                            Optional xChart As Object = Nothing, _
                            Optional PicChart As PictureBox = Nothing)

    Dim Item As TaskPanelGroupItem
    Dim StrGroupTitle As String
    Dim StrTemp As String
    Dim IntGroupIconIndex As Integer
    Dim StrItemToolTip As String
    Dim StrGroupToolTip As String

    If SystemOptions.UserInterface = EnglishInterface Then
        If SystemOptions.SysRegisterState = DevelopVersion Then
            'Stop
        End If
    End If

    If IntTransType = 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "«·„»Ì⁄« "
        Else
            StrGroupTitle = "Sales"
        End If

        IntGroupIconIndex = Me.imlViewIcons.ListImages("Sales").Index
    ElseIf IntTransType = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "«·„‘ —Ì« "
        Else
            StrGroupTitle = "Purchases"
        End If

        IntGroupIconIndex = Me.imlViewIcons.ListImages("Pur").Index
    ElseIf IntTransType = 9 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "„— Ã⁄ «·„»Ì⁄« "
        Else
            StrGroupTitle = "Retrun Sales"
        End If

        IntGroupIconIndex = Me.imlViewIcons.ListImages("ReSales").Index
    ElseIf IntTransType = 5 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "„— Ã⁄ «·„‘ —Ì« "
        Else
            StrGroupTitle = "Retrun Purchases"
        End If

        IntGroupIconIndex = Me.imlViewIcons.ListImages("RePur").Index
    End If

    'Set xGroup = TaskPanel1.Groups.Add(TaskPanelGroupsIDs.TskGroupSales, StrGroupTitle)
    xGroup.Caption = StrGroupTitle
    xGroup.IconIndex = IntGroupIconIndex

    If SystemOptions.UserInterface = ArabicInterface Then
        StrGroupToolTip = "⁄—÷ „⁄·Ê„«  ⁄‰ " & StrGroupTitle & " Œ·«· «·ÌÊ„ «·Õ«·Ï"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "„À· ..."
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "⁄œœ «·›Ê« Ì— Ê«·≈Ã„«·Ï ..., Ê≈Ã„«·Ï «·›Ê« Ì— «·‰ﬁœÌ… Ê«·√Ã·…"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "ﬂ„« Ì⁄—÷ «„«„ﬂ √Œ— 5 ›Ê« Ì— ”Ã·  Œ·«· «·ÌÊ„ ·≈” ⁄—«÷Â«"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Ê√Ì÷« —”„ »Ì«‰Ï ·≈⁄ÿ«∆ﬂ „ﬁ«—‰… »Ì‰ ÕÃ„ " & StrGroupTitle & " Œ·«· «·√”»Ê⁄"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Ê√Ì÷« ≈Ã„«·Ï  " & StrGroupTitle & " Œ·«· «·√”»Ê⁄ «·Õ«·Ï "
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Ê√Ì÷« ≈Ã„«·Ï  " & StrGroupTitle & " Œ·«· «·‘Â— «·Õ«·Ï "
        xGroup.ToolTip = StrGroupToolTip
        Set Item = xGroup.Items.Add(TaskPnlTransGroupItemsIDs.TskItemDayValTrans, "⁄œœ ›Ê« Ì— «·ÌÊ„:", xtpTaskItemTypeLink)
        StrItemToolTip = "⁄—÷ ›Ê« Ì— " & StrGroupTitle & " «· Ï ”Ã·  «·ÌÊ„."
        Item.ToolTip = StrItemToolTip
        StrTemp = " ≈Ã„«·Ï ﬁÌ„…" & StrGroupTitle & " : "
        xGroup.Items.Add TaskPnlTransGroupItemsIDs.TskItemDayValTrans, "≈Ã„«·Ï ﬁÌ„… " & StrGroupTitle & " : ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "⁄œœ «·√’‰«› «· Ï œŒ·  ›Ï ⁄„·Ì«  " & StrGroupTitle & " : ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, " ﬁÌ„… " & StrGroupTitle & " «·‰ﬁœÌ… : ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "ﬁÌ„… " & StrGroupTitle & " «·√Ã·… : ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "ﬁÌ„… «·Ã“¡ «·‰ﬁœÏ „‰ " & StrGroupTitle & " «·√Ã·… : ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "√Œ— ›« Ê—… »Ì⁄:", xtpTaskItemTypeLink
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrGroupToolTip = "Show information about " & StrGroupTitle & " in current day"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "like..."
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Count of Bills,Totals,Total of Cash and Total of Due"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "And Show the last 5 Bills in current day to Browse it"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Also Dispaly a chart to show " & StrGroupTitle & " in this week"
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Totals of " & StrGroupTitle & " in the cuurent Week "
        StrGroupToolTip = StrGroupToolTip & Chr(13) & "Totals of " & StrGroupTitle & " in the cuurent month. "
        xGroup.ToolTip = StrGroupToolTip
    
        Set Item = xGroup.Items.Add(TaskPnlTransGroupItemsIDs.TskItemDayValTrans, "Today Invoices Count:", xtpTaskItemTypeLink)
        StrItemToolTip = "Show " & StrGroupTitle & " Invoices Which Recorded Today"
        Item.ToolTip = StrItemToolTip
        StrTemp = "" & StrGroupTitle & " Totals: "
        xGroup.Items.Add TaskPnlTransGroupItemsIDs.TskItemDayValTrans, StrGroupTitle & " Totals: ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "Count of Items in " & StrGroupTitle & " : ", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "Cash " & StrGroupTitle & ":", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "Due " & StrGroupTitle & ":", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "Cash Value From Due " & StrGroupTitle & ":", xtpTaskItemTypeLink
        xGroup.Items.Add 0, "Last" & StrGroupTitle & " Invoice:", xtpTaskItemTypeLink
    End If

    If Not FG Is Nothing Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Set Item = xGroup.Items.Add(0, "√Œ— 5 ›Ê« Ì— ”Ã·  «·ÌÊ„", xtpTaskItemTypeText)
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Set Item = xGroup.Items.Add(0, "Last 5 Invoices Recorded Today", xtpTaskItemTypeText)
        End If

        Item.Bold = True
        Item.TextColor = vbRed
        Set Item = xGroup.Items.Add(0, "", xtpTaskItemTypeControl)
        FG.Visible = True
        ModTaskPanel.SetupGrid FG
        PicFg.Visible = True
        PicFg.Height = (FG.RowHeight(0) * 6) + 100
        Set FG.Container = PicFg
        Set Item.Control = PicFg
    End If

    If Not PicChart Is Nothing Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTemp = "—”„ »Ì«‰Ï ÌÊ÷Õ ÕÃ„ " & StrGroupTitle & " Â–« «·√”»Ê⁄"
        Else
            StrTemp = "Chart To Show The " & StrGroupTitle & " In this Week"
        End If

        Set Item = xGroup.Items.Add(0, StrTemp, xtpTaskItemTypeText)
        Item.Bold = True
        Item.TextColor = vbRed
        Set Item = xGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    
        PicChart.Visible = True
        xChart.Visible = True

        If SystemOptions.UserInterface = ArabicInterface Then
            xChart.SetMessageText "NoData", "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        End If

        Set xChart.Container = PicChart
        Item.SetMargins 0, 0, 0, 0
        Set Item.Control = PicChart
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Set Item = xGroup.Items.Add(0, "„⁄·Ê„«  «Œ—Ï „Â„…", xtpTaskItemTypeText)
        Item.Bold = True
        Item.TextColor = vbBlack
        StrTemp = "ÕÃ„ " & StrGroupTitle & " ›Ï «·√”»Ê⁄ «·Õ«·Ï:"
        StrItemToolTip = "≈÷€ÿ Â‰« Õ Ï Ì „ ⁄—÷  ﬁ—Ì— »‹ " & StrGroupTitle & " ›Ï «·√”»Ê⁄ «·Õ«·Ï "
        Set Item = xGroup.Items.Add(TaskPnlTransGroupItemsIDs.TskItemWeekValTrans, StrTemp, xtpTaskItemTypeLink)
        Item.ToolTip = StrItemToolTip
    
        StrTemp = "ÕÃ„ " & StrGroupTitle & " ›Ï «·‘Â— «·Ã«—Ï:"
        StrItemToolTip = "≈÷€ÿ Â‰« Õ Ï Ì „ ⁄—÷  ﬁ—Ì— »‹ " & StrGroupTitle & " ›Ï «·‘Â— «·Ã«—Ï "
        Set Item = xGroup.Items.Add(TaskPnlTransGroupItemsIDs.TskItemMonthValTrans, StrTemp, xtpTaskItemTypeLink)
        Item.ToolTip = StrItemToolTip
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        Set Item = xGroup.Items.Add(0, "More Important Information:", xtpTaskItemTypeText)
        Item.Bold = True
        Item.TextColor = vbBlack
        StrTemp = "Volume Of " & StrGroupTitle & " In this Week:"
        StrItemToolTip = "Click Here to Show a Report about " & StrGroupTitle & " in the Current Week"
        Set Item = xGroup.Items.Add(TaskPnlTransGroupItemsIDs.TskItemWeekValTrans, StrTemp, xtpTaskItemTypeLink)
        Item.ToolTip = StrItemToolTip
    
        StrTemp = "Volume Of " & StrGroupTitle & "in the Current Month:"
        StrItemToolTip = "Click Here to Show a Report about " & StrGroupTitle & " in the Current Month"
        Set Item = xGroup.Items.Add(TaskPnlTransGroupItemsIDs.TskItemMonthValTrans, StrTemp, xtpTaskItemTypeLink)
        Item.ToolTip = StrItemToolTip
    End If

End Sub

Private Sub PicFg_Resize(Index As Integer)
    On Error Resume Next
    Me.FG(Index).Move PicFg(Index).ScaleLeft, PicFg(Index).ScaleTop, PicFg(Index).ScaleWidth, PicFg(Index).ScaleHeight
End Sub

Public Property Get PanelType() As Integer
    PanelType = m_PanelType
End Property

Public Property Let PanelType(ByVal vNewValue As Integer)
    m_PanelType = vNewValue

    If m_PanelType = 3 Then
        Me.TimManAlram.interval = 5000
    End If

    CreateTaskPanel
End Property

Private Sub TimManAlram_Timer()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Popup As XtremeSuiteControls.PopUpControl
    Dim Item As XtremeSuiteControls.PopupControlItem
    Dim lastindex As Integer
    Dim i As Integer
    Dim Msg  As String
    Dim StrIconPath As String
    Dim J  As Integer
    Dim SngLeft As Single
    Dim SngRight As Single

    Static BolProgress As Boolean

    On Error GoTo ErrTrap

    If SystemOptions.SysMantainceAllow = False Then
        TimManAlram.Enabled = False
        Exit Sub
    End If

    StrSQL = "SELECT dbo.TblManAlram.TableID, dbo.TblManAlram.TransID," & "dbo.Transactions.Transaction_Serial, dbo.TblManAlram.AlramDate,                        dbo.TblManAl" & "ram.AlramPriority, dbo.TblManAlram.AlramMsg, dbo.TblUsers.UserName, dbo.TblManAl" & "ram.State FROM         dbo.Transactions INNER JOIN                       dbo.Tbl" & "ManAlram ON dbo.Transactions.Transaction_ID = dbo.TblManAlram.TransID INNER JOIN" & "                       dbo.TblUsers ON dbo.TblManAlram.UserID = dbo.TblUsers.Use" & "rID"
    StrSQL = StrSQL + " Where dbo.TblManAlram.State=0"
    StrSQL = StrSQL + " Order By dbo.TblManAlram.TableID ASC"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText + adAsyncExecute + adAsyncFetch

    Do While rs.State = adStateExecuting
        DoEvents
    Loop

    If rs.BOF Or rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    'Do While Me.PopUpControl.count > 1
    '    Unload PopUpControl(PopUpControl.UBound)
    'Loop

    If rs.RecordCount <= 5 Then

        For i = 0 To rs.RecordCount - 1

            If i = 0 Then
                Set Popup = Me.PopUpControl(0)
            Else
                lastindex = PopUpControl.count
                Load PopUpControl(lastindex)
                Set Popup = PopUpControl(lastindex)
            End If

            Popup.RightToLeft = True
            Popup.Animation = xtpPopupAnimationUnfold
            Popup.AnimateDelay = 256
            Popup.ShowDelay = 4000
            Popup.Transparency = 255
            '--------------------------
            Popup.RemoveAllItems
            Popup.Icons.RemoveAll
            Popup.VisualTheme = xtpPopupThemeOffice2003
            Popup.setSize 270, 100
            Msg = " Ã„Ì⁄ ÃÂ«“" & "-" & rs("UserName").value
            Set Item = Popup.AddItem(50, 27, 200, 45, Msg)
            Item.Bold = True
            Set Item = Popup.AddItem(50, 45, 270, 65, "—ﬁ„ «·›« Ê—… " & rs("Transaction_Serial").value)
            Msg = IIf(IsNull(rs("AlramMsg").value), "", rs("AlramMsg").value)
            Set Item = Popup.AddItem(50, 65, 170, 95, Msg)
            Item.TextColor = RGB(0, 61, 178)
            Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
            Set Item = Popup.AddItem(12, 20, 12, 27, "")
            StrIconPath = App.path & "\Garphics\Icons\1.ico"

            If Dir(StrIconPath) <> "" Then
                Item.SetIcon LoadIcon(StrIconPath, 32, 32), xtpPopupItemIconNormal
            End If
        
            Set Item = Popup.AddItem(250, 10, 266, 26, "")
            StrIconPath = App.path & "\Garphics\Icons\ico00002.ico"

            If Dir(StrIconPath) <> "" Then
                Item.SetIcon LoadIcon(StrIconPath, 16, 16), xtpPopupItemIconNormal
            End If

            Item.id = IDCLOSE
            Item.Button = True

            For J = 1 To rs("AlramPriority").value
                SngLeft = 12 + ((J - 1) * 12)
                SngRight = 28 + ((J - 1) * 18)
                Set Item = Popup.AddItem(SngLeft, 65, SngRight, 81, "")
                StrIconPath = App.path & "\Garphics\Icons\flag.ico"

                If Dir(StrIconPath) <> "" Then
                    Item.SetIcon LoadIcon(StrIconPath, 16, 16), xtpPopupItemIconNormal
                End If

                Item.Button = False
            Next J

            '-----------------------
            If i = 0 Then
                PopUpControl(0).Show
            Else
                PopUpControl(i).right = PopUpControl(i - 1).right
                PopUpControl(i).bottom = (PopUpControl(i - 1).bottom - PopUpControl(i - 1).Height)
                PopUpControl(i).AnimateDelay = PopUpControl(i - 1).AnimateDelay + 256
                PopUpControl(i).ShowDelay = PopUpControl(i - 1).ShowDelay + 1000
                PopUpControl(i).Show
            End If

            rs("State").value = 1
            rs.update
            rs.MoveNext
        Next i

    Else
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SetOffice2003Theme(Popup As XtremeSuiteControls.PopUpControl)
    Dim Item As PopupControlItem
    
    Popup.RemoveAllItems
    Popup.Icons.RemoveAll
    
    Set Item = Popup.AddItem(50, 27, 200, 45, " Ã„Ì⁄ ÃÂ«“")
    Item.Bold = True
    
    Set Item = Popup.AddItem(50, 45, 270, 65, "—ﬁ„ «·›« Ê—… ")
    
    Set Item = Popup.AddItem(50, 65, 170, 95, "See comments below" & vbCrLf & "Thanks.")
    Item.TextColor = RGB(0, 61, 178)
    Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    
    '    Set Item = Popup.AddItem(12, 20, 12, 27, "")
    '    Item.SetIcon LoadIcon("Icons\Letter.ico", 32, 32), xtpPopupItemIconNormal
    '
    '    Set Item = Popup.AddItem(250, 10, 266, 26, "")
    '    Item.SetIcon LoadIcon("Icons\ico00002.ico", 16, 16), xtpPopupItemIconNormal
    '    Item.Id = IDCLOSE
    '    Item.Button = True
    '
    '    Set Item = Popup.AddItem(12, 65, 28, 81, "")
    '    Item.SetIcon LoadIcon("Icons\flag.ico", 16, 16), xtpPopupItemIconNormal
    '    Item.Button = True
    '
    '    Set Item = Popup.AddItem(30, 65, 46, 81, "")
    '    Item.SetIcon LoadIcon("Icons\cross.ico", 16, 16), xtpPopupItemIconNormal
    '    Item.Button = True
   
    Popup.VisualTheme = xtpPopupThemeOffice2003
    Popup.setSize 270, 100

End Sub

Private Sub AddManTip()
    Dim Wrap As String
    Dim BolRtl As Boolean
    Dim Msg As String

    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Opt(0), "»ÃÀ ⁄‰ ﬁÿ⁄… «Ê ’‰› œŒ· «·’Ì«‰…", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Opt(1), "≈–« ﬂ‰   —Ìœ «·»ÕÀ ⁄‰ ÃÂ«“ ﬁœ  „  Ã„Ì⁄Â «„ ·«", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl TxtSearch(0), "√œŒ· Â‰« —ﬁ„ «·√Ì’«· «·Œ«’ »⁄„·Ì… «·’Ì«‰…" & Wrap & "·Ì „ «·»ÕÀ »Ê«”ÿ… —ﬁ„ «·√Ì’«·", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl TxtSearch(1), "√œŒ· Â‰« —ﬁ„ «· ﬂ  «·Œ«’ »«·ﬁÿ⁄… «Ê «·’‰› " & Wrap & "·Ì „ «·»ÕÀ »Ê«”ÿ… Â–« «·—ﬁ„", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl TxtSearch(2), "√œŒ· Â‰« —ﬁ„ «·”Ì—Ì«· «·Œ«’ »«·ﬁÿ⁄… «·„—«œ «·»ÕÀ ⁄‰Â« ", BolRtl
    End With

    '
    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl DcboCustomer(0), "«”„ «·⁄„Ì· ... «Œ — «”„ «·⁄„Ì· ·Ì „" & Wrap & "«·»ÕÀ ⁄‰ «·’Ì«‰… «·Œ«’… »Â–« «·⁄„Ì·", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl TxtSearch(3), "«”„ «·⁄„Ì· ... «œŒ· «”„ «·⁄„Ì· ·Ì „" & Wrap & "«·»ÕÀ ⁄‰ «·’Ì«‰… «·Œ«’… »Â–« «·⁄„Ì·", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdSeacrh, "≈÷€ÿ Â‰« Õ Ï  »œ« ⁄„·Ì… «·»ÃÀ" & Wrap & "»‰«¡ ⁄·Ï «·⁄Ê«„· «·„Õœœ…", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl FgSearch, "‰ «∆Ã «·»ÕÀ", BolRtl
    End With

    With TTP
        .Create Me.hWnd, "»ÕÀ «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl DcboCustomer(1), "«”„ «·⁄„Ì· ... «Œ — «”„ «·⁄„Ì· ·Ì „" & Wrap & "«·»ÕÀ ⁄‰ «·ÃÂ«“ «·Œ«’ »«·⁄„Ì·", BolRtl
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SetManDefaluts()
    Dim Dcombos As ClsDataCombos

    Me.Opt(0).value = True
    Opt_Click 0
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 1, Me.DcboCustomer(0), True
    Dcombos.GetCustomersSuppliers 1, Me.DcboCustomer(1), True
    Dcombos.GetItemsNames Me.DcboItemName
End Sub

Public Sub SetDcboSearch()
    'Set cSearchDcbo(0) = New clsDCboSearch
    'Set cSearchDcbo(0).Client = Me.DcboCustomer(0)
    'Set cSearchDcbo(1) = New clsDCboSearch
    'Set cSearchDcbo(1).Client = Me.DcboCustomer(1)
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboItemName
End Sub

Private Sub UnloadArrayControls()
    Dim i As Integer

    'Unload All The Controls Array in the Form
    'Note:  you must unload the
    'Controls before the Container
    For i = 1 To FG.UBound
        FG(i).Visible = False

        DoEvents
        Unload FG(i)
    Next i

    For i = 1 To PicFg.UBound
        PicFg(i).Visible = False

        DoEvents
        Unload PicFg(i)
    Next i

    For i = 1 To itemchart.UBound
        itemchart(i).Visible = False

        DoEvents
        Unload itemchart(i)
    Next i

    For i = 1 To PicChart.UBound
        PicChart(i).Visible = False

        DoEvents
        Unload PicChart(i)
    Next i

    For i = 1 To PopUpControl.UBound
        PopUpControl(i).Hide

        DoEvents
        Unload PopUpControl(i)
    Next i

End Sub

Private Sub ClearTaskPanel()
    Dim i As Integer, J As Integer

    With Me.TaskPanel1

        For i = 1 To Me.TaskPanel1.Groups.count
            Me.TaskPanel1.Groups(i).Items.Clear
        Next i

    End With

End Sub

Private Sub CreateInternetNewsBar()
    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem

    If SystemOptions.UserInterface = ArabicInterface Then
        Set Group = TaskPanel1.Groups.Add(0, "«·√Œ»«— «·⁄«·„Ì…")
        Set Item = Group.Items.Add(0, "·«ÌÊÃœ ≈ ’«· »«·„Êﬁ⁄", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("News").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        
        Set Group = TaskPanel1.Groups.Add(0, "√Œ»«— «·»Ê—’…(„’—)")
        Group.IconIndex = Me.imlViewIcons.ListImages("EgStock").Index
        Set Item = Group.Items.Add(0, "·«ÌÊÃœ ≈ ’«· »«·„Êﬁ⁄", xtpTaskItemTypeText)
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "√Œ»«— «·»Ê—’…(«·”⁄ÊœÌ…)")
        Set Item = Group.Items.Add(0, "·«ÌÊÃœ ≈ ’«· »«·„Êﬁ⁄", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("SAStock").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "«”⁄«— «·⁄„·« („’—)")
        Set Item = Group.Items.Add(0, "·«ÌÊÃœ ≈ ’«· »«·„Êﬁ⁄", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("CreditCart").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "«”⁄«— «·⁄„·« («·”⁄ÊœÏ)")
        Set Item = Group.Items.Add(0, "·«ÌÊÃœ ≈ ’«· »«·„Êﬁ⁄", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("SA").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "«”⁄«— «·–Â»(„’—)")
        Set Item = Group.Items.Add(0, "·«ÌÊÃœ ≈ ’«· »«·„Êﬁ⁄", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("Gold").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "‰’ÌÕ… «·ÌÊ„")
        Set Item = Group.Items.Add(0, "≈ ﬁ «··Â ÕÌÀ„« ﬂ‰ ", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("Tip").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        Set Group = TaskPanel1.Groups.Add(0, "World News")
        Set Item = Group.Items.Add(0, "No Website Connection", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("News").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        
        Set Group = TaskPanel1.Groups.Add(0, "Bourse Egypt")
        Group.IconIndex = Me.imlViewIcons.ListImages("EgStock").Index
        Set Item = Group.Items.Add(0, "No Website Connection", xtpTaskItemTypeText)
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "Bourse Saudi Arabia")
        Set Item = Group.Items.Add(0, "·No Website Connection", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("SAStock").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "Currency Converter Egypt")
        Set Item = Group.Items.Add(0, "No Website Connection", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("CreditCart").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "Currency Converter Saudi Arabia")
        Set Item = Group.Items.Add(0, "No Website Connection", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("SA").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "Gold Prices Egypt")
        Set Item = Group.Items.Add(0, "No Website Connection", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("Gold").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
        Set Group = TaskPanel1.Groups.Add(0, "Today Tip")
        Set Item = Group.Items.Add(0, "≈ ﬁ «··Â ÕÌÀ„« ﬂ‰ ", xtpTaskItemTypeText)
        Group.IconIndex = Me.imlViewIcons.ListImages("Tip").Index
        Item.Bold = True
        Item.SetMargins 0, 0, 0, 0
    End If

End Sub
