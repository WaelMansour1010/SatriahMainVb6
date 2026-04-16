VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOtheExpensAqar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12795
   Icon            =   "FrmOtheExpensAqar.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   12795
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
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Text            =   "modflag"
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8430
      Left            =   0
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   0
      Width           =   12795
      _cx             =   22569
      _cy             =   14870
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   870
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   7515
         Width           =   12780
         _cx             =   22543
         _cy             =   1535
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   435
            Left            =   11595
            TabIndex        =   44
            Top             =   285
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":6852
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   435
            Left            =   8850
            TabIndex        =   45
            Top             =   285
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":D0B4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   435
            Left            =   10170
            TabIndex        =   46
            Top             =   285
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":D44E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   435
            Left            =   7635
            TabIndex        =   47
            Top             =   285
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":13CB0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   435
            Left            =   3390
            TabIndex        =   48
            Top             =   285
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":1404A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   435
            Left            =   255
            TabIndex        =   49
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":145E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   435
            Left            =   5985
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   285
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":1497E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   435
            Left            =   4560
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   285
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   767
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":1B1E0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   435
            Left            =   1410
            TabIndex        =   111
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   285
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
            ButtonStyle     =   1
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":1B57A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   780
         Left            =   0
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   5985
         Width           =   12780
         _cx             =   22543
         _cy             =   1376
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
         Begin VB.CommandButton CMDSENDSMS 
            Caption         =   "«—”«· —”«·Â"
            Height          =   315
            Left            =   6780
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   330
            Width           =   1005
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   570
            Left            =   120
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   120
            Width           =   6615
            _cx             =   11668
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
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   240
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   120
               Width           =   765
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   120
               Width           =   690
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   240
               Index           =   1
               Left            =   1845
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   120
               Width           =   2055
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   240
               Index           =   0
               Left            =   4995
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   120
               Width           =   1350
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   9030
            TabIndex        =   58
            Top             =   225
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   8
            Left            =   11700
            TabIndex        =   59
            Top             =   225
            Width           =   870
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frm2 
         Height          =   705
         Left            =   0
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   750
         Width           =   12780
         _cx             =   22543
         _cy             =   1244
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
         Begin VB.CommandButton cmdSaveR 
            Caption         =   "Õ›Ÿ"
            Height          =   225
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   167
            Top             =   480
            Width           =   1185
         End
         Begin VB.CheckBox chkIsLegalAffairs 
            Alignment       =   1  'Right Justify
            Caption         =   "«Õ«·… ··‘∆Ê‰ «·ﬁ«‰Ê‰Ì…"
            Enabled         =   0   'False
            Height          =   255
            Left            =   10530
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   420
            Width           =   2055
         End
         Begin VB.TextBox txtLegalAffairs 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5160
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   420
            Width           =   2565
         End
         Begin VB.TextBox txtNoteSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   120
            Width           =   1770
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·‘—ﬂ…"
            Height          =   975
            Left            =   4965
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   1920
            Visible         =   0   'False
            Width           =   3810
            Begin MSDataListLib.DataCombo DcbAccount4 
               Height          =   315
               Left            =   120
               TabIndex        =   97
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount3 
               Height          =   315
               Left            =   120
               TabIndex        =   98
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   13
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   14
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·‘—ﬂ…"
            Height          =   975
            Left            =   11025
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1920
            Visible         =   0   'False
            Width           =   1950
            Begin VB.TextBox TxtStay1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   91
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox TxtCivilin1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   90
               Top             =   240
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   12
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   600
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   11
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   10
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   9
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·„ÊŸ›"
            Height          =   975
            Left            =   1290
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1920
            Visible         =   0   'False
            Width           =   3690
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount1 
               Height          =   315
               Left            =   120
               TabIndex        =   85
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   7
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   6
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·„ÊŸ›"
            Height          =   975
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1800
            Visible         =   0   'False
            Width           =   2280
            Begin VB.TextBox TxtCivilin 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   79
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtStay 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   78
               Top             =   600
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   3
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   1
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   4
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   5
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   600
               Width           =   300
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10275
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   120
            Visible         =   0   'False
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   7395
            TabIndex        =   62
            Top             =   120
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   268566529
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   11655
            TabIndex        =   63
            Top             =   1800
            Visible         =   0   'False
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmOtheExpensAqar.frx":21DDC
            Height          =   315
            Left            =   105
            TabIndex        =   101
            Top             =   30
            Width           =   3900
            _ExtentX        =   6879
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
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   315
            Left            =   5160
            TabIndex        =   134
            Top             =   120
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker txtLegalAffairsDate 
            Height          =   285
            Left            =   600
            TabIndex        =   165
            Top             =   390
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            _Version        =   393216
            Format          =   268566529
            CurrentDate     =   38784
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«Õ«·… ··‘∆Ê‰ «·ﬁ«‰Ê‰Ì…"
            Height          =   255
            Left            =   2775
            TabIndex        =   166
            Top             =   420
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”»» «·«Õ«·…"
            Height          =   285
            Index           =   29
            Left            =   7860
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   450
            Width           =   810
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄ "
            Height          =   240
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label Labelbank 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„›—œ"
            Height          =   255
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1800
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblcode 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   270
            Left            =   11910
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   135
            Width           =   870
         End
         Begin VB.Label lbldate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Left            =   9120
            TabIndex        =   64
            Top             =   150
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   780
         Left            =   0
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   0
         Width           =   12780
         _cx             =   22543
         _cy             =   1376
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
         BackColor       =   16777215
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
         Begin ImpulseButton.ISButton btnLast 
            Height          =   285
            Left            =   525
            TabIndex        =   68
            Top             =   225
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":21DF1
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   285
            Left            =   1695
            TabIndex        =   69
            Top             =   225
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":2218B
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   285
            Left            =   855
            TabIndex        =   70
            Top             =   225
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":22525
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   285
            Left            =   1275
            TabIndex        =   71
            Top             =   225
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
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
            ButtonImage     =   "FrmOtheExpensAqar.frx":228BF
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   555
            Left            =   11625
            Picture         =   "FrmOtheExpensAqar.frx":22C59
            Stretch         =   -1  'True
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Index           =   2
            Left            =   6915
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   225
            Width           =   4455
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4545
         Index           =   3
         Left            =   0
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   1440
         Width           =   12780
         _cx             =   22543
         _cy             =   8017
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
         Caption         =   " "
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
         Begin VB.TextBox TxtAccountNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   4305
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   960
            Width           =   2115
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   3840
            Width           =   11640
         End
         Begin VB.TextBox TxtMobile 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   4305
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   2115
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   10695
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   960
            Width           =   870
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   10695
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   870
         End
         Begin VB.TextBox TxtSearch 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   10695
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   870
         End
         Begin VB.ComboBox DcbTypID 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   2985
         End
         Begin MSDataListLib.DataCombo DcbIqara 
            Height          =   315
            Left            =   7620
            TabIndex        =   1
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
            Top             =   240
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitNo 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   240
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitType 
            Height          =   315
            Left            =   4305
            TabIndex        =   2
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   240
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   315
            Left            =   7620
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   600
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   7620
            TabIndex        =   9
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   960
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   2505
            Index           =   1
            Left            =   120
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   1230
            Width           =   12615
            _cx             =   22251
            _cy             =   4419
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
            Caption         =   "»Ì«‰«  «·›« Ê—…"
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
            Begin VB.TextBox TxtName 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   6000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   1080
               Width           =   4665
            End
            Begin VB.TextBox TxtAccountBank 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   8880
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   660
               Width           =   1785
            End
            Begin VB.TextBox TxtBillNo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   8880
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   255
               Width           =   1785
            End
            Begin VB.TextBox TxtValuee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   6000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   240
               Width           =   1785
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1065
               Index           =   0
               Left            =   6000
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   1440
               Width           =   6495
               _cx             =   11456
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
               Caption         =   " «·›« Ê—… «·„” Õﬁ… ⁄‰ «·› —…"
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
               Begin MSComCtl2.DTPicker FromDate 
                  Height          =   285
                  Left            =   2490
                  TabIndex        =   18
                  Top             =   180
                  Width           =   1635
                  _ExtentX        =   2884
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   251789313
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   285
                  Left            =   2490
                  TabIndex        =   20
                  Top             =   585
                  Width           =   1635
                  _ExtentX        =   2884
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   251789313
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal FromDateH 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   19
                  Top             =   180
                  Width           =   1920
                  _ExtentX        =   3387
                  _ExtentY        =   503
               End
               Begin Dynamic_Byte.NourHijriCal ToDateH 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   21
                  Top             =   585
                  Width           =   1920
                  _ExtentX        =   3387
                  _ExtentY        =   503
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   255
                  Left            =   4245
                  TabIndex        =   124
                  Top             =   210
                  Width           =   1035
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   240
                  Left            =   4245
                  TabIndex        =   123
                  Top             =   630
                  Width           =   1035
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2265
               Index           =   4
               Left            =   0
               TabIndex        =   135
               TabStop         =   0   'False
               Top             =   135
               Width           =   5895
               _cx             =   10398
               _cy             =   3995
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
               Begin VSFlex8UCtl.VSFlexGrid UnitsGrid 
                  Height          =   1515
                  Left            =   120
                  TabIndex        =   22
                  Top             =   390
                  Width           =   5685
                  _cx             =   10028
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
                  Rows            =   1
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOtheExpensAqar.frx":2842B
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   300
                  Left            =   3600
                  TabIndex        =   136
                  Top             =   1845
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmOtheExpensAqar.frx":284EA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÊÕœ«  «·„‘ —ﬂ… "
                  ForeColor       =   &H00FF0000&
                  Height          =   450
                  Index           =   45
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   135
                  Width           =   1770
               End
            End
            Begin MSDataListLib.DataCombo DcbBanck 
               Height          =   315
               Left            =   6000
               TabIndex        =   16
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   660
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   195
               Index           =   26
               Left            =   7830
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   660
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ ’«Õ» «·Õ”«»"
               Height          =   225
               Index           =   25
               Left            =   10980
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   1080
               Width           =   1590
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ Õ”«» «·»‰ﬂ «·„”œœ „‰Â"
               Height          =   225
               Index           =   6
               Left            =   10980
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   660
               Width           =   1590
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·›« Ê—…"
               Height          =   225
               Index           =   1
               Left            =   11220
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   255
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·›« Ê—…"
               Height          =   225
               Index           =   7
               Left            =   7620
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   255
               Width           =   1110
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   2505
            Index           =   2
            Left            =   120
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   1320
            Width           =   12615
            _cx             =   22251
            _cy             =   4419
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
            Caption         =   " »Ì«‰«  «· ’›Ì…"
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
            Begin VB.TextBox TxtDiscount2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   7350
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   1710
               Width           =   1005
            End
            Begin VB.TextBox TxtTotalAfterIns 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000001&
               Height          =   300
               Left            =   3840
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   2040
               Width           =   2145
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   9240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1680
               Width           =   1305
            End
            Begin VB.TextBox TxtNet 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000001&
               Height          =   300
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   2040
               Width           =   2145
            End
            Begin VB.TextBox TxtWindows 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   1680
               Width           =   2145
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000001&
               Height          =   300
               Left            =   7920
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   2040
               Width           =   2625
            End
            Begin VB.TextBox TxtPaints 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3840
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   1710
               Width           =   2145
            End
            Begin VB.TextBox TxtElectricity 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   675
               Width           =   2145
            End
            Begin VB.TextBox TxtDelayDay 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   7920
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   675
               Width           =   2625
            End
            Begin VB.TextBox TxtInsurance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   7920
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   1365
               Width           =   2625
            End
            Begin VB.TextBox TxtNoliquidation 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
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
               Left            =   7920
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1005
               Width           =   2625
            End
            Begin VB.ComboBox DcbStatusOper 
               Height          =   315
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   345
               Width           =   2625
            End
            Begin VB.TextBox TxtMaintOther 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   1365
               Width           =   2145
            End
            Begin VB.TextBox TxtMaintClean 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3840
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   1365
               Width           =   2145
            End
            Begin VB.TextBox TxtMaintkitchen 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   345
               Width           =   2145
            End
            Begin VB.TextBox TxtMaintDoors 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
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
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   1005
               Width           =   2145
            End
            Begin VB.TextBox TxtMaintCondition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
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
               Left            =   3840
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   1005
               Width           =   2145
            End
            Begin VB.TextBox TxtMaintenance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3840
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   345
               Width           =   2145
            End
            Begin VB.TextBox TxtRemainRent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3840
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   675
               Width           =   2145
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·Œ’„"
               Height          =   300
               Index           =   30
               Left            =   8070
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   1710
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì »⁄œ «·Œ’„"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   28
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   2040
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·Œ’„"
               Height          =   300
               Index           =   27
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   1680
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì »⁄œ Œ’„ «· «„Ì‰"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   24
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   2040
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê«›–"
               Height          =   180
               Index           =   21
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   1680
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’Ì«‰… „ÿ»Œ"
               Height          =   180
               Index           =   13
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   345
               Width           =   1110
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "›« Ê—… «·ﬂÂ—»«¡"
               Height          =   195
               Index           =   17
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   675
               Width           =   1110
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               ForeColor       =   &H00800000&
               Height          =   300
               Index           =   23
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   2040
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "œÂ«‰« "
               Height          =   180
               Index           =   22
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   1710
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «Ì«„ «· «ŒÌ— ›Ì ”œ«œ «·«ÌÃ«—"
               Height          =   195
               Index           =   20
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   675
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " √„Ì‰ «·„” √Ã—"
               Height          =   300
               Index           =   9
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   1365
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «· ’›Ì… «·ÌœÊÌ"
               Height          =   195
               Index           =   19
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   1005
               Width           =   1470
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «· ’›Ì…"
               Height          =   180
               Index           =   0
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   345
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ—Ï"
               Height          =   180
               Index           =   18
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   1365
               Width           =   1110
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’Ì«‰… „ﬂÌ›« "
               Height          =   195
               Index           =   11
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   1005
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ÿ«›…"
               Height          =   180
               Index           =   16
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   1365
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’Ì«‰… «»Ê«»"
               Height          =   195
               Index           =   12
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   1005
               Width           =   1110
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’Ì«‰… ﬂÂ—»«¡/”»«ﬂ…"
               Height          =   180
               Index           =   10
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   345
               Width           =   1950
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„ »ﬁÌ «·«ÌÃ«—"
               Height          =   180
               Index           =   8
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   675
               Width           =   1950
            End
         End
         Begin MSComCtl2.DTPicker PayedDate 
            Height          =   315
            Left            =   1785
            TabIndex        =   11
            Top             =   960
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Format          =   250871809
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal PayedDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ’›Ì…"
            Height          =   195
            Left            =   3240
            TabIndex        =   130
            Top             =   990
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ Õ‹  «·ﬂÂ—»«¡"
            Height          =   195
            Index           =   3
            Left            =   6195
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   960
            Width           =   1560
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Left            =   11715
            TabIndex        =   110
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃÊ«·"
            Height          =   195
            Index           =   0
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„’›Ì"
            Height          =   195
            Index           =   37
            Left            =   11445
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «·„” √Ã—"
            Height          =   195
            Index           =   5
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   600
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «·⁄ﬁ«— "
            Height          =   195
            Index           =   4
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·ÊÕœ…"
            Height          =   195
            Index           =   15
            Left            =   6495
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÊÕœ…"
            Height          =   195
            Index           =   14
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   195
            Index           =   2
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   600
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   780
         Left            =   0
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   6720
         Width           =   12780
         _cx             =   22543
         _cy             =   1376
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
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E2E9E9&
            Caption         =   "› Õ ”‰œ ﬁ»÷"
            Height          =   375
            Left            =   1320
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ «·”‰œ« "
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   160
            Top             =   240
            Width           =   1095
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   615
            Index           =   8
            Left            =   7800
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   120
            Width           =   4815
            _cx             =   8493
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
            Begin VB.OptionButton Optx 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ·"
               Height          =   195
               Index           =   0
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton Optx 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ··⁄ﬁ«—"
               Height          =   195
               Index           =   1
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton Optx 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ··„” √Ã—"
               Height          =   195
               Index           =   3
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ŒÌ«—«  «·⁄—÷"
               Height          =   270
               Index           =   15
               Left            =   3600
               TabIndex        =   143
               Top             =   120
               Width           =   1110
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   615
            Index           =   6
            Left            =   3000
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   120
            Width           =   4695
            _cx             =   8281
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
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   -240
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   120
               Width           =   1515
            End
            Begin VB.CommandButton Command9 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   1320
               RightToLeft     =   -1  'True
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   120
               Width           =   1035
            End
            Begin VB.CommandButton Command8 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ‘› Õ”«»"
               Height          =   375
               Left            =   120
               MaskColor       =   &H00C0E0FF&
               RightToLeft     =   -1  'True
               Style           =   1  'Graphical
               TabIndex        =   145
               Top             =   120
               Width           =   1155
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   120
               Width           =   930
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   -10680
      TabIndex        =   75
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   -360
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
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   11400
      Top             =   2280
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
            Picture         =   "FrmOtheExpensAqar.frx":28A84
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":28E1E
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":291B8
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":29552
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":298EC
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":29C86
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":2A020
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtheExpensAqar.frx":2A5BA
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmOtheExpensAqar"
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
 Dim ii As Long

Private Sub cmdSaveR_Click()
'RsSavRec.Resync
  RsSavRec.Fields("LegalAffairsDate").value = txtLegalAffairsDate.value
        
        RsSavRec("LegalAffairs").value = Trim(Me.txtLegalAffairs.text)
        If chkIsLegalAffairs.value = vbChecked Then
            RsSavRec.Fields("isLegalAffairs").value = 1
        Else
            RsSavRec.Fields("isLegalAffairs").value = 0
        End If
RsSavRec.update
RsSavRec.Resync adAffectCurrent
End Sub

Private Sub CMDSENDSMS_Click()
'0 manual
'1 save
'2 Print

SendMessage (0)
End Sub
Function SendMessage(currentOpt As Integer)
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", opt)
  If opt = currentOpt Then
  
      CompanyName = cOptions.ArabCompanyName '& CHR(13) & CurrentBranchName
     companyphone = cOptions.Company_Mobile
  '«·„” √Ã—
  
 
'0  " ’›Ì…"
'1 "›« Ê—… ﬂÂ—»«¡"
 
If DcbTypID.ListIndex = 0 Then
 Msg = " ÌÊÃœ „»·€  " & TxtTotal.text & "  „” Õﬁ „‰  ’›Ì… ··ÊÕœ…    " & DcbUnitNo.text & " »«·⁄ﬁ«—" & DcbIqara.text
 Else
 Msg = "  „ «’œ«— ›« Ê—… ﬂÂ—»«¡ »„»·€ " & TxtValuee.text & "  ··ÊÕœ…    " & DcbUnitNo.text & " »«·⁄ﬁ«—" & DcbIqara.text
 End If
 
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(val(dcCustomer.BoundText)))
 

MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function
 
 
 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = " »‰« ⁄·Ï  ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡ »—ﬁ„" & CHR(13)
des = des & "  " & txtNoteSerial1.text & " " & CHR(13)
des = des & "   ··„” √Ã—   " & " " & Me.dcCustomer.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblOtheExpensAqar"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1.text)
If val(DcbTypID.ListIndex) = 0 Then
Notevalue = val(TxtNet.text)
Else
Notevalue = val(TxtValuee.text)
End If
 notytype = 9067
 BranchID = val(Dcbranch.BoundText)
NoteDate = (XPDtbTrans.value)
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, Txt_DateHigri.value         ', recordDateH.value
                                              TxtNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial
                                     Else
                                                 If TxtNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.text = NoteID
                                                                TxtNoteSerial.text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.text), BranchID, user_id, NoteDate, Notevalue, des
RsSavRec.Resync adAffectCurrent
     End If
End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date, Optional NotValue As Double, Optional des As String)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
   ' Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCode As String
    Dim rs As New ADODB.Recordset
   ' Dim notes_serial As String
    Dim notes_id As String
   ' Dim j As Integer
   ' Dim ColumnName As String
    Dim RenreAccount As String
    Msg = des

 notes_id = general_noteid
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
    Dim Branch As Integer
    
    BranchID = 1
line_no = 1
    BranchID = val(Dcbranch.BoundText)
 StrAccountCode = get_account_code_branch(143, my_branch)
 RenreAccount = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, RenreAccount, NotValue, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, NotValue, 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
            
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function

Private Sub btnQuery_Click()
FrmIqarWaiverSet.m_RetrunType = 8
Load FrmIqarWaiverSet
FrmIqarWaiverSet.m_RetrunType = 8
FrmIqarWaiverSet.show
End Sub

Private Sub Cmd_Click()
RemoveGridRow
End Sub
Private Sub RemoveGridRow()
If Me.TxtModFlg.text <> "R" Then
    With Me.UnitsGrid
        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With
 End If
End Sub

Private Sub Command1_Click()
If Me.TxtModFlg.text = "R" Then
Unload FrmSanadatOFContract
Load FrmSanadatOFContract
FrmSanadatOFContract.Label1(0).Caption = txtNoteSerial1.text
FrmSanadatOFContract.Indx = 1
'FrmSanadatOFContract.TxtNotID.Text = val(TxtNotID.Text)
FrmSanadatOFContract.TxtContNo.text = val(TxtSerial1.text)
FrmSanadatOFContract.show
End If
End Sub
 Function CheckPayetTranss(ContNo As Long) As Boolean
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
CheckPayetTranss = False
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *   from dbo.Notes Where (ContNo =" & ContNo & ") and dbo.Notes.CashingType=13 "
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   CheckPayetTranss = True
   Else
   CheckPayetTranss = False
   End If
End Function
Private Sub Command10_Click()
    If checkApility("FrmOtheExpensAqar") = False Then
                Exit Sub
            End If
FrmCashing1.show
FrmCashing1.newrecord
FrmCashing1.DCboCashType.ListIndex = 13
FrmCashing1.Dcbranch.BoundText = Me.Dcbranch.BoundText
FrmCashing1.txtContractNo.text = txtNoteSerial1.text
FrmCashing1.TxtContNo.text = TxtSerial1.text

End Sub

Private Sub Command8_Click()
If val(dcCustomer.BoundText) <> 0 Then
Dim StrTempAccountCode As String
            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
                   StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
            ShowReport StrTempAccountCode, dcCustomer.text, FirstPeriod, Date
End If
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub DcbIqara_Change()
DcbIqara_Click (0)
End Sub
Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub
Private Sub DcbIqara_Click(Area As Integer)
DcbUnitType_Change
If val(DcbIqara.BoundText) = 0 Then: Exit Sub
    Dim EmpCode  As String
    GetIqarCode , , DcbIqara.BoundText, EmpCode
    Me.TxtSearch.text = EmpCode
End Sub

Private Sub DcboEmp_Change()
DcboEmp_Click (0)
End Sub

Private Sub DcboEmp_Click(Area As Integer)
    If val(DcboEmp.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmp.BoundText, EmpCode
    TxtEmployeeID.text = EmpCode
End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    If ChekSanNumber(Current_branch, 64) = True Then
          txtNoteSerial1.text = ""
      End If
      TxtNoteSerial.text = ""
End Sub

Private Sub DcbTypID_Change()
Ele(2).Visible = False
Ele(1).Visible = False
If val(DcbTypID.ListIndex) = 0 Then
Label1(37).Caption = "«·„’›Ì"
Label4.Caption = " «—ÌŒ «· ’›Ì…"
Ele(2).Visible = True
Else
Label4.Caption = " «—ÌŒ «·”œ«œ"
Ele(1).Visible = True
Label1(37).Caption = "„”ƒÊ· «·ﬂÂ—»«¡"
End If
End Sub

Private Sub DcbTypID_Click()
DcbTypID_Change
End Sub

Private Sub DcbUnitType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub
Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub
Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
Set Dcombos = New ClsDataCombos
If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)
idd1 = val(DcbUnitType.BoundText)
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
End If
End Sub
Private Sub DcbUnitNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub dcCustomer_Click(Area As Integer)
dcCustomer_Change
End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Function GetCutMobil(Optional CutID As Double) As String
Dim sql As String
Dim Rs3 As ADODB.Recordset
sql = " SELECT     CusID, Cus_mobile"
sql = sql & " From dbo.TblCustemers"
sql = sql & " Where (cusid = " & CutID & ")"
Set Rs3 = New ADODB.Recordset
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetCutMobil = IIf(IsNull(Rs3("Cus_mobile").value), "", Rs3("Cus_mobile").value)
Else
GetCutMobil = ""
End If
End Function

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    If ChekSanNumber(Current_branch, 64) = True Then
         txtNoteSerial1.Enabled = False
    Else
         txtNoteSerial1.Enabled = True
     End If
                           If SystemOptions.SpecialVersion = True Then
Ele(6).Visible = False
    End If
    conection = "select * from TblOtheExpensAqar where  BranchID in(" & Current_branchSql & ") order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetBanks Me.DcbBanck
If SystemOptions.UserInterface = ArabicInterface Then
With DcbStatusOper
.Clear
.AddItem "⁄«œÌ…"
.AddItem "Â—Ê»"
End With
With DcbTypID
.Clear
.AddItem " ’›Ì…"
.AddItem "›« Ê—… ﬂÂ—»«¡"
End With
Else
With DcbStatusOper
.Clear
.AddItem "Normal"
.AddItem "Escape"
End With
With DcbTypID
.Clear
.AddItem "Evacuation"
.AddItem "Electricity"
End With
End If


If SystemOptions.CanEditLegalAffairs Then
    chkIsLegalAffairs.Enabled = True
    txtLegalAffairs.Enabled = True
    txtLegalAffairsDate.Enabled = True
Else
    chkIsLegalAffairs.Enabled = False
    txtLegalAffairs.Enabled = False
    txtLegalAffairsDate.Enabled = False
End If

ReloadCombos
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
    ''/////
    Dim sql As String
    Dim i As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    If Me.TxtModFlg.text = "E" Then
     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    Cn.Execute "Delete from TblOtheExpensAqarDet Where ExpensAqID=" & val(TxtSerial1.text) & ""
    End If
    If txtNoteSerial1.text = "" Then
     txtNoteSerial1.text = Voucher_coding(val(Dcbranch.BoundText), XPDtbTrans.value, 64, 64)
     End If
    RsSavRec.Fields("NoteSerial1").value = IIf(Me.txtNoteSerial1 <> "", Trim(txtNoteSerial1.text), Null)
    RsSavRec.Fields("DelayDay").value = val(TxtDelayDay.text)
    RsSavRec.Fields("Noliquidation").value = val(TxtNoliquidation.text)
    RsSavRec.Fields("Paints").value = val(TxtPaints.text)
    RsSavRec.Fields("Windows").value = val(TxtWindows.text)
    RsSavRec.Fields("Total").value = val(TxtTotal.text)
    RsSavRec.Fields("Net").value = val(TxtNet.text)
    RsSavRec.Fields("BankID").value = val(Me.DcbBanck.BoundText)
    RsSavRec.Fields("Name").value = TxtName.text
    RsSavRec.Fields("TotalAfterIns").value = val(TxtTotalAfterIns.text)
    RsSavRec.Fields("Discount").value = val(TxtDiscount.text)
    RsSavRec.Fields("Discount2").value = val(TxtDiscount2.text)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecordDateH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("Valuee").value = val(TxtValuee.text)
    RsSavRec.Fields("AqarID").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val((DcbIqara.BoundText)), Null)
    RsSavRec.Fields("UnitTypID").value = IIf(val(DcbUnitType.BoundText) <> 0, val(DcbUnitType.BoundText), Null)
    RsSavRec.Fields("UnitID").value = IIf(val(DcbUnitNo.BoundText) <> 0, val(DcbUnitNo.BoundText), Null)
    RsSavRec.Fields("CusID").value = IIf(val(dcCustomer.BoundText) <> 0, val(dcCustomer.BoundText), Null)
    RsSavRec.Fields("EmpID").value = IIf(val(DcboEmp.BoundText) <> 0, val(DcboEmp.BoundText), Null)
    RsSavRec.Fields("Mobile").value = TxtMobile.text
    RsSavRec.Fields("TypID").value = val(DcbTypID.ListIndex)
    RsSavRec.Fields("BillNo").value = TxtBillNo.text
    RsSavRec.Fields("AccountNo").value = TxtAccountNo.text
    RsSavRec.Fields("AccountBank").value = TxtAccountBank.text
    RsSavRec.Fields("Remarks").value = TxtRemarks.text
    RsSavRec.Fields("PayedDateH").value = PayedDateH.value
    RsSavRec.Fields("PayedDate").value = PayedDate.value
     RsSavRec.Fields("LegalAffairsDate").value = txtLegalAffairsDate.value
        
        RsSavRec("LegalAffairs").value = Trim(Me.txtLegalAffairs.text)
        If chkIsLegalAffairs.value = vbChecked Then
            RsSavRec.Fields("isLegalAffairs").value = 1
        Else
            RsSavRec.Fields("isLegalAffairs").value = 0
        End If

    
    RsSavRec.Fields("Fromdateh").value = FromDateH.value
    RsSavRec.Fields("Fromdate").value = FromDate.value
    RsSavRec.Fields("TodateH").value = ToDateH.value
    RsSavRec.Fields("Todate").value = ToDate.value
    If val(DcbTypID.ListIndex) = 0 Then
    RsSavRec.Fields("StatusOper").value = val(Me.DcbStatusOper.ListIndex)
    Else
    RsSavRec.Fields("StatusOper").value = -1
    End If
    RsSavRec.Fields("RemainRent").value = val(TxtRemainRent.text)
    RsSavRec.Fields("Electricity").value = val(TxtElectricity.text)
    RsSavRec.Fields("Maintenance").value = val(TxtMaintenance.text)
    RsSavRec.Fields("MaintCondition").value = val(TxtMaintCondition.text)
    RsSavRec.Fields("MaintDoors").value = val(TxtMaintDoors.text)
    RsSavRec.Fields("MaintKitchen").value = val(TxtMaintkitchen.text)
    RsSavRec.Fields("MaintClean").value = val(TxtMaintClean.text)
    RsSavRec.Fields("MaintOther").value = val(TxtMaintOther.text)
    RsSavRec.Fields("Insurance").value = val(TxtInsurance.text)
    RsSavRec.update
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = "Select * from TblOtheExpensAqarDet where 1=-1"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
     With Me.UnitsGrid
     For i = 1 To .rows - 1
     If val(.TextMatrix(i, .ColIndex("UnitID"))) <> 0 Then
     Rs3.AddNew
     Rs3("ExpensAqID").value = val(TxtSerial1.text)
     Rs3("UnitID").value = IIf(val(.TextMatrix(i, .ColIndex("UnitID"))) = 0, Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
     Rs3("UnitType").value = IIf(val(.TextMatrix(i, .ColIndex("UnitType"))) = 0, 0, val(.TextMatrix(i, .ColIndex("UnitType"))))
     Rs3.update
     End If
     Next i
     End With
     createVoucher
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
               
                Me.Refresh
                FiLLTXT
                
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
                Me.Refresh
                FiLLTXT
                
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
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
    
    Me.TxtNoteSerial.text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
    Me.TxtNoteID.text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
    
    Me.txtNoteSerial1.text = IIf(IsNull(RsSavRec.Fields("NoteSerial1").value), "", RsSavRec.Fields("NoteSerial1").value)
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), "", RsSavRec.Fields("RecordDateH").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    TxtValuee.text = IIf(IsNull(RsSavRec.Fields("Valuee").value), "", RsSavRec.Fields("Valuee").value)
    Me.DcbIqara.BoundText = IIf(IsNull(RsSavRec.Fields("AqarID").value), "", RsSavRec.Fields("AqarID").value)
    Me.DcbUnitType.BoundText = IIf(IsNull(RsSavRec.Fields("UnitTypID").value), "", RsSavRec.Fields("UnitTypID").value)
    Me.DcbUnitNo.BoundText = IIf(IsNull(RsSavRec.Fields("UnitID").value), "", RsSavRec.Fields("UnitID").value)
    Me.dcCustomer.BoundText = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    Me.DcboEmp.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    TxtMobile.text = IIf(IsNull(RsSavRec.Fields("Mobile").value), "", RsSavRec.Fields("Mobile").value)
    Me.DcbTypID.ListIndex = IIf(IsNull(RsSavRec.Fields("TypID").value), -1, RsSavRec.Fields("TypID").value)
    TxtBillNo.text = IIf(IsNull(RsSavRec.Fields("BillNo").value), "", RsSavRec.Fields("BillNo").value)
    TxtAccountNo.text = IIf(IsNull(RsSavRec.Fields("AccountNo").value), "", RsSavRec.Fields("AccountNo").value)
    TxtAccountBank.text = IIf(IsNull(RsSavRec.Fields("AccountBank").value), "", RsSavRec.Fields("AccountBank").value)
    TxtRemarks.text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    PayedDate.value = IIf(IsNull(RsSavRec.Fields("PayedDate").value), Date, RsSavRec.Fields("PayedDate").value)
    PayedDateH.value = IIf(IsNull(RsSavRec.Fields("PayedDateH").value), ToHijriDate(PayedDate.value), RsSavRec.Fields("PayedDateH").value)
    FromDate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value)
    FromDateH.value = IIf(IsNull(RsSavRec.Fields("FromDateH").value), ToHijriDate(FromDate.value), RsSavRec.Fields("FromDateH").value)
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
    ToDateH.value = IIf(IsNull(RsSavRec.Fields("ToDateH").value), ToHijriDate(ToDate.value), RsSavRec.Fields("ToDateH").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbStatusOper.ListIndex = IIf(IsNull(RsSavRec.Fields("StatusOper").value), -1, RsSavRec.Fields("StatusOper").value)
    TxtRemainRent.text = IIf(IsNull(RsSavRec.Fields("RemainRent").value), 0, RsSavRec.Fields("RemainRent").value)
    TxtElectricity.text = IIf(IsNull(RsSavRec.Fields("Electricity").value), 0, RsSavRec.Fields("Electricity").value)
    TxtMaintenance.text = IIf(IsNull(RsSavRec.Fields("Maintenance").value), 0, RsSavRec.Fields("Maintenance").value)
    TxtMaintCondition.text = IIf(IsNull(RsSavRec.Fields("MaintCondition").value), 0, RsSavRec.Fields("MaintCondition").value)
    TxtMaintDoors.text = IIf(IsNull(RsSavRec.Fields("MaintDoors").value), 0, RsSavRec.Fields("MaintDoors").value)
    TxtMaintkitchen.text = IIf(IsNull(RsSavRec.Fields("MaintKitchen").value), 0, RsSavRec.Fields("MaintKitchen").value)
    TxtMaintClean.text = IIf(IsNull(RsSavRec.Fields("MaintClean").value), 0, RsSavRec.Fields("MaintClean").value)
    TxtMaintOther.text = IIf(IsNull(RsSavRec.Fields("MaintOther").value), 0, RsSavRec.Fields("MaintOther").value)
    TxtInsurance.text = IIf(IsNull(RsSavRec.Fields("Insurance").value), 0, RsSavRec.Fields("Insurance").value)
    ''///
    TxtDelayDay.text = IIf(IsNull(RsSavRec.Fields("DelayDay").value), "", RsSavRec.Fields("DelayDay").value)
    TxtNoliquidation.text = IIf(IsNull(RsSavRec.Fields("Noliquidation").value), "", RsSavRec.Fields("Noliquidation").value)
    TxtPaints.text = IIf(IsNull(RsSavRec.Fields("Paints").value), "", RsSavRec.Fields("Paints").value)
    TxtWindows.text = IIf(IsNull(RsSavRec.Fields("Windows").value), "", RsSavRec.Fields("Windows").value)
    TxtTotal.text = IIf(IsNull(RsSavRec.Fields("Total").value), 0, RsSavRec.Fields("Total").value)
    TxtNet.text = IIf(IsNull(RsSavRec.Fields("Net").value), 0, RsSavRec.Fields("Net").value)
    Me.DcbBanck.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value)
    TxtName.text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    TxtTotalAfterIns.text = IIf(IsNull(RsSavRec.Fields("TotalAfterIns").value), 0, RsSavRec.Fields("TotalAfterIns").value)
    TxtDiscount.text = IIf(IsNull(RsSavRec.Fields("Discount").value), 0, RsSavRec.Fields("Discount").value)
    TxtDiscount2.text = IIf(IsNull(RsSavRec.Fields("Discount2").value), 0, RsSavRec.Fields("Discount2").value)
    
    
     
    If IsNull(RsSavRec.Fields("isLegalAffairs").value) Then
        chkIsLegalAffairs.value = vbUnchecked
    Else
        If RsSavRec.Fields("isLegalAffairs").value = 0 Then
            chkIsLegalAffairs.value = vbUnchecked
        Else
            chkIsLegalAffairs.value = vbChecked
        End If
      
    End If
    
    Me.txtLegalAffairs.text = RsSavRec("LegalAffairs").value & ""
        
    txtLegalAffairsDate.value = IIf(IsNull(RsSavRec.Fields("LegalAffairsDate").value), Date, RsSavRec.Fields("LegalAffairsDate").value)
    
  FillGrid
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
ErrTrap:
End Sub
Sub FillGrid()
Dim sql As String
Dim i As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
UnitsGrid.Clear flexClearScrollable, flexClearEverything
UnitsGrid.rows = 2
sql = " SELECT     dbo.TblOtheExpensAqarDet.ID, dbo.TblOtheExpensAqarDet.ExpensAqID, dbo.TblOtheExpensAqarDet.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee,"
sql = sql & "                       dbo.TblOtheExpensAqarDet.unitid , dbo.TblAqarDetai.unitno"
sql = sql & " FROM         dbo.TblOtheExpensAqarDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblAqarDetai ON dbo.TblOtheExpensAqarDet.UnitID = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblAkarUnit ON dbo.TblOtheExpensAqarDet.UnitType = dbo.TblAkarUnit.id"
sql = sql & " Where (dbo.TblOtheExpensAqarDet.ExpensAqID = " & val(TxtSerial1.text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
With UnitsGrid
.rows = Rs3.RecordCount + 1
For i = 1 To Rs3.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs3("UnitID").value), 0, Rs3("UnitID").value)
.TextMatrix(i, .ColIndex("UnitType")) = IIf(IsNull(Rs3("UnitType").value), 0, Rs3("UnitType").value)
.TextMatrix(i, .ColIndex("unitno")) = IIf(IsNull(Rs3("unitno").value), "", Rs3("unitno").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("name").value), "", Rs3("name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("namee").value), "", Rs3("namee").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub

Private Sub FromDate_Change()
 If Me.TxtModFlg.text <> "R" Then
              FromDateH.value = ToHijriDate(FromDate.value)
   End If
End Sub

Private Sub Fromdateh_LostFocus()
 If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            FromDate.value = ToGregorianDate(FromDateH.value)
   End If
End Sub

Private Sub ISButton1_Click()
print_report
SendMessage (1)
End Sub

Private Sub ISButton3_Click()
On Error Resume Next
ShowAttachments TxtSerial1.text, "25042017111"
ErrTrap:
End Sub

Private Sub Optx_Click(index As Integer)
'On Error Resume Next
Dim My_SQL As String
RsSavRec.Close

Select Case index

Case 0
 My_SQL = " select * from TblOtheExpensAqar "
 My_SQL = My_SQL & "  where  BranchID in(" & Current_branchSql & ")"
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

Case 1
 My_SQL = " select * from TblOtheExpensAqar "
 My_SQL = My_SQL & "  where  BranchID in(" & Current_branchSql & ")"
 My_SQL = My_SQL & " and AqarID=" & val(DcbIqara.BoundText)
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
Case 3

 My_SQL = " select * from TblOtheExpensAqar "
 My_SQL = My_SQL & "  where  BranchID in(" & Current_branchSql & ")"
 My_SQL = My_SQL & " and CusID=" & val(Me.dcCustomer.BoundText)
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       
End Select
BtnFirst_Click
End Sub
Private Sub PayedDate_Change()
  If Me.TxtModFlg.text <> "R" Then
              PayedDateH.value = ToHijriDate(PayedDate.value)
   End If
End Sub

Private Sub PayedDateH_LostFocus()
 If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            PayedDate.value = ToGregorianDate(PayedDateH.value)
   End If
End Sub

Private Sub ToDate_Change()
 If Me.TxtModFlg.text <> "R" Then
              ToDateH.value = ToHijriDate(ToDate.value)
   End If
End Sub

Private Sub ToDateH_LostFocus()
 If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            ToDate.value = ToGregorianDate(ToDateH.value)
   End If
End Sub

Private Sub Txt_DateHigri_LostFocus()
 If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
   End If
End Sub


Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
    End If
End Sub
Private Sub dcCustomer_Change()
  If val(dcCustomer.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , dcCustomer.BoundText, EmpCode
    Me.Text15.text = EmpCode
   Me.TxtMobile.text = GetCutMobil(val(dcCustomer.BoundText))
End Sub

Private Sub TxtDelayDay_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDelayDay.text, 0)
End Sub

Private Sub txtDiscount_Change()
ClaCulte
End Sub

Private Sub TxtDiscount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDiscount.text, 0)
End Sub

Private Sub TxtElectricity_Change()
ClaCulte
End Sub

Private Sub TxtElectricity_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtElectricity.text, 0)
End Sub

Private Sub TxtEmployeeID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.text, True)
End If
End Sub

Private Sub TxtInsurance_Change()
ClaCulte
End Sub

Private Sub TxtInsurance_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtInsurance.text, 0)
End Sub

Private Sub TxtMaintClean_Change()
ClaCulte
End Sub

Private Sub TxtMaintClean_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMaintClean.text, 0)
End Sub

Private Sub TxtMaintCondition_Change()
ClaCulte
End Sub

Private Sub TxtMaintCondition_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMaintCondition.text, 0)
End Sub

Private Sub TxtMaintDoors_Change()
ClaCulte
End Sub

Private Sub TxtMaintDoors_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMaintDoors.text, 0)
End Sub

Private Sub TxtMaintenance_Change()
ClaCulte
End Sub

Private Sub TxtMaintenance_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMaintenance.text, 0)
End Sub

Private Sub TxtMaintkitchen_Change()
ClaCulte
End Sub

Private Sub TxtMaintkitchen_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMaintkitchen.text, 0)
End Sub

Private Sub TxtMaintOther_Change()
ClaCulte
End Sub

Private Sub TxtMaintOther_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMaintOther.text, 0)
End Sub

Private Sub TxtPaints_Change()
ClaCulte
End Sub

Private Sub TxtPaints_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPaints.text, 0)
End Sub

Private Sub TxtRemainRent_Change()
ClaCulte
End Sub

Private Sub TxtRemainRent_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtRemainRent.text, 0)
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim AqrID As Double
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.text, AqrID
        DcbIqara.BoundText = AqrID
        DcbIqara_Click (0)
    End If
End Sub

Sub ClaCulte()
If Me.TxtModFlg.text <> "R" Then
Me.TxtTotal.text = val(TxtMaintCondition.text) + val(TxtRemainRent.text) + val(TxtMaintenance.text) + val(TxtMaintClean.text) + val(TxtPaints.text)
Me.TxtTotal.text = val(Me.TxtTotal.text) + val(Me.TxtMaintkitchen.text) + val(Me.TxtElectricity.text) + val(Me.TxtMaintDoors.text) + val(Me.TxtMaintOther.text) + val(Me.TxtWindows.text)
TxtTotalAfterIns.text = val(Me.TxtTotal.text) - val(Me.TxtInsurance.text)
TxtNet.text = val(Me.TxtTotalAfterIns.text) - val(Me.TxtDiscount.text)
End If
End Sub

Private Sub TxtValuee_Change()
If Me.TxtModFlg.text <> "R" And val(DcbTypID.ListIndex) = 1 Then
TxtNet.text = val(TxtValuee.text)
End If
End Sub

Private Sub TxtValuee_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValuee.text, 0)
End Sub

Private Sub TxtWindows_Change()
ClaCulte
End Sub

Private Sub TxtWindows_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtWindows.text, 0)
End Sub

Private Sub UnitsGrid_AfterEdit(ByVal row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
With UnitsGrid
 Select Case .ColKey(Col)
 Case "Name"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitType"), False, True)
                .TextMatrix(row, .ColIndex("UnitType")) = StrAccountCode
Case "unitno"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(row, .ColIndex("UnitID")) = StrAccountCode
           If row = .rows - 1 Then
            .rows = .rows + 1
        End If

  End Select
End With
End Sub

Private Sub UnitsGrid_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
If Me.TxtModFlg.text <> "R" Then
With UnitsGrid
Select Case .ColKey(Col)
Case "Name"
If val(DcbIqara.BoundText) <> 0 And val(DcbUnitType.BoundText) <> 0 And val(DcbUnitNo.BoundText) <> 0 Then
Cancel = False
Else
Cancel = True
End If
Case "unitno"
If (.TextMatrix(row, .ColIndex("Name"))) <> "" Then
Cancel = False
Else
Cancel = True
End If
End Select
End With
End If
End Sub

Private Sub UnitsGrid_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid
        Select Case .ColKey(Col)
            Case "Name"
              .TextMatrix(row, .ColIndex("unitno")) = ""
              .TextMatrix(row, .ColIndex("UnitType")) = 0
                StrSQL = "select * from TblAkarUnit"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
   
             Case "unitno"
                StrSQL = "select * from dbo.TblAqarDetai  where id<>" & val(DcbUnitNo.BoundText) & " and  Aqarid=" & val(DcbIqara.BoundText) & " and unittype=" & val(.TextMatrix(row, .ColIndex("UnitType")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = UnitsGrid.BuildComboList(rs, "unitno", "ID")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
        End Select

    End With
End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.text <> "R" Then
              Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
      If ChekSanNumber(Current_branch, 64) = True Then
          txtNoteSerial1.text = ""
      End If
      TxtNoteSerial.text = ""
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
       If Dcbranch.text = "" Or (Me.Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Please select Branch  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
            Dcbranch.SetFocus
         End If
         End If
     If DcbIqara.text = "" Or val(Me.DcbIqara.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  «Œ Ì«— «·⁄ﬁ«—", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Real Estate  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            DcbIqara.SetFocus
            Exit Sub
         End If
          If DcbUnitType.text = "" Or val(Me.DcbUnitType.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— ‰Ê⁄ «·ÊÕœ…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Type Unit  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
           DcbUnitType.SetFocus
            Exit Sub
         End If
      If DcbUnitNo.text = "" Or val(Me.DcbUnitNo.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— —ﬁ„ «·ÊÕœ…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select  Unit  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
             DcbUnitNo.SetFocus
            Exit Sub
         End If
      If DcbTypID.text = "" Or val(Me.DcbTypID.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— ‰Ê⁄ «·⁄„·Ì…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Type ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
           DcbTypID.SetFocus
            Exit Sub
      End If
      
        If dcCustomer.text = "" Or val(Me.dcCustomer.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «·„” √Ã—", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
 
            Else
            MsgBox "Please Select  Renter  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
                    dcCustomer.SetFocus
            Exit Sub
         End If
         Dim TxtNoteSerial1str As String
my_branch = val(Me.Dcbranch.BoundText)
    If txtNoteSerial1.text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 64, 64)
    
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                  ' TxtNoteSerial1.text = TxtNoteSerial1str
                        '             txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
                    End If
                End If
    End If
'------------------------ txtmodflg type -------------------
    Select Case Me.TxtModFlg.text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           SendMessage (0)
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblOtheExpensAqar", "ID", "")
    Me.TxtSerial1.text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    txtNoteSerial1 = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 64, 64)
    RsSavRec.Fields("NoteSerial1").value = IIf(Me.txtNoteSerial1 <> "", Trim(txtNoteSerial1.text), Null)
    FiLLRec
ErrTrap:
End Sub
' change id search
Private Sub TxtSerial1_Change()
    On Error GoTo ErrTrap
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
ErrTrap:
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long, Optional NoteID As Long = 0)
    On Error GoTo ErrTrap
    If NoteID = 0 Then
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
      
    Else
      RsSavRec.Find "ID=" & NoteID, , adSearchForward, 1
    End If
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
    On Error GoTo ErrTrap
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
    FiLLTXT
     BtnLast_Click
ErrTrap:
End Sub
' delet sub
Private Sub btnDelete_Click()
    On Error GoTo ErrTrap
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    If CheckPayetTranss(val(TxtSerial1.text)) = True Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› ⁄·ÌÂ Õ—ﬂ… „ﬁ»Ê÷« "
    Exit Sub
    End If
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
                 Cn.Execute "Delete from TblOtheExpensAqarDet Where ExpensAqID=" & val(TxtSerial1.text) & ""
                 StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                RsSavRec.delete
  
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               End If
               LabCurrRec.Caption = 0
               LabCountRec.Caption = 0
               
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete this record because of its connection with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        Select Case Me.TxtModFlg.text
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
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
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

Private Sub TxtModFlg_Change()
    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
              
       
    ElseIf TxtModFlg.text = "R" Then
       
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        If TxtSerial1.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
            ISButton1.Enabled = True
    End If
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.text = "E" Then
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        Me.btnQuery.Enabled = False
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
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
              Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    
    FiLLTXT
    
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    If TxtSerial1.text <> "" Then

    If CheckPayetTranss(val(TxtSerial1.text)) = True Then
    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ⁄·ÌÂ Õ—ﬂ… „ﬁ»Ê÷« "
   ' Exit Sub
    End If
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
       UnitsGrid.rows = UnitsGrid.rows + 1
    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
          Else
           Msg = "Sorry ..." & CHR(13)
            Msg = Msg & "You can not edit this record now" & CHR(13)
            Msg = Msg & "It is in use by another user on the network"
          End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    Frm2.Enabled = True
    clear_all Me
           Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(143, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Exit Sub
                
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       Exit Sub
                End If
            End If
    TxtModFlg.text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = branch_id
    DcbStatusOper.ListIndex = 0
    DcbTypID.ListIndex = 0
    DcbTypID_Change
    UnitsGrid.Clear flexClearScrollable, flexClearEverything
    UnitsGrid.rows = 2
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
      clear_all Me
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
    FiLLTXT
    
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
              If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Public Function ReloadCombos()
Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
 Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
 Dcombos.GetIqar DcbIqara
 Dcombos.getAkarUnit Me.DcbUnitType
 Dcombos.GetSalesRepData Me.DcboEmp
End Function
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
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
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
    
  sql = " SELECT     dbo.TblOtheExpensAqar.RecordDateH, dbo.TblOtheExpensAqar.RecordDate, dbo.TblOtheExpensAqar.BranchID, dbo.TblBranchesData.branch_name, "
  sql = sql & "                    dbo.TblBranchesData.branch_namee, dbo.TblOtheExpensAqar.Valuee, dbo.TblOtheExpensAqar.Mobile, dbo.TblOtheExpensAqar.EmpID, dbo.TblEmployee.Emp_Name,"
  sql = sql & "                    dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3,"
  sql = sql & "                    dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
  sql = sql & "                    dbo.TblOtheExpensAqar.TypID, dbo.TblOtheExpensAqar.BillNo, dbo.TblOtheExpensAqar.AccountNo, dbo.TblOtheExpensAqar.AccountBank,"
  sql = sql & "                    dbo.TblOtheExpensAqar.Remarks, dbo.TblOtheExpensAqar.PayedDateH, dbo.TblOtheExpensAqar.PayedDate, dbo.TblOtheExpensAqar.FromDateH,"
  sql = sql & "                    dbo.TblOtheExpensAqar.FromDate, dbo.TblOtheExpensAqar.ToDateH, dbo.TblOtheExpensAqar.ToDate, dbo.TblOtheExpensAqar.StatusOper,"
  sql = sql & "                    dbo.TblOtheExpensAqar.RemainRent, dbo.TblOtheExpensAqar.Electricity, dbo.TblOtheExpensAqar.Maintenance, dbo.TblOtheExpensAqar.MaintCondition,"
  sql = sql & "                    dbo.TblOtheExpensAqar.MaintDoors, dbo.TblOtheExpensAqar.MaintKitchen, dbo.TblOtheExpensAqar.MaintClean, dbo.TblOtheExpensAqar.MaintOther,"
  sql = sql & "                    dbo.TblOtheExpensAqar.Insurance, dbo.TblOtheExpensAqar.DelayDay, dbo.TblOtheExpensAqar.Noliquidation, dbo.TblOtheExpensAqar.Paints,"
  sql = sql & "                    dbo.TblOtheExpensAqar.Windows, dbo.TblOtheExpensAqar.Total, dbo.TblOtheExpensAqar.Net, dbo.TblOtheExpensAqar.Name AS NameAccount,"
  sql = sql & "                    dbo.TblOtheExpensAqar.TotalAfterIns, dbo.TblOtheExpensAqar.Discount, dbo.TblOtheExpensAqar.BankID, dbo.BanksData.BankName, dbo.BanksData.BankNamee,"
  sql = sql & "                    dbo.TblOtheExpensAqar.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode,"
  sql = sql & "                    dbo.TblOtheExpensAqar.UnitID, TblAqarDetai_1.unitno, dbo.TblOtheExpensAqar.UnitTypID, TblAkarUnit_1.name, TblAkarUnit_1.namee,"
  sql = sql & "                    dbo.TblOtheExpensAqar.Aqarid , dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblOtheExpensAqar.ID"
  sql = sql & "     FROM         dbo.TblAqar RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblOtheExpensAqar ON dbo.TblAqar.Aqarid = dbo.TblOtheExpensAqar.AqarID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblAkarUnit TblAkarUnit_1 ON dbo.TblOtheExpensAqar.UnitTypID = TblAkarUnit_1.id LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblAqarDetai TblAqarDetai_1 ON dbo.TblOtheExpensAqar.UnitID = TblAqarDetai_1.Id LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblCustemers ON dbo.TblOtheExpensAqar.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  sql = sql & "                    dbo.BanksData ON dbo.TblOtheExpensAqar.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblEmployee ON dbo.TblOtheExpensAqar.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBranchesData ON dbo.TblOtheExpensAqar.BranchID = dbo.TblBranchesData.branch_id"
  sql = sql & "  Where (dbo.TblOtheExpensAqar.ID = " & val(TxtSerial1.text) & ")"
    
    If val(DcbTypID.ListIndex) = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOtheExpensAqar.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOtheExpensAqar.rpt"
        End If
    Else
           If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOtheExpensAqarElectr.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOtheExpensAqarElectr.rpt"
        End If
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
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If val(DcbTypID.ListIndex) <> 0 Then
     xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtValuee.text), "0.00"), 0, True, ".")
     Else
     xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(Me.TxtNet.text), "0.00"), 0, True, ".")
    End If
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Proof of  Insurances"
    ' labell name
    Label1(35).Caption = "GL"
    Command9.Caption = "Print  GL"
    Command8.Caption = "Account"
     
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblOtheExpensAqar"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub

'+++++++++++++++++++++++++++++++++ en
