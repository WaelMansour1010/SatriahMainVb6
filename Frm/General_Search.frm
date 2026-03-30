VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form General_Search 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9090
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   12615
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12615
   Begin VB.Frame Frame24 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   294
      Top             =   600
      Width           =   12612
      Begin VB.ComboBox CboCalType1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   324
         Top             =   5640
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Frame Frame26 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·Ã“«¡"
         Height          =   2055
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   311
         Top             =   3360
         Width           =   5535
         Begin VB.TextBox Text28 
            Alignment       =   1  'Right Justify
            Height          =   1080
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   312
            Top             =   840
            Width           =   4275
         End
         Begin MSDataListLib.DataCombo DcbSanction1 
            Height          =   315
            Left            =   240
            TabIndex        =   313
            Top             =   360
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «·Ã“«¡"
            Height          =   330
            Index           =   53
            Left            =   4485
            TabIndex        =   315
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·Ã“«¡"
            Height          =   285
            Index           =   52
            Left            =   4440
            TabIndex        =   314
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame Frame25 
         Height          =   852
         Left            =   480
         TabIndex        =   309
         Top             =   7200
         Width           =   11772
         Begin ALLButtonS.ALLButton ALLButton24 
            Height          =   492
            Left            =   120
            TabIndex        =   310
            Top             =   240
            Width           =   11532
            _ExtentX        =   20346
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Œ—ÊÃ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0000
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·⁄„·Ì…"
         Height          =   645
         Index           =   14
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   304
         Top             =   3360
         Width           =   4455
         Begin MSComCtl2.DTPicker DTPicker9 
            Height          =   330
            Left            =   1890
            TabIndex        =   305
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker10 
            Height          =   330
            Left            =   90
            TabIndex        =   306
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   51
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   308
            Top             =   180
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   50
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   307
            Top             =   210
            Width           =   420
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   645
         Index           =   13
         Left            =   9930
         RightToLeft     =   -1  'True
         TabIndex        =   299
         Top             =   3360
         Width           =   2715
         Begin VB.TextBox Text27 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   301
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox Text26 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   300
            Top             =   180
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   43
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   303
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   42
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   302
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   1395
         Index           =   12
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   295
         Top             =   3960
         Width           =   7095
         Begin VB.ComboBox CboCalType 
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   320
            Top             =   960
            Width           =   5175
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   296
            Top             =   240
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DcbEmp12 
            Height          =   315
            Left            =   360
            TabIndex        =   297
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCComponent 
            Height          =   315
            Left            =   360
            TabIndex        =   323
            Top             =   600
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„›—œ «·Œ’„"
            Height          =   255
            Index           =   7
            Left            =   6045
            RightToLeft     =   -1  'True
            TabIndex        =   322
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   255
            Index           =   5
            Left            =   6045
            RightToLeft     =   -1  'True
            TabIndex        =   321
            Top             =   990
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   285
            Index           =   39
            Left            =   5910
            TabIndex        =   298
            Top             =   240
            Width           =   1005
         End
      End
      Begin ALLButtonS.ALLButton ALLButton25 
         Height          =   375
         Left            =   5520
         TabIndex        =   316
         Top             =   6600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid8 
         Height          =   2625
         Left            =   0
         TabIndex        =   317
         Top             =   600
         Width           =   12555
         _cx             =   22146
         _cy             =   4630
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"General_Search.frx":0038
         ScrollTrack     =   -1  'True
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
      Begin ALLButtonS.ALLButton ALLButton26 
         Height          =   375
         Left            =   720
         TabIndex        =   318
         Top             =   6600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "„”Õ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":01AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ  «·Œ’Ê„«  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   319
         Top             =   120
         Width           =   12615
      End
   End
   Begin VB.Frame Frame20 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   263
      Top             =   600
      Width           =   12612
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   1395
         Index           =   15
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   281
         Top             =   3960
         Width           =   7095
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   283
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtSearchCode2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   282
            Top             =   600
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DcbEmp10 
            Height          =   315
            Left            =   360
            TabIndex        =   284
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEmp11 
            Height          =   315
            Left            =   360
            TabIndex        =   285
            Top             =   600
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDept 
            Height          =   315
            Left            =   360
            TabIndex        =   286
            Top             =   960
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   285
            Index           =   44
            Left            =   5910
            TabIndex        =   289
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ«∆„ »«·«‰–«—"
            Height          =   285
            Index           =   56
            Left            =   5670
            TabIndex        =   288
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ”„"
            Height          =   285
            Index           =   54
            Left            =   6120
            TabIndex        =   287
            Top             =   960
            Width           =   765
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   645
         Index           =   16
         Left            =   9930
         RightToLeft     =   -1  'True
         TabIndex        =   276
         Top             =   3360
         Width           =   2715
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   278
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox Text23 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   277
            Top             =   180
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   45
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   280
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   46
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   279
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·⁄„·Ì…"
         Height          =   645
         Index           =   17
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   271
         Top             =   3360
         Width           =   4455
         Begin MSComCtl2.DTPicker DTPicker15 
            Height          =   330
            Left            =   1890
            TabIndex        =   272
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker16 
            Height          =   330
            Left            =   90
            TabIndex        =   273
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   47
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   275
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   48
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   274
            Top             =   180
            Width           =   255
         End
      End
      Begin VB.Frame Frame23 
         Height          =   852
         Left            =   480
         TabIndex        =   269
         Top             =   7200
         Width           =   11772
         Begin ALLButtonS.ALLButton ALLButton21 
            Height          =   492
            Left            =   120
            TabIndex        =   270
            Top             =   240
            Width           =   11532
            _ExtentX        =   20346
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Œ—ÊÃ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":01CB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·Ã“«¡"
         Height          =   2055
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   264
         Top             =   3360
         Width           =   5535
         Begin VB.TextBox txtRemark 
            Alignment       =   1  'Right Justify
            Height          =   1080
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   265
            Top             =   840
            Width           =   4275
         End
         Begin MSDataListLib.DataCombo DcbSanction 
            Height          =   315
            Left            =   240
            TabIndex        =   266
            Top             =   360
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·Ã“«¡"
            Height          =   285
            Index           =   49
            Left            =   4440
            TabIndex        =   268
            Top             =   360
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «·Ã“«¡"
            Height          =   330
            Index           =   31
            Left            =   4485
            TabIndex        =   267
            Top             =   1200
            Width           =   840
         End
      End
      Begin ALLButtonS.ALLButton ALLButton22 
         Height          =   375
         Left            =   5520
         TabIndex        =   290
         Top             =   6600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":01E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid7 
         Height          =   2625
         Left            =   0
         TabIndex        =   291
         Top             =   720
         Width           =   12555
         _cx             =   22146
         _cy             =   4630
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"General_Search.frx":0203
         ScrollTrack     =   -1  'True
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
      Begin ALLButtonS.ALLButton ALLButton23 
         Height          =   375
         Left            =   720
         TabIndex        =   292
         Top             =   6600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "„”Õ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":0387
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ  «‰–«— „ÊŸ›"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   293
         Top             =   120
         Width           =   12615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   217
      Top             =   600
      Width           =   12612
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ ‰Â«Ì…«·«Ã«“…"
         Height          =   1035
         Index           =   11
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   255
         Top             =   5520
         Width           =   5295
         Begin MSComCtl2.DTPicker DTPicker7 
            Height          =   330
            Left            =   2250
            TabIndex        =   256
            Top             =   270
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   330
            Left            =   2250
            TabIndex        =   257
            Top             =   630
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
            Height          =   315
            Left            =   120
            TabIndex        =   258
            Top             =   270
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
            Height          =   315
            Left            =   120
            TabIndex        =   259
            Top             =   630
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   37
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   261
            Top             =   330
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   38
            Left            =   4785
            RightToLeft     =   -1  'True
            TabIndex        =   260
            Top             =   660
            Width           =   255
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·„»«‘—…"
         Height          =   1035
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   248
         Top             =   3360
         Width           =   5295
         Begin MSComCtl2.DTPicker FromStartDate 
            Height          =   330
            Left            =   2280
            TabIndex        =   249
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker ToStartDate 
            Height          =   330
            Left            =   2280
            TabIndex        =   250
            Top             =   630
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal FromStartDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   251
            Top             =   270
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal ToStartDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   252
            Top             =   630
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   36
            Left            =   4815
            RightToLeft     =   -1  'True
            TabIndex        =   254
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   35
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   253
            Top             =   330
            Width           =   420
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ »œ«Ì… «·«Ã«“…"
         Height          =   1035
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   241
         Top             =   4440
         Width           =   5295
         Begin MSComCtl2.DTPicker FromEndDate 
            Height          =   330
            Left            =   2250
            TabIndex        =   242
            Top             =   270
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker ToEndDate 
            Height          =   330
            Left            =   2250
            TabIndex        =   243
            Top             =   630
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal FromEndDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   244
            Top             =   270
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal ToEndDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   245
            Top             =   630
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   34
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   247
            Top             =   330
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   33
            Left            =   4785
            RightToLeft     =   -1  'True
            TabIndex        =   246
            Top             =   660
            Width           =   255
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   1515
         Index           =   8
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   233
         Top             =   3960
         Width           =   7095
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   238
            Top             =   240
            Width           =   1335
         End
         Begin VB.Frame Frame21 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„»«‘—…"
            Height          =   735
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   234
            Top             =   600
            Width           =   6615
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   495
               Index           =   1
               Left            =   2040
               TabIndex        =   235
               Top             =   120
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "≈· Õ«ﬁ „ÊŸ› ÃœÌœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   495
               Index           =   2
               Left            =   120
               TabIndex        =   236
               Top             =   120
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "⁄Êœ… „‰ ≈Ã«“…"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   495
               Index           =   0
               Left            =   3840
               TabIndex        =   237
               Top             =   120
               Width           =   1095
               _Version        =   786432
               _ExtentX        =   1931
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "«·ﬂ·"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin MSDataListLib.DataCombo DcbEmp 
            Height          =   315
            Left            =   360
            TabIndex        =   239
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   285
            Index           =   32
            Left            =   5910
            TabIndex        =   240
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   645
         Index           =   7
         Left            =   9930
         RightToLeft     =   -1  'True
         TabIndex        =   228
         Top             =   3360
         Width           =   2715
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   230
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   229
            Top             =   180
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   30
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   29
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   231
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·⁄„·Ì…"
         Height          =   645
         Index           =   6
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   223
         Top             =   3360
         Width           =   4455
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   330
            Left            =   1890
            TabIndex        =   224
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   330
            Left            =   90
            TabIndex        =   225
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   28
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   227
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   27
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   226
            Top             =   180
            Width           =   255
         End
      End
      Begin VB.Frame Frame19 
         Height          =   852
         Left            =   480
         TabIndex        =   218
         Top             =   7200
         Width           =   11772
         Begin ALLButtonS.ALLButton ALLButton18 
            Height          =   492
            Left            =   120
            TabIndex        =   219
            Top             =   240
            Width           =   11532
            _ExtentX        =   20346
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Œ—ÊÃ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":03A3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin ALLButtonS.ALLButton ALLButton19 
         Height          =   375
         Left            =   720
         TabIndex        =   220
         Top             =   6000
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":03BF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid6 
         Height          =   2625
         Left            =   0
         TabIndex        =   222
         Top             =   720
         Width           =   12555
         _cx             =   22146
         _cy             =   4630
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
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"General_Search.frx":03DB
         ScrollTrack     =   -1  'True
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
      Begin ALLButtonS.ALLButton ALLButton20 
         Height          =   375
         Left            =   720
         TabIndex        =   262
         Top             =   6600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "„”Õ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":061E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ „»«‘—… ⁄„·"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   221
         Top             =   120
         Width           =   12615
      End
   End
   Begin VB.Frame fram_empadv 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   197
      Top             =   -120
      Width           =   12612
      Begin VB.Frame Frame18 
         Height          =   852
         Left            =   480
         TabIndex        =   214
         Top             =   7200
         Width           =   11772
         Begin ALLButtonS.ALLButton ALLButton16 
            Height          =   492
            Left            =   120
            TabIndex        =   215
            Top             =   240
            Width           =   11532
            _ExtentX        =   20346
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":063A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Fra 
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   645
         Index           =   5
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   204
         Top             =   5760
         Width           =   3795
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "≈·Ï"
            Height          =   192
            Index           =   24
            Left            =   1140
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   240
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   192
            Index           =   23
            Left            =   2532
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   672
         Index           =   4
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   199
         Top             =   5880
         Width           =   4452
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   336
            Left            =   2400
            TabIndex        =   200
            Top             =   240
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   336
            Left            =   240
            TabIndex        =   201
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   312
            Index           =   22
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "≈·Ï"
            Height          =   192
            Index           =   20
            Left            =   1812
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   300
            Width           =   492
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
         Height          =   4785
         Left            =   0
         TabIndex        =   209
         Top             =   840
         Width           =   12555
         _cx             =   22140
         _cy             =   8446
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
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"General_Search.frx":0656
         ScrollTrack     =   -1  'True
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
      Begin MSDataListLib.DataCombo DataCombo8 
         Height          =   288
         Left            =   8316
         TabIndex        =   210
         Top             =   6480
         Width           =   2892
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo9 
         CausesValidation=   0   'False
         Height          =   288
         Left            =   8316
         TabIndex        =   211
         Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
         Top             =   6840
         Width           =   2892
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
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
      Begin ALLButtonS.ALLButton ALLButton17 
         Height          =   372
         Left            =   720
         TabIndex        =   216
         Top             =   6720
         Width           =   4212
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":0781
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„” Œœ„"
         Height          =   288
         Index           =   26
         Left            =   11220
         RightToLeft     =   -1  'True
         TabIndex        =   213
         Top             =   6876
         Width           =   948
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   312
         Index           =   25
         Left            =   11196
         RightToLeft     =   -1  'True
         TabIndex        =   212
         Top             =   6480
         Width           =   972
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ ”·› „ÊŸ›Ì‰            "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   198
         Top             =   120
         Width           =   12615
      End
   End
   Begin VB.Frame Fram_Visa 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   166
      Top             =   -120
      Width           =   12372
      Begin VB.Frame Frame16 
         Height          =   732
         Left            =   120
         TabIndex        =   194
         Top             =   7320
         Width           =   12012
         Begin ALLButtonS.ALLButton ALLButton15 
            Height          =   492
            Left            =   240
            TabIndex        =   195
            Top             =   120
            Width           =   11652
            _ExtentX        =   20558
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":079D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   1035
         Index           =   3
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   177
         Top             =   5640
         Width           =   4212
         Begin MSComCtl2.DTPicker EndDateFrom 
            Height          =   330
            Left            =   2370
            TabIndex        =   178
            Top             =   270
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker EndDateTo 
            Height          =   330
            Left            =   2370
            TabIndex        =   179
            Top             =   630
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal EndDateFromH 
            Height          =   315
            Left            =   720
            TabIndex        =   180
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal EndDateToH 
            Height          =   315
            Left            =   720
            TabIndex        =   181
            Top             =   630
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   192
            Index           =   14
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   192
            Index           =   13
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   660
            Width           =   492
         End
      End
      Begin VB.TextBox TxtHodono 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   176
         Top             =   5400
         Width           =   4455
      End
      Begin VB.TextBox TxtVisa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   175
         Top             =   5040
         Width           =   4455
      End
      Begin VB.TextBox TxtOrder 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   174
         Top             =   4680
         Width           =   4455
      End
      Begin VB.Frame Fra 
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   1035
         Index           =   0
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   167
         Top             =   6000
         Width           =   4335
         Begin MSComCtl2.DTPicker StarDateFrom 
            Height          =   330
            Left            =   2370
            TabIndex        =   168
            Top             =   270
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   42234
         End
         Begin MSComCtl2.DTPicker StarDateTo 
            Height          =   330
            Left            =   2370
            TabIndex        =   169
            Top             =   630
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            UpDown          =   -1  'True
            CurrentDate     =   42234
         End
         Begin Dynamic_Byte.NourHijriCal StarDateFromH 
            Height          =   315
            Left            =   720
            TabIndex        =   170
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal StarDateToH 
            Height          =   315
            Left            =   720
            TabIndex        =   171
            Top             =   630
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   10
            Left            =   3735
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   9
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   330
            Width           =   660
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg2 
         Height          =   3708
         Left            =   0
         TabIndex        =   184
         Top             =   720
         Width           =   12312
         _cx             =   21717
         _cy             =   6540
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483633
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
         Rows            =   50
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"General_Search.frx":07B9
         ScrollTrack     =   -1  'True
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
      Begin MSDataListLib.DataCombo DcbNtionality 
         Height          =   288
         Left            =   360
         TabIndex        =   185
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   4680
         Width           =   4212
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCity 
         Height          =   288
         Left            =   360
         TabIndex        =   186
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3960
         Width           =   4212
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ALLButtonS.ALLButton ALLButton13 
         Height          =   372
         Left            =   360
         TabIndex        =   193
         Top             =   6840
         Width           =   4212
         _ExtentX        =   7435
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "General_Search.frx":09F5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ «· √‘Ì—«             "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -8280
         TabIndex        =   196
         Top             =   120
         Width           =   20652
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   21
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   192
         Top             =   2820
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   192
         Index           =   19
         Left            =   11112
         RightToLeft     =   -1  'True
         TabIndex        =   191
         Top             =   4680
         Width           =   648
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «· √‘Ì—…"
         Height          =   192
         Index           =   18
         Left            =   10932
         RightToLeft     =   -1  'True
         TabIndex        =   190
         Top             =   5040
         Width           =   876
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·ÕœÊœ"
         Height          =   192
         Index           =   17
         Left            =   11088
         RightToLeft     =   -1  'True
         TabIndex        =   189
         Top             =   5400
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Ã‰”Ì…"
         Height          =   192
         Index           =   16
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   188
         Top             =   4680
         Width           =   528
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„œÌ‰…"
         Height          =   192
         Index           =   15
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   187
         Top             =   5040
         Visible         =   0   'False
         Width           =   432
      End
   End
   Begin VB.Frame Fram_Advreq 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   145
      Top             =   -120
      Width           =   12612
      Begin VB.Frame Fra 
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   645
         Index           =   2
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   154
         Top             =   5520
         Width           =   3795
         Begin VB.TextBox TxtIDTO 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox TxtIDFrom 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   8
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   195
            Index           =   7
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   240
            Width           =   300
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   1035
         Index           =   1
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   5400
         Width           =   2532
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   90
            TabIndex        =   150
            Top             =   270
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   90
            TabIndex        =   151
            Top             =   630
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   427032579
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   195
            Index           =   6
            Left            =   1740
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   5
            Left            =   1695
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   660
            Width           =   255
         End
      End
      Begin VB.Frame Frame17 
         Height          =   732
         Left            =   120
         TabIndex        =   146
         Top             =   6960
         Width           =   12375
         Begin ALLButtonS.ALLButton ALLButton14 
            Height          =   492
            Left            =   240
            TabIndex        =   147
            Top             =   120
            Width           =   12012
            _ExtentX        =   21193
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0A11
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   4548
         Left            =   120
         TabIndex        =   159
         Top             =   840
         Width           =   12312
         _cx             =   21717
         _cy             =   8022
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483633
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
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"General_Search.frx":0A2D
         ScrollTrack     =   -1  'True
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
      Begin MSDataListLib.DataCombo DCEmp_Name 
         Height          =   312
         Left            =   8400
         TabIndex        =   160
         Top             =   6480
         Width           =   2772
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   372
         Index           =   0
         Left            =   480
         TabIndex        =   161
         Top             =   6480
         Width           =   1968
         _ExtentX        =   3466
         _ExtentY        =   661
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
         BackStyle       =   0
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin MSDataListLib.DataCombo DCUser 
         CausesValidation=   0   'False
         Height          =   288
         Left            =   3840
         TabIndex        =   162
         Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
         Top             =   6600
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label returntype 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   372
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   165
         Top             =   4656
         Width           =   1332
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„” Œœ„"
         Height          =   288
         Index           =   12
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   164
         Top             =   6600
         Width           =   948
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   552
         Index           =   11
         Left            =   11040
         RightToLeft     =   -1  'True
         TabIndex        =   163
         Top             =   6480
         Width           =   972
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ ÿ·» ”·›… ‰ﬁœÌ…            "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -8040
         TabIndex        =   148
         Top             =   120
         Width           =   20652
      End
   End
   Begin VB.Frame Fram_treat 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   -120
      Width           =   12612
      Begin VB.Frame Frame15 
         Height          =   2292
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   4560
         Width           =   12372
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   960
            Width           =   4152
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   1320
            Width           =   4152
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   1080
            Width           =   4512
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   126
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Height          =   360
            Left            =   9360
            TabIndex        =   125
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   14280
            TabIndex        =   124
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   312
            Left            =   1920
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   600
            Width           =   2748
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   427032579
            CurrentDate     =   37140
         End
         Begin ALLButtonS.ALLButton ALLButton12 
            Height          =   372
            Left            =   480
            TabIndex        =   129
            Top             =   1680
            Width           =   4092
            _ExtentX        =   7223
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "»ÕÀ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0B59
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Height          =   288
            Left            =   6120
            TabIndex        =   130
            Top             =   720
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Height          =   288
            Left            =   6120
            TabIndex        =   131
            Top             =   1440
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "General_Search.frx":0B75
            Height          =   288
            Left            =   6120
            TabIndex        =   132
            Top             =   1800
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·Â« ›"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   4200
            TabIndex        =   143
            Top             =   1320
            Width           =   1572
         End
         Begin VB.Label lblbranch 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   252
            Index           =   4
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1800
            Width           =   852
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   720
            TabIndex        =   139
            Top             =   3720
            Width           =   1092
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·«ﬁ«„…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   138
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ÿ·»"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   137
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·„ÊŸ›"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   136
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ÃÊ«“"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   4200
            TabIndex        =   135
            Top             =   960
            Width           =   1572
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ÊŸÌ›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   134
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ "
            Height          =   300
            Index           =   4
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   600
            Width           =   1320
         End
      End
      Begin VB.Frame Frame14 
         Height          =   732
         Left            =   120
         TabIndex        =   121
         Top             =   6960
         Width           =   12375
         Begin ALLButtonS.ALLButton ALLButton11 
            Height          =   492
            Left            =   240
            TabIndex        =   122
            Top             =   120
            Width           =   12012
            _ExtentX        =   21193
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0B8A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   720
         Width           =   12372
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
            Height          =   3540
            Left            =   120
            TabIndex        =   119
            Top             =   120
            Width           =   12240
            _cx             =   21590
            _cy             =   6244
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
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"General_Search.frx":0BA6
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
         Begin ALLButtonS.ALLButton ALLButton10 
            Height          =   375
            Left            =   120
            TabIndex        =   120
            Tag             =   "Delete Row"
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "General_Search.frx":0CC5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ «·Ï „‰ ÌÂ„… «·«„—            "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -8040
         TabIndex        =   141
         Top             =   120
         Width           =   20652
      End
   End
   Begin VB.Frame Fram_adv 
      Height          =   8292
      Left            =   -240
      RightToLeft     =   -1  'True
      TabIndex        =   91
      Top             =   -120
      Width           =   12612
      Begin VB.Frame Frame12 
         Height          =   3855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   112
         Top             =   720
         Width           =   12372
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
            Height          =   3540
            Left            =   120
            TabIndex        =   113
            Top             =   120
            Width           =   12240
            _cx             =   21590
            _cy             =   6244
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
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"General_Search.frx":0CE1
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
         Begin ALLButtonS.ALLButton ALLButton9 
            Height          =   375
            Left            =   120
            TabIndex        =   114
            Tag             =   "Delete Row"
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "General_Search.frx":0E38
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame11 
         Height          =   732
         Left            =   120
         TabIndex        =   110
         Top             =   6960
         Width           =   12375
         Begin ALLButtonS.ALLButton ALLButton8 
            Height          =   492
            Left            =   240
            TabIndex        =   111
            Top             =   120
            Width           =   12012
            _ExtentX        =   21193
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0E54
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame10 
         Height          =   2292
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   4560
         Width           =   12372
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   14280
            TabIndex        =   96
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Height          =   360
            Left            =   9360
            TabIndex        =   95
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   94
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   1080
            Width           =   4512
         End
         Begin MSComCtl2.DTPicker lastPayDateFrom 
            Height          =   312
            Left            =   960
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   600
            Width           =   2748
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   427032579
            CurrentDate     =   37140
         End
         Begin ALLButtonS.ALLButton ALLButton7 
            Height          =   372
            Left            =   960
            TabIndex        =   98
            Top             =   1680
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "»ÕÀ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0E70
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   288
            Left            =   6120
            TabIndex        =   99
            Top             =   720
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   288
            Left            =   6120
            TabIndex        =   100
            Top             =   1440
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "General_Search.frx":0E8C
            Height          =   288
            Left            =   6120
            TabIndex        =   101
            Top             =   1800
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSComCtl2.DTPicker currentdayFrom 
            Height          =   312
            Left            =   960
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2748
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   427032579
            CurrentDate     =   37140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ ’—› «Œ— »œ·"
            Height          =   420
            Index           =   3
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   600
            Width           =   1320
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ÊŸÌ›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   108
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "«·› —… «·Õ«·Ì… «·„” Õﬁ…"
            ForeColor       =   &H00000000&
            Height          =   492
            Left            =   3840
            TabIndex        =   107
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·„ÊŸ›"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   106
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ÿ·»"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   105
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„” Õﬁ ’—›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   104
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   720
            TabIndex        =   103
            Top             =   3720
            Width           =   1092
         End
         Begin VB.Label lblbranch 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   252
            Index           =   3
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1800
            Width           =   852
         End
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ ÿ·» ’—› »œ· ”ﬂ‰ „ﬁœ„            "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -8040
         TabIndex        =   115
         Top             =   120
         Width           =   20652
      End
   End
   Begin VB.Frame Fram_Passports 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   -120
      Width           =   12612
      Begin VB.Frame Frame9 
         Height          =   2652
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   4560
         Width           =   12372
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   600
            Width           =   4152
         End
         Begin VB.TextBox numbEkama 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1440
            Width           =   4512
         End
         Begin VB.TextBox Nationality3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1080
            Width           =   4512
         End
         Begin VB.TextBox remark 
            Alignment       =   1  'Right Justify
            Height          =   444
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   1320
            Width           =   4212
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   70
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Height          =   360
            Left            =   9360
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   14280
            TabIndex        =   68
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComCtl2.DTPicker recorddate 
            Height          =   312
            Left            =   2760
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   960
            Width           =   1548
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   427032579
            CurrentDate     =   37140
         End
         Begin ALLButtonS.ALLButton ALLButton6 
            Height          =   372
            Left            =   120
            TabIndex        =   73
            Top             =   1920
            Width           =   4212
            _ExtentX        =   7435
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "»ÕÀ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0EA1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   288
            Left            =   6120
            TabIndex        =   74
            Top             =   720
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo JobTypeName3 
            Height          =   288
            Left            =   6000
            TabIndex        =   75
            Top             =   2400
            Visible         =   0   'False
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo branch_no 
            Bindings        =   "General_Search.frx":0EBD
            Height          =   288
            Left            =   6120
            TabIndex        =   76
            Top             =   1800
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label lblbranch 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   252
            Index           =   2
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1920
            Width           =   852
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   720
            TabIndex        =   85
            Top             =   3720
            Width           =   1092
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ã‰”Ì…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   84
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   83
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·„ÊŸ›"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   82
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ÃÊ«“"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   720
            Width           =   1092
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "«·€—÷"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   4440
            TabIndex        =   80
            Top             =   1440
            Width           =   1212
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·«ﬁ«„…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   79
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ÊŸÌ›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   78
            Top             =   2400
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "«· «—ÌŒ "
            Height          =   300
            Index           =   2
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1080
            Width           =   1080
         End
      End
      Begin VB.Frame Frame8 
         Height          =   852
         Left            =   120
         TabIndex        =   65
         Top             =   7320
         Width           =   12375
         Begin ALLButtonS.ALLButton ALLButton5 
            Height          =   492
            Left            =   240
            TabIndex        =   66
            Top             =   240
            Width           =   12012
            _ExtentX        =   21193
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":0ED2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   720
         Width           =   12372
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   3660
            Left            =   0
            TabIndex        =   63
            Top             =   120
            Width           =   12240
            _cx             =   21590
            _cy             =   6456
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
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"General_Search.frx":0EEE
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
         Begin ALLButtonS.ALLButton ALLButton4 
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Tag             =   "Delete Row"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "General_Search.frx":1007
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ  ”·Ì„ ÃÊ«“ ”›— ·„ÊŸ›                "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -7920
         TabIndex        =   87
         Top             =   120
         Width           =   20652
      End
   End
   Begin VB.Frame Fram_businessjob 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   -120
      Width           =   12612
      Begin VB.Frame Frame5 
         Height          =   3855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   720
         Width           =   12372
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   3540
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   12240
            _cx             =   21590
            _cy             =   6244
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
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"General_Search.frx":1023
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
         Begin ALLButtonS.ALLButton ALLButton3 
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Tag             =   "Delete Row"
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "General_Search.frx":1166
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame4 
         Height          =   732
         Left            =   120
         TabIndex        =   55
         Top             =   7320
         Width           =   12375
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   492
            Left            =   240
            TabIndex        =   56
            Top             =   120
            Width           =   12012
            _ExtentX        =   21193
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":1182
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2772
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   4560
         Width           =   12372
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   14280
            TabIndex        =   36
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox advanceID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Height          =   360
            Left            =   9360
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   34
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox task 
            Alignment       =   1  'Right Justify
            Height          =   684
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   1440
            Width           =   4212
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   312
            Left            =   2760
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   960
            Width           =   1548
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   427032579
            CurrentDate     =   37140
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   372
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   4212
            _ExtentX        =   7435
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "»ÕÀ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":119E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo emp_Name 
            Height          =   288
            Left            =   6120
            TabIndex        =   39
            Top             =   720
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Nationality 
            Height          =   288
            Left            =   6120
            TabIndex        =   40
            Top             =   1080
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Location 
            Height          =   288
            Left            =   6120
            TabIndex        =   41
            Top             =   1440
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo JobTypeName 
            Height          =   288
            Left            =   6120
            TabIndex        =   42
            Top             =   1800
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo manager 
            Height          =   288
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   4152
            _ExtentX        =   7329
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo branch_name 
            Bindings        =   "General_Search.frx":11BA
            Height          =   288
            Left            =   6120
            TabIndex        =   44
            Top             =   2160
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «·„Â„…"
            Height          =   300
            Index           =   1
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ÊŸÌ›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   53
            Top             =   1800
            Width           =   1572
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„Êﬁ⁄"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   52
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "„Â„… «·⁄„·"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   4200
            TabIndex        =   51
            Top             =   1440
            Width           =   1212
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„œÌ—"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·„ÊŸ›"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   49
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "«·—ﬁ„ «·ÿ·»"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   48
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ã‰”Ì…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   47
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   0
            Left            =   720
            TabIndex        =   46
            Top             =   3720
            Width           =   1092
         End
         Begin VB.Label lblbranch 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   252
            Index           =   1
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2220
            Width           =   852
         End
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ „Â„… ⁄„·                "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -8040
         TabIndex        =   60
         Top             =   120
         Width           =   20652
      End
   End
   Begin VB.Frame fram_empMove 
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -120
      Width           =   12612
      Begin VB.Frame fram1 
         Height          =   2772
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   4560
         Width           =   12372
         Begin VB.TextBox txtreson 
            Alignment       =   1  'Right Justify
            Height          =   684
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1440
            Width           =   4212
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   13
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox txtid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Height          =   360
            Left            =   9360
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TXTEnd_user_id 
            Height          =   285
            Left            =   14280
            TabIndex        =   11
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   312
            Left            =   2760
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   960
            Width           =   1548
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   427032579
            CurrentDate     =   37140
         End
         Begin ALLButtonS.ALLButton btnSearch 
            Height          =   372
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   4092
            _ExtentX        =   7223
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "»ÕÀ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":11CF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   288
            Left            =   6120
            TabIndex        =   17
            Top             =   720
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcmbFromDepart 
            Height          =   288
            Left            =   6120
            TabIndex        =   18
            Top             =   1080
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcmbFromProject 
            Height          =   288
            Left            =   6120
            TabIndex        =   19
            Top             =   1440
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboJobsType 
            Height          =   288
            Left            =   6120
            TabIndex        =   20
            Top             =   1800
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcmbManagerID 
            Height          =   288
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   4152
            _ExtentX        =   7329
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "General_Search.frx":11EB
            Height          =   288
            Left            =   6120
            TabIndex        =   8
            Top             =   2160
            Width           =   4512
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label lblbranch 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   252
            Index           =   0
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   2220
            Width           =   852
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   720
            TabIndex        =   30
            Top             =   3720
            Width           =   1092
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬁ”„"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   29
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "«·—ﬁ„ «·ÿ·»"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·„ÊŸ›"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10680
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„œÌ—"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "”»» «·‰ﬁ·"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   4200
            TabIndex        =   25
            Top             =   1440
            Width           =   1212
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„Êﬁ⁄"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   24
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ÊŸÌ›…"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   10680
            TabIndex        =   23
            Top             =   1800
            Width           =   1572
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «·‰ﬁ·"
            Height          =   300
            Index           =   0
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1080
            Width           =   1080
         End
      End
      Begin VB.Frame Frame7 
         Height          =   732
         Left            =   120
         TabIndex        =   6
         Top             =   7320
         Width           =   12375
         Begin ALLButtonS.ALLButton btnOk 
            Height          =   492
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   12012
            _ExtentX        =   21193
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "„Ê«›ﬁ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "General_Search.frx":1200
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   12372
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   3540
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   12240
            _cx             =   21590
            _cy             =   6244
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
            Rows            =   3
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"General_Search.frx":121C
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Tag             =   "Delete Row"
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "General_Search.frx":137A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "»ÕÀ ÿ·» ‰ﬁ· „ÊŸ›                "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   612
         Left            =   -8040
         TabIndex        =   2
         Top             =   120
         Width           =   20652
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Language  «··€…"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "EN"
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
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "General_Search.frx":1396
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "General_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Long
Dim last_root As Integer
Dim last_geeral As Integer
Dim last_branch As Integer
Dim mod_flad As String
Dim first_run  As Boolean
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset
Dim current_terms As String
Dim current_opr As String
Dim NewGrid As New ClsGrid
Dim expanses_account As String
Public Index As Integer

Public send_form As String



Private Sub ALLButton1_Click()

   Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 Dim MySQL As String
  
MySQL = MySQL & "  SELECT dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee,"
     MySQL = MySQL & "               dbo.TblUsers.UserName, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpDepartments.DepartmentName,"
      MySQL = MySQL & "              dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpJobsTypes.JobTypeID, dbo.TblEmpJobOrder.Branch_NO,"
      MySQL = MySQL & "              dbo.TblEmpGrades.name, dbo.TblEmpGrades.namee, dbo.TblEmpJobOrder.AdvanceID, dbo.TblEmpJobOrder.interval, dbo.TblEmpJobOrder.AdvanceDate,"
       MySQL = MySQL & "             dbo.TblEmpJobOrder.DeparmentID AS Expr1, dbo.TblEmpJobOrder.gradeID, dbo.TblEmpJobOrder.JobTypeID AS Expr2, dbo.TblEmpJobOrder.basicSalary,"
     MySQL = MySQL & "               dbo.TblEmpJobOrder.nationalId, dbo.TblEmpJobOrder.TransportTypeID, dbo.TblEmpJobOrder.JobLocation, dbo.TblEmpJobOrder.startDate, dbo.TblEmpJobOrder.startTime,"
      MySQL = MySQL & "              dbo.TblEmpJobOrder.EndDate, dbo.TblEmpJobOrder.EndTime, dbo.TblEmpJobOrder.PaymentVchrNo, dbo.TblEmpJobOrder.PaymentVchrValue,"
    MySQL = MySQL & "                dbo.TblEmpJobOrder.carExpenses, dbo.TblEmpJobOrder.HousingExpenses, dbo.TblEmpJobOrder.JobExpenses, dbo.TblEmpJobOrder.oldExpenses,"
        MySQL = MySQL & "            dbo.TblEmpJobOrder.foodExpenses, dbo.TblEmpJobOrder.carExpenses2, dbo.TblEmpJobOrder.TicketExpenses, dbo.TblEmpJobOrder.totalExpenses,"
       MySQL = MySQL & "             dbo.TblEmpJobOrder.Nationality, dbo.TblEmpJobOrder.Visa1, dbo.TblEmpJobOrder.ticketGo, dbo.TblEmpJobOrder.ticketBack, dbo.TblEmpJobOrder.Task,"
  MySQL = MySQL & "                  dbo.TblEmpJobOrder.Reason, dbo.TblEmpJobOrder.Visa2, dbo.TblEmpJobOrder.Visa3, dbo.TblEmpJobOrder.ok, dbo.TblEmpJobOrder.notok, dbo.TblEmpJobOrder.Manager,"
  MySQL = MySQL & "                  TblEmployee_2.Emp_Name AS Manager_name, dbo.TblEmployee.Emp_Name AS Ceo_Name"
MySQL = MySQL & "  FROM     dbo.TblEmployee RIGHT OUTER JOIN"
      MySQL = MySQL & "              dbo.TblEmpJobOrder LEFT OUTER JOIN"
   MySQL = MySQL & "                 dbo.TblEmpGrades ON dbo.TblEmpJobOrder.gradeID = dbo.TblEmpGrades.gradeid ON dbo.TblEmployee.Emp_ID = dbo.TblEmpJobOrder.Ceo LEFT OUTER JOIN"
    MySQL = MySQL & "                dbo.TblEmployee AS TblEmployee_2 ON dbo.TblEmpJobOrder.Manager = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
   MySQL = MySQL & "                 dbo.TblEmpJobsTypes ON dbo.TblEmpJobOrder.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
  MySQL = MySQL & "                  dbo.TblEmpDepartments ON dbo.TblEmpJobOrder.DeparmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                    dbo.TblUsers ON dbo.TblEmpJobOrder.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                   dbo.TblEmployee AS TblEmployee_1 ON dbo.TblEmpJobOrder.Emp_id = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblEmpJobOrder.Branch_NO = dbo.TblBranchesData.branch_id"

   MySQL = MySQL & "            where 1 = 1 "
   
 If advanceID.Text <> "" Then
        MySQL = MySQL + " and advanceID = " + advanceID.Text
 End If
 
  If emp_Name.BoundText <> "" Then
         MySQL = MySQL & " and TblEmployee_1.emp_Name = '" & (emp_Name.Text) & "'"
  End If
  
  If Nationality.BoundText <> "" Then
        MySQL = MySQL + " and Nationality  = " & Nationality.Text & ""
  End If
  
  If JobTypeName.BoundText <> "" Then
        MySQL = MySQL + " and JobTypeName = '" & (JobTypeName.Text) & "'"
  End If
  
  If branch_name.BoundText <> "" Then
       MySQL = MySQL + " and  branch_name  = '" & (branch_name.Text) & "'"
  End If


      
  If task.Text <> "" Then
        MySQL = MySQL + " and task like '%" & task.Text & "%'"
  End If
  
   
   '///////////////////////////////////
    If Not IsNull(Me.FromDate.value) Then
          MySQL = MySQL + " and  startDate >= " & SQLDate(Me.FromDate.value, True) & ""
    End If
    
  
    
           '///////////////////////////////////

    
   rs.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
             Fg_Journal.Rows = Fg_Journal.FixedRows
                Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
        
    Else
    retrive1
    End If
End Sub
Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    StrSQL = "SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmbarkation.recorddate, dbo.TblEmbarkation.Indate, "
    StrSQL = StrSQL + "                  dbo.TblEmbarkation.indateH, dbo.TblEmbarkation.workdate, dbo.TblEmbarkation.workdateH, dbo.TblEmbarkation.Remark, dbo.TblEmbarkation.PostedDate,"
    StrSQL = StrSQL + "                   dbo.TblEmbarkation.Approved, dbo.TblEmbarkation.Vac_new, dbo.TblEmbarkation.vac_Bak, dbo.TblEmbarkation.ApprovVacPeriod,"
    StrSQL = StrSQL + "                   dbo.TblEmbarkation.ActiveVacPeriod, dbo.TblEmbarkation.MoveVacBalance, dbo.TblEmbarkation.Join_Work, dbo.TblEmbarkation.stratDateH,"
    StrSQL = StrSQL + "                   dbo.TblEmbarkation.EndDateH , dbo.TblEmbarkation.stratDate, dbo.TblEmbarkation.EndDate, dbo.TblEmbarkation.UnPaid_Dis, dbo.TblEmbarkation.ID ,dbo.TblEmbarkation.Emp_ID"
    StrSQL = StrSQL + "      FROM         dbo.TblEmbarkation LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblEmployee ON dbo.TblEmbarkation.Emp_ID = dbo.TblEmployee.Emp_ID"
    BolBegine = False
    StrWhere = ""

    If val(Me.Text19.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.ID >=" & val(Me.Text19.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.ID>=" & val(Me.Text19.Text) & ""
        End If
    End If

    If val(Me.Text20.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.ID<=" & val(Me.Text20.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.ID<=" & val(Me.Text20.Text) & ""
        End If
    End If

    If Me.DcbEmp.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.Emp_ID=" & Me.DcbEmp.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.Emp_ID=" & Me.DcbEmp.BoundText & ""
        End If
    End If



    If Not IsNull(Me.DTPicker5.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.recorddate >=" & SQLDate(Me.DTPicker5.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.recorddate >=" & SQLDate(Me.DTPicker5.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker6.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.recorddate <=" & SQLDate(Me.DTPicker6.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.recorddate <=" & SQLDate(Me.DTPicker6.value, True) & ""
        End If
    End If
''''
    If Not IsNull(Me.FromStartDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.workdate >=" & SQLDate(Me.FromStartDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.workdate >=" & SQLDate(Me.FromStartDate.value, True) & ""
        End If
    End If
        If Not IsNull(Me.ToStartDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.workdate <=" & SQLDate(Me.ToStartDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.workdate <=" & SQLDate(Me.ToStartDate.value, True) & ""
        End If
    End If
''''''
    '-----------------------------------
       If Not IsNull(Me.FromEndDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.stratDate >=" & SQLDate(Me.FromEndDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.stratDate >=" & SQLDate(Me.FromEndDate.value, True) & ""
        End If
    End If
       If Not IsNull(Me.ToEndDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.stratDate <=" & SQLDate(Me.ToEndDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.stratDate <=" & SQLDate(Me.ToEndDate.value, True) & ""
        End If
    End If
    ''''''''''''''''''''''''''''
           If Not IsNull(Me.DTPicker7.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.EndDate >=" & SQLDate(Me.DTPicker7.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.EndDate >=" & SQLDate(Me.DTPicker7.value, True) & ""
        End If
    End If
       If Not IsNull(Me.DTPicker8.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.EndDate <=" & SQLDate(Me.DTPicker8.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.EndDate <=" & SQLDate(Me.DTPicker8.value, True) & ""
        End If
    End If
    ''''''''''''''''''''''''''
If Rd(1).value = True Then
      If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.Vac_new=1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.Vac_new = 1"
        End If
End If
If Rd(2).value = True Then
      If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmbarkation.vac_Bak=1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmbarkation.vac_Bak=1"
        End If
End If
'''''''''///////////////



    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblEmbarkation.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.VSFlexGrid6
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("recorddate").value)) Then
                    .TextMatrix(i, .ColIndex("recorddate")) = Format(rs("recorddate").value, "yyyy/M/d")
                End If
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                .TextMatrix(i, .ColIndex("workdate")) = IIf(IsNull(rs("workdate").value), "", rs("workdate").value)
                .TextMatrix(i, .ColIndex("workdateH")) = IIf(IsNull(rs("workdateH").value), "", rs("workdateH").value)
                
                .TextMatrix(i, .ColIndex("stratDate")) = IIf(IsNull(rs("stratDate").value), "", rs("stratDate").value)
                .TextMatrix(i, .ColIndex("stratDateH")) = IIf(IsNull(rs("stratDateH").value), "", rs("stratDateH").value)
                
                .TextMatrix(i, .ColIndex("EndDateH")) = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
                .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(rs("EndDate").value), "", rs("EndDate").value)
                If Not IsNull(rs("Vac_new").value) Then
                If rs("Vac_new").value = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Vac_new")) = " «· Õ«ﬁ „ÊŸ› ÃœÌœ"
                Else
                .TextMatrix(i, .ColIndex("Vac_new")) = " Join New Employees"
                End If
                End If
                End If
                
                       If Not IsNull(rs("vac_Bak").value) Then
                If rs("vac_Bak").value = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Vac_new")) = "  ⁄Êœ… „‰ ≈Ã«“…"
                Else
                .TextMatrix(i, .ColIndex("Vac_new")) = " Return From Vacation"
                End If
                End If
                End If
                               
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ID"), .Rows - 1, .ColIndex("ID"))
        End With

    End If

End Sub
Private Sub GetData1()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    StrSQL = "SELECT     dbo.TblEmployeeWarrning.id, dbo.TblEmployeeWarrning.recorddate, dbo.TblEmployeeWarrning.Emp_ID, dbo.TblEmployeeWarrning.Remark, "
    StrSQL = StrSQL + "                  dbo.TblEmployeeWarrning.Remark1, dbo.TblEmployeeWarrning.Remark2, dbo.TblEmployeeWarrning.MaxSan, dbo.TblEmployeeWarrning.Freq,"
    StrSQL = StrSQL + "                  dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployeeWarrning.DeptID,"
    StrSQL = StrSQL + "                  dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployeeWarrning.EmpID2,"
    StrSQL = StrSQL + "                  TblEmployee_1.Emp_Name AS Emp_Name2, TblEmployee_1.Fullcode AS Fullcode2, TblEmployee_1.Emp_Namee AS Emp_Namee2,"
    StrSQL = StrSQL + "                  dbo.TblEmployeeWarrning.SanctionID , dbo.TblAdminSanction.name, dbo.TblAdminSanction.NameE"
    StrSQL = StrSQL + "   FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblEmployeeWarrning LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblAdminSanction ON dbo.TblEmployeeWarrning.SanctionID = dbo.TblAdminSanction.ID ON TblEmployee_1.Emp_ID = dbo.TblEmployeeWarrning.EmpID2 ON"
    StrSQL = StrSQL + "                   dbo.TblEmpDepartments.DeparmentID = dbo.TblEmployeeWarrning.DeptID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblEmployee ON dbo.TblEmployeeWarrning.Emp_ID = dbo.TblEmployee.Emp_ID"
    
    StrWhere = ""

    If val(Me.Text23.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.ID >=" & val(Me.Text23.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.ID>=" & val(Me.Text23.Text) & ""
        End If
    End If

    If val(Me.Text22.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.ID<=" & val(Me.Text22.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.ID<=" & val(Me.Text22.Text) & ""
        End If
    End If

    If Me.DcbEmp10.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.Emp_ID=" & Me.DcbEmp10.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.Emp_ID=" & Me.DcbEmp10.BoundText & ""
        End If
    End If
   If Me.DcbEmp11.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.EmpID2=" & Me.DcbEmp11.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.EmpID2=" & Me.DcbEmp11.BoundText & ""
        End If
    End If
 If Me.DcbSanction.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.SanctionID=" & Me.DcbSanction.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.SanctionID=" & Me.DcbSanction.BoundText & ""
        End If
    End If

 If Me.DcbDept.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.DeptID=" & Me.DcbDept.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.DeptID=" & Me.DcbDept.BoundText & ""
        End If
    End If
    
    If Not IsNull(Me.DTPicker15.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.recorddate >=" & SQLDate(Me.DTPicker15.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.recorddate >=" & SQLDate(Me.DTPicker15.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker16.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.recorddate <=" & SQLDate(Me.DTPicker16.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.recorddate <=" & SQLDate(Me.DTPicker16.value, True) & ""
        End If
    End If
   If txtRemark.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployeeWarrning.Remark like N'%" & Me.txtRemark.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployeeWarrning.Remark like N'%" & Me.txtRemark.Text & "%'"
        End If
    End If


    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblEmployeeWarrning.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.VSFlexGrid7
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("recorddate").value)) Then
                    .TextMatrix(i, .ColIndex("recorddate")) = Format(rs("recorddate").value, "yyyy/M/d")
                End If
               If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Emp_Name2")) = IIf(IsNull(rs("Emp_Name2").value), "", rs("Emp_Name2").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Emp_Name2")) = IIf(IsNull(rs("Emp_Namee2").value), "", rs("Emp_Namee2").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                End If
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ID"), .Rows - 1, .ColIndex("ID"))
        End With

    End If

End Sub
Private Sub GetData2()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    StrSQL = "SELECT     dbo.tblKhsmEdafa.KhsmEdafa_ID, dbo.tblKhsmEdafa.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
    StrSQL = StrSQL + "                  dbo.tblKhsmEdafa.KhsmEdafa_Date, dbo.tblKhsmEdafa.KhsmEdafa_Type, dbo.tblKhsmEdafa.KhsmEdafa_Value, dbo.tblKhsmEdafa.KhsmEdafa_Code,"
    StrSQL = StrSQL + "                    dbo.tblKhsmEdafa.RcDate, dbo.tblKhsmEdafa.CalculateValueType, dbo.tblKhsmEdafa.Resonalarm, dbo.tblKhsmEdafa.Val, dbo.tblKhsmEdafa.AlrmOrder,"
    StrSQL = StrSQL + "                    dbo.tblKhsmEdafa.Mofrd, dbo.mofrad.name, dbo.mofrad.nameE, dbo.tblKhsmEdafa.SanctionID, dbo.TblAdminSanction.Name AS Expr1,"
    StrSQL = StrSQL + "                    dbo.TblAdminSanction.NameE AS Expr2"
    StrSQL = StrSQL + "    FROM         dbo.tblKhsmEdafa LEFT OUTER JOIN"
    StrSQL = StrSQL + "                    dbo.TblAdminSanction ON dbo.tblKhsmEdafa.SanctionID = dbo.TblAdminSanction.ID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                    dbo.mofrad ON dbo.tblKhsmEdafa.Mofrd = dbo.mofrad.id LEFT OUTER JOIN"
    StrSQL = StrSQL + "                    dbo.TblEmployee ON dbo.tblKhsmEdafa.Emp_ID = dbo.TblEmployee.Emp_ID"
    
    StrWhere = ""

    If val(Me.Text23.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.KhsmEdafa_ID >=" & val(Me.Text27.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.KhsmEdafa_ID>=" & val(Me.Text27.Text) & ""
        End If
    End If

    If val(Me.Text22.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.KhsmEdafa_ID<=" & val(Me.Text26.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.KhsmEdafa_ID<=" & val(Me.Text26.Text) & ""
        End If
    End If
  If val(Me.CboCalType.ListIndex) <> -1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.CalculateValueType=" & Me.CboCalType.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.CalculateValueType=" & Me.CboCalType.ListIndex & ""
        End If
    End If
    
    If Me.DcbEmp12.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.Emp_ID=" & Me.DcbEmp12.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.Emp_ID=" & Me.DcbEmp12.BoundText & ""
        End If
    End If
   If Me.DCComponent.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.Mofrd=" & Me.DCComponent.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.Mofrd=" & Me.DCComponent.BoundText & ""
        End If
    End If
 If Me.DcbSanction1.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.SanctionID=" & Me.DcbSanction1.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.SanctionID=" & Me.DcbSanction1.BoundText & ""
        End If
    End If


    
    If Not IsNull(Me.DTPicker9.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.RcDate >=" & SQLDate(Me.DTPicker9.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.RcDate >=" & SQLDate(Me.DTPicker9.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker10.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.RcDate <=" & SQLDate(Me.DTPicker10.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.RcDate <=" & SQLDate(Me.DTPicker10.value, True) & ""
        End If
    End If
   If Text28.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblKhsmEdafa.Resonalarm like N'%" & Me.Text28.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblKhsmEdafa.Resonalarm like N'%" & Me.Text28.Text & "%'"
        End If
    End If


    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.tblKhsmEdafa.KhsmEdafa_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.VSFlexGrid8
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("KhsmEdafa_ID").value), "", rs("KhsmEdafa_ID").value)
                        
                If Not (IsNull(rs("RcDate").value)) Then
                    .TextMatrix(i, .ColIndex("recorddate")) = Format(rs("RcDate").value, "yyyy/M/d")
                End If
               If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("nameM")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Expr1").value), "", rs("Expr1").value)
                Else
                .TextMatrix(i, .ColIndex("nameM")) = IIf(IsNull(rs("nameE").value), "", rs("nameE").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Expr2").value), "", rs("Expr2").value)
                End If
                CboCalType1.ListIndex = IIf(IsNull(rs("CalculateValueType").value), -1, rs("CalculateValueType").value)
                .TextMatrix(i, .ColIndex("TypDis")) = CboCalType1.Text
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(rs("Resonalarm").value), "", rs("Resonalarm").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ID"), .Rows - 1, .ColIndex("ID"))
        End With

    End If

End Sub
Private Sub ALLButton11_Click()
On Error Resume Next
Dim ID As Integer
ID = VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, VSFlexGrid4.ColIndex("ID"))
FrmTreament.Retrive (ID)
End Sub

Private Sub ALLButton12_Click()
 Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 Dim MySQL As String
  
 VSFlexGrid4.Rows = VSFlexGrid4.FixedRows
  
 MySQL = " SELECT dbo.TblTreatment.ID, dbo.TblTreatment.RecordDate, dbo.TblTreatment.BranchID, dbo.TblTreatment.NationalID, dbo.TblTreatment.ProjectID, dbo.TblTreatment.JobID,"
    MySQL = MySQL & "                             dbo.TblTreatment.IqamaID, dbo.TblTreatment.IqamaFrom, dbo.TblTreatment.EndDate, dbo.TblTreatment.ExpairDate, dbo.TblTreatment.LongTreatment,"
     MySQL = MySQL & "                            dbo.TblTreatment.TreatmentDate, dbo.TblTreatment.ComputerID, dbo.TblTreatment.Telephone, dbo.TblTreatment.IqamaEx, dbo.TblTreatment.IqamaNew,"
     MySQL = MySQL & "                            dbo.TblTreatment.PasNew, dbo.TblTreatment.EnterNew, dbo.TblTreatment.FinalTrea, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
     MySQL = MySQL & "                            dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Namee,"
      MySQL = MySQL & "                           dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblTreatment.EmpID,"
       MySQL = MySQL & "                          dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
          MySQL = MySQL & "                       dbo.EmpGroupDep.GroupName, dbo.TblEmployee.Nationality, dbo.TblEmployee.Fullcode, dbo.TblTreatment.Iqama, dbo.TblTreatment.ExpairDateH,"
          MySQL = MySQL & "                       dbo.TblTreatment.EndDateH , dbo.TblTreatment.pasid, dbo.TblEmployee.DateEndekamaH, dbo.TblEmployee.NumEkama"
    MySQL = MySQL & "           FROM     dbo.TblBranchesData RIGHT OUTER JOIN"
        MySQL = MySQL & "                         dbo.TblTreatment LEFT OUTER JOIN"
          MySQL = MySQL & "                       dbo.EmpGroupDep ON dbo.TblTreatment.ProjectID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
           MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblTreatment.JobID = dbo.TblEmpJobsTypes.JobTypeID ON dbo.TblBranchesData.branch_id = dbo.TblTreatment.BranchID LEFT OUTER JOIN"
                MySQL = MySQL & "                 dbo.TblEmployee ON dbo.TblTreatment.EmpID = dbo.TblEmployee.Emp_ID"

 
 MySQL = MySQL & "  WHERE     1 = 1 "
  


   
   
 If Text3.Text <> "" Then
        MySQL = MySQL + " and id = " + Text8.Text
 End If
   
  If DataCombo5.BoundText <> "" Then
         MySQL = MySQL & " and TblTreatment.empid = " & val(DataCombo5.BoundText)
  End If
  
  If Text16.Text <> "" Then
        MySQL = MySQL + " and Pasid  = " & Text16.Text & ""
  End If
  
   If Text15.Text <> "" Then
        MySQL = MySQL + " and telephone  = " & Text15.Text & ""
  End If
   
  
 If Text14.Text <> "" Then
  StrSQL = StrSQL + " and NumEkama = '" & (Text14.Text) & "'"
 End If
  
  If DataCombo6.BoundText <> "" Then
        MySQL = MySQL + " and JobTypeName = '" & (DataCombo6.Text) & "'"
  End If
  
  If DataCombo7.BoundText <> "" Then
       MySQL = MySQL + " and  TblTreatment.branchid  = " & val(DataCombo7.BoundText)
  End If

   
   '///////////////////////////////////
    If Not IsNull(Me.DTPicker2.value) Then
          MySQL = MySQL + " and  recorddate = " & SQLDate(Me.DTPicker2.value, True) & ""
    End If
    
  
           '///////////////////////////////////

    
   rs.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
             Fg_Journal.Rows = Fg_Journal.FixedRows
                Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
        
    Else
    Retrive5

End If
End Sub

Private Sub ALLButton13_Click()
    
 fg2.Rows = fg2.FixedRows
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = "SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
StrSQL = StrSQL & "                      dbo.TbVisaDeti.NotionalID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "                      dbo.TbVisa.OrderNo, dbo.TbVisa.VisaNo, dbo.TbVisa.Priod, dbo.TbVisa.DMYPriod, dbo.TbVisa.StarDate, dbo.TbVisa.StarDateH, dbo.TbVisa.EndDate,"
StrSQL = StrSQL & "                      dbo.TbVisa.EndDateH , dbo.TbVisa.ID AS IDM "
StrSQL = StrSQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TbVisa LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TbVisaDeti ON dbo.TbVisa.ID = dbo.TbVisaDeti.VisaID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments ON dbo.TbVisaDeti.CityID = dbo.TblCountriesGovernments.GovernmentID ON"
StrSQL = StrSQL & "                      dbo.Nationality.id = dbo.TbVisaDeti.NotionalID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TbVisaDeti.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & ""
 
    BolBegine = False
    StrWhere = ""



    If val(Me.DcbCity.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisaDeti.CityID=" & Me.DcbCity.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbVisaDeti.CityID =" & Me.DcbCity.BoundText & ""
        End If
    End If
    

    If val(Me.DcbNtionality.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisaDeti.NotionalID=" & Me.DcbNtionality.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbVisaDeti.NotionalID =" & Me.DcbNtionality.BoundText & ""
        End If
    End If

    If Me.TxtOrder.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisa.OrderNo like '%" & Me.TxtOrder.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.OrderNo like '%" & Me.TxtOrder.Text & "%'"
        End If
    End If
   If Me.TxtVisa.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisa.VisaNo like '%" & Me.TxtVisa.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.VisaNo like '%" & Me.TxtVisa.Text & "%'"
        End If
    End If
       If Me.TxtHodono.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisaDeti.HododNo like '%" & Me.TxtHodono.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisaDeti.HododNo like '%" & Me.TxtHodono.Text & "%'"
        End If
    End If

    If Not IsNull(Me.StarDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.StarDate >=" & SQLDate(Me.StarDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.StarDate >=" & SQLDate(Me.StarDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.StarDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.StarDate <=" & SQLDate(Me.StarDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.StarDate <=" & SQLDate(Me.StarDateTo.value, True) & ""
        End If
    End If
    
    If Not IsNull(Me.EndDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.EndDate >=" & SQLDate(Me.EndDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.EndDate >=" & SQLDate(Me.EndDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.EndDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.EndDate <=" & SQLDate(Me.EndDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.EndDate <=" & SQLDate(Me.EndDateTo.value, True) & ""
        End If
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TbVisa.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg2
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDM").value), "", rs("IDM").value)
                        
                If Not (IsNull(rs("EndDate").value)) Then
                    .TextMatrix(i, .ColIndex("EndDate")) = Format(rs("EndDate").value, "yyyy/M/d")
                End If
                    If Not (IsNull(rs("StarDate").value)) Then
                    .TextMatrix(i, .ColIndex("StarDate")) = Format(rs("StarDate").value, "yyyy/M/d")
                End If
             
                .TextMatrix(i, .ColIndex("StarDateH")) = IIf(IsNull(rs("StarDateH").value), "", rs("StarDateH").value)
                .TextMatrix(i, .ColIndex("EndDateH")) = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", rs("name").value)
               
                Else
                 .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                 .TextMatrix(i, .ColIndex("HododNo")) = IIf(IsNull(rs("HododNo").value), "", rs("HododNo").value)
                 .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
                  .TextMatrix(i, .ColIndex("VisaNo")) = IIf(IsNull(rs("VisaNo").value), "", rs("VisaNo").value)
                 .TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
              
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If


End Sub

Private Sub ALLButton14_Click()
On Error Resume Next
Dim ID As Integer
ID = FG.TextMatrix(FG.Row, FG.ColIndex("AdvanceID"))
If send_form = "advreq" Then
FrmEmpsAdvanceRequest.Retrive (ID)
ElseIf send_form = "advreqPayment" Then
FrmPayments.TxtAdvance.Text = ID
End If
End Sub

Private Sub ALLButton15_Click()
On Error Resume Next
Dim ID As Integer
ID = fg2.TextMatrix(fg2.Row, fg2.ColIndex("ID"))
FrmVisa.Retrive (ID)
End Sub

Private Sub ALLButton16_Click()
On Error Resume Next
Dim ID As Integer
ID = VSFlexGrid5.TextMatrix(VSFlexGrid5.Row, VSFlexGrid5.ColIndex("AdvanceID"))
FrmEmpsAdvance.Retrive (ID)

End Sub

Private Sub ALLButton17_Click()

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = "  SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblUsers.UserName, dbo.TblEmpAdvance.AdvanceID,"
StrSQL = StrSQL & "   dbo.TblEmpAdvance.AdvanceDate , dbo.TblEmpAdvance.AdvanceValue, dbo.TblBoxesData.Boxname"
StrSQL = StrSQL & "   FROM     dbo.TblEmployee INNER JOIN"
StrSQL = StrSQL & "   dbo.TblEmpAdvance ON dbo.TblEmployee.Emp_ID = dbo.TblEmpAdvance.Emp_ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblUsers ON dbo.TblEmpAdvance.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblBoxesData ON dbo.TblEmpAdvance.BoxID = dbo.TblBoxesData.BoxID"
    BolBegine = False
    StrWhere = ""

    If val(Me.Text17.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceID >=" & val(Me.Text17.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvance.AdvanceID >=" & val(Me.Text17.Text) & ""
        End If
    End If

    If val(Me.Text18.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceID <=" & val(Me.Text18.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvance.AdvanceID <=" & val(Me.Text18.Text) & ""
        End If
    End If

    If Me.DataCombo8.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DataCombo8.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_ID=" & Me.DataCombo8.BoundText & ""
        End If
    End If

    If Me.DataCombo9.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvance.UserID=" & Me.DataCombo9.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvance.UserID=" & Me.DataCombo9.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DTPicker3.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceDate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvance.AdvanceDate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker4.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceDate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvance.AdvanceDate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        End If
    End If
    If BolBegine = True Then
    '-----------------------------------
        StrWhere = StrWhere & " And ( not(noteid is null) and TblEmpAdvance.AdvanceType =0)"
    Else
        StrWhere = StrWhere & " Where ( not(noteid is null) and TblEmpAdvance.AdvanceType =0)"
    End If
    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblEmpAdvance.AdvanceID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.VSFlexGrid5
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("AdvanceID")) = IIf(IsNull(rs("AdvanceID").value), "", rs("AdvanceID").value)
                        
                If Not (IsNull(rs("AdvanceDate").value)) Then
                    .TextMatrix(i, .ColIndex("AdvanceDate")) = Format(rs("AdvanceDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("AdvanceValue")) = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If





End Sub

Private Sub ALLButton18_Click()
Unload Me
End Sub

Private Sub ALLButton19_Click()
GetData
End Sub

Private Sub ALLButton2_Click()
Dim ID As Integer
Dim Row As Integer
Row = VSFlexGrid1.Row
ID = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("ID")))
FrmBusinessJob.Retrive (ID)
End Sub

Private Sub ALLButton20_Click()
            clear_all Me
DTPicker5.value = ""
DTPicker6.value = ""
FromStartDate.value = ""
ToStartDate.value = ""
DTPicker7.value = ""
DTPicker8.value = ""
ToEndDate.value = ""
FromEndDate.value = ""
End Sub

Private Sub ALLButton21_Click()
Unload Me
End Sub

Private Sub ALLButton22_Click()
GetData1
End Sub

Private Sub ALLButton23_Click()
            clear_all Me
DTPicker15.value = ""
DTPicker16.value = ""
End Sub

Private Sub ALLButton24_Click()
Unload Me
End Sub

Private Sub ALLButton25_Click()
GetData2
End Sub

Private Sub ALLButton26_Click()
            clear_all Me
DTPicker9.value = ""
DTPicker10.value = ""
End Sub

Private Sub ALLButton5_Click()
On Error Resume Next
Dim ID As Integer
ID = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("ID"))
FrmPassports.Retrive (ID)

End Sub

Private Sub ALLButton6_Click()

   Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 Dim MySQL As String
  
 MySQL = " SELECT     dbo.TblPassports.Emp_ID, dbo.TblPassports.recorddate, dbo.TblPassports.Remark, dbo.TblPassports.Posted, dbo.TblPassports.PostedDate, "
 MySQL = MySQL & "  dbo.TblPassports.returned, dbo.TblPassports.returnedDate, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
 MySQL = MySQL & "  dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
 MySQL = MySQL & " dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.pasplace, dbo.TblEmployee.placeEkama, dbo.TblEmployee.Nationality, dbo.TblEmployee.NumEkama,"
 MySQL = MySQL & " dbo.TblPassports.id , dbo.TblPassports.Branch_NO"
 MySQL = MySQL & " FROM         dbo.TblPassports INNER JOIN"
 MySQL = MySQL & " dbo.TblEmployee ON dbo.TblPassports.Emp_ID = dbo.TblEmployee.Emp_ID"
 
 MySQL = MySQL & "  WHERE     1 = 1 "
  


   
   
 If Text3.Text <> "" Then
        MySQL = MySQL + " and id = " + Text3.Text
 End If
   
  If DataCombo1.BoundText <> "" Then
         MySQL = MySQL & " and TblPassports.emp_id = " & val(DataCombo1.BoundText)
  End If
  
  If Nationality3.Text <> "" Then
        MySQL = MySQL + " and Nationality  = " & Nationality3.Text & ""
  End If
  
 If numbEkama.Text <> "" Then
  StrSQL = StrSQL + " and NumbEkama = '" & (numbEkama.Text) & "'"
 End If
  
  If JobTypeName3.BoundText <> "" Then
        'MySQL = MySQL + " and JobTypeName = '" & (JobTypeName3.text) & "'"
  End If
  
  If branch_no.BoundText <> "" Then
       MySQL = MySQL + " and  branch_no  = " & val(branch_no.BoundText)
  End If


      
  If remark.Text <> "" Then
        MySQL = MySQL + " and remark like '%" & remark.Text & "%'"
  End If
  
   
   '///////////////////////////////////
    If Not IsNull(Me.recorddate.value) Then
          MySQL = MySQL + " and  recorddate >= " & SQLDate(Me.recorddate.value, True) & ""
    End If
    
  
    
           '///////////////////////////////////

    
   rs.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
             VSFlexGrid2.Rows = VSFlexGrid2.FixedRows
                Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
        
    Else
    Retrive3

End If

End Sub

Private Sub ALLButton7_Click()
 VSFlexGrid3.Rows = VSFlexGrid3.FixedRows

 Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 Dim MySQL As String
  

MySQL = "  SELECT     dbo.TblAdvancedHousing.AdvanceID, dbo.TblAdvancedHousing.Branch_NO, dbo.TblAdvancedHousing.Emp_id, dbo.TblAdvancedHousing.UserID, "
MySQL = MySQL & "  dbo.TblAdvancedHousing.AdvanceDate, dbo.TblAdvancedHousing.DeparmentID, dbo.TblAdvancedHousing.JobTypeID, dbo.TblAdvancedHousing.basicSalary,"
MySQL = MySQL & "  dbo.TblAdvancedHousing.payValue, dbo.TblAdvancedHousing.lastPayDateFrom, dbo.TblAdvancedHousing.lastPayDateTo, dbo.TblAdvancedHousing.currentdayFrom,"
MySQL = MySQL & "  dbo.TblAdvancedHousing.currentdayTo, dbo.TblAdvancedHousing.PaymentVchrNo, dbo.TblAdvancedHousing.Remarks, dbo.TblAdvancedHousing.Posted,"
MySQL = MySQL & "  dbo.TblAdvancedHousing.PostedDate, dbo.TblAdvancedHousing.NoteSerial, dbo.TblAdvancedHousing.Approved, dbo.TblAdvancedHousing.Transaction_ID,"
MySQL = MySQL & "  dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name, dbo.TblUsers.UserName, dbo.TblEmpDepartments.DepartmentName,"
MySQL = MySQL & "  dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_Code,"
MySQL = MySQL & "  dbo.TblEmployee.emp_name , dbo.TblEmployee.Emp_Namee"
MySQL = MySQL & "  FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
MySQL = MySQL & "  dbo.TblAdvancedHousing LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblEmployee ON dbo.TblAdvancedHousing.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblEmpJobsTypes ON dbo.TblAdvancedHousing.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID ON"
MySQL = MySQL & "  dbo.TblEmpDepartments.DeparmentID = dbo.TblAdvancedHousing.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblUsers ON dbo.TblAdvancedHousing.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblBranchesData ON dbo.TblAdvancedHousing.Branch_NO = dbo.TblBranchesData.branch_id"
 
 MySQL = MySQL & "  WHERE     1 = 1 "
  
 If Text12.Text <> "" Then
        MySQL = MySQL + " and  Advanceid = " + Text12.Text
 End If
   
  If Text9.Text <> "" Then
         MySQL = MySQL & " and payValue = " & val(Text9.Text)
  End If
  
  If Nationality3.Text <> "" Then
        MySQL = MySQL + " and Nationality  = " & Nationality3.Text & ""
  End If
  
 If numbEkama.Text <> "" Then
  StrSQL = StrSQL + " and NumbEkama = '" & (numbEkama.Text) & "'"
 End If
  
 If DataCombo2.BoundText <> "" Then
        MySQL = MySQL + " and TblAdvancedHousing.emp_id = " & val(DataCombo2.BoundText)
  End If
  
  
  If DataCombo3.BoundText <> "" Then
        MySQL = MySQL + " and JobTypeName = '" & (DataCombo3.Text) & "'"
  End If
  
  If DataCombo4.BoundText <> "" Then
       MySQL = MySQL + " and  branch_name  = '" & val(DataCombo4.Text) & "'"
  End If


      
   '///////////////////////////////////
    If Not IsNull(Me.lastPayDateFrom.value) Then
          MySQL = MySQL + " and  lastPayDateFrom = " & SQLDate(Me.lastPayDateFrom.value, True) & ""
    End If
    
     If Not IsNull(Me.currentdayFrom.value) Then
          MySQL = MySQL + " and  currentdayFrom = " & SQLDate(Me.currentdayFrom.value, True) & ""
    End If
    
           '///////////////////////////////////

    
   rs.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
             VSFlexGrid3.Rows = Fg_Journal.FixedRows
                Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
        
    Else
    Retrive4

End If

End Sub

Private Sub ALLButton8_Click()
Dim ID As Integer
Dim Row As Integer
Row = VSFlexGrid3.Row
ID = val(VSFlexGrid3.TextMatrix(VSFlexGrid3.Row, VSFlexGrid3.ColIndex("ID")))
FrmAdvancedHousingpayments.Retrive (ID)
End Sub

Private Sub BtnOK_Click()
Dim ID As Integer
Dim Row As Integer
Row = Fg_Journal.Row
ID = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("ID")))
FormEmpMoveDepartment.Retrive (ID)
End Sub
Private Sub btnSearch_Click()

Fg_Journal.Rows = Fg_Journal.FixedRows

   Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 Dim MySQL As String
  

MySQL = "   SELECT dbo.TblMoveEmp1.BranchID, dbo.TblMoveEmp1.ID, dbo.TblMoveEmp1.RecordDate, dbo.TblMoveEmp1.EmpID, dbo.TblMoveEmp1.FromDepart, dbo.TblMoveEmp1.ToDepart,"
MySQL = MySQL + "                     dbo.TblMoveEmp1.ManagerID, dbo.TblMoveEmp1.moveDate, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL + "                     TblEmployee_1.Emp_Code AS mangercode, TblEmployee_1.Emp_Name AS mangername, TblEmployee_1.Emp_Namee AS mangernamee,"
 MySQL = MySQL + "                    TblEmpDepartments_2.DepartmentName AS deptom, TblEmpDepartments_2.DepartmentNamee AS deptomE, TblEmpDepartments_1.DepartmentName AS depfrom,"
   MySQL = MySQL + "                  TblEmpDepartments_1.DepartmentNamee AS depfrome, dbo.TblMoveEmp1.JobID, TblEmpJobsTypes_1.JobTypeName, TblEmpJobsTypes_1.JobTypeNamee,"
   MySQL = MySQL + "                  TblEmployee_1.Emp_Name AS Namea, TblEmployee_1.Emp_Name1 AS Namea1, TblEmployee_1.Emp_Name2 AS Namea2, TblEmployee_1.Emp_Name3 AS Namea3,"
    MySQL = MySQL + "                 TblEmployee_1.Emp_Name4 AS Namea4, TblEmployee_1.Emp_Namee4 AS Namee4, TblEmployee_1.Emp_Namee3 AS Namee3, TblEmployee_1.Emp_Namee2 AS Namee2,"
   MySQL = MySQL + "                  TblEmployee_1.Emp_Namee1 AS Namee1, dbo.TblMoveEmp1.ProjectFrom, dbo.TblMoveEmp1.ProjectTo, dbo.TblMoveEmp1.basicSalary, dbo.TblMoveEmp1.Reson,"
     MySQL = MySQL + "                EmpGroupDep_2.GroupName AS frmdep, EmpGroupDep_1.GroupName AS todep, dbo.TblMoveEmp1.JobTo, TblEmpJobsTypes_1.JobTypeName AS namejob,"
       MySQL = MySQL + "              TblEmpJobsTypes_1.JobTypeNamee AS nameejob, dbo.TblMoveEmp1.UserID, dbo.TblMoveEmp1.posted, TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name,"
       MySQL = MySQL + "              TblEmployee_2.Emp_Namee, TblEmpJobsTypes_2.JobTypeName AS frmjob, TblEmpJobsTypes_2.JobTypeNamee AS frmjobE,"
    MySQL = MySQL + "                 TblEmployee_2.BignDateWork AS BignDateWork1, TblEmployee_2.Fullcode"
MySQL = MySQL + "   FROM     dbo.TblEmployee AS TblEmployee_1 RIGHT OUTER JOIN"
   MySQL = MySQL + "                 dbo.TblMoveEmp1 LEFT OUTER JOIN"
     MySQL = MySQL + "                dbo.TblEmpJobsTypes AS TblEmpJobsTypes_1 ON dbo.TblMoveEmp1.JobTo = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
       MySQL = MySQL + "              dbo.EmpGroupDep AS EmpGroupDep_1 ON dbo.TblMoveEmp1.ProjectTo = EmpGroupDep_1.GroupID LEFT OUTER JOIN"
        MySQL = MySQL + "             dbo.EmpGroupDep AS EmpGroupDep_2 ON dbo.TblMoveEmp1.ProjectFrom = EmpGroupDep_2.GroupID LEFT OUTER JOIN"
       MySQL = MySQL + "              dbo.TblBranchesData ON dbo.TblMoveEmp1.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        MySQL = MySQL + "             dbo.TblEmployee AS TblEmployee_2 ON dbo.TblMoveEmp1.EmpID = TblEmployee_2.Emp_ID ON TblEmployee_1.Emp_ID = dbo.TblMoveEmp1.ManagerID LEFT OUTER JOIN"
         MySQL = MySQL + "            dbo.TblEmpDepartments AS TblEmpDepartments_1 ON dbo.TblMoveEmp1.FromDepart = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
           MySQL = MySQL + "          dbo.TblEmpDepartments AS TblEmpDepartments_2 ON dbo.TblMoveEmp1.ToDepart = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
          MySQL = MySQL + "           dbo.TblEmpJobsTypes AS TblEmpJobsTypes_2 ON dbo.TblMoveEmp1.JobID = TblEmpJobsTypes_2.JobTypeID"

   MySQL = MySQL & "  where 1 =1   "
   
 If txtid.Text <> "" Then
        MySQL = MySQL + " and ID = " + txtid.Text
 End If
  
  
  If DcboEmpName.BoundText <> "" Then
         MySQL = MySQL & " and empID = " & val(DcboEmpName.BoundText)
  End If
  
  If DcmbFromDepart.BoundText <> "" Then
        MySQL = MySQL + " and  FromDepart  = " & val(DcmbFromDepart.Text)
  End If
  
  If dcmbFromProject.BoundText <> "" Then
        MySQL = MySQL + " and  ProjectFrom  = " & val(dcmbFromProject.Text)
  End If
  
  If DcboJobsType.BoundText <> "" Then
        MySQL = MySQL + " and TblEmpJobsTypes_1.JobTypeName = '" & (DcboJobsType.Text) & "'"
  End If
  
  If dcBranch.BoundText <> "" Then
       MySQL = MySQL + " and  TblMoveEmp1.branchID  = " & val(dcBranch.Text)
  End If

  If DcmbManagerID.BoundText <> "" Then
       MySQL = MySQL + " and  ManagerID  = " & val(DcmbManagerID.BoundText)
  End If
      
  If txtreson.Text <> "" Then
        MySQL = MySQL + " and reson like '%" & txtreson.Text & "%'"
  End If
  
   
   '///////////////////////////////////
    If Not IsNull(Me.FromDate.value) Then
          MySQL = MySQL + " and  MoveDate = " & SQLDate(Me.FromDate.value, True) & ""
    End If
    
      
          '///////////////////////////////////

    
   rs.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
             Fg_Journal.Rows = Fg_Journal.FixedRows
                Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
        
    Else
    Retrive

End If
End Sub



Private Sub Cmd_Click(Index As Integer)
 FG.Rows = FG.FixedRows
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  '  StrSQL = "SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name," & "dbo.TblUsers.UserName, dbo.TblEmpAdvance.AdvanceID,dbo.TblEmpAdvance.AdvanceDate ," & "dbo.TblEmpAdvance.AdvanceValue, dbo.TblBoxesData.BoxName "
  '  StrSQL = StrSQL + " FROM dbo.TblEmployee INNER JOIN"
  '  StrSQL = StrSQL + " dbo.TblEmpAdvanceRequest ON dbo.TblEmployee.Emp_ID = dbo.TblEmpAdvance.Emp_ID INNER JOIN"
  '  StrSQL = StrSQL + " dbo.TblUsers ON dbo.TblEmpAdvanceRequest.UserID = dbo.TblUsers.UserID  "
    
    
   StrSQL = "SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,dbo.TblUsers.UserName,"
 StrSQL = StrSQL + " dbo.TblEmpAdvanceRequest.AdvanceID , dbo.TblEmpAdvanceRequest.AdvanceDate, dbo.TblEmpAdvanceRequest.AdvanceValue"
 StrSQL = StrSQL + " FROM dbo.TblEmployee INNER JOIN dbo.TblEmpAdvanceRequest ON dbo.TblEmployee.Emp_ID = dbo.TblEmpAdvanceRequest.Emp_ID"
StrSQL = StrSQL + "  INNER JOIN dbo.TblUsers ON dbo.TblEmpAdvanceRequest.UserID = dbo.TblUsers.UserID  "
    
    
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvanceRequest.AdvanceID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvanceRequest.AdvanceID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvanceRequest.AdvanceID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvanceRequest.AdvanceID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If

    If Me.DCEmp_Name.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & val(Me.DCEmp_Name.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_ID=" & val(Me.DCEmp_Name.BoundText) & ""
        End If
    End If

    If Me.DCUser.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvanceRequest.UserID=" & Me.DCUser.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvanceRequest.UserID=" & Me.DCUser.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvanceRequest.AdvanceDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvanceRequest.AdvanceDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAdvanceRequest.AdvanceDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAdvanceRequest.AdvanceDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    
    If Me.returntype = 1 Or Me.returntype = 2 Then
     StrSQL = StrSQL & " and approved=1"
    End If
    
    StrSQL = StrSQL & " Order By dbo.TblEmpAdvanceRequest.AdvanceID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
         '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
'                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
            '    Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("AdvanceID")) = IIf(IsNull(rs("AdvanceID").value), "", rs("AdvanceID").value)
                        
                If Not (IsNull(rs("AdvanceDate").value)) Then
                    .TextMatrix(i, .ColIndex("AdvanceDate")) = Format(rs("AdvanceDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("AdvanceValue")) = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
'                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
            Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If




End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        ''Call Reload(Me)
 
    Else
        my_language = "A"
 
        ''Call Reload(Me)
    End If

End Sub










Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub

Private Sub DcbEmp_Click(Area As Integer)
  If val(DcbEmp.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcbEmp.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub

Private Sub DcbEmp10_Change()
DcbEmp10_Click (0)
End Sub

Private Sub DcbEmp10_Click(Area As Integer)
    If val(DcbEmp10.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmp10.BoundText, EmpCode
    Text21.Text = EmpCode
End Sub

Private Sub DcbEmp11_Change()
DcbEmp11_Click (0)
End Sub

Private Sub DcbEmp11_Click(Area As Integer)
    If val(DcbEmp11.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmp11.BoundText, EmpCode
    TxtSearchCode2.Text = EmpCode
End Sub

Private Sub DcbEmp12_Change()
DcbEmp12_Click (0)
End Sub

Private Sub DcbEmp12_Click(Area As Integer)
    If val(DcbEmp12.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmp12.BoundText, EmpCode
    Text24.Text = EmpCode
End Sub

Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 Else
      KeyAscii = 0
 End If
End Sub








Private Sub DTPicker7_Change()
If Not (IsNull(DTPicker7.value)) Then
NourHijriCal1.value = ToHijriDate(DTPicker7.value)
End If
End Sub

Private Sub DTPicker8_Change()
If Not (IsNull(DTPicker8.value)) Then
NourHijriCal2.value = ToHijriDate(DTPicker8.value)
End If
End Sub

Private Sub Fg_Click()
ALLButton14_Click
End Sub

Private Sub Fg_Journal_Click()
BtnOK_Click
End Sub

Private Sub fg2_Click()
ALLButton15_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    TxtModFlg.Text = "R"
    Set rs = New ADODB.Recordset

       Set Dcombos = New ClsDataCombos


   StarDateFrom.value = Date
    StarDateTo.value = Date
    EndDateFrom.value = Date
    EndDateTo.value = Date
    DtpDateFrom.value = Date
    DtpDateTo.value = Date
    DTPicker2.value = Date
    lastPayDateFrom.value = Date
    currentdayFrom.value = Date
    recorddate.value = Date
    DTPicker1.value = Date
    FromDate.value = Date
 DTPicker3.value = Date
  DTPicker4.value = Date
  
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = " select id,name from mofrad where   MofrdDiscount=1"
    Else
        My_SQL = " select id,namee from mofrad where   MofrdDiscount=1"
    End If

    fill_combo DCComponent, My_SQL
    
    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboCalType
            .Clear
            .AddItem "Œ’„ √Ì«„ „‰ «·„— »"
            .AddItem "Œ’„ ﬁÌ„… ‰ﬁœÌ… „‰ «·„— »"
          End With
           With Me.CboCalType1
            .Clear
            .AddItem "Œ’„ √Ì«„ „‰ «·„— »"
            .AddItem "Œ’„ ﬁÌ„… ‰ﬁœÌ… „‰ «·„— »"
            End With
          Else
With Me.CboCalType
            .Clear
            .AddItem "Days From Salry"
            .AddItem "Value"
        End With
        
        With Me.CboCalType1
            .Clear
            .AddItem "Days From Salry"
            .AddItem "Value"
        End With

    End If
    DTPicker10.value = ""
  DTPicker9.value = ""
    StarDateFrom.value = Null
    StarDateTo.value = Null
    EndDateFrom.value = Null
    EndDateTo.value = Null
    DtpDateFrom.value = Null
    DtpDateTo.value = Null
    DTPicker2.value = Null
    lastPayDateFrom.value = Null
    currentdayFrom.value = Null
    recorddate.value = Null
    DTPicker1.value = Null
    FromDate.value = Null
    DTPicker3.value = Null
  DTPicker4.value = Null
  
    
    
   '///////////////////Emp Move
      '  Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBranches Me.dcBranch
     Dcombos.GetEmpDepartments Me.DcmbFromDepart
   '  Dcombos.GetEmpDepartments Me.DcmbToDepart
     Dcombos.GetEmployees Me.DcmbManagerID
     Dcombos.GetEmpJobsTypes Me.DcboJobsType
 '  Dcombos.GetEmpJobsTypes Me.DcmbToJob
    Dcombos.GetEmpLocations Me.dcmbFromProject ' location
  ' Dcombos.GetEmpLocations Me.dcmbToProject ' location
    Dcombos.GetUsers DCUser
     
Dcombos.GetEmployees Me.DcbEmp12
     Dcombos.GetEmployees Me.emp_Name
    Dcombos.GetBranches Me.branch_name
     Dcombos.GetEmpDepartments Me.DcmbFromDepart
   '  Dcombos.GetEmpDepartments Me.DcmbToDepart
     Dcombos.GetEmployees Me.manager
     Dcombos.GetEmpJobsTypes Me.JobTypeName
 '  Dcombos.GetEmpJobsTypes Me.DcmbToJob
    Dcombos.GetEmpLocations Me.Location ' location
  ' Dcombos.GetEmpLocations Me.dcmbToProject ' location
  
  '//////////////////Passport
     Dcombos.GetEmployees Me.DataCombo1
    Dcombos.GetBranches Me.branch_no
     Dcombos.GetEmpDepartments Me.DcmbFromDepart
   '  Dcombos.GetEmpDepartments Me.DcmbToDepart
     Dcombos.GetEmployees Me.manager
     Dcombos.GetEmpJobsTypes Me.JobTypeName3
 '  Dcombos.GetEmpJobsTypes Me.DcmbToJob
    Dcombos.GetEmpLocations Me.Location ' location
  ' Dcombos.GetEmpLocations Me.dcmbToProject ' location
  
   '//////////////////Passport
     Dcombos.GetEmployees Me.DataCombo2
     Dcombos.GetBranches Me.DataCombo4
     Dcombos.GetEmpJobsTypes DataCombo6

      '//////////////////Passport
     Dcombos.GetEmployees Me.DataCombo5
     Dcombos.GetBranches Me.DataCombo7
     Dcombos.GetEmpJobsTypes DataCombo3
    
    'advancerequest
    Dcombos.GetEmployees Me.DCEmp_Name
   SetDtpickerDate Me.DTPicker9
    SetDtpickerDate Me.DTPicker10
    
    '////////////////Visa
      Dcombos.GetEmployees Me.DcbNtionality
     ' Dcombos.GetEmployees Me.DcbNtionality
          Dcombos.GetEmpDepartments Me.DcbDept
    Dcombos.GetAdminSanction Me.DcbSanction1
    Dcombos.GetAdminSanction Me.DcbSanction
    SetDtpickerDate Me.DTPicker15
    SetDtpickerDate Me.DTPicker16
Dcombos.GetEmployees Me.DcbEmp10
Dcombos.GetEmployees Me.DcbEmp11
'////////////////////////////////Emp Advance
Dcombos.GetEmployees Me.DataCombo8
      Dcombos.GetUsers Me.DataCombo9
      
     
    If SystemOptions.UserInterface = EnglishInterface And send_form <> "Embra" Then
    SetInterface Me
    ChangeLang
    End If
If send_form = "Khsm" Then
 If SystemOptions.UserInterface = EnglishInterface Then
    ' SetInterface Me
        ChangeLang4
    End If

Frame24.Visible = True

Frame20.Visible = False
Frame1.Visible = False
 fram_empadv.Visible = False
    fram_empMove.Visible = False
      Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
  '  fram_empMove.Visible = False
    Fram_Visa.Visible = False
    
 ElseIf send_form = "Warning" Then
   Frame24.Visible = False
 If SystemOptions.UserInterface = EnglishInterface Then
    ' SetInterface Me
        ChangeLang3
    End If



Frame20.Visible = True
Frame1.Visible = False
 fram_empadv.Visible = False
    fram_empMove.Visible = False
      Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
  '  fram_empMove.Visible = False
    Fram_Visa.Visible = False
      
ElseIf send_form = "Embra" Then
 If SystemOptions.UserInterface = EnglishInterface Then
     SetInterface Me
        ChangeLang2
    End If
    Frame24.Visible = False
    Frame20.Visible = False
    SetDtpickerDate Me.DTPicker5
SetDtpickerDate Me.DTPicker6
SetDtpickerDate Me.DTPicker7
SetDtpickerDate Me.DTPicker8
SetDtpickerDate Me.FromStartDate
SetDtpickerDate Me.ToStartDate
SetDtpickerDate Me.FromEndDate
SetDtpickerDate Me.ToEndDate
Dcombos.GetEmployees Me.DcbEmp
 Frame24.Visible = False
Frame1.Visible = True
 fram_empadv.Visible = False
    fram_empMove.Visible = False
      Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
  '  fram_empMove.Visible = False
    Fram_Visa.Visible = False
    
ElseIf send_form = "empmov" Then
Frame1.Visible = False
Frame20.Visible = False
 Frame24.Visible = False
 fram_empadv.Visible = False
    fram_empMove.Visible = True
      Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
  '  fram_empMove.Visible = False
    Fram_Visa.Visible = False
ElseIf send_form = "BJ" Then
Frame20.Visible = False
 Frame24.Visible = False
Frame1.Visible = False
 fram_empadv.Visible = False
    Fram_businessjob.Visible = True
    Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
   ' Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = False
ElseIf send_form = "Passports" Then
Frame20.Visible = False
 Frame24.Visible = False
Frame1.Visible = False
 fram_empadv.Visible = False
    Fram_Passports.Visible = True
     Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    'Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = False
ElseIf send_form = "adv" Then
Frame20.Visible = False
 Frame24.Visible = False
Frame1.Visible = False
 fram_empadv.Visible = False
    Fram_adv.Visible = True
     Fram_treat.Visible = False
    Fram_Advreq.Visible = False
   ' Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = False
ElseIf send_form = "treat" Then
Frame1.Visible = False
Frame20.Visible = False
 Frame24.Visible = False
 fram_empadv.Visible = False
    Fram_treat.Visible = True
    'Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = False
    
ElseIf send_form = "advreq" Or send_form = "advreqPayment" Then
Frame20.Visible = False
 Frame24.Visible = False
Frame1.Visible = False
 fram_empadv.Visible = False
    Fram_Advreq.Visible = True
  Fram_treat.Visible = False
   ' Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = False
ElseIf send_form = "visa" Then
Frame20.Visible = False
 Frame24.Visible = False
Frame1.Visible = False
    fram_empadv.Visible = False
    Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = True
ElseIf send_form = "emp_adv" Then
 Frame24.Visible = False
Frame20.Visible = False
Frame1.Visible = False
    fram_empadv.Visible = True
    Fram_treat.Visible = False
    Fram_Advreq.Visible = False
    Fram_adv.Visible = False
    Fram_Passports.Visible = False
    Fram_businessjob.Visible = False
    fram_empMove.Visible = False
    Fram_Visa.Visible = False
End If

 

End Sub



Private Sub ChangeLang()
  

  Label38.Caption = "bill no"
 Label37.Caption = "employee name"
  Label39.Caption = "Eqama No."
  Label32.Caption = "job"
  lblbranch(4).Caption = "Branch"
  lbl(4).Caption = "Date"
  Label34.Caption = "Passport No."
  Label42.Caption = "Telephone No"
  ALLButton12.Caption = "Search"
  ALLButton11.Caption = "Ok"
 Label41.Caption = "Search To hom Concern"
  With VSFlexGrid4
 .TextMatrix(0, .ColIndex("id")) = "No."
 .TextMatrix(0, .ColIndex("emp_name")) = "Employee Name"
.TextMatrix(0, .ColIndex("branch_no")) = "Branch"
.TextMatrix(0, .ColIndex("numbekama")) = "Eqama No"
.TextMatrix(0, .ColIndex("pasid")) = "Passport No"
.TextMatrix(0, .ColIndex("recordDate")) = " Date"
End With

'advanceHousing
Label30.Caption = "Billl no."
Label31.Caption = "Employee Name"
Label29.Caption = "Due Expenses"
Label35.Caption = "Job"
lbl(3).Caption = "date Last Expenses"
lblbranch(3).Caption = "Branch"
Label33.Caption = "Current Expenses"
ALLButton7.Caption = "Search"
ALLButton8.Caption = "Ok"
  Label36.Caption = "Search Advanced Housing Allowance"
  With VSFlexGrid3
   .TextMatrix(0, .ColIndex("JobTypeName")) = "Job Type "
    .TextMatrix(0, .ColIndex("recorddate")) = "Date "
      .TextMatrix(0, .ColIndex("remark")) = "Due Expenses "
        .TextMatrix(0, .ColIndex("lastpay")) = "Due Expenses Date "
 .TextMatrix(0, .ColIndex("id")) = "No."
 .TextMatrix(0, .ColIndex("emp_name")) = "Employee Name"
.TextMatrix(0, .ColIndex("branch_no")) = "Branch"
'.TextMatrix(0, .ColIndex("numbekama")) = "Eqama No"
End With

'visa

lbl(19).Caption = "bill no"
lbl(18).Caption = "Visa No"
lbl(17).Caption = "Boundary No"
lbl(16).Caption = "Nationality"
lbl(15).Caption = "City"
Fra(0).Caption = "Record Date"
Fra(3).Caption = "Record Date"
lbl(9).Caption = "from"
lbl(10).Caption = "To"
lbl(14).Caption = "from"
lbl(13).Caption = "TO"
ALLButton15.Caption = "OK"
ALLButton13.Caption = "search"
Label43.Caption = "Visa Search"
With fg2
 .TextMatrix(0, .ColIndex("Serial")) = "Serial."
  .TextMatrix(0, .ColIndex("GovernmentName")) = "City."
    .TextMatrix(0, .ColIndex("StarDateH")) = "Star Date Hijri."
 .TextMatrix(0, .ColIndex("OrderNo")) = "Order No"
.TextMatrix(0, .ColIndex("VisaNo")) = "Visa No"
.TextMatrix(0, .ColIndex("HododNo")) = "Boundary No"
.TextMatrix(0, .ColIndex("StarDate")) = "Start Date "
.TextMatrix(0, .ColIndex("EndDate")) = "End Date"
.TextMatrix(0, .ColIndex("EndDateH")) = "End Date Hijri "
.TextMatrix(0, .ColIndex("name")) = "nationality"
'.TextMatrix(0, .ColIndex("GovernmentName")) = "City"
End With
 


'passport
Label23.Caption = "Process NO"
Label19.Caption = "Employee Name"
Label24.Caption = "Nationality"
Label15.Caption = "Eqama No"
Label14.Caption = "JOb"
lblbranch(2).Caption = "Branch"
Label17.Caption = "Passport nO"
Label16.Caption = "Mission"
lbl(2).Caption = "Date"
ALLButton6.Caption = "Search"
ALLButton5.Caption = "Ok"
Label27.Caption = "Search Delivery of passports Employees "
  
With VSFlexGrid2
 .TextMatrix(0, .ColIndex("ID")) = "No."
 .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
.TextMatrix(0, .ColIndex("JobTypeName")) = "Type"
.TextMatrix(0, .ColIndex("branch_No")) = "Branch"
.TextMatrix(0, .ColIndex("recorddate")) = "Date"
.TextMatrix(0, .ColIndex("remark")) = "Mission"
End With
 
 
 
'business job
Label3.Caption = "Order no"
Label4.Caption = "employee name"
Label2.Caption = "Nationality"
Label11.Caption = "Location"
Label12.Caption = "job"
lblbranch(1).Caption = "Branch"
Label8.Caption = "Manager"
Label10.Caption = "Task "
ALLButton1.Caption = "Search"
 lbl(1).Caption = "Task Date"
ALLButton2.Caption = "Ok"
Label13.Caption = "Search Business Task"
With VSFlexGrid1
 .TextMatrix(0, .ColIndex("ID")) = "No."
 .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
.TextMatrix(0, .ColIndex("JobTypeName")) = "job"
.TextMatrix(0, .ColIndex("Manager_Name")) = "Manager"
.TextMatrix(0, .ColIndex("branch_Name")) = "Branch name"
.TextMatrix(0, .ColIndex("task")) = "task"
.TextMatrix(0, .ColIndex("startDate")) = "start Date"
End With


'emp move

'UserNameLabel20.Caption = " No"
Label6.Caption = "Employee Name"
Label18.Caption = "Department"
Label22.Caption = "Location"
Label7.Caption = "Job"
 lblbranch(0).Caption = "Branch"
 Label26.Caption = "Manager"
 lbl(0).Caption = "move date"
 Label21.Caption = "reason"
btnSearch.Caption = "search"
btnOk.Caption = "OK"
 Label9.Caption = "Search Employee Move"

With Fg_Journal
 .TextMatrix(0, .ColIndex("ID")) = "No."
 .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
.TextMatrix(0, .ColIndex("jobtypName")) = "job"
.TextMatrix(0, .ColIndex("Mangername")) = "Manager"
.TextMatrix(0, .ColIndex("branch_Name")) = "Branch name"
.TextMatrix(0, .ColIndex("reson")) = "reason"
.TextMatrix(0, .ColIndex("movedate")) = "move Date"
.TextMatrix(0, .ColIndex("value")) = "location"
End With
 
 
 Label20.Caption = "No."
 
'advreq
Fra(1).Caption = "Record Date"
Fra(2).Caption = "Process No"
lbl(11).Caption = "Employee Name"
lbl(6).Caption = "From"
lbl(5).Caption = "To"
Cmd(0).Caption = "Search"
ALLButton14.Caption = "Ok"
Label50.Caption = "Search Cash Advanced Request"
lbl(7).Caption = "From"
lbl(8).Caption = "to"
lbl(12).Caption = "user"
With FG
 .TextMatrix(0, .ColIndex("Serial")) = "Serial."
 .TextMatrix(0, .ColIndex("AdvanceID")) = "Process No "
.TextMatrix(0, .ColIndex("AdvanceDate")) = "Date"
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
.TextMatrix(0, .ColIndex("AdvanceValue")) = "value"
.TextMatrix(0, .ColIndex("BoxName")) = "Box Name"
'.TextMatrix(0, .ColIndex("movedate")) = "move Date"
.TextMatrix(0, .ColIndex("UserName")) = "User Name"
End With





 
 
 
'lbl(0).Caption = "Branch"
'    Label20.Caption = "No."
''    Label6.Caption = "Country"
'    Label15.Caption = "Begining Inv. No."
'    Label18.Caption = "Type"
'    Label22.Caption = "Currency"
'    Label7.Caption = "Bank"
'    Label1.Caption = "Name"
'    Label26.Caption = "Vendor"
'    Label21.Caption = "From Date"
'    Label25.Caption = "To Date"
'    lbl(22).Caption = "last shipment date"
'    lbl(21).Caption = "End Date"
'    btnSearch.Caption = "Search"
'    btnOk.Caption = "OK"
'    Me.Caption = "LC Search"
'    Label9.Caption = "LC Search"

    Label44.Caption = "Search Employee Advanced"
    ALLButton16.Caption = "Ok"
    Fra(4).Caption = "Record Date"
    ALLButton17.Caption = "Search"
    lbl(22).Caption = "From"
    lbl(20).Caption = "To"
    lbl(26).Caption = "User Name"
    lbl(25).Caption = "Employee"
    Fra(5).Caption = "Process No"
    lbl(23).Caption = "From"
    lbl(24).Caption = "To"
    
With VSFlexGrid5
 .TextMatrix(0, .ColIndex("Serial")) = "Serial."
 .TextMatrix(0, .ColIndex("AdvanceID")) = "Process No "
.TextMatrix(0, .ColIndex("AdvanceDate")) = "Date"
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
.TextMatrix(0, .ColIndex("AdvanceValue")) = "value"
.TextMatrix(0, .ColIndex("BoxName")) = "Box Name"
'.TextMatrix(0, .ColIndex("movedate")) = "move Date"
.TextMatrix(0, .ColIndex("UserName")) = "User Name"
End With
    
    
End Sub

Private Sub Retrive()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************


    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.Fg_Journal
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                
                
                  .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                  .TextMatrix(i, .ColIndex("reson")) = IIf(IsNull(rs("reson").value), "", rs("reson").value)
                  .TextMatrix(i, .ColIndex("movedate")) = IIf(IsNull(rs("movedate").value), "", rs("movedate").value)
                
                  .TextMatrix(i, .ColIndex("jobtypName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                 .TextMatrix(i, .ColIndex("Mangername")) = IIf(IsNull(rs("Mangername").value), "", rs("Mangername").value)
                  .TextMatrix(i, .ColIndex("movedate")) = IIf(IsNull(rs("moveDate").value), "", rs("moveDate").value)
                                
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub


Private Sub retrive1()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************


    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.VSFlexGrid1
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                               
                  .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("advanceID").value), "", rs("advanceID").value)
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                  .TextMatrix(i, .ColIndex("Task")) = IIf(IsNull(rs("task").value), "", rs("task").value)
                  .TextMatrix(i, .ColIndex("startdate")) = IIf(IsNull(rs("startdate").value), "", rs("startdate").value)
                .TextMatrix(i, .ColIndex("Manager_Name")) = IIf(IsNull(rs("Manager_Name").value), "", rs("Manager_Name").value)
                  .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub


Private Sub Retrive2()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************


    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.VSFlexGrid2
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                               
                  .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  .TextMatrix(i, .ColIndex("branch_No")) = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
                  .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("remark").value), "", rs("remark").value)
                  .TextMatrix(i, .ColIndex("recorddate")) = IIf(IsNull(rs("recorddate").value), "", rs("recorddate").value)
                '.TextMatrix(i, .ColIndex("Manager_Name")) = IIf(IsNull(rs("Manager_Name").value), "", rs("Manager_Name").value)
                 ' .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub


Private Sub Retrive3()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************
    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.VSFlexGrid2
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                               
                  .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  .TextMatrix(i, .ColIndex("branch_No")) = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
                  .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("remark").value), "", rs("remark").value)
                  .TextMatrix(i, .ColIndex("recorddate")) = IIf(IsNull(rs("recorddate").value), "", rs("recorddate").value)
                '.TextMatrix(i, .ColIndex("Manager_Name")) = IIf(IsNull(rs("Manager_Name").value), "", rs("Manager_Name").value)
                 ' .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub

Private Sub Retrive4()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************
    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.VSFlexGrid3
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                               
                  .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("advanceID").value), "", rs("advanceID").value)
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  .TextMatrix(i, .ColIndex("branch_No")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                  .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("payValue").value), "", rs("payValue").value)
                  .TextMatrix(i, .ColIndex("recorddate")) = IIf(IsNull(rs("advancedate").value), "", rs("advancedate").value)
                   .TextMatrix(i, .ColIndex("lastpay")) = IIf(IsNull(rs("lastpaydateto").value), "", rs("lastpaydateto").value)
                '.TextMatrix(i, .ColIndex("Manager_Name")) = IIf(IsNull(rs("Manager_Name").value), "", rs("Manager_Name").value)
                 ' .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub


Private Sub Retrive5()
    VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid4.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************
    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.VSFlexGrid4
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                               
                  .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  .TextMatrix(i, .ColIndex("branch_No")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                '  .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("payValue").value), "", rs("payValue").value)
                  .TextMatrix(i, .ColIndex("recorddate")) = IIf(IsNull(rs("recorddate").value), "", rs("recorddate").value)
                .TextMatrix(i, .ColIndex("numbekama")) = IIf(IsNull(rs("numekama").value), "", rs("numekama").value)
                .TextMatrix(i, .ColIndex("pasid")) = IIf(IsNull(rs("pasid").value), "", rs("pasid").value)
                '.TextMatrix(i, .ColIndex("Manager_Name")) = IIf(IsNull(rs("Manager_Name").value), "", rs("Manager_Name").value)
                 ' .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub

Private Sub ChangeLang3()
    ALLButton23.Caption = "Clear"
    ALLButton22.Caption = "Search"
    ALLButton21.Caption = "Exit"
    
     Label46.Caption = "Search Warning Employee"
    ' labell name
   Fra(16).Caption = "Trans No"
   ' Me.lbl(14).Caption = "Trans ID"
    lbl(46).Caption = "From"
    lbl(47).Caption = "From"
    Frame22.Caption = "Data of Sanction"
    lbl(49).Caption = "Sanction"
    lbl(31).Caption = "Reason"
    lbl(45).Caption = "To"
    lbl(48).Caption = "To"
    Fra(17).Caption = "Trans Date"
    lbl(56).Caption = "Warning By"
    lbl(44).Caption = "Employee"
    lbl(54).Caption = "Department"

With VSFlexGrid7
.TextMatrix(0, .ColIndex("Serial")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No"
.TextMatrix(0, .ColIndex("recorddate")) = "Trans Date"
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
.TextMatrix(0, .ColIndex("Emp_Name2")) = "Warning By"
.TextMatrix(0, .ColIndex("DepartmentName")) = "Department"
.TextMatrix(0, .ColIndex("Name")) = "Sanction"

.TextMatrix(0, .ColIndex("Remark")) = "Reason"
End With

  End Sub


Private Sub ChangeLang4()
    ALLButton26.Caption = "Clear"
    ALLButton25.Caption = "Search"
    ALLButton24.Caption = "Exit"
    
     Label47.Caption = "Search Punishment"
    ' labell name
   Fra(13).Caption = "Trans No"
   ' Me.lbl(14).Caption = "Trans ID"
    lbl(43).Caption = "From"
    lbl(50).Caption = "From"
    Frame26.Caption = "Data of Sanction"
    lbl(52).Caption = "Sanction"
    lbl(53).Caption = "Reason"
    lbl(42).Caption = "To"
    lbl(51).Caption = "To"
    Fra(14).Caption = "Trans Date"
  
    lbl(39).Caption = "Employee"
 Label1(5).Caption = "Type"
 Label1(7).Caption = "Component"
With VSFlexGrid8
.TextMatrix(0, .ColIndex("Serial")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No"
.TextMatrix(0, .ColIndex("recorddate")) = "Trans Date"
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
.TextMatrix(0, .ColIndex("nameM")) = "Component"
.TextMatrix(0, .ColIndex("TypDis")) = "Type"
.TextMatrix(0, .ColIndex("Name")) = "Sanction"

.TextMatrix(0, .ColIndex("Remark")) = "Reason"
End With

  End Sub



Private Sub ChangeLang2()
    ALLButton20.Caption = "Clear"
    ALLButton19.Caption = "Search"
    ALLButton18.Caption = "Exit"
    
     Label45.Caption = "Search Start a Work"
    ' labell name
   Fra(7).Caption = "Trans No"
   ' Me.lbl(14).Caption = "Trans ID"
    lbl(29).Caption = "From"
    lbl(35).Caption = "From"
    lbl(34).Caption = "From"
    lbl(37).Caption = "From"
    lbl(28).Caption = "From"
    
    lbl(30).Caption = "To"
    lbl(27).Caption = "To"
    lbl(36).Caption = "To"
    lbl(33).Caption = "To"
    lbl(38).Caption = "To"
    Fra(10).Caption = "Satar Work Date"
    Fra(9).Caption = "Satar Vacation Date"
    Fra(11).Caption = "End Vacation Date"
    Fra(6).Caption = "Trans Date"
    
    lbl(32).Caption = "Employee"
    
Rd(0).Caption = "All"
Rd(0).RightToLeft = False
Rd(1).Caption = "New Employee"
Rd(1).RightToLeft = False
Rd(2).Caption = "Return Vacation"
Rd(2).RightToLeft = False
Frame21.Caption = "Type"
With VSFlexGrid6
.TextMatrix(0, .ColIndex("Serial")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No"
.TextMatrix(0, .ColIndex("recorddate")) = "Trans Date"
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
.TextMatrix(0, .ColIndex("Vac_new")) = "Type"
.TextMatrix(0, .ColIndex("workdate")) = "Start Work"
.TextMatrix(0, .ColIndex("workdateH")) = "Start Work"

.TextMatrix(0, .ColIndex("stratDate")) = "Start Vacation"
.TextMatrix(0, .ColIndex("stratDateH")) = "Start Vacation"
.TextMatrix(0, .ColIndex("EndDate")) = "End Vacation"
.TextMatrix(0, .ColIndex("EndDateH")) = "End Vacation"

End With

  End Sub

Private Sub FromEndDate_Change()
If Not (IsNull(FromEndDate.value)) Then
FromEndDateH.value = ToHijriDate(FromEndDate.value)
End If
End Sub

Private Sub FromEndDateH_LostFocus()
VBA.Calendar = vbCalGreg
            FromEndDate.value = ToGregorianDate(FromEndDateH.value)
End Sub

Private Sub FromStartDate_Change()
If Not (IsNull(FromStartDate.value)) Then
FromStartDateH.value = ToHijriDate(FromStartDate.value)
End If
End Sub

Private Sub FromStartDateH_LostFocus()
VBA.Calendar = vbCalGreg
            FromStartDate.value = ToGregorianDate(FromStartDateH.value)
End Sub

Private Sub NourHijriCal1_LostFocus()
VBA.Calendar = vbCalGreg
            DTPicker7.value = ToGregorianDate(NourHijriCal1.value)
End Sub

Private Sub NourHijriCal2_LostFocus()
VBA.Calendar = vbCalGreg
            DTPicker8.value = ToGregorianDate(NourHijriCal2.value)
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text21.Text, EmpID
        DcbEmp10.BoundText = EmpID
    End If
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text24.Text, EmpID
        DcbEmp12.BoundText = EmpID
    End If
End Sub

Private Sub ToEndDate_Change()
If Not (IsNull(ToEndDate.value)) Then
ToEndDateH.value = ToHijriDate(ToEndDate.value)
End If
End Sub

Private Sub ToEndDateH_LostFocus()
VBA.Calendar = vbCalGreg
            ToEndDate.value = ToGregorianDate(ToEndDateH.value)
End Sub

Private Sub ToStartDate_Change()
If Not (IsNull(ToStartDate.value)) Then
ToStartDateH.value = ToHijriDate(ToStartDate.value)
End If
End Sub

Private Sub ToStartDateH_LostFocus()
VBA.Calendar = vbCalGreg
            ToStartDate.value = ToGregorianDate(ToStartDateH.value)
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)

    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then

    ElseIf KeyAscii = vbKeyBack Then

    Else
      KeyAscii = 0
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
 
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcbEmp.BoundText = EmpID
    End If
End Sub

Private Sub TxtSearchCode2_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode2.Text, EmpID
        DcbEmp11.BoundText = EmpID
    End If
End Sub

Private Sub VSFlexGrid1_Click()
ALLButton2_Click
End Sub

Private Sub VSFlexGrid2_Click()
ALLButton5_Click
End Sub

Private Sub VSFlexGrid3_Click()
ALLButton8_Click
End Sub

Private Sub VSFlexGrid4_Click()
ALLButton11_Click
End Sub

Private Sub VSFlexGrid5_Click()
ALLButton16_Click
End Sub

Private Sub VSFlexGrid6_Click()
With VSFlexGrid6
FrmEmbarkation.Retrive val(.TextMatrix(.Row, .ColIndex("ID")))
End With
End Sub

Private Sub VSFlexGrid7_Click()
With VSFlexGrid7
If Index = 1 Then
FRmEmployeeWarning.Retrive val(.TextMatrix(.Row, .ColIndex("ID")))
ElseIf Index = 2 Then
FrmKhsm.TxtAlrmOrder.Text = val(.TextMatrix(.Row, .ColIndex("ID")))
End If
End With
End Sub

Private Sub VSFlexGrid8_Click()
With VSFlexGrid8
FrmKhsm.FindRec (val(.TextMatrix(.Row, .ColIndex("ID"))))
End With
End Sub
